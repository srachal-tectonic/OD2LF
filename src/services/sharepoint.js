const config = require('../config');
const logger = require('../logger');
const { graphClient } = require('./graph-auth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SITE_ID = config.microsoft.siteId;

// In-memory delta token storage. In production with multiple dynos, use a database.
let deltaToken = null;
let subscriptionId = null;

// ── Subscription Management ──────────────────────────────────

const EXPECTED_NOTIFICATION_URL = () => `${config.appUrl}/api/webhook`;
const EXPECTED_RESOURCE = `sites/${SITE_ID}/drive/root`;
const RENEW_IF_LESS_THAN_MS = 48 * 60 * 60 * 1000; // renew if <48h until expiry

let currentSubscription = null; // cached full subscription object

async function listSubscriptions() {
  const client = graphClient();
  const response = await client.get(`${GRAPH_BASE}/subscriptions`);
  return response.data.value || [];
}

async function findMatchingSubscription() {
  const expectedUrl = EXPECTED_NOTIFICATION_URL();
  const subs = await listSubscriptions();
  return subs.find(
    (s) => s.notificationUrl === expectedUrl && s.resource === EXPECTED_RESOURCE
  );
}

async function createSubscription() {
  const client = graphClient();
  const expiration = new Date();
  expiration.setDate(expiration.getDate() + 29); // Max ~30 days

  const payload = {
    changeType: 'updated',
    notificationUrl: EXPECTED_NOTIFICATION_URL(),
    resource: EXPECTED_RESOURCE,
    expirationDateTime: expiration.toISOString(),
    clientState: config.microsoft.webhookClientState,
  };

  logger.info('Creating SharePoint webhook subscription', {
    notificationUrl: payload.notificationUrl,
    expiration: payload.expirationDateTime,
  });

  const response = await client.post(`${GRAPH_BASE}/subscriptions`, payload);
  subscriptionId = response.data.id;
  currentSubscription = response.data;

  logger.info('Subscription created', { subscriptionId });

  // Get initial delta token so we only see future changes
  await initializeDeltaToken();

  return response.data;
}

async function renewSubscription() {
  if (!subscriptionId) {
    logger.warn('No subscription to renew, creating new one');
    return createSubscription();
  }

  const client = graphClient();
  const expiration = new Date();
  expiration.setDate(expiration.getDate() + 29);

  try {
    const response = await client.patch(
      `${GRAPH_BASE}/subscriptions/${subscriptionId}`,
      { expirationDateTime: expiration.toISOString() }
    );

    logger.info('Subscription renewed', {
      subscriptionId,
      newExpiration: response.data.expirationDateTime,
    });

    currentSubscription = response.data;
    return response.data;
  } catch (err) {
    if (err.response?.status === 404) {
      logger.warn('Subscription not found, creating new one');
      return createSubscription();
    }
    throw err;
  }
}

/**
 * Idempotent: guarantees exactly one Graph subscription matching our
 * notificationUrl + resource. Safe to call on every startup and by the
 * daily cron. Adopts an existing sub, renews if close to expiry, and
 * only creates when truly missing.
 */
async function ensureSubscription() {
  const existing = await findMatchingSubscription();

  if (!existing) {
    logger.info('No matching subscription found; creating new one');
    return createSubscription();
  }

  subscriptionId = existing.id;
  currentSubscription = existing;

  const msUntilExpiry = new Date(existing.expirationDateTime).getTime() - Date.now();

  if (msUntilExpiry <= 0) {
    logger.warn('Adopted subscription is expired; recreating', {
      subscriptionId: existing.id,
    });
    return createSubscription();
  }

  // Prime delta token so processNewItems() doesn't replay history
  if (!deltaToken) {
    try {
      await initializeDeltaToken();
    } catch (err) {
      logger.warn('Failed to initialize delta token on adopt', { error: err.message });
    }
  }

  if (msUntilExpiry < RENEW_IF_LESS_THAN_MS) {
    logger.info('Adopted subscription close to expiry; renewing', {
      subscriptionId: existing.id,
      hoursUntilExpiry: Math.round(msUntilExpiry / 3600000),
    });
    return renewSubscription();
  }

  logger.info('Adopted existing subscription', {
    subscriptionId: existing.id,
    expires: existing.expirationDateTime,
  });
  return existing;
}

function getSubscriptionId() {
  return subscriptionId;
}

function getSubscriptionState() {
  return {
    id: subscriptionId,
    expirationDateTime: currentSubscription?.expirationDateTime || null,
    notificationUrl: currentSubscription?.notificationUrl || null,
    resource: currentSubscription?.resource || null,
  };
}

function setSubscriptionId(id) {
  subscriptionId = id;
}

// ── Delta Query ──────────────────────────────────────────────

// Drive-root delta is used (not folder-scoped) because folder-scoped delta on
// SharePoint document libraries has known instability — tokens can be
// invalidated by background indexing and return 410 Gone even when nothing in
// the scoped folder has changed. Drive-root delta is reliable; the cost is
// receiving events for items outside the target, which we filter out below.

async function initializeDeltaToken() {
  const client = graphClient();
  const response = await client.get(
    `${GRAPH_BASE}/sites/${SITE_ID}/drive/root/delta?token=latest`
  );

  const deltaLink = response.data['@odata.deltaLink'];
  if (deltaLink) {
    deltaToken = new URL(deltaLink).searchParams.get('token');
    logger.info('Delta token initialized');
  }
}

async function getChangedItems() {
  const client = graphClient();
  const baseUrl = `${GRAPH_BASE}/sites/${SITE_ID}/drive/root/delta`;

  let url = baseUrl;
  if (deltaToken) {
    url += `?token=${encodeURIComponent(deltaToken)}`;
  }

  const allChanges = [];

  while (url) {
    let response;
    try {
      response = await client.get(url);
    } catch (err) {
      // 410 Gone: delta token has been invalidated by SharePoint. Reset and
      // re-prime so the next webhook starts from a fresh state. Changes that
      // occurred during the gap will be missed by delta but typically arrive
      // again as their own webhooks (e.g. ongoing uploads).
      if (err.response?.status === 410) {
        logger.warn('Delta returned 410 Gone; resetting token and re-priming', {
          hadToken: !!deltaToken,
        });
        deltaToken = null;
        try {
          await initializeDeltaToken();
        } catch (initErr) {
          logger.error('Failed to re-initialize delta token after 410', {
            error: initErr.message,
          });
        }
        return [];
      }
      throw err;
    }

    const data = response.data;
    allChanges.push(...(data.value || []));

    if (data['@odata.nextLink']) {
      url = data['@odata.nextLink'];
    } else if (data['@odata.deltaLink']) {
      deltaToken = new URL(data['@odata.deltaLink']).searchParams.get('token');
      url = null;
    } else {
      url = null;
    }
  }

  return allChanges;
}

// Returns the path segments below the configured target folder, or null
// if the parent is outside the target subtree.
//   parent === target               → []
//   parent === target/A             → ['A']
//   parent === target/A/B           → ['A', 'B']
function relativeSegmentsFromTarget(parentPath) {
  if (!parentPath) return null;
  const targetPath = config.microsoft.targetFolderPath;
  if (parentPath.endsWith(targetPath)) return [];
  const marker = `${targetPath}/`;
  const idx = parentPath.indexOf(marker);
  if (idx === -1) return null;
  return parentPath.slice(idx + marker.length).split('/').filter(Boolean);
}

async function getNewItemsInTarget() {
  const allChanges = await getChangedItems();
  const targetPath = config.microsoft.targetFolderPath;

  // Only log items inside the target subtree so debug output isn't drowned
  // out by activity elsewhere on the drive.
  let inTargetCount = 0;
  for (const item of allChanges) {
    const segments = relativeSegmentsFromTarget(item.parentReference?.path);
    if (segments === null) continue;
    inTargetCount++;
    logger.debug('Delta item in target', {
      name: item.name,
      isFolder: item.folder !== undefined,
      isFile: item.file !== undefined,
      deleted: item.deleted !== undefined,
      relativePath: segments.join('/') || '(target root)',
    });
  }
  if (allChanges.length > 0) {
    logger.debug('Delta query summary', {
      totalChanges: allChanges.length,
      inTarget: inTargetCount,
    });
  }

  const newFolders = allChanges.filter((item) => {
    const isFolder = item.folder !== undefined;
    const isNotDeleted = item.deleted === undefined;
    const parentPath = item.parentReference?.path || '';
    const isInTarget = parentPath.endsWith(targetPath);
    return isFolder && isNotDeleted && isInTarget;
  });

  // Files: match anywhere within the target subtree, not just direct children.
  // Files dropped via folder upload arrive in the delta after the folder event,
  // with parentPath ending in <subfolder>, not the target itself.
  // Deduplicate by item id — Graph delta often returns the same file multiple
  // times (e.g. .docx triggers create + metadata update + indexing events).
  const seenFileIds = new Set();
  const newFiles = [];
  for (const item of allChanges) {
    if (item.file === undefined || item.deleted !== undefined) continue;
    const segments = relativeSegmentsFromTarget(item.parentReference?.path);
    if (segments === null) continue;
    if (seenFileIds.has(item.id)) continue;
    seenFileIds.add(item.id);
    newFiles.push({ ...item, relativeSegments: segments });
  }

  if (newFolders.length > 0 || newFiles.length > 0) {
    logger.info('Delta query found new items in target', {
      totalChanges: allChanges.length,
      newFoldersInTarget: newFolders.length,
      newFilesInTarget: newFiles.length,
      targetPath,
    });
  } else {
    logger.debug('Delta query: no matching items', {
      totalChanges: allChanges.length,
      targetPath,
    });
  }

  return { newFolders, newFiles };
}

async function getNewFoldersInTarget() {
  const { newFolders } = await getNewItemsInTarget();
  return newFolders;
}

// ── Download File Content ────────────────────────────────────

async function getFolderContents(folderId) {
  const client = graphClient();
  const response = await client.get(
    `${GRAPH_BASE}/sites/${SITE_ID}/drive/items/${folderId}/children`
  );
  return response.data.value || [];
}

async function downloadFile(itemId) {
  const client = graphClient();
  // Get download URL
  const metaResponse = await client.get(
    `${GRAPH_BASE}/sites/${SITE_ID}/drive/items/${itemId}`
  );

  const downloadUrl = metaResponse.data['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    throw new Error(`No download URL for item ${itemId}`);
  }

  // Download the actual file content
  const fileResponse = await require('axios').get(downloadUrl, {
    responseType: 'arraybuffer',
  });

  return {
    name: metaResponse.data.name,
    mimeType: metaResponse.data.file?.mimeType || 'application/octet-stream',
    size: metaResponse.data.size,
    content: Buffer.from(fileResponse.data),
  };
}

module.exports = {
  createSubscription,
  renewSubscription,
  ensureSubscription,
  listSubscriptions,
  findMatchingSubscription,
  getSubscriptionId,
  getSubscriptionState,
  setSubscriptionId,
  getNewFoldersInTarget,
  getNewItemsInTarget,
  getFolderContents,
  downloadFile,
};

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

  let url = `${GRAPH_BASE}/sites/${SITE_ID}/drive/root/delta`;
  if (deltaToken) {
    url += `?token=${encodeURIComponent(deltaToken)}`;
  }

  const allChanges = [];

  while (url) {
    const response = await client.get(url);
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

async function getNewItemsInTarget() {
  const allChanges = await getChangedItems();
  const targetPath = config.microsoft.targetFolderPath;

  // Log all changes so we can see the actual parentReference.path format
  for (const item of allChanges) {
    logger.debug('Delta item', {
      name: item.name,
      isFolder: item.folder !== undefined,
      isFile: item.file !== undefined,
      deleted: item.deleted !== undefined,
      parentPath: item.parentReference?.path,
    });
  }

  const newFolders = allChanges.filter((item) => {
    const isFolder = item.folder !== undefined;
    const isNotDeleted = item.deleted === undefined;
    const parentPath = item.parentReference?.path || '';
    const isInTarget = parentPath.endsWith(targetPath);
    return isFolder && isNotDeleted && isInTarget;
  });

  // Deduplicate by item id — Graph delta often returns the same file multiple
  // times (e.g. .docx triggers create + metadata update + indexing events)
  const seenFileIds = new Set();
  const newFiles = allChanges.filter((item) => {
    const isFile = item.file !== undefined;
    const isNotDeleted = item.deleted === undefined;
    const parentPath = item.parentReference?.path || '';
    const isInTarget = parentPath.endsWith(targetPath);
    if (isFile && isNotDeleted && isInTarget) {
      if (seenFileIds.has(item.id)) return false;
      seenFileIds.add(item.id);
      return true;
    }
    return false;
  });

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

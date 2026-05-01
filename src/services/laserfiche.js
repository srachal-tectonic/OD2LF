const axios = require('axios');
const https = require('https');
const FormData = require('form-data');
const config = require('../config');
const logger = require('../logger');

let token = null;
let tokenExpiry = 0;

const baseUrl = `${config.laserfiche.serverUrl}/LFRepositoryAPI/v1/Repositories/${config.laserfiche.repoName}`;

const httpClient = axios.create({
  baseURL: baseUrl,
  httpsAgent: new https.Agent({
    rejectUnauthorized: config.laserfiche.verifySsl,
  }),
});

// Add auth interceptor
httpClient.interceptors.request.use((reqConfig) => {
  if (token) {
    reqConfig.headers.Authorization = `Bearer ${token}`;
  }
  return reqConfig;
});

// ── Authentication ───────────────────────────────────────────

async function authenticate() {
  // Return cached token if valid (with 2-minute buffer)
  if (token && Date.now() < tokenExpiry - 2 * 60 * 1000) {
    return token;
  }

  const response = await httpClient.post(
    '/Token',
    new URLSearchParams({
      grant_type: 'password',
      username: config.laserfiche.username,
      password: config.laserfiche.password,
    }),
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
  );

  token = response.data.access_token;
  // Default expiry is typically 15 minutes
  tokenExpiry = Date.now() + (response.data.expires_in || 900) * 1000;

  logger.info('Laserfiche authenticated', {
    expiresIn: response.data.expires_in,
  });

  return token;
}

async function ensureAuthenticated() {
  if (!token || Date.now() >= tokenExpiry - 2 * 60 * 1000) {
    await authenticate();
  }
}

// ── Folder Operations ────────────────────────────────────────

async function createFolder(parentId, folderName) {
  await ensureAuthenticated();

  logger.debug('Creating Laserfiche folder', { parentId, folderName });

  try {
    const response = await httpClient.post(
      `/Entries/${parentId}/Laserfiche.Repository.Folder/children`,
      { entryType: 'Folder', name: folderName }
    );

    logger.info('Laserfiche folder created', {
      parentId,
      folderName,
      entryId: response.data.id,
    });

    return response.data;
  } catch (err) {
    logger.error('Laserfiche folder create failed', {
      parentId,
      folderName,
      status: err.response?.status,
      body: err.response?.data,
    });
    throw err;
  }
}

// Pages through @odata.nextLink so we don't miss entries when the parent
// folder has more children than LF's default page size — the prior single
// GET was the source of 409s when an existing subfolder fell off page 1.
async function getFolderChildren(entryId) {
  await ensureAuthenticated();

  logger.debug('Getting folder children', { entryId });

  const all = [];
  let url = `/Entries/${encodeURIComponent(entryId)}/Laserfiche.Repository.Folder/children`;

  while (url) {
    const response = await httpClient.get(url);
    const data = response.data || {};
    const page = data.value || data.Value || [];
    all.push(...page);

    const next =
      data['@odata.nextLink'] ||
      data['@nextLink'] ||
      data.nextLink ||
      null;
    url = next || null;
  }

  return { value: all };
}

async function findOrCreateFolder(parentId, folderName) {
  await ensureAuthenticated();

  const findExisting = async () => {
    const listing = await getFolderChildren(parentId);
    const entries = listing.value || listing.Value || [];
    return entries.find(
      (e) => e.name === folderName && e.entryType === 'Folder'
    );
  };

  try {
    const existing = await findExisting();
    if (existing) {
      logger.info('Laserfiche folder found', { folderName, entryId: existing.id });
      return existing;
    }
  } catch (err) {
    logger.warn('Folder lookup failed, will attempt create', {
      folderName,
      error: err.response?.status || err.message,
      url: err.config?.url,
    });
  }

  try {
    return await createFolder(parentId, folderName);
  } catch (err) {
    // 409 from create means the folder already exists (race or stale listing).
    // Re-list and locate it so callers get a usable id instead of a hard fail.
    if (err.response?.status === 409) {
      logger.info('Folder create 409, re-listing to locate existing', {
        parentId,
        folderName,
      });
      try {
        const existing = await findExisting();
        if (existing) {
          logger.info('Recovered existing folder after 409', {
            folderName,
            entryId: existing.id,
          });
          return existing;
        }
      } catch (relistErr) {
        logger.error('Re-list after 409 failed', {
          folderName,
          error: relistErr.response?.status || relistErr.message,
        });
      }
    }
    throw err;
  }
}

// ── Document Upload ──────────────────────────────────────────

async function uploadDocument(parentFolderId, fileName, fileBuffer, mimeType) {
  await ensureAuthenticated();

  const form = new FormData();
  form.append('electronicDocument', fileBuffer, {
    filename: fileName,
    contentType: mimeType || null,
  });

  try {
    // Use the same upload pattern as the working app: /Entries/{id}/{fileName}
    const response = await httpClient.post(
      `/Entries/${parentFolderId}/${encodeURIComponent(fileName)}?autoRename=true`,
      form,
      {
        headers: form.getHeaders(),
        maxContentLength: Infinity,
        maxBodyLength: Infinity,
      }
    );

    logger.info('Document uploaded to Laserfiche', {
      fileName,
      parentFolderId,
      entryId: response.data.id,
    });

    return response.data;
  } catch (err) {
    logger.error('Laserfiche document upload failed', {
      fileName,
      parentFolderId,
      status: err.response?.status,
      body: err.response?.data,
    });
    throw err;
  }
}

// ── Logout ───────────────────────────────────────────────────

async function logout() {
  if (token) {
    try {
      await httpClient.delete('/Token');
    } catch (err) {
      // Ignore logout errors
    }
    token = null;
    tokenExpiry = 0;
  }
}

module.exports = {
  authenticate,
  createFolder,
  findOrCreateFolder,
  getFolderChildren,
  uploadDocument,
  logout,
};

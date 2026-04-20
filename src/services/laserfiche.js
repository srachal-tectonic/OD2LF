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
}

async function getFolderChildren(entryId) {
  await ensureAuthenticated();

  logger.debug('Getting folder children', { entryId });

  const response = await httpClient.get(
    `/Entries/${encodeURIComponent(entryId)}/Laserfiche.Repository.Folder/children`
  );
  return response.data;
}

async function findOrCreateFolder(parentId, folderName) {
  await ensureAuthenticated();

  try {
    const listing = await getFolderChildren(parentId);
    const entries = listing.value || listing.Value || [];
    const existing = entries.find(
      (e) => e.name === folderName && e.entryType === 'Folder'
    );

    if (existing) {
      logger.info('Laserfiche folder found', { folderName, entryId: existing.id });
      return existing;
    }
  } catch (err) {
    logger.warn('Folder lookup failed, will create new', {
      folderName,
      error: err.response?.status || err.message,
      url: err.config?.url,
    });
  }

  return createFolder(parentId, folderName);
}

// ── Document Upload ──────────────────────────────────────────

async function uploadDocument(parentFolderId, fileName, fileBuffer, mimeType) {
  await ensureAuthenticated();

  const form = new FormData();
  form.append('electronicDocument', fileBuffer, {
    filename: fileName,
    contentType: mimeType || null,
  });

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

#!/usr/bin/env node

/**
 * Bulk Upload Script
 *
 * Uploads an entire SharePoint folder (with all subfolders and files)
 * into Laserfiche, preserving the folder structure.
 *
 * Usage:
 *   node src/scripts/upload-folder.js "<SharePoint folder path>"
 *
 * The folder path is relative to the SharePoint site's default drive root.
 */

require('dotenv').config();

const axios = require('axios');
const config = require('../config');
const logger = require('../logger');
const { graphClient } = require('../services/graph-auth');
const laserfiche = require('../services/laserfiche');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SITE_ID = config.microsoft.siteId;

// ── Stats tracking ──────────────────────────────────────────

const stats = {
  foldersCreated: 0,
  filesUploaded: 0,
  filesFailed: 0,
  totalBytes: 0,
  errors: [],
};

// ── SharePoint helpers ──────────────────────────────────────

async function getDriveItemByPath(folderPath) {
  const client = graphClient();
  const encodedPath = encodeURIComponent(folderPath).replace(/%2F/g, '/');
  const url = `${GRAPH_BASE}/sites/${SITE_ID}/drive/root:/${encodedPath}`;
  const response = await client.get(url);
  return response.data;
}

async function listChildren(itemId) {
  const client = graphClient();
  const allItems = [];
  let url = `${GRAPH_BASE}/sites/${SITE_ID}/drive/items/${itemId}/children?$top=200`;

  while (url) {
    const response = await client.get(url);
    const data = response.data;
    allItems.push(...(data.value || []));
    url = data['@odata.nextLink'] || null;
  }

  return allItems;
}

async function downloadFile(itemId) {
  const client = graphClient();
  const metaResponse = await client.get(
    `${GRAPH_BASE}/sites/${SITE_ID}/drive/items/${itemId}`
  );

  const downloadUrl = metaResponse.data['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    throw new Error(`No download URL for item ${itemId}`);
  }

  const fileResponse = await axios.get(downloadUrl, {
    responseType: 'arraybuffer',
  });

  return {
    name: metaResponse.data.name,
    mimeType: metaResponse.data.file?.mimeType || 'application/octet-stream',
    size: metaResponse.data.size,
    content: Buffer.from(fileResponse.data),
  };
}

// ── Recursive upload ────────────────────────────────────────

async function uploadFolderRecursive(spItemId, lfParentFolderId, depth = 0) {
  const children = await listChildren(spItemId);
  const indent = '  '.repeat(depth);

  const folders = children.filter((c) => c.folder !== undefined);
  const files = children.filter((c) => c.file !== undefined);

  // Upload files in this folder
  for (const file of files) {
    try {
      console.log(`${indent}  -> ${file.name} (${formatBytes(file.size)})`);
      const downloaded = await downloadFile(file.id);

      await laserfiche.uploadDocument(
        lfParentFolderId,
        downloaded.name,
        downloaded.content,
        downloaded.mimeType
      );

      stats.filesUploaded++;
      stats.totalBytes += downloaded.size;
    } catch (err) {
      stats.filesFailed++;
      stats.errors.push({ file: file.name, error: err.message });
      console.error(`${indent}  !! FAILED: ${file.name} - ${err.message}`);
    }
  }

  // Recurse into subfolders
  for (const folder of folders) {
    console.log(`${indent}[FOLDER] ${folder.name}/`);

    try {
      const lfFolder = await laserfiche.findOrCreateFolder(
        lfParentFolderId,
        folder.name
      );
      stats.foldersCreated++;

      await uploadFolderRecursive(folder.id, lfFolder.id, depth + 1);
    } catch (err) {
      stats.errors.push({ folder: folder.name, error: err.message });
      console.error(`${indent}  !! FAILED to create folder: ${folder.name} - ${err.message}`);
    }
  }
}

// ── Utilities ───────────────────────────────────────────────

function formatBytes(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${(bytes / Math.pow(k, i)).toFixed(1)} ${sizes[i]}`;
}

function formatDuration(ms) {
  const seconds = Math.floor(ms / 1000);
  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = seconds % 60;
  if (minutes > 0) {
    return `${minutes}m ${remainingSeconds}s`;
  }
  return `${seconds}.${String(ms % 1000).padStart(3, '0')}s`;
}

// ── Main ────────────────────────────────────────────────────

async function main() {
  const folderPath = process.argv[2];

  if (!folderPath) {
    console.error('Usage: node src/scripts/upload-folder.js "<SharePoint folder path>"');
    console.error('Example: node src/scripts/upload-folder.js "General/SBA Loans/_Completed Loans/<folder name>"');
    process.exit(1);
  }

  const lfDestinationId = config.laserfiche.destinationFolderId;

  console.log('='.repeat(70));
  console.log('BULK FOLDER UPLOAD: SharePoint -> Laserfiche');
  console.log('='.repeat(70));
  console.log(`Source:      SharePoint:/${folderPath}`);
  console.log(`Destination: Laserfiche folder ID ${lfDestinationId}`);
  console.log('='.repeat(70));
  console.log();

  const startTime = Date.now();

  try {
    // Authenticate with Laserfiche
    console.log('[1/3] Authenticating with Laserfiche...');
    await laserfiche.authenticate();
    console.log('      Laserfiche authenticated.\n');

    // Resolve the SharePoint folder
    console.log('[2/3] Resolving SharePoint folder...');
    const rootItem = await getDriveItemByPath(folderPath);

    if (!rootItem.folder) {
      console.error(`ERROR: "${folderPath}" is not a folder.`);
      process.exit(1);
    }

    const childCount = rootItem.folder.childCount;
    console.log(`      Found: "${rootItem.name}" (${childCount} direct children)\n`);

    // Create the root folder in Laserfiche
    console.log('[3/3] Starting recursive upload...\n');
    const lfRootFolder = await laserfiche.findOrCreateFolder(
      lfDestinationId,
      rootItem.name
    );
    stats.foldersCreated++;

    console.log(`[FOLDER] ${rootItem.name}/ (LF ID: ${lfRootFolder.id})`);

    // Recursively upload everything
    await uploadFolderRecursive(rootItem.id, lfRootFolder.id, 1);

  } catch (err) {
    console.error(`\nFATAL ERROR: ${err.message}`);
    if (err.response?.data) {
      console.error('Response:', JSON.stringify(err.response.data, null, 2));
    }
    logger.error('Bulk upload failed', { error: err.message, stack: err.stack });
  } finally {
    await laserfiche.logout();
  }

  const elapsed = Date.now() - startTime;

  // Print summary
  console.log();
  console.log('='.repeat(70));
  console.log('UPLOAD COMPLETE');
  console.log('='.repeat(70));
  console.log(`Time elapsed:    ${formatDuration(elapsed)}`);
  console.log(`Folders created: ${stats.foldersCreated}`);
  console.log(`Files uploaded:  ${stats.filesUploaded}`);
  console.log(`Files failed:    ${stats.filesFailed}`);
  console.log(`Total data:      ${formatBytes(stats.totalBytes)}`);

  if (stats.errors.length > 0) {
    console.log(`\nErrors (${stats.errors.length}):`);
    for (const e of stats.errors) {
      console.log(`  - ${e.file || e.folder}: ${e.error}`);
    }
  }

  console.log('='.repeat(70));
}

main();

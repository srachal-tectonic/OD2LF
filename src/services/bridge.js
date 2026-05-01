const sharepoint = require('./sharepoint');
const laserfiche = require('./laserfiche');
const config = require('../config');
const logger = require('../logger');

/**
 * Main bridge logic: detect new folders and files in the SharePoint target folder,
 * download them and upload to Laserfiche.
 *
 * - New folders: create a matching folder in Laserfiche, then transfer all files inside.
 * - New files (directly in target): upload straight to the Laserfiche destination folder.
 */
async function processNewItems() {
  const { newFolders, newFiles } = await sharepoint.getNewItemsInTarget();

  let foldersProcessed = 0;
  let foldersFailed = 0;
  let filesProcessed = 0;
  let filesFailed = 0;

  if (newFolders.length === 0 && newFiles.length === 0) {
    logger.debug('No new folders or files in target path, skipping');
    return { foldersProcessed: 0, foldersFailed: 0, filesProcessed: 0, filesFailed: 0 };
  }

  logger.info('Bridge processing pass starting', {
    newFolders: newFolders.length,
    newFiles: newFiles.length,
  });

  // Cache LF folder ids per relative path within this pass so we don't
  // re-list the same parent for every file in the same subfolder.
  const folderCache = new Map();
  folderCache.set('', config.laserfiche.destinationFolderId);

  async function ensureLfFolderPath(segments) {
    let parentId = config.laserfiche.destinationFolderId;
    let key = '';
    for (const seg of segments) {
      key = key ? `${key}/${seg}` : seg;
      let id = folderCache.get(key);
      if (!id) {
        const lfFolder = await laserfiche.findOrCreateFolder(parentId, seg);
        id = lfFolder.id;
        folderCache.set(key, id);
      }
      parentId = id;
    }
    return parentId;
  }

  // ── Eagerly create direct-child folders ─────────────────────
  // Files arriving in subsequent delta pages will populate these via the
  // file loop below. Creating eagerly keeps empty folders visible in LF.
  for (const folder of newFolders) {
    try {
      logger.info('Processing new folder', {
        folderName: folder.name,
        folderId: folder.id,
      });
      await ensureLfFolderPath([folder.name]);
      foldersProcessed++;
    } catch (folderErr) {
      foldersFailed++;
      logger.error('Failed to create folder in Laserfiche', {
        folderName: folder.name,
        folderId: folder.id,
        error: folderErr.message,
        status: folderErr.response?.status,
        responseBody: folderErr.response?.data,
      });
    }
  }

  // ── Process new files anywhere under the target ─────────────
  for (const file of newFiles) {
    const segments = file.relativeSegments || [];
    const relativePath = segments.join('/') || '(target root)';
    try {
      logger.info('Processing new file', {
        fileName: file.name,
        fileId: file.id,
        relativePath,
      });

      const targetFolderId = await ensureLfFolderPath(segments);
      const downloaded = await sharepoint.downloadFile(file.id);

      await laserfiche.uploadDocument(
        targetFolderId,
        downloaded.name,
        downloaded.content,
        downloaded.mimeType
      );

      filesProcessed++;
      logger.info('File transferred successfully', {
        fileName: downloaded.name,
        size: downloaded.size,
        relativePath,
      });
    } catch (fileErr) {
      filesFailed++;
      logger.error('Failed to transfer file', {
        fileName: file.name,
        fileId: file.id,
        relativePath,
        error: fileErr.message,
        status: fileErr.response?.status,
        responseBody: fileErr.response?.data,
      });
    }
  }

  logger.info('Bridge processing complete', {
    foldersProcessed,
    foldersFailed,
    filesProcessed,
    filesFailed,
    totalSeen: newFolders.length + newFiles.length,
  });
  return { foldersProcessed, foldersFailed, filesProcessed, filesFailed };
}

// Keep backward-compatible export
async function processNewFolders() {
  const result = await processNewItems();
  return result.foldersProcessed;
}

module.exports = { processNewFolders, processNewItems };

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
  let filesProcessed = 0;

  if (newFolders.length === 0 && newFiles.length === 0) {
    logger.debug('No new folders or files in target path, skipping');
    return { foldersProcessed: 0, filesProcessed: 0 };
  }

  // ── Process new folders (existing behaviour) ────────────────
  for (const folder of newFolders) {
    try {
      logger.info('Processing new folder', {
        folderName: folder.name,
        folderId: folder.id,
      });

      // Create matching folder in Laserfiche
      const lfFolder = await laserfiche.findOrCreateFolder(
        config.laserfiche.destinationFolderId,
        folder.name
      );

      // Get all files in the SharePoint folder
      const contents = await sharepoint.getFolderContents(folder.id);
      const files = contents.filter((item) => item.file !== undefined);

      logger.info('Found files in folder', {
        folderName: folder.name,
        fileCount: files.length,
      });

      for (const file of files) {
        try {
          const downloaded = await sharepoint.downloadFile(file.id);

          await laserfiche.uploadDocument(
            lfFolder.id,
            downloaded.name,
            downloaded.content,
            downloaded.mimeType
          );

          logger.info('File transferred successfully', {
            fileName: downloaded.name,
            size: downloaded.size,
          });
        } catch (fileErr) {
          logger.error('Failed to transfer file', {
            fileName: file.name,
            folderId: folder.id,
            error: fileErr.message,
          });
        }
      }

      foldersProcessed++;
      logger.info('Folder processing complete', {
        folderName: folder.name,
        filesProcessed: files.length,
      });
    } catch (folderErr) {
      logger.error('Failed to process folder', {
        folderName: folder.name,
        folderId: folder.id,
        error: folderErr.message,
        status: folderErr.response?.status,
        responseBody: folderErr.response?.data,
      });
    }
  }

  // ── Process new files directly in target folder ─────────────
  for (const file of newFiles) {
    try {
      logger.info('Processing new file in target folder', {
        fileName: file.name,
        fileId: file.id,
      });

      const downloaded = await sharepoint.downloadFile(file.id);

      await laserfiche.uploadDocument(
        config.laserfiche.destinationFolderId,
        downloaded.name,
        downloaded.content,
        downloaded.mimeType
      );

      filesProcessed++;
      logger.info('File transferred successfully', {
        fileName: downloaded.name,
        size: downloaded.size,
      });
    } catch (fileErr) {
      logger.error('Failed to transfer file from target folder', {
        fileName: file.name,
        fileId: file.id,
        error: fileErr.message,
      });
    }
  }

  logger.info('Bridge processing complete', { foldersProcessed, filesProcessed });
  return { foldersProcessed, filesProcessed };
}

// Keep backward-compatible export
async function processNewFolders() {
  const result = await processNewItems();
  return result.foldersProcessed;
}

module.exports = { processNewFolders, processNewItems };

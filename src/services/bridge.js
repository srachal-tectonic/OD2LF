const sharepoint = require('./sharepoint');
const laserfiche = require('./laserfiche');
const config = require('../config');
const logger = require('../logger');

/**
 * Main bridge logic: detect new folders and files in the SharePoint target folder,
 * download them and upload to Laserfiche.
 *
 * Two sources of files in a pass:
 * - Delta-detected files: catch ongoing uploads (file events fire as items are
 *   indexed by SharePoint).
 * - Recursive enumeration of new folders: catch folders that were MOVED into
 *   the target — moves don't fire individual events for child items, so delta
 *   alone never sees them.
 *
 * uploadedFileIds dedups across both sources within an app instance so we
 * don't double-upload a file picked up by enumeration that later also fires
 * its own delta event.
 */

const uploadedFileIds = new Set();

async function collectFilesRecursive(spFolderId, baseSegments, accumMap) {
  const items = await sharepoint.getFolderContents(spFolderId);
  for (const item of items) {
    if (item.file !== undefined) {
      if (!accumMap.has(item.id)) {
        accumMap.set(item.id, { file: item, segments: baseSegments });
      }
    } else if (item.folder !== undefined) {
      await collectFilesRecursive(
        item.id,
        [...baseSegments, item.name],
        accumMap
      );
    }
  }
}

async function processNewItems() {
  const { newFolders, newFiles } = await sharepoint.getNewItemsInTarget();

  let foldersProcessed = 0;
  let foldersFailed = 0;
  let filesProcessed = 0;
  let filesFailed = 0;
  let filesSkipped = 0;

  if (newFolders.length === 0 && newFiles.length === 0) {
    logger.debug('No new folders or files in target path, skipping');
    return { foldersProcessed: 0, foldersFailed: 0, filesProcessed: 0, filesFailed: 0 };
  }

  logger.info('Bridge processing pass starting', {
    newFolders: newFolders.length,
    deltaFiles: newFiles.length,
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

  // Aggregate every file we should process, keyed by SP item id.
  const filesToProcess = new Map();
  for (const f of newFiles) {
    filesToProcess.set(f.id, { file: f, segments: f.relativeSegments || [] });
  }

  // ── Process new folders + recursively enumerate their contents ──
  for (const folder of newFolders) {
    try {
      logger.info('Processing new folder', {
        folderName: folder.name,
        folderId: folder.id,
      });
      await ensureLfFolderPath([folder.name]);
      foldersProcessed++;

      // Enumerate existing contents — required to catch MOVED folders whose
      // children don't generate individual delta events.
      const before = filesToProcess.size;
      await collectFilesRecursive(folder.id, [folder.name], filesToProcess);
      const added = filesToProcess.size - before;
      if (added > 0) {
        logger.info('Enumerated existing contents of new folder', {
          folderName: folder.name,
          filesFound: added,
        });
      }
    } catch (folderErr) {
      foldersFailed++;
      logger.error('Failed to process new folder', {
        folderName: folder.name,
        folderId: folder.id,
        error: folderErr.message,
        status: folderErr.response?.status,
        responseBody: folderErr.response?.data,
      });
    }
  }

  // ── Process all files (delta-detected + enumerated) ─────────
  for (const { file, segments } of filesToProcess.values()) {
    const relativePath = segments.join('/') || '(target root)';

    if (uploadedFileIds.has(file.id)) {
      filesSkipped++;
      logger.debug('Skipping already-uploaded file', {
        fileName: file.name,
        fileId: file.id,
        relativePath,
      });
      continue;
    }

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

      uploadedFileIds.add(file.id);
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
    filesSkipped,
    totalFilesSeen: filesToProcess.size,
  });
  return { foldersProcessed, foldersFailed, filesProcessed, filesFailed };
}

// Keep backward-compatible export
async function processNewFolders() {
  const result = await processNewItems();
  return result.foldersProcessed;
}

module.exports = { processNewFolders, processNewItems };

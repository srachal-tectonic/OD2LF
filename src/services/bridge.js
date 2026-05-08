const sharepoint = require('./sharepoint');
const laserfiche = require('./laserfiche');
const config = require('../config');
const logger = require('../logger');
const { summarizeBody } = require('../logger');

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

async function collectFilesRecursive(spFolderId, baseSegments, accumMap, folderSegmentsList) {
  const items = await sharepoint.getFolderContents(spFolderId);
  for (const item of items) {
    if (item.file !== undefined) {
      if (!accumMap.has(item.id)) {
        accumMap.set(item.id, { file: item, segments: baseSegments });
      }
    } else if (item.folder !== undefined) {
      const childSegments = [...baseSegments, item.name];
      // Record the subfolder so we can ensure it in LF even when it (or its
      // entire subtree) contains no files. Without this, empty subfolders
      // never have their path walked by an upload and so never get created.
      folderSegmentsList.push(childSegments);
      await collectFilesRecursive(item.id, childSegments, accumMap, folderSegmentsList);
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

  // Subfolders discovered during recursive enumeration (so we can ensure
  // empty ones in LF; folders with files would otherwise be created as a
  // side-effect of the file-upload path-walk).
  const enumeratedFolderSegments = [];
  let subfoldersEnsured = 0;
  let subfoldersFailed = 0;

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
      const filesBefore = filesToProcess.size;
      const foldersBefore = enumeratedFolderSegments.length;
      await collectFilesRecursive(
        folder.id,
        [folder.name],
        filesToProcess,
        enumeratedFolderSegments
      );
      const filesAdded = filesToProcess.size - filesBefore;
      const foldersAdded = enumeratedFolderSegments.length - foldersBefore;
      if (filesAdded > 0 || foldersAdded > 0) {
        logger.info('Enumerated existing contents of new folder', {
          folderName: folder.name,
          filesFound: filesAdded,
          subfoldersFound: foldersAdded,
        });
      }
    } catch (folderErr) {
      foldersFailed++;
      logger.error('Failed to process new folder', {
        folderName: folder.name,
        folderId: folder.id,
        error: folderErr.message,
        status: folderErr.response?.status,
        responseBody: summarizeBody(folderErr.response?.data),
      });
    }
  }

  // ── Ensure every enumerated subfolder exists in LF ──────────
  // Idempotent: ensureLfFolderPath caches results within this pass and
  // findOrCreateFolder handles existing folders without duplicating them.
  for (const segments of enumeratedFolderSegments) {
    const relativePath = segments.join('/');
    try {
      await ensureLfFolderPath(segments);
      subfoldersEnsured++;
    } catch (subErr) {
      subfoldersFailed++;
      logger.error('Failed to ensure subfolder in Laserfiche', {
        relativePath,
        error: subErr.message,
        status: subErr.response?.status,
        responseBody: summarizeBody(subErr.response?.data),
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
        responseBody: summarizeBody(fileErr.response?.data),
      });
    }
  }

  logger.info('Bridge processing complete', {
    foldersProcessed,
    foldersFailed,
    subfoldersEnsured,
    subfoldersFailed,
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

function saveBase64Picture_(job, data) {
  if (!data.imageBase64) {
    throw new Error('Image absente de la requête.');
  }

  const mimeType = normalizeMimeType_(data.mimeType);
  const extension = getExtensionFromMimeType_(mimeType);
  const displayName = sanitizeFilename_(job.displayName || 'licencie');
  const fileName = `${displayName}_${sanitizeFilename_(job.userId)}.${extension}`;
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const props = PropertiesService.getDocumentProperties();

  deleteDrivePicture_(job.userId);
  removeFilesWithName_(folder, fileName);

  const bytes = Utilities.base64Decode(data.imageBase64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  const driveUrl =
    `https://drive.google.com/uc?export=view&id=${file.getId()}`;

  props.setProperty(
    `${CONFIG.DRIVE_URL_PROPERTY_PREFIX}${job.userId}`,
    driveUrl
  );
  props.setProperty(
    `${CONFIG.DRIVE_FILE_ID_PROPERTY_PREFIX}${job.userId}`,
    file.getId()
  );

  return {
    fileId: file.getId(),
    driveUrl
  };
}

function updatePictureCell_(job, driveUrl) {
  const spreadsheet = SpreadsheetApp.openById(job.spreadsheetId);
  const sheet = spreadsheet
    .getSheets()
    .find((currentSheet) => currentSheet.getSheetId() === Number(job.sheetId));

  if (!sheet) {
    throw new Error(`Onglet introuvable : ${job.sheetId}`);
  }

  sheet
    .getRange(job.pictureCell)
    .setFormula(`=IMAGE("${driveUrl}")`);
}

function deleteDrivePicture_(userId) {
  const props = PropertiesService.getDocumentProperties();
  const fileIdKey = `${CONFIG.DRIVE_FILE_ID_PROPERTY_PREFIX}${userId}`;
  const urlKey = `${CONFIG.DRIVE_URL_PROPERTY_PREFIX}${userId}`;
  let fileId = props.getProperty(fileIdKey);

  if (!fileId) {
    fileId = extractDriveFileId_(props.getProperty(urlKey));
  }

  if (fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (error) {
      console.warn(`Impossible de placer le fichier ${fileId} dans la corbeille`, error);
    }
  }

  props.deleteProperty(fileIdKey);
  props.deleteProperty(urlKey);
  CacheService.getScriptCache().remove(urlKey);
}

function removeFilesWithName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);

  while (files.hasNext()) {
    const file = files.next();

    try {
      file.setTrashed(true);
    } catch (error) {
      console.warn(
        `Impossible de supprimer l'ancien fichier ${file.getId()} (${fileName})`,
        error
      );
    }
  }
}

function extractDriveFileId_(url) {
  if (!url) {
    return null;
  }

  const queryMatch = url.match(/[?&]id=([^&]+)/);
  if (queryMatch) {
    return queryMatch[1];
  }

  const pathMatch = url.match(/\/d\/([^/]+)/);
  return pathMatch ? pathMatch[1] : null;
}

function normalizeMimeType_(mimeType) {
  const allowed = ['image/jpeg', 'image/png', 'image/webp', 'image/gif'];
  return allowed.includes(mimeType) ? mimeType : 'image/jpeg';
}

function getExtensionFromMimeType_(mimeType) {
  const extensions = {
    'image/jpeg': 'jpg',
    'image/png': 'png',
    'image/webp': 'webp',
    'image/gif': 'gif'
  };

  return extensions[mimeType] || CONFIG.DEFAULT_EXTENSION;
}

function sanitizeFilename_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'licencie';
}


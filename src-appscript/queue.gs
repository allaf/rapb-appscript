function ensureQueueSheet_() {
  const spreadsheet = getProjectSpreadsheet_();
  let queueSheet = spreadsheet.getSheetByName(CONFIG.QUEUE_SHEET_NAME);

  if (!queueSheet) {
    queueSheet = spreadsheet.insertSheet(CONFIG.QUEUE_SHEET_NAME);
    queueSheet.getRange(1, 1, 1, JOB_HEADERS.length).setValues([JOB_HEADERS]);
    queueSheet.setFrozenRows(1);
  }

  if (!queueSheet.isSheetHidden()) {
    queueSheet.hideSheet();
  }

  return queueSheet;
}

function getProjectSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty(CONFIG.SPREADSHEET_ID_PROPERTY);

  if (!spreadsheetId) {
    const activeSpreadsheet = SpreadsheetApp.getActive();

    if (!activeSpreadsheet) {
      throw new Error(
        'Google Sheet non configuré. Exécute setupProject() depuis la feuille.'
      );
    }

    spreadsheetId = activeSpreadsheet.getId();
    props.setProperty(CONFIG.SPREADSHEET_ID_PROPERTY, spreadsheetId);
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function createPictureJob_(context) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const queueSheet = ensureQueueSheet_();
    const existing = findOpenJob_(
      queueSheet,
      context.sheetId,
      context.pictureCell.getA1Notation(),
      context.userId
    );

    if (existing) {
      return existing;
    }

    const job = {
      requestId: Utilities.getUuid(),
      spreadsheetId: context.spreadsheetId,
      sheetId: context.sheetId,
      pictureCell: context.pictureCell.getA1Notation(),
      userId: context.userId,
      profileUrl: context.profileUrl,
      displayName: context.displayName,
      status: 'PENDING',
      requestedAt: new Date()
    };

    queueSheet.appendRow([
      job.requestId,
      job.spreadsheetId,
      job.sheetId,
      job.pictureCell,
      job.userId,
      job.profileUrl,
      job.displayName,
      job.status,
      job.requestedAt,
      '',
      '',
      '',
      '',
      ''
    ]);

    return job;
  } finally {
    lock.releaseLock();
  }
}

function findOpenJob_(queueSheet, sheetId, pictureCell, userId) {
  const lastRow = queueSheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const rows = queueSheet
    .getRange(2, 1, lastRow - 1, JOB_HEADERS.length)
    .getValues();

  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const job = rowToJob_(rows[index], index + 2);

    if (
      String(job.sheetId) === String(sheetId) &&
      job.pictureCell === pictureCell &&
      String(job.userId) === String(userId) &&
      ['PENDING', 'PROCESSING'].includes(job.status)
    ) {
      return job;
    }
  }

  return null;
}

function listPendingJobs_() {
  const queueSheet = ensureQueueSheet_();
  const lastRow = queueSheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return queueSheet
    .getRange(2, 1, lastRow - 1, JOB_HEADERS.length)
    .getValues()
    .map((row, index) => rowToJob_(row, index + 2))
    .filter((job) => job.status === 'PENDING')
    .slice(0, CONFIG.MAX_PENDING_JOBS)
    .map(toPublicJob_);
}

function claimJob_(requestId) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const job = findJobByRequestId_(requestId);

    if (!job) {
      throw new Error('Demande introuvable.');
    }

    if (job.status !== 'PENDING') {
      throw new Error(`La demande est déjà ${job.status}.`);
    }

    const queueSheet = ensureQueueSheet_();
    queueSheet.getRange(job.rowNumber, JOB_COLUMNS.STATUS).setValue('PROCESSING');
    queueSheet.getRange(job.rowNumber, JOB_COLUMNS.STARTED_AT).setValue(new Date());

    job.status = 'PROCESSING';
    return toPublicJob_(job);
  } finally {
    lock.releaseLock();
  }
}

function completeJob_(job, fileId, driveUrl) {
  const queueSheet = ensureQueueSheet_();

  queueSheet
    .getRange(job.rowNumber, JOB_COLUMNS.STATUS)
    .setValue('DONE');
  queueSheet
    .getRange(job.rowNumber, JOB_COLUMNS.FINISHED_AT)
    .setValue(new Date());
  queueSheet
    .getRange(job.rowNumber, JOB_COLUMNS.ERROR)
    .clearContent();
  queueSheet
    .getRange(job.rowNumber, JOB_COLUMNS.FILE_ID)
    .setValue(fileId);
  queueSheet
    .getRange(job.rowNumber, JOB_COLUMNS.DRIVE_URL)
    .setValue(driveUrl);
}

function failJob_(requestId, errorMessage) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const job = findJobByRequestId_(requestId);

    if (!job) {
      throw new Error('Demande introuvable.');
    }

    const queueSheet = ensureQueueSheet_();
    queueSheet.getRange(job.rowNumber, JOB_COLUMNS.STATUS).setValue('ERROR');
    queueSheet
      .getRange(job.rowNumber, JOB_COLUMNS.FINISHED_AT)
      .setValue(new Date());
    queueSheet
      .getRange(job.rowNumber, JOB_COLUMNS.ERROR)
      .setValue(String(errorMessage || 'Erreur inconnue').slice(0, 1000));

    return toPublicJob_(job);
  } finally {
    lock.releaseLock();
  }
}

function findJobByRequestId_(requestId) {
  const queueSheet = ensureQueueSheet_();
  const lastRow = queueSheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const requestIds = queueSheet
    .getRange(2, JOB_COLUMNS.REQUEST_ID, lastRow - 1, 1)
    .getDisplayValues();

  for (let index = 0; index < requestIds.length; index += 1) {
    if (requestIds[index][0] === requestId) {
      const rowNumber = index + 2;
      const row = queueSheet
        .getRange(rowNumber, 1, 1, JOB_HEADERS.length)
        .getValues()[0];
      return rowToJob_(row, rowNumber);
    }
  }

  return null;
}

function rowToJob_(row, rowNumber) {
  return {
    rowNumber,
    requestId: String(row[JOB_COLUMNS.REQUEST_ID - 1] || ''),
    spreadsheetId: String(row[JOB_COLUMNS.SPREADSHEET_ID - 1] || ''),
    sheetId: row[JOB_COLUMNS.SHEET_ID - 1],
    pictureCell: String(row[JOB_COLUMNS.PICTURE_CELL - 1] || ''),
    userId: String(row[JOB_COLUMNS.USER_ID - 1] || ''),
    profileUrl: String(row[JOB_COLUMNS.PROFILE_URL - 1] || ''),
    displayName: String(row[JOB_COLUMNS.DISPLAY_NAME - 1] || ''),
    status: String(row[JOB_COLUMNS.STATUS - 1] || ''),
    requestedAt: row[JOB_COLUMNS.REQUESTED_AT - 1],
    startedAt: row[JOB_COLUMNS.STARTED_AT - 1],
    finishedAt: row[JOB_COLUMNS.FINISHED_AT - 1],
    error: String(row[JOB_COLUMNS.ERROR - 1] || ''),
    fileId: String(row[JOB_COLUMNS.FILE_ID - 1] || ''),
    driveUrl: String(row[JOB_COLUMNS.DRIVE_URL - 1] || '')
  };
}

function toPublicJob_(job) {
  return {
    requestId: job.requestId,
    spreadsheetId: job.spreadsheetId,
    sheetId: job.sheetId,
    pictureCell: job.pictureCell,
    userId: job.userId,
    profileUrl: job.profileUrl,
    displayName: job.displayName,
    status: job.status
  };
}

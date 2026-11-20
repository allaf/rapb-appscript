const PALMARES_JOB_COLUMNS = Object.freeze({
  REQUEST_ID: 1,
  SPREADSHEET_ID: 2,
  SHEET_ID: 3,
  TARGET_CELL: 4,
  USER_ID: 5,
  PROFILE_URL: 6,
  DISPLAY_NAME: 7,
  STATUS: 8,
  REQUESTED_AT: 9,
  STARTED_AT: 10,
  FINISHED_AT: 11,
  ERROR: 12
});

const PALMARES_JOB_HEADERS = Object.freeze([
  'requestId',
  'spreadsheetId',
  'sheetId',
  'targetCell',
  'userId',
  'profileUrl',
  'displayName',
  'status',
  'requestedAt',
  'startedAt',
  'finishedAt',
  'error'
]);

function ensurePalmaresQueueSheet_() {
  const spreadsheet = getProjectSpreadsheet_();
  let queueSheet = spreadsheet.getSheetByName(CONFIG.PALMARES_QUEUE_SHEET_NAME);

  if (!queueSheet) {
    queueSheet = spreadsheet.insertSheet(CONFIG.PALMARES_QUEUE_SHEET_NAME);
    queueSheet
      .getRange(1, 1, 1, PALMARES_JOB_HEADERS.length)
      .setValues([PALMARES_JOB_HEADERS]);
    queueSheet.setFrozenRows(1);
  }

  if (!queueSheet.isSheetHidden()) {
    queueSheet.hideSheet();
  }

  return queueSheet;
}

function requestPalmaresForSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActive();
    const sheet = spreadsheet.getActiveSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < 5) {
      throw new Error('Aucun boxeur à traiter dans cet onglet.');
    }

    const contexts = [];
    const seenUsers = new Set();

    for (let row = 5; row <= lastRow; row += 1) {
      const linkCell = sheet.getRange(row, 2);
      const profileUrl = getUrlFromCell_(linkCell);

      if (!profileUrl) {
        continue;
      }

      const match = profileUrl.match(CONFIG.PROFILE_URL_PATTERN);

      if (!match) {
        continue;
      }

      const userId = decodeURIComponent(match[1]);

      if (seenUsers.has(userId)) {
        continue;
      }

      seenUsers.add(userId);

      contexts.push({
        spreadsheetId: spreadsheet.getId(),
        sheetId: sheet.getSheetId(),
        targetCell: sheet.getRange(row, 7).getA1Notation(),
        userId,
        profileUrl,
        displayName: `${sheet.getRange(row, 2).getDisplayValue().trim()} ${sheet.getRange(row, 3).getDisplayValue().trim()}`.trim()
      });
    }

    if (contexts.length === 0) {
      throw new Error(
        'Aucun lien vers une fiche licencié FFBoxe trouvé en colonne B.'
      );
    }

    const jobs = contexts.map((context) => createPalmaresJob_(context));

    SpreadsheetApp.getActive().toast(
      jobs.length === 1
        ? '1 demande de palmarès créée.'
        : `${jobs.length} demandes de palmarès créées.`,
      'Palmarès'
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(error.message);
  }
}

function createPalmaresJob_(context) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const queueSheet = ensurePalmaresQueueSheet_();
    const existing = findOpenPalmaresJob_(
      queueSheet,
      context.sheetId,
      context.targetCell,
      context.userId
    );

    if (existing) {
      return existing;
    }

    const job = {
      requestId: Utilities.getUuid(),
      spreadsheetId: context.spreadsheetId,
      sheetId: context.sheetId,
      targetCell: context.targetCell,
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
      job.targetCell,
      job.userId,
      job.profileUrl,
      job.displayName,
      job.status,
      job.requestedAt,
      '',
      '',
      ''
    ]);

    return job;
  } finally {
    lock.releaseLock();
  }
}

function findOpenPalmaresJob_(
  queueSheet,
  sheetId,
  targetCell,
  userId
) {
  const lastRow = queueSheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const rows = queueSheet
    .getRange(
      2,
      1,
      lastRow - 1,
      PALMARES_JOB_HEADERS.length
    )
    .getValues();

  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const job = rowToPalmaresJob_(rows[index], index + 2);

    if (
      String(job.sheetId) === String(sheetId) &&
      job.targetCell === targetCell &&
      String(job.userId) === String(userId) &&
      ['PENDING', 'PROCESSING'].includes(job.status)
    ) {
      return job;
    }
  }

  return null;
}

function listPendingPalmaresJobs_() {
  const queueSheet = ensurePalmaresQueueSheet_();
  const lastRow = queueSheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return queueSheet
    .getRange(
      2,
      1,
      lastRow - 1,
      PALMARES_JOB_HEADERS.length
    )
    .getValues()
    .map((row, index) => rowToPalmaresJob_(row, index + 2))
    .filter((job) => job.status === 'PENDING')
    .slice(0, CONFIG.PALMARES_MAX_PENDING_JOBS)
    .map(toPublicPalmaresJob_);
}

function claimPalmaresJob_(requestId) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const job = findPalmaresJobByRequestId_(requestId);

    if (!job) {
      throw new Error('Demande de palmarès introuvable.');
    }

    if (job.status !== 'PENDING') {
      throw new Error(`La demande est déjà ${job.status}.`);
    }

    const queueSheet = ensurePalmaresQueueSheet_();
    queueSheet
      .getRange(job.rowNumber, PALMARES_JOB_COLUMNS.STATUS)
      .setValue('PROCESSING');
    queueSheet
      .getRange(job.rowNumber, PALMARES_JOB_COLUMNS.STARTED_AT)
      .setValue(new Date());

    job.status = 'PROCESSING';

    return toPublicPalmaresJob_(job);
  } finally {
    lock.releaseLock();
  }
}

function completePalmaresJob_(job, palmares) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const spreadsheet = SpreadsheetApp.openById(job.spreadsheetId);
    const sheet = spreadsheet.getSheetById(Number(job.sheetId));

    if (!sheet) {
      throw new Error('Onglet cible introuvable.');
    }

    sheet.getRange(job.targetCell).setValue(String(palmares || ''));

    const queueSheet = ensurePalmaresQueueSheet_();

    queueSheet
      .getRange(job.rowNumber, PALMARES_JOB_COLUMNS.STATUS)
      .setValue('DONE');
    queueSheet
      .getRange(job.rowNumber, PALMARES_JOB_COLUMNS.FINISHED_AT)
      .setValue(new Date());
    queueSheet
      .getRange(job.rowNumber, PALMARES_JOB_COLUMNS.ERROR)
      .clearContent();
  } finally {
    lock.releaseLock();
  }
}

function failPalmaresJob_(requestId, errorMessage) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const job = findPalmaresJobByRequestId_(requestId);

    if (!job) {
      throw new Error('Demande de palmarès introuvable.');
    }

    const queueSheet = ensurePalmaresQueueSheet_();

    queueSheet
      .getRange(job.rowNumber, PALMARES_JOB_COLUMNS.STATUS)
      .setValue('ERROR');
    queueSheet
      .getRange(job.rowNumber, PALMARES_JOB_COLUMNS.FINISHED_AT)
      .setValue(new Date());
    queueSheet
      .getRange(job.rowNumber, PALMARES_JOB_COLUMNS.ERROR)
      .setValue(
        String(errorMessage || 'Erreur inconnue').slice(0, 1000)
      );

    return toPublicPalmaresJob_(job);
  } finally {
    lock.releaseLock();
  }
}

function findPalmaresJobByRequestId_(requestId) {
  const queueSheet = ensurePalmaresQueueSheet_();
  const lastRow = queueSheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const requestIds = queueSheet
    .getRange(
      2,
      PALMARES_JOB_COLUMNS.REQUEST_ID,
      lastRow - 1,
      1
    )
    .getDisplayValues();

  for (let index = 0; index < requestIds.length; index += 1) {
    if (requestIds[index][0] === requestId) {
      const rowNumber = index + 2;
      const row = queueSheet
        .getRange(
          rowNumber,
          1,
          1,
          PALMARES_JOB_HEADERS.length
        )
        .getValues()[0];

      return rowToPalmaresJob_(row, rowNumber);
    }
  }

  return null;
}

function rowToPalmaresJob_(row, rowNumber) {
  return {
    rowNumber,
    requestId: String(
      row[PALMARES_JOB_COLUMNS.REQUEST_ID - 1] || ''
    ),
    spreadsheetId: String(
      row[PALMARES_JOB_COLUMNS.SPREADSHEET_ID - 1] || ''
    ),
    sheetId: row[PALMARES_JOB_COLUMNS.SHEET_ID - 1],
    targetCell: String(
      row[PALMARES_JOB_COLUMNS.TARGET_CELL - 1] || ''
    ),
    userId: String(
      row[PALMARES_JOB_COLUMNS.USER_ID - 1] || ''
    ),
    profileUrl: String(
      row[PALMARES_JOB_COLUMNS.PROFILE_URL - 1] || ''
    ),
    displayName: String(
      row[PALMARES_JOB_COLUMNS.DISPLAY_NAME - 1] || ''
    ),
    status: String(
      row[PALMARES_JOB_COLUMNS.STATUS - 1] || ''
    ),
    requestedAt:
      row[PALMARES_JOB_COLUMNS.REQUESTED_AT - 1],
    startedAt:
      row[PALMARES_JOB_COLUMNS.STARTED_AT - 1],
    finishedAt:
      row[PALMARES_JOB_COLUMNS.FINISHED_AT - 1],
    error: String(
      row[PALMARES_JOB_COLUMNS.ERROR - 1] || ''
    )
  };
}

function toPublicPalmaresJob_(job) {
  return {
    requestId: job.requestId,
    spreadsheetId: job.spreadsheetId,
    sheetId: job.sheetId,
    targetCell: job.targetCell,
    userId: job.userId,
    profileUrl: job.profileUrl,
    displayName: job.displayName,
    status: job.status
  };
}

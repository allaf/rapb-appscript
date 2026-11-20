const CONFIG = Object.freeze({
  MENU_NAME: 'Photos licenciés',
  QUEUE_SHEET_NAME: '_PHOTO_JOBS',
  DRIVE_FOLDER_ID: '1W_HF_Hx_XH43jAetyG9SplIRMMsCHRyY',
  API_TOKEN_PROPERTY: 'FFBOXE_API_TOKEN',
  SPREADSHEET_ID_PROPERTY: 'FFBOXE_SPREADSHEET_ID',
  DRIVE_URL_PROPERTY_PREFIX: 'drive_',
  DRIVE_FILE_ID_PROPERTY_PREFIX: 'drive_file_',
  MAX_PENDING_JOBS: 10,
  PROFILE_URL_PATTERN: /\/personnes\/fiche\/([^/?#]+)(?:[/?#]|$)/,
  DEFAULT_EXTENSION: 'jpg',
  WEB_APP_URL_PROPERTY: 'https://script.google.com/macros/s/AKfycbztm8-iiPORObX7PsdW8CdVezGu3ub7bTEI1AqxppDh1C_vxo6XxWC5VH6Ab281TaXC/exec',
});

const JOB_COLUMNS = Object.freeze({
  REQUEST_ID: 1,
  SPREADSHEET_ID: 2,
  SHEET_ID: 3,
  PICTURE_CELL: 4,
  USER_ID: 5,
  PROFILE_URL: 6,
  DISPLAY_NAME: 7,
  STATUS: 8,
  REQUESTED_AT: 9,
  STARTED_AT: 10,
  FINISHED_AT: 11,
  ERROR: 12,
  FILE_ID: 13,
  DRIVE_URL: 14
});

const JOB_HEADERS = Object.freeze([
  'requestId',
  'spreadsheetId',
  'sheetId',
  'pictureCell',
  'userId',
  'profileUrl',
  'displayName',
  'status',
  'requestedAt',
  'startedAt',
  'finishedAt',
  'error',
  'fileId',
  'driveUrl'
]);

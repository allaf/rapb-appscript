const FFB_REGISTRATION_PROPERTY = 'FFB_PENDING_REGISTRATION';

/**
 * Prépare l'inscription FFB à partir de la ligne actuellement sélectionnée.
 *
 * Cette première version est volontairement limitée au tableau ENFANTS.
 */
function prepareFfbRegistration() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range) {
    throw new Error('Sélectionne une cellule sur la ligne de l’enfant.');
  }

  const row = range.getRow();

  const config = REGISTRATION_IMPORT.SECTIONS.ENFANTS;
  const section = findSection_(sheet, config.title);

  if (row <= section.headerRow || row > section.endRow) {
    throw new Error(
      'La ligne sélectionnée ne se trouve pas dans le tableau ENFANTS.'
    );
  }

  const columns = config.columns;

  const firstName = getCellText_(
    sheet,
    row,
    columns.firstName
  );

  const lastName = getCellText_(
    sheet,
    row,
    columns.lastName
  );

  const email = getCellText_(
    sheet,
    row,
    columns.email
  );

  const phone = getCellText_(
    sheet,
    row,
    columns.phone
  );

  const birthDate = getBirthDateText_(
    sheet,
    row,
    columns.birthDate
  );

  const parentFirstName = getCellText_(
    sheet,
    row,
    columns.parentFirstName
  );

  let parentLastName = getCellText_(
    sheet,
    row,
    columns.parentLastName
  );

  /*
   * Si le nom du parent n'est pas renseigné,
   * on reprend celui de l'enfant.
   */
  if (!parentLastName) {
    parentLastName = lastName;
  }

  if (!firstName) {
    throw new Error('Le prénom de l’enfant est vide.');
  }

  if (!lastName) {
    throw new Error('Le nom de l’enfant est vide.');
  }

  if (!birthDate) {
    throw new Error('La date de naissance est vide.');
  }

  if (!email) {
    throw new Error('L’adresse e-mail est vide.');
  }

  if (!phone) {
    throw new Error('Le téléphone est vide.');
  }

  if (!parentFirstName) {
    throw new Error('Le prénom du parent est vide.');
  }

  const registration = {
    type: 'ENFANT',

    firstName,
    lastName: lastName.toUpperCase(),

    birthDate,

    email,
    phone: digitsOnly_(phone),

    parentFirstName,
    parentLastName: parentLastName.toUpperCase(),

    postalCode: '34000',
    city: 'MONTPELLIER',

    sheetId: sheet.getSheetId(),
    sheetName: sheet.getName(),
    row,

    createdAt: new Date().toISOString()
  };

  PropertiesService
    .getScriptProperties()
    .setProperty(
      FFB_REGISTRATION_PROPERTY,
      JSON.stringify(registration)
    );


  const preview = [
    `Type : ${registration.type}`,
    `Prénom : ${registration.firstName}`,
    `Nom : ${registration.lastName}`,
    `Date de naissance : ${registration.birthDate}`,
    `Email : ${registration.email}`,
    `Téléphone : ${registration.phone}`,
    '',
    `Prénom parent : ${registration.parentFirstName}`,
    `Nom parent : ${registration.parentLastName}`,
    '',
    `Code postal : ${registration.postalCode}`,
    `Ville : ${registration.city}`,
    '',
    `Onglet : ${registration.sheetName}`,
    `Ligne : ${registration.row}`
  ].join('\n');

  SpreadsheetApp.getUi().alert(
    'Inscription FFB préparée',
    preview,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function getPendingFfbRegistration_() {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty(FFB_REGISTRATION_PROPERTY);

  if (!value) {
    return null;
  }

  return JSON.parse(value);
}


function getCellText_(sheet, row, column) {
  if (!column) {
    return '';
  }

  return sheet
    .getRange(row, column)
    .getDisplayValue()
    .trim();
}


/**
 * Renvoie toujours la date au format JJ/MM/AAAA.
 */
function getBirthDateText_(sheet, row, column) {
  const cell = sheet.getRange(row, column);
  const value = cell.getValue();

  if (
    value instanceof Date &&
    !Number.isNaN(value.getTime())
  ) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );
  }

  const text = cell.getDisplayValue().trim();

  const match = text.match(
    /(\d{1,2})\/(\d{1,2})\/(\d{4})/
  );

  if (!match) {
    return text;
  }

  return [
    match[1].padStart(2, '0'),
    match[2].padStart(2, '0'),
    match[3]
  ].join('/');
}
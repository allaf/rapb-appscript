const REGISTRATION_IMPORT = {
  SECTIONS: {
    LOISIRS: {
      title: 'LOISIRS',
      columns: {
        amount: 1,
        firstName: 4,
        lastName: 5,
        email: 7,
        phone: 8,
        birthDate: 9,
        remark: 10,
        whatsapp: 11
      }
    },

    AMATEURS: {
      title: 'AMATEURS',
      columns: {
        amount: 1,
        firstName: 4,
        lastName: 5,
        email: 7,
        phone: 8,
        birthDate: 9,
        remark: 10,
        whatsapp: 11
      }
    },

    ENFANTS: {
      title: 'ENFANTS',
      columns: {
        amount: 1,
        firstName: 4,
        lastName: 5,
        email: 7,
        phone: 8,
        remark: 9,
        parentFirstName: 10,
        parentLastName: 11,
        birthDate: 12
      }
    }
  }
};

function showPdfImportDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const template = HtmlService.createTemplateFromFile(
    'ImportPdfDialog'
  );

  template.sheetId = sheet.getSheetId();
  template.sheetName = sheet.getName();

  const html = template
    .evaluate()
    .setWidth(480)
    .setHeight(310);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Importer une inscription'
  );
}

function importRegistrationPdf(payload) {
  try {
    if (!payload || !payload.base64) {
      throw new Error('Aucun fichier PDF reçu.');
    }

    const fileName = String(
      payload.fileName || 'inscription.pdf'
    );

    if (!fileName.toLowerCase().endsWith('.pdf')) {
      throw new Error(
        'Le fichier sélectionné doit être un PDF.'
      );
    }

    const bytes = Utilities.base64Decode(
      payload.base64
    );

    if (bytes.length > 10 * 1024 * 1024) {
      throw new Error(
        'Le PDF dépasse la taille maximale de 10 Mo.'
      );
    }

    const pdfBlob = Utilities.newBlob(
      bytes,
      'application/pdf',
      fileName
    );

    const sheetId = Number(payload.sheetId);
    const spreadsheet = SpreadsheetApp.getActive();

    const sheet = spreadsheet
      .getSheets()
      .find(
        (candidate) =>
          candidate.getSheetId() === sheetId
      );

    if (!sheet) {
      throw new Error(
        'L’onglet sélectionné est introuvable.'
      );
    }

    const text = extractTextFromPdf_(pdfBlob);
    const registration = parseRegistrationPdf_(text);
    const result = writeRegistration_(
      sheet,
      registration
    );

    return {
      firstName: registration.firstName,
      lastName: registration.lastName,
      section: registration.section,
      row: result.row,
      created: result.created,
      paymentMethod: registration.paymentMethod,
      amount:
        registration.paymentMethod === 'CB'
          ? registration.amount
          : null
    };
  } catch (error) {
    console.error(error.stack || error.message);
    throw new Error(
      `Import impossible : ${error.message}`
    );
  }
}

function authorizePdfImport() {
  const temporaryDocument = DocumentApp.create(
    `AUTORISATION_IMPORT_PDF_${Date.now()}`
  );

  const documentId = temporaryDocument.getId();

  temporaryDocument
    .getBody()
    .appendParagraph('Document temporaire d’autorisation');

  temporaryDocument.saveAndClose();

  DriveApp
    .getFileById(documentId)
    .setTrashed(true);

  SpreadsheetApp.getUi().alert(
    'Autorisation accordée',
    'L’import des PDF est maintenant autorisé.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function extractTextFromPdf_(pdfBlob) {
  let temporaryDocumentId = null;

  try {
    const document = Drive.Files.create(
      {
        name: `IMPORT_TEMP_${Date.now()}`,
        mimeType: 'application/vnd.google-apps.document'
      },
      pdfBlob,
      {
        fields: 'id'
      }
    );

    temporaryDocumentId = document.id;

    for (let attempt = 1; attempt <= 5; attempt += 1) {
      try {
        const text = DocumentApp
          .openById(temporaryDocumentId)
          .getBody()
          .getText();

        if (text && text.trim()) {
          return text;
        }
      } catch (error) {
        if (attempt === 5) {
          throw error;
        }
      }

      Utilities.sleep(500);
    }

    throw new Error(
      'Aucun texte n’a pu être extrait du PDF.'
    );
  } finally {
    if (temporaryDocumentId) {
      try {
        Drive.Files.remove(temporaryDocumentId);
      } catch (error) {
        console.error(
          `Document temporaire non supprimé : ${error.message}`
        );
      }
    }
  }
}

function parseRegistrationPdf_(text) {
  const normalizedText = normalizeText_(text);

  let section;

  if (normalizedText.includes('boxe enfant/ado')) {
    section = 'ENFANTS';
  } else if (
    normalizedText.includes('boxe competiteur/amateur')
  ) {
    section = 'AMATEURS';
  } else if (
    normalizedText.includes('boxe loisir adulte')
  ) {
    section = 'LOISIRS';
  } else {
    throw new Error(
      'Le type d’inscription n’a pas été reconnu.'
    );
  }

  const paymentMethod = extractField_(
    text,
    ['Mode de paiement']
  );

  const registration = {
    section,
    firstName: extractRequiredField_(text, ['Prénom']),
    lastName: extractRequiredField_(text, ['Nom']),
    birthDate: extractRequiredField_(
      text,
      ['Date de naissance']
    ),
    phone: extractField_(
      text,
      ['Téléphone portable']
    ),
    email: extractRequiredField_(
      text,
      ['Email pratiquant', 'email parent', 'Email']
    ),
    paymentMethod: normalizePaymentMethod_(paymentMethod),
    amount: extractAmount_(text),
    whatsapp: hasAcceptedWhatsapp_(text),
    parentFirstName: '',
    parentLastName: ''
  };

  if (section === 'ENFANTS') {
    registration.parentFirstName = extractField_(
      text,
      ['prenom du parent', 'prénom du parent']
    );

    registration.parentLastName = extractField_(
      text,
      ['nom du parent']
    );
  }

  return registration;
}

function writeRegistration_(sheet, registration) {
  const config =
    REGISTRATION_IMPORT.SECTIONS[registration.section];

  const section = findSection_(sheet, config.title);
  const existingRow = findExistingRegistration_(
    sheet,
    section,
    config.columns,
    registration
  );

  const row =
    existingRow ||
    findFirstEmptyRow_(
      sheet,
      section,
      config.columns
    );

  if (!row) {
    throw new Error(
      `Aucune ligne vide disponible dans le tableau ${config.title}.`
    );
  }

  const created = !existingRow;
  const columns = config.columns;

  setIfBlank_(
    sheet,
    row,
    columns.firstName,
    registration.firstName
  );

  setIfBlank_(
    sheet,
    row,
    columns.lastName,
    registration.lastName.toUpperCase()
  );

  setIfBlank_(
    sheet,
    row,
    columns.email,
    registration.email
  );

  setPhoneIfBlank_(
    sheet,
    row,
    columns.phone,
    registration.phone
  );

  setDateIfBlank_(
    sheet,
    row,
    columns.birthDate,
    registration.birthDate
  );

  if (columns.parentFirstName) {
    setIfBlank_(
      sheet,
      row,
      columns.parentFirstName,
      registration.parentFirstName
    );
  }

  if (columns.parentLastName) {
    setIfBlank_(
      sheet,
      row,
      columns.parentLastName,
      registration.parentLastName.toUpperCase()
    );
  }

  /*
   * La cotisation est renseignée uniquement
   * lorsque le paiement est effectué par CB.
   */
  if (registration.paymentMethod === 'CB') {
    sheet
      .getRange(row, columns.amount)
      .setValue(registration.amount);
  }

  if (
    registration.whatsapp &&
    columns.whatsapp
  ) {
    setIfBlank_(
      sheet,
      row,
      columns.whatsapp,
      'Oui'
    );
  }

  updateImportRemark_(
    sheet,
    row,
    columns.remark,
    registration.paymentMethod,
    created
  );

  return {
    row,
    created
  };
}

function findSection_(sheet, sectionTitle) {
  const values = sheet.getDataRange().getDisplayValues();
  const normalizedTitle = normalizeText_(sectionTitle);

  for (let row = 0; row < values.length; row += 1) {
    const containsTitle = values[row].some(
      (value) => normalizeText_(value) === normalizedTitle
    );

    if (!containsTitle) {
      continue;
    }

    /*
     * Recherche de la ligne d’en-tête située
     * juste sous le titre du tableau.
     */
    const lastCandidate = Math.min(
      row + 6,
      values.length
    );

    for (
      let headerIndex = row + 1;
      headerIndex < lastCandidate;
      headerIndex += 1
    ) {
      const firstNameHeader = normalizeText_(
        values[headerIndex][3] || ''
      );

      const lastNameHeader = normalizeText_(
        values[headerIndex][4] || ''
      );

      if (
        firstNameHeader === 'prenom' &&
        lastNameHeader === 'nom'
      ) {
        return {
          titleRow: row + 1,
          headerRow: headerIndex + 1,
          endRow: findSectionEndRow_(
            values,
            headerIndex + 1,
            sectionTitle
          )
        };
      }
    }
  }

  throw new Error(
    `Le tableau ${sectionTitle} est introuvable dans l’onglet actif.`
  );
}

function findSectionEndRow_(
  values,
  headerRow,
  currentSectionTitle
) {
  const otherTitles = Object.values(
    REGISTRATION_IMPORT.SECTIONS
  )
    .map((section) => section.title)
    .filter((title) => title !== currentSectionTitle)
    .map(normalizeText_);

  for (
    let rowIndex = headerRow;
    rowIndex < values.length;
    rowIndex += 1
  ) {
    const containsAnotherTitle = values[rowIndex].some(
      (value) =>
        otherTitles.includes(normalizeText_(value))
    );

    if (containsAnotherTitle) {
      return rowIndex;
    }
  }

  return values.length;
}

function findExistingRegistration_(
  sheet,
  section,
  columns,
  registration
) {
  const startRow = section.headerRow + 1;
  const numberOfRows = section.endRow - startRow + 1;

  if (numberOfRows <= 0) {
    return null;
  }

  const values = sheet
    .getRange(startRow, 1, numberOfRows, sheet.getLastColumn())
    .getDisplayValues();

  const searchedEmail = normalizeText_(registration.email);
  const searchedPhone = digitsOnly_(registration.phone);
  const searchedFirstName = normalizeText_(
    registration.firstName
  );
  const searchedLastName = normalizeText_(
    registration.lastName
  );
  const searchedBirthDate = dateKey_(
    registration.birthDate
  );

  for (let index = 0; index < values.length; index += 1) {
    const row = values[index];

    const email = normalizeText_(
      row[columns.email - 1] || ''
    );

    const phone = digitsOnly_(
      row[columns.phone - 1] || ''
    );

    const firstName = normalizeText_(
      row[columns.firstName - 1] || ''
    );

    const lastName = normalizeText_(
      row[columns.lastName - 1] || ''
    );

    const birthDate = dateKey_(
      row[columns.birthDate - 1] || ''
    );

    const sameEmail =
      searchedEmail && email === searchedEmail;

    const samePhone =
      searchedPhone && phone === searchedPhone;

    const sameIdentity =
      firstName === searchedFirstName &&
      lastName === searchedLastName &&
      birthDate === searchedBirthDate;

    if (sameEmail || samePhone || sameIdentity) {
      return startRow + index;
    }
  }

  return null;
}

function findFirstEmptyRow_(sheet, section, columns) {
  const startRow = section.headerRow + 1;

  for (
    let row = startRow;
    row <= section.endRow;
    row += 1
  ) {
    const firstName = sheet
      .getRange(row, columns.firstName)
      .getDisplayValue()
      .trim();

    const lastName = sheet
      .getRange(row, columns.lastName)
      .getDisplayValue()
      .trim();

    const email = sheet
      .getRange(row, columns.email)
      .getDisplayValue()
      .trim();

    if (!firstName && !lastName && !email) {
      return row;
    }
  }

  return null;
}

function updateImportRemark_(
  sheet,
  row,
  column,
  paymentMethod,
  created
) {
  const cell = sheet.getRange(row, column);
  const currentValue = cell.getDisplayValue().trim();
  const additions = [];

  if (
    paymentMethod === 'CB' &&
    !/\bCB\b/i.test(currentValue)
  ) {
    additions.push('CB');
  }

  if (
    created &&
    !currentValue
      .toLowerCase()
      .includes('ligne créée par le script')
  ) {
    additions.push('Ligne créée par le script');
  }

  if (additions.length === 0) {
    return;
  }

  cell.setValue(
    [currentValue, ...additions]
      .filter(Boolean)
      .join(' - ')
  );
}

function setIfBlank_(sheet, row, column, value) {
  if (!column || value === '' || value == null) {
    return;
  }

  const cell = sheet.getRange(row, column);

  if (!cell.getDisplayValue().trim()) {
    cell.setValue(value);
  }
}


function setPhoneIfBlank_(sheet, row, column, phone) {
  const cell = sheet.getRange(row, column);

  if (cell.getDisplayValue().trim()) {
    return;
  }

  /*
   * On envoie une chaîne afin de conserver
   * le zéro initial du numéro.
   */
  cell.setValue(digitsOnly_(phone));
}


function setDateIfBlank_(sheet, row, column, dateText) {
  const cell = sheet.getRange(row, column);

  if (cell.getDisplayValue().trim()) {
    return;
  }

  const match = String(dateText).match(
    /(\d{2})\/(\d{2})\/(\d{4})/
  );

  if (!match) {
    cell.setValue(dateText);
    return;
  }

  const date = new Date(
    Number(match[3]),
    Number(match[2]) - 1,
    Number(match[1])
  );

  cell.setValue(date);
}


function extractRequiredField_(text, labels) {
  const value = extractField_(text, labels);

  if (!value) {
    throw new Error(
      `Champ introuvable dans le PDF : ${labels[0]}`
    );
  }

  return value;
}

function extractField_(text, labels) {
  for (const label of labels) {
    const escapedLabel = label.replace(
      /[.*+?^${}()|[\]\\]/g,
      '\\$&'
    );

    const pattern = new RegExp(
      `(?:^|\\n)\\s*${escapedLabel}\\s*:\\s*([^\\n]+)`,
      'i'
    );

    const match = text.match(pattern);

    if (match) {
      return match[1].trim();
    }
  }

  return '';
}

function extractAmount_(text) {
  const match = text.match(
    /somme de\s*:\s*\*+\s*([\d\s]+(?:[,.]\d{2})?)\s*Euros/i
  );

  if (!match) {
    throw new Error(
      'Le montant de la cotisation est introuvable.'
    );
  }

  return Number(
    match[1]
      .replace(/\s/g, '')
      .replace(',', '.')
  );
}

function hasAcceptedWhatsapp_(text) {
  return /M['’]ajouter[^\n]*:\s*Oui/i.test(text);
}

function normalizePaymentMethod_(paymentMethod) {
  const normalized = normalizeText_(paymentMethod);

  if (normalized === 'cb') {
    return 'CB';
  }

  if (normalized.includes('espece')) {
    return 'Espèces';
  }

  if (normalized.includes('cheque')) {
    return 'Chèque';
  }

  return paymentMethod.trim();
}

function normalizeText_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function digitsOnly_(value) {
  return String(value || '').replace(/\D/g, '');
}

function dateKey_(value) {
  const match = String(value || '').match(
    /(\d{1,2})\/(\d{1,2})\/(\d{4})/
  );

  if (!match) {
    return normalizeText_(value);
  }

  return [
    match[1].padStart(2, '0'),
    match[2].padStart(2, '0'),
    match[3]
  ].join('/');
}

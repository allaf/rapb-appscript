function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu(CONFIG.MENU_NAME)
    .addItem(
      'Récupérer ou actualiser la photo',
      'requestSelectedUserPicture'
    )
    .addSeparator()
    .addItem('Supprimer la photo', 'deleteSelectedUserPicture')
    .addSeparator()
    .addItem(
      'Configurer l’URL de déploiement',
      'configureWebAppUrl'
    )
    .addItem(
      'Afficher la configuration API',
      'showApiSettings'
    )
    .addToUi();

  ui.createMenu('Inscriptions')
    .addItem('Importer un PDF', 'showPdfImportDialog')
    .addItem('Préparer inscription FFB', 'prepareFfbRegistration'   )
    .addToUi();
  
  ui.createMenu('Palmarès')
    .addItem('Récupérer les palmarès', 'requestPalmaresForSheet')
    .addToUi();

}

function configureWebAppUrl() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    'Configurer l’URL de déploiement',
    'Colle l’URL Apps Script terminant par /exec :',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const webAppUrl = result.getResponseText().trim();

  if (
    !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(
      webAppUrl
    )
  ) {
    ui.alert(
      'Cette URL n’est pas reconnue. Elle doit commencer par ' +
        'https://script.google.com/macros/s/ et terminer par /exec.'
    );
    return;
  }

  PropertiesService.getScriptProperties().setProperty(
    CONFIG.WEB_APP_URL_PROPERTY,
    webAppUrl
  );

  ui.alert('URL de déploiement enregistrée.');
}

function setupProject() {
  PropertiesService.getScriptProperties().setProperty(
    CONFIG.SPREADSHEET_ID_PROPERTY,
    SpreadsheetApp.getActive().getId()
  );
  ensureQueueSheet_();
  const token = getOrCreateApiToken_();

  SpreadsheetApp.getUi().alert(
    'Configuration terminée',
    'La file de demandes a été créée et masquée.\n\n' +
      'Jeton à renseigner dans l’extension :\n' +
      token,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showApiSettings() {
  const token = getOrCreateApiToken_();

  const webAppUrl = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.WEB_APP_URL_PROPERTY);

  SpreadsheetApp.getUi().alert(
    'Configuration API',
    'URL Web App :\n' +
      (webAppUrl ||
        'Utilise le menu « Configurer l’URL de déploiement ».') +
      '\n\nJeton :\n' +
      token,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function requestSelectedUserPicture() {
  try {
    const contexts = getSelectedLicenceContexts_();
    const jobs = contexts.map((context) => createPictureJob_(context));

    SpreadsheetApp.getActive().toast(
      jobs.length === 1
        ? `Demande créée pour le licencié ${jobs[0].userId}`
        : `${jobs.length} demandes de photos créées`,
      CONFIG.MENU_NAME
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(error.message);
  }
}

function deleteSelectedUserPicture() {
  try {
    const context = getSelectedLicenceContext_();
    const ui = SpreadsheetApp.getUi();
    const confirmation = ui.alert(
      'Supprimer la photo',
      `Supprimer la photo du licencié ${context.userId} ?`,
      ui.ButtonSet.YES_NO
    );

    if (confirmation !== ui.Button.YES) {
      return;
    }

    deleteDrivePicture_(context.userId);
    context.pictureCell.clearContent();

    SpreadsheetApp.getActive().toast(
      `Photo du licencié ${context.userId} supprimée`,
      CONFIG.MENU_NAME
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(error.message);
  }
}

function getSelectedLicenceContexts_() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const range = sheet.getActiveRange();
  const contexts = [];
  const seenUsers = new Set();

  for (let row = 1; row <= range.getNumRows(); row += 1) {
    for (let column = 1; column <= range.getNumColumns(); column += 1) {
      const linkCell = range.getCell(row, column);
      const profileUrl = getUrlFromCell_(linkCell);

      if (!profileUrl) {
        continue;
      }

      const match = profileUrl.match(CONFIG.PROFILE_URL_PATTERN);

      if (!match) {
        continue;
      }

      const userId = decodeURIComponent(match[1]);
      const key = `${sheet.getSheetId()}:${linkCell.getA1Notation()}`;

      if (seenUsers.has(key)) {
        continue;
      }

      if (linkCell.getColumn() <= 1) {
        throw new Error(
          `La cellule ${linkCell.getA1Notation()} ne possède aucune cellule à gauche.`
        );
      }

      seenUsers.add(key);

      contexts.push({
        spreadsheetId: spreadsheet.getId(),
        sheetId: sheet.getSheetId(),
        linkCell: linkCell.getA1Notation(),
        pictureCell: linkCell.offset(0, -1),
        userId,
        profileUrl,
        displayName: linkCell.getDisplayValue().trim()
      });
    }
  }

  if (contexts.length === 0) {
    throw new Error(
      'La sélection ne contient aucun lien vers une fiche licencié FFBoxe.'
    );
  }

  return contexts;
}

function getSelectedLicenceContext_() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const linkCell = sheet.getActiveCell();

  const row = linkCell.getRow();
  const col = linkCell.getColumn();
  const profileUrl = getUrlFromCell_(linkCell);

  if (!profileUrl) {
    throw new Error(
      'La cellule sélectionnée ne contient aucun lien.'
    );
  }

  const match = profileUrl.match(
    CONFIG.PROFILE_URL_PATTERN
  );

  if (!match) {
    throw new Error(
      `Le lien trouvé n’est pas reconnu : ${profileUrl}`
    );
  }

  if (col <= 1) {
    throw new Error(
      'La cellule du lien ne possède aucune cellule à gauche.'
    );
  }

  const firstName = linkCell.getDisplayValue().trim();
  const lastName = linkCell.offset(0, 1).getDisplayValue().trim();

  return {
    spreadsheetId: spreadsheet.getId(),
    sheetId: sheet.getSheetId(),
    pictureCell: sheet.getRange(row, col - 1),
    userId: decodeURIComponent(match[1]),
    profileUrl,
    displayName: `${linkCell.getDisplayValue().trim()}_${linkCell.offset(0, 1).getDisplayValue().trim()}` };
}

function getUrlFromCell_(cell) {
  const richText = cell.getRichTextValue();
  return richText ? richText.getLinkUrl() : null;
}

function isLicenceProfileUrl_(url) {
  return Boolean(url && CONFIG.PROFILE_URL_PATTERN.test(url));
}


function doGet(e) {

  try {
    assertApiToken_(e && e.parameter ? e.parameter.token : null);

    const action = e.parameter.action || 'pendingJobs';

    if (action === 'pendingRegistration') {
      return jsonResponse_({
        success: true,
        registration: getPendingFfbRegistration_()
      });
    }

    if (action === 'pendingJobs') {
      return jsonResponse_({
        success: true,
        jobs: listPendingJobs_()
      });
    }

    throw new Error(`Action GET inconnue : ${action}`);
  } catch (error) {
    return jsonResponse_({
      success: false,
      error: error.message
    });
  }
}

function getPendingFfbRegistration_() {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty('FFB_PENDING_REGISTRATION');

  if (!value) return null;

  return JSON.parse(value);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    assertApiToken_(data.token);

    if (data.action === 'claimJob') {
      return jsonResponse_({
        success: true,
        job: claimJob_(data.requestId)
      });
    }

    if (data.action === 'uploadPicture') {
      return uploadPictureResponse_(data);
    }

    if (data.action === 'failJob') {
      failJob_(data.requestId, data.error);
      return jsonResponse_({ success: true });
    }

    throw new Error(`Action POST inconnue : ${data.action}`);
  } catch (error) {
    return jsonResponse_({
      success: false,
      error: error.message
    });
  }
}

function uploadPictureResponse_(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const job = findJobByRequestId_(data.requestId);

    if (!job) {
      throw new Error('Demande introuvable.');
    }

    if (!['PENDING', 'PROCESSING'].includes(job.status)) {
      throw new Error(`La demande est déjà ${job.status}.`);
    }

    if (String(job.userId) !== String(data.userId)) {
      throw new Error('Le licencié ne correspond pas à la demande.');
    }

    const picture = saveBase64Picture_(job, data);
    updatePictureCell_(job, picture.driveUrl);
    completeJob_(job, picture.fileId, picture.driveUrl);

    return jsonResponse_({
      success: true,
      requestId: job.requestId,
      userId: job.userId,
      driveUrl: picture.driveUrl
    });
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateApiToken_() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty(CONFIG.API_TOKEN_PROPERTY);

  if (!token) {
    token = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty(CONFIG.API_TOKEN_PROPERTY, token);
  }

  return token;
}

function assertApiToken_(providedToken) {
  const expectedToken = getOrCreateApiToken_();

  if (!providedToken || providedToken !== expectedToken) {
    throw new Error('Accès refusé.');
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

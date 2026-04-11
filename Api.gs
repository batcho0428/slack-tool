/**
 * Next.js frontend integration endpoint.
 * POST JSON: { action: string, payload?: object }
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return _apiJson({ ok: false, error: 'Empty request body' });
    }

    const body = JSON.parse(e.postData.contents);
    if (!_isApiAuthorized(body)) {
      return _apiJson({ ok: false, error: 'Unauthorized request' });
    }

    const action = body && body.action ? String(body.action) : '';
    const payload = body && body.payload ? body.payload : {};

    if (!action) return _apiJson({ ok: false, error: 'Missing action' });

    const result = _dispatchApiAction(action, payload);
    return _apiJson({ ok: true, result: result });
  } catch (err) {
    return _apiJson({ ok: false, error: _apiErrorString(err) });
  }
}

function _isApiAuthorized(body) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('FRONTEND_API_SHARED_SECRET');
  if (!expected) {
    throw new Error('Script property FRONTEND_API_SHARED_SECRET is not set');
  }

  const actual = body && body.authToken ? String(body.authToken) : '';
  return actual === expected;
}

function _apiJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _apiErrorString(err) {
  try {
    return err && err.message ? String(err.message) : String(err);
  } catch (e) {
    return 'Unknown error';
  }
}

function _dispatchApiAction(action, payload) {
  if (payload && Object.prototype.toString.call(payload.__args) === '[object Array]') {
    var fn = this[action];
    if (typeof fn !== 'function') {
      throw new Error('Unsupported action: ' + action);
    }
    return fn.apply(null, payload.__args);
  }

  switch (action) {
    case 'ping':
      return { now: new Date().getTime() };
    case 'getLoginUser':
      return getLoginUser(payload.sessionToken);
    case 'requestLoginOtp':
      return requestLoginOtp(payload.email);
    case 'verifyLoginOtp':
      return verifyLoginOtp(payload.email, payload.code);
    case 'getAuthUrl':
      return { url: getAuthUrl() };
    case 'handleSlackOAuthCode':
      return handleSlackOAuthCode(payload.code, payload.redirectUri);
    case 'getScriptUrl':
      return { url: getScriptUrl() };
    case 'getSearchOptions':
      return getSearchOptions();
    case 'searchRecipients':
      return searchRecipients(payload.criteria || {});
    case 'sendDMs':
      return sendDMs(payload.sessionToken, payload.message, payload.recipients || []);
    case 'getChannels':
      return getChannels(payload.sessionToken);
    case 'inviteToChannel':
      return inviteToChannel(payload.sessionToken, payload.channelId, payload.recipients || []);
    case 'getUserProfile':
      return getUserProfile(payload.sessionToken, payload.targetEmail);
    case 'updateUserProfile':
      return updateUserProfile(payload.sessionToken, payload.formData || {}, payload.targetEmail);
    case 'createUser':
      return createUser(payload.sessionToken, payload.userObj || {});
    case 'listSurveys':
      return listSurveys(payload.sessionToken);
    case 'getSurveyDetails':
      return getSurveyDetails(payload.sessionToken, payload.spreadsheetRef, payload.rowIndex);
    case 'listCollections':
      return listCollections(payload.sessionToken);
    case 'createCollection':
      return createCollection(payload.sessionToken, payload.payload || {});
    case 'updateCollection':
      return updateCollection(payload.sessionToken, payload.collectionId, payload.payload || {});
    case 'deleteCollection':
      return deleteCollection(payload.sessionToken, payload.collectionId);
    case 'fetchCollectionSummary':
      return fetchCollectionSummary(payload.sessionToken, payload.collectionId);
    case 'getCollectionRowDetails':
      return getCollectionRowDetails(payload.sessionToken, payload.collectionId, payload.recipientEmail);
    case 'recordPayment':
      return recordPayment(
        payload.sessionToken,
        payload.collectionId,
        payload.recipientEmail,
        payload.amount,
        payload.type,
        payload.handlerEmail
      );
    case 'recordPaymentWithChange':
      return recordPaymentWithChange(
        payload.sessionToken,
        payload.collectionId,
        payload.recipientEmail,
        payload.receivedAmount,
        payload.expectedAmount,
        payload.handlerEmail
      );
    case 'createRosterCsv':
      return createRosterCsv(payload.sessionToken, payload.params || {});
    default:
      throw new Error('Unsupported action: ' + action);
  }
}

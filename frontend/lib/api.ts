type ApiEnvelope<T> = {
  ok: boolean;
  result?: T;
  error?: string;
};

async function callApi<T>(action: string, payload?: Record<string, unknown>): Promise<T> {
  const res = await fetch('/api/gas', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action, payload: payload || {} })
  });

  const text = await res.text();
  let json: ApiEnvelope<T>;
  try {
    json = JSON.parse(text) as ApiEnvelope<T>;
  } catch (e) {
    const snippet = String(text || '').slice(0, 120);
    throw new Error(`API応答がJSONではありません: ${snippet}`);
  }
  if (!json.ok) {
    throw new Error(json.error || 'API request failed');
  }
  return json.result as T;
}

export const gasApi = {
  ping: () => callApi<{ now: number }>('ping'),
  getLoginUser: (sessionToken: string) => callApi('getLoginUser', { sessionToken }),
  requestLoginOtp: (email: string) => callApi('requestLoginOtp', { email }),
  verifyLoginOtp: (email: string, code: string) => callApi('verifyLoginOtp', { email, code }),
  getAuthUrl: () => callApi<{ url: string }>('getAuthUrl'),
  getScriptUrl: () => callApi<{ url: string }>('getScriptUrl'),
  getSearchOptions: () => callApi('getSearchOptions'),
  searchRecipients: (criteria: Record<string, unknown>) => callApi('searchRecipients', { criteria }),
  sendDMs: (sessionToken: string, message: string, recipients: unknown[]) =>
    callApi('sendDMs', { sessionToken, message, recipients }),
  getChannels: (sessionToken: string) => callApi('getChannels', { sessionToken }),
  inviteToChannel: (sessionToken: string, channelId: string, recipients: unknown[]) =>
    callApi('inviteToChannel', { sessionToken, channelId, recipients }),
  getUserProfile: (sessionToken: string, targetEmail?: string) =>
    callApi('getUserProfile', { sessionToken, targetEmail }),
  updateUserProfile: (sessionToken: string, formData: Record<string, unknown>, targetEmail?: string) =>
    callApi('updateUserProfile', { sessionToken, formData, targetEmail }),
  listSurveys: (sessionToken: string) => callApi('listSurveys', { sessionToken }),
  getSurveyDetails: (sessionToken: string, spreadsheetRef: string, rowIndex?: number) =>
    callApi('getSurveyDetails', { sessionToken, spreadsheetRef, rowIndex }),
  listCollections: (sessionToken: string) => callApi('listCollections', { sessionToken }),
  createCollection: (sessionToken: string, payload: Record<string, unknown>) =>
    callApi('createCollection', { sessionToken, payload }),
  updateCollection: (sessionToken: string, collectionId: string, payload: Record<string, unknown>) =>
    callApi('updateCollection', { sessionToken, collectionId, payload }),
  deleteCollection: (sessionToken: string, collectionId: string) =>
    callApi('deleteCollection', { sessionToken, collectionId }),
  fetchCollectionSummary: (sessionToken: string, collectionId: string) =>
    callApi('fetchCollectionSummary', { sessionToken, collectionId }),
  recordPayment: (
    sessionToken: string,
    collectionId: string,
    recipientEmail: string,
    amount: number,
    type: string,
    handlerEmail: string
  ) => callApi('recordPayment', { sessionToken, collectionId, recipientEmail, amount, type, handlerEmail }),
  recordPaymentWithChange: (
    sessionToken: string,
    collectionId: string,
    recipientEmail: string,
    receivedAmount: number,
    expectedAmount: number,
    handlerEmail: string
  ) =>
    callApi('recordPaymentWithChange', {
      sessionToken,
      collectionId,
      recipientEmail,
      receivedAmount,
      expectedAmount,
      handlerEmail
    }),
  createRosterCsv: (sessionToken: string, params: Record<string, unknown>) =>
    callApi('createRosterCsv', { sessionToken, params })
};

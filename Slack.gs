function sendDMs(sessionToken, message, recipients) {
  const { token, email: senderEmail } = getUserToken(sessionToken);
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  let successCount = 0;
  const failedList = [];
  const time = new Date();

  recipients.forEach((r) => {
    try {
      const uid = getSlackID(token, r.email);
      if (!uid) throw new Error("Slackアカウントなし");
      const text = message.replace(/{mention}/g, `<@${uid}>`);
      const res = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
        method: "post",
        contentType: "application/json",
        headers: { "Authorization": "Bearer " + token },
        payload: JSON.stringify({ channel: uid, text: text, unfurl_links: false, unfurl_media: false }),
        muteHttpExceptions: true
      });
      const json = JSON.parse(res.getContentText());
      if (!json.ok) throw new Error(json.error || "Unknown Error");
      successCount++;
      logSheet.appendRow([time, senderEmail, r.email, "DM Success", ""]);
    } catch (e) {
      failedList.push({ email: r.email, error: e.message });
      logSheet.appendRow([time, senderEmail, r.email, "DM Failed", e.message]);
    }
    Utilities.sleep(500);
  });
  return { success: successCount, failed: failedList };
}
function getChannels(sessionToken) {
  const { token } = getUserToken(sessionToken);
  const fetch = function(types) {
    let channels = [];
    let cursor = "";
    do {
      const url = `https://slack.com/api/conversations.list?types=${types}&exclude_archived=true&limit=200&cursor=${cursor}`;
      const res = UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
      const json = JSON.parse(res.getContentText());
      if (!json.ok) throw new Error(json.error);
      if (json.channels) channels = channels.concat(json.channels);
      cursor = (json.response_metadata && json.response_metadata.next_cursor) ? json.response_metadata.next_cursor : "";
    } while (cursor);
    return channels;
  };

  try {
    try {
      return fetch("public_channel,private_channel")
        .map(c => ({ id: c.id, name: c.name, is_private: c.is_private }))
        .sort((a, b) => a.name.localeCompare(b.name));
    } catch (e) {
      if (e.message === 'missing_scope') {
        return fetch("public_channel")
          .map(c => ({ id: c.id, name: c.name, is_private: c.is_private }))
          .sort((a, b) => a.name.localeCompare(b.name));
      }
      throw e;
    }
  } catch (e) {
    throw new Error("チャンネル一覧取得失敗: " + e.message);
  }
}
function inviteToChannel(sessionToken, channelId, recipients) {
  const { token, email: senderEmail } = getUserToken(sessionToken);
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  let successCount = 0;
  const failedList = [];
  const time = new Date();

  recipients.forEach((r) => {
    try {
      const uid = getSlackID(token, r.email);
      if (!uid) throw new Error("Slackアカウントなし");
      const res = UrlFetchApp.fetch("https://slack.com/api/conversations.invite", {
        method: "post", contentType: "application/json", headers: { "Authorization": "Bearer " + token },
        payload: JSON.stringify({ channel: channelId, users: uid }), muteHttpExceptions: true
      });
      const json = JSON.parse(res.getContentText());
      if (!json.ok) {
        if (json.error === 'already_in_channel') throw new Error("既に参加済みです");
        else if (json.error === 'not_in_channel') throw new Error("実行者(あなた)がこのチャンネルに参加していません");
        else throw new Error(json.error);
      }
      successCount++;
      logSheet.appendRow([time, senderEmail, r.email, "Invite Success (User)", channelId]);
    } catch (e) {
      failedList.push({ email: r.email, error: e.message });
      logSheet.appendRow([time, senderEmail, r.email, "Invite Failed", `${channelId}: ${e.message}`]);
    }
    Utilities.sleep(500);
  });
  return { success: successCount, failed: failedList };
}

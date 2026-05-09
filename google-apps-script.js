/**
 * 10 分鐘諮詢預約自動化
 *
 * 使用方式：
 * 1. 到 https://script.google.com 建立新專案。
 * 2. 貼上此檔案內容。
 * 3. 左側「服務」新增 Calendar API 進階服務，並在 Google Cloud 同步啟用 Google Calendar API。
 * 4. 部署為 Web app，執行身分選「我」，存取權選「任何人」。
 * 5. 把 Web app URL 貼回 10分鐘諮詢表單.html 的 APPS_SCRIPT_URL。
 */

const CONFIG = {
  OWNER_EMAIL: 'astroalys@astroalys.com',
  CALENDAR_ID: 'primary',
  SPREADSHEET_ID: '1Ljk0cA078jHDUzBu8AryjMPNekdqK3P7ynNVoirejAk',
  SHEET_NAME: '10分鐘諮詢預約',
  TIME_ZONE: 'Asia/Taipei',
  DURATION_MINUTES: 10,
  LOOKAHEAD_DAYS: 28,
  ALLOWED_WEEKDAYS: [4, 5],
  ALLOWED_TIMES: ['11:30', '12:00', '12:30'],
  EVENT_TITLE_PREFIX: '10分鐘諮詢',
  FORM_TOKEN: 'alys-consult-2026-x9Kp4Qm7Vz',
};

function doGet(e) {
  try {
    validateToken(e.parameter || {});

    if ((e.parameter || {}).action !== 'availability') {
      return callbackResponse({ ok: false, reason: 'unknown_action' }, e);
    }

    return callbackResponse({ ok: true, slots: getAvailableSlots() }, e);
  } catch (error) {
    return callbackResponse({ ok: false, reason: 'server_error', message: String(error) }, e);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    validateToken(e.parameter || {});

    const data = normalizePayload(e.parameter || {});
    validatePayload(data);

    const start = new Date(data.slot_iso);
    const end = new Date(start.getTime() + CONFIG.DURATION_MINUTES * 60 * 1000);
    validateSlot(start);

    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!calendar) throw new Error('找不到 Google Calendar：' + CONFIG.CALENDAR_ID);

    const conflicts = calendar.getEvents(start, end);

    if (conflicts.length > 0) {
      appendRow(data, start, end, '時段已被占用', '');
      sendUnavailableEmail(data);
      notifyOwnerUnavailable(data, start);
      return jsonResponse({ ok: false, reason: 'slot_unavailable' });
    }

    const event = createMeetEvent(data, start, end);
    const meetLink = getMeetLink(event);

    appendRow(data, start, end, '預約成功', meetLink);
    sendClientSuccessEmail(data, start, end, meetLink);
    sendOwnerSuccessEmail(data, start, end, meetLink);

    return jsonResponse({ ok: true, meetLink: meetLink });
  } catch (error) {
    notifyOwnerError(error, e && e.parameter ? e.parameter : {});
    return jsonResponse({ ok: false, reason: 'server_error', message: String(error) });
  } finally {
    lock.releaseLock();
  }
}

function validateToken(params) {
  if (!params || params.form_token !== CONFIG.FORM_TOKEN) {
    throw new Error('Invalid form token');
  }
}

function getAvailableSlots() {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) throw new Error('找不到 Google Calendar：' + CONFIG.CALENDAR_ID);

  const now = new Date();
  const slots = [];

  for (let dayOffset = 0; dayOffset <= CONFIG.LOOKAHEAD_DAYS; dayOffset += 1) {
    const date = new Date(now.getTime() + dayOffset * 24 * 60 * 60 * 1000);
    const weekday = Number(Utilities.formatDate(date, CONFIG.TIME_ZONE, 'u')) % 7;
    const jsWeekday = weekday === 0 ? 0 : weekday;

    if (!CONFIG.ALLOWED_WEEKDAYS.includes(jsWeekday)) continue;

    const dateValue = Utilities.formatDate(date, CONFIG.TIME_ZONE, 'yyyy-MM-dd');

    CONFIG.ALLOWED_TIMES.forEach(function(time) {
      const start = new Date(dateValue + 'T' + time + ':00+08:00');
      const end = new Date(start.getTime() + CONFIG.DURATION_MINUTES * 60 * 1000);

      if (start <= now) return;
      if (calendar.getEvents(start, end).length > 0) return;

      slots.push({
        iso: dateValue + 'T' + time + ':00+08:00',
        label: Utilities.formatDate(start, CONFIG.TIME_ZONE, 'MM/dd（E）HH:mm'),
      });
    });
  }

  return slots;
}

function normalizePayload(raw) {
  return {
    submitted_at: new Date(),
    name: trim(raw.name),
    email: trim(raw.email),
    change_goal: trim(raw.change_goal),
    timeline: trim(raw.timeline),
    past_attempts: trim(raw.past_attempts),
    six_month_worry: trim(raw.six_month_worry),
    slot_iso: trim(raw.slot_iso),
    slot_label: trim(raw.slot_label),
    timezone: trim(raw.timezone) || CONFIG.TIME_ZONE,
    confirm: raw.confirm ? '已確認' : '',
  };
}

function trim(value) {
  return value == null ? '' : String(value).trim();
}

function validatePayload(data) {
  const requiredFields = ['name', 'email', 'change_goal', 'timeline', 'past_attempts', 'six_month_worry', 'slot_iso', 'confirm'];
  requiredFields.forEach(function(field) {
    if (!data[field]) throw new Error('Missing required field: ' + field);
  });

  if (!/^\S+@\S+\.\S+$/.test(data.email)) {
    throw new Error('Invalid email: ' + data.email);
  }
}

function validateSlot(start) {
  if (Number.isNaN(start.getTime())) throw new Error('Invalid slot_iso');

  const weekday = Number(Utilities.formatDate(start, CONFIG.TIME_ZONE, 'u')) % 7;
  const jsWeekday = weekday === 0 ? 0 : weekday;
  const time = Utilities.formatDate(start, CONFIG.TIME_ZONE, 'HH:mm');

  if (!CONFIG.ALLOWED_WEEKDAYS.includes(jsWeekday)) {
    throw new Error('Slot weekday is not allowed: ' + jsWeekday);
  }

  if (!CONFIG.ALLOWED_TIMES.includes(time)) {
    throw new Error('Slot time is not allowed: ' + time);
  }

  if (start <= new Date()) {
    throw new Error('Slot is in the past');
  }
}

function createMeetEvent(data, start, end) {
  const description = [
    '你的 10 分鐘諮詢已預約成功。',
    '',
    '請於預約時間點擊 Google Meet 連結進入會議。',
    '',
    '若臨時需要調整時間，請透過 LINE@ 聯繫：',
    'https://lin.ee/OYa16Dv',
    '',
    'Alys',
  ].join('\n');

  const event = {
    summary: CONFIG.EVENT_TITLE_PREFIX + '｜' + data.name,
    description: description,
    start: {
      dateTime: start.toISOString(),
      timeZone: CONFIG.TIME_ZONE,
    },
    end: {
      dateTime: end.toISOString(),
      timeZone: CONFIG.TIME_ZONE,
    },
    attendees: [
      { email: data.email },
      { email: CONFIG.OWNER_EMAIL },
    ],
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: 'hangoutsMeet' },
      },
    },
    reminders: {
      useDefault: false,
      overrides: [
        { method: 'email', minutes: 24 * 60 },
        { method: 'popup', minutes: 30 },
      ],
    },
  };

  return Calendar.Events.insert(event, CONFIG.CALENDAR_ID, {
    conferenceDataVersion: 1,
    sendUpdates: 'all',
  });
}

function getMeetLink(event) {
  if (event.hangoutLink) return event.hangoutLink;

  const entryPoints = event.conferenceData && event.conferenceData.entryPoints;
  if (!entryPoints) return '';

  const video = entryPoints.find(function(entry) { return entry.entryPointType === 'video'; });
  return video ? video.uri : '';
}

function appendRow(data, start, end, status, meetLink) {
  const sheet = getSheet();
  sheet.appendRow([
    data.submitted_at,
    status,
    data.name,
    data.email,
    data.change_goal,
    data.timeline,
    data.past_attempts,
    data.six_month_worry,
    data.slot_label,
    start,
    end,
    meetLink,
  ]);
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow([
      '填寫時間',
      '狀態',
      '姓名或暱稱',
      'Email',
      'Q1 目前最想改變',
      'Q2 改變期限',
      'Q3 過去嘗試',
      'Q4 半年後擔心',
      '選擇時段',
      '開始時間',
      '結束時間',
      'Google Meet',
    ]);
  }

  return sheet;
}

function sendClientSuccessEmail(data, start, end, meetLink) {
  const subject = 'Alys 巫｜10分鐘諮詢行前通知';
  const body = [
    'Hi ' + data.name + '，',
    '',
    '這是你預約的 10 分鐘諮詢時段 😊',
    '',
    '時間：' + formatDateTime(start) + ' - ' + Utilities.formatDate(end, CONFIG.TIME_ZONE, 'HH:mm'),
    '',
    'Google Meet：' + meetLink,
    '',
    '這 10 分鐘裡，我會陪你一起整理目前的狀態，',
    '看看現在真正卡住的地方，可能是什麼。',
    '有時候，不是你不夠努力，',
    '而是有些模式已經默默影響你很久了。',
    '',
    '---',
    '',
    '諮詢前，你可以稍微想一下：',
    '',
    '・你現在最困擾的是什麼？',
    '',
    '・這個狀態持續多久了？',
    '',
    '不需要準備很多，',
    '',
    '帶著你現在最真實的狀態來就可以了',
    '',
    '---',
    '',
    '如果臨時需要調整時間，',
    '',
    '請提前透過 Line@ 跟我聯繫。',
    '',
    'Line@：https://lin.ee/OYa16Dv',
    '',
    '我們線上見 😊',
    '',
    'Alys',
  ].join('\n');

  MailApp.sendEmail(data.email, subject, body, { name: 'Alys 巫' });
}

function sendOwnerSuccessEmail(data, start, end, meetLink) {
  const subject = '新的 10 分鐘諮詢預約｜' + data.name;
  const body = [
    '你收到一筆新的 10 分鐘諮詢預約。',
    '',
    '姓名：' + data.name,
    'Email：' + data.email,
    '時間：' + formatDateTime(start) + ' - ' + Utilities.formatDate(end, CONFIG.TIME_ZONE, 'HH:mm'),
    'Google Meet：' + meetLink,
    '',
    '諮詢前問卷',
    '',
    '1. 目前最想改變的是什麼？',
    data.change_goal,
    '',
    '2. 希望這次改變最晚多久內開始發生？',
    data.timeline,
    '',
    '3. 過去為了解決這個問題，做過哪些嘗試？',
    data.past_attempts,
    '',
    '4. 如果半年後狀態完全沒變，最擔心的是什麼？',
    data.six_month_worry,
  ].join('\n');

  MailApp.sendEmail(CONFIG.OWNER_EMAIL, subject, body, { name: '10 分鐘諮詢預約系統' });
}

function sendUnavailableEmail(data) {
  const subject = '你選擇的 10 分鐘諮詢時段已被預約';
  const body = [
    data.name + ' 你好，',
    '',
    '你剛剛選擇的時段：' + data.slot_label + ' 目前已經被預約或暫停開放。',
    '',
    '請回到表單重新選擇其他時段。',
    '',
    'Alys',
    '天台的占卜巫',
  ].join('\n');

  MailApp.sendEmail(data.email, subject, body, { name: '天台的占卜巫 Alys' });
}

function notifyOwnerUnavailable(data, start) {
  MailApp.sendEmail(
    CONFIG.OWNER_EMAIL,
    '有人選到已被占用的 10 分鐘諮詢時段',
    '姓名：' + data.name + '\nEmail：' + data.email + '\n時段：' + formatDateTime(start),
    { name: '10 分鐘諮詢預約系統' }
  );
}

function notifyOwnerError(error, data) {
  MailApp.sendEmail(
    CONFIG.OWNER_EMAIL,
    '10 分鐘諮詢預約系統發生錯誤',
    '錯誤：' + String(error) + '\n\n原始資料：\n' + JSON.stringify(data, null, 2),
    { name: '10 分鐘諮詢預約系統' }
  );
}

function formatDateTime(date) {
  return Utilities.formatDate(date, CONFIG.TIME_ZONE, 'yyyy/MM/dd（E）HH:mm');
}

function callbackResponse(payload, e) {
  const params = e && e.parameter ? e.parameter : {};
  const callback = params.callback;

  if (callback && /^[A-Za-z_$][\w.$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse(payload);
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}





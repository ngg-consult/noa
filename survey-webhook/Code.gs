/**
 * SurveyLegend → Google Sheets + התראה על תשובה שלילית
 *
 * הסקריפט מקבל כל תשובה שנשלחת מה-Webhook של SurveyLegend, כותב אותה לגיליון
 * ושולח מייל כאשר התשובה עומדת באחד מכללי "תשובה שלילית" שמוגדרים למטה.
 *
 * שלב 1 (עכשיו): להשאיר NEGATIVE_RULES ריק, לשלוח תשובת ניסיון ולבדוק בלשונית Raw
 *               איך נראים הנתונים.
 * שלב 2: למלא את NEGATIVE_RULES לפי שמות השדות שהתקבלו בפועל.
 */

// ===================== הגדרות =====================

const CONFIG = {
  // מחרוזת סודית שמתווספת לכתובת ה-Webhook (?key=...). לשנות לערך אקראי משלך.
  SECRET_KEY: 'CHANGE_ME_TO_A_RANDOM_STRING',

  // נמענים להתראה, מופרדים בפסיקים.
  ALERT_EMAILS: 'client@example.com',

  // שם הסקר שיופיע בכותרת המייל.
  SURVEY_NAME: 'סקר שביעות רצון נוסעים',

  RESPONSES_SHEET: 'Responses',
  RAW_SHEET: 'Raw',
  ALERTS_SHEET: 'Alerts',
};

/**
 * כללי תשובה שלילית. תשובה שעומדת בכלל אחד לפחות תפעיל התראה.
 *   field    – חלק משם השדה כפי שמופיע בכותרת העמודה בלשונית Responses (ללא תלות באותיות גדולות/קטנות).
 *              '*' = כל שדה.
 *   values   – רשימת תשובות טקסט שנחשבות שליליות (התאמה מדויקת, ללא רווחים בקצוות).
 *   maxScore – ציון מספרי שממנו ומטה התשובה שלילית (למשל 2 בסולם 1–5).
 *
 * דוגמה:
 *   { field: 'satisfaction', maxScore: 2 },
 *   { field: '*', values: ['לא מרוצה', 'גרוע'] },
 */
const NEGATIVE_RULES = [
];

// ===================== נקודות כניסה =====================

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!e || !e.parameter || e.parameter.key !== CONFIG.SECRET_KEY) {
      return respond_({ ok: false, error: 'unauthorized' });
    }

    const rawBody = e.postData ? e.postData.contents : '';
    logRaw_(rawBody);

    const payloads = parsePayload_(e);
    payloads.forEach(handleResponse_);

    return respond_({ ok: true, received: payloads.length });
  } catch (err) {
    logRaw_('ERROR: ' + err + '\n' + (e && e.postData ? e.postData.contents : ''));
    return respond_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return respond_({ ok: true, status: 'webhook is running' });
}

// ===================== עיבוד =====================

function parsePayload_(e) {
  const body = e.postData ? e.postData.contents : '';
  let data;
  try {
    data = JSON.parse(body);
  } catch (_) {
    // פורמט Web Form: השדות מגיעים כפרמטרים רגילים.
    data = Object.assign({}, e.parameter);
    delete data.key;
  }
  // אם האפשרות Append responses שולחת כמה תשובות יחד – לטפל בכל אחת בנפרד.
  return Array.isArray(data) ? data : [data];
}

function handleResponse_(payload) {
  const flat = flatten_(payload);
  flat['_received_at'] = new Date();

  appendToSheet_(CONFIG.RESPONSES_SHEET, flat);

  const hits = findNegatives_(flat);
  if (hits.length) {
    sendAlert_(flat, hits);
    appendToSheet_(CONFIG.ALERTS_SHEET, {
      _received_at: flat['_received_at'],
      reasons: hits.map(h => h.field + ': ' + h.value).join(' | '),
    });
  }
}

function flatten_(obj, prefix, out) {
  out = out || {};
  prefix = prefix || '';
  if (obj === null || obj === undefined) {
    if (prefix) out[prefix] = '';
    return out;
  }
  if (Array.isArray(obj)) {
    if (obj.every(v => v === null || typeof v !== 'object')) {
      out[prefix] = obj.join(', ');
    } else {
      obj.forEach((v, i) => flatten_(v, prefix ? prefix + '.' + i : String(i), out));
    }
    return out;
  }
  if (typeof obj === 'object') {
    Object.keys(obj).forEach(k => flatten_(obj[k], prefix ? prefix + '.' + k : k, out));
    return out;
  }
  out[prefix] = obj;
  return out;
}

function findNegatives_(flat) {
  const hits = [];
  Object.keys(flat).forEach(key => {
    if (key.charAt(0) === '_') return;
    const value = flat[key];
    const text = String(value).trim();
    NEGATIVE_RULES.forEach(rule => {
      const fieldMatch = rule.field === '*' || key.toLowerCase().indexOf(String(rule.field).toLowerCase()) !== -1;
      if (!fieldMatch) return;
      const valueMatch = rule.values && rule.values.indexOf(text) !== -1;
      const num = Number(text);
      const scoreMatch = rule.maxScore !== undefined && text !== '' && !isNaN(num) && num <= rule.maxScore;
      if (valueMatch || scoreMatch) hits.push({ field: key, value: text });
    });
  });
  return hits;
}

// ===================== גיליון =====================

function appendToSheet_(name, record) {
  const sheet = getSheet_(name);
  let headers = sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];

  const missing = Object.keys(record).filter(k => headers.indexOf(k) === -1);
  if (missing.length) {
    headers = headers.concat(missing);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const row = headers.map(h => (h in record ? record[h] : ''));
  sheet.appendRow(row);
}

function logRaw_(text) {
  getSheet_(CONFIG.RAW_SHEET).appendRow([new Date(), text]);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ===================== מייל =====================

function sendAlert_(flat, hits) {
  const reasons = hits.map(h =>
    '<li><b>' + escape_(h.field) + '</b>: ' + escape_(h.value) + '</li>').join('');
  const details = Object.keys(flat)
    .filter(k => k.charAt(0) !== '_')
    .map(k => '<tr><td style="padding:4px 8px;border:1px solid #ddd">' + escape_(k) +
              '</td><td style="padding:4px 8px;border:1px solid #ddd">' + escape_(flat[k]) + '</td></tr>')
    .join('');

  const html =
    '<div dir="rtl" style="font-family:Arial,sans-serif">' +
    '<p>התקבלה תשובה שלילית בסקר <b>' + escape_(CONFIG.SURVEY_NAME) + '</b>.</p>' +
    '<p>הסיבות שזוהו:</p><ul>' + reasons + '</ul>' +
    '<p>פרטי התשובה המלאה:</p>' +
    '<table style="border-collapse:collapse">' + details + '</table>' +
    '</div>';

  MailApp.sendEmail({
    to: CONFIG.ALERT_EMAILS,
    subject: 'התראה: תשובה שלילית – ' + CONFIG.SURVEY_NAME,
    htmlBody: html,
  });
}

function escape_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===================== בדיקה ידנית =====================

/** הרצה ידנית מתוך העורך: מדמה תשובה ובודקת שהגיליון והמייל עובדים. */
function testWebhook() {
  const fake = {
    parameter: { key: CONFIG.SECRET_KEY },
    postData: { contents: JSON.stringify({ test: true, question: 'שביעות רצון', answer: 1 }) },
  };
  Logger.log(doPost(fake).getContent());
}

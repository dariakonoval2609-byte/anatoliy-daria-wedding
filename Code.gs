/**
 * Google Apps Script backend for the wedding RSVP form.
 * Bind this script to the guest spreadsheet, run setupWeddingRsvp() once,
 * then deploy it as a web app to receive guest responses.
 */
const RESPONSE_TAB_NAME = "Відповіді гостей";
const INVITE_CODE_HEADER = "Код запрошення";
const RESPONSE_HEADERS = [
  "Код запрошення", "ПІБ", "Присутність", "Гість +1", "Кількість дітей",
  "Проживання", "Трансфер", "Другий день", "Алкоголь", "Харчові обмеження / побажання", "Оновлено"
];

const WEDDING_SITE_ORIGIN = "https://dariakonoval2609-byte.github.io";

/** Return only the guest data needed to render their personal invitation. */
function doGet(e) {
  const action = String(e && e.parameter && e.parameter.action || "");
  const code = String(e && e.parameter && e.parameter.invite || "").trim();
  if (action !== "guest" || !code) return guestMessageHtml_({ type: "wedding-guest", invite: code, guest: null });
  try {
    const ss = getWeddingSpreadsheet_();
    const sheet = ss.getSheets()[0];
    const headers = getHeaderMap_(sheet);
    const required = ["Прізвище", "Ім'я", "По-батькові", INVITE_CODE_HEADER];
    required.forEach(name => { if (headers[name] === undefined) throw new Error("Не знайдено колонку: " + name); });
    const rows = sheet.getDataRange().getDisplayValues();
    const matchingRow = rows.slice(1).find(row => constantTimeEqual_(row[headers[INVITE_CODE_HEADER]], code));
    if (!matchingRow) return guestMessageHtml_({ type: "wedding-guest", invite: code, guest: null });

    const lastName = cellByHeader_(matchingRow, headers, ["Прізвище"]);
    const firstName = cellByHeader_(matchingRow, headers, ["Ім'я"]);
    const patronymic = cellByHeader_(matchingRow, headers, ["По-батькові"]);
    const relationship = cellByHeader_(matchingRow, headers, ["Зв'язки з подружжям", "Зв’язки з подружжям", "Зв'язок із подружжям"]);
    const statusPlusOne = normalize_(cellByHeader_(matchingRow, headers, ["Статус гостя (+1)"]));
    const plusOneName = cellByHeader_(matchingRow, headers, ["Ім'я запрошеного гостя +1", "Ім’я запрошеного гостя +1"]);
    const lodging = normalize_(cellByHeader_(matchingRow, headers, ["Заселення", "Проживання"]));
    const transfer = normalize_(cellByHeader_(matchingRow, headers, ["Потреба у Трансфері", "Трансфер"]));
    const staying = /^(так|yes|true|1|потрібне|потрібен|буду|потрібно)$/.test(lodging);
    const needsTransfer = /^(так|yes|true|1|потрібне|потрібен|буду|потрібно)$/.test(transfer);
    let scheduleHeader = "Розклад весілля";
    if (staying && needsTransfer) scheduleHeader = "Розклад весілля з трансфером і поселенням";
    else if (staying) scheduleHeader = "Розклад весілля з поселенням";
    else if (needsTransfer) scheduleHeader = "Розклад весілля з трансфером";
    const scheduleText = cellByHeader_(matchingRow, headers, [scheduleHeader]) || cellByHeader_(matchingRow, headers, ["Розклад весілля"]);
    const invitationText = cellByHeader_(matchingRow, headers, ["Теплі слова запрошення", "Теплі слова запрошення для гостя"]);
    const guest = {
      fullName: [lastName, firstName, patronymic].filter(Boolean).join(" "),
      firstName: firstName,
      relationship: relationship,
      plusOneAllowed: /^(так|yes|true|1|\+|є)$/.test(statusPlusOne),
      plusOneName: plusOneName,
      invitationText: invitationText,
      scheduleText: scheduleText
    };
    return guestMessageHtml_({ type: "wedding-guest", invite: code, guest: guest });
  } catch (error) {
    console.error(error);
    return guestMessageHtml_({ type: "wedding-guest", invite: code, guest: null });
  }
}

function setupWeddingRsvp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Спочатку відкрийте Apps Script з потрібної Google-таблиці.");
  PropertiesService.getScriptProperties().setProperty("WEDDING_SPREADSHEET_ID", ss.getId());
  const guestSheet = ss.getSheets()[0];
  const headerMap = getHeaderMap_(guestSheet);
  ["Прізвище", "Ім'я", "По-батькові"].forEach(name => {
    if (headerMap[name] === undefined) throw new Error("У першому рядку таблиці не знайдено колонку: " + name);
  });
  let codeColumn = headerMap[INVITE_CODE_HEADER];
  if (codeColumn === undefined) {
    codeColumn = guestSheet.getLastColumn();
    guestSheet.getRange(1, codeColumn + 1).setValue(INVITE_CODE_HEADER);
    codeColumn += 1;
  } else {
    codeColumn += 1;
  }
  const lastRow = guestSheet.getLastRow();
  if (lastRow > 1) {
    const range = guestSheet.getRange(2, codeColumn, lastRow - 1, 1);
    const codes = range.getValues();
    let changed = false;
    codes.forEach(row => {
      if (!String(row[0] || "").trim()) {
        row[0] = Utilities.getUuid().replace(/-/g, "");
        changed = true;
      }
    });
    if (changed) range.setValues(codes);
  }
  let responseSheet = ss.getSheetByName(RESPONSE_TAB_NAME);
  if (!responseSheet) responseSheet = ss.insertSheet(RESPONSE_TAB_NAME);
  responseSheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setValues([RESPONSE_HEADERS]);
  responseSheet.setFrozenRows(1);
  return "Готово: створено коди запрошень і вкладку «" + RESPONSE_TAB_NAME + "».";
}

function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.payload
      ? e.parameter.payload
      : (e && e.postData ? e.postData.contents : "{}");
    const payload = JSON.parse(raw);
    const code = String(payload.inviteCode || "").trim();
    const nonce = String(payload.nonce || "").replace(/[^A-Za-z0-9_-]/g, "").slice(0, 100);
    if (!code) return htmlResponse_(false, "Відсутній код запрошення.", nonce);

    const ss = getWeddingSpreadsheet_();
    const guestSheet = ss.getSheets()[0];
    const map = getHeaderMap_(guestSheet);
    const required = ["Прізвище", "Ім'я", "По-батькові", INVITE_CODE_HEADER];
    required.forEach(name => { if (map[name] === undefined) throw new Error("Не знайдено колонку: " + name); });
    const rows = guestSheet.getDataRange().getValues();
    const codeIndex = map[INVITE_CODE_HEADER];
    const matchingRow = rows.slice(1).find(row => constantTimeEqual_(row[codeIndex], code));
    if (!matchingRow) return htmlResponse_(false, "Код запрошення не знайдено. Перевірте персональне посилання.", nonce);

    const invitedName = [matchingRow[map["Прізвище"]], matchingRow[map["Ім'я"]], matchingRow[map["По-батькові"]]]
      .map(value => String(value || "").trim()).filter(Boolean).join(" ");
    if (normalize_(payload.guestName) !== normalize_(invitedName)) {
      return htmlResponse_(false, "ПІБ не збігається із запрошенням. Перевірте написання.", nonce);
    }
    const plusOneName = clean_(payload.plusOne, 160);
    const plusOneStatus = map["Статус гостя (+1)"] === undefined ? "" : normalize_(matchingRow[map["Статус гостя (+1)"]]);
    const plusOneAllowed = /^(так|yes|true|1|\+|є)$/.test(plusOneStatus);
    if (plusOneName && !plusOneAllowed) {
      return htmlResponse_(false, "Це запрошення не передбачає гостя +1.", nonce);
    }
    const listedPlusOne = map["Ім'я запрошеного гостя +1"] === undefined ? "" : String(matchingRow[map["Ім'я запрошеного гостя +1"]] || "").trim();
    if (plusOneAllowed && listedPlusOne && plusOneName && normalize_(plusOneName) !== normalize_(listedPlusOne)) {
      return htmlResponse_(false, "Будь ласка, вкажіть ім’я гостя +1, зазначене у запрошенні.", nonce);
    }
    if (!["Так", "Ні"].includes(String(payload.attendance || ""))) {
      return htmlResponse_(false, "Оберіть відповідь щодо присутності.", nonce);
    }

    const responseSheet = ss.getSheetByName(RESPONSE_TAB_NAME);
    if (!responseSheet) throw new Error("Спершу запустіть setupWeddingRsvp().");
    const alcohol = Array.isArray(payload.alcohol) ? payload.alcohol.join(", ") : String(payload.alcohol || "");
    const record = [
      code,
      invitedName,
      clean_(payload.attendance, 20),
      plusOneName,
      clean_(payload.children, 20),
      clean_(payload.lodging, 20),
      clean_(payload.transfer, 20),
      clean_(payload.secondDay, 20),
      clean_(alcohol, 300),
      clean_(payload.notes, 1000),
      new Date()
    ];
    const lastRow = responseSheet.getLastRow();
    let existingRow = 0;
    if (lastRow > 1) {
      const existingCodes = responseSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      const offset = existingCodes.findIndex(row => constantTimeEqual_(row[0], code));
      if (offset >= 0) existingRow = offset + 2;
    }
    if (existingRow) responseSheet.getRange(existingRow, 1, 1, record.length).setValues([record]);
    else responseSheet.appendRow(record);
    return htmlResponse_(true, "Дякуємо! Вашу відповідь збережено.", nonce);
  } catch (error) {
    console.error(error);
    const raw = e && e.parameter && e.parameter.payload ? e.parameter.payload : "{}";
    let nonce = "";
    try { nonce = String(JSON.parse(raw).nonce || "").replace(/[^A-Za-z0-9_-]/g, "").slice(0, 100); } catch (_) {}
    return htmlResponse_(false, "Не вдалося зберегти відповідь. Повідомте, будь ласка, молодят.", nonce);
  }
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach((value, index) => { map[String(value || "").trim()] = index; });
  return map;
}

function cellByHeader_(row, headerMap, candidates) {
  const keys = Object.keys(headerMap);
  for (const candidate of candidates) {
    if (headerMap[candidate] !== undefined) return String(row[headerMap[candidate]] || "").trim();
    const normalizedCandidate = normalize_(candidate).replace(/[’‘`]/g, "'");
    const matchedKey = keys.find(key => normalize_(key).replace(/[’‘`]/g, "'") === normalizedCandidate);
    if (matchedKey !== undefined) return String(row[headerMap[matchedKey]] || "").trim();
  }
  return "";
}

function guestMessageHtml_(data) {
  const safeData = JSON.stringify(data)
    .replace(/</g, "\\u003c").replace(/>/g, "\\u003e").replace(/&/g, "\\u0026");
  const html = "<!doctype html><meta charset=\"utf-8\"><script>window.parent.postMessage(" + safeData + ", \"" + WEDDING_SITE_ORIGIN + "\");</script>";
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWeddingSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty("WEDDING_SPREADSHEET_ID");
  if (!id) throw new Error("Спершу запустіть setupWeddingRsvp() з Google-таблиці.");
  return SpreadsheetApp.openById(id);
}

function normalize_(value) {
  return String(value || "").normalize("NFC").trim().replace(/\s+/g, " ").toLocaleLowerCase("uk");
}

function clean_(value, maxLength) {
  return String(value || "").replace(/[\u0000-\u001f]/g, " ").trim().slice(0, maxLength);
}

function constantTimeEqual_(a, b) {
  const left = String(a || "");
  const right = String(b || "");
  if (left.length !== right.length) return false;
  let difference = 0;
  for (let i = 0; i < left.length; i++) difference |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return difference === 0;
}

function htmlResponse_(ok, message, nonce) {
  const data = JSON.stringify({ type: "wedding-rsvp", ok: ok, message: message, nonce: nonce })
    .replace(/</g, "\\u003c").replace(/>/g, "\\u003e").replace(/&/g, "\\u0026");
  const html = "<!doctype html><meta charset=\"utf-8\"><script>window.parent.postMessage(" + data + ", \"*\");</script>";
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


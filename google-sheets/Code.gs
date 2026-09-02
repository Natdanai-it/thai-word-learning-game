/**
 * คลังคำไทย — ระบบรับคะแนนออนไลน์และแดชบอร์ดครู
 * สำคัญ: เปลี่ยนรหัสด้านล่างก่อน Deploy ทุกครั้ง
 */
const TEACHER_PIN = "2468";
const SCORE_SHEET_NAME = "คะแนนนักเรียน";
const MAX_DASHBOARD_ROWS = 5000;
const SCORE_HEADERS = [
  "รหัสรายการ",
  "เวลาที่รับข้อมูล",
  "เวลาจากเครื่องนักเรียน",
  "ชื่อ-นามสกุล",
  "ชั้น/ห้อง",
  "เลขที่",
  "รหัสด่าน",
  "บทเรียน",
  "โหมดเกม",
  "คะแนนที่ได้",
  "คะแนนเต็ม",
  "คำที่ยังไม่แม่น"
];
const ALLOWED_LESSON_IDS = ["choy", "legend", "song", "horse-poem"];
const ALLOWED_MODES = ["คลังคำ", "เติมบริบท", "ท้าทาย 30 วิ", "จับคู่ความหมาย", "AR ล่าคำ", "ทบทวนคำที่ผิด"];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Dashboard")
    .setTitle("แดชบอร์ดครู · คลังคำไทย")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function doPost(event) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return jsonResponse_({ status: "busy", message: "กรุณาลองส่งอีกครั้ง" });
  }

  try {
    const raw = JSON.parse((event && event.postData && event.postData.contents) || "{}");
    const payload = normalizePayload_(raw);
    const validationMessage = validatePayload_(payload);
    if (validationMessage) {
      return jsonResponse_({ status: "error", message: validationMessage });
    }

    const sheet = getScoreSheet_();
    if (isDuplicate_(sheet, payload.attemptId)) {
      return jsonResponse_({ status: "duplicate", attemptId: payload.attemptId });
    }

    sheet.appendRow([
      payload.attemptId,
      new Date(),
      payload.clientTime,
      payload.studentName,
      payload.classroom,
      payload.studentNumber,
      payload.lessonId,
      payload.lesson,
      payload.mode,
      payload.score,
      payload.total,
      payload.wrongWords
    ]);

    return jsonResponse_({ status: "ok", attemptId: payload.attemptId });
  } catch (error) {
    return jsonResponse_({ status: "error", message: "รูปแบบข้อมูลไม่ถูกต้อง" });
  } finally {
    lock.releaseLock();
  }
}

/** เรียกจาก Dashboard.html ด้วย google.script.run เท่านั้น */
function getDashboardData(pin) {
  verifyTeacherPin_(pin);
  const sheet = getScoreSheet_();
  const lastRow = sheet.getLastRow();
  const startRow = Math.max(2, lastRow - MAX_DASHBOARD_ROWS + 1);
  const rowCount = Math.max(0, lastRow - startRow + 1);
  const timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  if (!rowCount) {
    return { generatedAt: new Date().toISOString(), totalStored: 0, truncated: false, rows: [] };
  }

  const values = sheet.getRange(startRow, 1, rowCount, SCORE_HEADERS.length).getValues();
  const rows = values.reverse().map(function (row) {
    return {
      attemptId: displayText_(row[0]),
      receivedAt: dateText_(row[1], timezone),
      clientTime: dateText_(row[2], timezone),
      studentName: displayText_(row[3]),
      classroom: displayText_(row[4]),
      studentNumber: displayText_(row[5]),
      lessonId: displayText_(row[6]),
      lesson: displayText_(row[7]),
      mode: displayText_(row[8]),
      score: Number(row[9]) || 0,
      total: Number(row[10]) || 100,
      wrongWords: displayText_(row[11])
    };
  });

  return {
    generatedAt: new Date().toISOString(),
    totalStored: Math.max(0, lastRow - 1),
    truncated: lastRow - 1 > MAX_DASHBOARD_ROWS,
    rows: rows
  };
}

function normalizePayload_(raw) {
  const wrongWords = Array.isArray(raw.wrongWords) ? raw.wrongWords.join(" | ") : raw.wrongWords;
  return {
    attemptId: safeText_(raw.attemptId, 100),
    clientTime: safeText_(raw.clientTime, 60),
    studentName: safeText_(raw.studentName, 120),
    classroom: safeText_(raw.classroom, 30),
    studentNumber: safeText_(raw.studentNumber, 20),
    lessonId: safeText_(raw.lessonId, 30),
    lesson: safeText_(raw.lesson, 100),
    mode: safeText_(raw.mode, 50),
    score: Number(raw.score),
    total: Number(raw.total),
    wrongWords: safeText_(wrongWords, 800)
  };
}

function validatePayload_(payload) {
  if (!payload.attemptId || !payload.studentName || !payload.classroom || !payload.studentNumber) return "ข้อมูลนักเรียนไม่ครบ";
  if (ALLOWED_LESSON_IDS.indexOf(payload.lessonId) === -1) return "ไม่พบรหัสด่าน";
  if (ALLOWED_MODES.indexOf(payload.mode) === -1) return "ไม่พบโหมดเกม";
  if (!isFinite(payload.score) || !isFinite(payload.total) || payload.score < 0 || payload.total <= 0 || payload.score > payload.total || payload.total > 1000) return "คะแนนไม่ถูกต้อง";
  return "";
}

function verifyTeacherPin_(pin) {
  const supplied = String(pin == null ? "" : pin).trim();
  if (!supplied || supplied !== String(TEACHER_PIN)) {
    throw new Error("PIN ไม่ถูกต้อง");
  }
}

function getScoreSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("กรุณาเปิด Apps Script จาก Google Sheet ปลายทาง");

  let sheet = spreadsheet.getSheetByName(SCORE_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SCORE_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SCORE_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SCORE_HEADERS.length)
      .setBackground("#173b67")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
    sheet.getRange("B:C").setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.autoResizeColumns(1, SCORE_HEADERS.length);
  }
  return sheet;
}

function isDuplicate_(sheet, attemptId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  return Boolean(sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(attemptId)
    .matchEntireCell(true)
    .findNext());
}

function safeText_(value, maxLength) {
  const text = String(value == null ? "" : value).trim().slice(0, maxLength || 500);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function displayText_(value) {
  return String(value == null ? "" : value).replace(/^'(?=[=+\-@])/, "");
}

function dateText_(value, timezone) {
  if (!value) return "";
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return displayText_(value);
  return Utilities.formatDate(date, timezone, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

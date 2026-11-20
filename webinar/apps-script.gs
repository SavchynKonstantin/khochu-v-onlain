/**
 * Приймач реєстрацій на вебінар.
 * Вставляється в Google Таблицю: Розширення → Apps Script.
 * Кожна реєстрація з сайту = новий рядок на аркуші «Реєстрації».
 */

var SHEET_NAME = 'Реєстрації';
var HEADERS = ['Дата і час', "Ім'я", 'Телефон', 'utm_source', 'utm_medium', 'utm_campaign', 'Сторінка'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // щоб дві одночасні реєстрації не записались в один рядок
  try {
    var p = (e && e.parameter) || {};
    var name = clean(p.name, 80);
    var phone = clean(p.phone, 20);
    if (!name || !phone) return reply({ ok: false, error: 'empty' });

    var sheet = getSheet();
    sheet.appendRow([
      new Date(),
      safe(name),
      "'" + phone,               // апостроф: Google не перетворить номер на число
      safe(clean(p.utm_source, 100)),
      safe(clean(p.utm_medium, 100)),
      safe(clean(p.utm_campaign, 100)),
      safe(clean(p.page, 300))
    ]);
    return reply({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Відкрийте посилання /exec у браузері — побачите цей текст, значить все працює
function doGet() {
  return ContentService.createTextOutput('Приймач реєстрацій працює ✅');
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('dd.mm.yyyy hh:mm');
  }
  return sheet;
}

function clean(v, max) {
  return String(v || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
}

// Захист: текст, що починається з = + - @, Google сприйняв би як формулу
function safe(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

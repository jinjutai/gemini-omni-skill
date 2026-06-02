// ============================================================
//  Google Apps Script — Skill Download Tracker
//  วิธีใช้: คัดลอก code นี้ไปวางใน script.google.com
//  แล้ว Deploy → New deployment → Web app
// ============================================================

const SHEET_NAME = 'Downloads';

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // สร้าง sheet และ header ถ้ายังไม่มี
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Timestamp', 'IP', 'Country', 'Country Code', 'Region', 'City', 'User Agent']);
      sheet.setFrozenRows(1);
      // จัดสีหัวตาราง
      sheet.getRange(1, 1, 1, 7)
        .setBackground('#4f46e5')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    const p = e.parameter;
    sheet.appendRow([
      new Date(),
      p.ip           || 'Unknown',
      p.country      || 'Unknown',
      p.country_code || '',
      p.region       || 'Unknown',
      p.city         || 'Unknown',
      p.ua           || 'Unknown',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

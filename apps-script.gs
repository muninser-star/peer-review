// Код для Google Apps Script — принимает оценки со страницы и пишет в таблицу.
// Как установить — см. инструкцию ниже в чате.

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Ответы') || ss.insertSheet('Ответы');

    var HEADER = [
      'Время', 'Период', 'Оценивающий', 'Кого оценивают', 'Команда',
      'Оценка комфорта (1-10)', 'Соответствие роли',
      'Сильные стороны', 'Слабые стороны', 'Что улучшить'
    ];

    // Шапка: создаём, если её нет (надёжно — по первой ячейке)
    if (sheet.getRange('A1').getValue() !== 'Время') {
      if (sheet.getLastRow() > 0) sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
      sheet.setFrozenRows(1);
    }

    // По строке на каждого оценённого коллегу
    (data.rows || []).forEach(function (r) {
      sheet.appendRow([
        new Date(), data.period, data.evaluator, r.target, r.team,
        r.comfort_score, r.role, r.strong, r.weak, r.improve
      ]);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

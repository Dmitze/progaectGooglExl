// === Настройки для экспорта ===
const EXPORT_CONFIG = {
  SHEET_NAME: "2 Бат Загальна",
  RANGE: "A1:K27", // Включая шапку и данные!
  DOC_NAME: "2 Бат Загальна (експорт)",
  WORD_FILE_NAME: "2_Бат_Загальна.docx"
};

// === Основная функция экспорта ===
function exportSheetRangeToWord() {
  try {
    // 1. Открываем таблицу и нужный лист
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(EXPORT_CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`Лист "${EXPORT_CONFIG.SHEET_NAME}" не знайдено!`);

    // 2. Получаем значения диапазона
    const values = sheet.getRange(EXPORT_CONFIG.RANGE).getValues();

    // 3. Создаём Google Документ и вставляем таблицу
    const doc = DocumentApp.create(EXPORT_CONFIG.DOC_NAME);
    doc.getBody().appendTable(values);
    doc.saveAndClose();

    // 4. Экспортируем Google Документ как .docx (Word)
    const token = ScriptApp.getOAuthToken();
    const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${doc.getId()}&exportFormat=docx`;
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const blob = response.getBlob().setName(EXPORT_CONFIG.WORD_FILE_NAME);
    const file = DriveApp.createFile(blob);

    // 5. Показываем ссылку пользователю
    showWordExportSuccessDialog(file.getUrl());
  } catch (error) {
    showWordExportErrorDialog(error.message);
  }
}

// === Вспомогательные диалоговые окна ===
function showWordExportSuccessDialog(url) {
  const html = HtmlService
    .createHtmlOutput(`<div style="font-size:1.2em;">
      ✅ <b>Word-файл створено!</b><br>
      <a href="${url}" target="_blank">Завантажити файл</a>
    </div>`)
    .setWidth(350)
    .setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, "Експорт до Word");
}

function showWordExportErrorDialog(message) {
  const html = HtmlService
    .createHtmlOutput(`<div style="color:#b71c1c;font-size:1.2em;">
      ❌ Помилка експорту:<br>
      <span style="white-space:pre-line;">${message}</span>
    </div>`)
    .setWidth(380)
    .setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, "Помилка експорту");
}

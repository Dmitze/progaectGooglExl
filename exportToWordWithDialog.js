// === КОНФИГУРАЦИЯ ===
const DEFAULT_DOC_NAME = "Експортований лист";
const DEFAULT_WORD_FILE_NAME = "ExportedSheet.docx";

// === Показать диалог выбора листа и діапазона ===
function showExportToWordDialog() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const sheetOptions = sheets.map(s => `<option value="${s.getName()}">${s.getName()}</option>`).join("");
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif;">
      <h3>Експорт до Word</h3>
      <label>Лист:</label>
      <select id="sheetName">${sheetOptions}</select>
      <br><br>
      <label>Діапазон (наприклад, A1:K27):</label>
      <input type="text" id="range" value="A1:K27" style="width:120px;">
      <br><br>
      <label>Ім'я Word-файла:</label>
      <input type="text" id="wordName" value="${DEFAULT_WORD_FILE_NAME}" style="width:180px;">
      <br><br>
      <button onclick="exportNow()" style="font-size:1.1em;">Експортувати</button>
      <div id="status" style="margin-top:15px;"></div>
      <script>
        function exportNow() {
          const sheet = document.getElementById('sheetName').value;
          const range = document.getElementById('range').value;
          const wordName = document.getElementById('wordName').value;
          document.getElementById('status').innerHTML = '⏳ Зачекайте...';
          google.script.run
            .withSuccessHandler(url => {
              document.getElementById('status').innerHTML =
                '<b>✅ Word-файл створено!</b><br><a href="'+url+'" target="_blank">Завантажити</a>';
            })
            .withFailureHandler(err => {
              document.getElementById('status').innerHTML =
                '<b style="color:red;">❌ ' + (err.message || err) + '</b>';
            })
            .exportSheetRangeToWordCustom(sheet, range, wordName);
        }
      </script>
    </div>
  `).setWidth(400).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, "Експорт до Word");
}

// === Основная функция экспорта (вызывается из диалога) ===
function exportSheetRangeToWordCustom(sheetName, rangeA1, wordFileName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Лист "${sheetName}" не знайдено!`);
    const values = sheet.getRange(rangeA1).getValues();
    if (!values || !values.length) throw new Error("Діапазон порожній або невірний!");

    const doc = DocumentApp.create(wordFileName.replace(/\.docx$/i, ''));
    doc.getBody().appendTable(values);
    doc.saveAndClose();

    const token = ScriptApp.getOAuthToken();
    const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${doc.getId()}&exportFormat=docx`;
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const blob = response.getBlob().setName(wordFileName.endsWith('.docx') ? wordFileName : wordFileName + ".docx");
    const file = DriveApp.createFile(blob);

    return file.getUrl();
  } catch (e) {
    throw new Error(e && e.message ? e.message : e);
  }
}

// Открытие HTML-формы
function showWordExportFullForm() {
  const html = HtmlService.createHtmlOutputFromFile('WordExportForm')
    .setWidth(520).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, "Генератор Word-звіту");
}

// Получение списка всех листов для формы
function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

// Получение данных диапазона для предварительного просмотра
function getPreviewData(sheetName, rangeA1) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getRange(rangeA1).getValues();
}

// Основная функция генерации Word-файла
function generateWordReport(formData) {
  /*
    formData = {
      header: { boss, date, order },
      title,
      description,
      tables: [
        { sheet: ..., range: ... },
        ...
      ]
    }
  */
  const doc = DocumentApp.create('Word звіт');
  const body = doc.getBody();

  // Шапка
  body.appendParagraph(`Начальник: ${formData.header.boss}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(`Дата: ${formData.header.date}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(`Наказ: ${formData.header.order}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(''); // пробел

  // Заголовок
  body.appendParagraph(formData.title).setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // Описание
  body.appendParagraph(formData.description).setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // Таблицы
  formData.tables.forEach((t, idx) => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheet);
    if (!sheet) return;
    const values = sheet.getRange(t.range).getValues();
    body.appendParagraph(`Таблиця ${idx+1}: ${t.sheet} ${t.range}`);
    body.appendTable(values);
    body.appendParagraph(''); // пробел
  });

  doc.saveAndClose();

  // Экспортируем в .docx
  const token = ScriptApp.getOAuthToken();
  const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${doc.getId()}&exportFormat=docx`;
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob = response.getBlob().setName("WordReport.docx");
  const file = DriveApp.createFile(blob);

  return file.getUrl();
}

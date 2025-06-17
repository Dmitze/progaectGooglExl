/**
 * Контроль орфографії і форматів у ключових полях + орфографія через LanguageTool
 * Групування помилок по листах, перевірка унікальності, діапазонів дат, орфографія для "Опис"
 */

const VALIDATION_FIELDS = [
  { name: "Email", regex: /^[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}$/, message: "Некоректний email", unique: true },
  { name: "Дата", regex: /^\d{4}-\d{2}-\d{2}$/, message: "Дата повинна бути у форматі РРРР-ММ-ДД", dateRange: { min: "2024-01-01", max: "2025-12-31" } },
  { name: "Телефон", regex: /^\+?\d{10,15}$/, message: "Некоректний номер телефону", unique: false },
  { name: "ID", regex: /^[A-Za-z0-9\-]+$/, message: "Некоректний ID", unique: true }
];

// Додає пункт меню "Валідація"
function addValidationMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Валідація")
    .addItem("Перевірити орфографію/формати", "runValidation")
    .addToUi();
}

// Головна функція перевірки
function runValidation() {
  const ui = SpreadsheetApp.getUi();
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let issuesBySheet = {};
  let checkedRows = 0;
  // Для перевірки унікальності
  const uniqueValues = {};
  VALIDATION_FIELDS.filter(f => f.unique).forEach(f => {
    uniqueValues[f.name] = {};
  });

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0].map(h => h.trim());
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      checkedRows++;
      VALIDATION_FIELDS.forEach(field => {
        const idx = headers.indexOf(field.name);
        if (idx >= 0) {
          const value = (row[idx] || '').toString().trim();
          // Порожнє
          if (!value) {
            addIssue(issuesBySheet, sheet.getName(), `Рядок ${r+1}: поле "${field.name}" порожнє`);
          } else {
            // Формат
            if (!field.regex.test(value)) {
              addIssue(issuesBySheet, sheet.getName(), `Рядок ${r+1}: поле "${field.name}" — ${field.message} ("${value}")`);
            }
            // Перевірка діапазону дат
            if (field.dateRange && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
              const dateVal = new Date(value);
              const minDate = new Date(field.dateRange.min);
              const maxDate = new Date(field.dateRange.max);
              if (dateVal < minDate || dateVal > maxDate) {
                addIssue(issuesBySheet, sheet.getName(), `Рядок ${r+1}: дата "${value}" поза дозволеним діапазоном (${field.dateRange.min} — ${field.dateRange.max})`);
              }
            }
          }
          // Для унікальних полів: збір даних для перевірки дублювання
          if (field.unique && value) {
            if (!uniqueValues[field.name][value]) uniqueValues[field.name][value] = [];
            uniqueValues[field.name][value].push({ sheet: sheet.getName(), row: r+1 });
          }
        }
      });

      // === Орфографія для "Опис" ===
      const descIdx = headers.indexOf("Опис");
      if (descIdx >= 0) {
        const desc = (row[descIdx] || '').toString();
        if (desc.length > 0) {
          // Для теста можно ограничить количество проверок, если строк много!
          try {
            const spellingErrors = checkSpellingWithLanguageTool(desc, "uk");
            spellingErrors.forEach(err => {
              addIssue(issuesBySheet, sheet.getName(), `Рядок ${r+1}: орфографія ("${desc.substr(err.offset, err.length)}"): ${err.message}. Варіанти: ${err.replacements.join(', ')}`);
            });
          } catch(e) {
            addIssue(issuesBySheet, sheet.getName(), `Рядок ${r+1}: помилка при перевірці орфографії (API): ${e.message}`);
          }
        }
      }
    }
  });

  // Перевірка дублювання унікальних полів
  Object.keys(uniqueValues).forEach(fieldName => {
    Object.keys(uniqueValues[fieldName]).forEach(val => {
      if (uniqueValues[fieldName][val].length > 1) {
        const places = uniqueValues[fieldName][val].map(e => `[${e.sheet} рядок ${e.row}]`).join(', ');
        const firstSheet = uniqueValues[fieldName][val][0].sheet;
        addIssue(issuesBySheet, firstSheet, `Дубльоване значення "${val}" у полі "${fieldName}": ${places}`);
      }
    });
  });

  // Формування результату
  const totalIssues = Object.values(issuesBySheet).reduce((a, b) => a + b.length, 0);
  if (totalIssues === 0) {
    ui.alert(`✅ Всі ключові поля у ${checkedRows} рядках у порядку!`);
  } else {
    let html = `<div style="font-family:monospace;font-size:13px;max-height:500px;overflow:auto;color:#8b0000;"><b>Знайдено проблем: ${totalIssues}</b><br>`;
    Object.keys(issuesBySheet).forEach(sheetName => {
      html += `<b>${sheetName}:</b><ul><li>${issuesBySheet[sheetName].join('</li><li>')}</li></ul>`;
    });
    html += `</div>`;
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(700).setHeight(600), "Проблеми валідації");
  }
}

// Групування помилок по листу
function addIssue(obj, sheet, issue) {
  if (!obj[sheet]) obj[sheet] = [];
  obj[sheet].push(issue);
}

/**
 * Проверяет орфографию текста через LanguageTool API
 * @param {string} text - текст для проверки
 * @param {string} lang - язык (например, "uk" для украинского, "ru" для русского, "en-US" для английского)
 * @returns {Array} массив ошибок [{offset, length, message, replacements[]}]
 */
function checkSpellingWithLanguageTool(text, lang) {
  const apiUrl = "https://api.languagetoolplus.com/v2/check";
  const payload = {
    text: text,
    language: lang || "uk"
  };
  const options = {
    method: "post",
    payload: payload,
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(apiUrl, options);
  if (response.getResponseCode() === 200) {
    const result = JSON.parse(response.getContentText());
    return result.matches.map(m => ({
      offset: m.offset,
      length: m.length,
      message: m.message,
      replacements: (m.replacements || []).map(r => r.value)
    }));
  }
  return [];
}

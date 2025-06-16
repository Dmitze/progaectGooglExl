

// === Диалог восстановления по истории ===
function showRestoreHistoryDialog() {
  const html = HtmlService.createHtmlOutputFromFile('history_restore_dialog')
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Відновлення даних за історією");
}

// === Получить список листов и дат, где были изменения ===
function getHistoryRestoreData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(sh => sh.getName());

  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return {sheets, datesBySheet:{}};

  const data = logSheet.getDataRange().getValues();
  const datesBySheet = {};

  for (let i = 1; i < data.length; i++) {
    const sheet = data[i][2];
    const dt = data[i][0];
    if (!sheet || !dt) continue;

    const iso = new Date(dt).toISOString();
    if (!datesBySheet[sheet]) datesBySheet[sheet] = [];
    if (!datesBySheet[sheet].includes(iso)) datesBySheet[sheet].push(iso);
  }

  for (const sh in datesBySheet) {
    datesBySheet[sh].sort((a, b) => new Date(b) - new Date(a));
  }

  return {sheets, datesBySheet};
}

// === Предпросмотр таблицы на определенную дату с пагинацией ===
function getSheetPreviewOnDate(sheetName, restoreIsoString, page, pageSize) {
  page = Math.max(1, Number(page) || 1);
  pageSize = Math.max(1, Number(pageSize) || 30);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { preview: [], totalRows: 0, totalCols: 0 };

  const vals = sheet.getDataRange().getValues();
  const totalRows = vals.length;
  const totalCols = vals[0]?.length || 0;

  if (totalRows === 0 || totalCols === 0) {
    return { preview: [], totalRows: 0, totalCols: 0 };
  }

  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  let preview = vals.map(row => row.slice());

  if (logSheet) {
    const data = logSheet.getDataRange().getValues();
    const restoreDate = new Date(restoreIsoString);
    const previewCells = {};
    
    for (let i = 1; i < data.length; i++) {
      const dt = new Date(data[i][0]);
      if (data[i][2] === sheetName && dt <= restoreDate) {
        const addr = data[i][3], oldV = data[i][5];
        previewCells[addr] = oldV;
      }
    }

    Object.keys(previewCells).forEach(addr => {
      try {
        const rng = sheet.getRange(addr);
        const r = rng.getRow() - 1, c = rng.getColumn() - 1;
        if (r >= 0 && c >= 0 && r < preview.length && c < preview[0].length) {
          preview[r][c] = previewCells[addr];
        }
      } catch (e) {}
    });
  }

  const startRow = Math.max(0, (page - 1) * pageSize);
  const endRow = Math.min(startRow + pageSize, totalRows);
  const paginated = preview.slice(startRow, endRow);

  return {
    preview: paginated,
    totalRows,
    totalCols
  };
}

// === Восстановление данных на выбранную дату ===
function restoreSheetToDate(sheetName, restoreIsoString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) throw new Error('Лист "Лог змін" не знайдено!');

  const data = logSheet.getDataRange().getValues();
  const restoreDate = new Date(restoreIsoString);
  const changes = [];

  for (let i = 1; i < data.length; i++) {
    const dt = new Date(data[i][0]);
    if (data[i][2] === sheetName && dt <= restoreDate) {
      const addr = data[i][3], oldV = data[i][5];
      changes.push({dt, addr, oldV});
    }
  }

  const lastCellValues = {};
  for (const ch of changes) {
    lastCellValues[ch.addr] = ch.oldV;
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Лист не знайдено!");

  for (let addr in lastCellValues) {
    try {
      sheet.getRange(addr).setValue(lastCellValues[addr]);
    } catch (e) {}
  }

  return Object.keys(lastCellValues).length;
}

// === Вызов из формы: запуск восстановления ===
function restoreHistoryDialogAction(sheetName, restoreIsoString) {
  try {
    const restored = restoreSheetToDate(sheetName, restoreIsoString);
    return `Відновлено ${restored} комірок листа "${sheetName}". Перевірте дані!`;
  } catch (e) {
    return "Помилка: " + e.message;
  }
}

// === Экспорт в Excel ===
function exportSheetToExcel(sheetName, restoreIsoString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Лист не знайдено!");

  const preview = getSheetPreviewOnDate(sheetName, restoreIsoString, 1, sheet.getLastRow());
  const tempSS = SpreadsheetApp.create(`Export_${sheetName}_${new Date().toISOString()}`);
  const tempSheet = tempSS.getSheets()[0];
  tempSheet.clear();
  tempSheet.getRange(1, 1, preview.preview.length, preview.preview[0]?.length || 0).setValues(preview.preview);
  
  tempSS.getSheets().forEach(s => {
    if (s.getName() !== tempSheet.getName()) tempSS.deleteSheet(s);
  });

  const blob = tempSS.getBlob().setContentType(MimeType.MICROSOFT_EXCEL);
  const file = DriveApp.createFile(blob);
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);

  return file.getUrl();
}

// === Сравнение текущего состояния с сохранённым ===
function getSheetDiffOnDate(sheetName, restoreIsoString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) throw new Error('Лист "Лог змін" не знайдено!');

  const data = logSheet.getDataRange().getValues();
  const restoreDate = new Date(restoreIsoString);
  const previewCells = {};

  for (let i = 1; i < data.length; i++) {
    const dt = new Date(data[i][0]);
    if (data[i][2] === sheetName && dt <= restoreDate) {
      const addr = data[i][3], oldV = data[i][5];
      previewCells[addr] = oldV;
    }
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Лист не знайдено!");

  const diff = [];
  Object.keys(previewCells).forEach(addr => {
    try {
      const rng = sheet.getRange(addr);
      const currentV = rng.getValue();
      const oldV = previewCells[addr];
      if ((currentV ?? "") !== (oldV ?? "")) {
        diff.push({
          addr: addr,
          oldValue: oldV,
          newValue: currentV
        });
      }
    } catch (e) {}
  });

  return diff;
}

// === Создание листа "Лог змін", если его нет ===
function setupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    const headers = [[
      "Час зміни",
      "Користувач",
      "Аркуш",
      "Адреса",
      "Тип дії",
      "Було",
      "Стало",
      "Формула",
      "Важлива зміна"
    ]];
    logSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    logSheet.autoResizeColumns(1, headers[0].length);
  }
}

// === Настройки ===
const SHEET_NAMES = [
    "2 Бат Загальна", "Ударні БпЛА", "Розвідувальні БпЛА", "НРК", "ППО",
    "НСО БТ", "АТ", "Засоби ураження", "ЗББ та Р", "РЕБ", "Оптика", "РЛС"
  ];
  const LOG_SHEET_NAME = "Лог змін";
  const COLOR_GREEN = "#b6d7a8";
  const IMPORTANT_RANGES = {
    "2 Бат Загальна": ["A1:C5"],
    "АТ": ["B2:D6"]
  };
  
  // === Меню при відкритті файлу ===
  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("Дії з таблицею")
      .addItem("Ручна перевірка змін", "checkChanges")
      .addItem("Звіт по діям користувачів", "showUsersActionReport")
      .addSeparator()
      .addItem("Відновити дані за історією", "showRestoreHistoryDialog")
      .addItem("Пошук по історії змін", "showHistorySearchDialog")
      .addSeparator()
      .addItem("Оновити дашборд", "createOrUpdateDashboard")
      .addItem("Додати коментар до комірки", "showAddCommentDialog")
      .addItem("Переглянути коментарі", "showCommentsDialog")
      .addSeparator()
      .addItem("Перевірити орфографію/формати", "runValidation")
      .addItem("Створити публічний звіт (копію)", "createPublicReport")
      .addSeparator()
      // Новые пункты: экспорт и архивация логов
      .addItem("Експорт логу у Excel", "exportLogSheetAsExcel")
      .addItem("Експорт логу у CSV", "exportLogSheetAsCSV")
      .addItem("Архівація логів", "archiveLogHistory")
      .addItem("Створити тригер на архівацію", "createDailyArchiveTrigger")
      .addItem("Видалити старі бэкапи", "cleanupOldBackups");
  
    menu.addToUi();
  
    setupLogSheet();
  }
  
  // === Основна перевірка змін ===
  function checkChanges() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    SHEET_NAMES.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      if (!Array.isArray(values)) return;
      const currentHash = JSON.stringify(values);
  
      const storedHashKey = `prevDataHash_${sheetName}`;
      const storedValuesKey = `prevValues_${sheetName}`;
      const storedHash = props.getProperty(storedHashKey);
  
      let oldValues = [];
      if (storedHash) {
        const old = props.getProperty(storedValuesKey);
        oldValues = old ? JSON.parse(old) : values.map(row => row.map(() => null));
      } else {
        oldValues = values.map(row => row.map(() => null));
      }
  
      // Логування змін значень
      if (
        storedHash &&
        storedHash !== currentHash &&
        Array.isArray(values) &&
        Array.isArray(oldValues) &&
        values.length > 0 &&
        oldValues.length > 0
      ) {
        highlightChanges(sheet, oldValues, values);
        logChanges(sheet, oldValues, values);
      }
  
      // Логування додавання/видалення рядків
      if (
        Array.isArray(oldValues) &&
        Array.isArray(values) &&
        oldValues.length !== values.length
      ) {
        const type = oldValues.length < values.length ? "Додано рядок" : "Видалено рядок";
        logRowOrColumnAction(sheet, type, oldValues.length, values.length);
      }
  
      // Логування додавання/видалення стовпців
      if (
        Array.isArray(oldValues) && Array.isArray(values) &&
        oldValues.length > 0 && values.length > 0 &&
        oldValues[0].length !== values[0].length
      ) {
        const type = oldValues[0].length < values[0].length ? "Додано стовпець" : "Видалено стовпець";
        logRowOrColumnAction(sheet, type, oldValues[0].length, values[0].length);
      }
  
      props.setProperty(storedHashKey, currentHash);
      props.setProperty(storedValuesKey, JSON.stringify(values));
    });
  }
  
  // === Підсвітка змінених комірок ===
  function highlightChanges(sheet, oldValues, newValues) {
    if (!Array.isArray(newValues) || !Array.isArray(oldValues)) return;
    for (let row = 0; row < newValues.length; row++) {
      for (let col = 0; col < newValues[row].length; col++) {
        const oldValue = (oldValues[row] || [])[col];
        const newValue = newValues[row][col];
        if (oldValue !== newValue) {
          const cell = sheet.getRange(row + 1, col + 1);
          // Додаємо кольори відповідно до типу зміни
          if ((oldValue === "" || oldValue === null) && newValue !== "") {
            // Було пусто -> стало щось (додавання)
            cell.setBackground("#b6d7a8"); // Зелений
          } else if (oldValue !== "" && (newValue === "" || newValue === null)) {
            // Було щось -> стало пусто (видалення)
            cell.setBackground("#ea9999"); // Червоний
          } else {
            // Будь-яка інша зміна (оновлення)
            cell.setBackground("#ffe599"); // Жовтий
          }
        }
      }
    }
  }
  
  // === Логування змін значень з типом дії ===
  function logChanges(sheet, oldValues, newValues) {
    if (!sheet || !Array.isArray(newValues) || !Array.isArray(oldValues)) return;
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const user = Session.getActiveUser().getEmail();
    const time = new Date();
    let changes = [];
    for (let row = 0; row < newValues.length; row++) {
      const newRow = newValues[row] || [];
      const oldRow = oldValues[row] || [];
      for (let col = 0; col < newRow.length; col++) {
        const oldValue = (oldRow[col] !== undefined ? oldRow[col] : "");
        const newValue = (newRow[col] !== undefined ? newRow[col] : "");
        if (oldValue !== newValue) {
          const cell = sheet.getRange(row + 1, col + 1);
          let formula = cell.getFormula();
          if (formula) formula = "=" + formula;
          const important = isImportantCell(sheet.getName(), row + 1, col + 1) ? "Так" : "Ні";
          // Тип дії для кожної зміни:
          let changeType = "";
          if ((oldValue === "" || oldValue === null) && newValue !== "") {
            changeType = "Додано значення";
          } else if (oldValue !== "" && (newValue === "" || newValue === null)) {
            changeType = "Видалено значення";
          } else {
            changeType = "Змінено";
          }
          changes.push([
            time,
            user,
            sheet.getName(),
            cell.getA1Notation(),
            changeType,
            oldValue,
            newValue,
            formula || "",
            important
          ]);
        }
      }
    }
    if (changes.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, changes.length, 9).setValues(changes);
    }
  }
  
  // === Логування додавання/видалення рядків/стовпців ===
  function logRowOrColumnAction(sheet, type, oldLen, newLen) {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const user = Session.getActiveUser().getEmail();
    const time = new Date();
    let actionDesc = "";
    let sheetName = (sheet && typeof sheet.getName === "function") ? sheet.getName() : "[невідомий лист]";
  
    if (type === "Додано рядок") {
      actionDesc = `Було ${oldLen}, стало ${newLen}`;
    } else if (type === "Видалено рядок") {
      actionDesc = `Було ${oldLen}, стало ${newLen}`;
    } else if (type === "Додано стовпець") {
      actionDesc = `Було ${oldLen}, стало ${newLen}`;
    } else if (type === "Видалено стовпець") {
      actionDesc = `Було ${oldLen}, стало ${newLen}`;
    } else {
      actionDesc = `Невідомий тип зміни: ${type}`;
    }
    logSheet.appendRow([
      time,
      user,
      sheetName,
      "",
      type,
      actionDesc,
      "",
      "",
      ""
    ]);
  }
  
  // === Визначення важливих комірок ===
  function isImportantCell(sheetName, row, col) {
    if (!IMPORTANT_RANGES[sheetName]) return false;
    for (const rangeStr of IMPORTANT_RANGES[sheetName]) {
      const range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(rangeStr);
      if (
        row >= range.getRow() &&
        row < range.getRow() + range.getNumRows() &&
        col >= range.getColumn() &&
        col < range.getColumn() + range.getNumColumns()
      ) {
        return true;
      }
    }
    return false;
  }
  
  // === Створення аркуша для логів ===
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
  
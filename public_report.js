
function createPublicReport() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    const dateStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd_HH-mm");
    const publicName = ss.getName() + ` [Публічний звіт ${dateStr}]`;
    const newFile = file.makeCopy(publicName);
    const newSs = SpreadsheetApp.openById(newFile.getId());

    // Додаємо новий sheet для звіту
    const mainSheet = newSs.insertSheet("Публічний звіт");

    // Видаляємо всі інші sheet-и
    newSs.getSheets().forEach(s => {
      if (s.getSheetName() !== mainSheet.getSheetName()) newSs.deleteSheet(s);
    });

    // Генеруємо і записуємо звіт
    const report = buildPublicReport();
    if (!report || !report.summary) {
      Logger.log("Неможливо згенерувати звіт: buildPublicReport() повернув порожній об'єкт");
      mainSheet.getRange(1, 1).setValue("Дані для звіту відсутні.");
    } else {
      fillReportSheet(mainSheet, report);
    }

    // Відкрити доступ "тільки для читання для всіх"
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Показати посилання користувачу
    try {
      ui.alert("Публічний звіт створено!\n\nПерейдіть сюди: " + newFile.getUrl());
    } catch (e) {
      Logger.log("URL публічного звіту: " + newFile.getUrl());
    }
  } catch (e) {
    try {
      SpreadsheetApp.getUi().alert("Помилка при створенні звіту: " + e.message);
    } catch (err) {
      Logger.log("Помилка при створенні звіту: " + e.message);
    }
  }
}

/**
 * Функція для відкриття HTML-дашборду у діалозі
 */
function showPublicReportDialog() {
  var html = HtmlService.createHtmlOutputFromFile('public_report_dialog')
    .setWidth(950).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "Публічний звіт (дашборд)");
}

/**
 * Головний генератор звіту для таблиці та HTML-дашборду
 */
function buildPublicReport() {
  const logs = getAllHistoryLogs();
  if (!logs || logs.length === 0) {
    return {
      summary: [["Дані відсутні"]],
      chartHeader: [],
      chartData: [],
      topSheetsHeader: [],
      topSheets: [],
      actionsHeader: [],
      actions: [],
      lastChangesHeader: [],
      lastChanges: [],
      description: [["Немає опису"]],
      notes: [["Звіт пустий"]]
    };
  }
  const meta = getSheetMeta();
  const stat = calcPublicStats(logs);
  const topSheets = Object.entries(stat.sheetStats)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  // Графік змін (останні 14 днів)
  const dailyData = [];
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  for (let i = 13; i >= 0; i--) {
    const d = new Date(now.getTime() - i * 24 * 60 * 60 * 1000);
    const dateStr = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    dailyData.push([dateStr, stat.dailyStats[dateStr] || 0]);
  }

  // Останні зміни (10 записів)
  const lastChanges = logs.slice(-10).map(log => [
    log.date || log.dateTime || "",
    log.sheet || "",
    log.action || "",
    log.address || "",
    log.oldValue ? "..." : "",
    log.newValue ? "..." : ""
  ]);

  return {
    summary: [
      ["Зведення / Статистика"],
      ["Всього змін", stat.total],
      ["Унікальних аркушів", stat.uniqueSheets],
      ["Типів змін", Object.keys(stat.actionStats).length],
      ["Остання зміна", stat.lastChangeDate || ""],
      []
    ],
    chartHeader: [["Графік змін за останні 14 днів"], ["Дата", "Кількість змін"]],
    chartData: dailyData,
    topSheetsHeader: [["Топ аркушів"], ["Аркуш", "Кількість змін"]],
    topSheets: topSheets,
    actionsHeader: [["Типи змін"], ["Тип", "Кількість"]],
    actions: Object.entries(stat.actionStats).sort((a, b) => b[1] - a[1]),
    lastChangesHeader: [["Останні зміни (без ідентифікаторів)"], ["Дата", "Аркуш", "Дія", "Адреса", "Було", "Стало"]],
    lastChanges: lastChanges,
    description: [
      ["Опис таблиці / Контакти"],
      [meta.description || "Ця таблиця — агрегований публічний звіт активності без персональних даних."],
      ["Відповідальний:", meta.contact || "—"]
    ],
    notes: [
      ["Примітки"],
      ["У цьому звіті видалені чутливі поля та індивідуальні ідентифікатори користувачів."],
      ["Email-и та коментарі не включені."],
      ["Більше інформації — у приватній версії."]
    ]
  };
}

/**
 * Записує всі блоки у аркуш
 */
function fillReportSheet(sheet, report) {
  let row = 1;
  function writeBlock(block) {
    block.forEach(arr => { sheet.getRange(row++, 1, 1, arr.length).setValues([arr]); });
    row++;
  }
  writeBlock(report.summary);
  writeBlock(report.chartHeader);
  report.chartData.forEach(arr => { sheet.getRange(row++, 1, 1, arr.length).setValues([arr]); });
  row++;
  writeBlock(report.topSheetsHeader);
  report.topSheets.forEach(arr => { sheet.getRange(row++, 1, 1, arr.length).setValues([arr]); });
  row++;
  writeBlock(report.actionsHeader);
  report.actions.forEach(arr => { sheet.getRange(row++, 1, 1, arr.length).setValues([arr]); });
  row++;
  writeBlock(report.lastChangesHeader);
  report.lastChanges.forEach(arr => { sheet.getRange(row++, 1, 1, arr.length).setValues([arr]); });
  row++;
  writeBlock(report.description);
  row++;
  writeBlock(report.notes);

  // Додаємо графік
  try {
    const chartStart = findHeaderRow(sheet, "Графік змін за останні 14 днів") + 1;
    const chartEnd = chartStart + report.chartData.length - 1;
    const chartRange = sheet.getRange(chartStart, 1, report.chartData.length, 2);
    const chart = sheet.newChart()
      .addRange(chartRange)
      .setChartType(Charts.ChartType.LINE)
      .setOption('title', 'Графік змін за останні 14 днів')
      .setPosition(chartEnd + 2, 1, 0, 0)
      .build();
    sheet.insertChart(chart);
  } catch (e) {
    Logger.log("Графік не створено: " + e.message);
  }

  sheet.autoResizeColumns(1, 6);
}

/**
 * Пошук рядка з заголовком для розміщення графіка
 */
function findHeaderRow(sheet, header) {
  if (!sheet) return 1;
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().indexOf(header) > -1) return i + 1;
  }
  return 1;
}

/**
 * Повертає мета-інфо (опис, контакт)
 */
function getSheetMeta() {
  // Можна зробити окремий службовий аркуш з мета-інфо або прописати тут вручну
  return {
    description: "Таблиця для агрегованого моніторингу змін у бойових та сервісних аркушах.",
    contact: "info@yourdomain.com"
  };
}

/**
 * Підрахунок статистики для звіту
 */
function calcPublicStats(logs) {
  if (!logs || logs.length === 0) {
    return {
      total: 0, uniqueSheets: 0, lastChangeDate: "", sheetStats: {}, actionStats: {}, dailyStats: {}
    };
  }
  const stats = {
    total: logs.length,
    uniqueSheets: 0,
    lastChangeDate: "",
    sheetStats: {},
    actionStats: {},
    dailyStats: {}
  };
  let lastDate = null;
  logs.forEach(log => {
    if (!log.sheet) return;
    stats.sheetStats[log.sheet] = (stats.sheetStats[log.sheet] || 0) + 1;
    stats.actionStats[log.action] = (stats.actionStats[log.action] || 0) + 1;
    // Дата для dailyStats
    let dateStr = "";
    if (log.date) dateStr = log.date.split(" ")[0];
    else if (log.dateTime) dateStr = log.dateTime.split(" ")[0];
    if (dateStr) stats.dailyStats[dateStr] = (stats.dailyStats[dateStr] || 0) + 1;
    // Остання дата
    const d = log.dateTime ? new Date(log.dateTime) : (log.date ? new Date(log.date) : null);
    if (d && (!lastDate || d > lastDate)) lastDate = d;
  });
  stats.uniqueSheets = Object.keys(stats.sheetStats).length;
  stats.lastChangeDate = lastDate ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : "";
  return stats;
}

/**
 * Отримання логів для звіту
 */
function getAllHistoryLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Лог змін");
  if (!logSheet) return [];
  const data = logSheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const log = {};
    headers.forEach((key, i) => log[key] = row[i]);
    return log;
  });
}

// === Тестові та дебаг-функції ===

function testPublicReport() {
  const fakeLogs = [
    { date: "2025-06-16 09:00:00", sheet: "Sheet1", action: "Змінено", address: "A1", oldValue: "a", newValue: "b" },
    { date: "2025-06-16 10:00:00", sheet: "Sheet1", action: "Додано значення", address: "B2", oldValue: "", newValue: "c" },
    { date: "2025-06-15 11:00:00", sheet: "Sheet2", action: "Видалено значення", address: "C3", oldValue: "d", newValue: "" },
    { date: "2025-06-15 12:00:00", sheet: "Sheet1", action: "Змінено", address: "A2", oldValue: "e", newValue: "f" },
    { date: "2025-06-14 13:00:00", sheet: "Sheet3", action: "Змінено", address: "D4", oldValue: "g", newValue: "h" }
  ];
  const stat = calcPublicStats(fakeLogs);
  Logger.log("СТАТИСТИКА ЗВІТУ:");
  Logger.log(stat);
  const report = buildPublicReportWithLogs(fakeLogs);
  Logger.log("Звіт:");
  Logger.log(report);
}

function buildPublicReportWithLogs(logs) {
  const meta = getSheetMeta();
  const stat = calcPublicStats(logs);
  const topSheets = Object.entries(stat.sheetStats)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  const dailyData = [];
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  for (let i = 13; i >= 0; i--) {
    const d = new Date(now.getTime() - i * 24 * 60 * 60 * 1000);
    const dateStr = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    dailyData.push([dateStr, stat.dailyStats[dateStr] || 0]);
  }
  const lastChanges = logs.slice(-10).map(log => [
    log.date || log.dateTime || "",
    log.sheet || "",
    log.action || "",
    log.address || "",
    log.oldValue ? "..." : "",
    log.newValue ? "..." : ""
  ]);
  return {
    summary: [
      ["Зведення / Статистика"],
      ["Всього змін", stat.total],
      ["Унікальних аркушів", stat.uniqueSheets],
      ["Типів змін", Object.keys(stat.actionStats).length],
      ["Остання зміна", stat.lastChangeDate || ""],
      []
    ],
    chartHeader: [["Графік змін за останні 14 днів"], ["Дата", "Кількість змін"]],
    chartData: dailyData,
    topSheetsHeader: [["Топ аркушів"], ["Аркуш", "Кількість змін"]],
    topSheets: topSheets,
    actionsHeader: [["Типи змін"], ["Тип", "Кількість"]],
    actions: Object.entries(stat.actionStats).sort((a, b) => b[1] - a[1]),
    lastChangesHeader: [["Останні зміни (без ідентифікаторів)"], ["Дата", "Аркуш", "Дія", "Адреса", "Було", "Стало"]],
    lastChanges: lastChanges,
    description: [
      ["Опис таблиці / Контакти"],
      [meta.description || "Тестовий звіт"],
      ["Відповідальний:", meta.contact || "—"]
    ],
    notes: [
      ["Примітки"],
      ["У цьому звіті видалені чутливі поля та індивідуальні ідентифікатори."],
      ["Email-и та коментарі не включені."],
      ["Більше інформації — у приватній версії."]
    ]
  };
}

function debugPublicReportData() {
  const logs = getAllHistoryLogs();
  const stat = calcPublicStats(logs);
  Logger.log("СТАТИСТИКА ЗВІТУ:");
  Logger.log(stat);
  const report = buildPublicReport();
  Logger.log("Звіт:");
  Logger.log(report);
}

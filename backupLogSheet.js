/**
 * Экспорт листа "Лог змін" как файла CSV или Excel
 * @param {string} format - "csv" или "xlsx"
 * @returns {string} URL созданного файла
 */
function backupLogSheetToDrive(format = "xlsx") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME || "Лог змін");
  if (!logSheet) throw new Error("Лист 'Лог змін' не знайдено!");

  const data = logSheet.getDataRange().getValues();

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const fileName = `backup_log_${timestamp}`;

  let fileUrl = "";

  if (format === "csv") {
    // CSV
    let csv = data.map(row => row.map(cell => `"${(cell + "").replace(/"/g, '""')}`).join(",")).join("\n");
    const blob = Utilities.newBlob(csv, MimeType.CSV, `${fileName}.csv`);
    const file = DriveApp.createFile(blob);
    fileUrl = file.getUrl();
  } else if (format === "xlsx") {
    // Excel
    const tempSS = SpreadsheetApp.create(fileName);
    const tempSheet = tempSS.getSheets()[0];
    tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    const blob = tempSS.getBlob().setContentType(MimeType.MICROSOFT_EXCEL);
    const file = DriveApp.createFile(blob.setName(`${fileName}.xlsx`));
    DriveApp.getFileById(tempSS.getId()).setTrashed(true); // Удаляем временную таблицу

    fileUrl = file.getUrl();
  }

  return fileUrl;
}

/**
 * Архивирует лог в отдельный файл CSV и сохраняет его на Google Диск
 */
function archiveLogHistory() {
  const folderName = "Резервні копії / Логи";
  const folderIter = DriveApp.getFoldersByName(folderName);
  let folder;

  if (folderIter.hasNext()) {
    folder = folderIter.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME || "Лог змін");
  if (!logSheet) throw new Error("Лист 'Лог змін' не знайдено!");

  const data = logSheet.getDataRange().getValues();
  if (data.length <= 1) return; // Только заголовки — не сохраняем

  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const csv = data.map(row =>
    row.map(cell => `"${(cell || "").toString().replace(/"/g, '""')}`).join(",")
  ).join("\n");

  const blob = Utilities.newBlob(csv, MimeType.CSV, `log_archive_${timestamp}.csv`);
  folder.createFile(blob);

  // Очищаем лог после архивации (по желанию)
  logSheet.getRange(2, 1, logSheet.getLastRow()-1, logSheet.getLastColumn()).clearContent();

  Logger.log(`Лог за ${timestamp} архівовано.`);
}
/**
 * Создает триггер на ежедневную архивацию лога
 */
function createDailyArchiveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === "archiveLogHistory");
  if (!exists) {
    ScriptApp.newTrigger("archiveLogHistory")
      .timeBased()
      .atHour(23)
      .everyDays(1)
      .create();
    Logger.log("Тригер на щоденну архівацію створено.");
  }
}

/**
 * Удаляет архивные логи старше N дней
 * @param {number} daysToKeep - Сколько дней хранить
 */
function cleanupOldBackups(daysToKeep = 30) {
  const folder = DriveApp.getFoldersByName("Резервні копії / Логи").next();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getLastUpdated() < cutoffDate) {
      file.setTrashed(true);
    }
  }

  Logger.log("Старі бэкапи видалені.");
}

/**
 * Вызов: через меню — экспортирует лог как Excel
 */
function exportLogSheetAsExcel() {
  try {
    const url = backupLogSheetToDrive("xlsx");
    SpreadsheetApp.getUi().alert("Файл Excel створено", `Завантажити: ${url}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Помилка", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Вызов: через меню — экспортирует лог как CSV
 */
function exportLogSheetAsCSV() {
  try {
    const url = backupLogSheetToDrive("csv");
    SpreadsheetApp.getUi().alert("Файл CSV створено", `Завантажити: ${url}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Помилка", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

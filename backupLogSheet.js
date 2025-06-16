/**
 * ID папки Google Drive для збереження резервних копій
 * Заміни цей рядок на актуальний ID потрібної папки!
 */
const LOG_BACKUP_FOLDER_ID = '1kHedr0VsFe75Uh94_adQobZR8l9FViHS'; 

/**
 * Експорт листа "Лог змін" як файла CSV або Excel у конкретну папку Google Drive
 * @param {string} format - "csv" або "xlsx"
 * @returns {string} URL створеного файлу
 */
function backupLogSheetToDrive(format = "xlsx") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(typeof LOG_SHEET_NAME !== 'undefined' ? LOG_SHEET_NAME : "Лог змін");
  if (!logSheet) throw new Error("Лист 'Лог змін' не знайдено!");

  let folder = null;
  try {
    folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  } catch (e) {
    throw new Error("Папку для резервних копій не знайдено! Перевірте BACKUP_FOLDER_ID.");
  }

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length === 0) throw new Error("Немає даних для експорту!");

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const fileName = `backup_log_${timestamp}`;

  let fileUrl = "";

  if (format === "csv") {
    // CSV з усіма заголовками
    let csv = data.map(row => row.map(cell => `"${(cell !== null && cell !== undefined ? cell.toString().replace(/"/g, '""') : "")}"`).join(",")).join("\n");
    const blob = Utilities.newBlob(csv, MimeType.CSV, `${fileName}.csv`);
    const file = folder.createFile(blob);
    fileUrl = file.getUrl();
  } else if (format === "xlsx") {
    // Створюємо тимчасову таблицю, копіюємо дані, експортуємо як Excel
    const tempSS = SpreadsheetApp.create(fileName);
    const tempSheet = tempSS.getSheets()[0];
    tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    // Видалення зайвих листів, якщо є (Google може додати пусті листи)
    tempSS.getSheets().forEach(s => {
      if (s.getName() !== tempSheet.getName()) tempSS.deleteSheet(s);
    });

    const blob = tempSS.getBlob().setContentType(MimeType.MICROSOFT_EXCEL);
    const file = folder.createFile(blob.setName(`${fileName}.xlsx`));
    DriveApp.getFileById(tempSS.getId()).setTrashed(true); // Видаляємо тимчасову таблицю

    fileUrl = file.getUrl();
  } else {
    throw new Error("Непідтримуваний формат: " + format);
  }

  return fileUrl;
}

/**
 * Архівує лог у файл CSV і зберігає у конкретну папку, очищає лог
 */
function archiveLogHistory() {
  let folder;
  try {
    folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  } catch (e) {
    throw new Error("Папку для резервних копій не знайдено! Перевірте BACKUP_FOLDER_ID.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(typeof LOG_SHEET_NAME !== 'undefined' ? LOG_SHEET_NAME : "Лог змін");
  if (!logSheet) throw new Error("Лист 'Лог змін' не знайдено!");

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return; // Тільки заголовки — не зберігаємо

  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  const csv = data.map(row =>
    row.map(cell => `"${(cell !== null && cell !== undefined ? cell.toString().replace(/"/g, '""') : "")}"`).join(",")
  ).join("\n");

  const blob = Utilities.newBlob(csv, MimeType.CSV, `log_archive_${timestamp}.csv`);
  folder.createFile(blob);

  // Очищаємо лог після архівації (лишаємо тільки заголовки)
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow()-1, logSheet.getLastColumn()).clearContent();
  }

  Logger.log(`Лог за ${timestamp} архівовано.`);
}

/**
 * Створює тригер на щоденну архівацію лога (23:00)
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
 * Видаляє архіви логів старше N днів тільки з бекап-папки
 * @param {number} daysToKeep - Скільки днів зберігати
 */
function cleanupOldBackups(daysToKeep = 30) {
  let folder;
  try {
    folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  } catch (e) {
    throw new Error("Папку для резервних копій не знайдено! Перевірте BACKUP_FOLDER_ID.");
  }

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  const files = folder.getFiles();
  let deleted = 0;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getLastUpdated() < cutoffDate) {
      file.setTrashed(true);
      deleted++;
    }
  }

  Logger.log(`Старі бекапи (${deleted} шт.) видалені.`);
}

/**
 * Виводить діалогове вікно для експорту логу у Excel
 * Перевіряє контекст запуску: якщо не через UI, виводить тільки Logger.log
 */
function exportLogSheetAsExcel() {
  try {
    const url = backupLogSheetToDrive("xlsx");
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert("Файл Excel створено", `Завантажити: ${url}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Файл Excel створено: " + url);
    }
  } catch (e) {
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert("Помилка", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Помилка: " + e.message);
    }
  }
}

/**
 * Виводить діалогове вікно для експорту логу у CSV
 * Перевіряє контекст запуску: якщо не через UI, виводить тільки Logger.log
 */
function exportLogSheetAsCSV() {
  try {
    const url = backupLogSheetToDrive("csv");
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert("Файл CSV створено", `Завантажити: ${url}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Файл CSV створено: " + url);
    }
  } catch (e) {
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert("Помилка", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Помилка: " + e.message);
    }
  }
}

/**
 * Перевірка чи доступний SpreadsheetApp.getUi()
 * (потрібно для уникнення помилки "Cannot call getUi from this context")
 */
function isUiAvailable() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

function getLogFilesList() {
  const folder = DriveApp.getFolderById(LOG_BACKUP_FOLDER_ID);
  const files = folder.getFiles();
  let list = [];
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (!/\.csv$|\.xlsx$/i.test(name)) continue; // тільки csv/xlsx
    list.push({
      id: file.getId(),
      name: name,
      date: file.getLastUpdated()
    });
  }
  // Сортуємо від новіших до старіших
  list.sort((a, b) => b.date - a.date);
  return list;
}

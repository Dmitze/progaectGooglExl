
/**
 * Отримати список архівів логів у TMP_FOLDER_ID (csv/xlsx)
 * Повертає масив {id, name, url, date}
 */
function getLogArchivesList() {
  const folder = DriveApp.getFolderById(TMP_FOLDER_ID);
  const files = folder.getFiles();
  let list = [];
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (!/\.csv$|\.xlsx$/i.test(name)) continue;
    list.push({
      id: file.getId(),
      name: name,
      url: file.getUrl(),
      date: file.getLastUpdated()
    });
  }
  // Сортуємо від новіших до старіших
  list.sort((a, b) => b.date - a.date);
  return list;
}

/**
 * Отримати вміст CSV-файлу як масив рядків для попереднього перегляду
 * @param {string} fileId
 * @param {number} limitRows - максимальна кількість рядків (щоб не навантажувати UI)
 * @return {Array<Array<string>>}
 */
function getCsvPreview(fileId, limitRows = 100) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const text = blob.getDataAsString('UTF-8');
  const lines = text.split(/\r?\n/).slice(0, limitRows);
  return lines.map(l => parseCsvRow(l));
}

/**
 * Дуже простий парсер CSV-рядка (для попереднього перегляду)
 */
function parseCsvRow(line) {
  // Простий варіант: розбиває по ",", прибирає лапки (НЕ універсальний для складних CSV!)
  return line.match(/("([^"]|"")*"|[^,]*)/g)
    .map(cell => cell.replace(/^"|"$/g, '').replace(/""/g, '"'));
}

/**
 * Google Apps Script не дозволяє зробити прямий download, але можна повернути лінк на файл у Drive
 * або створити копію (наприклад, для xlsx можна просто дати лінк на файл)
 * @param {string} fileId
 * @return {string} URL
 */
function getFileDownloadUrl(fileId) {
  const file = DriveApp.getFileById(fileId);
  // Для Drive це завжди лінк на файл, а для прямого download можна використовувати export через API
  return file.getUrl();
}

/**
 * (Опціонально) Зробити копію архіву з новою назвою і повернути посилання
 * @param {string} fileId
 * @returns {string} URL нової копії
 */
function makeCopyOfArchive(fileId) {
  const file = DriveApp.getFileById(fileId);
  const copy = file.makeCopy('COPY_' + file.getName(), DriveApp.getFolderById(TMP_FOLDER_ID));
  return copy.getUrl();
}

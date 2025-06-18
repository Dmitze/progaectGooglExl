
/** ID папки Google Drive для зберігання логів, архівів, кешу, тимчасових файлів */
const TMP_FOLDER_ID = '14EX4nx7NACIv0qCnJL0soVivEmI4SG9G';

/**
 * Отримати посилання на папку для тимчасових/архівних файлів.
 * Використовуйте цю функцію у всіх модулях для доступу до папки.
 */
function getTmpFolderOrThrow() {
  try {
    return DriveApp.getFolderById(TMP_FOLDER_ID);
  } catch (e) {
    throw new Error("Папку для тимчасових файлів не знайдено! Перевірте TMP_FOLDER_ID.");
  }
}

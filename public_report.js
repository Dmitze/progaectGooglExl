/**
 * Публічний/читабельний звіт для неавторизованих
 */

function addPublicReportMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Публічний звіт")
    .addItem("Створити публічний звіт (копію)", "createPublicReport")
    .addToUi();
}

// Копіює дані у новий файл, видаляючи чутливу інформацію
function createPublicReport() {
  // 1. Створити копію файлу/аркуша
  // 2. Видалити/замінити чутливі дані
  // 3. Надати доступ для читання за посиланням
  SpreadsheetApp.getUi().alert("Публічний звіт створено (логіка ще не реалізована).");
}

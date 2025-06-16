/**
 * Динамічний дашборд: створення і оновлення листа з підсумками
 */

function createOrUpdateDashboard() {
  // Створити/оновити лист "Дашборд"
  // 1. Порахувати кількість змін, активність, топ-користувачів
  // 2. Побудувати графік активності (опційно)
  // (Тут буде твоя логіка)
}

// Можна додати кнопку в меню:
function addDashboardMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Дашборд")
    .addItem("Оновити дашборд", "createOrUpdateDashboard")
    .addToUi();
}

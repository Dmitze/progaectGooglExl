/**
 * Пошук по історії змін: користувач, дата, комірка, аркуш
 */

function addHistorySearchMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Пошук змін")
    .addItem("Пошук змін", "showHistorySearchDialog")
    .addToUi();
}

// Відкриває форму для пошуку змін (можна HtmlService)
function showHistorySearchDialog() {
  SpreadsheetApp.getUi().alert("Тут буде форма для пошуку по історії змін.");
}

// Приклад функції для пошуку за критеріями
function findChanges(criteria) {
  // criteria: {user, dateFrom, dateTo, cell, sheet}
  // Повертає масив змін з "Лог змін", що задовольняють умовам
  // (Тут буде твоя логіка)
}

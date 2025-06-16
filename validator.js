/**
 * Контроль орфографії і форматів у ключових полях
 */

function addValidationMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Валідація")
    .addItem("Перевірити орфографію/формати", "runValidation")
    .addToUi();
}

// Основна функція перевірки
function runValidation() {
  // 1. Перевірити дати, email, номери
  // 2. (Опціонально) орфографія, якщо буде API
  SpreadsheetApp.getUi().alert("Тут буде результат перевірки даних.");
}
/**
 * Система коментарів і обговорень до комірки
 */

function addCommentMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Коментарі")
    .addItem("Додати коментар до комірки", "showAddCommentDialog")
    .addItem("Переглянути коментарі", "showCommentsDialog")
    .addToUi();
}

// Діалог для додавання коментаря
function showAddCommentDialog() {
  SpreadsheetApp.getUi().alert("Тут буде форма для додавання коментаря до комірки.");
}

// Діалог для перегляду/обговорення коментарів
function showCommentsDialog() {
  SpreadsheetApp.getUi().alert("Тут буде історія коментарів для виділеної комірки.");
}

// Записати коментар до окремого листа
function addCellComment(sheetName, cellA1, author, text, parentId) {
  // parentId — для підтримки відповідей
  // (Тут буде твоя логіка)
}

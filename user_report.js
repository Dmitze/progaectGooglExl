// === Звіт по діям користувачів ===
function showUsersActionReport() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    SpreadsheetApp.getUi().alert("Лист 'Лог змін' не знайдено.");
    return;
  }
  const data = logSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("Немає записаних змін для звіту.");
    return;
  }

  // Індекси: "Користувач"=1, "Тип дії"=4
  const userStats = {};
  const userActions = {};

  for (let i = 1; i < data.length; i++) {
    const user = data[i][1] || "[невідомо]";
    const type = (data[i][4] || "").trim();
    if (!userStats[user]) userStats[user] = { "Додано значення":0, "Видалено значення":0, "Змінено":0, "Додано рядок":0, "Видалено рядок":0, "Додано стовпець":0, "Видалено стовпець":0, "Зміна значення":0 };
    if (!userActions[user]) userActions[user] = [];
    // Накопичуємо основні типи
    if (userStats[user][type] !== undefined) userStats[user][type]++;
    else if (type) userStats[user][type] = 1;
    // Детальні дії
    const dt = (data[i][0] || "");
    const sheet = (data[i][2] || "");
    const cell = (data[i][3] || "");
    const oldV = (data[i][5] || "");
    const newV = (data[i][6] || "");
    userActions[user].push(`${dt}: [${sheet} ${cell}] ${type} | було: "${oldV}" → стало: "${newV}"`);
  }

  let report = "Звіт по діям користувачів:\n\n";
  for (const user in userStats) {
    report += `Користувач: ${user}\n`;
    report += `  Додано значень:     ${userStats[user]["Додано значення"] || 0}\n`;
    report += `  Видалено значень:  ${userStats[user]["Видалено значення"] || 0}\n`;
    report += `  Змінено значень:   ${userStats[user]["Змінено"] || 0}\n`;
    report += `  Додано рядків:     ${userStats[user]["Додано рядок"] || 0}\n`;
    report += `  Видалено рядків:   ${userStats[user]["Видалено рядок"] || 0}\n`;
    report += `  Додано стовпців:   ${userStats[user]["Додано стовпець"] || 0}\n`;
    report += `  Видалено стовпців: ${userStats[user]["Видалено стовпець"] || 0}\n`;
    report += `  Зміна значення:    ${userStats[user]["Зміна значення"] || 0}\n`;
    report += `  Всього дій:        ${userActions[user].length}\n\n`;
  }
  // Опціонально: деталізований список
  report += "\nДеталізований перелік (для копіювання):\n";
  for (const user in userActions) {
    report += `\n${user}:\n`;
    for (const action of userActions[user]) {
      report += `  ${action}\n`;
    }
  }

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<pre style="font-size:13px;max-height:600px;overflow:auto">${report}</pre>`).setWidth(600).setHeight(700),
    "Звіт по діям користувачів"
  );
}

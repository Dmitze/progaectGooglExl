

/**
 * Модуль для інтеграції з history_search_dialog.html.
 * Повністю готовий для роботи з фронтендом.
 * Підтримка пошуку, фільтрації, аналітики, перевірки ролі адміна.
 */

// === 1. Відкрити діалогове вікно для пошуку змін ===
function showHistorySearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('history_search_dialog')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Пошук по історії змін');
}

// === 2. Отримати всі записи з листа "Лог змін" ===
function getAllHistoryLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Лог змін");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (!data.length) return [];
  const headers = data.shift();
  // Очікується структура: [Дата/час, Користувач, Аркуш, Адреса, Дія, Було, Стало, ...]
  return data.map(row => ({
    date: row[0] ? formatDate(row[0]) : "",
    dateTime: formatDateTime(row[0]),
    user: row[1] || "",
    sheet: row[2] || "",
    address: row[3] || "",
    action: row[4] || "",
    oldValue: row[5] || "",
    newValue: row[6] || ""
  }));
}

// === 3. Чи є користувач адміном? ===
function isUserAdmin() {
  const userEmail = Session.getActiveUser().getEmail();
  // Список адміністраторів — відредагуйте під себе:
  const adminEmails = [
    "admin@example.com",
    "your-admin@email.com"
  ];
  return adminEmails.includes(userEmail);
}

// === 4. Відновити стан аркуша за вибраним записом логу ===
function restoreSheetFromLog(log) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(log.sheet);
    if (!sheet) throw new Error(`Аркуш "${log.sheet}" не знайдено`);
    const range = sheet.getRange(log.address);
    range.setValue(log.oldValue); // oldValue — те, що було ДО зміни
    return "Стан аркуша успішно відновлено";
  } catch (e) {
    throw new Error(`Помилка відновлення: ${e.message}`);
  }
}

// === 5. Зібрати статистику по історії змін ===
function getHistoryAnalytics() {
  const logs = getAllHistoryLogs();
  const users = {};
  const sheets = {};
  const days = {};
  logs.forEach(log => {
    if (log.user) users[log.user] = (users[log.user] || 0) + 1;
    if (log.sheet) sheets[log.sheet] = (sheets[log.sheet] || 0) + 1;
    if (log.date) days[log.date] = (days[log.date] || 0) + 1;
  });
  return { users, sheets, days };
}

// === 6. Форматування дати ===
function formatDate(date) {
  if (!(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function formatDateTime(date) {
  if (!(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

// === 7. Експорт всієї історії у CSV (опціонально, для окремої кнопки) ===
function exportHistoryToCSV() {
  const logs = getAllHistoryLogs();
  if (!logs.length) throw new Error("Немає даних для експорту");
  const headers = ['Дата/час', 'Аркуш', 'Користувач', 'Дія', 'Адреса', 'Було', 'Стало'];
  const rows = [headers].concat(
    logs.map(r => [
      r.dateTime || r.date, r.sheet, r.user, r.action, r.address, r.oldValue, r.newValue
    ])
  );
  const csv = rows.map(row => row.map(cell =>
    `"${(cell||'').toString().replace(/"/g,'""')}"`
  ).join(',')).join('\r\n');
  // Повертаємо CSV-рядок для подальшої роботи на клієнті
  return csv;
}

/**
 * ДОДАТКОВО: Відкрити діалог з аналітикою (опціонально)
 */
function showHistoryAnalytics() {
  const html = HtmlService.createHtmlOutput(`
    <script>
      google.script.run.withSuccessHandler(function(stats){
        let content = '';
        function makeTable(title, data) {
          let table = '<b>'+title+'</b><table border="1" cellspacing="0" cellpadding="5">';
          table += '<tr><th>Назва</th><th>Кількість</th></tr>';
          for (let key in data) {
            table += '<tr><td>'+key+'</td><td>'+data[key]+'</td></tr>';
          }
          table += '</table>';
          return table;
        }
        content += makeTable('Кількість змін по користувачах:', stats.users);
        content += makeTable('Кількість змін по аркушах:', stats.sheets);
        content += makeTable('Кількість змін по днях:', stats.days);
        document.getElementById('analytics').innerHTML = content;
      }).getHistoryAnalytics();
    </script>
    <div style="padding:20px;font-family:sans-serif;">
      <h3>Аналітика активності</h3>
      <div id="analytics"></div>
      <button onclick="google.script.host.close()" style="margin-top:15px;padding:8px 16px;border:none;background:#1976d2;color:white;border-radius:4px;cursor:pointer;">Закрити</button>
    </div>
  `).setWidth(600).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Аналітика активності');
}

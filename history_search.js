/**
 * Пошук по історії змін: користувач, дата, комірка, аркуш
 */

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

/**
 * Открывает диалоговое окно с формой поиска
 */
function showHistorySearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('history_search_dialog')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Пошук по історії змін');
}

/**
 * Получает все записи из листа "Лог змін"
 * @returns {Array} Массив объектов с историей изменений
 */
function getAllHistoryLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Лог змін");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Убираем заголовок

  return data.map(row => ({
    date: row[0] ? formatDate(row[0]) : "",
    dateTime: formatDateTime(row[0]),
    user: row[1],
    sheet: row[2],
    address: row[3],
    action: row[4],
    oldValue: row[5],
    newValue: row[6]
  }));
}

/**
 * Проверяет, является ли пользователь администратором
 * @returns {Boolean}
 */
function isUserAdmin() {
  const userEmail = Session.getActiveUser().getEmail();
  const adminEmails = ["admin@example.com"]; // Замени на своих админов
  return adminEmails.includes(userEmail);
}

/**
 * Восстанавливает состояние листа по выбранной записи лога
 * @param {Object} log - запись из лога
 */
function restoreSheetFromLog(log) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(log.sheet);
    if (!sheet) throw new Error(`Аркуш "${log.sheet}" не знайдено`);

    const range = sheet.getRange(log.address);
    range.setValue(log.newValue); // Или oldValue? зависит от логики

    return "Стан аркуша успішно відновлено";
  } catch (e) {
    throw new Error(`Помилка відновлення: ${e.message}`);
  }
}

/**
 * Собирает статистику по истории изменений
 * @returns {Object} аналитика по пользователям, листам, дням
 */
function getHistoryAnalytics() {
  const logs = getAllHistoryLogs();

  const users = {};
  const sheets = {};
  const days = {};

  logs.forEach(log => {
    users[log.user] = (users[log.user] || 0) + 1;
    sheets[log.sheet] = (sheets[log.sheet] || 0) + 1;

    const day = log.date;
    days[day] = (days[day] || 0) + 1;
  });

  return { users, sheets, days };
}

// === Вспомогательные функции ===

function formatDate(date) {
  if (!(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatDateTime(date) {
  if (!(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

/**
 * Экспортирует ВСЮ историю в CSV
 */
function exportHistoryToCSV() {
  const html = HtmlService.createHtmlOutput(`
    <script>
      google.script.run.withSuccessHandler(function(logs){
        if (!logs || !logs.length) {
          alert("Немає даних для експорту");
          google.script.host.close();
          return;
        }

        const headers = ['Дата/час', 'Аркуш', 'Користувач', 'Дія', 'Адреса', 'Було', 'Стало'];
        const rows = [headers].concat(logs.map(r => [
          r.dateTime || r.date, r.sheet, r.user, r.action, r.address, r.oldValue, r.newValue
        ]));
        const csv = rows.map(row => row.map(cell =>
          \\""${(cell||'').toString().replace(/"/g,'""')}\\""
        ).join(',')).join('\\r\\n');

        const blob = new Blob([csv], {type:'text/csv'});
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'full_history_export.csv';
        document.body.appendChild(a); a.click();
        setTimeout(() => {
          URL.revokeObjectURL(url);
          a.remove();
          google.script.host.close();
        }, 600);
      }).getAllHistoryLogs();
    </script>
    <div style="padding:20px;font-family:sans-serif;">Експорт історії у форматі CSV...</div>
  `).setWidth(350).setHeight(120);

  SpreadsheetApp.getUi().showModalDialog(html, 'Експорт історії');
}

/**
 * Показывает аналитику по всему логу
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

function initDialog() {
  showStatus('Завантаження даних…');
  google.script.run.withSuccessHandler(function(logs){
    console.log('Получены логи:', logs); // Отладочное сообщение
    allLogs = logs || [];
    fillFilters(allLogs);
    resetSearch(); // Показати всі логи
    // Перевірити права адміна
    google.script.run.withSuccessHandler(function(admin){
      isAdmin = !!admin;
      document.getElementById('adminActions').style.display = isAdmin ? '' : 'none';
    }).isUserAdmin && google.script.run.isUserAdmin();
  }).getAllHistoryLogs();
}

function fillFilters(logs) {
  const sheets = new Set(), users = new Set();
  logs.forEach(row => {
    sheets.add(row.sheet);
    users.add(row.user);
  });
  const sheetSel = document.getElementById('searchSheet');
  const userSel = document.getElementById('searchUser');
  sheetSel.innerHTML = '<option value="">Всі</option>';
  Array.from(sheets).sort().forEach(s => {
    if (s) sheetSel.innerHTML += `<option>${escapeHtml(s)}</option>`;
  });
  userSel.innerHTML = '<option value="">Всі</option>';
  Array.from(users).sort().forEach(u => {
    if (u) userSel.innerHTML += `<option>${escapeHtml(u)}</option>`;
  });
}

function getDashboardStats() {
  const logs = getAllHistoryLogs();
  return calculateStats(logs);
}

/**
 * Динамічний дашборд: створення і оновлення листа з підсумками
 */
function createOrUpdateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboardSheet = ss.getSheetByName("Дашборд");

  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet("Дашборд");
  } else {
    dashboardSheet.clear();
  }

  const logs = getAllHistoryLogs();

  if (!logs || !logs.length) {
    dashboardSheet.getRange("A1").setValue("Немає даних для формування дашборду.");
    return;
  }

  const stats = calculateStats(logs);

  // === Заголовок ===
  dashboardSheet.appendRow(["📊 Підсумковий дашборд"]);

  // === Общая статистика ===
  dashboardSheet.appendRow(["Всього змін", stats.total]);
  dashboardSheet.appendRow(["Унікальні користувачі", stats.uniqueUsers]);
  dashboardSheet.appendRow(["Аркушів у журналі", stats.uniqueSheets]);

  // === ТОП користувачів ===
  dashboardSheet.appendRow(["🏆 Топ користувачів"]);
  Object.entries(stats.userStats).forEach(([user, count]) => {
    dashboardSheet.appendRow([user, count]);
  });

  // === ТОП аркушів ===
  dashboardSheet.appendRow(["📁 Найчастіше редаговані аркуші"]);
  Object.entries(stats.sheetStats).forEach(([sheet, count]) => {
    dashboardSheet.appendRow([sheet, count]);
  });

  // === Типи змін ===
  dashboardSheet.appendRow(["⚙️ Типи змін"]);
  Object.entries(stats.actionStats).forEach(([action, count]) => {
    dashboardSheet.appendRow([action, count]);
  });

  // === Активність по днях ===
  dashboardSheet.appendRow(["📅 Активність за останні 7 днів"]);
  dashboardSheet.appendRow(["Дата", "Кількість змін"]);
  stats.dailyStats.forEach(entry => {
    dashboardSheet.appendRow([entry.date, entry.count]);
  });

  dashboardSheet.autoResizeColumns(1, 2);
}

/**
 * Розраховує статистику на основі логів
 * @param {Array} logs — записи из "Лог змін"
 * @returns {Object} объект со статистикой
 */
function calculateStats(logs) {
  const stats = {
    total: logs.length,
    uniqueUsers: 0,
    uniqueSheets: 0,
    userStats: {},
    sheetStats: {},
    actionStats: {},
    dailyStats: {}
  };

  logs.forEach(log => {
    if (log && log.user) {
      stats.userStats[log.user] = (stats.userStats[log.user] || 0) + 1;
    }

    if (log && log.sheet) {
      stats.sheetStats[log.sheet] = (stats.sheetStats[log.sheet] || 0) + 1;
    }

    if (log && log.action) {
      stats.actionStats[log.action] = (stats.actionStats[log.action] || 0) + 1;
    }

    if (log && log.date) {
      const date = new Date(log.date).toISOString().split('T')[0];
      stats.dailyStats[date] = (stats.dailyStats[date] || 0) + 1;
    }
  });

  // Уникальные значения
  stats.uniqueUsers = Object.keys(stats.userStats).length;
  stats.uniqueSheets = Object.keys(stats.sheetStats).length;

  // Сортировка топов
  stats.topUsers = Object.entries(stats.userStats)
    .filter(([key, val]) => key !== 'undefined')
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  stats.topSheets = Object.entries(stats.sheetStats)
    .filter(([key, val]) => key !== 'undefined')
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  stats.topActions = Object.entries(stats.actionStats)
    .filter(([key, val]) => key !== 'undefined')
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  // Активность по дням за последние 7 дней
  const now = new Date();
  const weekAgo = new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000); // 7 дней назад
  const dailyData = [];

  for (let d = new Date(weekAgo); d <= now; d.setDate(d.getDate() + 1)) {
    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    dailyData.push({ date: dateStr, count: stats.dailyStats[dateStr] || 0 });
  }

  return {
    ...stats,
    dailyStats: dailyData
  };
}

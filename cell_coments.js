/**
 * === Система коментарів до комірки ===
 * v3.1 — WOW-логіка, безпечні JSON, автопідсвітка, дерево, згадки, теги, лайки, фільтри, аналітика
 */

const COMMENT_SHEET_NAME = "Коментарі";
const COMMENT_HEADERS = ["ID", "Sheet", "Cell", "Datetime", "Author", "Text", "ParentID", "Resolved", "Likes", "Mentions", "Tags"];

// === Меню ===
function addCommentMenu() {
  SpreadsheetApp.getUi()
    .createMenu("Коментарі")
    .addItem("Додати коментар до комірки", "showAddCommentDialog")
    .addItem("Переглянути коментарі", "showCommentsDialog")
    .addItem("Останні коментарі", "showRecentCommentsDialog")
    .addToUi();
}

// === Встановлюємо активну комірку ===
function getActiveCellLocation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  return {
    sheetName: sheet.getName(),
    cellA1: cell.getA1Notation()
  };
}

// === Створюємо лист "Коментарі", якщо його немає ===
function setupCommentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let commentSheet = ss.getSheetByName(COMMENT_SHEET_NAME);

  if (!commentSheet) {
    commentSheet = ss.insertSheet(COMMENT_SHEET_NAME);
    commentSheet.getRange(1, 1, 1, COMMENT_HEADERS.length).setValues([COMMENT_HEADERS]);
    commentSheet.hideSheet(); // приховуємо лист
  }

  return commentSheet;
}

// === Генеруємо унікальний ID для коментаря ===
function generateCommentId() {
  return Utilities.getUuid();
}

// === Додаємо новий коментар до листа ===
function saveCellComment(sheetName, cellA1, text, parentId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const commentSheet = setupCommentSheet();

    const user = Session.getActiveUser().getEmail() || "unknown";
    const now = new Date();
    const datetime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const id = generateCommentId();
    const mentions = extractMentions(text);
    const tags = extractTags(text);
    const resolved = "no";

    // Зберігаємо коментар
    commentSheet.appendRow([
      id,
      sheetName,
      cellA1,
      datetime,
      user,
      text,
      parentId || "",
      resolved,
      "[]", // Likes (JSON array)
      JSON.stringify(mentions),
      JSON.stringify(tags)
    ]);

    // Автоматичне повідомлення @згаданим користувачам
    notifyMentionedUsers(user, sheetName, cellA1, mentions, text);

    // Підсвічуємо комірку
    highlightCellWithComments(sheetName, cellA1);

    return {
      success: true,
      message: "✅ Коментар додано",
      commentId: id
    };

  } catch (e) {
    Logger.log("Помилка додавання коментаря: " + e.message);
    return {
      success: false,
      message: "❌ Не вдалося додати коментар: " + e.message
    };
  }
}

// === Отримуємо всі коментарі для аркуша/комірки або всі, якщо без параметрів ===
function getCellComments(sheetName, cellA1) {
  const commentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMMENT_SHEET_NAME);
  if (!commentSheet) return [];

  const data = commentSheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const comments = data.slice(1).map(row => ({
    id: row[0],
    sheet: row[1],
    cell: row[2],
    datetime: row[3],
    author: anonymizeUser(row[4]),
    text: row[5],
    parentId: row[6] || null,
    resolved: row[7],
    likes: safeParseJson(row[8]),
    mentions: safeParseJson(row[9]),
    tags: safeParseJson(row[10])
  }));

  if (sheetName && cellA1) {
    return comments.filter(c => c.sheet === sheetName && c.cell === cellA1);
  }
  return comments;
}

// === Додаємо лайк до коментаря ===
function likeComment(commentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(COMMENT_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === commentId) {
      const likes = safeParseJson(data[i][8]);
      const currentUser = Session.getActiveUser().getEmail();
      if (!likes.includes(currentUser)) {
        likes.push(currentUser);
        sheet.getRange(i + 1, 9).setValue(JSON.stringify(likes));
        return { success: true, likes: likes.length };
      }
      return { success: true, likes: likes.length };
    }
  }
  return { success: false, message: "Не знайдено коментаря" };
}

// === Позначити коментар як вирішений ===
function markAsResolved(commentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(COMMENT_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === commentId) {
      sheet.getRange(i + 1, 8).setValue("yes");
      return { success: true, message: "📌 Коментар відмічено як вирішений" };
    }
  }

  return { success: false, message: "❌ Коментар не знайдено" };
}

// === Обробка згадок ===
function extractMentions(text) {
  text = text || "";
  const mentionRegex = /@([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+)/g;
  const matches = text.match(mentionRegex) || [];
  return matches.map(m => m.slice(1));
}

function extractTags(text) {
  text = text || "";
  const tagRegex = /#(\w+)/g;
  const matches = text.match(tagRegex) || [];
  return matches.map(m => m.slice(1));
}

// === Нотифікація згаданим користувачам ===
function notifyMentionedUsers(author, sheetName, cellA1, mentions, text) {
  if (!mentions || mentions.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const link = ss.getUrl() +
    `#gid=${ss.getSheetByName(sheetName).getSheetId()}&range=${cellA1}`;

  mentions.forEach(email => {
    try {
      MailApp.sendEmail({
        to: email,
        subject: `[Згадка] ${author} згадав вас у "${sheetName}"!`,
        body: `${author} згадав вас у комірці ${cellA1}:\n\n"${text}"\n\n👉 Перегляньте обговорення: ${link}`
      });
    } catch (e) {
      Logger.log("Не вдалося надіслати email " + email + ": " + e.message);
    }
  });
}

// === Дерево коментарів (для UI/вкладених відповідей) ===
function buildCommentTree(comments) {
  comments = comments || [];
  const map = {};
  const tree = [];
  comments.forEach(comment => {
    map[comment.id] = { ...comment, replies: [] };
  });
  comments.forEach(comment => {
    if (comment.parentId && map[comment.parentId]) {
      map[comment.parentId].replies.push(map[comment.id]);
    } else if (!comment.parentId) {
      tree.push(map[comment.id]);
    }
  });
  return tree;
}

// === WOW: Фільтрація/аналітика ===
function getTopCommenters() {
  const comments = getCellComments();
  const stats = {};
  comments.forEach(c => {
    stats[c.author] = (stats[c.author] || 0) + 1;
  });
  return Object.entries(stats).sort((a, b) => b[1] - a[1]).slice(0, 5);
}
function getCommentsByTag(tag) {
  return getCellComments().filter(c => c.tags.includes(tag));
}
function getCommentsByUser(user) {
  return getCellComments().filter(c => c.author === user);
}
function getUnresolvedComments() {
  return getCellComments().filter(c => c.resolved !== "yes");
}

// === WOW: Пошук по коментарях (по тексту, тегу, автору, resolved) ===
function searchComments(params) {
  params = params || {};
  const {query, tag, onlyMine, onlyOpen} = params;
  const user = Session.getActiveUser().getEmail();
  return getCellComments().filter(c =>
    (!query || (c.text && c.text.toLowerCase().includes(query.toLowerCase())))
    && (!tag || (c.tags && c.tags.includes(tag)))
    && (!onlyMine || c.author === user)
    && (!onlyOpen || c.resolved !== "yes")
  );
}

// === Підсвічування комірки з коментарем ===
function highlightCellWithComments(sheetName, cellA1) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  try {
    sheet.getRange(cellA1).setBackground("#ffeeba"); // жовтий колор для коментованої комірки
  } catch(e) { Logger.log(e.message); }
}

// === Автоматичне підсвічування для всіх комірок з коментарями ===
function highlightAllCommentedCells() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const comments = getCellComments();
  const sheetMap = {};
  comments.forEach(c => {
    if (!sheetMap[c.sheet]) sheetMap[c.sheet] = new Set();
    sheetMap[c.sheet].add(c.cell);
  });
  Object.keys(sheetMap).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.getDataRange().setBackground(null); // clear
      sheetMap[sheetName].forEach(cellA1 => {
        try { sheet.getRange(cellA1).setBackground("#ffeeba"); } catch(e){}
      });
    }
  });
}

// === Тригер на зміну (автоматичне підсвічування) ===
function onEdit(e) {
  highlightAllCommentedCells();
}

// === Анонімізація автора (опціонально) ===
function anonymizeUser(email) {
  if (!email || email === '') return 'Невідомий';
  return `Користувач #${Math.abs(hashCode(email)) % 100}`;
}
function hashCode(str) {
  str = str || "";
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash = Math.floor(hash);
  }
  return hash;
}

// === Безпечний JSON parse (завжди повертає масив) ===
function safeParseJson(str) {
  if (!str) return [];
  try {
    var val = JSON.parse(str);
    if (Array.isArray(val)) return val;
    return [];
  } catch { return []; }
}

// === Останні коментарі ===
function getRecentComments(limit) {
  return getCellComments().sort((a, b) => new Date(b.datetime) - new Date(a.datetime)).slice(0, limit || 10);
}
function showRecentCommentsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('comments_dialog')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🕒 Останні коментарі');
}

// === Форма для додавання коментаря ===
function showAddCommentDialog() {
  const html = HtmlService.createHtmlOutputFromFile('add_comment_dialog')
    .setWidth(540)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '📝 Додати коментар');
}

// === Форма для перегляду коментарів ===
function showCommentsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('view_comments_dialog')
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '💬 Коментарі до комірки');
}

function addCellCommentFromDialog(text, tags, mentions, isQuestion, parentId, attachments) {
  // твоя логіка, наприклад:
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var sheetName = sheet.getName();
  var cellA1 = cell.getA1Notation();
  var author = Session.getActiveUser().getEmail() || "Unknown";
  var id = addCellComment(sheetName, cellA1, author, text, parentId, tags, mentions, isQuestion ? "question" : "", attachments);
  highlightCellWithComments(sheetName, cellA1);
  notifyMentionedUsers(mentions, sheetName, cellA1, text, id);
  return id;
}

function addCellComment(sheetName, cellA1, author, text, parentId, tags, mentions, type, attachments) {
  const commentSheet = setupCommentSheet();
  const now = new Date();
  const datetime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const id = generateCommentId();
  commentSheet.appendRow([
    id,
    sheetName,
    cellA1,
    datetime,
    author,
    text,
    parentId || "",
    "no",
    "[]", // Likes
    JSON.stringify(mentions || []),
    JSON.stringify(tags || [])
  ]);
  return id;
}

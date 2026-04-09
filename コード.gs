const PAGE_SIZE = 10;

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('サークルイベント一覧')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getEvents(searchQuery = '', page = 1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  if (!sheet) return { events: [], totalPages: 0 };
  
  const values = sheet.getDataRange().getValues();
  values.shift(); // ヘッダー削除

  let allEvents = values.map((row, index) => {
    return {
      rowNumber: index + 2,
      date: row[0] ? Utilities.formatDate(new Date(row[0]), "JST", "yyyy/MM/dd") : "",
      title: row[1] || "",
      location: row[2] || "",
      description: row[3] || "",
      author: row[4] || "",
      status: row[5] || '予定',
      report: row[6] || ''
    };
  }).filter(e => e.title !== ""); // 空行除外

  // 1. 検索フィルタリング
  if (searchQuery && searchQuery.trim() !== "") {
    const q = searchQuery.toLowerCase();
    allEvents = allEvents.filter(e => 
      e.title.toLowerCase().includes(q) || 
      e.location.toLowerCase().includes(q) || 
      e.description.toLowerCase().includes(q) ||
      e.date.includes(q) // 日付でも検索できるように追加
    );
  }

  // 2. 日付順に並び替え（★降順：新しい日付が上）
  allEvents.sort((a, b) => new Date(b.date) - new Date(a.date));

  // 3. ページング処理
  const totalCount = allEvents.length;
  const totalPages = Math.ceil(totalCount / PAGE_SIZE) || 1;
  const start = (page - 1) * PAGE_SIZE;
  const pagedEvents = allEvents.slice(start, start + PAGE_SIZE);

  return {
    events: pagedEvents,
    currentPage: page,
    totalPages: totalPages,
    totalCount: totalCount
  };
}

// 完了報告用
function submitCompletion(rowNumber, comment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  sheet.getRange(rowNumber, 6).setValue('完了');
  sheet.getRange(rowNumber, 7).setValue(comment);
  return true;
}

// 新規投稿用
function createEventAndSendEmail(eventData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  sheet.appendRow([eventData.date, eventData.title, eventData.location, eventData.description, eventData.author, '予定', '']);
  MailApp.sendEmail('kitayama@enaworks.net', `【新着】${eventData.title}`, `新しい投稿がありました。\n日付:${eventData.date}\n内容:${eventData.description}`);
  return true;
}

// 編集用
function updateEvent(rowNumber, updatedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  sheet.getRange(rowNumber, 1, 1, 4).setValues([[updatedData.date, updatedData.title, updatedData.location, updatedData.description]]);
  return true;
}

// 削除用
function deleteEvent(rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  sheet.deleteRow(rowNumber);
  return true;
}

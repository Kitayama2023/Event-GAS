/**
 * サークルイベント管理システム - サーバー側スクリプト
 * * 必要シート:
 * 1. 「イベント一覧」 (A:日付, B:イベント名, C:場所, D:詳細, E:投稿者, F:ステータス, G:完了コメント)
 * 2. 「変更履歴」   (A:更新日時, B:行番号, C:イベント名, D:変更内容)
 */

const PAGE_SIZE = 10; // 1ページあたりの表示件数
const NOTIFICATION_EMAIL = 'kitayama@enaworks.net'; // 通知先アドレス

// 1. ウェブアプリを表示する
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('サークルイベントポータル')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. イベント一覧を取得する（検索・ソート・ページング込）
function getEvents(searchQuery = '', page = 1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  if (!sheet) return { events: [], totalPages: 1, currentPage: 1, totalCount: 0 };
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { events: [], totalPages: 1, currentPage: 1, totalCount: 0 };
  
  // A列からG列までの全データを取得
  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  // オブジェクトの配列に変換（空行は除外）
  let allEvents = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (!row[1]) continue; // タイトル(B列)がなければスキップ

    allEvents.push({
      rowNumber: i + 2,
      date: row[0] instanceof Date ? Utilities.formatDate(row[0], "JST", "yyyy/MM/dd") : String(row[0]),
      title: String(row[1]),
      location: String(row[2]),
      description: String(row[3]),
      author: String(row[4]),
      status: row[5] || '予定',
      report: row[6] || ''
    });
  }

  // 検索フィルタリング
  if (searchQuery && searchQuery.trim() !== "") {
    const q = searchQuery.toLowerCase();
    allEvents = allEvents.filter(e => 
      e.title.toLowerCase().includes(q) || 
      e.location.toLowerCase().includes(q) || 
      e.description.toLowerCase().includes(q) ||
      e.date.includes(q)
    );
  }

  // 日付の降順ソート（最新・未来の日付が上）
  allEvents.sort((a, b) => {
    const dateA = new Date(a.date).getTime();
    const dateB = new Date(b.date).getTime();
    return dateB - dateA;
  });

  // ページング処理
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

// 3. 新規イベントを投稿し、メールを送る
function createEventAndSendEmail(eventData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');

  // スプレッドシートへ追記
  sheet.appendRow([
    eventData.date,
    eventData.title,
    eventData.location,
    eventData.description,
    eventData.author,
    '予定', // ステータス
    ''      // 完了コメント
  ]);

  // メール通知（新規投稿時のみ送信）
  const subject = `【新着イベント】${eventData.title}`;
  const body = `サークルの新しい予定が投稿されました。

【日付】${eventData.date}
【タイトル】${eventData.title}
【場所】${eventData.location}
【投稿者】${eventData.author}

【詳細】
${eventData.description}

※アプリから確認・編集が可能です。`;

  MailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
  return true;
}

// 4. イベント内容を編集（保存）する
function updateEvent(rowNumber, updatedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('イベント一覧');
  const historySheet = ss.getSheetByName('変更履歴');
  
  // 履歴用：変更前データの取得
  const oldData = eventSheet.getRange(rowNumber, 1, 1, 4).getValues()[0];
  const oldTitle = oldData[1];
  
  let changes = [];
  if (oldData[0] != updatedData.date) changes.push(`日付: ${updatedData.date}`);
  if (oldData[1] != updatedData.title) changes.push(`タイトル: ${updatedData.title}`);
  if (oldData[2] != updatedData.location) changes.push(`場所: ${updatedData.location}`);
  if (oldData[3] != updatedData.description) changes.push("詳細を変更");

  // 一覧シートの更新 (A-D列)
  eventSheet.getRange(rowNumber, 1, 1, 4).setValues([[
    updatedData.date,
    updatedData.title,
    updatedData.location,
    updatedData.description
  ]]);
  
  // 変更があれば履歴シートに記録
  if (historySheet && changes.length > 0) {
    historySheet.appendRow([
      new Date(),
      rowNumber,
      oldTitle,
      changes.join(', ')
    ]);
  }
  
  // ※ 編集時はメールを送らない設定
  return true;
}

// 5. イベントを削除する
function deleteEvent(rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('イベント一覧');
  const historySheet = ss.getSheetByName('変更履歴');
  
  const oldTitle = eventSheet.getRange(rowNumber, 2).getValue();
  
  // 履歴に記録
  if (historySheet) {
    historySheet.appendRow([new Date(), rowNumber, oldTitle, "イベントを削除しました"]);
  }
  
  eventSheet.deleteRow(rowNumber);
  
  // ※ 削除時はメールを送らない設定
  return true;
}

// 6. 完了報告を書き込む
function submitCompletion(rowNumber, comment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  
  // F列(6):ステータス, G列(7):完了コメント
  sheet.getRange(rowNumber, 6).setValue('完了');
  sheet.getRange(rowNumber, 7).setValue(comment);
  
  return true;
}

// 7. 定期メールチェック（メール投稿用）
function processEventEmails() {
  const threads = GmailApp.search('label:イベント投稿 is:unread');
  if (threads.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      if (message.isUnread()) {
        const body = message.getPlainBody();
        const sender = message.getFrom();
        
        const dateMatch = body.match(/【日付】(.*)/);
        const titleMatch = body.match(/【タイトル】(.*)/);
        const locationMatch = body.match(/【場所】(.*)/);
        const descriptionMatch = body.match(/【本文】([\s\S]*)/);

        if (dateMatch && titleMatch && locationMatch) {
          const cleanSender = sender.replace(/^.*<([^>]+)>.*$/, '$1');
          sheet.appendRow([
            dateMatch[1].trim(), 
            titleMatch[1].trim(), 
            locationMatch[1].trim(), 
            descriptionMatch ? descriptionMatch[1].trim() : "", 
            cleanSender, 
            '予定', 
            ''
          ]);
        }
        message.markRead();
      }
    });
  });
}

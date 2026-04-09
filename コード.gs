// 1. サイトを表示する関数（変更なし）
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('サークルイベント一覧')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. データを取得する関数（★修正箇所）
// 1ページあたりの表示件数
const PAGE_SIZE = 10;

function getEvents(searchQuery = '', page = 1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  const values = sheet.getDataRange().getValues();
  values.shift(); // ヘッダー削除

  // 1. 全データをオブジェクト化（ここではまだ並べ替えないので .reverse() は削除）
  let allEvents = values.map((row, index) => {
    return {
      rowNumber: index + 2,
      date: Utilities.formatDate(new Date(row[0]), "JST", "yyyy/MM/dd"),
      title: row[1],
      location: row[2],
      description: row[3],
      author: row[4],
      status: row[5] || '予定',
      report: row[6] || ''
    };
  });

  // 2. ★ここで「イベント日付の降順」に並び替える処理を追加★
  allEvents.sort((a, b) => {
    const timeA = new Date(a.date).getTime();
    const timeB = new Date(b.date).getTime();
    return timeB - timeA; // 降順（timeA - timeB にすると昇順になります）
  });

  // 3. 検索フィルタリング
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    allEvents = allEvents.filter(e => 
      e.title.toLowerCase().includes(q) || 
      e.location.toLowerCase().includes(q) || 
      e.description.toLowerCase().includes(q)
    );
  }

  // 4. ページング処理
  const totalCount = allEvents.length;
  const totalPages = Math.ceil(totalCount / PAGE_SIZE);
  const start = (page - 1) * PAGE_SIZE;
  const pagedEvents = allEvents.slice(start, start + PAGE_SIZE);

  return {
    events: pagedEvents,
    currentPage: page,
    totalPages: totalPages,
    totalCount: totalCount
  };
}

// 3. 【新規追加】完了報告をシートに書き込む関数
function submitCompletion(rowNumber, comment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  
  // 指定された行のF列(6)に「完了」、G列(7)にコメントを書き込む
  sheet.getRange(rowNumber, 6).setValue('完了');
  sheet.getRange(rowNumber, 7).setValue(comment);
  
  return true; // 成功したことをフロントに返す
}


function processEventEmails() {
  // 「イベント投稿」ラベルが付いた【未読】メールを検索
  const threads = GmailApp.search('label:MS関係 is:unread');
  
  if (threads.length === 0) {
    return; // 新しい未読メールがなければここで処理を終了
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');

  threads.forEach(thread => {
    const messages = thread.getMessages();
    
    messages.forEach(message => {
      // 未読メッセージのみを処理対象とする
      if (message.isUnread()) {
        const body = message.getPlainBody(); // メールの本文を取得
        const sender = message.getFrom();    // 送信者を取得
        
        // 正規表現を使って、フォーマットから各項目を抽出
        const dateMatch = body.match(/【日付】(.*)/);
        const titleMatch = body.match(/【タイトル】(.*)/);
        const locationMatch = body.match(/【場所】(.*)/);
        const descriptionMatch = body.match(/【本文】([\s\S]*)/); // 【本文】以降すべてを取得

        // 日付・タイトル・場所が正しく入力されている場合のみシートに追加
        if (dateMatch && titleMatch && locationMatch) {
          const date = dateMatch[1].trim();
          const title = titleMatch[1].trim();
          const location = locationMatch[1].trim();
          const description = descriptionMatch ? descriptionMatch[1].trim() : "";
          
          // 送信者のアドレスのみを抽出（例: "名前 <email@example.com>" -> "email@example.com"）
          const cleanSender = sender.replace(/^.*<([^>]+)>.*$/, '$1');

          // スプレッドシートの最終行にデータを追加
          // 順番: 日付, イベント名, 場所, 詳細, 投稿者
          sheet.appendRow([date, title, location, description, cleanSender]);
        }
        
        // 処理が終わったメールを「既読」にする（次回以降の二重登録を防ぐため）
        message.markRead();
      }
    });
  });
}

// ▼▼▼ Code.gs の一番下に追加 ▼▼▼

function createEventAndSendEmail(eventData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');

  // 1. スプレッドシートに新しいイベントを追加
  sheet.appendRow([
    eventData.date,
    eventData.title,
    eventData.location,
    eventData.description,
    eventData.author,
    '予定', // ステータス初期値
    ''      // 完了コメント初期値
  ]);

// 2. メール送信処理（★固定アドレスへ送信するように変更）
  const targetEmail = 'neropi-2022@yahoo.co.jp'; 

  const subject = `【新着イベント】${eventData.title}`;
  const body = `サークルの新しいイベントが投稿されました！

【日付】${eventData.date}
【イベント名】${eventData.title}
【場所】${eventData.location}
【投稿者】${eventData.author}

【詳細】
${eventData.description}

※詳しくはサークルのイベント一覧サイトをご確認ください。
`;

  // 固定アドレス宛にメールを送信
  MailApp.sendEmail(targetEmail, subject, body);

  return true;
}


// ▼▼▼ Code.gs の末尾に追加 ▼▼▼

function updateEvent(rowNumber, updatedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント一覧');
  
  // 指定された行の A列(1)からD列(4)までを新しいデータで上書きする
  // 順序: [日付, タイトル, 場所, 詳細]
  const range = sheet.getRange(rowNumber, 1, 1, 4);
  range.setValues([[
    updatedData.date,
    updatedData.title,
    updatedData.location,
    updatedData.description
  ]]);
  
  return true;
}

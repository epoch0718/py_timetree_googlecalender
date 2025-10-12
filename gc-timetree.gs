// このスクリプトが管理するイベントであることを示すための「目印」
const SYNC_TAG = '[TimeTree]'; // 見えない文字(ゼロ幅スペース)で囲んだタグ

/**
 * 12時間表記(AM/PM)の時刻文字列をパースして、Dateオブジェクトを生成するヘルパー関数
 * @param {string} dateStr - "YYYY-MM-DD"形式の日付文字列
 * @param {string} timeStr - "H:MM AM/PM"形式の時刻文字列
 * @returns {Date}
 */
function parseDateTime(dateStr, timeStr) {
  // この関数は変更なし
  const dateParts = dateStr.split('-');
  const year = parseInt(dateParts[0], 10);
  const month = parseInt(dateParts[1], 10) - 1;
  const day = parseInt(dateParts[2], 10);

  const timeMatch = timeStr.match(/(\d+):(\d+)\s(AM|PM)/);
  if (!timeMatch) return null; // 不正な形式の場合はnullを返す

  let hour = parseInt(timeMatch[1], 10);
  const minute = parseInt(timeMatch[2], 10);
  const ampm = timeMatch[3];

  if (ampm === 'PM' && hour < 12) {
    hour += 12;
  }
  if (ampm === 'AM' && hour === 12) {
    hour = 0;
  }

  return new Date(year, month, day, hour, minute);
}

/**
 * イベントのユニークなキーを生成する関数
 * @param {string} title - イベントのタイトル
 * @param {string} dateStr - "YYYY-MM-DD"形式の日付文字列
 * @returns {string} - "タイトル-YYYY-MM-DD" 形式のキー
 */
function createEventKey(title, dateStr) {
  return `${title}-${dateStr}`;
}

/**
 * WebアプリとしてPOSTリクエストを受け取ったときに実行されるメイン関数
 * @param {Object} e - POSTリクエストのイベントオブジェクト
 */
function doPost(e) {
  const logs = [];
  let statusMessage = "";
  let createdCount = 0;
  let updatedCount = 0;
  let deletedCount = 0;

  let debugExistingEvents = []; 
  
  try {
    const timetreeEvents = JSON.parse(e.postData.contents);
    //const calendar = CalendarApp.getDefaultCalendar();
    const calendar = CalendarApp.getCalendarById('epoch.making.glass@gmail.com'); 
    
    
    logs.push(`Received ${timetreeEvents.length} events to process from TimeTree.`);
    if (timetreeEvents.length === 0) {
      logs.push("No events to process. Sync finished.");
      return ContentService.createTextOutput(JSON.stringify({ status: "No events received.", logs: logs })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- 1. 既存の同期済みイベントの名簿を作成 ---
    const firstEventDate = new Date(timetreeEvents[0].date);
    const year = firstEventDate.getFullYear();
    const month = firstEventDate.getMonth();
    const firstDayOfMonth = new Date(year, month, 1);
    const lastDayOfMonth = new Date(year, month + 1, 0, 23, 59, 59);

    const existingEvents = calendar.getEvents(firstDayOfMonth, lastDayOfMonth)
      .filter(event => {
        try {
          //return event.getDescription().includes(SYNC_TAG);
          return event.getDescription().startsWith(SYNC_TAG);
        } catch (err) {
          return false;
        }
      });

    // デバッグのため、取得した既存イベントの情報を整形して保存
    debugExistingEvents = existingEvents.map(event => {
      const startTime = event.getStartTime();
      const dateStrJST = Utilities.formatDate(startTime, "Asia/Tokyo", "yyyy-MM-dd");
      return {
        title: event.getTitle(),
        startTime_raw: startTime.toString(),
        startTime_jst_formatted: dateStrJST,
        generated_key: createEventKey(event.getTitle(), dateStrJST)
      };
    });
      
    // 高速で検索できるように、既存イベントをMapに変換する
    const googleEventsMap = new Map();
    existingEvents.forEach(event => {
      /*
      const key = createEventKey(event.getTitle(), Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
      */
      // 1. イベントの開始時刻をDateオブジェクトとして取得
      const startTime = event.getStartTime();
      
      // 2. タイムゾーンを指定して、"yyyy-MM-dd"形式の文字列に変換
      //    これならプロジェクト設定に依存せず、常に日本時間で日付が計算される
      const dateStr = Utilities.formatDate(startTime, "Asia/Tokyo", "yyyy-MM-dd");
      
      // 3. 修正された日付文字列でキーを生成
      const key = createEventKey(event.getTitle(), dateStr);

      googleEventsMap.set(key, event);

    });
    logs.push(`Found ${googleEventsMap.size} existing synced events in Google Calendar for ${year}-${month + 1}.`);

    // --- 2. TimeTreeの予定を名簿と照合し、更新または新規作成 ---
    timetreeEvents.forEach(ttEvent => {
      const key = createEventKey(ttEvent.title, ttEvent.date);
      //const options = { description: SYNC_TAG + "\n" + (ttEvent.memo || "") };
      const memo = ttEvent.memo || "";
      // SYNC_TAG とメモを単純に改行で結合する
      const description = `${SYNC_TAG}\n${memo}`;
      const options = { description: description };

      if (googleEventsMap.has(key)) {
        // **【更新処理】** 既存の予定が見つかった場合
        const existingEvent = googleEventsMap.get(key);
        let needsUpdate = false;

        // メモの比較と更新
        if (existingEvent.getDescription() !== options.description) {
          existingEvent.setDescription(options.description);
          needsUpdate = true;
        }

        // 時間の比較と更新
        if (ttEvent.time) { // 時間指定イベントの場合
          const newStartTime = parseDateTime(ttEvent.date, ttEvent.time);
          const newEndTime = new Date(newStartTime.getTime() + (60 * 60 * 1000));
          if (existingEvent.getStartTime().getTime() !== newStartTime.getTime() || existingEvent.getEndTime().getTime() !== newEndTime.getTime()) {
            existingEvent.setTime(newStartTime, newEndTime);
            needsUpdate = true;
          }
        }
        
        if (needsUpdate) {
          logs.push(`Updating event: '${ttEvent.title}'`);
          updatedCount++;
        }
        
        googleEventsMap.delete(key); // 処理済みとして名簿から削除

      } else {
        // **【新規作成処理】** 既存の予定が見つからなかった場合
        logs.push(`Creating new event: '${ttEvent.title}'`);
        if (ttEvent.time) {
          const startTime = parseDateTime(ttEvent.date, ttEvent.time);
          const endTime = new Date(startTime.getTime() + (60 * 60 * 1000));
          calendar.createEvent(ttEvent.title, startTime, endTime, options);
        } else {
          const eventDate = new Date(ttEvent.date);
          const utcDate = new Date(eventDate.getUTCFullYear(), eventDate.getUTCMonth(), eventDate.getUTCDate());
          calendar.createAllDayEvent(ttEvent.title, utcDate, options);
        }
        createdCount++;
      }
    });

    // --- 3. 名簿に残り、不要になった予定を削除 ---
    if (googleEventsMap.size > 0) {
      logs.push(`Deleting ${googleEventsMap.size} events that no longer exist in TimeTree.`);
      googleEventsMap.forEach(eventToDelete => {
        logs.push(` - Deleting: '${eventToDelete.getTitle()}'`);
        eventToDelete.deleteEvent();
        deletedCount++;
      });
    }

    statusMessage = `Sync complete. Created: ${createdCount}, Updated: ${updatedCount}, Deleted: ${deletedCount}.`;
    logs.push(statusMessage);

  } catch (error) {
    statusMessage = "Error processing request: " + error.toString();
    logs.push(statusMessage, error.stack);
    logs.push("Received data: " + (e ? e.postData.contents : "N/A"));
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(
      { 
        status: statusMessage, 
        logs: logs,
        debug_data: {
          google_calendar_events_found: debugExistingEvents
        } 
      }))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * GitHub Actionsのワークフローを起動します。
 */
function triggerGitHubActionsWorkflow() {
  // --- ▼ あなたの情報に合わせて変更してください ▼ ---
  const GITHUB_OWNER = 'epoch0718'; // あなたのGitHubユーザー名
  const GITHUB_REPO = 'py_timetree_googlecalender'; // リポジトリ名
  const WORKFLOW_FILE_NAME = 'main.yml';      // ワークフローのファイル名
  const GIT_BRANCH = 'main';                  // 対象のブランチ名
  // --- ▲ 設定ここまで ▲ ---

  // スクリプトプロパティから安全にPATを取得
  const GITHUB_PAT = PropertiesService.getScriptProperties().getProperty('GITHUB_PAT');
  
  if (!GITHUB_PAT) {
    Logger.log('エラー: スクリプトプロパティに GITHUB_PAT が設定されていません。');
    return;
  }

  const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/actions/workflows/${WORKFLOW_FILE_NAME}/dispatches`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Accept': 'application/vnd.github.v3+json',
      'Authorization': `token ${GITHUB_PAT}`
    },
    payload: JSON.stringify({
      'ref': GIT_BRANCH
    })
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    // レスポンスコード204は成功を意味する
    if (response.getResponseCode() === 204) {
      Logger.log('GitHub Actionsワークフローの起動に成功しました。');
    } else {
      Logger.log(`ワークフローの起動に失敗しました。ステータスコード: ${response.getResponseCode()}, レスポンス: ${response.getContentText()}`);
    }
  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.toString()}`);
  }
}

/**
 * 【手動テスト用】doPost関数をエディタから実行するための関数
 * Pythonから送られてくるのと同じ形式のテストデータを擬似的に作成し、doPostを呼び出します。
 */
function testDoPost() {
  // --- ▼ Pythonから送られてくるデータをここに再現します ▼ ---
  const testEvents = [
    {
      "date": "2025-10-12", // 既存
      "time": "5:00 PM",   // 既存
      "title": "10/12 1700",
      "memo": "メモメモ111"
    },
    {
      "date": "2025-10-14", // 新規
      "time": "7:00 PM",  // 新規
      "title": "手動テストイベント 2",
      "memo": "これも手動テストです。\n改行もOK。"
    }
  ];
  // --- ▲ テストデータここまで ▲ ---

  // doPostに渡すための、擬似的なイベントオブジェクト`e`を作成
  const mockEventObject = {
    postData: {
      contents: JSON.stringify(testEvents)
    }
  };

  Logger.log("★★★ testDoPost を開始します ★★★");
  
  // 擬似的な`e`を使って、doPost関数を呼び出す
  const response = doPost(mockEventObject);
  
  // doPostから返ってきた結果をログに出力
  Logger.log("--- doPostからのレスポンス ---");
  Logger.log(response.getContent());
  
  Logger.log("★★★ testDoPost が完了しました ★★★");
}
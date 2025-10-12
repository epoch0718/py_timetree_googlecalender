/**
 * @OnlyCurrentDoc
 * NotionとGoogleカレンダーを双方向同期するスクリプト
 * - 今日から指定日数先までの予定を同期対象とする
 * - 変更は差分ではなく、期間内の全アイテムを比較して反映
 * - 削除（キャンセル/アーカイブ）も双方向に反映
 * - 1分トリガーでの実行を想定
 * - GCイベント説明へのNotionリンク追加機能は削除
 * - Notion DBクエリのペイロードを修正
 * - トリガー自動設定機能を追加
 */

// --- 設定項目 ---
// ★★★ 以下を環境に合わせて変更してください ★★★
// 【注意】セキュリティのため、APIキーなどはスクリプトプロパティでの管理を推奨します
const NOTION_API_KEY = 'notionのAPIキー'; // 
const NOTION_DATABASE_ID = '287c32e7c65781ddb2aec4ebfdad083d'; // 
const CALENDAR_ID = 'epoch.making.glass@gmail.com'; // 
const SYNC_RANGE_DAYS = 10; // 同期対象とする日数（今日から何日先までか）

// Notionデータベースのプロパティ名（実際のプロパティ名に合わせてください）
const NOTION_PROPS = {
  name: 'タイトル',         // ページタイトル (必須)
  date: '実行日',         // 日付プロパティ (必須)
  gcEventId: 'GC Event ID', // Google CalendarのイベントIDを格納するテキストプロパティ (必須)
  gcLink: 'GC Link',        // Google Calendarへのリンクを格納するURLプロパティ (任意)
  memo: 'メモ'
};
// ★★★ 設定項目ここまで ★★★

// --- グローバル定数 ---
const GC_EXT_PROP_NOTION_PAGE_ID = 'notionPageId'; // GCイベントの拡張プロパティに保存するNotionページIDのキー
const SCRIPT_LOCK = LockService.getScriptLock();
const MAX_LOCK_WAIT_SECONDS = 10; // ロックの最大待機時間（秒）
const MAX_EXECUTION_TIME_SECONDS = 330; // GASの最大実行時間（秒）- 少し余裕を持たせる (6分 = 360秒)
const NOTION_API_BASE_URL = 'https://api.notion.com/v1';
const NOTION_API_VERSION = '2022-06-28';
const TIMESTAMP_COMPARISON_BUFFER_MS = 5000; // タイムスタンプ比較時の許容誤差 (5秒)
const MAX_API_RETRIES = 3; // APIリトライ回数
const RETRY_WAIT_BASE_MS = 1000; // リトライ待機時間の基本値
const TRIGGER_FUNCTION_NAME = 'mainSyncTrigger'; // トリガーで実行する関数名

// --- メイン処理 ---
/**
 * 同期処理のメイン関数。トリガーで呼び出されることを想定。
 */
function mainSyncTrigger() {
  const scriptStartTime = new Date().getTime();
  Logger.log(`同期処理を開始します... 開始時刻: ${new Date(scriptStartTime).toLocaleString()}`);

  if (!SCRIPT_LOCK.tryLock(MAX_LOCK_WAIT_SECONDS * 1000)) {
    Logger.log('他のプロセスが実行中のため、今回の同期処理をスキップします。');
    return;
  }

  let errorOccurred = false;
  try {
    // APIサービスの有効性チェック
    if (typeof Calendar === 'undefined') {
      throw new Error("Google Calendar API 詳細サービスが無効です。「サービス」+から追加してください。");
    }
    if (!NOTION_API_KEY || NOTION_API_KEY === 'YOUR_NOTION_API_KEY') { // YOUR_NOTION_API_KEY は初期値のプレースホルダとしてチェック
       throw new Error("Notion APIキーが設定されていません。コード内の NOTION_API_KEY を設定してください。");
    }
     if (!NOTION_DATABASE_ID || NOTION_DATABASE_ID === 'YOUR_NOTION_DATABASE_ID') { // YOUR_NOTION_DATABASE_ID は初期値のプレースホルダとしてチェック
       throw new Error("NotionデータベースIDが設定されていません。コード内の NOTION_DATABASE_ID を設定してください。");
    }
     if (!CALENDAR_ID || CALENDAR_ID === 'primary' && !Session.getActiveUser().getEmail()) { // primary でメールアドレスが取得できない場合もエラー
         // primary の場合、 Calendar.CalendarList.get('primary') で存在確認する方がより確実だが、API有効化が必要
         const cal = CalendarApp.getCalendarById(CALENDAR_ID); // CalendarAppで存在確認
         if (!cal) {
             throw new Error(`カレンダーID '${CALENDAR_ID}' が見つからないか、アクセス権がありません。コード内の CALENDAR_ID を確認してください。`);
         }
     }


    // 1. Google Calendar -> Notion 同期
    Logger.log("--- Google Calendar -> Notion 同期開始 ---");
    syncGoogleCalendarToNotion(scriptStartTime);
    Logger.log("--- Google Calendar -> Notion 同期終了 ---");

    if (isTimeRunningOut(scriptStartTime)) return; // 時間切れチェック

    // 2. Notion -> Google Calendar 同期
    Logger.log("--- Notion -> Google Calendar 同期開始 ---");
    syncNotionToGoogleCalendar(scriptStartTime);
    Logger.log("--- Notion -> Google Calendar 同期終了 ---");

    if (!isTimeRunningOut(scriptStartTime)) {
      const elapsedTime = (new Date().getTime() - scriptStartTime) / 1000;
      Logger.log(`同期処理が正常に完了しました。経過時間: ${elapsedTime.toFixed(1)}秒`);
    }

  } catch (error) {
    errorOccurred = true;
    Logger.log(`同期処理中にエラーが発生しました: ${error}\n${error.stack || ''}`);
    // 必要に応じてエラー通知（メール送信など）をここに追加
  } finally {
    SCRIPT_LOCK.releaseLock();
    // Logger.log('スクリプトロックを解放しました。'); // ログ削減
    if (errorOccurred) {
      Logger.log("同期処理はエラーにより終了しました。");
    } else if (isTimeRunningOut(scriptStartTime)) {
        Logger.log(`同期処理は実行時間制限 (${MAX_EXECUTION_TIME_SECONDS}秒) により中断されました。`);
    }
  }
}

/**
 * GASの実行時間制限が近づいているかチェック
 */
function isTimeRunningOut(startTime) {
  const elapsedTimeSeconds = (new Date().getTime() - startTime) / 1000;
  if (elapsedTimeSeconds >= MAX_EXECUTION_TIME_SECONDS) {
    Logger.log(`GAS実行時間制限 (${MAX_EXECUTION_TIME_SECONDS}秒) 超過のため中断。経過時間: ${elapsedTimeSeconds.toFixed(1)}秒`);
    return true;
  }
  return false;
}

// --- Google Calendar -> Notion 同期 ---

/**
 * Google Calendarの変更をNotionに同期する
 */
function syncGoogleCalendarToNotion(scriptStartTime) {
  const today = new Date();
  today.setHours(0, 0, 0, 0); // 今日の0時0分0秒

  const syncStartDate = new Date(today);
  const syncEndDate = new Date(today);
  syncEndDate.setDate(syncEndDate.getDate() + SYNC_RANGE_DAYS); // SYNC_RANGE_DAYS日後の0時

  // Logger.log(`GC同期対象期間: ${syncStartDate.toISOString()} (timeMin) から ${syncEndDate.toISOString()} (timeMax) まで`); // ログ削減

  // Google Calendarからイベントを取得
  const events = getAllGcEventsInRange(syncStartDate, syncEndDate, scriptStartTime);
  if (events === null) {
    Logger.log("GCイベントの取得に失敗したため、GC->Notion同期を中断します。");
    return;
  }
  if (events.length === 0) {
      // Logger.log("GC: 同期対象期間内にイベントはありません。"); // ログ削減
      return;
  }

  Logger.log(`GC: 同期対象期間内のイベント ${events.length} 件を処理します。`);

  let processedCount = 0;
  let createdCount = 0;
  let updatedCount = 0;
  let deletedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  for (const event of events) {
    if (isTimeRunningOut(scriptStartTime)) return; // 時間切れチェック

    const eventId = event.id;
    const status = event.status;
    const gcUpdatedTime = event.updated ? new Date(event.updated) : null;
    const notionPageIdFromEvent = event.extendedProperties?.private?.[GC_EXT_PROP_NOTION_PAGE_ID];

    try {
      // タイトル必須チェック
      if (!event.summary || event.summary.trim() === '') {
        Logger.log(`[スキップ] GCイベント(${eventId})のタイトルが空です。`);
        skippedCount++;
        const pageToArchive = notionPageIdFromEvent ? getNotionPageById(notionPageIdFromEvent) : findNotionPageByGcEventId(eventId);
        if (pageToArchive && !pageToArchive.archived) {
            Logger.log(`  -> タイトルが空になったため、Notionページ (${pageToArchive.id}) をアーカイブします。`);
            deleteNotionPage(pageToArchive.id);
        }
        continue;
      }

      // --- 削除 (Cancelled) 処理 ---
      if (status === 'cancelled') {
        // Logger.log(`  GC Event ID: ${eventId} - アクション: GCイベントキャンセル処理`); // ログ削減
        let pageFoundAndDeleted = false;
        if (notionPageIdFromEvent) {
             const pageById = getNotionPageById(notionPageIdFromEvent);
             if(pageById && !pageById.archived) {
                 // Logger.log(`  -> 拡張プロパティで見つかったNotionページ (${notionPageIdFromEvent}) をアーカイブします。`); // ログ削減
                 if (deleteNotionPage(notionPageIdFromEvent)) { deletedCount++; processedCount++; pageFoundAndDeleted = true; }
                 else { errorCount++; }
             } else if (pageById && pageById.archived) { skippedCount++; pageFoundAndDeleted = true; }
        }
        if (!pageFoundAndDeleted) {
            const notionPage = findNotionPageByGcEventId(eventId);
            if (notionPage && !notionPage.archived) {
              // Logger.log(`  -> DB検索で見つかったNotionページ (${notionPage.id}) をアーカイブします。`); // ログ削減
              if (deleteNotionPage(notionPage.id)) { deletedCount++; processedCount++; pageFoundAndDeleted = true; }
              else { errorCount++; }
            } else if (notionPage && notionPage.archived) { skippedCount++; pageFoundAndDeleted = true; }
        }
        if (!pageFoundAndDeleted) { /* Logger.log(`  -> 対応ページなしorアーカイブ済 スキップ。`); */ skippedCount++; } // ログ削減
        continue;
      }

      // --- 作成 / 更新 (Confirmed / Tentative) 処理 ---
      let targetNotionPage = null;
      let foundBy = "";
      if (notionPageIdFromEvent) {
          //targetNotionPage = getNotionPageById(notionPageIdFromEvent);
          if (targetNotionPage) { foundBy = `拡張プロパティ`; }
      }
      if (!targetNotionPage) {
          targetNotionPage = findNotionPageByGcEventId(eventId);
          if (targetNotionPage) { foundBy = `GC Event ID`; }
      }

      if (targetNotionPage) {
        // 更新処理
        if (targetNotionPage.archived) {
            // Logger.log(`  [スキップ] 対応するNotionページ (${targetNotionPage.id}) はアーカイブ済。`); // ログ削減
            skippedCount++;
            continue;
        }
        const notionLastEditedTime = new Date(targetNotionPage.last_edited_time);
        if (gcUpdatedTime && notionLastEditedTime && notionLastEditedTime.getTime() > gcUpdatedTime.getTime() + TIMESTAMP_COMPARISON_BUFFER_MS) {
          // Logger.log(`  [スキップ] Notion(${targetNotionPage.id}) の最終編集がGCより新しいため更新しません。`); // ログ削減
          skippedCount++;
           if (!notionPageIdFromEvent) { addNotionPageIdToGcEvent(eventId, targetNotionPage.id); } // 再連携
        } else {
          if (updateNotionPageFromGcEvent(targetNotionPage.id, event)) {
             updatedCount++; processedCount++;
             if (!notionPageIdFromEvent) { addNotionPageIdToGcEvent(eventId, targetNotionPage.id); } // 再連携
          } else { errorCount++; }
        }
      } else {
        // 新規作成処理
        // Logger.log(`  GC Event ID: ${eventId} - 対応Notionページなし、新規作成。`); // ログ削減
        const newPage = createNotionPageFromGcEvent(event);
        if (newPage?.id) {
          createdCount++; processedCount++;
          addNotionPageIdToGcEvent(eventId, newPage.id);
        } else {
          Logger.log(`  [エラー] Notionページの新規作成失敗 (GC ID: ${eventId})`); errorCount++;
        }
      }
    } catch (e) { errorCount++; Logger.log(`[エラー] GC Event (${eventId}) 処理中: ${e}\n${e.stack || ''}`); }
  }
  Logger.log(`GC -> Notion 同期結果: 処理=${processedCount}, 新規=${createdCount}, 更新=${updatedCount}, 削除=${deletedCount}, スキップ=${skippedCount}, エラー=${errorCount}`);
}

/**
 * 指定期間内のGoogle Calendarイベントを全て取得する（ページング対応）
 */
function getAllGcEventsInRange(startDate, endDate, scriptStartTime) {
  const allEvents = []; let nextPageToken = null;
  const syncOptions = { maxResults: 250, singleEvents: true, orderBy: 'startTime', showDeleted: true, timeMin: startDate.toISOString(), timeMax: endDate.toISOString(), fields: "items(id,status,summary,start,end,updated,htmlLink,description,extendedProperties/private),nextPageToken" };
  let attempt = 0;
  do {
    if (isTimeRunningOut(scriptStartTime)) return null;
    if (nextPageToken) { syncOptions.pageToken = nextPageToken; }
    try {
      const eventList = Calendar.Events.list(CALENDAR_ID, syncOptions);
      if (eventList.items) { allEvents.push(...eventList.items); }
      nextPageToken = eventList.nextPageToken; attempt = 0;
    } catch (e) {
        Logger.log(`GCイベント取得エラー: ${e}`);
        if (e.details && e.details.code === 403 && e.details.message.includes('Rate Limit Exceeded')) {
            attempt++;
            if (attempt <= MAX_API_RETRIES) { const waitTime = RETRY_WAIT_BASE_MS * Math.pow(2, attempt -1); Logger.log(`-> Rate Limit超過。${waitTime / 1000}秒待機してリトライ (${attempt}/${MAX_API_RETRIES})`); Utilities.sleep(waitTime); continue; }
            else { Logger.log(`-> リトライ上限 (${MAX_API_RETRIES}回) 超過。取得中止。`); return null; }
        } else if (e.details) { Logger.log(`-> 詳細: ${JSON.stringify(e.details)}`); }
         Logger.log("-> GCイベント取得中に回復不能エラー発生。"); return null;
    }
  } while (nextPageToken);
  return allEvents;
}

// --- Notion -> Google Calendar 同期 ---

/**
 * Notionの変更をGoogle Calendarに同期する
 */
function syncNotionToGoogleCalendar(scriptStartTime) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const syncStartDate = new Date(today); const syncEndDate = new Date(today); syncEndDate.setDate(syncEndDate.getDate() + SYNC_RANGE_DAYS);

  const pages = getAllNotionPagesInDateRange(syncStartDate, syncEndDate, scriptStartTime);
   if (pages === null) { Logger.log("Notionページの取得失敗。Notion->GC同期中断。"); return; }
   if (pages.length === 0) { /* Logger.log("Notion: 同期対象期間内ページなし。"); */ return; } // ログ削減

  Logger.log(`Notion: 同期対象期間内のページ ${pages.length} 件を処理。`);
  let processedCount = 0, createdCount = 0, updatedCount = 0, deletedCount = 0, skippedCount = 0, errorCount = 0;

  for (const page of pages) {
    if (isTimeRunningOut(scriptStartTime)) return;
    const pageId = page.id; const isArchived = page.archived; const notionLastEditedTime = new Date(page.last_edited_time); const gcEventId = page.properties[NOTION_PROPS.gcEventId]?.rich_text?.[0]?.plain_text?.trim();

    try {
      // アーカイブ処理
      if (isArchived) {
        if (gcEventId) { if (deleteGcEvent(gcEventId)) { deletedCount++; processedCount++; } else { skippedCount++; } }
        else { skippedCount++; }
        continue;
      }

      // アクティブページの処理 (タイトル・日付チェック)
       const notionTitle = page.properties[NOTION_PROPS.name]?.title?.[0]?.plain_text?.trim() || "";
       if (notionTitle === "") { Logger.log(`[スキップ] Notionページ(${pageId})タイトル空。`); skippedCount++; if (gcEventId) { deleteGcEvent(gcEventId); } continue; }
       if (!page.properties[NOTION_PROPS.date]?.date?.start) { Logger.log(`[スキップ] Notionページ(${pageId})日付未設定。`); skippedCount++; if (gcEventId) { deleteGcEvent(gcEventId); } continue; }

      let existingGcEvent = null;
      if (gcEventId) { existingGcEvent = getGcEventById(gcEventId); }

      if (existingGcEvent) {
        // 更新処理
        const gcUpdatedTime = existingGcEvent.updated ? new Date(existingGcEvent.updated) : null;
        if (existingGcEvent.status === 'cancelled') { Logger.log(`-> GC(${gcEventId})キャンセル済、新規扱い。`); existingGcEvent = null; clearGcEventIdFromNotion(pageId); }
        else if (gcUpdatedTime && notionLastEditedTime && gcUpdatedTime.getTime() > notionLastEditedTime.getTime() + TIMESTAMP_COMPARISON_BUFFER_MS) { /* Logger.log(`[スキップ] GC(${gcEventId})がNotionより新。`); */ skippedCount++; } // ログ削減
        else { if (updateGcEventFromNotionPage(gcEventId, page)) { updatedCount++; processedCount++; } else { errorCount++; } }
      }

      // 新規作成処理
      if (!existingGcEvent) {
         const newEvent = createGcEventFromNotionPage(page);
         if (newEvent?.id) { createdCount++; processedCount++; updateNotionWithGcEventId(pageId, newEvent.id); addNotionPageIdToGcEvent(newEvent.id, pageId); }
         else { Logger.log(`[エラー] GCイベント新規作成失敗 (Notion ID: ${pageId})`); errorCount++; }
      }
    } catch (e) { errorCount++; Logger.log(`[エラー] Notion Page (${pageId}) 処理中: ${e}\n${e.stack || ''}`); }
  }
  Logger.log(`Notion -> GC 同期結果: 処理=${processedCount}, 新規=${createdCount}, 更新=${updatedCount}, 削除=${deletedCount}, スキップ=${skippedCount}, エラー=${errorCount}`);
}

/**
 * 指定期間内のNotionページを全て取得する（ページング対応）
 */
function getAllNotionPagesInDateRange(startDate, endDate, scriptStartTime) {
  const allPages = []; let nextCursor = null;
  const startDateStr = startDate.toISOString().split('T')[0]; const endDateStr = endDate.toISOString().split('T')[0];
  const filter = { and: [ { property: NOTION_PROPS.date, date: { on_or_after: startDateStr } }, { property: NOTION_PROPS.date, date: { on_or_before: endDateStr } } ] };
  let attempt = 0;
  do {
    if (isTimeRunningOut(scriptStartTime)) return null;
    const payload = { filter: filter, page_size: 100 }; // database_id は不要
    if (nextCursor) { payload.start_cursor = nextCursor; }
    try {
      const response = callNotionApi(`/databases/${NOTION_DATABASE_ID}/query`, 'post', payload);
      if (response.results) { allPages.push(...response.results); }
      nextCursor = response.next_cursor; attempt = 0;
      if (!response.has_more) { break; }
    } catch (e) {
      Logger.log(`Notionページ取得エラー: ${e}`);
       if (e.message.includes('Rate limit exceeded') || e.message.includes('status 429')) {
            attempt++;
            if (attempt <= MAX_API_RETRIES) { const waitTime = RETRY_WAIT_BASE_MS * Math.pow(2, attempt -1); Logger.log(`-> Rate Limit超過。${waitTime / 1000}秒待機してリトライ (${attempt}/${MAX_API_RETRIES})`); Utilities.sleep(waitTime); continue; }
            else { Logger.log(`-> リトライ上限 (${MAX_API_RETRIES}回) 超過。取得中止。`); return null; }
        } else { Logger.log("-> Notionページ取得中に回復不能エラー発生。"); return null; }
    }
  } while (nextCursor);
  return allPages;
}

// --- Google Calendar API Helper --- (変更なし)
function getGcEventById(eventId) { try { return Calendar.Events.get(CALENDAR_ID, eventId, {fields: "id,status,updated,summary,start,end,description,extendedProperties/private"}); } catch (e) { if (e.message.includes('Not Found')) { return null; } else if (e instanceof ReferenceError) { throw e; } else { Logger.log(`[getGcEventById] GC取得エラー(ID:${eventId}):${e}`); return null; } } }
function createGcEventFromNotionPage(page) { const pageId = page.id; try { const res = buildGcEventResourceFromNotionPage(page); return Calendar.Events.insert(res, CALENDAR_ID); } catch (e) { Logger.log(`[createGcEvent] GC作成失敗(Notion:${pageId}):${e}${e.details?JSON.stringify(e.details):''}`); if (e instanceof ReferenceError) throw e; return null; } }

function updateGcEventFromNotionPage(eventId, page) {
   const pageId = page.id; 
   try { 
    const res = buildGcEventResourceFromNotionPage(page); 
    
    /*
    const patch = {summary: res.summary, start: res.start, end: res.end, extendedProperties: res.extendedProperties}; 
    */

    const patch = { summary: res.summary, start: res.start, end: res.end, description: res.description, extendedProperties: res.extendedProperties };

    return Calendar.Events.patch(patch, CALENDAR_ID, eventId); 
    } catch (e) {
       Logger.log(`[updateGcEvent] GC更新失敗(ID:${eventId},Notion:${pageId}):${e}${e.details?JSON.stringify(e.details):''}`); 
       if (e.message.includes('Not Found')) return null; 
       else if (e instanceof ReferenceError) throw e; return null; } 
}

function deleteGcEvent(eventId) { if (!eventId) return false; try { Calendar.Events.remove(CALENDAR_ID, eventId); return true; } catch (e) { if (e.message.includes('Not Found')) { return false; } else if (e instanceof ReferenceError) { Logger.log("[deleteGcEvent] GC API無効"); throw e; } else { Logger.log(`[deleteGcEvent] GC削除エラー(ID:${eventId}):${e}`); return false; } } }
function addNotionPageIdToGcEvent(eventId, pageId) { if (!eventId || !pageId) return false; try { const event = getGcEventById(eventId); if (!event) { Logger.log(`[addNotionPageIdToGcEvent] GC(${eventId})見つからず追記不可`); return false; } const props = event.extendedProperties?.private || {}; if (props[GC_EXT_PROP_NOTION_PAGE_ID] === pageId) return true; const resource = {extendedProperties:{private:{...props,[GC_EXT_PROP_NOTION_PAGE_ID]:pageId}}}; Calendar.Events.patch(resource, CALENDAR_ID, eventId); return true; } catch (e) { Logger.log(`[addNotionPageIdToGcEvent] GC(${eventId})へのID(${pageId})追記エラー:${e}`); if (e instanceof ReferenceError) throw e; return false; } }

// --- Notion API Helper --- (変更なし)
function callNotionApi(endpoint, method = 'get', payload = null, muteHttpExceptions = true) { const options = { method: method, contentType: 'application/json', headers: { 'Authorization': `Bearer ${NOTION_API_KEY}`, 'Notion-Version': NOTION_API_VERSION }, muteHttpExceptions: muteHttpExceptions }; if (payload && (method === 'post' || method === 'patch')) { options.payload = JSON.stringify(payload); } let response, responseBody, responseCode, attempt = 0; while (attempt <= MAX_API_RETRIES) { try { response = UrlFetchApp.fetch(NOTION_API_BASE_URL + endpoint, options); responseCode = response.getResponseCode(); responseBody = response.getContentText(); if (responseCode >= 200 && responseCode < 300) { try { return JSON.parse(responseBody); } catch (parseError) { throw new Error(`Notion API応答JSONパースエラー:${parseError.message}`); } } else if (responseCode === 429) { attempt++; if (attempt <= MAX_API_RETRIES) { const wait = RETRY_WAIT_BASE_MS * Math.pow(2, attempt -1); Logger.log(`[callNotionApi]Rate limit(429)。${wait/1000}秒待機リトライ(${attempt}/${MAX_API_RETRIES})`); Utilities.sleep(wait); continue; } else { throw new Error(`Notion API Rate limit exceeded after ${MAX_API_RETRIES} retries.`); } } else { let msg = `Notion APIエラー:Status ${responseCode}`; try { const err = JSON.parse(responseBody); msg += ` - ${err.code}:${err.message}`; } catch (e) { msg += `\nBody:${responseBody.substring(0,500)}`; } throw new Error(msg); } } catch (fetchError) { throw new Error(`UrlFetchAppエラー:${fetchError.message}`); } } throw new Error(`Notion API call failed after ${MAX_API_RETRIES} retries.`); }
function getNotionPageById(pageId) { if (!pageId) return null; try { return callNotionApi(`/pages/${pageId}`, 'get'); } catch (e) { if (e.message.includes('status 404')) return null; Logger.log(`[getNotionPageById]ページ取得エラー(ID:${pageId}):${e}`); return null; } }
function findNotionPageByGcEventId(gcEventId, onlyActive = false) { if (!gcEventId) return null; const payload = { filter:{property:NOTION_PROPS.gcEventId,rich_text:{equals:gcEventId}}, page_size:1 }; try { const response = callNotionApi(`/databases/${NOTION_DATABASE_ID}/query`, 'post', payload); if (response.results?.length > 0) { const page = response.results[0]; if (onlyActive && page.archived) return null; return page; } else return null; } catch (e) { Logger.log(`[findNotionPageByGcEventId]DB検索エラー(GC ID:${gcEventId}):${e}`); return null; } }
function createNotionPageFromGcEvent(event) { const eventId=event.id; try { const props = buildNotionPropertiesFromGcEvent(event); const payload = {parent:{database_id:NOTION_DATABASE_ID}, properties:props}; return callNotionApi('/pages', 'post', payload); } catch (e) { Logger.log(`[createNotionPage]Notion作成失敗(GC:${eventId}):${e}`); return null; } }
function updateNotionPageFromGcEvent(pageId, event) { const eventId=event.id; try { const props = buildNotionPropertiesFromGcEvent(event); if (!props || Object.keys(props).length === 0) return true; const payload = {properties:props}; callNotionApi(`/pages/${pageId}`, 'patch', payload); return true; } catch (e) { Logger.log(`[updateNotionPage]Notion更新失敗(ID:${pageId},GC:${eventId}):${e}`); return false; } }
function deleteNotionPage(pageId) { if (!pageId) return false; try { callNotionApi(`/pages/${pageId}`, 'patch', {archived:true}); return true; } catch (e) { Logger.log(`[deleteNotionPage]Notionアーカイブ失敗(ID:${pageId}):${e}`); return false; } }
function updateNotionWithGcEventId(pageId, eventId) { if (!pageId || !eventId) return false; try { const payload = {properties:{[NOTION_PROPS.gcEventId]:{rich_text:[{type:"text", text:{content:eventId}}]}}}; callNotionApi(`/pages/${pageId}`, 'patch', payload); return true; } catch (e) { Logger.log(`[updateNotionWithGcEventId]NotionへのGC ID(${eventId})書込失敗(Page:${pageId}):${e}`); return false; } }
function clearGcEventIdFromNotion(pageId) { if (!pageId) return false; try { const payload = {properties:{[NOTION_PROPS.gcEventId]:{rich_text:[]}}}; callNotionApi(`/pages/${pageId}`, 'patch', payload); return true; } catch (e) { Logger.log(`[clearGcEventIdFromNotion]Notion(${pageId})のGC IDクリア失敗:${e}`); return false; } }

// --- Data Conversion Helpers --- (変更なし)
function buildNotionPropertiesFromGcEvent(event) { 
  const props = {}; 
  const eventId = event.id; 
  const title = event.summary?.trim()||''; 
  if(!title) Logger.log(`[buildNotionProps]警告:GC(${eventId})タイトル空`); props[NOTION_PROPS.name]={title:[{text:{content:title}}]}; 
  const{startDate,endDate,isAllDay}=parseGcDates(event.start,event.end); 
  if(!startDate) throw new Error(`Invalid start date for GC(${eventId})`); 
  const startStr=getNotionDateTimeString(startDate,isAllDay); 
  const dateProp={start:startStr}; 
  if(endDate&&!isAllDay){
    const endStr=getNotionDateTimeString(endDate,false);
    if(endStr)dateProp.end=endStr;
  } else if(endDate&&isAllDay){
    dateProp.end=null;
  } 
  props[NOTION_PROPS.date]={date:dateProp}; 
  props[NOTION_PROPS.gcEventId]={rich_text:[{type:"text",text:{content:eventId}}]}; 
  if(NOTION_PROPS.gcLink&&event.htmlLink){
    props[NOTION_PROPS.gcLink]={url:event.htmlLink};
  }else if(NOTION_PROPS.gcLink){
    props[NOTION_PROPS.gcLink]={url:null};
  } 

 // ★★★ ここからが追加部分 ★★★
  // "memo"プロパティが設定されていれば、GCのdescriptionをそこに設定
  if (NOTION_PROPS.memo) {
    const description = event.description || '';
    // Notionのテキストプロパティは2000文字の制限があるため、超える場合は切り詰める
    props[NOTION_PROPS.memo] = {
      rich_text: [{
        type: "text",
        text: { content: description.substring(0, 2000) }
      }]
    };
  }
  // ★★★ 追加ここまで ★★★
  return props;
}

function buildGcEventResourceFromNotionPage(page){
  const pageId=page.id;const props=page.properties; 
  const summary=props[NOTION_PROPS.name]?.title?.[0]?.plain_text?.trim()||''; 
  if(!summary)throw new Error(`Notion(${pageId})タイトル空`); 
  const dateProp=props[NOTION_PROPS.date]?.date; 
  if(!dateProp?.start)throw new Error(`Notion(${pageId})日付不正`); 
  const{start,end,isAllDay,timeZone}=parseNotionDate(dateProp); 
  
  const resource={
    summary:summary,
    start:{},
    end:{},
    extendedProperties:{private:{[GC_EXT_PROP_NOTION_PAGE_ID]:pageId}}
  }; 

  // ★★★ ここからが追加部分 ★★★
  // "memo"プロパティが設定されていれば、その内容をdescriptionに追加
  if (NOTION_PROPS.memo) {
    const memoProp = props[NOTION_PROPS.memo];
    const description = memoProp?.rich_text?.map(rt => rt.plain_text).join('') || '';
    resource.description = description;
  }
  // ★★★ 追加ここまで ★★★

  if(isAllDay){resource.start.date=start.toISOString().split('T')[0]; 
  const gcEnd=new Date(end?end.getTime():start.getTime()); gcEnd.setDate(gcEnd.getDate()+1); resource.end.date=gcEnd.toISOString().split('T')[0];}else{resource.start.dateTime=start.toISOString(); resource.end.dateTime=end?end.toISOString():new Date(start.getTime()+3600000).toISOString(); 
  const tz=timeZone||Session.getScriptTimeZone(); resource.start.timeZone=tz; resource.end.timeZone=tz;} 
  
  return resource;
}

function parseGcDates(gcStart, gcEnd){let start, end, isAllDay=false; try{if(gcStart.dateTime){start=new Date(gcStart.dateTime);if(gcEnd?.dateTime)end=new Date(gcEnd.dateTime);isAllDay=false;}else if(gcStart.date){start=new Date(gcStart.date+'T00:00:00Z');if(gcEnd?.date){const exclEnd=new Date(gcEnd.date+'T00:00:00Z');end=new Date(exclEnd.getTime()-86400000);}else{end=new Date(start.getTime());}isAllDay=true;} if(start&&end&&!isNaN(start.valueOf())&&!isNaN(end.valueOf())&&end<start)end=new Date(start.getTime()); if(start&&isNaN(start.valueOf()))start=null;if(end&&isNaN(end.valueOf()))end=null;}catch(e){Logger.log(`GC日付パースエラー:${e}. Start:${JSON.stringify(gcStart)},End:${JSON.stringify(gcEnd)}`);start=null;end=null;} return{startDate:start,endDate:end,isAllDay};}
function parseNotionDate(notionDateProp){let start, end, isAllDay=false, timeZone=null; try{if(!notionDateProp||!notionDateProp.start)return{start:null,end:null,isAllDay,timeZone}; const startStr=notionDateProp.start; const endStr=notionDateProp.end; timeZone=notionDateProp.time_zone; if(startStr.includes('T')){start=new Date(startStr); if(endStr&&endStr.includes('T'))end=new Date(endStr); else if(endStr)end=new Date(endStr+'T00:00:00'+(timeZone?'':'Z')); isAllDay=false;}else{start=new Date(startStr+'T00:00:00Z'); if(endStr)end=new Date(endStr+'T00:00:00Z'); else end=new Date(start.getTime()); isAllDay=true; timeZone=null;} if(start&&end&&!isNaN(start.valueOf())&&!isNaN(end.valueOf())&&end<start)end=new Date(start.getTime()); if(start&&isNaN(start.valueOf()))start=null; if(end&&isNaN(end.valueOf()))end=null;}catch(e){Logger.log(`Notion日付パースエラー:${e}. DateProp:${JSON.stringify(notionDateProp)}`);start=null;end=null;} return{start,end,isAllDay,timeZone};}
function getNotionDateTimeString(dateObj, isAllDay){if(!dateObj||!(dateObj instanceof Date)||isNaN(dateObj.valueOf()))return null; try{if(isAllDay){const y=dateObj.getUTCFullYear(); const m=(dateObj.getUTCMonth()+1).toString().padStart(2,'0'); const d=dateObj.getUTCDate().toString().padStart(2,'0'); return `${y}-${m}-${d}`;}else{return dateObj.toISOString();}}catch(e){Logger.log(`Notion日付文字列変換エラー:${e}.Date:${dateObj},isAllDay:${isAllDay}`); return null;}}

// --- トリガー設定関数 ---

/**
 * mainSyncTrigger を1分ごとに実行するトリガーを設定します。
 * 既存の同名関数用トリガーは削除されます。
 * 【重要】この関数を一度手動で実行してトリガーを設定してください。
 */
function setupTrigger() {
  // 既存のトリガーを削除
  deleteTriggers();

  // 新しいトリガーを作成
  try {
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .timeBased()
      .everyMinutes(1)
      .create();
    Logger.log(`トリガーを設定しました: ${TRIGGER_FUNCTION_NAME} を1分ごとに実行します。`);
    SpreadsheetApp.getUi().alert('1分ごとの自動同期トリガーを設定しました。'); // スプレッドシートに紐づいている場合
  } catch (e) {
    Logger.log(`トリガーの設定に失敗しました: ${e}`);
    SpreadsheetApp.getUi().alert(`トリガーの設定に失敗しました: ${e}`); // スプレッドシートに紐づいている場合
  }
}

/**
 * mainSyncTrigger を実行する時間主導型トリガーをすべて削除します。
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME &&
        trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
      Logger.log(`既存のトリガー (ID: ${trigger.getUniqueId()}) を削除しました。`);
    }
  });
  if (deletedCount > 0) {
    Logger.log(`${deletedCount}件の既存トリガーを削除しました。`);
  } else {
     Logger.log(`削除対象の既存トリガーは見つかりませんでした。`);
  }
}


// --- 初期化用関数 --- (変更なし)
/**
 * スクリプトプロパティを削除して、次回の同期を完全同期にする（デバッグ用）
 */
function initializeSyncState() {
  if (!SCRIPT_LOCK.tryLock(MAX_LOCK_WAIT_SECONDS * 1000)) { Logger.log("[initializeSyncState] ロック取得失敗、スキップ"); return; }
  try {
    PropertiesService.getScriptProperties().deleteProperty('NOTION_API_KEY');
    PropertiesService.getScriptProperties().deleteProperty('NOTION_DATABASE_ID');
    PropertiesService.getScriptProperties().deleteProperty('CALENDAR_ID');
    Logger.log('[initializeSyncState] 関連するスクリプトプロパティを削除しました（必要に応じて手動で再設定してください）。');
  } finally { SCRIPT_LOCK.releaseLock(); }
}

// --- 手動実行用のテスト関数 ---

/**
 * 【手動実行用】Google CalendarからNotionへの同期のみを実行します。
 */
function runGcToNotionSync() {
  const startTime = new Date().getTime();
  Logger.log("★★★ 手動実行: Google Calendar -> Notion 同期を開始します ★★★");
  
  if (!SCRIPT_LOCK.tryLock(MAX_LOCK_WAIT_SECONDS * 1000)) {
    Logger.log('他のプロセスが実行中のため、手動実行を中止しました。');
    return;
  }
  
  try {
    syncGoogleCalendarToNotion(startTime);
    Logger.log("★★★ 手動実行: Google Calendar -> Notion 同期が完了しました ★★★");
  } catch (error) {
    Logger.log(`★★★ 手動実行中にエラーが発生しました: ${error}\n${error.stack || ''} ★★★`);
  } finally {
    SCRIPT_LOCK.releaseLock();
  }
}

/**
 * 【手動実行用】NotionからGoogle Calendarへの同期のみを実行します。
 */
function runNotionToGcSync() {
  const startTime = new Date().getTime();
  Logger.log("★★★ 手動実行: Notion -> Google Calendar 同期を開始します ★★★");

  if (!SCRIPT_LOCK.tryLock(MAX_LOCK_WAIT_SECONDS * 1000)) {
    Logger.log('他のプロセスが実行中のため、手動実行を中止しました。');
    return;
  }

  try {
    syncNotionToGoogleCalendar(startTime);
    Logger.log("★★★ 手動実行: Notion -> Google Calendar 同期が完了しました ★★★");
  } catch (error) {
    Logger.log(`★★★ 手動実行中にエラーが発生しました: ${error}\n${error.stack || ''} ★★★`);
  } finally {
    SCRIPT_LOCK.releaseLock();
  }
}
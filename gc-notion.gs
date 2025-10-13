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
const NOTION_API_KEY = 'notionのAPIキーを記入してください'; // 
const NOTION_DATABASE_ID = '287c32e7c65781ddb2aec4ebfdad083d'; // 
const CALENDAR_ID = 'epoch.making.glass@gmail.com'; // 
const SYNC_RANGE_DAYS_BEFORE = 7; // 同期対象とする日数（今日から何日前までか）
const SYNC_RANGE_DAYS = 10; // 同期対象とする日数（今日から何日先までか）
const TRIGGER_EVERY_MINUTES = 10;//setupTrigger で何分単位に実行するか決める

// Notionデータベースのプロパティ名（実際のプロパティ名に合わせてください）
const NOTION_PROPS = {
  name: 'タイトル',         // ページタイトル (必須)
  date: '実行日',         // 日付プロパティ (必須)
  gcEventId: 'GC Event ID', // Google CalendarのイベントIDを格納するテキストプロパティ (必須)
  gcLink: 'GC Link',        // Google Calendarへのリンクを格納するURLプロパティ (任意)
  memo: 'メモ',
  price: '単価',
  endDate: '終了日' 
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
  Logger.log('START mainSyncTrigger()');
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
 * Google Calendarの変更をNotionに同期する（API呼び出し最適化版）
 * ループ内でのAPI呼び出しを避け、最初にNotionのデータを一括取得することでクォータ超過を防ぎます。
 */
function syncGoogleCalendarToNotion(scriptStartTime) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const syncStartDate = new Date(today);
  syncStartDate.setDate(syncStartDate.getDate() - SYNC_RANGE_DAYS_BEFORE);

  const syncEndDate = new Date(today);
  syncEndDate.setDate(syncEndDate.getDate() + SYNC_RANGE_DAYS);

  // 1. Google Calendarから同期対象期間のイベントをすべて取得
  const gcEvents = getAllGcEventsInRange(syncStartDate, syncEndDate, scriptStartTime);
  if (gcEvents === null) {
    Logger.log("GCイベントの取得に失敗したため、GC->Notion同期を中断します。");
    return;
  }
  if (gcEvents.length === 0) {
    Logger.log("GC: 同期対象期間内にイベントはありません。");
    // この場合でもNotion側の孤児データを削除する処理に進むべきだが、今回はシンプルに終了
    return;
  }
  Logger.log(`GC: 同期対象期間内のイベント ${gcEvents.length} 件を処理します。`);

  // 2. Notionから対応する期間のページを「一度に」すべて取得
  const notionPages = getAllNotionPagesInDateRange(syncStartDate, syncEndDate, scriptStartTime);
  if (notionPages === null) {
    Logger.log("Notionページの取得に失敗したため、GC->Notion同期を中断します。");
    return;
  }

  // 3. Notionページを高速検索できるようにMapに変換する（キーは「GC Event ID」）
  const notionPagesMap = new Map();
  notionPages.forEach(page => {
    const gcEventId = page.properties[NOTION_PROPS.gcEventId]?.rich_text?.[0]?.plain_text?.trim();
    if (gcEventId) {
      notionPagesMap.set(gcEventId, page);
    }
  });
  Logger.log(`Notion: 同期対象期間内のページ ${notionPages.length} 件を名簿に登録しました。`);

  let createdCount = 0, updatedCount = 0, deletedCount = 0, skippedCount = 0, errorCount = 0;

  // 4. GCイベントを一つずつループし、Notionの名簿と照合する
  for (const event of gcEvents) {
    if (isTimeRunningOut(scriptStartTime)) return;

    const eventId = event.id;
    const status = event.status;
    
    try {
      const targetNotionPage = notionPagesMap.get(eventId);

      // --- 削除 (Cancelled) 処理 ---
      if (status === 'cancelled') {
        if (targetNotionPage && !targetNotionPage.archived) {
          if (deleteNotionPage(targetNotionPage.id)) {
            deletedCount++;
          } else {
            errorCount++;
          }
        } else {
          // 対応するページがないか、すでにアーカイブ済み
          skippedCount++;
        }
        // 処理済みのページをMapから削除
        if (targetNotionPage) notionPagesMap.delete(eventId);
        continue;
      }
      
      // タイトル必須チェック（タイトルがなければNotion側も削除）
      if (!event.summary || event.summary.trim() === '') {
        Logger.log(`[スキップ] GCイベント(${eventId})のタイトルが空です。`);
        skippedCount++;
        if (targetNotionPage && !targetNotionPage.archived) {
            Logger.log(`  -> タイトルが空になったため、Notionページ (${targetNotionPage.id}) をアーカイブします。`);
            deleteNotionPage(targetNotionPage.id);
        }
        if (targetNotionPage) notionPagesMap.delete(eventId);
        continue;
      }
      
      // --- 作成 / 更新 処理 ---
      if (targetNotionPage) {
        // 【更新処理】対応するNotionページが見つかった場合
        if (targetNotionPage.archived) {
            skippedCount++;
        } else {
            const gcUpdatedTime = event.updated ? new Date(event.updated) : null;
            const notionLastEditedTime = new Date(targetNotionPage.last_edited_time);
            
            // Notion側が新しい場合は更新しない
            if (gcUpdatedTime && notionLastEditedTime && notionLastEditedTime.getTime() > gcUpdatedTime.getTime() + TIMESTAMP_COMPARISON_BUFFER_MS) {
                skippedCount++;
            } else {
                if (updateNotionPageFromGcEvent(targetNotionPage.id, event)) {
                    updatedCount++;
                } else {
                    errorCount++;
                }
            }
        }
        // 処理済みのページをMapから削除
        notionPagesMap.delete(eventId);
      } else {
        // 【新規作成処理】対応するNotionページが見つからなかった場合
        const newPage = createNotionPageFromGcEvent(event);
        if (newPage?.id) {
          createdCount++;
          // GCイベント側に新しいNotion Page IDを書き込む
          addNotionPageIdToGcEvent(eventId, newPage.id);
        } else {
          errorCount++;
        }
      }
    } catch (e) {
      errorCount++;
      Logger.log(`[エラー] GC Event (${eventId}) 処理中に予期せぬエラー: ${e}\n${e.stack || ''}`);
    }
  }
  
  // 5. ループ終了後、Mapにまだ残っているNotionページを処理する
  // これらは、GC側では対応するイベントが（期間外に移動するか削除されて）見つからなかったページ
  if (notionPagesMap.size > 0) {
    Logger.log(`GCに対応するイベントが見つからなかったNotionページが ${notionPagesMap.size} 件あります。アーカイブします。`);
    notionPagesMap.forEach(pageToArchive => {
      if (!pageToArchive.archived) {
        Logger.log(` - アーカイブ: ${pageToArchive.properties[NOTION_PROPS.name]?.title?.[0]?.plain_text || pageToArchive.id}`);
        if(deleteNotionPage(pageToArchive.id)) {
          deletedCount++;
        } else {
          errorCount++;
        }
      }
    });
  }

  Logger.log(`GC -> Notion 同期結果: 新規=${createdCount}, 更新=${updatedCount}, 削除=${deletedCount}, スキップ=${skippedCount}, エラー=${errorCount}`);
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
 * Notionの変更をGoogle Calendarに同期する（API呼び出し最適化版）
 */
function syncNotionToGoogleCalendar(scriptStartTime) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const syncStartDate = new Date(today);
  syncStartDate.setDate(syncStartDate.getDate() - SYNC_RANGE_DAYS_BEFORE); 

  const syncEndDate = new Date(today);
  syncEndDate.setDate(syncEndDate.getDate() + SYNC_RANGE_DAYS);

  // 1. Notionから同期対象期間のページをすべて取得
  const notionPages = getAllNotionPagesInDateRange(syncStartDate, syncEndDate, scriptStartTime);
  if (notionPages === null) {
    Logger.log("Notionページの取得に失敗したため、Notion->GC同期を中断します。");
    return;
  }
   if (notionPages.length === 0) {
    Logger.log("Notion: 同期対象期間内にページはありません。");
    // この場合でもGC側の孤児データを削除する処理に進むべきだが、今回はシンプルに終了
    return;
  }
  Logger.log(`Notion: 同期対象期間内のページ ${notionPages.length} 件を処理します。`);

  // 2. Google Calendarから対応する期間のイベントを「一度に」すべて取得
  const gcEvents = getAllGcEventsInRange(syncStartDate, syncEndDate, scriptStartTime);
  if (gcEvents === null) {
    Logger.log("GCイベントの取得に失敗したため、Notion->GC同期を中断します。");
    return;
  }

  // 3. GCイベントを高速検索できるようにMapに変換する（キーは「Notion Page ID」）
  const gcEventsMap = new Map();
  gcEvents.forEach(event => {
    const notionPageId = event.extendedProperties?.private?.[GC_EXT_PROP_NOTION_PAGE_ID];
    if (notionPageId) {
      gcEventsMap.set(notionPageId, event);
    }
  });
  Logger.log(`GC: 同期対象期間内のイベント ${gcEvents.length} 件を名簿に登録しました。`);
  
  let createdCount = 0, updatedCount = 0, deletedCount = 0, skippedCount = 0, errorCount = 0;

  // 4. Notionのページを一つずつループし、GCの名簿と照合する
  for (const page of notionPages) {
    if (isTimeRunningOut(scriptStartTime)) return;

    const pageId = page.id;
    const isArchived = page.archived;
    
    try {
      const existingGcEvent = gcEventsMap.get(pageId);

      // --- 削除 (Archived) 処理 ---
      if (isArchived) {
        if (existingGcEvent && existingGcEvent.status !== 'cancelled') {
          if (deleteGcEvent(existingGcEvent.id)) {
            deletedCount++;
          } else {
            errorCount++;
          }
        } else {
          skippedCount++;
        }
        if (existingGcEvent) gcEventsMap.delete(pageId);
        continue;
      }
      
      // タイトル・日付必須チェック
      const notionTitle = page.properties[NOTION_PROPS.name]?.title?.[0]?.plain_text?.trim() || "";
      if (notionTitle === "" || !page.properties[NOTION_PROPS.date]?.date?.start) {
        Logger.log(`[スキップ] Notionページ(${pageId})のタイトルまたは日付が空です。`);
        skippedCount++;
        if (existingGcEvent && existingGcEvent.status !== 'cancelled') {
            Logger.log(`  -> 対応するGCイベント (${existingGcEvent.id}) を削除します。`);
            deleteGcEvent(existingGcEvent.id);
        }
        if (existingGcEvent) gcEventsMap.delete(pageId);
        continue;
      }

      // --- 作成 / 更新 処理 ---
      if (existingGcEvent) {
        // 【更新処理】対応するGCイベントが見つかった場合
        if (existingGcEvent.status === 'cancelled') {
            // GC側でキャンセル済みの場合は、Notion側から再作成（復活）させる
            if (createGcEventFromNotionPage(page)) { createdCount++; } else { errorCount++; }
        } else {
            const notionLastEditedTime = new Date(page.last_edited_time);
            const gcUpdatedTime = existingGcEvent.updated ? new Date(existingGcEvent.updated) : null;
            
            // GC側が新しい場合は更新しない
            if (gcUpdatedTime && notionLastEditedTime && gcUpdatedTime.getTime() > notionLastEditedTime.getTime() + TIMESTAMP_COMPARISON_BUFFER_MS) {
                skippedCount++;
            } else {
                if (updateGcEventFromNotionPage(existingGcEvent.id, page)) {
                    updatedCount++;
                } else {
                    errorCount++;
                }
            }
        }
        // 処理済みのイベントをMapから削除
        gcEventsMap.delete(pageId);
      } else {
        // 【新規作成処理】対応するGCイベントが見つからなかった場合
        const newEvent = createGcEventFromNotionPage(page);
        if (newEvent?.id) {
          createdCount++;
          // 作成したGCイベントにNotion Page IDを書き込む
          addNotionPageIdToGcEvent(newEvent.id, pageId);
          // Notionページに新しいGC Event IDを書き込む
          updateNotionWithGcEventId(pageId, newEvent.id);
        } else {
          errorCount++;
        }
      }
    } catch (e) {
      errorCount++;
      Logger.log(`[エラー] Notion Page (${pageId}) 処理中に予期せぬエラー: ${e}\n${e.stack || ''}`);
    }
  }

  // 5. ループ終了後、Mapにまだ残っているGCイベントを処理する
  // これらは、Notion側に対応するページが見つからなかったイベント
  if (gcEventsMap.size > 0) {
    Logger.log(`Notionに対応するページが見つからなかったGCイベントが ${gcEventsMap.size} 件あります。削除します。`);
    gcEventsMap.forEach(eventToDelete => {
        if (eventToDelete.status !== 'cancelled') {
            Logger.log(` - 削除: ${eventToDelete.summary || eventToDelete.id}`);
            if(deleteGcEvent(eventToDelete.id)) {
                deletedCount++;
            } else {
                errorCount++;
            }
        }
    });
  }
  
  Logger.log(`Notion -> GC 同期結果: 新規=${createdCount}, 更新=${updatedCount}, 削除=${deletedCount}, スキップ=${skippedCount}, エラー=${errorCount}`);
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
function updateNotionPageFromGcEvent(pageId, event) {
   const eventId=event.id; 
   try { 
    const props = buildNotionPropertiesFromGcEvent(event); 
    if (!props || Object.keys(props).length === 0) return true; 
    const payload = {properties:props}; callNotionApi(`/pages/${pageId}`, 'patch', payload); return true; 
   } catch (e) {
     Logger.log(`[updateNotionPage]Notion更新失敗(ID:${pageId},GC:${eventId}):${e}`); return false; 
   } 
}

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


  // ★★★ ここからが追加部分 ★★★
  // --- 2. 終了日（endDate）プロパティの処理 ---
  // "endDate"プロパティが設定されていれば、startDateから年月日を抽出して設定
  if (NOTION_PROPS.endDate && startDate) {
    // startDate（Dateオブジェクト）を "YYYY-MM-DD" 形式の文字列に変換
    // タイムゾーンはスクリプトの実行環境に依存するが、通常はこれで問題ない
    const year = startDate.getFullYear();
    const month = (startDate.getMonth() + 1).toString().padStart(2, '0');
    const day = startDate.getDate().toString().padStart(2, '0');
    const yyyymmdd = `${year}-${month}-${day}`;
    
    // Notionの日付プロパティの形式で設定
    props[NOTION_PROPS.endDate] = {
      date: {
        start: yyyymmdd
        // 終了日プロパティに期間は不要なため、endは設定しない
      }
    };
  }
  // ★★★ 追加ここまで ★★★


  props[NOTION_PROPS.gcEventId]={rich_text:[{type:"text",text:{content:eventId}}]}; 
  if(NOTION_PROPS.gcLink&&event.htmlLink){
    props[NOTION_PROPS.gcLink]={url:event.htmlLink};
  }else if(NOTION_PROPS.gcLink){
    props[NOTION_PROPS.gcLink]={url:null};
  } 

  const description = event.description || '';
  // ★★★ ここからが追加部分 ★★★
  // "memo"プロパティが設定されていれば、GCのdescriptionをそこに設定
  if (NOTION_PROPS.memo) {
    // Notionのテキストプロパティは2000文字の制限があるため、超える場合は切り詰める
    props[NOTION_PROPS.memo] = {
      rich_text: [{
        type: "text",
        text: { content: description.substring(0, 2000) }
      }]
    };
  }
  // ★★★ 追加ここまで ★★★

  // "price"プロパティが設定されていれば、descriptionから金額を解析
  if (NOTION_PROPS.price) {
    // 正規表現で "[金額]10000" や "[価格] 12000" のようなパターンを探す
    // 数字の前の空白はあってもなくてもOK (\s*)
    // 数字はカンマ区切りでもOK ([\d,]+)
    const priceMatch = description.match(/\[(?:金額|価格)\]\s*([\d,]+)/);
    
    if (priceMatch && priceMatch[1]) {
      // マッチした場合、カンマを除去して数値に変換
      const priceValue = parseInt(priceMatch[1].replace(/,/g, ''), 10);
      
      // Notionの数値プロパティの形式で設定
      if (!isNaN(priceValue)) {
        props[NOTION_PROPS.price] = {
          number: priceValue
        };
      }
    } else {
      // マッチしなかった場合、またはすでに値が入っている場合はどうするか？
      // ここでは、マッチしなかった場合は何もしない（nullをセットしない）仕様とします。
      // nullをセットすると、既存の値がクリアされてしまうため。
    }
  }

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
      .everyMinutes(TRIGGER_EVERY_MINUTES)
      .create();
    Logger.log(`トリガーを設定しました: ${TRIGGER_FUNCTION_NAME} を${TRIGGER_EVERY_MINUTES}分ごとに実行します。`);
  } catch (e) {
    Logger.log(`トリガーの設定に失敗しました: ${e}`);
  }
}

/**
 * mainSyncTrigger を実行する時間主導型トリガーをすべて削除します。
 */
function deleteTriggers() {
  forceReleaseScriptLock();
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

// ... (既存のコード) ...


// ★★★ この関数を一番下に追加 ★★★

/**
 * 【緊急用】スクリプトロックを強制的に解放します。
 * 通常の実行で「他のプロセスが実行中」エラーが続く場合に、この関数を手動で実行してください。
 */
function forceReleaseScriptLock() {
  Logger.log("スクリプトロックの強制解放を試みます...");
  try {
    LockService.getScriptLock().releaseLock();
    Logger.log("ロックを正常に解放しました。");
  } catch (e) {
    Logger.log(`ロックの解放中にエラーが発生しました（おそらく、ロックは既に解放されています）: ${e}`);
  }
}
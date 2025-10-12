const SYNC_TAG = '\u200B[ttsync]\u200B'; // 見えない文字(ゼロ幅スペース)で囲んだタグ

/**
 * 12時間表記(AM/PM)の時刻文字列をパースして、Dateオブジェクトを生成するヘルパー関数
 * @param {string} dateStr - "YYYY-MM-DD"形式の日付文字列
 * @param {string} timeStr - "H:MM AM/PM"形式の時刻文字列
 * @returns {Date}
 */
function parseDateTime(dateStr, timeStr) {
  const dateParts = dateStr.split('-');
  const year = parseInt(dateParts[0], 10);
  const month = parseInt(dateParts[1], 10) - 1; // 月は0-indexed
  const day = parseInt(dateParts[2], 10);

  const timeMatch = timeStr.match(/(\d+):(\d+)\s(AM|PM)/);
  let hour = parseInt(timeMatch[1], 10);
  const minute = parseInt(timeMatch[2], 10);
  const ampm = timeMatch[3];

  if (ampm === 'PM' && hour < 12) {
    hour += 12;
  }
  if (ampm === 'AM' && hour === 12) { // 深夜12時(12:00 AM)のケース
    hour = 0;
  }

  return new Date(year, month, day, hour, minute);
}

/**
 * WebアプリとしてPOSTリクエストを受け取ったときに実行されるメイン関数
 * @param {Object} e - POSTリクエストのイベントオブジェクト
 */
function doPost(e) {
  const logs = [];
  let statusMessage = "";
  let deletedCount = 0;
  let createdCount = 0;

  try {
    const events = JSON.parse(e.postData.contents);
    const calendar = CalendarApp.getDefaultCalendar();
    
    logs.push(`Received ${events.length} events to process.`);

    // --- 1. 既存の同期済みイベントを削除 ---
    if (events.length > 0) {
      const firstEventDate = new Date(events[0].date);
      const year = firstEventDate.getFullYear();
      const month = firstEventDate.getMonth();

      const firstDayOfMonth = new Date(year, month, 1);
      const lastDayOfMonth = new Date(year, month + 1, 0, 23, 59, 59);

      logs.push(`Cleaning up events for ${year}-${month + 1}.`);

      // 対象月の全てのイベントを取得
      const allEventsInMonth = calendar.getEvents(firstDayOfMonth, lastDayOfMonth);
      
      // その中から、説明欄に目印(SYNC_TAG)が含まれるものをフィルタリング
      const eventsToDelete = allEventsInMonth.filter(event => {
        try {
          return event.getDescription().includes(SYNC_TAG);
        } catch (err) {
          return false;
        }
      });
      
      if (eventsToDelete.length > 0) {
        logs.push(`Found ${eventsToDelete.length} existing synced events to delete.`);
        eventsToDelete.forEach(event => {
          event.deleteEvent();
        });
        deletedCount = eventsToDelete.length;
      } else {
        logs.push("No existing synced events to delete.");
      }
    }

    // --- 2. 新しいイベントを登録 ---
    logs.push("Creating new events from TimeTree data...");
    events.forEach(eventData => {
      const title = eventData.title;
      const dateStr = eventData.date;
      const timeStr = eventData.time;
      const options = { description: SYNC_TAG };

      logs.push(`Processing: '${title}'`);

      if (timeStr) {
        const startTime = parseDateTime(dateStr, timeStr);
        const endTime = new Date(startTime.getTime() + (60 * 60 * 1000));
        calendar.createEvent(title, startTime, endTime, options);
      } else {
        const eventDate = new Date(dateStr);
        const utcDate = new Date(eventDate.getUTCFullYear(), eventDate.getUTCMonth(), eventDate.getUTCDate());
        calendar.createAllDayEvent(title, utcDate, options);
      }
      createdCount++;
    });
    
    statusMessage = `Sync complete. Deleted: ${deletedCount}, Created: ${createdCount}.`;
    logs.push(statusMessage);
    
  } catch (error) {
    statusMessage = "Error processing request: " + error.toString() + " at line " + error.lineNumber;
    logs.push(statusMessage);
    logs.push("Stack: " + error.stack);
    logs.push("Received data: " + e.postData.contents);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
        status: statusMessage,
        logs: logs 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
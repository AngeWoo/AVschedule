// ====================================================================
//  檔案：code.gs (AV放送立願報名行事曆 - 後端 API)
//  版本：0815_CSV_Confirmed (確認使用 CSV URL 讀取資料)
// ====================================================================

const BACKEND_VERSION = "0815_Email_Ready"; // *** 後端版本號 ***

// --- 全域設定 ---
// [重要] SHEET_ID 仍然需要，用於所有寫入、修改、刪除操作
const SHEET_ID = '1gBIlhEKPQHBvslY29veTdMEJeg2eVcaJx_7A-8cTWIM';
const EVENTS_SHEET_NAME = 'Events';
const SIGNUPS_SHEET_NAME = 'Signups';
const LOGS_SHEET_NAME = 'Logs';
const LINE_SEND_LOGS_SHEET_NAME = 'LineSendLogs';

// --- CSV 公開網址設定 (用於唯讀操作) ---
const EVENTS_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=795926947&single=true&output=csv';
const SIGNUPS_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=767193030&single=true&output=csv';
const NEWS_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTRBfLLti6pcfbynd6kjfZT6Gp5trB8PdXvikHoTNMgLsDzDYsGYmexOqEw6ZkrIedyeAd6DE0bHpso/pub?gid=0&single=true&output=csv';
// 註：Logs 工作表為「寫入專用」，記錄操作日誌時無法透過唯讀的 CSV 進行，故以下 URL 不會被使用。
// const LOGS_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=214281427&single=true&output=csv';


// -------------------- API 核心路由函式 --------------------

function doGet(e) {
  let responsePayload;
  try {
    if (!e || !e.parameter || !e.parameter.payload || !e.parameter.callback) {
      return HtmlService.createHtmlOutput("API 後端已部署。");
    }

    const request = JSON.parse(e.parameter.payload);
    const functionName = request.functionName;
    const params = request.params || {};
    let result;

    console.log(`doGet: 接收到函式呼叫 - ${functionName}，參數: ${JSON.stringify(params)}`);

    switch (functionName) {
      case 'getEventsAndSignups': result = getEventsAndSignups(); break;
      case 'getNewsAnnouncements': result = getNewsAnnouncements(); break;
      case 'getUnifiedSignups': result = getUnifiedSignups(params.searchText, params.startDateStr, params.endDateStr); break;
      case 'getStatsData': result = getStatsData(params.startDateStr, params.endDateStr); break;
      // --- 以下為寫入操作，維持使用 SpreadsheetApp ---
      case 'batchAddSignups': result = batchAddSignups(params.userName, params.entries); break;
      case 'addSignup': result = addSignup(params.eventId, params.userName, params.position); break;
      case 'addBackupSignup': result = addBackupSignup(params.eventId, params.userName, params.position); break;
      case 'removeSignup': result = removeSignup(params.eventId, params.userName); break;
      case 'modifySignup': result = modifySignup(params.signupId, params.newPosition); break;
      case 'createTempSheetAndExport': result = createTempSheetAndExport(params.startDateStr, params.endDateStr); break;
      // --- 分享功能依賴讀取函式，會自動使用新方法 ---
      case 'getSignupsAsTextForToday': result = getSignupsAsTextForToday(); break;
      case 'getSignupsAsTextForTomorrow': result = getSignupsAsTextForTomorrow(); break;
      case 'getSignupsAsTextForDateRange': result = getSignupsAsText(params.startDateStr, params.endDateStr); break;
      case 'sendSignupsToEmail': result = sendSignupsToEmail(params.startDateStr, params.endDateStr, params.targetEmail, params.customBody); break;
      // --- LINE Messaging API ---
      case 'setupDailyTrigger': result = setupDaily20Trigger_MessageAPI(params.lin_to); break;
      case 'manualTriggerDailyNotify': result = dailyNotifyTomorrow_MessageAPI(params.lin_to); break;
      case 'debugLineAuth': result = debugLineAuth(params.lin_to); break;
      case 'setLineToken': result = setLineToken(params.line_channel_access_token); break;
      case 'getLineTargets': result = getLineTargets_(); break;
      case 'clearLineTo': result = clearLineTo(); break;
      case 'removeLineTarget': result = removeLineTarget(params.targetId); break;
      case 'updateLineTargetAlias': result = updateLineTargetAlias(params.targetId, params.alias); break;
      case 'getLineWebhookStatus': result = getLineWebhookStatus_(); break;
      case 'getLineSendLogs': result = getLineSendLogs_(params.limit); break;
      case 'deleteLineSendLog': result = deleteLineSendLog_(params.sheet_row); break;
      case 'getCalendarDownloadLinks': result = getCalendarDownloadLinks_(); break;
      case 'setCalendarDownloadLinks': result = setCalendarDownloadLinks_(params.current_month_url, params.next_month_url); break;
      case 'sendMarqueeAnnouncementsToLine': result = sendMarqueeAnnouncementsToLine_MessageAPI(params.lin_to); break;
      case 'sendCustomLineMessage': result = sendCustomLineMessage_MessageAPI(params.message, params.lin_to); break;
      default: throw new Error(`未知的函式名稱: ${functionName}`);
    }

    responsePayload = {
      status: 'success',
      data: result,
      version: BACKEND_VERSION
    };
  } catch (err) {
    console.error(`doGet 執行失敗: ${err.message}\n${err.stack}`);
    responsePayload = { status: 'error', message: err.message, version: BACKEND_VERSION };
  }

  const callbackFunctionName = e.parameter.callback;
  const jsonpResponse = `${callbackFunctionName}(${JSON.stringify(responsePayload)})`;

  return ContentService.createTextOutput(jsonpResponse).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  try {
    props.setProperty(PROP_LINE_WEBHOOK_LAST_AT, new Date().toISOString());
    const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '';
    if (!rawBody) {
      return ContentService.createTextOutput(JSON.stringify({ ok: true, ignored: true, reason: 'empty_body' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    let body = null;
    try { body = JSON.parse(rawBody); } catch (parseErr) { body = null; }
    if (body && Array.isArray(body.events)) {
      const result = lineWebhook_(body);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, ignored: true, reason: 'unsupported_post_payload' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    props.setProperty(PROP_LINE_WEBHOOK_LAST_ERROR, String(err && err.message ? err.message : err));
    console.error(`doPost webhook 失敗: ${err.message}\n${err.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message ? err.message : err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------- 輔助函式 --------------------
function padTime(timeValue) {
  if (timeValue instanceof Date) {
    try {
      const scriptTimeZone = Session.getScriptTimeZone();
      return Utilities.formatDate(timeValue, scriptTimeZone, 'HH:mm');
    } catch (e) {
      return '00:00';
    }
  }
  let trimmedTime = String(timeValue || '').trim();
  if (/^\d:\d\d$/.test(trimmedTime)) { return '0' + trimmedTime; }
  return trimmedTime;
}

function logOperation(action, eventId, userName, position, result, details) {
  try {
    // 日誌記錄為寫入操作，必須使用 SpreadsheetApp
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    if (!logSheet) {
      console.error(`Log sheet '${LOGS_SHEET_NAME}' not found. Cannot log operation.`);
      return;
    }
    const timestamp = new Date();
    const scriptTimeZone = Session.getScriptTimeZone();
    const formattedTimestamp = Utilities.formatDate(timestamp, scriptTimeZone, 'yyyy/MM/dd HH:mm:ss');

    logSheet.appendRow([formattedTimestamp, action, eventId, userName, position, result, details]);
    SpreadsheetApp.flush();
    console.log(`Logged: ${action} for Event ${eventId}, User ${userName}, Position ${position}, Result: ${result}, Details: ${details}`);
  } catch (logError) {
    console.error(`Failed to log operation to sheet: ${logError.message}\n${logError.stack}`);
  }
}

/**
 * 從指定的 URL 獲取並解析 CSV 內容，並使用快取。
 * @param {string} url CSV 檔案的 URL。
 * @param {string} cacheKey 用於快取的鍵。
 * @returns {Array<Array<string>>} 解析後的 CSV 資料。
 */
function fetchAndParseCsvWithCache(url, cacheKey) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    console.log(`fetchAndParseCsvWithCache: '${cacheKey}' 從快取中讀取。`);
    return JSON.parse(cachedData);
  }
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      const csvContent = response.getContentText('UTF-8');
      const data = Utilities.parseCsv(csvContent);
      // 快取 90 秒，避免請求頻繁，同時確保資料有一定即時性
      cache.put(cacheKey, JSON.stringify(data), 90);
      console.log(`fetchAndParseCsvWithCache: '${cacheKey}' 已擷取並快取。`);
      return data;
    } else {
      throw new Error(`無法擷取 CSV 資料。狀態碼: ${responseCode}`);
    }
  } catch (e) {
    console.error(`fetchAndParseCsvWithCache: 擷取 ${url} 失敗: ${e.message}`);
    throw new Error(`從 ${url} 讀取資料時發生錯誤。`);
  }
}

function invalidateMasterDataCache_() {
  const cache = CacheService.getScriptCache();
  ['events_data', 'signups_data'].forEach(cacheKey => cache.remove(cacheKey));
  console.log('invalidateMasterDataCache_: 已清除 events_data / signups_data 快取。');
}


// -------------------- 資料獲取與處理函式 (已修改為 CSV 模式) --------------------

/**
 * 讀取 "News" 的 CSV 網址並回傳有效的公告訊息。
 */
function getNewsAnnouncements() {
  try {
    const data = fetchAndParseCsvWithCache(NEWS_CSV_URL, 'news_data');
    if (!data || data.length <= 1) {
      console.warn(`公告 CSV 中沒有資料或只有標頭。`);
      return [];
    }

    data.shift(); // 移除標頭

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const scriptTimeZone = Session.getScriptTimeZone();

    const messages = data.map(row => {
      const dateValue = row[0];
      const messageTitle = row[1];
      const messageContent = row[2];

      if (dateValue && messageTitle && messageContent) {
        try {
          const messageDate = new Date(dateValue);
          if (isNaN(messageDate.getTime())) return null; // 略過無效日期

          messageDate.setHours(0, 0, 0, 0);

          // 只顯示今天及未來的公告
          if (messageDate >= today) {
            const dateStr = Utilities.formatDate(messageDate, scriptTimeZone, 'yyyy/MM/dd');
            return `${dateStr}：${messageTitle} - ${messageContent}`;
          }
        } catch (e) {
          console.error(`處理公告日期時發生錯誤: ${dateValue}. Error: ${e.message}`);
          return null;
        }
      }
      return null;
    }).filter(Boolean); // 過濾掉所有 null 的項目

    console.log(`getNewsAnnouncements: 找到 ${messages.length} 則有效公告。`);
    return messages;
  } catch (err) {
    console.error(`getNewsAnnouncements 失敗: ${err.message}`, err.stack);
    // 發生錯誤時回傳空陣列，避免前端頁面崩潰
    return [];
  }
}

/**
 * 從 Google Sheet 直接讀取 Events 和 Signups 的主要資料。
 * 查詢 / 刪除後重查需要即時反映，不能依賴發布 CSV 的刷新延遲。
 */
function getMasterData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const eventsSheet = ss.getSheetByName(EVENTS_SHEET_NAME);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);

    if (!eventsSheet || !signupsSheet) {
      throw new Error(`找不到必要工作表：${EVENTS_SHEET_NAME} / ${SIGNUPS_SHEET_NAME}`);
    }

    const eventsData = eventsSheet.getDataRange().getValues();
    const signupsData = signupsSheet.getDataRange().getValues();

    console.log(`getMasterData (Sheet): 成功讀取資料。`);
    console.log(`getMasterData: Events Sheet 讀取到 ${eventsData.length} 行 (含標頭)。`);
    console.log(`getMasterData: Signups Sheet 讀取到 ${signupsData.length} 行 (含標頭)。`);

    const eventsDataWithoutHeader = eventsData.slice(1);
    const signupsDataWithoutHeader = signupsData.slice(1);
    const scriptTimeZone = Session.getScriptTimeZone();
    const eventsMap = new Map();

    eventsDataWithoutHeader.forEach(row => {
      const eventId = row[0];
      if (eventId && row[2]) {
        try {
          const dateObj = new Date(row[2]);
          if (isNaN(dateObj.getTime())) {
            console.warn(`getMasterData: 無效日期格式，跳過活動 ID ${eventId}: ${row[2]}`);
            return;
          }
          eventsMap.set(eventId, {
            title: row[1] || '未知事件',
            dateString: Utilities.formatDate(dateObj, scriptTimeZone, 'yyyy/MM/dd'),
            dateObj: dateObj,
            startTime: padTime(row[3]),
            endTime: padTime(row[4]),
            maxAttendees: parseInt(row[5], 10) || 999,
            description: row[7] || ''
          });
        } catch (e) {
          console.error(`getMasterData: 處理活動 ID ${eventId} 的日期時發生錯誤: ${row[2]}. 錯誤: ${e.message}`);
        }
      }
    });

    return { eventsData: eventsDataWithoutHeader, signupsData: signupsDataWithoutHeader, eventsMap };
  } catch (err) {
    console.error(`getMasterData (Sheet) 失敗: ${err.message}`, err.stack);
    throw new Error(`從 Google Sheet 讀取資料時發生錯誤: ${err.message}`);
  }
}

function getEventsAndSignups() {
  try {
    const { eventsData, signupsData } = getMasterData();
    const scriptTimeZone = Session.getScriptTimeZone();
    const signupsByEventId = new Map();
    signupsData.forEach(row => {
      const eventId = row[1]; // Column B: EventID
      if (!eventId) return;
      if (!signupsByEventId.has(eventId)) { signupsByEventId.set(eventId, []); }
      signupsByEventId.get(eventId).push({ user: row[2], position: row[4] }); // Column C: User, Column E: Position
    });

    return eventsData.map(row => {
      const [eventId, title, rawDate, rawStartTime, rawEndTime, maxAttendees, positions, description] = row;
      if (!eventId || !title || !rawDate || !rawStartTime) {
        return null;
      }

      const paddedStartTime = padTime(rawStartTime);
      const paddedEndTime = padTime(rawEndTime);

      let datePart;
      try {
        const dateObj = new Date(rawDate);
        if (isNaN(dateObj.getTime())) {
          return null;
        }
        datePart = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy-MM-dd");
      } catch (e) {
        return null;
      }

      const signups = signupsByEventId.get(eventId) || [];
      const signupCount = signups.length;
      const parsedMaxAttendees = parseInt(maxAttendees, 10) || 999;
      const isFull = signupCount >= parsedMaxAttendees;

      let eventColor, textColor;
      if (isFull) { eventColor = '#e74c3c'; textColor = 'white'; }
      else if (signupCount > 0) { eventColor = '#0d6efd'; textColor = 'white'; }
      else { eventColor = '#adb5bd'; textColor = '#212529'; }

      return {
        id: eventId, title: title, start: `${datePart}T${paddedStartTime}`,
        end: paddedEndTime ? `${datePart}T${paddedEndTime}` : null,
        backgroundColor: eventColor, borderColor: eventColor, textColor: textColor,
        extendedProps: {
          full_title: title, description: description, maxAttendees: parsedMaxAttendees,
          signups, positions: (positions || '').split(',').map(p => p.trim()).filter(p => p),
          startTime: paddedStartTime, endTime: paddedEndTime
        }
      };
    }).filter(e => e !== null);
  } catch (err) { console.error("getEventsAndSignups 失敗:", err.message, err.stack); throw err; }
}

function getUnifiedSignups(searchText, startDateStr, endDateStr) {
  try {
    const { signupsData, eventsMap } = getMasterData();
    const searchLower = searchText ? searchText.toLowerCase() : '';
    const dayMap = ['日', '一', '二', '三', '四', '五', '六'];

    let start = null;
    let end = null;
    try {
      if (startDateStr) {
        start = new Date(startDateStr);
        start.setHours(0, 0, 0, 0);
      }
      if (endDateStr) {
        end = new Date(endDateStr);
        end.setHours(23, 59, 59, 999);
      }
    } catch (e) {
      throw new Error('日期格式不正確。');
    }

    let filteredSignups = signupsData.filter((row) => {
      if (row.length < 5) {
        return false;
      }

      const eventId = row[1]; // Column B: EventID
      const eventDetail = eventsMap.get(eventId);

      if (!eventDetail) {
        return false;
      }
      if (!eventDetail.dateObj || isNaN(eventDetail.dateObj.getTime())) {
        return false;
      }

      const eventDate = eventDetail.dateObj;

      if (start && eventDate < start) { return false; }
      if (end && eventDate > end) { return false; }

      if (searchLower) {
        const user = String(row[2] || '').toLowerCase(); // Column C: User
        const position = String(row[4] || '').toLowerCase(); // Column E: Position
        const eventTitle = String(eventDetail.title || '').toLowerCase();
        const eventDescription = String(eventDetail.description || '').toLowerCase();

        if (!user.includes(searchLower) &&
          !position.includes(searchLower) &&
          !eventTitle.includes(searchLower) &&
          !eventDescription.includes(searchLower)) {
          return false;
        }
      }
      return true;
    });

    const mappedSignups = filteredSignups.map(row => {
      const eventId = row[1];
      const eventDetail = eventsMap.get(eventId) || { title: '未知事件', dateString: '日期無效' };
      let eventDayOfWeek = eventDetail.dateObj ? dayMap[eventDetail.dateObj.getDay()] : '';

      let formattedTimestamp = '';
      if (row[3] instanceof Date) { // 如果是 Date 物件
        try {
          formattedTimestamp = Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
        } catch (e) {
          formattedTimestamp = 'Invalid Date';
        }
      } else { // 如果是從 CSV 來的字串
        formattedTimestamp = row[3] || '';
      }

      return {
        signupId: row[0], eventId: eventId, eventTitle: eventDetail.title,
        eventDate: eventDetail.dateString, eventDayOfWeek: eventDayOfWeek,
        startTime: eventDetail.startTime || '',
        endTime: eventDetail.endTime || '',
        user: row[2], timestamp: formattedTimestamp, position: row[4]
      };
    });

    mappedSignups.sort((a, b) => {
      const dateA = new Date(a.eventDate);
      const dateB = new Date(b.eventDate);
      if (isNaN(dateA.getTime()) || isNaN(dateB.getTime())) {
        return 0;
      }
      const dateDiff = dateA.getTime() - dateB.getTime();
      if (dateDiff !== 0) return dateDiff;

      // 日期相同時，依照開始時間排序
      const timeA = a.startTime || '';
      const timeB = b.startTime || '';
      return timeA.localeCompare(timeB);
    });

    return mappedSignups;
  } catch (err) {
    console.error("getUnifiedSignups 失敗:", err.message, err.stack);
    throw err;
  }
}

function getStatsData(startDateStr, endDateStr) {
  try {
    const allSignups = getUnifiedSignups('', startDateStr, endDateStr);
    if (!allSignups || allSignups.length === 0) { return { labels: [], data: [], fullDetails: [] }; }
    const { eventsMap } = getMasterData();
    const statsByEventId = {};
    allSignups.forEach(signup => {
      const eventId = signup.eventId;
      if (!statsByEventId[eventId]) {
        const eventInfoFromMap = eventsMap.get(eventId) || {};
        const eventInfo = {
          title: eventInfoFromMap.title || signup.eventTitle, date: eventInfoFromMap.dateString || signup.eventDate,
          dayOfWeek: signup.eventDayOfWeek, startTime: eventInfoFromMap.startTime || '',
          endTime: eventInfoFromMap.endTime || '', maxAttendees: eventInfoFromMap.maxAttendees || 999
        };
        statsByEventId[eventId] = { count: 0, signups: [], eventInfo: eventInfo, label: `${eventInfo.title} (${eventInfo.date})` };
      }
      statsByEventId[eventId].count++;
      statsByEventId[eventId].signups.push({ user: signup.user, position: signup.position });
    });
    const processedData = Object.values(statsByEventId);
    processedData.sort((a, b) => new Date(a.eventInfo.date).getTime() - new Date(b.eventInfo.date).getTime());
    const labels = processedData.map(item => item.label);
    const data = processedData.map(item => item.count);
    const fullDetails = processedData.map(item => {
      item.signups.sort((a, b) => a.position.localeCompare(b.position));
      return { eventInfo: item.eventInfo, signups: item.signups };
    });
    return { labels, data, fullDetails };
  } catch (err) { console.error("getStatsData 失敗:", err.message, err.stack); throw err; }
}

function getSignupsAsTextForToday() {
  const scriptTimeZone = Session.getScriptTimeZone();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, scriptTimeZone, 'yyyy-MM-dd');
  const result = getSignupsAsText(todayStr, todayStr);
  if (result && result.status === 'success' && result.text) {
    const dateDisplay = Utilities.formatDate(today, scriptTimeZone, 'MM/dd');
    result.text = result.text.replace(`日期：${todayStr} ~ ${todayStr}`, `日期：今天 (${dateDisplay})`);
  }
  return result;
}

function getSignupsAsTextForTomorrow() {
  const scriptTimeZone = Session.getScriptTimeZone();
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, scriptTimeZone, 'yyyy-MM-dd');
  const result = getSignupsAsText(tomorrowStr, tomorrowStr);
  if (result && result.status === 'success' && result.text) {
    const dateDisplay = Utilities.formatDate(tomorrow, scriptTimeZone, 'MM/dd');
    result.text = result.text.replace(`日期：${tomorrowStr} ~ ${tomorrowStr}`, `日期：明天 (${dateDisplay})`);
  }
  return result;
}

function getSignupsAsText(startDateStr, endDateStr) {
  const allSignups = getUnifiedSignups('', startDateStr, endDateStr);
  if (!allSignups || allSignups.length === 0) {
    return { status: 'nodata' };
  }
  const eventsGroup = {};
  allSignups.forEach(signup => {
    const key = `${signup.eventDate} ${signup.eventTitle}`;
    if (!eventsGroup[key]) {
      eventsGroup[key] = [];
    }
    eventsGroup[key].push(`- ${signup.position}: ${signup.user}`);
  });
  let formattedText = `【AV放送立願報名記錄】\n日期：${startDateStr} ~ ${endDateStr}\n----------\n\n`;
  Object.keys(eventsGroup).sort().forEach(eventKey => {
    formattedText += `【${eventKey}】\n`;
    formattedText += eventsGroup[eventKey].join('\n');
    formattedText += '\n\n';
  });
  formattedText += '查看最新報名資訊：\nhttps://angewoo.github.io/AVschedule/';
  return { status: 'success', text: formattedText };
}

function sendSignupsToEmail(startDateStr, endDateStr, targetEmail, customBody) {
  if (!targetEmail) return { status: 'error', message: '未提供 Email 地址。' };

  let body = customBody;
  if (!body) {
    const result = getSignupsAsText(startDateStr, endDateStr);
    if (result.status !== 'success') return result;
    body = result.text;
  }

  try {
    // 注意: 使用 Google 內建 MailApp 發送，寄件者為執行腳本的 Google 帳號。
    // GAS 不支援直接 SMTP 連線，若需使用外部 SMTP (如 smtp2go)，建議改用其 HTTP API。
    MailApp.sendEmail({
      to: targetEmail,
      subject: `AV放送立願報名記錄 (${startDateStr} ~ ${endDateStr})`,
      body: body
    });
    return { status: 'success', message: 'Email 已發送。' };
  } catch (e) {
    console.error('sendSignupsToEmail error: ' + e.toString());
    return { status: 'error', message: '發送 Email 失敗: ' + e.toString() };
  }
}

// -------------------- 主要資料寫入函式 (維持不變，繼續使用 SpreadsheetApp) --------------------

function getEventEndDateTimeInfo_(eventData) {
  const eventTitle = String(eventData && eventData[1] || '');
  const eventDatePart = new Date(eventData && eventData[2]);
  if (isNaN(eventDatePart.getTime())) {
    return { error: `活動 (${eventTitle || '未知活動'}) 的日期格式無效，請管理者檢查。` };
  }
  const endTimeStr = padTime(eventData && eventData[4] || '23:59');
  if (!/^\d\d:\d\d$/.test(endTimeStr)) {
    return { error: `活動 (${eventTitle || '未知活動'}) 的結束時間格式錯誤，請管理者檢查。應為 HH:mm 格式。` };
  }
  const timeParts = endTimeStr.split(':').map(Number);
  return {
    endTimeStr: endTimeStr,
    eventEndDateTime: new Date(
      eventDatePart.getFullYear(),
      eventDatePart.getMonth(),
      eventDatePart.getDate(),
      timeParts[0],
      timeParts[1],
      0
    )
  };
}

function batchAddSignups(userName, entries) {
  const normalizedUserName = String(userName || '').trim();
  const normalizedEntries = Array.isArray(entries)
    ? entries.map((entry, index) => ({
      request_index: index,
      eventId: String(entry && entry.eventId || '').trim(),
      position: String(entry && entry.position || '').trim()
    }))
    : [];

  if (!normalizedUserName) {
    return { status: 'error', message: '缺少姓名。', results: [] };
  }
  if (!normalizedEntries.length) {
    return { status: 'error', message: '未提供任何批次立願項目。', results: [] };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    normalizedEntries.forEach(entry => {
      logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', '系統忙碌中，鎖定失敗');
    });
    return { status: 'error', message: '系統忙碌中，請稍後再試。', results: [] };
  }

  const results = [];
  const pendingRows = [];
  const pendingMetas = [];
  let writeCompleted = false;

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const eventsSheet = ss.getSheetByName(EVENTS_SHEET_NAME);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    const eventRows = eventsSheet.getDataRange().getValues();
    const signupsData = signupsSheet.getDataRange().getValues();
    const eventsById = new Map();
    const signupsByEventId = new Map();
    const requestedEventIds = new Set();

    eventRows.forEach(row => {
      const eventId = String(row && row[0] || '').trim();
      if (eventId) {
        eventsById.set(eventId, row);
      }
    });
    signupsData.forEach(row => {
      const eventId = String(row && row[1] || '').trim();
      if (!eventId) return;
      if (!signupsByEventId.has(eventId)) {
        signupsByEventId.set(eventId, []);
      }
      signupsByEventId.get(eventId).push(row);
    });

    normalizedEntries.forEach((entry, index) => {
      const resultItem = {
        request_index: index,
        event_id: entry.eventId,
        event_title: '',
        position: entry.position,
        status: 'error',
        message: ''
      };

      if (!entry.eventId || !entry.position) {
        resultItem.message = '缺少活動或崗位資訊。';
        results.push(resultItem);
        logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', resultItem.message);
        return;
      }

      const eventData = eventsById.get(entry.eventId);
      if (!eventData) {
        resultItem.message = '找不到此活動。';
        results.push(resultItem);
        logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', resultItem.message);
        return;
      }

      resultItem.event_title = String(eventData[1] || '');

      if (requestedEventIds.has(entry.eventId)) {
        resultItem.message = '同一法會僅能批次立願一個崗位。';
        results.push(resultItem);
        logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', resultItem.message);
        return;
      }
      requestedEventIds.add(entry.eventId);

      const eventEndInfo = getEventEndDateTimeInfo_(eventData);
      if (eventEndInfo.error) {
        resultItem.message = eventEndInfo.error;
        results.push(resultItem);
        logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', resultItem.message);
        return;
      }

      if (new Date() > eventEndInfo.eventEndDateTime) {
        resultItem.message = '此活動已結束，無法報名。';
        results.push(resultItem);
        logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', resultItem.message);
        return;
      }

      const currentSignups = signupsByEventId.get(entry.eventId) || [];
      if (currentSignups.some(row => String(row && row[2] || '').trim() === normalizedUserName)) {
        resultItem.message = '您已經報名過此活動了。';
        results.push(resultItem);
        logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', resultItem.message);
        return;
      }

      const existingPositionHolder = currentSignups.find(row => String(row && row[4] || '').trim() === entry.position);
      if (existingPositionHolder) {
        resultItem.message = `此崗位目前由 [${existingPositionHolder[2]}] 報名。`;
        results.push(resultItem);
        logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', resultItem.message);
        return;
      }

      const maxAttendees = parseInt(eventData[5], 10) || 999;
      if (currentSignups.length >= maxAttendees) {
        resultItem.message = '活動總人數已額滿。';
        results.push(resultItem);
        logOperation('batchAddSignup', entry.eventId, normalizedUserName, entry.position, '失敗', resultItem.message);
        return;
      }

      const newSignupId = 'su' + new Date().getTime() + '_' + index;
      const newRowData = [newSignupId, entry.eventId, normalizedUserName, new Date(), entry.position];
      pendingRows.push(newRowData);
      pendingMetas.push({
        eventId: entry.eventId,
        eventTitle: resultItem.event_title,
        position: entry.position,
        signupId: newSignupId
      });
      currentSignups.push(newRowData);
      signupsByEventId.set(entry.eventId, currentSignups);
      results.push(Object.assign({}, resultItem, {
        status: 'success',
        message: '報名成功！'
      }));
    });

    if (pendingRows.length > 0) {
      const startRow = signupsSheet.getLastRow() + 1;
      signupsSheet.getRange(startRow, 1, pendingRows.length, 5).setValues(pendingRows);
      SpreadsheetApp.flush();

      const writtenRows = signupsSheet.getRange(startRow, 1, pendingRows.length, 5).getValues();
      const isVerified = writtenRows.length === pendingRows.length && writtenRows.every((row, index) =>
        String(row && row[0] || '') === String(pendingRows[index][0]) &&
        String(row && row[1] || '') === String(pendingRows[index][1]) &&
        String(row && row[2] || '') === String(pendingRows[index][2]) &&
        String(row && row[4] || '') === String(pendingRows[index][4])
      );

      if (!isVerified) {
        for (let i = 0; i < pendingRows.length; i++) {
          signupsSheet.deleteRow(startRow);
        }
        SpreadsheetApp.flush();
        pendingMetas.forEach(meta => {
          logOperation('batchAddSignup', meta.eventId, normalizedUserName, meta.position, '失敗', '批次立願資料寫入驗證失敗，已回滾');
        });
        throw new Error('批次立願資料寫入驗證失敗，請稍後再試。');
      }

      writeCompleted = true;
      invalidateMasterDataCache_();
      pendingMetas.forEach(meta => {
        logOperation('batchAddSignup', meta.eventId, normalizedUserName, meta.position, '成功', `批次立願成功 (${meta.eventTitle || meta.eventId})`);
      });
    }

    const successCount = results.filter(item => item.status === 'success').length;
    const failedCount = results.length - successCount;
    return {
      status: 'success',
      requested_count: normalizedEntries.length,
      success_count: successCount,
      failed_count: failedCount,
      results: results
    };
  } catch (err) {
    console.error('batchAddSignups 失敗:', err.message, err.stack);
    if (!writeCompleted) {
      pendingMetas.forEach(meta => {
        logOperation('batchAddSignup', meta.eventId, normalizedUserName, meta.position, '失敗', `後端錯誤: ${err.message}`);
      });
    }
    throw new Error('批次立願時發生後端錯誤，請稍後再試。');
  } finally {
    lock.releaseLock();
  }
}

function addSignup(eventId, userName, position) {
  if (!eventId || !userName || !position) {
    logOperation('addSignup', eventId, userName, position, '失敗', '缺少必要資訊');
    return { status: 'error', message: '缺少必要資訊。' };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    logOperation('addSignup', eventId, userName, position, '失敗', '系統忙碌中，鎖定失敗');
    return { status: 'error', message: '系統忙碌中，請稍後再試。' };
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const eventsSheet = ss.getSheetByName(EVENTS_SHEET_NAME);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);

    const eventData = eventsSheet.getDataRange().getValues().find(row => row[0] === eventId);
    if (!eventData) {
      logOperation('addSignup', eventId, userName, position, '失敗', '找不到此活動');
      return { status: 'error', message: '找不到此活動。' };
    }

    const eventDatePart = new Date(eventData[2]);
    if (isNaN(eventDatePart.getTime())) {
      logOperation('addSignup', eventId, userName, position, '失敗', `活動日期格式無效: ${eventData[2]}`);
      return { status: 'error', message: `活動 (${eventData[1]}) 的日期格式無效，請管理者檢查。` };
    }
    const endTimeStr = padTime(eventData[4] || '23:59');
    if (!/^\d\d:\d\d$/.test(endTimeStr)) {
      logOperation('addSignup', eventId, userName, position, '失敗', `活動結束時間格式錯誤: ${endTimeStr}`);
      return { status: 'error', message: `活動 (${eventData[1]}) 的結束時間格式錯誤，請管理者檢查。應為 HH:mm 格式。` };
    }
    const [hours, minutes] = endTimeStr.split(':').map(Number);
    const eventEndDateTime = new Date(eventDatePart.getFullYear(), eventDatePart.getMonth(), eventDatePart.getDate(), hours, minutes, 0);
    if (new Date() > eventEndDateTime) {
      logOperation('addSignup', eventId, userName, position, '失敗', '活動已結束，無法報名');
      return { status: 'error', message: '此活動已結束，無法報名。' };
    }

    const currentSignups = signupsSheet.getDataRange().getValues().filter(row => row[1] === eventId);

    if (currentSignups.some(row => row[2] === userName)) {
      logOperation('addSignup', eventId, userName, position, '重複報名', '您已經報名過此活動了');
      return { status: 'error', message: '您已經報名過此活動了。' };
    }

    const existingPositionHolder = currentSignups.find(row => row[4] === position);
    if (existingPositionHolder) {
      logOperation('addSignup', eventId, userName, position, '崗位重複', `此崗位已被 ${existingPositionHolder[2]} 報名`);
      return { status: 'confirm_backup', message: `此崗位目前由 [${existingPositionHolder[2]}] 報名，您要改為報名備援嗎？` };
    }

    const maxAttendees = eventData[5];
    if (currentSignups.length >= maxAttendees) {
      logOperation('addSignup', eventId, userName, position, '額滿', '活動總人數已額滿');
      return { status: 'error', message: '活動總人數已額滿。' };
    }

    const newSignupId = 'su' + new Date().getTime();
    const newRowData = [newSignupId, eventId, userName, new Date(), position];
    signupsSheet.appendRow(newRowData);
    SpreadsheetApp.flush();

    const lastRowValues = signupsSheet.getRange(signupsSheet.getLastRow(), 1, 1, 5).getValues()[0];
    if (lastRowValues[0] === newSignupId && lastRowValues[2] === userName) {
      invalidateMasterDataCache_();
      logOperation('addSignup', eventId, userName, position, '成功', '報名資料寫入並驗證成功');
      return { status: 'success', message: '報名成功！' };
    } else {
      if (lastRowValues[0] === newSignupId) {
        signupsSheet.deleteRow(signupsSheet.getLastRow());
        logOperation('addSignup', eventId, userName, position, '失敗', '報名資料寫入驗證失敗，已嘗試回滾');
      } else {
        logOperation('addSignup', eventId, userName, position, '失敗', '報名資料寫入驗證失敗，無法回滾');
      }
      throw new Error('報名資料寫入驗證失敗，請稍後再試。');
    }
  } catch (err) {
    console.error("addSignup 失敗:", err.message, err.stack);
    logOperation('addSignup', eventId, userName, position, '失敗', `後端錯誤: ${err.message}`);
    throw new Error('報名時發生後端錯誤，請聯繫管理員。');
  } finally { lock.releaseLock(); }
}

function addBackupSignup(eventId, userName, position) {
  if (!eventId || !userName || !position) {
    logOperation('addBackupSignup', eventId, userName, position, '失敗', '缺少必要資訊');
    return { status: 'error', message: '缺少必要資訊。' };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    logOperation('addBackupSignup', eventId, userName, position, '失敗', '系統忙碌中，鎖定失敗');
    return { status: 'error', message: '系統忙碌中，請稍後再試。' };
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    const newSignupId = 'su' + new Date().getTime();
    const backupPosition = `${position} (備援)`;
    const newRowData = [newSignupId, eventId, userName, new Date(), backupPosition];
    signupsSheet.appendRow(newRowData);
    SpreadsheetApp.flush();

    const lastRowValues = signupsSheet.getRange(signupsSheet.getLastRow(), 1, 1, 5).getValues()[0];
    if (lastRowValues[0] === newSignupId && lastRowValues[2] === userName) {
      invalidateMasterDataCache_();
      logOperation('addBackupSignup', eventId, userName, backupPosition, '成功', '備援報名資料寫入並驗證成功');
      return { status: 'success' };
    } else {
      if (lastRowValues[0] === newSignupId) {
        signupsSheet.deleteRow(signupsSheet.getLastRow());
        logOperation('addBackupSignup', eventId, userName, backupPosition, '失敗', '備援報名資料寫入驗證失敗，已嘗試回滾');
      } else {
        logOperation('addBackupSignup', eventId, userName, backupPosition, '失敗', '備援報名資料寫入驗證失敗，無法回滾');
      }
      throw new Error('備援報名資料寫入驗證失敗，請稍後再試。');
    }
  } catch (err) {
    console.error("addBackupSignup 失敗:", err.message, err.stack);
    logOperation('addBackupSignup', eventId, userName, position, '失敗', `後端錯誤: ${err.message}`);
    throw new Error('備援報名時發生後端錯誤。');
  } finally { lock.releaseLock(); }
}

function removeSignup(eventId, userName) {
  if (!eventId || !userName) {
    logOperation('removeSignup', eventId, userName, '', '失敗', '缺少事件ID或姓名');
    return { status: 'error', message: '缺少事件ID或姓名。' };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    logOperation('removeSignup', eventId, userName, '', '失敗', '系統忙碌中，鎖定失敗');
    return { status: 'error', message: '系統忙碌中，無法取消報名，請稍後再試。' };
  }
  let removedPosition = '';
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    const signupsData = signupsSheet.getDataRange().getValues();

    for (let i = signupsData.length - 1; i >= 1; i--) {
      if (signupsData[i][1] === eventId && signupsData[i][2] === userName) {
        let rowToDeleteIndex = i + 1;
        removedPosition = signupsData[i][4];
        signupsSheet.deleteRow(rowToDeleteIndex);
        SpreadsheetApp.flush();
        invalidateMasterDataCache_();
        logOperation('removeSignup', eventId, userName, removedPosition, '成功', `已刪除行 ${rowToDeleteIndex}`);
        return { status: 'success', message: '已為您取消報名。' };
      }
    }
    logOperation('removeSignup', eventId, userName, '', '找不到', '找不到您的報名紀錄');
    return { status: 'error', message: '找不到您的報名紀錄。' };
  } catch (err) {
    console.error("removeSignup 失敗:", err.message, err.stack);
    logOperation('removeSignup', eventId, userName, removedPosition, '失敗', `後端錯誤: ${err.message}`);
    throw new Error('取消報名時發生錯誤。');
  } finally { lock.releaseLock(); }
}

function modifySignup(signupId, newPosition) {
  if (!signupId || !newPosition) {
    logOperation('modifySignup', signupId, '', newPosition, '失敗', '缺少報名ID或新崗位資訊');
    return { status: 'error', message: '缺少報名ID或新崗位資訊。' };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    logOperation('modifySignup', signupId, '', newPosition, '失敗', '系統忙碌中，鎖定失敗');
    return { status: 'error', message: '系統忙碌中，請稍後再試。' };
  }
  let eventId = '';
  let userName = '';
  let oldPosition = '';
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    const signupsData = signupsSheet.getDataRange().getValues();

    let targetRowIndex = -1;
    for (let i = 1; i < signupsData.length; i++) {
      if (signupsData[i][0] === signupId) {
        targetRowIndex = i + 1;
        eventId = signupsData[i][1];
        userName = signupsData[i][2];
        oldPosition = signupsData[i][4];
        break;
      }
    }
    if (targetRowIndex === -1) {
      logOperation('modifySignup', '', '', newPosition, '失敗', `找不到指定的報名紀錄ID: ${signupId}`);
      throw new Error('找不到指定的報名紀錄ID。');
    }

    const positionTaken = signupsData.some(row =>
      row[1] === eventId &&
      row[4] === newPosition &&
      row[0] !== signupId
    );

    if (positionTaken) {
      logOperation('modifySignup', eventId, userName, `${oldPosition} -> ${newPosition}`, '崗位重複', `無法修改，崗位 [${newPosition}] 已被其他人報名`);
      return { status: 'error', message: `無法修改，崗位 [${newPosition}] 已被其他人報名。` };
    }

    signupsSheet.getRange(targetRowIndex, 5).setValue(newPosition);
    SpreadsheetApp.flush();

    invalidateMasterDataCache_();
    logOperation('modifySignup', eventId, userName, `${oldPosition} -> ${newPosition}`, '成功', '崗位已成功修改');
    return { status: 'success', message: '崗位已成功修改！' };
  } catch (err) {
    console.error("modifySignup 失敗:", err.message, err.stack);
    logOperation('modifySignup', eventId, userName, `${oldPosition || ''} -> ${newPosition}`, '失敗', `後端錯誤: ${err.message}`);
    throw new Error('修改報名時發生後端錯誤。');
  } finally {
    lock.releaseLock();
  }
}

function createTempSheetAndExport(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const spreadsheetId = ss.getId();
  let tempSheet = null; let tempFile = null;
  try {
    const tempSheetName = "匯出報表_" + new Date().getTime();
    tempSheet = ss.insertSheet(tempSheetName);
    const data = getUnifiedSignups('', startDateStr, endDateStr);
    if (!data || data.length === 0) {
      ss.deleteSheet(tempSheet);
      return { status: 'nodata', message: '在選定的日期範圍內沒有任何報名記錄。' };
    }
    const title = `AV放送立願報名記錄 (${startDateStr || '所有'} - ${endDateStr || '所有'})`;
    tempSheet.getRange("A1:E1").merge().setValue(title).setFontWeight("bold").setFontSize(15).setHorizontalAlignment('center');
    const headers = ["活動日期", "活動", "崗位", "報名者", "報名時間"];
    const fields = ["eventDate", "eventTitle", "position", "user", "timestamp"];
    tempSheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3").setFontSize(15).setHorizontalAlignment('center');
    const outputData = data.map(row => fields.map(field => row[field] || ""));
    if (outputData.length > 0) { tempSheet.getRange(3, 1, outputData.length, headers.length).setValues(outputData).setFontSize(15).setHorizontalAlignment('left'); }
    tempSheet.autoResizeColumns(1, 5);
    SpreadsheetApp.flush();
    const blob = Drive.Files.export(spreadsheetId, MimeType.MICROSOFT_EXCEL, { gid: tempSheet.getSheetId(), alt: 'media' });
    blob.setName(`AV立願報名記錄_${new Date().toISOString().slice(0, 10)}.xlsx`);
    tempFile = DriveApp.createFile(blob);
    tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const downloadUrl = tempFile.getWebContentLink();
    scheduleDeletion(tempFile.getId());
    ss.deleteSheet(tempSheet);
    return { status: 'success', url: downloadUrl, fileId: tempFile.getId() };
  } catch (e) {
    console.error("createTempSheetAndExport 失敗: " + e.toString());
    if (tempFile) { try { DriveApp.getFileById(tempFile.getId()).setTrashed(true); } catch (f) { console.error("無法刪除臨時檔案: " + f.toString()); } }
    if (tempSheet && ss.getSheetByName(tempSheet.getName())) { ss.deleteSheet(tempSheet); }
    throw new Error('匯出時發生錯誤: ' + e.toString());
  }
}
function scheduleDeletion(fileId) { if (!fileId) return; try { const trigger = ScriptApp.newTrigger('triggeredDeleteHandler').timeBased().after(24 * 60 * 60 * 1000).create(); PropertiesService.getScriptProperties().setProperty(trigger.getUniqueId(), fileId); } catch (e) { console.error(`排程刪除檔案 '${fileId}' 時失敗: ${e.toString()}`); } }
function triggeredDeleteHandler(e) { const triggerId = e.triggerUid; const scriptProperties = PropertiesService.getScriptProperties(); const fileId = scriptProperties.getProperty(triggerId); if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) { console.error(`刪除檔案 ${fileId} 失敗: ${err.toString()}`); } scriptProperties.deleteProperty(triggerId); } const allTriggers = ScriptApp.getProjectTriggers(); for (let i = 0; i < allTriggers.length; i++) { if (allTriggers[i].getUniqueId() === triggerId) { ScriptApp.deleteTrigger(allTriggers[i]); break; } } }


// ===================== [LINE Messaging API (替代 LINE Notify)] =====================
const LINE_CHANNEL_ACCESS_TOKEN = 'MaSV+Q15S2BXNVYW6xroMRfJwIl5Oe8ZsRWRtuo9HpbQRssOXQceAUYPwuGQ9L8wa9pkbZyWudE+DFZ2MwRzHqBEkNUmM+gIen+YQyXncJGz4/MO+QuB4u6niSLPzUAM5wNacDP2uMXHc51laR2/3gdB04t89/1O/w1cDnyilFU='; // ★請至 LINE Developers Console 申請
const LINE_USER_ID = 'Ufefb59896be6a2030d5c97e25f00d5d7'; // ★請填入您的 User ID 或 Group ID
const LINE_TO_PROPERTY_KEY = 'AV_LIN_TO';
const LINE_TOKEN_PROPERTY_KEY = 'AV_LINE_CHANNEL_ACCESS_TOKEN';
const PROP_LINE_TARGETS = 'LINE_TARGETS';
const PROP_LINE_WEBHOOK_LAST_AT = 'LINE_WEBHOOK_LAST_AT';
const PROP_LINE_WEBHOOK_LAST_EVENT_COUNT = 'LINE_WEBHOOK_LAST_EVENT_COUNT';
const PROP_LINE_WEBHOOK_LAST_ERROR = 'LINE_WEBHOOK_LAST_ERROR';
const PROP_LINE_WEBHOOK_LAST_SOURCE = 'LINE_WEBHOOK_LAST_SOURCE';
const PROP_LINE_WEBHOOK_LAST_REPLY_STATUS = 'LINE_WEBHOOK_LAST_REPLY_STATUS';
const PROP_LINE_WEBHOOK_LAST_REPLY_ERROR = 'LINE_WEBHOOK_LAST_REPLY_ERROR';
const CALENDAR_LINK_CURRENT_MONTH_KEY = 'AV_CALENDAR_LINK_CURRENT_MONTH';
const CALENDAR_LINK_NEXT_MONTH_KEY = 'AV_CALENDAR_LINK_NEXT_MONTH';
const LINE_SUBSCRIPTION_MESSAGE_TYPES = ['tomorrow_signup', 'marquee_announcement'];
const LINE_EVENT_CACHE_TTL_SEC = 6 * 60 * 60;

function getScriptTimeZone_() {
  return Session.getScriptTimeZone() || 'Asia/Taipei';
}

function formatInScriptTimeZone_(date, pattern) {
  return Utilities.formatDate(date, getScriptTimeZone_(), pattern);
}

function resolveLineToken() {
  const fromProps = normalizeLineToken_(PropertiesService.getScriptProperties().getProperty(LINE_TOKEN_PROPERTY_KEY));
  if (fromProps) return fromProps;
  return normalizeLineToken_(LINE_CHANNEL_ACCESS_TOKEN);
}

function resolveLineTo(linTo) {
  const normalized = String(linTo || '').trim();
  if (normalized) {
    PropertiesService.getScriptProperties().setProperty(LINE_TO_PROPERTY_KEY, normalized);
    return normalized;
  }
  const saved = String(PropertiesService.getScriptProperties().getProperty(LINE_TO_PROPERTY_KEY) || '').trim();
  if (saved) return saved;
  return String(LINE_USER_ID || '').trim();
}

function parseLineTargets_(lineToRaw) {
  if (Array.isArray(lineToRaw)) {
    return lineToRaw.map(v => String(v || '').trim()).filter(Boolean);
  }
  return String(lineToRaw || '')
    .split(/[\s,;\n\r]+/)
    .map(v => v.trim())
    .filter(Boolean);
}

function dedupeLineTargetIds_(items) {
  const out = [];
  const seen = {};
  (Array.isArray(items) ? items : []).forEach(item => {
    const value = String(item || '').trim();
    if (!value || seen[value]) return;
    seen[value] = true;
    out.push(value);
  });
  return out;
}

function isValidLineTargetId_(targetId) {
  return /^[UCR][0-9a-fA-F]{32}$/.test(String(targetId || '').trim());
}

function findInvalidLineTargetIds_(targets) {
  return dedupeLineTargetIds_((Array.isArray(targets) ? targets : []).filter(targetId => !isValidLineTargetId_(targetId)));
}

function normalizeLineSubscribedMessageTypes_(types) {
  let list = [];
  if (Array.isArray(types)) {
    list = types;
  } else if (typeof types === 'string') {
    list = types.split(/[\s,;|/]+/);
  }
  const normalized = dedupeLineTargetIds_(list.map(item => {
    const raw = String(item || '').trim().toLowerCase();
    if (!raw) return '';
    if (raw === 'all' || raw === '全部') return 'all';
    if (raw === 'tomorrow_signup' || raw === '明天' || raw === '明天名單') return 'tomorrow_signup';
    if (raw === 'marquee_announcement' || raw === '公告' || raw === '跑馬燈' || raw === '跑馬燈公告') return 'marquee_announcement';
    return '';
  }));
  if (normalized.indexOf('all') >= 0 || normalized.length === 0) {
    return LINE_SUBSCRIPTION_MESSAGE_TYPES.slice();
  }
  return normalized.filter(type => LINE_SUBSCRIPTION_MESSAGE_TYPES.indexOf(type) >= 0);
}

function normalizeLineSourceStatus_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (normalized === 'blocked' || normalized === 'left' || normalized === 'inactive') return normalized;
  return 'active';
}

function normalizeLineTargetRecord_(item) {
  const row = item || {};
  const status = normalizeLineSourceStatus_(row.source_status || (row.notifications_enabled === false ? 'inactive' : 'active'));
  const notificationsEnabled = status === 'active' ? row.notifications_enabled !== false : false;
  return {
    target_id: String(row.target_id || '').trim(),
    source_type: String(row.source_type || '').trim(),
    user_id: String(row.user_id || '').trim(),
    group_id: String(row.group_id || '').trim(),
    room_id: String(row.room_id || '').trim(),
    bound_user_name: normalizeLineBoundUserName_(row.bound_user_name || ''),
    last_seen_at: String(row.last_seen_at || '').trim(),
    last_interaction_at: String(row.last_interaction_at || row.last_seen_at || '').trim(),
    last_event_type: String(row.last_event_type || '').trim(),
    source_status: status,
    notifications_enabled: notificationsEnabled,
    subscribed_message_types: normalizeLineSubscribedMessageTypes_(row.subscribed_message_types),
    alias: normalizeLineTargetAlias_(row.alias || '')
  };
}

function getRecordedLineTargets_() {
  return getLineTargets_();
}

function resolveLineTargetsForSend_(linTo, options) {
  const opts = options || {};
  const messageType = String(opts.message_type || '').trim();
  const respectPreferences = Boolean(opts.respect_preferences);
  const directTargets = parseLineTargets_(linTo);
  if (directTargets.length) {
    return dedupeLineTargetIds_(directTargets);
  }

  const savedTargetLineToRaw = String(PropertiesService.getScriptProperties().getProperty(LINE_TO_PROPERTY_KEY) || '').trim();
  const configuredTargets = parseLineTargets_(savedTargetLineToRaw);
  if (configuredTargets.length) {
    return dedupeLineTargetIds_(configuredTargets);
  }

  const recordedTargets = getRecordedLineTargets_();
  if (recordedTargets.length) {
    const filteredTargets = recordedTargets
      .filter(item => {
        const targetId = String(item && item.target_id || '').trim();
        if (!targetId) return false;
        const status = normalizeLineSourceStatus_(item && item.source_status || '');
        if (status === 'blocked' || status === 'left') return false;
        if (!respectPreferences) return true;
        if (item && item.notifications_enabled === false) return false;
        if (!messageType) return true;
        const subscribedTypes = normalizeLineSubscribedMessageTypes_(item && item.subscribed_message_types);
        return subscribedTypes.indexOf(messageType) >= 0;
      })
      .map(item => String(item && item.target_id || '').trim());
    return dedupeLineTargetIds_(filteredTargets);
  }

  return dedupeLineTargetIds_(parseLineTargets_(String(LINE_USER_ID || '').trim()));
}

function getLineTargets_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_LINE_TARGETS);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed)
      ? parsed.map(item => normalizeLineTargetRecord_(item)).filter(item => item.target_id)
      : [];
  } catch (e) {
    console.error('解析 LINE_TARGETS 失敗: ' + e.toString());
    return [];
  }
}

function setLineTargets_(items) {
  const arr = Array.isArray(items)
    ? items.map(item => normalizeLineTargetRecord_(item)).filter(item => item.target_id)
    : [];
  PropertiesService.getScriptProperties().setProperty(PROP_LINE_TARGETS, JSON.stringify(arr));
}

function normalizeLineTargetAlias_(rawAlias) {
  return String(rawAlias || '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
}

function normalizeLineBoundUserName_(rawValue) {
  return String(rawValue || '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 40);
}

function getLineWebhookStatus_() {
  const props = PropertiesService.getScriptProperties();
  const targets = getLineTargets_();
  const activeTargets = targets.filter(item => normalizeLineSourceStatus_(item.source_status) === 'active');
  const notificationsEnabledTargets = activeTargets.filter(item => item.notifications_enabled !== false);
  return {
    last_at: String(props.getProperty(PROP_LINE_WEBHOOK_LAST_AT) || ''),
    last_event_count: Number(props.getProperty(PROP_LINE_WEBHOOK_LAST_EVENT_COUNT) || 0),
    last_source: String(props.getProperty(PROP_LINE_WEBHOOK_LAST_SOURCE) || ''),
    last_error: String(props.getProperty(PROP_LINE_WEBHOOK_LAST_ERROR) || ''),
    last_reply_status: String(props.getProperty(PROP_LINE_WEBHOOK_LAST_REPLY_STATUS) || ''),
    last_reply_error: String(props.getProperty(PROP_LINE_WEBHOOK_LAST_REPLY_ERROR) || ''),
    targets_count: targets.length,
    active_targets_count: activeTargets.length,
    notifications_enabled_count: notificationsEnabledTargets.length
  };
}

function normalizeCalendarDownloadLink_(rawValue) {
  return String(rawValue || '').trim();
}

function validateCalendarDownloadLink_(url, label) {
  const value = normalizeCalendarDownloadLink_(url);
  if (!value) return '';
  if (!/^https?:\/\//i.test(value)) {
    throw new Error(`${label} 連結格式不正確，請輸入完整的 https:// 或 http:// 網址。`);
  }
  return value;
}

function getCalendarDownloadLinks_() {
  const props = PropertiesService.getScriptProperties();
  return {
    current_month_url: normalizeCalendarDownloadLink_(props.getProperty(CALENDAR_LINK_CURRENT_MONTH_KEY)),
    next_month_url: normalizeCalendarDownloadLink_(props.getProperty(CALENDAR_LINK_NEXT_MONTH_KEY))
  };
}

function setCalendarDownloadLinks_(currentMonthUrl, nextMonthUrl) {
  const props = PropertiesService.getScriptProperties();
  const normalizedCurrent = validateCalendarDownloadLink_(currentMonthUrl, '本月行事曆');
  const normalizedNext = validateCalendarDownloadLink_(nextMonthUrl, '次月行事曆');

  if (normalizedCurrent) props.setProperty(CALENDAR_LINK_CURRENT_MONTH_KEY, normalizedCurrent);
  else props.deleteProperty(CALENDAR_LINK_CURRENT_MONTH_KEY);

  if (normalizedNext) props.setProperty(CALENDAR_LINK_NEXT_MONTH_KEY, normalizedNext);
  else props.deleteProperty(CALENDAR_LINK_NEXT_MONTH_KEY);

  return {
    status: 'success',
    current_month_url: normalizedCurrent,
    next_month_url: normalizedNext
  };
}

function getLineSendLogs_(limit) {
  const requestedLimit = Number(limit || 50);
  const normalizedLimit = Math.min(Math.max(requestedLimit, 1), 500);
  const sheet = ensureLineSendLogSheet_();
  const totalRows = Math.max(sheet.getLastRow() - 1, 0);

  if (totalRows === 0) {
    return {
      rows: [],
      total_count: 0,
      returned_count: 0,
      limit: normalizedLimit,
      sheet_name: LINE_SEND_LOGS_SHEET_NAME
    };
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 18);
  const startRow = Math.max(2, sheet.getLastRow() - normalizedLimit + 1);
  const rowCount = sheet.getLastRow() - startRow + 1;
  const values = sheet.getRange(startRow, 1, rowCount, lastColumn).getDisplayValues();

  const rows = values
    .map((row, index) => ({
      sheet_row: startRow + index,
      timestamp: String(row[0] || ''),
      batch_id: String(row[1] || ''),
      trigger_source: String(row[2] || ''),
      message_type: String(row[3] || ''),
      function_name: String(row[4] || ''),
      line_to_input: String(row[5] || ''),
      resolved_target_count: Number(row[6] || 0),
      target: String(row[7] || ''),
      target_alias: String(row[8] || ''),
      chunk_index: Number(row[9] || 0),
      chunk_count: Number(row[10] || 0),
      message_length: Number(row[11] || 0),
      message_preview: String(row[12] || ''),
      message_body: String(row[13] || ''),
      sent: String(row[14] || '').toUpperCase() === 'TRUE',
      http_status: Number(row[15] || 0),
      error: String(row[16] || ''),
      response_body: String(row[17] || '')
    }))
    .reverse();

  return {
    rows: rows,
    total_count: totalRows,
    returned_count: rows.length,
    limit: normalizedLimit,
    sheet_name: LINE_SEND_LOGS_SHEET_NAME
  };
}

function deleteLineSendLog_(sheetRow) {
  const rowNumber = Number(sheetRow || 0);
  if (!Number.isInteger(rowNumber) || rowNumber < 2) {
    return { status: 'error', error: '無效的 log 列號。' };
  }

  const sheet = ensureLineSendLogSheet_();
  const lastRow = sheet.getLastRow();
  if (rowNumber > lastRow) {
    return { status: 'error', error: '指定的 log 不存在或已被刪除。' };
  }

  sheet.deleteRow(rowNumber);
  SpreadsheetApp.flush();
  return {
    status: 'success',
    deleted_row: rowNumber,
    remaining_count: Math.max(sheet.getLastRow() - 1, 0)
  };
}

function inferLineToType_(targetId) {
  const s = String(targetId || '');
  if (s.startsWith('U')) return 'user';
  if (s.startsWith('C')) return 'group';
  if (s.startsWith('R')) return 'room';
  return 'unknown';
}

function getLineEventTargetId_(event) {
  const source = (event && event.source) || {};
  return String(source.groupId || source.roomId || source.userId || '').trim();
}

function buildLineEventCacheKey_(event) {
  if (!event) return '';
  const webhookEventId = String(event.webhookEventId || '').trim();
  if (webhookEventId) return `line_evt_${webhookEventId}`;
  const parts = [
    String(event.type || ''),
    String(event.timestamp || ''),
    String(event.replyToken || ''),
    getLineEventTargetId_(event),
    String(event && event.message && event.message.id || ''),
    String(event && event.postback && event.postback.data || '')
  ];
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, parts.join('|'));
  return `line_evt_${Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '')}`;
}

function shouldProcessLineEvent_(event) {
  const cacheKey = buildLineEventCacheKey_(event);
  if (!cacheKey) return true;
  const cache = CacheService.getScriptCache();
  if (cache.get(cacheKey)) return false;
  cache.put(cacheKey, '1', LINE_EVENT_CACHE_TTL_SEC);
  return true;
}

function resolveLineTargetStatusFromEvent_(eventType, previousStatus) {
  if (eventType === 'unfollow') return 'blocked';
  if (eventType === 'leave') return 'left';
  if (eventType === 'follow' || eventType === 'join' || eventType === 'message' || eventType === 'postback') return 'active';
  return normalizeLineSourceStatus_(previousStatus);
}

function buildLineTargetFromEvent_(event, existingRecord, nowIso) {
  const source = (event && event.source) || {};
  const targetId = getLineEventTargetId_(event);
  if (!targetId) return null;
  const existing = normalizeLineTargetRecord_(existingRecord || { target_id: targetId });
  const nextStatus = resolveLineTargetStatusFromEvent_(String(event && event.type || '').trim(), existing.source_status);
  return normalizeLineTargetRecord_(Object.assign({}, existing, {
    target_id: targetId,
    source_type: String(source.type || inferLineToType_(targetId)),
    user_id: String(source.userId || existing.user_id || ''),
    group_id: String(source.groupId || existing.group_id || ''),
    room_id: String(source.roomId || existing.room_id || ''),
    last_seen_at: nowIso,
    last_interaction_at: nowIso,
    last_event_type: String(event && event.type || ''),
    source_status: nextStatus,
    notifications_enabled: nextStatus === 'active' ? existing.notifications_enabled !== false : false,
    subscribed_message_types: existing.subscribed_message_types
  }));
}

function normalizeLineCommandText_(text) {
  return String(text || '')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeLineDateInput_(rawDate) {
  const normalized = String(rawDate || '').trim().replace(/[./]/g, '-');
  const parsed = new Date(normalized);
  if (isNaN(parsed.getTime())) return '';
  return formatInScriptTimeZone_(parsed, 'yyyy-MM-dd');
}

function getCurrentWeekDateRange_() {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const day = today.getDay();
  const diffToMonday = day === 0 ? -6 : 1 - day;
  const start = new Date(today);
  start.setDate(today.getDate() + diffToMonday);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  return {
    startDateStr: formatInScriptTimeZone_(start, 'yyyy-MM-dd'),
    endDateStr: formatInScriptTimeZone_(end, 'yyyy-MM-dd')
  };
}

function getCurrentMonthDateRange_() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  return {
    startDateStr: formatInScriptTimeZone_(start, 'yyyy-MM-dd'),
    endDateStr: formatInScriptTimeZone_(end, 'yyyy-MM-dd')
  };
}

function buildLineQuickReplyItem_(label, text) {
  return {
    type: 'action',
    action: {
      type: 'message',
      label: String(label || '').slice(0, 20),
      text: String(text || '').slice(0, 300)
    }
  };
}

function buildLineTextMessage_(text, quickReplyItems) {
  const message = { type: 'text', text: String(text || '').slice(0, 5000) };
  const items = Array.isArray(quickReplyItems) ? quickReplyItems.filter(Boolean).slice(0, 13) : [];
  if (items.length) {
    message.quickReply = { items: items };
  }
  return message;
}

function buildLineMenuTemplateMessage_() {
  return {
    type: 'template',
    altText: 'AV 放送 Bot 功能選單',
    template: {
      type: 'buttons',
      title: 'AV 放送 Bot',
      text: '快速查名單與通知設定',
      actions: [
        { type: 'message', label: '今天名單', text: '今天名單' },
        { type: 'message', label: '明天名單', text: '明天名單' },
        { type: 'message', label: '本週名單', text: '本週名單' },
        { type: 'message', label: '通知設定', text: '通知設定' }
      ]
    }
  };
}

function buildLineNotificationTemplateMessage_() {
  return {
    type: 'template',
    altText: '通知設定',
    template: {
      type: 'buttons',
      title: '通知設定',
      text: '快速切換通知偏好',
      actions: [
        { type: 'message', label: '啟用通知', text: '啟用通知' },
        { type: 'message', label: '停止通知', text: '停止通知' },
        { type: 'message', label: '只收明天名單', text: '只收明天名單' },
        { type: 'message', label: '只收公告', text: '只收公告' }
      ]
    }
  };
}

function summarizeLineTargetPreferences_(targetRecord) {
  const target = normalizeLineTargetRecord_(targetRecord);
  const typeLabels = target.subscribed_message_types.map(type => type === 'tomorrow_signup' ? '明天名單' : '公告');
  const statusLabel = target.notifications_enabled ? '啟用' : '停用';
  const sourceStatus = target.source_status === 'blocked'
    ? '已封鎖'
    : (target.source_status === 'left' ? '已離開' : (target.source_status === 'inactive' ? '未啟用' : '正常'));
  return `通知：${statusLabel}\n訂閱：${typeLabels.join('、') || '無'}\n來源狀態：${sourceStatus}${target.bound_user_name ? `\n綁定姓名：${target.bound_user_name}` : ''}${target.alias ? `\n標註：${target.alias}` : ''}`;
}

function buildLineMenuMessages_(targetRecord) {
  const summary = summarizeLineTargetPreferences_(targetRecord);
  const intro = [
    '可用指令：',
    '1. 今天名單 / 明天名單 / 本週名單 / 本月名單',
    '2. 查詢 2026-03-08 或 查詢 2026-03-08 2026-03-12',
    '3. 今天有什麼法會 / 本週有哪些法會 / 本月有哪些法會 / 查活動 2026-03-08',
    '4. 我是 王小明 / 解除綁定 / 我的報名 / 我的本週報名',
    '5. 公告 / 本月行事曆 / 次月行事曆',
    '6. 通知設定 / 啟用通知 / 停止通知 / 訂閱全部 / 只收明天名單 / 只收公告',
    '',
    summary
  ].join('\n');
  return [
    buildLineTextMessage_(intro, [
      buildLineQuickReplyItem_('今天名單', '今天名單'),
      buildLineQuickReplyItem_('明天名單', '明天名單'),
      buildLineQuickReplyItem_('本週名單', '本週名單'),
      buildLineQuickReplyItem_('本月名單', '本月名單'),
      buildLineQuickReplyItem_('今天有什麼法會', '今天有什麼法會'),
      buildLineQuickReplyItem_('我的報名', '我的報名'),
      buildLineQuickReplyItem_('公告', '公告'),
      buildLineQuickReplyItem_('本月行事曆', '本月行事曆'),
      buildLineQuickReplyItem_('次月行事曆', '次月行事曆'),
      buildLineQuickReplyItem_('通知設定', '通知設定')
    ]),
    buildLineMenuTemplateMessage_()
  ];
}

function buildLineNotificationMessages_(targetRecord, leadText) {
  return [
    buildLineTextMessage_(`${leadText ? leadText + '\n\n' : ''}${summarizeLineTargetPreferences_(targetRecord)}`, [
      buildLineQuickReplyItem_('啟用通知', '啟用通知'),
      buildLineQuickReplyItem_('停止通知', '停止通知'),
      buildLineQuickReplyItem_('訂閱全部', '訂閱全部'),
      buildLineQuickReplyItem_('只收明天名單', '只收明天名單'),
      buildLineQuickReplyItem_('只收公告', '只收公告')
    ]),
    buildLineNotificationTemplateMessage_()
  ];
}

function canBindLineUserName_(targetRecord) {
  return String(targetRecord && targetRecord.source_type || '').trim() === 'user';
}

function applyLineTargetBoundUserName_(targetRecord, userName) {
  const target = normalizeLineTargetRecord_(targetRecord);
  target.bound_user_name = normalizeLineBoundUserName_(userName);
  target.last_interaction_at = new Date().toISOString();
  target.last_event_type = target.bound_user_name ? 'bind_user_name' : 'unbind_user_name';
  return target;
}

function getUserSignupsInRange_(userName, startDateStr, endDateStr) {
  const normalizedUserName = normalizeLineBoundUserName_(userName);
  if (!normalizedUserName) return [];
  return getUnifiedSignups('', startDateStr, endDateStr)
    .filter(item => normalizeLineBoundUserName_(item && item.user || '') === normalizedUserName);
}

function getEventsInRange_(startDateStr, endDateStr) {
  const start = startDateStr ? new Date(startDateStr) : null;
  const end = endDateStr ? new Date(endDateStr) : null;
  if (start) start.setHours(0, 0, 0, 0);
  if (end) end.setHours(23, 59, 59, 999);
  return getEventsAndSignups().filter(event => {
    const eventDate = new Date(event && event.start || '');
    if (isNaN(eventDate.getTime())) return false;
    if (start && eventDate < start) return false;
    if (end && eventDate > end) return false;
    return true;
  });
}

function buildMySignupsReplyMessages_(targetRecord, label, startDateStr, endDateStr) {
  const target = normalizeLineTargetRecord_(targetRecord);
  if (!canBindLineUserName_(target)) {
    return [buildLineTextMessage_('「我的報名」只支援 1 對 1 私訊使用。請改在個人聊天室中操作。')];
  }
  if (!target.bound_user_name) {
    return [buildLineTextMessage_('尚未綁定姓名。請先輸入「我是 王小明」。', [
      buildLineQuickReplyItem_('我是 王小明', '我是 王小明'),
      buildLineQuickReplyItem_('通知設定', '通知設定'),
      buildLineQuickReplyItem_('選單', '選單')
    ])];
  }
  const signups = getUserSignupsInRange_(target.bound_user_name, startDateStr, endDateStr);
  if (!signups.length) {
    return [buildLineTextMessage_(`${label}沒有 ${target.bound_user_name} 的報名資料。`, [
      buildLineQuickReplyItem_('今天名單', '今天名單'),
      buildLineQuickReplyItem_('本週有哪些法會', '本週有哪些法會'),
      buildLineQuickReplyItem_('通知設定', '通知設定')
    ])];
  }
  const body = [
    `【${label}】`,
    `姓名：${target.bound_user_name}`,
    `區間：${startDateStr} ~ ${endDateStr}`,
    '----------',
    ''
  ].join('\n') + signups.map((item, index) => (
    `${index + 1}. ${item.eventDate} (${item.eventDayOfWeek || '-'}) ${item.startTime || ''}${item.endTime ? ' - ' + item.endTime : ''}\n${item.eventTitle}\n崗位：${item.position}`
  )).join('\n\n');
  return buildLineTextMessagesFromLongText_(body, [
    buildLineQuickReplyItem_('我的本週報名', '我的本週報名'),
    buildLineQuickReplyItem_('本週有哪些法會', '本週有哪些法會'),
    buildLineQuickReplyItem_('通知設定', '通知設定')
  ]);
}

function buildEventsInRangeReplyMessages_(label, startDateStr, endDateStr) {
  const events = getEventsInRange_(startDateStr, endDateStr);
  if (!events.length) {
    return [buildLineTextMessage_(`${label}沒有法會資料。`, [
      buildLineQuickReplyItem_('今天名單', '今天名單'),
      buildLineQuickReplyItem_('公告', '公告'),
      buildLineQuickReplyItem_('選單', '選單')
    ])];
  }
  const body = [
    `【${label}】`,
    `區間：${startDateStr} ~ ${endDateStr}`,
    '----------',
    ''
  ].join('\n') + events.map((event, index) => {
    const startAt = new Date(event.start);
    const dateText = isNaN(startAt.getTime()) ? '' : formatInScriptTimeZone_(startAt, 'yyyy-MM-dd');
    const positions = Array.isArray(event.extendedProps && event.extendedProps.positions)
      ? event.extendedProps.positions.join('、')
      : '';
    const signupCount = Array.isArray(event.extendedProps && event.extendedProps.signups)
      ? event.extendedProps.signups.length
      : 0;
    const maxAttendees = Number(event.extendedProps && event.extendedProps.maxAttendees || 999);
    return `${index + 1}. ${dateText} ${event.extendedProps.startTime || ''}${event.extendedProps.endTime ? ' - ' + event.extendedProps.endTime : ''}\n${event.extendedProps.full_title || event.title}\n已報 ${signupCount}/${maxAttendees}${positions ? `\n崗位：${positions}` : ''}`;
  }).join('\n\n');
  return buildLineTextMessagesFromLongText_(body, [
    buildLineQuickReplyItem_('今天有什麼法會', '今天有什麼法會'),
    buildLineQuickReplyItem_('本週有哪些法會', '本週有哪些法會'),
    buildLineQuickReplyItem_('本月有哪些法會', '本月有哪些法會')
  ]);
}

function buildLineTextMessagesFromLongText_(text, quickReplyItems) {
  const chunks = splitLineMessageChunks_(String(text || '').trim(), 4500);
  if (!chunks.length) return [buildLineTextMessage_('目前沒有可顯示的內容。', quickReplyItems)];
  const limitedChunks = chunks.slice(0, 5);
  if (chunks.length > 5) {
    limitedChunks[4] = `${limitedChunks[4].slice(0, 4300)}\n\n...(內容過長，請改用網站查看完整資料)`;
  }
  return limitedChunks.map((chunk, index) => buildLineTextMessage_(
    chunk,
    index === limitedChunks.length - 1 ? quickReplyItems : null
  ));
}

function buildSignupsReplyMessages_(result, emptyText) {
  if (!result || result.status === 'nodata') {
    return [buildLineTextMessage_(emptyText || '查無資料。', [
      buildLineQuickReplyItem_('明天名單', '明天名單'),
      buildLineQuickReplyItem_('本週名單', '本週名單'),
      buildLineQuickReplyItem_('通知設定', '通知設定')
    ])];
  }
  if (result.status !== 'success' || !result.text) {
    return [buildLineTextMessage_('查詢失敗，請稍後再試。')];
  }
  return buildLineTextMessagesFromLongText_(result.text, [
    buildLineQuickReplyItem_('今天名單', '今天名單'),
    buildLineQuickReplyItem_('明天名單', '明天名單'),
    buildLineQuickReplyItem_('本週名單', '本週名單'),
    buildLineQuickReplyItem_('本月名單', '本月名單')
  ]);
}

function buildAnnouncementsReplyMessages_() {
  const announcements = getNewsAnnouncements();
  if (!announcements.length) {
    return [buildLineTextMessage_('目前沒有最新公告。')];
  }
  const body = `【跑馬燈公告】\n更新時間：${formatInScriptTimeZone_(new Date(), 'yyyy/MM/dd HH:mm')}\n\n${announcements.map((item, index) => `${index + 1}. ${item}`).join('\n')}`;
  return buildLineTextMessagesFromLongText_(body, [
    buildLineQuickReplyItem_('本月行事曆', '本月行事曆'),
    buildLineQuickReplyItem_('次月行事曆', '次月行事曆'),
    buildLineQuickReplyItem_('通知設定', '通知設定')
  ]);
}

function buildCalendarLinkReplyMessages_(linkType) {
  const links = getCalendarDownloadLinks_();
  const isCurrent = linkType !== 'next';
  const label = isCurrent ? '本月行事曆' : '次月行事曆';
  const url = isCurrent ? String(links.current_month_url || '') : String(links.next_month_url || '');
  if (!url) {
    return [buildLineTextMessage_(`${label} 連結尚未設定。`)];
  }
  return [buildLineTextMessage_(`【${label}】\n${url}`, [
    buildLineQuickReplyItem_('本月行事曆', '本月行事曆'),
    buildLineQuickReplyItem_('次月行事曆', '次月行事曆'),
    buildLineQuickReplyItem_('公告', '公告')
  ])];
}

function parseLineCommand_(text) {
  const normalized = normalizeLineCommandText_(text);
  const lower = normalized.toLowerCase();
  if (!normalized) return null;
  if (lower === 'help' || lower === 'menu' || normalized === '幫助' || normalized === '選單' || normalized === '功能') {
    return { type: 'menu' };
  }
  if (normalized === '今天名單') return { type: 'today' };
  if (normalized === '明天名單') return { type: 'tomorrow' };
  if (normalized === '本週名單') return { type: 'week' };
  if (normalized === '本月名單') return { type: 'month' };
  if (normalized === '我的報名') return { type: 'my_signups' };
  if (normalized === '我的本週報名') return { type: 'my_week_signups' };
  if (normalized === '今天有什麼法會') return { type: 'events_today' };
  if (normalized === '本週有哪些法會') return { type: 'events_week' };
  if (normalized === '本月有哪些法會') return { type: 'events_month' };
  if (normalized === '公告' || normalized === '最新公告') return { type: 'announcements' };
  if (normalized === '本月行事曆') return { type: 'calendar_current' };
  if (normalized === '次月行事曆') return { type: 'calendar_next' };
  if (normalized === '通知設定') return { type: 'notification_status' };
  if (normalized === '啟用通知') return { type: 'notification_update', mode: 'enable' };
  if (normalized === '停止通知' || normalized === '停用通知') return { type: 'notification_update', mode: 'disable' };
  if (normalized === '訂閱全部') return { type: 'notification_update', mode: 'all' };
  if (normalized === '只收明天名單') return { type: 'notification_update', mode: 'tomorrow_only' };
  if (normalized === '只收公告') return { type: 'notification_update', mode: 'marquee_only' };
  if (normalized === '解除綁定') return { type: 'unbind_user_name' };
  const bindMatch = normalized.match(/^(?:我是|綁定姓名|姓名綁定)\s+(.+)$/);
  if (bindMatch) {
    return { type: 'bind_user_name', userName: normalizeLineBoundUserName_(bindMatch[1]) };
  }
  const eventRangeMatch = normalized.match(/^查活動\s+(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})(?:\s+(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}))?$/);
  if (eventRangeMatch) {
    const startDateStr = normalizeLineDateInput_(eventRangeMatch[1]);
    const endDateStr = normalizeLineDateInput_(eventRangeMatch[2] || eventRangeMatch[1]);
    if (startDateStr && endDateStr) {
      return { type: 'events_range', startDateStr: startDateStr, endDateStr: endDateStr };
    }
  }
  const rangeMatch = normalized.match(/^(?:查詢|名單)\s+(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})(?:\s+(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}))?$/);
  if (rangeMatch) {
    const startDateStr = normalizeLineDateInput_(rangeMatch[1]);
    const endDateStr = normalizeLineDateInput_(rangeMatch[2] || rangeMatch[1]);
    if (startDateStr && endDateStr) {
      return { type: 'date_range', startDateStr: startDateStr, endDateStr: endDateStr };
    }
  }
  return null;
}

function applyLineTargetNotificationMode_(targetRecord, mode) {
  const target = normalizeLineTargetRecord_(targetRecord);
  if (mode === 'disable') {
    target.notifications_enabled = false;
  } else if (mode === 'enable') {
    target.notifications_enabled = true;
    if (!target.subscribed_message_types.length) {
      target.subscribed_message_types = LINE_SUBSCRIPTION_MESSAGE_TYPES.slice();
    }
    target.source_status = 'active';
  } else if (mode === 'tomorrow_only') {
    target.notifications_enabled = true;
    target.subscribed_message_types = ['tomorrow_signup'];
    target.source_status = 'active';
  } else if (mode === 'marquee_only') {
    target.notifications_enabled = true;
    target.subscribed_message_types = ['marquee_announcement'];
    target.source_status = 'active';
  } else {
    target.notifications_enabled = true;
    target.subscribed_message_types = LINE_SUBSCRIPTION_MESSAGE_TYPES.slice();
    target.source_status = 'active';
  }
  target.last_interaction_at = new Date().toISOString();
  target.last_event_type = 'preference_update';
  return target;
}

function buildLineCommandReplyMessages_(command, targetRecord) {
  if (!command) return [];
  if (command.type === 'menu') {
    return buildLineMenuMessages_(targetRecord);
  }
  if (command.type === 'today') {
    return buildSignupsReplyMessages_(getSignupsAsTextForToday(), '今天沒有報名資料。');
  }
  if (command.type === 'tomorrow') {
    return buildSignupsReplyMessages_(getSignupsAsTextForTomorrow(), '明天沒有報名資料。');
  }
  if (command.type === 'week') {
    const range = getCurrentWeekDateRange_();
    return buildSignupsReplyMessages_(getSignupsAsText(range.startDateStr, range.endDateStr), '本週沒有報名資料。');
  }
  if (command.type === 'month') {
    const range = getCurrentMonthDateRange_();
    return buildSignupsReplyMessages_(getSignupsAsText(range.startDateStr, range.endDateStr), '本月沒有報名資料。');
  }
  if (command.type === 'my_signups') {
    const range = getCurrentMonthDateRange_();
    return buildMySignupsReplyMessages_(targetRecord, '我的本月報名', range.startDateStr, range.endDateStr);
  }
  if (command.type === 'my_week_signups') {
    const range = getCurrentWeekDateRange_();
    return buildMySignupsReplyMessages_(targetRecord, '我的本週報名', range.startDateStr, range.endDateStr);
  }
  if (command.type === 'date_range') {
    return buildSignupsReplyMessages_(getSignupsAsText(command.startDateStr, command.endDateStr), '指定日期區間沒有報名資料。');
  }
  if (command.type === 'events_today') {
    const today = formatInScriptTimeZone_(new Date(), 'yyyy-MM-dd');
    return buildEventsInRangeReplyMessages_('今天法會', today, today);
  }
  if (command.type === 'events_week') {
    const range = getCurrentWeekDateRange_();
    return buildEventsInRangeReplyMessages_('本週法會', range.startDateStr, range.endDateStr);
  }
  if (command.type === 'events_month') {
    const range = getCurrentMonthDateRange_();
    return buildEventsInRangeReplyMessages_('本月法會', range.startDateStr, range.endDateStr);
  }
  if (command.type === 'events_range') {
    return buildEventsInRangeReplyMessages_('指定區間法會', command.startDateStr, command.endDateStr);
  }
  if (command.type === 'announcements') {
    return buildAnnouncementsReplyMessages_();
  }
  if (command.type === 'calendar_current') {
    return buildCalendarLinkReplyMessages_('current');
  }
  if (command.type === 'calendar_next') {
    return buildCalendarLinkReplyMessages_('next');
  }
  if (command.type === 'bind_user_name') {
    if (!canBindLineUserName_(targetRecord)) {
      return [buildLineTextMessage_('姓名綁定只支援 1 對 1 私訊。請改在個人聊天室中輸入。')];
    }
    if (!command.userName) {
      return [buildLineTextMessage_('請用「我是 王小明」這種格式綁定姓名。')];
    }
    const updatedTarget = applyLineTargetBoundUserName_(targetRecord, command.userName);
    return [buildLineTextMessage_(`已綁定姓名：${updatedTarget.bound_user_name}\n之後可直接輸入「我的報名」或「我的本週報名」。`, [
      buildLineQuickReplyItem_('我的報名', '我的報名'),
      buildLineQuickReplyItem_('我的本週報名', '我的本週報名'),
      buildLineQuickReplyItem_('通知設定', '通知設定')
    ])];
  }
  if (command.type === 'unbind_user_name') {
    if (!canBindLineUserName_(targetRecord)) {
      return [buildLineTextMessage_('解除綁定只支援 1 對 1 私訊。請改在個人聊天室中操作。')];
    }
    const updatedTarget = applyLineTargetBoundUserName_(targetRecord, '');
    return [buildLineTextMessage_(`已解除姓名綁定。${updatedTarget.bound_user_name ? '' : '\n如需重新設定，請輸入「我是 王小明」。'}`)];
  }
  if (command.type === 'notification_status') {
    return buildLineNotificationMessages_(targetRecord, '這是目前聊天室 / 對話的通知設定。');
  }
  if (command.type === 'notification_update') {
    const updatedTarget = applyLineTargetNotificationMode_(targetRecord, command.mode);
    const labelMap = {
      enable: '已啟用通知。',
      disable: '已停止此聊天室 / 對話的自動通知。',
      all: '已改為接收全部通知。',
      tomorrow_only: '已改為只接收明天名單。',
      marquee_only: '已改為只接收公告。'
    };
    return buildLineNotificationMessages_(updatedTarget, labelMap[command.mode] || '通知設定已更新。');
  }
  return [];
}

function buildLineWelcomeMessages_(event, targetRecord) {
  const eventType = String(event && event.type || '');
  const greeting = eventType === 'join'
    ? '已加入此群組 / 聊天室。'
    : '感謝加入 AV 放送 Bot。';
  return [
    buildLineTextMessage_(`${greeting}\n輸入「選單」可查看常用指令；也可以直接輸入「明天名單」、「公告」或「通知設定」。`, [
      buildLineQuickReplyItem_('選單', '選單'),
      buildLineQuickReplyItem_('明天名單', '明天名單'),
      buildLineQuickReplyItem_('公告', '公告'),
      buildLineQuickReplyItem_('通知設定', '通知設定')
    ]),
    buildLineMenuTemplateMessage_()
  ];
}

function lineWebhook_(payload) {
  const events = (payload && Array.isArray(payload.events)) ? payload.events : [];
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_LINE_WEBHOOK_LAST_EVENT_COUNT, String(events.length));
  props.deleteProperty(PROP_LINE_WEBHOOK_LAST_ERROR);
  props.deleteProperty(PROP_LINE_WEBHOOK_LAST_REPLY_ERROR);
  if (!events.length) {
    return { ok: true, received_events: 0, processed_events: 0, recorded: 0, replied_events: 0, data: getLineTargets_() };
  }

  const current = getLineTargets_();
  const map = {};
  current.forEach(item => {
    if (!item || !item.target_id) return;
    map[String(item.target_id)] = normalizeLineTargetRecord_(item);
  });

  let recorded = 0;
  let processedEvents = 0;
  let repliedEvents = 0;
  let skippedEvents = 0;
  const nowIso = new Date().toISOString();

  events.forEach(ev => {
    try {
      if (!shouldProcessLineEvent_(ev)) {
        skippedEvents += 1;
        return;
      }
      processedEvents += 1;
      const targetId = getLineEventTargetId_(ev);
      if (targetId) {
        props.setProperty(PROP_LINE_WEBHOOK_LAST_SOURCE, targetId);
        const existing = map[targetId] || null;
        const mergedTarget = buildLineTargetFromEvent_(ev, existing, nowIso);
        if (mergedTarget) {
          map[targetId] = mergedTarget;
          if (!existing || !existing.target_id) recorded += 1;
        }
      }

      const eventType = String(ev && ev.type || '').trim();
      let replyMessages = [];
      if (eventType === 'follow' || eventType === 'join') {
        replyMessages = buildLineWelcomeMessages_(ev, map[targetId] || {});
      } else if (eventType === 'message') {
        const message = ev && ev.message || {};
        if (String(message.type || '') === 'text') {
          const command = parseLineCommand_(message.text);
          if (command) {
            if (command.type === 'notification_update' && targetId && map[targetId]) {
              map[targetId] = applyLineTargetNotificationMode_(map[targetId], command.mode);
            } else if (command.type === 'bind_user_name' && targetId && map[targetId]) {
              map[targetId] = applyLineTargetBoundUserName_(map[targetId], command.userName);
            } else if (command.type === 'unbind_user_name' && targetId && map[targetId]) {
              map[targetId] = applyLineTargetBoundUserName_(map[targetId], '');
            }
            replyMessages = buildLineCommandReplyMessages_(command, map[targetId] || {});
          }
        }
      } else if (eventType === 'postback') {
        const command = parseLineCommand_(String(ev && ev.postback && ev.postback.data || '').replace(/^cmd:/, ''));
        if (command) {
          if (command.type === 'notification_update' && targetId && map[targetId]) {
            map[targetId] = applyLineTargetNotificationMode_(map[targetId], command.mode);
          } else if (command.type === 'bind_user_name' && targetId && map[targetId]) {
            map[targetId] = applyLineTargetBoundUserName_(map[targetId], command.userName);
          } else if (command.type === 'unbind_user_name' && targetId && map[targetId]) {
            map[targetId] = applyLineTargetBoundUserName_(map[targetId], '');
          }
          replyMessages = buildLineCommandReplyMessages_(command, map[targetId] || {});
        }
      }

      const replyToken = String(ev && ev.replyToken || '').trim();
      if (replyToken && replyMessages.length) {
        const replyResult = sendLineReplyMessages_(replyToken, replyMessages, {
          batch_id: Utilities.getUuid(),
          trigger_source: 'line_webhook',
          message_type: 'webhook_reply',
          function_name: 'lineWebhook_',
          line_to_input: targetId,
          resolved_target_count: targetId ? 1 : 0,
          target_id: targetId
        });
        if (replyResult && replyResult.sent) {
          repliedEvents += 1;
          props.setProperty(PROP_LINE_WEBHOOK_LAST_REPLY_STATUS, String(replyResult.status || 200));
          props.deleteProperty(PROP_LINE_WEBHOOK_LAST_REPLY_ERROR);
        } else if (replyResult && replyResult.error) {
          props.setProperty(PROP_LINE_WEBHOOK_LAST_REPLY_ERROR, String(replyResult.error));
        }
      }
    } catch (eventErr) {
      props.setProperty(PROP_LINE_WEBHOOK_LAST_ERROR, String(eventErr && eventErr.message ? eventErr.message : eventErr));
      console.error(`lineWebhook event 處理失敗: ${eventErr.message}\n${eventErr.stack}`);
    }
  });

  const merged = Object.keys(map).map(k => normalizeLineTargetRecord_(map[k])).sort((a, b) => {
    const aTs = new Date(a.last_seen_at || 0).getTime();
    const bTs = new Date(b.last_seen_at || 0).getTime();
    return bTs - aTs;
  });
  setLineTargets_(merged);
  return {
    ok: true,
    received_events: events.length,
    processed_events: processedEvents,
    skipped_events: skippedEvents,
    recorded: recorded,
    replied_events: repliedEvents,
    total: merged.length,
    data: merged
  };
}

function maskToken_(token) {
  const t = String(token || '');
  if (t.length <= 12) return t ? '***' : '';
  return `${t.slice(0, 6)}...${t.slice(-6)}`;
}

function normalizeLineToken_(rawToken) {
  return String(rawToken || '')
    .replace(/^Bearer\s+/i, '')
    .replace(/\s+/g, '')
    .trim();
}

function setLineToken(rawToken) {
  const token = normalizeLineToken_(rawToken);
  if (!token) {
    return { status: 'error', error: 'line_channel_access_token 不可為空。' };
  }
  if (token.includes('請在此貼上')) {
    return { status: 'error', error: 'line_channel_access_token 格式不正確。' };
  }
  PropertiesService.getScriptProperties().setProperty(LINE_TOKEN_PROPERTY_KEY, token);
  return {
    status: 'success',
    token_preview: maskToken_(token),
    token_length: token.length
  };
}

function debugLineAuth(linTo) {
  const token = resolveLineToken();
  const target = resolveLineTo(linTo);
  if (!token) {
    return { status: 'error', error: 'LINE token 未設定', token_preview: '', token_length: 0, lin_to: target || '' };
  }
  const verifyUrl = 'https://api.line.me/v2/bot/info';
  try {
    const resp = UrlFetchApp.fetch(verifyUrl, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const code = Number(resp.getResponseCode() || 0);
    const body = String(resp.getContentText() || '');
    if (code >= 200 && code < 300) {
      return {
        status: 'success',
        message: 'LINE token 驗證成功 (bot/info)',
        token_preview: maskToken_(token),
        token_length: token.length,
        lin_to: target || '',
        verify_code: code,
        verify_body: body
      };
    }
    return {
      status: 'error',
      error: `LINE token 驗證失敗 (${code})`,
      token_preview: maskToken_(token),
      token_length: token.length,
      lin_to: target || '',
      verify_code: code,
      verify_body: body
    };
  } catch (e) {
    return {
      status: 'error',
      error: 'LINE token 驗證例外: ' + String(e),
      token_preview: maskToken_(token),
      token_length: token.length,
      lin_to: target || ''
    };
  }
}

function clearLineTo() {
  PropertiesService.getScriptProperties().deleteProperty(LINE_TO_PROPERTY_KEY);
  return { status: 'success', lin_to: '' };
}

function removeLineTarget(targetId) {
  const normalizedId = String(targetId || '').trim();
  if (!normalizedId) {
    return { status: 'error', error: 'targetId is required' };
  }

  const current = getLineTargets_();
  const next = current.filter(item => String(item && item.target_id || '') !== normalizedId);
  if (next.length === current.length) {
    return { status: 'error', error: `LINE target not found: ${normalizedId}` };
  }

  setLineTargets_(next);
  const savedLineTo = String(PropertiesService.getScriptProperties().getProperty(LINE_TO_PROPERTY_KEY) || '').trim();
  if (savedLineTo === normalizedId) {
    PropertiesService.getScriptProperties().deleteProperty(LINE_TO_PROPERTY_KEY);
  }

  return {
    status: 'success',
    removed_target_id: normalizedId,
    data: next
  };
}

function updateLineTargetAlias(targetId, rawAlias) {
  const normalizedId = String(targetId || '').trim();
  if (!normalizedId) {
    return { status: 'error', error: 'targetId is required' };
  }

  const current = getLineTargets_();
  let updated = false;
  const next = current.map(item => {
    if (String(item && item.target_id || '') !== normalizedId) {
      return item;
    }
    updated = true;
    return Object.assign({}, item, {
      alias: normalizeLineTargetAlias_(rawAlias)
    });
  });

  if (!updated) {
    return { status: 'error', error: `LINE target not found: ${normalizedId}` };
  }

  setLineTargets_(next);
  const saved = next.find(item => String(item && item.target_id || '') === normalizedId) || {};
  return {
    status: 'success',
    target_id: normalizedId,
    alias: String(saved.alias || ''),
    message: `Alias updated for ${normalizedId}.`,
    data: next
  };
}

function summarizeLineMessagesForLog_(messages) {
  const list = Array.isArray(messages) ? messages : [];
  const parts = list.map(message => {
    if (!message || typeof message !== 'object') return '';
    if (message.type === 'text') return String(message.text || '');
    if (message.type === 'template') return `[template] ${String(message.altText || '')}`;
    if (message.type === 'flex') return `[flex] ${String(message.altText || '')}`;
    return `[${String(message.type || 'message')}]`;
  }).filter(Boolean);
  return parts.join('\n---\n');
}

function sendLineReplyMessages_(replyToken, messages, logContext) {
  const token = resolveLineToken();
  const normalizedReplyToken = String(replyToken || '').trim();
  const messageList = (Array.isArray(messages) ? messages : []).filter(Boolean).slice(0, 5);
  const context = Object.assign({
    batch_id: Utilities.getUuid(),
    trigger_source: 'line_webhook',
    message_type: 'webhook_reply',
    function_name: 'sendLineReplyMessages_',
    line_to_input: String(logContext && logContext.target_id || ''),
    resolved_target_count: 1,
    target_id: String(logContext && logContext.target_id || '')
  }, logContext || {});
  const logMessage = summarizeLineMessagesForLog_(messageList);

  if (!token || token.includes('請在此貼上')) {
    const errorText = 'LINE_CHANNEL_ACCESS_TOKEN 未設定。';
    logLineSendFailure_(context, logMessage, errorText);
    return { sent: false, status: 0, error: errorText };
  }
  if (!normalizedReplyToken) {
    const errorText = 'replyToken 不存在。';
    logLineSendFailure_(context, logMessage, errorText);
    return { sent: false, status: 0, error: errorText };
  }
  if (!messageList.length) {
    return { sent: false, status: 0, error: '無可回覆訊息。' };
  }

  let response;
  try {
    response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ replyToken: normalizedReplyToken, messages: messageList }),
      muteHttpExceptions: true
    });
  } catch (err) {
    const errorText = 'LINE reply request failed: ' + String(err);
    logLineSendFailure_(context, logMessage, errorText);
    return { sent: false, status: 0, error: errorText };
  }

  const status = Number(response.getResponseCode() || 0);
  const bodyText = String(response.getContentText() || '');
  const sent = status >= 200 && status < 300;
  const result = {
    target: String(context.target_id || ''),
    sent: sent,
    status: status,
    response_body: bodyText,
    error: sent ? '' : `LINE API 回應 ${status}: ${bodyText}`
  };
  appendLineSendLogs_(buildLineSendLogRows_(context, logMessage, [result]));
  return result;
}

function sendLinePushSingle_(token, target, text) {
  const reqBody = {
    to: target,
    messages: [{ type: 'text', text: text }]
  };
  let response;
  try {
    response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(reqBody),
      muteHttpExceptions: true
    });
  } catch (err) {
    return {
      target: target,
      sent: false,
      status: 0,
      error: 'LINE API request failed: ' + String(err)
    };
  }

  const status = Number(response.getResponseCode() || 0);
  const bodyText = String(response.getContentText() || '');
  const sent = status >= 200 && status < 300;
  return {
    target: target,
    sent: sent,
    status: status,
    response_body: bodyText,
    error: sent ? '' : `LINE API 回應 ${status}: ${bodyText}`
  };
}

function sendLineMessage(message, linTo, logContext) {
  const context = Object.assign({
    batch_id: Utilities.getUuid(),
    trigger_source: '',
    message_type: '',
    function_name: 'sendLineMessage',
    line_to_input: String(linTo || ''),
    respect_preferences: false
  }, logContext || {});
  const channelAccessToken = resolveLineToken();
  if (!channelAccessToken || channelAccessToken.includes('請在此貼上')) {
    const msg = 'LINE_CHANNEL_ACCESS_TOKEN 未設定。';
    console.error(msg);
    logLineSendFailure_(Object.assign({}, context, { resolved_target_count: 0 }), message, msg);
    return { status: 'error', error: msg };
  }
  const targets = resolveLineTargetsForSend_(linTo, {
    message_type: context.message_type,
    respect_preferences: context.respect_preferences
  });
  if (!targets.length || targets.some(t => t.includes('請在此貼上'))) {
    const hasDirectTarget = parseLineTargets_(linTo).length > 0;
    const hasConfiguredTarget = parseLineTargets_(String(PropertiesService.getScriptProperties().getProperty(LINE_TO_PROPERTY_KEY) || '').trim()).length > 0;
    const hasRecordedTarget = getRecordedLineTargets_().length > 0;
    const msg = (!hasDirectTarget && !hasConfiguredTarget && hasRecordedTarget && context.respect_preferences)
      ? '目前沒有符合此通知類型的訂閱對象。'
      : 'lin_to (LINE User ID / Group ID) 未設定。';
    console.error(msg);
    logLineSendFailure_(Object.assign({}, context, { resolved_target_count: 0 }), message, msg);
    return { status: 'error', error: msg };
  }
  const invalidTargets = findInvalidLineTargetIds_(targets);
  if (invalidTargets.length) {
    const msg = `line_to 格式不正確：${invalidTargets.join(', ')}`;
    console.error(msg);
    logLineSendFailure_(Object.assign({}, context, {
      resolved_target_count: targets.length
    }), message, msg);
    return {
      status: 'error',
      error: msg,
      invalid_targets: invalidTargets,
      targets: targets
    };
  }
  const results = targets.map(target => sendLinePushSingle_(channelAccessToken, target, message));
  appendLineSendLogs_(buildLineSendLogRows_(Object.assign({}, context, {
    resolved_target_count: targets.length
  }), message, results));
  const successCount = results.filter(r => r && r.sent).length;
  if (successCount > 0) {
    return { status: 'success', target_count: targets.length, success_count: successCount, targets: targets, results: results };
  }
  const firstError = results[0] && results[0].error ? results[0].error : '發送失敗';
  return { status: 'error', error: firstError, target_count: targets.length, targets: targets, results: results };
}

function dailyNotifyTomorrow_MessageAPI(linTo, logContext) { // 函式名稱維持不變以相容前端
  try {
    const res = getSignupsAsTextForTomorrow();
    let msg = '';
    if (!res || res.status === 'nodata') {
      msg = '\n（提醒）明日無報名資料。';
    } else if (res.status !== 'success' || !res.text) {
      msg = '\n取得明日報名文字失敗。';
    } else {
      msg = '\n' + res.text;
    }
    return sendLineMessage(msg, linTo, Object.assign({
      trigger_source: 'manual_or_api',
      message_type: 'tomorrow_signup',
      function_name: 'dailyNotifyTomorrow_MessageAPI',
      respect_preferences: !String(linTo || '').trim()
    }, logContext || {}));
  } catch (e) {
    console.error('[dailyNotify] 例外: ' + e.toString());
    logLineSendFailure_(Object.assign({
      batch_id: Utilities.getUuid(),
      trigger_source: 'manual_or_api',
      message_type: 'tomorrow_signup',
      function_name: 'dailyNotifyTomorrow_MessageAPI',
      line_to_input: String(linTo || ''),
      resolved_target_count: 0
    }, logContext || {}), '', e.toString());
    return { status: 'error', error: e.toString() };
  }
}

function dailyNotifyTomorrowAndMarqueeSeparate_MessageAPI(linTo) {
  try {
    const executionBatchId = Utilities.getUuid();
    const tomorrowResult = dailyNotifyTomorrow_MessageAPI(linTo, {
      batch_id: executionBatchId + '-tomorrow',
      trigger_source: 'daily_auto_trigger',
      message_type: 'tomorrow_signup',
      function_name: 'dailyNotifyTomorrow_MessageAPI'
    });
    const marqueeResult = sendMarqueeAnnouncementsToLine_MessageAPI(linTo, {
      batch_id: executionBatchId + '-marquee',
      trigger_source: 'daily_auto_trigger',
      message_type: 'marquee_announcement',
      function_name: 'sendMarqueeAnnouncementsToLine_MessageAPI'
    });
    const tomorrowOk = tomorrowResult && tomorrowResult.status === 'success';
    const marqueeOk = marqueeResult && marqueeResult.status === 'success';

    if (tomorrowOk && marqueeOk) {
      return {
        status: 'success',
        message: '已分開發送 2 則通知（明天名單、跑馬燈公告）。',
        results: {
          tomorrow: tomorrowResult,
          marquee: marqueeResult
        }
      };
    }

    return {
      status: 'error',
      error: '部分通知發送失敗。',
      results: {
        tomorrow: tomorrowResult,
        marquee: marqueeResult
      }
    };
  } catch (e) {
    console.error('[dailyNotifySeparate] 例外: ' + e.toString());
    return { status: 'error', error: e.toString() };
  }
}

function sendMarqueeAnnouncementsToLine_MessageAPI(linTo, logContext) {
  try {
    const announcements = getNewsAnnouncements();
    const list = Array.isArray(announcements) ? announcements : [];
    const scriptTimeZone = Session.getScriptTimeZone() || 'Asia/Taipei';
    const generatedAt = Utilities.formatDate(new Date(), scriptTimeZone, 'yyyy/MM/dd HH:mm');
    let messageBody = '';

    if (!list.length) {
      messageBody = '目前沒有最新公告。';
    } else {
      messageBody = list.map((item, index) => `${index + 1}. ${item}`).join('\n');
    }

    let fullMessage = `【跑馬燈公告】\n更新時間：${generatedAt}\n\n${messageBody}`;
    if (fullMessage.length > 4900) {
      fullMessage = fullMessage.slice(0, 4868) + '\n...(內容過長已截斷)';
    }

    const sendResult = sendLineMessage(fullMessage, linTo, Object.assign({
      trigger_source: 'manual_or_api',
      message_type: 'marquee_announcement',
      function_name: 'sendMarqueeAnnouncementsToLine_MessageAPI',
      respect_preferences: !String(linTo || '').trim()
    }, logContext || {}));
    sendResult.announcement_count = list.length;
    return sendResult;
  } catch (e) {
    console.error('[sendMarqueeAnnouncementsToLine] 失敗: ' + e.toString());
    logLineSendFailure_(Object.assign({
      batch_id: Utilities.getUuid(),
      trigger_source: 'manual_or_api',
      message_type: 'marquee_announcement',
      function_name: 'sendMarqueeAnnouncementsToLine_MessageAPI',
      line_to_input: String(linTo || ''),
      resolved_target_count: 0
    }, logContext || {}), '', e.toString());
    return { status: 'error', error: e.toString(), announcement_count: 0 };
  }
}

function ensureLineSendLogSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(LINE_SEND_LOGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LINE_SEND_LOGS_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'BatchId',
      'TriggerSource',
      'MessageType',
      'FunctionName',
      'LineToInput',
      'ResolvedTargetCount',
      'Target',
      'TargetAlias',
      'ChunkIndex',
      'ChunkCount',
      'MessageLength',
      'MessagePreview',
      'MessageBody',
      'Sent',
      'HttpStatus',
      'Error',
      'ResponseBody'
    ]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getLineTargetAliasMap_() {
  const aliasMap = {};
  getLineTargets_().forEach(item => {
    const targetId = String(item && item.target_id || '').trim();
    if (!targetId) return;
    aliasMap[targetId] = String(item && item.alias || '');
  });
  return aliasMap;
}

function buildLineSendLogRows_(context, message, results) {
  const list = Array.isArray(results) ? results : [];
  if (!list.length) return [];

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  const aliasMap = getLineTargetAliasMap_();
  const text = String(message || '');
  const preview = text.length > 160 ? text.slice(0, 160) + '...' : text;
  const batchId = String(context && context.batch_id || Utilities.getUuid());
  const resolvedTargetCount = Number(context && context.resolved_target_count || list.length || 0);
  const chunkIndex = Number(context && context.chunk_index || 1);
  const chunkCount = Number(context && context.chunk_count || 1);

  return list.map(result => {
    const target = String(result && result.target || '');
    return [
      timestamp,
      batchId,
      String(context && context.trigger_source || ''),
      String(context && context.message_type || ''),
      String(context && context.function_name || ''),
      String(context && context.line_to_input || ''),
      resolvedTargetCount,
      target,
      String(aliasMap[target] || ''),
      chunkIndex,
      chunkCount,
      text.length,
      preview,
      text,
      result && result.sent ? 'TRUE' : 'FALSE',
      Number(result && result.status || 0),
      String(result && result.error || ''),
      String(result && result.response_body || '')
    ];
  });
}

function appendLineSendLogs_(rows) {
  const entries = Array.isArray(rows) ? rows : [];
  if (!entries.length) return;

  try {
    const sheet = ensureLineSendLogSheet_();
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, entries.length, entries[0].length).setValues(entries);
    SpreadsheetApp.flush();
  } catch (err) {
    console.error(`Failed to write LINE send logs: ${err.message}\n${err.stack}`);
  }
}

function logLineSendFailure_(context, message, errorText) {
  appendLineSendLogs_(buildLineSendLogRows_(context, message, [{
    target: '',
    sent: false,
    status: 0,
    error: String(errorText || ''),
    response_body: ''
  }]));
}

function splitLineMessageChunks_(message, maxLength) {
  const text = String(message || '').replace(/\r\n/g, '\n');
  const limit = Math.max(1, Number(maxLength || 0));
  if (!text) return [];

  const chunks = [];
  let remaining = text;
  while (remaining.length > limit) {
    let splitAt = remaining.lastIndexOf('\n', limit);
    if (splitAt <= 0 || splitAt < Math.floor(limit * 0.6)) splitAt = limit;
    chunks.push(remaining.slice(0, splitAt));
    remaining = remaining.slice(splitAt).replace(/^\n+/, '');
  }
  if (remaining) chunks.push(remaining);
  return chunks;
}

function sendCustomLineMessage_MessageAPI(message, linTo) {
  const batchId = Utilities.getUuid();
  try {
    const normalizedMessage = String(message || '').replace(/\r\n/g, '\n');
    if (!normalizedMessage.trim()) {
      logLineSendFailure_({
        batch_id: batchId,
        trigger_source: 'manual_custom_message',
        message_type: 'custom_message',
        function_name: 'sendCustomLineMessage_MessageAPI',
        line_to_input: String(linTo || ''),
        resolved_target_count: 0
      }, normalizedMessage, '訊息內容不可為空。');
      return { status: 'error', error: '訊息內容不可為空。' };
    }

    const channelAccessToken = resolveLineToken();
    if (!channelAccessToken || channelAccessToken.includes('請在此貼上')) {
      logLineSendFailure_({
        batch_id: batchId,
        trigger_source: 'manual_custom_message',
        message_type: 'custom_message',
        function_name: 'sendCustomLineMessage_MessageAPI',
        line_to_input: String(linTo || ''),
        resolved_target_count: 0
      }, normalizedMessage, 'LINE_CHANNEL_ACCESS_TOKEN 未設定。');
      return { status: 'error', error: 'LINE_CHANNEL_ACCESS_TOKEN 未設定。' };
    }

    const targets = resolveLineTargetsForSend_(linTo);
    if (!targets.length || targets.some(t => t.includes('請在此貼上'))) {
      logLineSendFailure_({
        batch_id: batchId,
        trigger_source: 'manual_custom_message',
        message_type: 'custom_message',
        function_name: 'sendCustomLineMessage_MessageAPI',
        line_to_input: String(linTo || ''),
        resolved_target_count: 0
      }, normalizedMessage, 'lin_to (LINE User ID / Group ID) 未設定。');
      return { status: 'error', error: 'lin_to (LINE User ID / Group ID) 未設定。' };
    }
    const invalidTargets = findInvalidLineTargetIds_(targets);
    if (invalidTargets.length) {
      const errorText = `line_to 格式不正確：${invalidTargets.join(', ')}`;
      logLineSendFailure_({
        batch_id: batchId,
        trigger_source: 'manual_custom_message',
        message_type: 'custom_message',
        function_name: 'sendCustomLineMessage_MessageAPI',
        line_to_input: String(linTo || ''),
        resolved_target_count: targets.length
      }, normalizedMessage, errorText);
      return { status: 'error', error: errorText, invalid_targets: invalidTargets, targets: targets };
    }

    const chunks = splitLineMessageChunks_(normalizedMessage, 4900);
    const results = [];
    chunks.forEach((chunk, chunkIndex) => {
      const chunkResults = targets.map(target => {
        const sendResult = sendLinePushSingle_(channelAccessToken, target, chunk);
        sendResult.chunk_index = chunkIndex + 1;
        return sendResult;
      });
      appendLineSendLogs_(buildLineSendLogRows_({
        batch_id: batchId,
        trigger_source: 'manual_custom_message',
        message_type: 'custom_message',
        function_name: 'sendCustomLineMessage_MessageAPI',
        line_to_input: String(linTo || ''),
        resolved_target_count: targets.length,
        chunk_index: chunkIndex + 1,
        chunk_count: chunks.length
      }, chunk, chunkResults));
      results.push.apply(results, chunkResults);
    });

    const successCount = results.filter(r => r && r.sent).length;
    if (successCount > 0) {
      return {
        status: 'success',
        chunk_count: chunks.length,
        target_count: targets.length,
        success_count: successCount,
        targets: targets,
        results: results
      };
    }

    const firstError = results[0] && results[0].error ? results[0].error : '發送失敗';
    return {
      status: 'error',
      error: firstError,
      chunk_count: chunks.length,
      target_count: targets.length,
      targets: targets,
      results: results
    };
  } catch (e) {
    console.error('[sendCustomLineMessage] 失敗: ' + e.toString());
    logLineSendFailure_({
      batch_id: batchId,
      trigger_source: 'manual_custom_message',
      message_type: 'custom_message',
      function_name: 'sendCustomLineMessage_MessageAPI',
      line_to_input: String(linTo || ''),
      resolved_target_count: 0
    }, String(message || ''), e.toString());
    return { status: 'error', error: e.toString() };
  }
}

function setupDaily20Trigger_MessageAPI(linTo) {
  const funcName = 'dailyNotifyTomorrowAndMarqueeSeparate_MessageAPI';
  try {
    const targets = resolveLineTargetsForSend_(linTo);
    if (!targets.length || targets.some(t => t.includes('請在此貼上'))) {
      return { status: 'error', error: 'lin_to (LINE User ID / Group ID) 未設定。' };
    }
    const invalidTargets = findInvalidLineTargetIds_(targets);
    if (invalidTargets.length) {
      return {
        status: 'error',
        error: `line_to 格式不正確：${invalidTargets.join(', ')}`,
        invalid_targets: invalidTargets,
        targets: targets
      };
    }
    ensureLineSendLogSheet_();
    const triggers = ScriptApp.getProjectTriggers() || [];
    const managedHandlerNames = [
      'dailyNotifyTomorrow_MessageAPI',
      'dailyNotifyTomorrowAndMarqueeSeparate_MessageAPI'
    ];
    triggers.forEach(t => {
      const handlerName = t.getHandlerFunction ? t.getHandlerFunction() : '';
      if (managedHandlerNames.indexOf(handlerName) >= 0) {
        ScriptApp.deleteTrigger(t);
      }
    });
    ScriptApp.newTrigger(funcName)
      .timeBased()
      .atHour(20)
      .everyDays(1)
      .inTimezone('Asia/Taipei')
      .create();
    console.log('[setupDailyTrigger] 已建立每日 20:00 排程 (Messaging API)。');
    return { status: 'success', lin_to: resolveLineTo(linTo), target_count: targets.length };
  } catch (e) {
    console.error('[setupDailyTrigger] 失敗: ' + e.toString());
    return { status: 'error', error: e.toString() };
  }
}

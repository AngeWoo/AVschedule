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
      case 'getLineWebhookStatus': result = getLineWebhookStatus_(); break;
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
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_LINE_WEBHOOK_LAST_AT, new Date().toISOString());
    const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '';
    if (rawBody) {
      let body = null;
      try { body = JSON.parse(rawBody); } catch (parseErr) { body = null; }
      if (body && Array.isArray(body.events)) {
        const result = lineWebhook_(body);
        return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  } catch (err) {
    PropertiesService.getScriptProperties().setProperty(PROP_LINE_WEBHOOK_LAST_ERROR, String(err && err.message ? err.message : err));
    console.error(`doPost webhook 失敗: ${err.message}\n${err.stack}`);
  }
  return doGet(e);
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
 * 從 CSV 網址獲取 Events 和 Signups 的主要資料。
 */
function getMasterData() {
  try {
    const eventsData = fetchAndParseCsvWithCache(EVENTS_CSV_URL, 'events_data');
    const signupsData = fetchAndParseCsvWithCache(SIGNUPS_CSV_URL, 'signups_data');

    console.log(`getMasterData (CSV): 成功讀取資料。`);
    console.log(`getMasterData: Events CSV 讀取到 ${eventsData.length} 行 (含標頭)。`);
    console.log(`getMasterData: Signups CSV 讀取到 ${signupsData.length} 行 (含標頭)。`);

    // 移除標頭，回傳純資料部分
    const eventsDataWithoutHeader = eventsData.slice(1);
    const signupsDataWithoutHeader = signupsData.slice(1);

    const scriptTimeZone = Session.getScriptTimeZone();
    const eventsMap = new Map();

    eventsDataWithoutHeader.forEach(row => {
      // CSV Columns: ID,Title,Date,StartTime,EndTime,MaxAttendees,Positions,Description
      const eventId = row[0];
      if (eventId && row[2]) { // 檢查 ID 和 Date 存在
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
    console.error(`getMasterData (CSV) 失敗: ${err.message}`, err.stack);
    throw new Error(`從 CSV 讀取資料時發生錯誤: ${err.message}`);
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

function getRecordedLineTargetIds_() {
  return getLineTargets_()
    .map(item => String(item && item.target_id ? item.target_id : '').trim())
    .filter(Boolean);
}

function resolveLineTargetsForSend_(linTo) {
  const targetLineToRaw = resolveLineTo(linTo);
  const configuredTargets = parseLineTargets_(targetLineToRaw);
  const recordedTargets = getRecordedLineTargetIds_();
  const merged = configuredTargets.concat(recordedTargets);
  const deduped = [];
  const seen = {};
  merged.forEach(t => {
    const key = String(t || '').trim();
    if (!key || seen[key]) return;
    seen[key] = true;
    deduped.push(key);
  });
  return deduped;
}

function getLineTargets_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_LINE_TARGETS);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    console.error('解析 LINE_TARGETS 失敗: ' + e.toString());
    return [];
  }
}

function setLineTargets_(items) {
  const arr = Array.isArray(items) ? items : [];
  PropertiesService.getScriptProperties().setProperty(PROP_LINE_TARGETS, JSON.stringify(arr));
}

function getLineWebhookStatus_() {
  const props = PropertiesService.getScriptProperties();
  const targets = getLineTargets_();
  return {
    last_at: String(props.getProperty(PROP_LINE_WEBHOOK_LAST_AT) || ''),
    last_event_count: Number(props.getProperty(PROP_LINE_WEBHOOK_LAST_EVENT_COUNT) || 0),
    last_source: String(props.getProperty(PROP_LINE_WEBHOOK_LAST_SOURCE) || ''),
    last_error: String(props.getProperty(PROP_LINE_WEBHOOK_LAST_ERROR) || ''),
    targets_count: targets.length
  };
}

function inferLineToType_(targetId) {
  const s = String(targetId || '');
  if (s.startsWith('U')) return 'user';
  if (s.startsWith('C')) return 'group';
  if (s.startsWith('R')) return 'room';
  return 'unknown';
}

function lineWebhook_(payload) {
  const events = (payload && Array.isArray(payload.events)) ? payload.events : [];
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_LINE_WEBHOOK_LAST_EVENT_COUNT, String(events.length));
  props.deleteProperty(PROP_LINE_WEBHOOK_LAST_ERROR);
  if (!events.length) {
    return { ok: true, received_events: 0, recorded: 0, data: getLineTargets_() };
  }

  const current = getLineTargets_();
  const map = {};
  current.forEach(item => {
    if (!item || !item.target_id) return;
    map[String(item.target_id)] = item;
  });

  let recorded = 0;
  const nowIso = new Date().toISOString();
  events.forEach(ev => {
    const source = (ev && ev.source) || {};
    const targetId = String(source.groupId || source.roomId || source.userId || '').trim();
    if (!targetId) return;
    props.setProperty(PROP_LINE_WEBHOOK_LAST_SOURCE, targetId);

    const sourceType = String(source.type || inferLineToType_(targetId));
    const existing = map[targetId] || {};
    map[targetId] = {
      target_id: targetId,
      source_type: sourceType,
      user_id: String(source.userId || ''),
      group_id: String(source.groupId || ''),
      room_id: String(source.roomId || ''),
      last_seen_at: nowIso
    };
    if (!existing.target_id) recorded += 1;
  });

  const merged = Object.keys(map).map(k => map[k]).sort((a, b) => {
    const aTs = new Date(a.last_seen_at || 0).getTime();
    const bTs = new Date(b.last_seen_at || 0).getTime();
    return bTs - aTs;
  });
  setLineTargets_(merged);
  return { ok: true, received_events: events.length, recorded: recorded, total: merged.length, data: merged };
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

function sendLineMessage(message, linTo) {
  const channelAccessToken = resolveLineToken();
  if (!channelAccessToken || channelAccessToken.includes('請在此貼上')) {
    const msg = 'LINE_CHANNEL_ACCESS_TOKEN 未設定。';
    console.error(msg);
    return { status: 'error', error: msg };
  }
  const targets = resolveLineTargetsForSend_(linTo);
  if (!targets.length || targets.some(t => t.includes('請在此貼上'))) {
    const msg = 'lin_to (LINE User ID / Group ID) 未設定。';
    console.error(msg);
    return { status: 'error', error: msg };
  }
  const results = targets.map(target => sendLinePushSingle_(channelAccessToken, target, message));
  const successCount = results.filter(r => r && r.sent).length;
  if (successCount > 0) {
    return { status: 'success', target_count: targets.length, success_count: successCount, results: results };
  }
  const firstError = results[0] && results[0].error ? results[0].error : '發送失敗';
  return { status: 'error', error: firstError, target_count: targets.length, results: results };
}

function dailyNotifyTomorrow_MessageAPI(linTo) { // 函式名稱維持不變以相容前端
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
    return sendLineMessage(msg, linTo);
  } catch (e) {
    console.error('[dailyNotify] 例外: ' + e.toString());
    return { status: 'error', error: e.toString() };
  }
}

function setupDaily20Trigger_MessageAPI(linTo) {
  const funcName = 'dailyNotifyTomorrow_MessageAPI';
  try {
    const targets = resolveLineTargetsForSend_(linTo);
    if (!targets.length || targets.some(t => t.includes('請在此貼上'))) {
      return { status: 'error', error: 'lin_to (LINE User ID / Group ID) 未設定。' };
    }
    const triggers = ScriptApp.getProjectTriggers() || [];
    triggers.forEach(t => {
      if (t.getHandlerFunction && t.getHandlerFunction() === funcName) {
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

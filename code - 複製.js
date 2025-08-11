// ====================================================================
//  檔案：code.gs (AV放送立願報名行事曆 - 後端 API)
//  版本：6.27 (新增操作日誌寫入 Logs 分頁)
// ====================================================================

const BACKEND_VERSION = "0716"; // *** 後端版本號 ***

// --- 已填入您提供的 CSV 網址 ---
const EVENTS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=795926947&single=true&output=csv";
const SIGNUPS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=767193030&single=true&output=csv";
// ---------------------------------------------------

// --- 全域設定 ---
const SHEET_ID = '1gBIlhEKPQHBvslY29veTdMEJeg2eVcaJx_7A-8cTWIM';
const EVENTS_SHEET_NAME = 'Events';
const SIGNUPS_SHEET_NAME = 'Signups';
const LOGS_SHEET_NAME = 'Logs'; // 新增 Logs 分頁名稱

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

    console.log(`doGet: 接收到函式呼叫 - ${functionName}，參數: ${JSON.stringify(params)}`); // Debug: Log incoming call

    switch (functionName) {
      case 'getEventsAndSignups': result = getEventsAndSignups(); break;
      case 'getUnifiedSignups': result = getUnifiedSignups(params.searchText, params.startDateStr, params.endDateStr); break;
      case 'getStatsData': result = getStatsData(params.startDateStr, params.endDateStr); break;
      case 'addSignup': result = addSignup(params.eventId, params.userName, params.position); break;
      case 'addBackupSignup': result = addBackupSignup(params.eventId, params.userName, params.position); break;
      case 'removeSignup': result = removeSignup(params.eventId, params.userName); break;
      case 'modifySignup': result = modifySignup(params.signupId, params.newPosition); break;
      case 'createTempSheetAndExport': result = createTempSheetAndExport(params.startDateStr, params.endDateStr); break;
      case 'getSignupsAsTextForToday': result = getSignupsAsTextForToday(); break;
      case 'getSignupsAsTextForTomorrow': result = getSignupsAsTextForTomorrow(); break;
      case 'getSignupsAsTextForDateRange': result = getSignupsAsText(params.startDateStr, params.endDateStr); break;
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
  return doGet(e);
}

// -------------------- 輔助函式 --------------------
function padTime(timeStr) {
  let trimmedTime = String(timeStr || '').trim();
  if (trimmedTime.includes('GMT') && trimmedTime.includes(':')) {
    try {
      const dateObj = new Date(trimmedTime);
      if (!isNaN(dateObj.getTime())) {
        const scriptTimeZone = Session.getScriptTimeZone();
        return Utilities.formatDate(dateObj, scriptTimeZone, 'HH:mm');
      }
    } catch(e) {}
  }
  if (/^\d:\d\d$/.test(trimmedTime)) { return '0' + trimmedTime; }
  return trimmedTime;
}

/**
 * 將操作日誌記錄到 Google Sheet 的 Logs 分頁。
 * @param {string} action 操作類型 (例如: 'addSignup', 'removeSignup', 'modifySignup')
 * @param {string} eventId 相關活動的 ID
 * @param {string} userName 執行操作的使用者名稱
 * @param {string} position 相關的崗位 (例如: 新崗位, 舊崗位 -> 新崗位)
 * @param {string} result 操作結果 (例如: '成功', '失敗', '重複報名', '額滿', '找不到')
 * @param {string} details 操作的詳細描述
 */
function logOperation(action, eventId, userName, position, result, details) {
  try {
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
    // 立即刷新，確保日誌被寫入
    SpreadsheetApp.flush(); 
    console.log(`Logged: ${action} for Event ${eventId}, User ${userName}, Position ${position}, Result: ${result}, Details: ${details}`);
  } catch (logError) {
    console.error(`Failed to log operation to sheet: ${logError.message}\n${logError.stack}`);
  }
}


// -------------------- 資料獲取與處理函式 --------------------
function getMasterData() {
  if (EVENTS_CSV_URL.includes("在此貼上") || SIGNUPS_CSV_URL.includes("在此貼上")) {
    throw new Error("後端程式碼尚未設定 Events 和 Signups 的 CSV 網址。");
  }
  try {
    const cacheBuster = '&_t=' + new Date().getTime();
    const requests = [ 
      { url: EVENTS_CSV_URL + cacheBuster, muteHttpExceptions: true }, 
      { url: SIGNUPS_CSV_URL + cacheBuster, muteHttpExceptions: true } 
    ];
    const responses = UrlFetchApp.fetchAll(requests);
    const eventsResponse = responses[0];
    const signupsResponse = responses[1];
    if (eventsResponse.getResponseCode() !== 200) throw new Error(`無法獲取 Events CSV 資料。錯誤碼: ${eventsResponse.getResponseCode()}`);
    if (signupsResponse.getResponseCode() !== 200) throw new Error(`無法獲取 Signups CSV 資料。錯誤碼: ${signupsResponse.getResponseCode()}`);
    const eventsData = Utilities.parseCsv(eventsResponse.getContentText());
    const signupsData = Utilities.parseCsv(signupsResponse.getContentText());
    
    console.log(`getMasterData: 成功從 CSV 獲取資料。`); // Debug
    console.log(`getMasterData: Events CSV 讀取到 ${eventsData.length} 行 (含標頭)。`); // Debug
    console.log(`getMasterData: Signups CSV 讀取到 ${signupsData.length} 行 (含標頭)。`); // Debug

    eventsData.shift(); // 移除標頭
    signupsData.shift(); // 移除標頭

    console.log(`getMasterData: Events Data (移除標頭後): ${eventsData.length} 筆`); // Debug
    console.log(`getMasterData: Signups Data (移除標頭後): ${signupsData.length} 筆`); // Debug
    // console.log('getMasterData: Sample Events Data (first 2 rows):', eventsData.slice(0, 2)); // Debug: Sample data
    // console.log('getMasterData: Sample Signups Data (first 2 rows):', signupsData.slice(0, 2)); // Debug: Sample data


    const scriptTimeZone = Session.getScriptTimeZone();
    const eventsMap = new Map();
    eventsData.forEach(row => {
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
        } catch(e) {
            console.error(`getMasterData: 處理活動 ID ${eventId} 的日期時發生錯誤: ${row[2]}. 錯誤: ${e.message}`);
        }
      }
    });
    console.log(`getMasterData: 建立 eventsMap 完成。共 ${eventsMap.size} 個有效活動。`); // Debug
    return { eventsData, signupsData, eventsMap };
  } catch (err) {
    console.error(`getMasterData (CSV) 失敗: ${err.message}`, err.stack);
    throw new Error(`從 CSV 網址讀取資料時發生錯誤: ${err.message}`);
  }
}

function getEventsAndSignups() {
  try {
    const { eventsData, signupsData } = getMasterData(); // 這會呼叫一次 getMasterData
    const scriptTimeZone = Session.getScriptTimeZone();
    const signupsByEventId = new Map();
    signupsData.forEach(row => {
      const eventId = row[1];
      if (!eventId) return;
      if (!signupsByEventId.has(eventId)) { signupsByEventId.set(eventId, []); }
      signupsByEventId.get(eventId).push({ user: row[2], position: row[4] });
    });
    console.log(`getEventsAndSignups: signupsByEventId 建立完成。共 ${signupsByEventId.size} 個活動有報名資料。`); // Debug

    return eventsData.map(row => {
      const [eventId, title, rawDateString, rawStartTime, rawEndTime, maxAttendees, positions, description] = row;
      if (!eventId || !title || !rawDateString || !rawStartTime) {
          console.warn(`getEventsAndSignups: 跳過無效活動資料: ${JSON.stringify(row)}`);
          return null;
      }
      
      const paddedStartTime = padTime(rawStartTime);
      const paddedEndTime = padTime(rawEndTime);
      
      let datePart;
      try {
          const dateObj = new Date(rawDateString);
          if (isNaN(dateObj.getTime())) {
              console.warn(`getEventsAndSignups: 無效日期字串，跳過活動 ID ${eventId}: ${rawDateString}`);
              return null;
          }
          datePart = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy-MM-dd");
      } catch (e) {
          console.error(`getEventsAndSignups: 處理日期時發生錯誤 ${rawDateString} for event ${eventId}: ${e.message}`);
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
  } catch(err) { console.error("getEventsAndSignups 失敗:", err.message, err.stack); throw err; }
}

function getUnifiedSignups(searchText, startDateStr, endDateStr) {
  console.log(`getUnifiedSignups: 接收參數 - searchText: "${searchText}", startDateStr: "${startDateStr}", endDateStr: "${endDateStr}"`); // Debug
  try {
    const { signupsData, eventsMap } = getMasterData(); // 再次呼叫 getMasterData 確保資料是最新的
    console.log(`getUnifiedSignups: 從 getMasterData 取得 ${signupsData.length} 筆報名資料，${eventsMap.size} 筆活動資料。`); // Debug
    const searchLower = searchText ? searchText.toLowerCase() : '';
    const dayMap = ['日', '一', '二', '三', '四', '五', '六'];

    let start = null;
    let end = null;
    try {
        if (startDateStr) {
            start = new Date(startDateStr);
            start.setHours(0, 0, 0, 0); // 確保從當天0點開始
            console.log(`getUnifiedSignups: 解析開始日期為 ${start}`); // Debug
        }
        if (endDateStr) {
            end = new Date(endDateStr);
            end.setHours(23, 59, 59, 999); // 確保到當天23:59:59.999結束
            console.log(`getUnifiedSignups: 解析結束日期為 ${end}`); // Debug
        }
    } catch(e) {
        console.error(`getUnifiedSignups: 日期解析錯誤 - startDateStr: ${startDateStr}, endDateStr: ${endDateStr}. 錯誤: ${e.message}`);
        throw new Error('日期格式不正確。');
    }

    let filteredSignups = signupsData.filter((row, index) => {
      // 確保row有足夠的元素，避免 undefined 錯誤
      if (row.length < 5) {
          console.warn(`getUnifiedSignups: 跳過格式不正確的報名資料列 (索引 ${index}): ${JSON.stringify(row)}`);
          return false;
      }

      const eventId = row[1];
      const eventDetail = eventsMap.get(eventId);

      if (!eventDetail) {
          console.warn(`getUnifiedSignups: 報名資料列 (索引 ${index}) 的 Event ID: "${eventId}" 在 eventsMap 中找不到，跳過。`); // Debug
          return false; 
      }
      if (!eventDetail.dateObj || isNaN(eventDetail.dateObj.getTime())) { 
          console.warn(`getUnifiedSignups: 活動 ID: "${eventId}" 的日期物件無效，跳過。`); // Debug
          return false; 
      }

      const eventDate = eventDetail.dateObj;
      
      // 日期過濾
      if (start && eventDate < start) {
          // console.log(`getUnifiedSignups: 跳過報名 - 日期早於開始日期。事件日期: ${eventDate.toISOString().slice(0, 10)}, 開始日期: ${start.toISOString().slice(0, 10)}`); // Debug
          return false;
      }
      if (end && eventDate > end) {
          // console.log(`getUnifiedSignups: 跳過報名 - 日期晚於結束日期。事件日期: ${eventDate.toISOString().slice(0, 10)}, 結束日期: ${end.toISOString().slice(0, 10)}`); // Debug
          return false;
      }

      // 關鍵字過濾
      if (searchLower) {
        const user = String(row[2] || '').toLowerCase();
        const position = String(row[4] || '').toLowerCase();
        const eventTitle = String(eventDetail.title || '').toLowerCase();
        const eventDescription = String(eventDetail.description || '').toLowerCase();
        
        if (!user.includes(searchLower) && 
            !position.includes(searchLower) && 
            !eventTitle.includes(searchLower) && 
            !eventDescription.includes(searchLower)) {
          // console.log(`getUnifiedSignups: 跳過報名 - 不符合關鍵字 "${searchLower}"。姓名: ${user}, 崗位: ${position}, 活動: ${eventTitle}`); // Debug
          return false;
        }
      }
      // console.log(`getUnifiedSignups: 報名資料列 (索引 ${index}) 通過所有過濾條件。`); // Debug
      return true;
    });

    console.log(`getUnifiedSignups: 經過過濾後，剩餘 ${filteredSignups.length} 筆報名資料。`); // Debug

    const mappedSignups = filteredSignups.map(row => {
      const eventId = row[1];
      const eventDetail = eventsMap.get(eventId) || { title: '未知事件', dateString: '日期無效'};
      let eventDayOfWeek = eventDetail.dateObj ? dayMap[eventDetail.dateObj.getDay()] : '';
      const rawTimestamp = row[3] || '';
      return {
        signupId: row[0], eventId: eventId, eventTitle: eventDetail.title, 
        eventDate: eventDetail.dateString, eventDayOfWeek: eventDayOfWeek,
        startTime: eventDetail.startTime || '',
        endTime: eventDetail.endTime || '',
        user: row[2], timestamp: rawTimestamp, position: row[4]
      };
    });
    
    mappedSignups.sort((a, b) => {
        // 確保日期有效性再進行比較
        const dateA = new Date(a.eventDate);
        const dateB = new Date(b.eventDate);
        if (isNaN(dateA.getTime()) || isNaN(dateB.getTime())) {
            console.warn(`getUnifiedSignups: 排序時發現無效日期: A: ${a.eventDate}, B: ${b.eventDate}`); // Debug
            return 0; // 如果日期無效，則不改變順序
        }
        return dateA.getTime() - dateB.getTime();
    });

    console.log(`getUnifiedSignups: 最終回傳 ${mappedSignups.length} 筆報名資料。`); // Debug
    // console.log('getUnifiedSignups: Final mappedSignups (first 5):', mappedSignups.slice(0, 5)); // Debug: Sample final data
    return mappedSignups;
  } catch(err) { 
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
      item.signups.sort((a,b) => a.position.localeCompare(b.position));
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

// -------------------- 主要資料寫入函式 --------------------

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
    
    // 檢查活動是否存在
    const eventData = eventsSheet.getDataRange().getValues().find(row => row[0] === eventId);
    if (!eventData) {
      logOperation('addSignup', eventId, userName, position, '失敗', '找不到此活動');
      return { status: 'error', message: '找不到此活動。' };
    }

    // 檢查活動日期時間是否有效且未結束
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
    
    // 檢查是否已報名
    if (currentSignups.some(row => row[2] === userName)) { 
      logOperation('addSignup', eventId, userName, position, '重複報名', '您已經報名過此活動了');
      return { status: 'error', message: '您已經報名過此活動了。' }; 
    }
    
    const existingPositionHolder = currentSignups.find(row => row[4] === position);
    // 檢查崗位是否被佔用
    if (existingPositionHolder) { 
      logOperation('addSignup', eventId, userName, position, '崗位重複', `此崗位已被 ${existingPositionHolder[2]} 報名`);
      return { status: 'confirm_backup', message: `此崗位目前由 [${existingPositionHolder[2]}] 報名，您要改為報名備援嗎？` }; 
    }

    const maxAttendees = eventData[5];
    // 檢查是否額滿
    if (currentSignups.length >= maxAttendees) { 
      logOperation('addSignup', eventId, userName, position, '額滿', '活動總人數已額滿');
      return { status: 'error', message: '活動總人數已額滿。' }; 
    }

    const newSignupId = 'su' + new Date().getTime();
    const newRowData = [newSignupId, eventId, userName, new Date(), position];
    signupsSheet.appendRow(newRowData);
    SpreadsheetApp.flush(); // 確保資料寫入

    // 寫入驗證 (再次確認是否成功寫入，並與日誌同步)
    const lastRowValues = signupsSheet.getRange(signupsSheet.getLastRow(), 1, 1, 5).getValues()[0];
    if (lastRowValues[0] === newSignupId && lastRowValues[2] === userName) {
      logOperation('addSignup', eventId, userName, position, '成功', '報名資料寫入並驗證成功');
      return { status: 'success', message: '報名成功！' };
    } else {
      console.error(`寫入驗證失敗！預期寫入 [${newSignupId}, ${userName}]，但最後一列是 [${lastRowValues[0]}, ${lastRowValues[2]}]`);
      if (lastRowValues[0] === newSignupId) { 
        signupsSheet.deleteRow(signupsSheet.getLastRow()); // 嘗試回滾
        logOperation('addSignup', eventId, userName, position, '失敗', '報名資料寫入驗證失敗，已嘗試回滾');
      } else {
        logOperation('addSignup', eventId, userName, position, '失敗', '報名資料寫入驗證失敗，無法回滾');
      }
      throw new Error('報名資料寫入驗證失敗，請稍後再試。');
    }
  } catch(err) { 
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
    SpreadsheetApp.flush(); // 確保資料寫入

    const lastRowValues = signupsSheet.getRange(signupsSheet.getLastRow(), 1, 1, 5).getValues()[0];
    if (lastRowValues[0] === newSignupId && lastRowValues[2] === userName) {
      logOperation('addBackupSignup', eventId, userName, backupPosition, '成功', '備援報名資料寫入並驗證成功');
      return { status: 'success' };
    } else {
      console.error(`備援寫入驗證失敗！預期寫入 [${newSignupId}, ${userName}]，但最後一列是 [${lastRowValues[0]}, ${lastRowValues[2]}]`);
      if (lastRowValues[0] === newSignupId) { 
        signupsSheet.deleteRow(signupsSheet.getLastRow()); // 嘗試回滾
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
  let rowToDeleteIndex = -1;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    const signupsData = signupsSheet.getDataRange().getValues();
    
    // 從後往前找，避免刪除行後索引變動
    for (let i = signupsData.length - 1; i >= 1; i--) { 
      if (signupsData[i][1] === eventId && signupsData[i][2] === userName) { 
        rowToDeleteIndex = i + 1; // Apps Script 的行號是 1-based
        removedPosition = signupsData[i][4]; // 獲取崗位信息
        signupsSheet.deleteRow(rowToDeleteIndex); 
        logOperation('removeSignup', eventId, userName, removedPosition, '成功', `已刪除行 ${rowToDeleteIndex}`);
        return { status: 'success', message: '已為您取消報名。' }; 
      } 
    }
    // 如果找不到
    logOperation('removeSignup', eventId, userName, '', '找不到', '找不到您的報名紀錄');
    return { status: 'error', message: '找不到您的報名紀錄。' };
  } catch(err) { 
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
  let targetRowIndex = -1;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    const signupsData = signupsSheet.getDataRange().getValues();
    
    // 尋找目標行
    for (let i = 1; i < signupsData.length; i++) { // 從第2行(索引1)開始
      if (signupsData[i][0] === signupId) {
        targetRowIndex = i + 1; // Google Sheet 的行號是 1-based
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

    // 檢查新崗位是否已被佔用
    const positionTaken = signupsData.some(row => 
      row[1] === eventId && // 同一個活動
      row[4] === newPosition && // 新崗位已存在
      row[0] !== signupId // 但不是目前正在修改的這一行
    );

    if (positionTaken) { 
      logOperation('modifySignup', eventId, userName, `${oldPosition} -> ${newPosition}`, '崗位重複', `無法修改，崗位 [${newPosition}] 已被其他人報名`);
      return { status: 'error', message: `無法修改，崗位 [${newPosition}] 已被其他人報名。` }; 
    }

    signupsSheet.getRange(targetRowIndex, 5).setValue(newPosition); // 修改崗位 (第5欄)
    SpreadsheetApp.flush(); // 確保資料寫入

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
    blob.setName(`AV立願報名記錄_${new Date().toISOString().slice(0,10)}.xlsx`); 
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
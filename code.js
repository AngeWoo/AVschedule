// ====================================================================
//  檔案：code.gs (AV放送立願報名行事曆 - 後端 API)
//  版本：6.3 (修正報名時間欄位顯示問題)
// ====================================================================

// --- 已填入您提供的 CSV 網址 ---
const EVENTS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=795926947&single=true&output=csv";
const SIGNUPS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=767193030&single=true&output=csv";
// ---------------------------------------------------

// --- 全域設定 ---
const SHEET_ID = '1gBIlhEKPQHBvslY29veTdMEJeg2eVcaJx_7A-8cTWIM';
const EVENTS_SHEET_NAME = 'Events';
const SIGNUPS_SHEET_NAME = 'Signups';

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

    switch (functionName) {
      case 'getEventsAndSignups': result = getEventsAndSignups(); break;
      case 'getAllSignups': result = getAllSignups(params.startDateStr, params.endDateStr); break;
      case 'getStatsData': result = getStatsData(params.startDateStr, params.endDateStr); break;
      case 'getMySignups': result = getMySignups(params.userName, params.startDateStr, params.endDateStr); break;
      case 'getUniqueSignupNames': result = getUniqueSignupNames(params.startDateStr, params.endDateStr); break;
      case 'addSignup': result = addSignup(params.eventId, params.userName, params.position); break;
      case 'addBackupSignup': result = addBackupSignup(params.eventId, params.userName, params.position); break;
      case 'removeSignup': result = removeSignup(params.eventId, params.userName); break;
      case 'createTempSheetAndExport': result = createTempSheetAndExport(params.startDateStr, params.endDateStr); break;
      case 'scheduleDeletion': scheduleDeletion(params.sheetName); result = { status: 'scheduled' }; break;
      case 'exportAndEmailReport': result = exportAndEmailReport(params.startDateStr, params.endDateStr, params.recipientEmail); break;
      case 'deleteAllTempSheets': result = deleteAllTempSheets(); break;
      case 'generateStatsPdfReport': result = generateStatsPdfReport(params.chartImageData, params.startDateStr, params.endDateStr, params.fullDetails); break;
      default: throw new Error(`未知的函式名稱: ${functionName}`);
    }

    responsePayload = { status: 'success', data: result };
  } catch (err) {
    console.error(`doGet 執行失敗: ${err.message}\n${err.stack}`);
    responsePayload = { status: 'error', message: err.message };
  }

  const callbackFunctionName = e.parameter.callback;
  const jsonpResponse = `${callbackFunctionName}(${JSON.stringify(responsePayload)})`;
  
  return ContentService.createTextOutput(jsonpResponse).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  return doGet(e);
}


// -------------------- 核心資料獲取函式 --------------------
function getMasterData() {
  if (EVENTS_CSV_URL.includes("在此貼上") || SIGNUPS_CSV_URL.includes("在此貼上")) {
    throw new Error("後端程式碼尚未設定 Events 和 Signups 的 CSV 網址。");
  }

  try {
    const requests = [
      { url: EVENTS_CSV_URL, muteHttpExceptions: true },
      { url: SIGNUPS_CSV_URL, muteHttpExceptions: true }
    ];
    const responses = UrlFetchApp.fetchAll(requests);
    
    const eventsResponse = responses[0];
    const signupsResponse = responses[1];

    if (eventsResponse.getResponseCode() !== 200) throw new Error(`無法獲取 Events CSV 資料。錯誤碼: ${eventsResponse.getResponseCode()}`);
    if (signupsResponse.getResponseCode() !== 200) throw new Error(`無法獲取 Signups CSV 資料。錯誤碼: ${signupsResponse.getResponseCode()}`);

    const eventsData = Utilities.parseCsv(eventsResponse.getContentText());
    const signupsData = Utilities.parseCsv(signupsResponse.getContentText());

    eventsData.shift(); // 移除標題行
    signupsData.shift(); // 移除標題行

    const scriptTimeZone = Session.getScriptTimeZone();
    const eventsMap = new Map();
    eventsData.forEach(row => {
      const eventId = row[0];
      if (eventId) {
        eventsMap.set(eventId, {
          title: row[1] || '未知事件',
          dateString: row[2] ? Utilities.formatDate(new Date(row[2]), scriptTimeZone, 'yyyy/MM/dd') : '',
          dateObj: new Date(row[2]),
          startTime: row[3] || '',
          endTime: row[4] || '',
          maxAttendees: parseInt(row[5], 10) || 999
        });
      }
    });
    return { eventsData, signupsData, eventsMap };
  } catch (err) {
    console.error(`getMasterData (CSV) 失敗: ${err.message}`, err.stack);
    throw new Error(`從 CSV 網址讀取資料時發生錯誤: ${err.message}`);
  }
}

// -------------------- 主要功能函式 (已修正) --------------------

function getEventsAndSignups() {
  try {
    const { eventsData, signupsData } = getMasterData();
    const scriptTimeZone = Session.getScriptTimeZone();

    const signupsByEventId = new Map();
    signupsData.forEach(row => {
      const eventId = row[1];
      if (!eventId) return;
      if (!signupsByEventId.has(eventId)) {
        signupsByEventId.set(eventId, []);
      }
      signupsByEventId.get(eventId).push({ user: row[2], position: row[4] });
    });
    
    const formatDatePart = (dateStr) => {
        try {
            return Utilities.formatDate(new Date(dateStr), scriptTimeZone, "yyyy-MM-dd");
        } catch (e) {
            console.error(`日期值 "${dateStr}" 無效，無法格式化。錯誤: ${e.message}`);
            return null;
        }
    };
    
    const getHourMinuteString = (dateObject) => {
        if (!dateObject || isNaN(dateObject.getTime())) return null;
        return Utilities.formatDate(dateObject, scriptTimeZone, "HH:mm");
    };

    const padTime = (timeStr) => {
      const trimmedTime = String(timeStr || '').trim();
      if (trimmedTime && /^\d:\d\d$/.test(trimmedTime)) {
        return '0' + trimmedTime;
      }
      return trimmedTime;
    };

    return eventsData.map(row => {
      const eventId        = row[0];
      const title          = row[1];
      const rawDateString  = row[2];
      const rawStartTime   = row[3];
      const rawEndTime     = row[4];
      const maxAttendees   = row[5];
      const positions      = row[6];
      const description    = row[7];

      if (!eventId || !title || !rawDateString || !rawStartTime) {
        console.warn(`跳過無效事件 (缺少必要欄位): ID=${eventId || 'N/A'}, Title=${title || 'N/A'}, Date=${rawDateString || 'N/A'}, StartTime=${rawStartTime || 'N/A'}`);
        return null;
      }

      const paddedStartTime = padTime(rawStartTime);
      const paddedEndTime = padTime(rawEndTime);
      
      let eventStartDateTime = null;
      let eventEndDateTime = null;

      try {
          const isoDatePart = rawDateString.replace(/\//g, '-');
          
          const startDateTimeString = `${isoDatePart}T${paddedStartTime}:00`;
          eventStartDateTime = new Date(startDateTimeString);

          if (isNaN(eventStartDateTime.getTime())) {
              console.error(`無法解析活動開始日期/時間: "${startDateTimeString}". 事件ID: ${eventId}. 跳過此事件。`);
              return null; 
          }

          if (paddedEndTime) {
              const endDateTimeString = `${isoDatePart}T${paddedEndTime}:00`;
              eventEndDateTime = new Date(endDateTimeString);
              if (isNaN(eventEndDateTime.getTime())) {
                  console.warn(`無法解析活動結束日期/時間: "${endDateTimeString}". 事件ID: ${eventId}. 此事件將不設定結束時間。`);
                  eventEndDateTime = null;
              }
          }
      } catch (e) {
          console.error(`處理事件日期/時間時發生未預期錯誤: ${e.message}. 事件ID: ${eventId}. 堆疊追蹤: ${e.stack}`);
          return null; 
      }
      
      const formattedStartTime = getHourMinuteString(eventStartDateTime);
      const formattedEndTime = getHourMinuteString(eventEndDateTime);

      if (!formattedStartTime) {
          console.error(`格式化後的開始時間為 null，跳過此事件: ID=${eventId}`);
          return null;
      }

      const signups = signupsByEventId.get(eventId) || [];
      const signupCount = signups.length;
      const parsedMaxAttendees = parseInt(maxAttendees, 10) || 999;
      const isFull = signupCount >= parsedMaxAttendees;
      
      let eventColor, textColor;
      if (isFull) { 
          eventColor = '#e74c3c';
          textColor = 'white'; 
      } 
      else if (signupCount > 0) { 
          eventColor = '#0d6efd';
          textColor = 'white'; 
      } 
      else { 
          eventColor = '#adb5bd';
          textColor = '#212529'; 
      }
      
      return {
        id: eventId,
        title: title,
        start: `${formatDatePart(rawDateString)}T${formattedStartTime}`, 
        end: formattedEndTime ? `${formatDatePart(rawDateString)}T${formattedEndTime}` : null,
        backgroundColor: eventColor,
        borderColor: eventColor,
        textColor: textColor,
        extendedProps: { 
            full_title: title, 
            description: description, 
            maxAttendees: parsedMaxAttendees, 
            signups, 
            positions: (positions || '').split(',').map(p => p.trim()).filter(p => p), 
            startTime: formattedStartTime,
            endTime: formattedEndTime
        }
      };
    }).filter(e => e !== null);
  } catch(err) { 
    console.error("getEventsAndSignups 失敗 (頂層捕獲):", err.message, err.stack);
    throw err; 
  }
}

function getAllSignups(startDateStr, endDateStr) {
  try {
    const { signupsData, eventsMap } = getMasterData();
    const scriptTimeZone = Session.getScriptTimeZone();

    let filteredSignups = signupsData;
    if (startDateStr && endDateStr) {
      const start = new Date(startDateStr);
      const end = new Date(endDateStr);
      end.setHours(23, 59, 59, 999);
      
      filteredSignups = signupsData.filter(row => {
        const eventDetail = eventsMap.get(row[1]); 
        return eventDetail && eventDetail.dateObj && eventDetail.dateObj >= start && eventDetail.dateObj <= end;
      });
    }
    
    const mappedSignups = filteredSignups.map(row => {
      const eventId = row[1];
      const eventDetail = eventsMap.get(eventId) || { title: '未知事件', dateString: '日期無效'};
      
      return {
        signupId: row[0],
        eventId: eventId,
        eventTitle: eventDetail.title,
        eventDate: eventDetail.dateString,
        user: row[2],
        // 【修改】直接抓取D欄的文字內容，不再進行日期解析
        timestamp: row[3] || '',
        position: row[4]
      };
    });

    // 【修改】只根據活動日期排序，因為報名時間現在是文字，無法準確排序
    mappedSignups.sort((a, b) => {
        const dateA = new Date(a.eventDate);
        const dateB = new Date(b.eventDate);
        return dateA.getTime() - dateB.getTime();
    });

    return mappedSignups;

  } catch(err) { console.error("getAllSignups 失敗:", err.message, err.stack); throw err; }
}

function getStatsData(startDateStr, endDateStr) {
  try {
    const { signupsData, eventsMap } = getMasterData();
    const scriptTimeZone = Session.getScriptTimeZone();

    const allSignups = getAllSignups(startDateStr, endDateStr);

    if (!allSignups || allSignups.length === 0) {
      return { labels: [], data: [], fullDetails: [] };
    }

    const statsByEventId = {};
    allSignups.forEach(signup => {
      const eventId = signup.eventId;
      if (!statsByEventId[eventId]) {
        const eventInfoFromMap = eventsMap.get(eventId) || {}; 
        const eventInfo = {
          title: eventInfoFromMap.title || signup.eventTitle,
          date: eventInfoFromMap.dateString || signup.eventDate,
          startTime: eventInfoFromMap.startTime || '',
          endTime: eventInfoFromMap.endTime || '',
          maxAttendees: eventInfoFromMap.maxAttendees || 999
        };
        statsByEventId[eventId] = {
          count: 0,
          signups: [],
          eventInfo: eventInfo,
          label: `${eventInfo.title} (${eventInfo.date})`
        };
      }
      statsByEventId[eventId].count++;
      statsByEventId[eventId].signups.push({ user: signup.user, position: signup.position });
    });

    const processedData = Object.values(statsByEventId);
    processedData.sort((a, b) => {
      const dateA = new Date(a.eventInfo.date).getTime();
      const dateB = new Date(b.eventInfo.date).getTime();
      return dateA - dateB;
    });

    const labels = processedData.map(item => item.label);
    const data = processedData.map(item => item.count);
    const fullDetails = processedData.map(item => {
      item.signups.sort((a,b) => a.position.localeCompare(b.position));
      return {
        eventInfo: item.eventInfo,
        signups: item.signups
      };
    });

    return { labels, data, fullDetails };
  } catch (err) {
    console.error("getStatsData 失敗:", err.message, err.stack);
    throw err;
  }
}

function addSignup(eventId, userName, position) {
  if (!eventId || !userName || !position) { return { status: 'error', message: '缺少必要資訊 (事件ID, 姓名或崗位)。' }; }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const eventsSheet = ss.getSheetByName(EVENTS_SHEET_NAME);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    const eventData = eventsSheet.getDataRange().getValues().find(row => row[0] === eventId);
    if (!eventData) return { status: 'error', message: '找不到此活動。' };
    const eventDatePart = new Date(eventData[2]);
    const eventTimePart = new Date(eventData[4]);
    const eventEndDateTime = new Date( eventDatePart.getFullYear(), eventDatePart.getMonth(), eventDatePart.getDate(), eventTimePart.getHours(), eventTimePart.getMinutes(), eventTimePart.getSeconds() );
    if (new Date() > eventEndDateTime) { return { status: 'error', message: '此活動已結束，無法再進行報名。' }; }
    const currentSignups = signupsSheet.getDataRange().getValues().filter(row => row[1] === eventId);
    if (currentSignups.some(row => row[2] === userName)) { return { status: 'error', message: '您已經報名過此活動了。' }; }
    const existingPositionHolder = currentSignups.find(row => row[4] === position);
    if (existingPositionHolder) { 
      return { status: 'confirm_backup', message: `此崗位目前由 [${existingPositionHolder[2]}] 報名，您要改為報名備援嗎？`, }; 
    }
    const maxAttendees = eventData[5];
    if (currentSignups.length >= maxAttendees) { return { status: 'error', message: '很抱歉，活動總人數已額滿。' }; }
    const newSignupId = 'su' + new Date().getTime();
    signupsSheet.appendRow([newSignupId, eventId, userName, new Date(), position]);
    return { status: 'success', message: '報名成功！' };
  } catch(err) { console.error("addSignup 失敗:", err.message, err.stack); throw new Error('報名時發生未預期的錯誤。');
  } finally { lock.releaseLock(); }
}

function removeSignup(eventId, userName) {
  if (!eventId || !userName) return { status: 'error', message: '缺少事件ID或姓名。' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    const eventsSheet = ss.getSheetByName(EVENTS_SHEET_NAME);
    const eventData = eventsSheet.getDataRange().getValues().find(row => row[0] === eventId);
    if (!eventData) return { status: 'error', message: '找不到對應的活動資訊。' };
    const eventDatePart = new Date(eventData[2]);
    const eventTimePart = new Date(eventData[4]);
    const eventEndDateTime = new Date( eventDatePart.getFullYear(), eventDatePart.getMonth(), eventDatePart.getDate(), eventTimePart.getHours(), eventTimePart.getMinutes(), eventTimePart.getSeconds() );
    if (new Date() > eventEndDateTime) { return { status: 'error', message: '此活動已結束，無法取消報名。' }; }
    const signupsData = signupsSheet.getDataRange().getValues();
    let found = false;
    for (let i = signupsData.length - 1; i >= 1; i--) { if (signupsData[i][1] === eventId && signupsData[i][2] === userName) { signupsSheet.deleteRow(i + 1); found = true; break; } }
    if (found) { return { status: 'success', message: '已為您取消報名。' }; }
    return { status: 'error', message: '找不到您的報名紀錄。' };
  } catch(err) { console.error("removeSignup 失敗:", err.message, err.stack);
    throw new Error('取消報名時發生未預期的錯誤。');
  } finally { lock.releaseLock(); }
}

// --- 其他輔助函式 (無需修改) ---
function addBackupSignup(eventId, userName, position) { if (!eventId || !userName || !position) { return { status: 'error', message: '缺少必要資訊。' }; } const lock = LockService.getScriptLock(); lock.waitLock(15000); try { const ss = SpreadsheetApp.openById(SHEET_ID); const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME); const finalPosition = `${position} (備援)`; const newSignupId = 'su' + new Date().getTime(); signupsSheet.appendRow([newSignupId, eventId, userName, new Date(), finalPosition]); return { status: 'success', message: '已成功為您報名備援！' }; } catch (err) { console.error("addBackupSignup 失敗:", err.message, err.stack); throw new Error('備援報名時發生錯誤。'); } finally { lock.releaseLock(); } }
function getMySignups(userName, startDateStr, endDateStr) { if (!userName) return []; try { return getAllSignups(startDateStr, endDateStr).filter(signup => signup.user === userName); } catch (err) { console.error("getMySignups 失敗:", err.message, err.stack); throw err; } }
function getUniqueSignupNames(startDateStr, endDateStr) { try { const signups = getAllSignups(startDateStr, endDateStr); if (!signups || signups.length === 0) { return []; } const nameSet = new Set(signups.map(s => s.user)); return Array.from(nameSet).sort((a, b) => a.localeCompare(b, 'zh-Hant')); } catch (err) { console.error("getUniqueSignupNames 失敗:", err.message, err.stack); throw err; } }
function createTempSheetAndExport(startDateStr, endDateStr) { const ss = SpreadsheetApp.openById(SHEET_ID); const spreadsheetId = ss.getId(); let tempSheet = null; let tempFile = null; try { const tempSheetName = "匯出報表_" + new Date().getTime(); tempSheet = ss.insertSheet(tempSheetName); const data = getAllSignups(startDateStr, endDateStr); if (!data || data.length === 0) { ss.deleteSheet(tempSheet); return { status: 'nodata', message: '在選定的日期範圍內沒有任何報名記錄。' }; } data.sort((a, b) => { const dateA = new Date(a.eventDate).getTime() || 0; const dateB = new Date(b.eventDate).getTime() || 0; if (dateA !== dateB) { return dateA - dateB; } return 0; }); const dateRangeText = `(${startDateStr || '所有'} - ${endDateStr || '所有'})`; const title = `AV放送立願報名記錄 ${dateRangeText}`; tempSheet.getRange("A1:E1").merge().setValue(title).setFontWeight("bold").setFontSize(15).setHorizontalAlignment('center'); const headers = ["活動日期", "活動", "崗位", "報名者", "報名時間"]; const fields = ["eventDate", "eventTitle", "position", "user", "timestamp"]; const headerRange = tempSheet.getRange(2, 1, 1, headers.length); headerRange.setValues([headers]).setFontWeight("bold").setBackground("#d9ead3").setFontSize(15).setHorizontalAlignment('center'); const outputData = data.map(row => fields.map(field => row[field] || "")); if (outputData.length > 0) { const dataRange = tempSheet.getRange(3, 1, outputData.length, headers.length); dataRange.setValues(outputData).setFontSize(15).setHorizontalAlignment('left'); } tempSheet.setColumnWidth(1, 160); tempSheet.setColumnWidth(2, 420); tempSheet.setColumnWidth(3, 220); tempSheet.setColumnWidth(4, 160); tempSheet.setColumnWidth(5, 280); SpreadsheetApp.flush(); const blob = Drive.Files.export(spreadsheetId, MimeType.MICROSOFT_EXCEL, { gid: tempSheet.getSheetId(), alt: 'media' }); blob.setName(`AV立願報名記錄_${new Date().toISOString().slice(0,10)}.xlsx`); tempFile = DriveApp.createFile(blob); tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); const downloadUrl = tempFile.getWebContentLink(); scheduleDeletion(tempFile.getId()); ss.deleteSheet(tempSheet); return { status: 'success', url: downloadUrl, fileId: tempFile.getId() }; } catch (e) { console.error("createTempSheetAndExport 失敗: " + e.toString()); if (tempFile) { try { DriveApp.getFileById(tempFile.getId()).setTrashed(true); } catch (f) { console.error("無法刪除臨時檔案: " + f.toString()); } } if (tempSheet && ss.getSheetByName(tempSheet.getName())) { ss.deleteSheet(tempSheet); } throw new Error('匯出時發生錯誤: ' + e.toString()); } }
function exportAndEmailReport(startDateStr, endDateStr, recipientEmail) { const ss = SpreadsheetApp.openById(SHEET_ID); const spreadsheetId = ss.getId(); let tempSheet = null; try { console.log(`開始處理郵件報表請求 - 收件人: ${recipientEmail}, 日期範圍: ${startDateStr} 到 ${endDateStr}`); if (!recipientEmail) { return { status: 'error', message: '未提供收件人Email地址。' }; } const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; if (!emailRegex.test(recipientEmail)) { return { status: 'error', message: '請輸入有效的 Email 地址格式。' }; } const tempSheetName = "郵寄報表_" + new Date().getTime(); tempSheet = ss.insertSheet(tempSheetName); const data = getAllSignups(startDateStr, endDateStr); if (!data || data.length === 0) { ss.deleteSheet(tempSheet); return { status: 'nodata', message: '在選定的日期範圍內沒有任何報名記錄。' }; } data.sort((a, b) => { const dateA = new Date(a.eventDate).getTime() || 0; const dateB = new Date(b.eventDate).getTime() || 0; return dateA - dateB; }); const dateRangeText = `${startDateStr || '所有'} - ${endDateStr || '所有'}`; const title = `AV放送立願報名記錄 ${dateRangeText}`; const generateTime = new Date().toLocaleString('zh-TW'); tempSheet.getRange("A1:E1").merge() .setValue(title) .setFontWeight("bold") .setFontSize(16) .setHorizontalAlignment('center') .setBackground("#FF8C42") .setFontColor("white"); tempSheet.getRange("A2:E2").merge() .setValue(`報表生成時間：${generateTime} | 總計：${data.length} 筆記錄`) .setFontSize(12) .setHorizontalAlignment('center') .setBackground("#FFDAB3") .setFontColor("#333333"); tempSheet.getRange("A3:E3").merge().setValue(""); const headers = ["活動日期", "活動名稱", "崗位", "報名者", "報名時間"]; const headerRange = tempSheet.getRange(4, 1, 1, headers.length); headerRange.setValues([headers]) .setFontWeight("bold") .setBackground("#F0F4F8") .setFontColor("#334455") .setFontSize(14) .setHorizontalAlignment('center') .setBorder(true, true, true, true, true, true, "#7B8D9F", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); const fields = ["eventDate", "eventTitle", "position", "user", "timestamp"]; const outputData = data.map(row => fields.map(field => row[field] || "")); if (outputData.length > 0) { const dataRange = tempSheet.getRange(5, 1, outputData.length, headers.length); dataRange.setValues(outputData) .setFontSize(12) .setHorizontalAlignment('left'); for (let i = 0; i < outputData.length; i++) { const rowRange = tempSheet.getRange(5 + i, 1, 1, headers.length); if (i % 2 === 0) { rowRange.setBackground("#FDFDFD"); } else { rowRange.setBackground("#F8F9FA"); } rowRange.setBorder(true, true, true, true, false, false, "#E0E6EE", SpreadsheetApp.BorderStyle.SOLID); } } tempSheet.setColumnWidth(1, 120); tempSheet.setColumnWidth(2, 350); tempSheet.setColumnWidth(3, 150); tempSheet.setColumnWidth(4, 120); tempSheet.setColumnWidth(5, 180); const uniqueEvents = [...new Set(data.map(item => item.eventTitle))].length; const uniqueUsers = [...new Set(data.map(item => item.user))].length; const summaryRow = 5 + outputData.length + 2; tempSheet.getRange(summaryRow, 1, 1, 5).merge() .setValue(`📊 統計摘要：共 ${uniqueEvents} 個活動 | ${uniqueUsers} 位不重複報名者 | ${data.length} 筆報名記錄`) .setFontWeight("bold") .setHorizontalAlignment('center') .setBackground("#E8F4FD") .setFontColor("#2C5282") .setBorder(true, true, true, true, true, true, "#4299E1", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); SpreadsheetApp.flush(); const rawBlob = Drive.Files.export(spreadsheetId, MimeType.MICROSOFT_EXCEL, { gid: tempSheet.getSheetId(), alt: 'media' }); const excelFileName = `AV立願報名記錄_${startDateStr}_${endDateStr}_${new Date().toISOString().slice(0,10)}.xlsx`; const excelBlob = Utilities.newBlob(rawBlob.getBytes(), MimeType.MICROSOFT_EXCEL, excelFileName); const emailSubject = `AV放送立願報名記錄報表 ${dateRangeText}`; const emailBody = ` <div style="font-family: 'Microsoft JhengHei', Arial, sans-serif; max-width: 600px; margin: 0 auto;"> <div style="background: linear-gradient(135deg, #FF8C42, #ff7f2b); color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0;"> <h2 style="margin: 0; font-size: 20px;">📡 AV放送立願報名記錄報表</h2> </div> <div style="background: #f8f9fa; padding: 20px; border-radius: 0 0 8px 8px; border: 1px solid #e9ecef;"> <p style="margin-top: 0; font-size: 14px; color: #333;">您好：</p> <p style="font-size: 14px; color: #333; line-height: 1.6;"> 附件是您所申請的 AV放送立願報名記錄報表，請查收。 </p> <div style="background: white; padding: 15px; border-radius: 6px; margin: 15px 0; border-left: 4px solid #FF8C42;"> <p style="margin: 0; color: #333;"><strong>📅 日期範圍：</strong>${dateRangeText}</p> <p style="margin: 5px 0 0 0; color: #333;"><strong>📊 包含資料：</strong>${data.length} 筆記錄</p> <p style="margin: 5px 0 0 0; color: #333;"><strong>🕒 生成時間：</strong>${generateTime}</p> <p style="margin: 5px 0 0 0; color: #333;"><strong>🎯 活動數量：</strong>${uniqueEvents} 個活動</p> <p style="margin: 5px 0 0 0; color: #333;"><strong>👥 報名人數：</strong>${uniqueUsers} 位不重複人員</p> </div> <div style="background: #fff3cd; padding: 12px; border-radius: 6px; border-left: 4px solid #ffc107;"> <p style="margin: 0; font-size: 13px; color: #856404;"> <strong>💡 使用提示：</strong>此 Excel 檔案可直接開啟，包含完整的格式設計、統計摘要和品牌配色。 </p> </div> <div style="background: #e8f4fd; padding: 12px; border-radius: 6px; border-left: 4px solid #4299e1; margin-top: 10px;"> <p style="margin: 0 0 5px 0; font-size: 13px; color: #2c5282; font-weight: bold;">✨ 報表特色：</p> <ul style="margin: 0; padding-left: 20px; font-size: 12px; color: #2c5282;"> <li>專業的標題和副標題設計</li> <li>交替行顏色便於閱讀</li> <li>完整的統計摘要資訊</li> <li>品牌色彩主題設計</li> </ul> </div> <p style="margin-bottom: 0; font-size: 13px; color: #666; margin-top: 15px;"> 此為系統自動發送郵件，請勿直接回覆。如有疑問請聯繫系統管理員。 </p> </div> <div style="text-align: center; color: #666; font-size: 12px; margin-top: 15px; padding: 10px;"> <p style="margin: 0;">© 2025 AV-Broadcasting Team. All Rights Reserved.</p> <p style="margin: 5px 0 0 0;">此郵件由 AV放送立願系統自動生成</p> </div> </div>`; console.log(`準備發送郵件到：${recipientEmail}`); MailApp.sendEmail({ to: recipientEmail, subject: emailSubject, htmlBody: emailBody, attachments: [excelBlob] }); console.log(`郵件發送成功：${recipientEmail}`); ss.deleteSheet(tempSheet); return { status: 'success', message: `報表已成功寄送至 ${recipientEmail}。共包含 ${data.length} 筆記錄，涵蓋 ${uniqueEvents} 個活動。` }; } catch (e) { console.error("exportAndEmailReport 失敗: " + e.toString()); console.error("錯誤堆疊: " + e.stack); if (tempSheet && ss.getSheetByName(tempSheet.getName())) { try { ss.deleteSheet(tempSheet); console.log("已清理臨時工作表"); } catch (deleteError) { console.error("刪除臨時工作表失敗: " + deleteError.toString()); } } return { status: 'error', message: '郵件寄送時發生錯誤: ' + e.toString() }; } }
function scheduleDeletion(fileId) { if (!fileId) return; try { const trigger = ScriptApp.newTrigger('triggeredDeleteHandler').timeBased().after(24 * 60 * 60 * 1000).create(); PropertiesService.getScriptProperties().setProperty(trigger.getUniqueId(), fileId); console.log(`已成功排程在 24 小時後刪除檔案 '${fileId}'。觸發器ID: ${trigger.getUniqueId()}`); } catch (e) { console.error(`排程刪除檔案 '${fileId}' 時失敗: ${e.toString()}`); } }
function triggeredDeleteHandler(e) { const triggerId = e.triggerUid; const scriptProperties = PropertiesService.getScriptProperties(); const fileId = scriptProperties.getProperty(triggerId); if (fileId) { console.log(`觸發器 ${triggerId} 已啟動，準備刪除檔案: ${fileId}`); deleteFileById(fileId); scriptProperties.deleteProperty(triggerId); } else { console.log(`觸發器 ${triggerId} 已啟動，但找不到對應的檔案 ID 屬性。`); } const allTriggers = ScriptApp.getProjectTriggers(); for (let i = 0; i < allTriggers.length; i++) { if (allTriggers[i].getUniqueId() === triggerId) { ScriptApp.deleteTrigger(allTriggers[i]); console.log(`已刪除一次性觸發器: ${triggerId}`); break; } } }
function deleteFileById(fileId) { if (!fileId) return; try { const file = DriveApp.getFileById(fileId); file.setTrashed(true); console.log(`已成功將檔案移至垃圾桶：'${fileId}'`); } catch (e) { console.error("deleteFileById 失敗: " + e.toString()); } }
function deleteAllTempSheets() { const lock = LockService.getScriptLock(); lock.waitLock(15000); try { const ss = SpreadsheetApp.openById(SHEET_ID); const allSheets = ss.getSheets(); const scriptProperties = PropertiesService.getScriptProperties(); const allTriggers = ScriptApp.getProjectTriggers(); let deletedSheetsCount = 0; let deletedFilesCount = 0; allSheets.forEach(sheet => { const sheetName = sheet.getName(); if (sheetName.startsWith("匯出報表_") || sheetName.startsWith("郵寄報表_")) { ss.deleteSheet(sheet); deletedSheetsCount++; console.log(`已手動刪除工作表: ${sheetName}`); } }); const scriptProps = PropertiesService.getScriptProperties().getProperties(); for (let propKey in scriptProps) { if (scriptProps[propKey].length === 33 && allTriggers.some(t => t.getUniqueId() === propKey && t.getHandlerFunction() === 'triggeredDeleteHandler')) { const fileIdToDelete = scriptProps[propKey]; try { deleteFileById(fileIdToDelete); PropertiesService.getScriptProperties().deleteProperty(propKey); deletedFilesCount++; } catch (f) { console.error(`嘗試刪除檔案 ${fileIdToDelete} 時失敗 (可能已不存在): ${f.toString()}`); } } } allTriggers.forEach(trigger => { if (trigger.getHandlerFunction() === 'triggeredDeleteHandler') { const triggerId = trigger.getUniqueId(); scriptProperties.deleteProperty(triggerId); ScriptApp.deleteTrigger(trigger); console.log(`已刪除觸發器 ID: ${triggerId}`); } }); if (deletedSheetsCount > 0 || deletedFilesCount > 0) { return { status: 'success', message: `已成功清除 ${deletedSheetsCount} 個工作表和 ${deletedFilesCount} 個檔案及其排程。` }; } else { return { status: 'nodata', message: '目前沒有任何可清除的報表連結或檔案。' }; } } catch (e) { console.error("deleteAllTempSheets 失敗: " + e.toString()); throw new Error('清除過程中發生錯誤: ' + e.toString()); } finally { lock.releaseLock(); } }
function deleteSheetByName(sheetName) { if (!sheetName) return; try { const ss = SpreadsheetApp.openById(SHEET_ID); const sheet = ss.getSheetByName(sheetName); if (sheet) { ss.deleteSheet(sheet); console.log(`已成功刪除工作表：'${sheetName}'`); } else { console.log(`嘗試刪除時，找不到工作表：'${sheetName}'`); } } catch (e) { console.error("deleteSheetByName 失敗: " + e.toString()); } }
function generateStatsPdfReport(chartImageData, startDateStr, endDateStr, fullDetails) { let tempDoc = null; let tempFile = null; try { console.log('開始生成PDF報表...'); console.log('圖表數據長度:', chartImageData ? chartImageData.length : 'null'); try { DriveApp.getRootFolder(); } catch (driveError) { console.error('Drive API錯誤:', driveError); throw new Error('Drive API未啟用或無權限。請檢查Advanced Google Services設定。'); } try { const testDoc = DocumentApp.create('test_doc_' + new Date().getTime()); DriveApp.getFileById(testDoc.getId()).setTrashed(true); } catch (docError) { console.error('DocumentApp錯誤:', docError); throw new Error('DocumentApp服務錯誤，請檢查權限設定。'); } const docTitle = `AV放送立願統計報表 (${startDateStr || '所有'} - ${endDateStr || '所有'})`; console.log('創建文檔:', docTitle); tempDoc = DocumentApp.create(docTitle); const body = tempDoc.getBody(); const titleParagraph = body.appendParagraph(docTitle); titleParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER); titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1); body.appendParagraph(''); if (chartImageData && chartImageData.length > 100) { try { console.log('處理圖表圖片...'); const base64Data = chartImageData.includes(',') ? chartImageData.split(',')[1] : chartImageData; if (!base64Data || base64Data.length < 100) { throw new Error('Base64數據無效'); } const decodedImageData = Utilities.base64Decode(base64Data); const blob = Utilities.newBlob(decodedImageData, MimeType.PNG, 'stats_chart.png'); const image = body.appendImage(blob); image.setWidth(500); image.setHeight(300); body.appendParagraph(''); console.log('圖表圖片插入成功'); } catch (imageError) { console.error('處理圖表圖片時發生錯誤:', imageError); body.appendParagraph('無法載入圖表圖片: ' + imageError.message); body.appendParagraph(''); } } else { body.appendParagraph('未提供圖表圖片或數據無效。'); body.appendParagraph(''); } body.appendParagraph('--- 詳細活動報名資訊 ---').setHeading(DocumentApp.ParagraphHeading.HEADING2); if (!fullDetails || fullDetails.length === 0) { body.appendParagraph('在選定日期範圍內沒有詳細報名數據。'); } else { fullDetails.forEach(eventData => { body.appendParagraph(''); body.appendParagraph(`活動: ${eventData.eventInfo.title}`) .setHeading(DocumentApp.ParagraphHeading.HEADING3); body.appendParagraph(`日期: ${eventData.eventInfo.date}`); body.appendParagraph(`時間: ${eventData.eventInfo.startTime}${eventData.eventInfo.endTime ? ' - ' + eventData.eventInfo.endTime : ''}`); body.appendParagraph(`已報名人數: ${eventData.signups.length} / 額滿人數: ${eventData.eventInfo.maxAttendees}`); if (eventData.signups && eventData.signups.length > 0) { body.appendParagraph('報名人員列表:'); eventData.signups.forEach(signup => { body.appendListItem(`${signup.position}: ${signup.user}`); }); } else { body.appendParagraph('無人報名'); } }); } console.log('保存文檔...'); tempDoc.saveAndClose(); const docId = tempDoc.getId(); console.log('文檔ID:', docId); let pdfBlob; try { console.log('嘗試使用Advanced Drive Service匯出PDF...'); pdfBlob = Drive.Files.export(docId, MimeType.PDF, {alt: 'media'}); console.log('使用Advanced Drive Service成功'); } catch (advancedDriveError) { console.error('Advanced Drive Service錯誤:', advancedDriveError); try { console.log('嘗試備用匯出方案...'); const url = `https://docs.google.com/document/d/${docId}/export?format=pdf`; const response = UrlFetchApp.fetch(url, { headers: { authorization: `Bearer ${ScriptApp.getOAuthToken()}` } }); pdfBlob = response.getBlob(); console.log('備用方案成功'); } catch (backupError) { console.error('備用方案也失敗:', backupError); throw new Error('無法匯出PDF，請檢查Drive API設定或權限。詳細錯誤: ' + advancedDriveError.message); } } pdfBlob.setName(`${docTitle}.pdf`); console.log('創建PDF檔案...'); tempFile = DriveApp.createFile(pdfBlob); tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); const downloadUrl = tempFile.getWebContentLink(); console.log('PDF檔案創建成功，URL:', downloadUrl); scheduleDeletion(tempFile.getId()); DriveApp.getFileById(docId).setTrashed(true); console.log('臨時文檔已刪除'); return { url: downloadUrl }; } catch (e) { console.error("generateStatsPdfReport 失敗: " + e.toString()); console.error("錯誤堆疊:", e.stack); if (tempDoc) { try { DriveApp.getFileById(tempDoc.getId()).setTrashed(true); console.log('已清理臨時文檔'); } catch (f) { console.error("無法刪除臨時 Doc: " + f.toString()); } } if (tempFile) { try { DriveApp.getFileById(tempFile.getId()).setTrashed(true); console.log('已清理臨時PDF檔案'); } catch (f) { console.error("無法刪除臨時 PDF 檔案: " + f.toString()); } } throw new Error('生成 PDF 報表時發生錯誤: ' + e.toString()); } }
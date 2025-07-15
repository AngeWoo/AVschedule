// ====================================================================
//  檔案：code.gs (AV放送立願報名行事曆 - 後端 API)
//  版本：6.24-debug (包含新增與刪除的強制日誌記錄)
// ====================================================================

// --- 已填入您提供的 CSV 網址 ---
const EVENTS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=795926947&single=true&output=csv";
const SIGNUPS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRT5YNZcSXbE6ULGft15cba4Mx8kK1Eb7bLftucmkmUGmTxkrA8vw5uerAP2dYqptBHnbmq_3QNOOJx/pub?gid=767193030&single=true&output=csv";
// ---------------------------------------------------

// --- 全域設定 ---
const SHEET_ID = '1gBIlhEKPQHBvslY29veTdMEJeg2eVcaJx_7A-8cTWIM';
const EVENTS_SHEET_NAME = 'Events';
const SIGNUPS_SHEET_NAME = 'Signups';
const LOGS_SHEET_NAME = 'Logs'; // 日誌工作表名稱

// -------------------- 偵錯輔助函式 --------------------
function writeLog(action, eventId, userName, position, result, details) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    if (logSheet) {
      logSheet.insertRowBefore(2).getRange(2, 1, 1, 7).setValues([[new Date(), action, eventId, userName, position, result, details || '']]);
    } else {
        console.error("日誌工作表 '" + LOGS_SHEET_NAME + "' 不存在，無法寫入日誌。");
    }
  } catch (e) {
    console.error("寫入日誌時發生嚴重錯誤: " + e.message);
  }
}

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
      case 'getUnifiedSignups': result = getUnifiedSignups(params.searchText, params.startDateStr, params.endDateStr); break;
      case 'getStatsData': result = getStatsData(params.startDateStr, params.endDateStr); break;
      case 'addSignup': result = addSignup(params.eventId, params.userName, params.position); break;
      case 'addBackupSignup': result = addBackupSignup(params.eventId, params.userName, params.position); break;
      case 'removeSignup': result = removeSignup(params.eventId, params.userName); break;
      case 'createTempSheetAndExport': result = createTempSheetAndExport(params.startDateStr, params.endDateStr); break;
      case 'getSignupsAsTextForToday': result = getSignupsAsTextForToday(); break;
      case 'getSignupsAsTextForTomorrow': result = getSignupsAsTextForTomorrow(); break;
      case 'getSignupsAsTextForDateRange': result = getSignupsAsText(params.startDateStr, params.endDateStr); break;
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

function doPost(e) { return doGet(e); }

// -------------------- 主要資料寫入函式 (偵錯版) --------------------

function addSignup(eventId, userName, position) {
  writeLog('addSignup', eventId, userName, position, '啟動', '開始處理報名請求');
  if (!eventId || !userName || !position) {
    writeLog('addSignup', eventId, userName, position, '失敗', '缺少必要資訊');
    return { status: 'error', message: '缺少必要資訊。' };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    writeLog('addSignup', eventId, userName, position, '失敗', '系統忙碌，無法取得鎖');
    return { status: 'error', message: '系統忙碌中，請稍後再試。' };
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const eventsSheet = ss.getSheetByName(EVENTS_SHEET_NAME);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    if (!eventsSheet || !signupsSheet) { throw new Error(`工作表不存在! Events: ${!!eventsSheet}, Signups: ${!!signupsSheet}`); }
    const eventData = eventsSheet.getDataRange().getValues().find(row => row[0] === eventId);
    if (!eventData) {
      writeLog('addSignup', eventId, userName, position, '失敗', '找不到活動');
      return { status: 'error', message: '找不到此活動。' };
    }
    const eventDatePart = new Date(eventData[2]);
    if (isNaN(eventDatePart.getTime())) {
      writeLog('addSignup', eventId, userName, position, '失敗', `活動日期格式無效: ${eventData[2]}`);
      return { status: 'error', message: `活動 (${eventData[1]}) 的日期格式無效，請管理者檢查。` };
    }
    const endTimeStr = padTime(eventData[4] || '23:59');
    if (!/^\d\d:\d\d$/.test(endTimeStr)) {
      writeLog('addSignup', eventId, userName, position, '失敗', `結束時間格式錯誤: ${endTimeStr}`);
      return { status: 'error', message: `活動 (${eventData[1]}) 的結束時間格式錯誤，請管理者檢查。應為 HH:mm 格式。` };
    }
    const [hours, minutes] = endTimeStr.split(':').map(Number);
    const eventEndDateTime = new Date(eventDatePart.getFullYear(), eventDatePart.getMonth(), eventDatePart.getDate(), hours, minutes, 0);
    if (new Date() > eventEndDateTime) {
      writeLog('addSignup', eventId, userName, position, '失敗', '活動已結束');
      return { status: 'error', message: '此活動已結束，無法報名。' };
    }
    const currentSignups = signupsSheet.getDataRange().getValues().filter(row => row[1] === eventId);
    if (currentSignups.some(row => row[2] === userName)) {
      writeLog('addSignup', eventId, userName, position, '失敗', '重複報名');
      return { status: 'error', message: '您已經報名過此活動了。' };
    }
    const existingPositionHolder = currentSignups.find(row => row[4] === position);
    if (existingPositionHolder) {
      writeLog('addSignup', eventId, userName, position, '確認備援', `崗位已被 ${existingPositionHolder[2]} 報名`);
      return { status: 'confirm_backup', message: `此崗位目前由 [${existingPositionHolder[2]}] 報名，您要改為報名備援嗎？` };
    }
    const maxAttendees = eventData[5];
    if (currentSignups.length >= maxAttendees) {
      writeLog('addSignup', eventId, userName, position, '失敗', '人數已額滿');
      return { status: 'error', message: '活動總人數已額滿。' };
    }
    const newSignupId = 'su' + new Date().getTime();
    const newRowData = [newSignupId, eventId, userName, new Date(), position];
    writeLog('addSignup', eventId, userName, position, '準備寫入', `資料: [${newRowData.join(', ')}]`);
    signupsSheet.appendRow(newRowData);
    SpreadsheetApp.flush(); 
    writeLog('addSignup', eventId, userName, position, '寫入完成', '已執行 appendRow 和 flush');
    const lastRowValues = signupsSheet.getRange(signupsSheet.getLastRow(), 1, 1, 5).getValues()[0];
    if (lastRowValues[0] === newSignupId && lastRowValues[2] === userName) {
      writeLog('addSignup', eventId, userName, position, '成功', '寫入並驗證成功');
      return { status: 'success', message: '報名成功！' };
    } else {
      const errorDetail = `寫入驗證失敗！預期 [${newSignupId}, ${userName}], 實際 [${lastRowValues[0]}, ${lastRowValues[2]}]`;
      writeLog('addSignup', eventId, userName, position, '失敗', errorDetail);
      if (lastRowValues[0] === newSignupId) { signupsSheet.deleteRow(signupsSheet.getLastRow()); }
      throw new Error('報名資料寫入驗證失敗，請稍後再試。');
    }
  } catch(err) { 
    writeLog('addSignup', eventId, userName, position, '程式碼錯誤', err.message);
    throw new Error('報名時發生後端錯誤，請聯繫管理員。');
  } finally { 
    lock.releaseLock();
  }
}

function addBackupSignup(eventId, userName, position) {
  writeLog('addBackupSignup', eventId, userName, position, '啟動', '開始處理備援報名');
  if (!eventId || !userName || !position) {
    writeLog('addBackupSignup', eventId, userName, position, '失敗', '缺少必要資訊');
    return { status: 'error', message: '缺少必要資訊。' };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    writeLog('addBackupSignup', eventId, userName, position, '失敗', '系統忙碌，無法取得鎖');
    return { status: 'error', message: '系統忙碌中，請稍後再試。' };
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    if (!signupsSheet) { throw new Error(`工作表不存在! Signups: ${!!signupsSheet}`); }
    const newSignupId = 'su' + new Date().getTime();
    const backupPosition = `${position} (備援)`;
    const newRowData = [newSignupId, eventId, userName, new Date(), backupPosition];
    writeLog('addBackupSignup', eventId, userName, position, '準備寫入', `資料: [${newRowData.join(', ')}]`);
    signupsSheet.appendRow(newRowData);
    SpreadsheetApp.flush();
    writeLog('addBackupSignup', eventId, userName, position, '寫入完成', '已執行 appendRow 和 flush');
    const lastRowValues = signupsSheet.getRange(signupsSheet.getLastRow(), 1, 1, 5).getValues()[0];
    if (lastRowValues[0] === newSignupId && lastRowValues[2] === userName) {
      writeLog('addBackupSignup', eventId, userName, position, '成功', '備援寫入並驗證成功');
      return { status: 'success' };
    } else {
      const errorDetail = `備援寫入驗證失敗！預期 [${newSignupId}, ${userName}], 實際 [${lastRowValues[0]}, ${lastRowValues[2]}]`;
      writeLog('addBackupSignup', eventId, userName, position, '失敗', errorDetail);
      if (lastRowValues[0] === newSignupId) { signupsSheet.deleteRow(signupsSheet.getLastRow()); }
      throw new Error('備援報名資料寫入驗證失敗，請稍後再試。');
    }
  } catch (err) {
    writeLog('addBackupSignup', eventId, userName, position, '程式碼錯誤', err.message);
    throw new Error('備援報名時發生後端錯誤。');
  } finally {
    lock.releaseLock();
  }
}

function removeSignup(eventId, userName) {
  writeLog('removeSignup', eventId, userName, '', '啟動', '開始處理刪除請求');
  if (!eventId || !userName) {
    writeLog('removeSignup', eventId, userName, '', '失敗', '缺少事件ID或姓名');
    return { status: 'error', message: '缺少事件ID或姓名。' };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    writeLog('removeSignup', eventId, userName, '', '失敗', '系統忙碌，無法取得鎖');
    return { status: 'error', message: '系統忙碌中，無法取消報名，請稍後再試。' };
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupsSheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
    if (!signupsSheet) { throw new Error(`工作表不存在! Signups: ${!!signupsSheet}`); }
    const signupsData = signupsSheet.getDataRange().getValues();
    let rowIndexToDelete = -1;
    let originalData = '';
    for (let i = signupsData.length - 1; i >= 1; i--) { 
      if (signupsData[i][1] === eventId && signupsData[i][2] === userName) { 
        rowIndexToDelete = i + 1;
        originalData = signupsData[i].join(', ');
        break; 
      } 
    }
    if (rowIndexToDelete > -1) {
      writeLog('removeSignup', eventId, userName, '', '準備刪除', `找到匹配紀錄於行 ${rowIndexToDelete}。資料: ${originalData}`);
      signupsSheet.deleteRow(rowIndexToDelete);
      SpreadsheetApp.flush();
      writeLog('removeSignup', eventId, userName, '', '成功', `已刪除行 ${rowIndexToDelete}`);
      return { status: 'success', message: '已為您取消報名。' };
    } else {
      writeLog('removeSignup', eventId, userName, '', '失敗', '在工作表中找不到匹配的報名紀錄');
      return { status: 'error', message: '找不到您的報名紀錄。' };
    }
  } catch(err) {
    writeLog('removeSignup', eventId, userName, '', '程式碼錯誤', err.message);
    throw new Error('取消報名時發生錯誤。');
  } finally {
    lock.releaseLock();
  }
}

// ====================================================================
//  其他函式 (保持不變)
// ====================================================================
function padTime(timeStr) { let trimmedTime = String(timeStr || '').trim(); if (trimmedTime.includes('GMT') && trimmedTime.includes(':')) { try { const dateObj = new Date(trimmedTime); if (!isNaN(dateObj.getTime())) { const scriptTimeZone = Session.getScriptTimeZone(); return Utilities.formatDate(dateObj, scriptTimeZone, 'HH:mm'); } } catch(e) {} } if (/^\d:\d\d$/.test(trimmedTime)) { return '0' + trimmedTime; } return trimmedTime; }
function getMasterData() { if (EVENTS_CSV_URL.includes("在此貼上") || SIGNUPS_CSV_URL.includes("在此貼上")) { throw new Error("後端程式碼尚未設定 Events 和 Signups 的 CSV 網址。"); } try { const cacheBuster = '&_t=' + new Date().getTime(); const requests = [ { url: EVENTS_CSV_URL + cacheBuster, muteHttpExceptions: true }, { url: SIGNUPS_CSV_URL + cacheBuster, muteHttpExceptions: true } ]; const responses = UrlFetchApp.fetchAll(requests); const eventsResponse = responses[0]; const signupsResponse = responses[1]; if (eventsResponse.getResponseCode() !== 200) throw new Error(`無法獲取 Events CSV 資料。錯誤碼: ${eventsResponse.getResponseCode()}`); if (signupsResponse.getResponseCode() !== 200) throw new Error(`無法獲取 Signups CSV 資料。錯誤碼: ${signupsResponse.getResponseCode()}`); const eventsData = Utilities.parseCsv(eventsResponse.getContentText()); const signupsData = Utilities.parseCsv(signupsResponse.getContentText()); eventsData.shift(); signupsData.shift(); const scriptTimeZone = Session.getScriptTimeZone(); const eventsMap = new Map(); eventsData.forEach(row => { const eventId = row[0]; if (eventId && row[2]) { try { const dateObj = new Date(row[2]); if (isNaN(dateObj.getTime())) { console.warn(`無效日期格式，跳過活動 ID ${eventId}: ${row[2]}`); return; } eventsMap.set(eventId, { title: row[1] || '未知事件', dateString: Utilities.formatDate(dateObj, scriptTimeZone, 'yyyy/MM/dd'), dateObj: dateObj, startTime: padTime(row[3]), endTime: padTime(row[4]), maxAttendees: parseInt(row[5], 10) || 999, description: row[7] || '' }); } catch(e) { console.error(`處理活動 ID ${eventId} 的日期時發生錯誤: ${row[2]}. 錯誤: ${e.message}`); } } }); return { eventsData, signupsData, eventsMap }; } catch (err) { console.error(`getMasterData (CSV) 失敗: ${err.message}`, err.stack); throw new Error(`從 CSV 網址讀取資料時發生錯誤: ${err.message}`); } }
function getEventsAndSignups() { try { const { eventsData, signupsData } = getMasterData(); const scriptTimeZone = Session.getScriptTimeZone(); const signupsByEventId = new Map(); signupsData.forEach(row => { const eventId = row[1]; if (!eventId) return; if (!signupsByEventId.has(eventId)) { signupsByEventId.set(eventId, []); } signupsByEventId.get(eventId).push({ user: row[2], position: row[4] }); }); return eventsData.map(row => { const [eventId, title, rawDateString, rawStartTime, rawEndTime, maxAttendees, positions, description] = row; if (!eventId || !title || !rawDateString || !rawStartTime) return null; const paddedStartTime = padTime(rawStartTime); const paddedEndTime = padTime(rawEndTime); const datePart = Utilities.formatDate(new Date(rawDateString), scriptTimeZone, "yyyy-MM-dd"); const signups = signupsByEventId.get(eventId) || []; const signupCount = signups.length; const parsedMaxAttendees = parseInt(maxAttendees, 10) || 999; const isFull = signupCount >= parsedMaxAttendees; let eventColor, textColor; if (isFull) { eventColor = '#e74c3c'; textColor = 'white'; } else if (signupCount > 0) { eventColor = '#0d6efd'; textColor = 'white'; } else { eventColor = '#adb5bd'; textColor = '#212529'; } return { id: eventId, title: title, start: `${datePart}T${paddedStartTime}`, end: paddedEndTime ? `${datePart}T${paddedEndTime}` : null, backgroundColor: eventColor, borderColor: eventColor, textColor: textColor, extendedProps: { full_title: title, description: description, maxAttendees: parsedMaxAttendees, signups, positions: (positions || '').split(',').map(p => p.trim()).filter(p => p), startTime: paddedStartTime, endTime: paddedEndTime } }; }).filter(e => e !== null); } catch(err) { console.error("getEventsAndSignups 失敗:", err.message, err.stack); throw err; } }
function getUnifiedSignups(searchText, startDateStr, endDateStr) { try { const { signupsData, eventsMap } = getMasterData(); const searchLower = searchText ? searchText.toLowerCase() : ''; const dayMap = ['日', '一', '二', '三', '四', '五', '六']; let filteredSignups = signupsData.filter(row => { const eventId = row[1]; const eventDetail = eventsMap.get(eventId); if (!eventDetail || !eventDetail.dateObj || isNaN(eventDetail.dateObj.getTime())) { return false; } const eventDate = eventDetail.dateObj; const start = startDateStr ? new Date(startDateStr) : null; const end = endDateStr ? new Date(endDateStr) : null; if (end) end.setHours(23, 59, 59, 999); if (start && eventDate < start) return false; if (end && eventDate > end) return false; if (searchLower) { const user = String(row[2] || '').toLowerCase(); const position = String(row[4] || '').toLowerCase(); const eventTitle = String(eventDetail.title || '').toLowerCase(); const eventDescription = String(eventDetail.description || '').toLowerCase(); if (!user.includes(searchLower) && !position.includes(searchLower) && !eventTitle.includes(searchLower) && !eventDescription.includes(searchLower)) { return false; } } return true; }); const mappedSignups = filteredSignups.map(row => { const eventId = row[1]; const eventDetail = eventsMap.get(eventId) || { title: '未知事件', dateString: '日期無效'}; let eventDayOfWeek = eventDetail.dateObj ? dayMap[eventDetail.dateObj.getDay()] : ''; const rawTimestamp = row[3] || ''; return { signupId: row[0], eventId: eventId, eventTitle: eventDetail.title, eventDate: eventDetail.dateString, eventDayOfWeek: eventDayOfWeek, user: row[2], timestamp: rawTimestamp, position: row[4] }; }); mappedSignups.sort((a, b) => new Date(a.eventDate).getTime() - new Date(b.eventDate).getTime()); return mappedSignups; } catch(err) { console.error("getUnifiedSignups 失敗:", err.message, err.stack); throw err; } }
function getStatsData(startDateStr, endDateStr) { try { const allSignups = getUnifiedSignups('', startDateStr, endDateStr); if (!allSignups || allSignups.length === 0) { return { labels: [], data: [], fullDetails: [] }; } const { eventsMap } = getMasterData(); const statsByEventId = {}; allSignups.forEach(signup => { const eventId = signup.eventId; if (!statsByEventId[eventId]) { const eventInfoFromMap = eventsMap.get(eventId) || {}; const eventInfo = { title: eventInfoFromMap.title || signup.eventTitle, date: eventInfoFromMap.dateString || signup.eventDate, dayOfWeek: signup.eventDayOfWeek, startTime: eventInfoFromMap.startTime || '', endTime: eventInfoFromMap.endTime || '', maxAttendees: eventInfoFromMap.maxAttendees || 999 }; statsByEventId[eventId] = { count: 0, signups: [], eventInfo: eventInfo, label: `${eventInfo.title} (${eventInfo.date})` }; } statsByEventId[eventId].count++; statsByEventId[eventId].signups.push({ user: signup.user, position: signup.position }); }); const processedData = Object.values(statsByEventId); processedData.sort((a, b) => new Date(a.eventInfo.date).getTime() - new Date(b.eventInfo.date).getTime()); const labels = processedData.map(item => item.label); const data = processedData.map(item => item.count); const fullDetails = processedData.map(item => { item.signups.sort((a,b) => a.position.localeCompare(b.position)); return { eventInfo: item.eventInfo, signups: item.signups }; }); return { labels, data, fullDetails }; } catch (err) { console.error("getStatsData 失敗:", err.message, err.stack); throw err; } }
function getSignupsAsTextForToday() { const scriptTimeZone = Session.getScriptTimeZone(); const today = new Date(); const todayStr = Utilities.formatDate(today, scriptTimeZone, 'yyyy-MM-dd'); const result = getSignupsAsText(todayStr, todayStr); if (result && result.status === 'success' && result.text) { const dateDisplay = Utilities.formatDate(today, scriptTimeZone, 'MM/dd'); result.text = result.text.replace(`日期：${todayStr} ~ ${todayStr}`, `日期：今天 (${dateDisplay})`); } return result; }
function getSignupsAsTextForTomorrow() { const scriptTimeZone = Session.getScriptTimeZone(); const tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate() + 1); const tomorrowStr = Utilities.formatDate(tomorrow, scriptTimeZone, 'yyyy-MM-dd'); const result = getSignupsAsText(tomorrowStr, tomorrowStr); if (result && result.status === 'success' && result.text) { const dateDisplay = Utilities.formatDate(tomorrow, scriptTimeZone, 'MM/dd'); result.text = result.text.replace(`日期：${tomorrowStr} ~ ${tomorrowStr}`, `日期：明天 (${dateDisplay})`); } return result; }
function getSignupsAsText(startDateStr, endDateStr) { const allSignups = getUnifiedSignups('', startDateStr, endDateStr); if (!allSignups || allSignups.length === 0) { return { status: 'nodata' }; } const eventsGroup = {}; allSignups.forEach(signup => { const key = `${signup.eventDate} ${signup.eventTitle}`; if (!eventsGroup[key]) { eventsGroup[key] = []; } eventsGroup[key].push(`- ${signup.position}: ${signup.user}`); }); let formattedText = `【AV放送立願報名記錄】\n日期：${startDateStr} ~ ${endDateStr}\n----------\n\n`; Object.keys(eventsGroup).sort().forEach(eventKey => { formattedText += `【${eventKey}】\n`; formattedText += eventsGroup[eventKey].join('\n'); formattedText += '\n\n'; }); return { status: 'success', text: formattedText }; }
function createTempSheetAndExport(startDateStr, endDateStr) { const ss = SpreadsheetApp.openById(SHEET_ID); const spreadsheetId = ss.getId(); let tempSheet = null; let tempFile = null; try { const tempSheetName = "匯出報表_" + new Date().getTime(); tempSheet = ss.insertSheet(tempSheetName); const data = getUnifiedSignups('', startDateStr, endDateStr); if (!data || data.length === 0) { ss.deleteSheet(tempSheet); return { status: 'nodata', message: '在選定的日期範圍內沒有任何報名記錄。' }; } const title = `AV放送立願報名記錄 (${startDateStr || '所有'} - ${endDateStr || '所有'})`; tempSheet.getRange("A1:E1").merge().setValue(title).setFontWeight("bold").setFontSize(15).setHorizontalAlignment('center'); const headers = ["活動日期", "活動", "崗位", "報名者", "報名時間"]; const fields = ["eventDate", "eventTitle", "position", "user", "timestamp"]; tempSheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3").setFontSize(15).setHorizontalAlignment('center'); const outputData = data.map(row => fields.map(field => row[field] || "")); if (outputData.length > 0) { tempSheet.getRange(3, 1, outputData.length, headers.length).setValues(outputData).setFontSize(15).setHorizontalAlignment('left'); } tempSheet.autoResizeColumns(1, 5); SpreadsheetApp.flush(); const blob = Drive.Files.export(spreadsheetId, MimeType.MICROSOFT_EXCEL, { gid: tempSheet.getSheetId(), alt: 'media' }); blob.setName(`AV立願報名記錄_${new Date().toISOString().slice(0,10)}.xlsx`); tempFile = DriveApp.createFile(blob); tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); const downloadUrl = tempFile.getWebContentLink(); scheduleDeletion(tempFile.getId()); ss.deleteSheet(tempSheet); return { status: 'success', url: downloadUrl, fileId: tempFile.getId() }; } catch (e) { console.error("createTempSheetAndExport 失敗: " + e.toString()); if (tempFile) { try { DriveApp.getFileById(tempFile.getId()).setTrashed(true); } catch (f) { console.error("無法刪除臨時檔案: " + f.toString()); } } if (tempSheet && ss.getSheetByName(tempSheet.getName())) { ss.deleteSheet(tempSheet); } throw new Error('匯出時發生錯誤: ' + e.toString()); } }
function scheduleDeletion(fileId) { if (!fileId) return; try { const trigger = ScriptApp.newTrigger('triggeredDeleteHandler').timeBased().after(24 * 60 * 60 * 1000).create(); PropertiesService.getScriptProperties().setProperty(trigger.getUniqueId(), fileId); } catch (e) { console.error(`排程刪除檔案 '${fileId}' 時失敗: ${e.toString()}`); } }
function triggeredDeleteHandler(e) { const triggerId = e.triggerUid; const scriptProperties = PropertiesService.getScriptProperties(); const fileId = scriptProperties.getProperty(triggerId); if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) { console.error(`刪除檔案 ${fileId} 失敗: ${err.toString()}`); } scriptProperties.deleteProperty(triggerId); } const allTriggers = ScriptApp.getProjectTriggers(); for (let i = 0; i < allTriggers.length; i++) { if (allTriggers[i].getUniqueId() === triggerId) { ScriptApp.deleteTrigger(allTriggers[i]); break; } } }
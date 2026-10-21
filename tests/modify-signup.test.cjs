const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const backendSource = fs.readFileSync(path.join(root, 'code.gs'), 'utf8');
const frontendSource = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

// Run the production functions with in-memory Sheets and a simulated JSONP transport.
function extractFunction(source, name, indent = '') {
  const declarations = new RegExp(`^${indent}(?:async )?function \\w+\\(`, 'gm');
  const matches = [...source.matchAll(declarations)];
  const index = matches.findIndex(match => match[0].includes(`function ${name}(`));
  assert.notEqual(index, -1, `Missing function: ${name}`);
  return source.slice(matches[index].index, matches[index + 1]?.index ?? source.length);
}

function createHarness(options = {}) {
  const rows = [
    ['SignupID', 'EventID', 'User', 'Timestamp', 'Position'],
    ['signup-test', 'event-test', '測試使用者', new Date('2026-10-21T06:00:00Z'), '原崗位']
  ];
  if (options.positionTaken) rows.push(['other-signup', 'event-test', '其他使用者', '', '新崗位']);
  const state = { rows, alerts: [], requests: [], writes: [], refreshes: 0, invalidations: 0, releases: 0 };
  const eventsData = [['event-test', '測試活動', '2026-10-21', '15:00', '16:00', 3, '原崗位,新崗位', '']];
  const quietConsole = { log() {}, error() {} };
  const back = vm.createContext({
    console: quietConsole, Date, BACKEND_VERSION: 'test', SHEET_ID: 'test-sheet', SIGNUPS_SHEET_NAME: 'Signups',
    Session: { getScriptTimeZone: () => 'Asia/Taipei' },
    Utilities: { formatDate: (_date, _zone, format) => format === 'yyyy-MM-dd' ? '2026-10-21' : '2026/10/21 14:00:00' },
    padTime: value => value,
    getMasterData: () => ({ eventsData, signupsData: rows.slice(1) }),
    logOperation() {},
    invalidateMasterDataCache_() { state.invalidations++; },
    LockService: { getScriptLock: () => ({
      tryLock: () => !options.busy,
      releaseLock() { state.releases++; }
    }) },
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: () => ({
        getDataRange: () => ({ getValues: () => rows.map(row => [...row]) }),
        getRange: (row, column) => ({ setValue(value) {
          if (options.writeFails) throw new Error('Simulated write failure');
          rows[row - 1][column - 1] = value;
          state.writes.push({ row, column, value });
        } })
      }) }),
      flush() {}
    },
    ContentService: {
      MimeType: { JAVASCRIPT: 'javascript' },
      createTextOutput: value => ({ value, setMimeType() { return this; } })
    }
  });
  vm.runInContext(['buildEventsAndSignups_', 'modifySignup', 'doGet']
    .map(name => extractFunction(backendSource, name)).join('\n'), back);

  function request(functionName, params) {
    const result = back.doGet({ parameter: { payload: JSON.stringify({ functionName, params }), callback: 'cb' } });
    return JSON.parse(result.value.slice(3, -1));
  }

  const front = {
    console: quietConsole,
    allCalendarEvents: back.buildEventsAndSignups_(),
    DATE_FMT_CA_TW: new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei' }),
    isMobile: false, SCRIPT_URL: 'https://example.invalid/exec',
    showLoading() {}, hideLoading() {},
    setTimeout: () => 1, clearTimeout() {},
    lastQuery: { params: {} },
    invalidateUnifiedSignupsCache() {},
    executeUnifiedSignupQuery: async (...args) => {
      assert.equal(args.at(-1), true);
      state.refreshes++;
      front.allCalendarEvents = back.buildEventsAndSignups_();
    },
    setupCalendarTab() {},
    Swal: { fire: async (...args) => {
      if (typeof args[0] === 'object') return { value: options.cancel ? undefined : (options.selection || '新崗位') };
      state.alerts.push(args);
    } }
  };
  front.getCalendarEventById = id => front.allCalendarEvents.find(event => event.id === id);
  front.window = front;
  front.document = {
    createElement: () => ({}), getElementById: () => null,
    head: { appendChild(script) {
      const url = new URL(script.src);
      const payload = JSON.parse(url.searchParams.get('payload'));
      state.requests.push(payload);
      const response = Object.hasOwn(options, 'response') ? options.response : request(payload.functionName, payload.params);
      front[url.searchParams.get('callback')](response);
    } }
  };
  vm.createContext(front);
  vm.runInContext(['getUnifiedSignups', 'callAppsScript', 'promptToModifySignup']
    .map(name => extractFunction(frontendSource, name, '        ')).join('\n'), front);
  return { state, front, request };
}

test('backend and inline frontend JavaScript parse', () => {
  new vm.Script(backendSource);
  for (const match of frontendSource.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) {
    if (match[1].trim()) new vm.Script(match[1]);
  }
});

test('query → JSONP → sheet write → refreshed query preserves the signup ID and new position', async () => {
  const { state, front } = createHarness();
  const [signup] = await front.getUnifiedSignups('', '', '');
  assert.equal(signup.signupId, 'signup-test');
  assert.equal(signup.timestamp, '2026/10/21 14:00:00');
  await front.promptToModifySignup(signup);
  assert.deepEqual(state.requests[0].params, { signupId: 'signup-test', newPosition: '新崗位' });
  assert.deepEqual(state.writes, [{ row: 2, column: 5, value: '新崗位' }]);
  assert.equal(state.invalidations, 1);
  assert.equal(state.releases, 1);
  assert.equal(state.refreshes, 1);
  assert.deepEqual(state.alerts, [['修改成功', '崗位已成功修改！', 'success']]);
  const [refreshed] = await front.getUnifiedSignups('', '', '');
  assert.equal(refreshed.signupId, 'signup-test');
  assert.equal(refreshed.position, '新崗位');
});

for (const [name, options, message] of [
  ['occupied position', { positionTaken: true }, '已被其他人報名'],
  ['lock unavailable', { busy: true }, '系統忙碌中'],
  ['sheet write failure', { writeFails: true }, '後端錯誤']
]) {
  test(`${name} reports failure without success or refreshing`, async () => {
    const { state, front } = createHarness(options);
    const [signup] = await front.getUnifiedSignups('', '', '');
    await front.promptToModifySignup(signup);
    assert.equal(state.writes.length, 0);
    assert.equal(state.refreshes, 0);
    assert.equal(state.rows[1][4], '原崗位');
    assert.equal(state.alerts.length, 1);
    assert.equal(state.alerts[0][0], '修改失敗');
    assert.equal(state.alerts[0][2], 'error');
    assert.ok(state.alerts[0][1].includes(message));
    if (!options.busy) assert.equal(state.releases, 1);
  });
}

test('missing ID is rejected by the API before writing', () => {
  const { state, request } = createHarness();
  const response = request('modifySignup', { newPosition: '新崗位' });
  assert.equal(response.status, 'error');
  assert.equal(response.message, '缺少報名ID或新崗位資訊。');
  assert.equal(state.writes.length, 0);
});

test('deleted signup is rejected without writing', () => {
  const { state, request } = createHarness();
  const response = request('modifySignup', { signupId: 'deleted-signup', newPosition: '新崗位' });
  assert.equal(response.status, 'error');
  assert.equal(state.writes.length, 0);
  assert.equal(state.releases, 1);
});

test('incomplete cached signup is blocked without sending a request', async () => {
  const { state, front } = createHarness();
  await front.promptToModifySignup({ eventId: 'event-test', user: '測試使用者', position: '原崗位' });
  assert.equal(state.requests.length, 0);
  assert.equal(state.alerts[0][0], '無法修改');
  assert.equal(state.alerts[0][2], 'error');
});

for (const [name, response] of [
  ['legacy nested error', { status: 'success', data: { status: 'error', message: '缺少報名ID或新崗位資訊。' } }],
  ['missing operation result', { status: 'success' }]
]) {
  test(`${name} cannot display success`, async () => {
    const { state, front } = createHarness({ response });
    const [signup] = await front.getUnifiedSignups('', '', '');
    await front.promptToModifySignup(signup);
    assert.equal(state.alerts[0][0], '修改失敗');
    assert.equal(state.alerts[0][2], 'error');
    assert.equal(state.refreshes, 0);
  });
}

for (const [name, options] of [['cancel', { cancel: true }], ['unchanged position', { selection: '原崗位' }]]) {
  test(`${name} sends no mutation`, async () => {
    const { state, front } = createHarness(options);
    const [signup] = await front.getUnifiedSignups('', '', '');
    await front.promptToModifySignup(signup);
    assert.equal(state.requests.length, 0);
    assert.equal(state.alerts.length, 0);
  });
}

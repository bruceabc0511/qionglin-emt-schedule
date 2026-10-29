/**
 * 芎林義消救護分隊 協勤排班系統 — Google Apps Script 後端
 * ------------------------------------------------------------
 * 部署方式：請參考專案根目錄 README.md
 *  1. 建立一份 Google 試算表 → 擴充功能 → Apps Script
 *  2. 貼上本檔與 appsscript.json
 *  3. 執行 setup() 一次（建立工作表、匯入人員、設定每月20號自動產生次月班表觸發器）
 *  4. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我；存取權：所有人）
 *  5. 將部署網址填入 js/config.js 的 API_URL
 */

// ======================= 系統設定 =======================
const CONFIG = {
  UNIT_NAME: '芎林義消救護分隊',
  SHIFT_HOURS: 4,                       // 每班時數（4 小時）
  STEP_HOURS: 1,                        // 起始時間間隔（1 小時一格：08-09、09-10…）
  OPEN_NEXT_MONTH_DAY: 20,              // 每月幾號自動開放選填次月班表（20 號）
  TIER: {
    1: { name: '一車', start: 8, end: 22 },
    2: { name: '二車', start: 8, end: 18 },
  },
  REQUIRED_T1_PER_DAY: 3,               // 每日一車應排班數（排班率分母）
  RATE_TARGET: 90,                      // 每月協勤率達標門檻（%）
  DEFAULT_PIN: '0000',                  // 新成員預設 PIN（首次登入強制變更）
  TOKEN_TTL_SEC: 6 * 60 * 60,           // 登入有效時間（最長 6 小時）
  ALLOW_PAST_EDIT: false,               // 是否允許一般成員修改過去日期的班
  MAX_LOGIN_FAIL: 5,                    // 連續登入失敗次數上限
};

const SHEET_MEMBERS = '人員清單';
const SHEET_SHIFTS = '排班紀錄';
const MEMBER_HEADERS = ['姓名', 'PIN雜湊', '角色', '狀態', '需變更PIN', '建立時間'];
const SHIFT_HEADERS = ['ID', '日期', '車別', '開始', '結束', '姓名', '登記時間', '異動註記'];

// 由 EXCEL「芎林救護義消服勤預定表」匯入的初始人員
const INITIAL_MEMBERS = [
  '賴致平', '林心彤', '許世良', '賴志昀', '官正軒', '蔡宗佑', '龍賢銨', '林月娥',
  '陳容清', '張謙加', '陳長鎧', '吳宛諭', '彭育琳', '羅艾婷', '陳佩秀', '陳姿璇',
  '官振瑋', '劉兆華', '劉劭倫', '閻欣慈', '謝桂湘', '鍾桂蓉', '潘麗芬', '洪如意',
];

// ======================= 初始化 =======================
/** 請在 Apps Script 編輯器中手動執行一次 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid());

  let mem = ss.getSheetByName(SHEET_MEMBERS);
  if (!mem) {
    mem = ss.insertSheet(SHEET_MEMBERS);
    mem.getRange(1, 1, 1, MEMBER_HEADERS.length).setValues([MEMBER_HEADERS])
      .setFontWeight('bold').setBackground('#66B3FF').setFontColor('#ffffff');
    mem.setFrozenRows(1);
    const now = new Date();
    const rows = [['管理員', hashPin_('管理員', '9999'), 'admin', '啟用', 'Y', now]];
    INITIAL_MEMBERS.forEach(n => rows.push([n, hashPin_(n, CONFIG.DEFAULT_PIN), 'member', '啟用', 'Y', now]));
    mem.getRange(2, 1, rows.length, MEMBER_HEADERS.length).setValues(rows);
    mem.autoResizeColumns(1, MEMBER_HEADERS.length);
  }

  let sh = ss.getSheetByName(SHEET_SHIFTS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SHIFTS);
    sh.getRange(1, 1, 1, SHIFT_HEADERS.length).setValues([SHIFT_HEADERS])
      .setFontWeight('bold').setBackground('#66B3FF').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:H').setNumberFormat('@');   // 全部以純文字儲存，避免日期被自動轉換
  } else if (sh.getLastColumn() < SHIFT_HEADERS.length) {
    sh.getRange(1, 1, 1, SHIFT_HEADERS.length).setValues([SHIFT_HEADERS])
      .setFontWeight('bold').setBackground('#66B3FF').setFontColor('#ffffff');
  }

  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  setupAutoOpenTrigger_();

  const t = new Date();
  renderMonthSheet_(t.getFullYear(), t.getMonth() + 1);
  if (t.getDate() >= CONFIG.OPEN_NEXT_MONTH_DAY) {
    const nx = new Date(t.getFullYear(), t.getMonth() + 1, 1);
    renderMonthSheet_(nx.getFullYear(), nx.getMonth() + 1);
  }
  Logger.log('初始化完成。管理員帳號：管理員 / PIN 9999；成員預設 PIN：' + CONFIG.DEFAULT_PIN);
}

/** 建立每日定時觸發器：每月 20 日自動建立次月排班表工作表 */
function setupAutoOpenTrigger_() {
  const fnName = 'dailyAutoOpenCheck_';
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === fnName);
  if (!exists) {
    ScriptApp.newTrigger(fnName).timeBased().everyDays(1).atHour(0).create();
  }
}

function dailyAutoOpenCheck_() {
  const now = new Date();
  if (now.getDate() >= CONFIG.OPEN_NEXT_MONTH_DAY) {
    const nx = new Date(now.getFullYear(), now.getMonth() + 1, 1);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(monthSheetName_(nx.getFullYear(), nx.getMonth() + 1))) {
      renderMonthSheet_(nx.getFullYear(), nx.getMonth() + 1);
    }
  }
}

/** 試算表選單：手動重新產生月表 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('排班系統')
    .addItem('重新產生本月排班表', 'menuRenderThisMonth_')
    .addItem('重新產生下月排班表', 'menuRenderNextMonth_')
    .addToUi();
}
function menuRenderThisMonth_() { const t = new Date(); renderMonthSheet_(t.getFullYear(), t.getMonth() + 1); }
function menuRenderNextMonth_() { const t = new Date(); t.setDate(1); t.setMonth(t.getMonth() + 1); renderMonthSheet_(t.getFullYear(), t.getMonth() + 1); }

// ======================= Web API 入口 =======================
function doGet() {
  return json_({ ok: true, data: { service: CONFIG.UNIT_NAME + ' 協勤排班 API', time: new Date() } });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const handler = ROUTES[req.action];
    if (!handler) throw new Error('未知的操作：' + req.action);
    return json_({ ok: true, data: handler(req) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

const ROUTES = {
  members: apiMembers,
  login: apiLogin,
  changePin: apiChangePin,
  getMonth: apiGetMonth,
  addShift: apiAddShift,
  updateShift: apiUpdateShift,
  deleteShift: apiDeleteShift,
  setTier2Lock: apiSetTier2Lock,
  adminListMembers: apiAdminListMembers,
  adminAddMember: apiAdminAddMember,
  adminRemoveMember: apiAdminRemoveMember,
  adminRestoreMember: apiAdminRestoreMember,
  adminResetPin: apiAdminResetPin,
};

// ======================= 帳號 / 登入 =======================
function apiMembers() {
  return readMembers_().filter(m => m.active).map(m => m.name);
}

/** 排班統計用人員清單（啟用中、非純管理帳號） */
function rosterNames_() {
  return readMembers_().filter(m => m.active && m.role !== 'admin').map(m => m.name);
}

function apiLogin(req) {
  const name = String(req.name || '').trim();
  const pin = String(req.pin || '');
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + name;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= CONFIG.MAX_LOGIN_FAIL) throw new Error('登入失敗次數過多，請 10 分鐘後再試');

  const m = readMembers_().find(x => x.name === name && x.active);
  if (!m || m.pinHash !== hashPin_(name, pin)) {
    cache.put(failKey, String(fails + 1), 600);
    throw new Error('姓名或 PIN 碼錯誤');
  }
  cache.remove(failKey);
  const token = Utilities.getUuid();
  cache.put('tok_' + token, name, CONFIG.TOKEN_TTL_SEC);
  return { token, name: m.name, role: m.role, mustChangePin: m.mustChange };
}

function apiChangePin(req) {
  const me = requireSession_(req.token);
  const newPin = String(req.newPin || '');
  if (!/^\d{4,8}$/.test(newPin)) throw new Error('PIN 碼須為 4～8 位數字');
  if (newPin === CONFIG.DEFAULT_PIN) throw new Error('請勿使用預設 PIN 碼');
  return withLock_(() => {
    const sh = getSheet_(SHEET_MEMBERS);
    const m = readMembers_().find(x => x.name === me.name);
    if (m.pinHash !== hashPin_(me.name, String(req.oldPin || ''))) throw new Error('原 PIN 碼錯誤');
    sh.getRange(m.row, 2).setValue(hashPin_(me.name, newPin));
    sh.getRange(m.row, 5).setValue('N');
    return true;
  });
}

// ======================= 二車鎖定管理 =======================
function getLockedT2Dates_(y, m) {
  const key = 'T2_LOCK_' + y + '-' + pad_(m);
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return [];
  try { return JSON.parse(raw); } catch (_) { return []; }
}

function saveLockedT2Dates_(y, m, dates) {
  const key = 'T2_LOCK_' + y + '-' + pad_(m);
  const uniq = Array.from(new Set(dates)).sort();
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(uniq));
  return uniq;
}

function apiSetTier2Lock(req) {
  requireAdmin_(req.token);
  const dates = Array.isArray(req.dates) ? req.dates.map(String) : [String(req.date || '')];
  const locked = !!req.locked;
  if (!dates.length || !/^\d{4}-\d{2}-\d{2}$/.test(dates[0])) throw new Error('日期格式錯誤');
  const [y, m] = dates[0].split('-').map(Number);
  return withLock_(() => {
    const set = new Set(getLockedT2Dates_(y, m));
    dates.forEach(d => { if (locked) set.add(d); else set.delete(d); });
    const updated = saveLockedT2Dates_(y, m, Array.from(set));
    safeRender_(dates[0]);
    return { lockedT2Dates: updated };
  });
}

// ======================= 排班 =======================
function apiGetMonth(req) {
  requireSession_(req.token);
  const y = Number(req.year), mo = Number(req.month);
  const prefix = y + '-' + pad_(mo) + '-';
  const shifts = readShifts_().filter(s => s.date.indexOf(prefix) === 0);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName(monthSheetName_(y, mo));
  return {
    year: y, month: mo,
    shifts: shifts.map(stripRow_),
    lockedT2Dates: getLockedT2Dates_(y, mo),
    members: rosterNames_(),
    sheetUrl: ss.getUrl() + (mSheet ? '#gid=' + mSheet.getSheetId() : ''),
    serverTime: Utilities.formatDate(new Date(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss"),
  };
}

function apiAddShift(req) {
  const me = requireSession_(req.token);
  const isAdmin = me.role === 'admin';
  const date = String(req.date), tier = Number(req.tier), start = Number(req.start);
  const targetName = (isAdmin && req.targetName) ? String(req.targetName).trim() : me.name;
  const adminNote = (isAdmin && targetName !== me.name) ? '管理員異動' : '';
  return withLock_(() => {
    const all = readShifts_();                       // 取得鎖之後才讀取最新資料 → 先送出者先得
    validateShift_(all, { date, tier, start, name: targetName, isAdmin });
    const rec = {
      id: Utilities.getUuid(), date, tier, start, end: start + CONFIG.SHIFT_HOURS, name: targetName,
      createdAt: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      adminNote,
    };
    const sh = getSheet_(SHEET_SHIFTS);
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 8).setNumberFormat('@')
      .setValues([[rec.id, rec.date, String(tier), String(start), String(rec.end), rec.name, rec.createdAt, rec.adminNote]]);
    SpreadsheetApp.flush();
    safeRender_(date);
    return rec;
  });
}

function apiUpdateShift(req) {
  const me = requireSession_(req.token);
  const isAdmin = me.role === 'admin';
  const tier = Number(req.tier), start = Number(req.start);
  return withLock_(() => {
    const all = readShifts_();
    const s = all.find(x => x.id === req.id);
    if (!s) throw new Error('找不到此班次，可能已被刪除');
    if (s.name !== me.name && !isAdmin) throw new Error('僅限本人或管理員可以修改排班');
    const targetName = (isAdmin && req.targetName) ? String(req.targetName).trim() : s.name;
    const isOther = isAdmin && (s.name !== me.name || targetName !== me.name);
    const adminNote = isOther ? '管理員異動' : s.adminNote;
    checkEditable_(s.date, isAdmin);
    validateShift_(all, { date: s.date, tier, start, name: targetName, excludeId: s.id, isAdmin });
    const sh = getSheet_(SHEET_SHIFTS);
    sh.getRange(s.row, 3, 1, 4).setNumberFormat('@')
      .setValues([[String(tier), String(start), String(start + CONFIG.SHIFT_HOURS), targetName]]);
    sh.getRange(s.row, 8).setNumberFormat('@').setValue(adminNote);
    SpreadsheetApp.flush();
    safeRender_(s.date);
    return Object.assign(stripRow_(s), { tier, start, end: start + CONFIG.SHIFT_HOURS, name: targetName, adminNote });
  });
}

function apiDeleteShift(req) {
  const me = requireSession_(req.token);
  const isAdmin = me.role === 'admin';
  return withLock_(() => {
    const s = readShifts_().find(x => x.id === req.id);
    if (!s) throw new Error('找不到此班次，可能已被刪除');
    if (s.name !== me.name && !isAdmin) throw new Error('僅限本人或管理員可以取消排班');
    checkEditable_(s.date, isAdmin);
    getSheet_(SHEET_SHIFTS).deleteRow(s.row);
    SpreadsheetApp.flush();
    safeRender_(s.date);
    return true;
  });
}

function validateShift_(all, p) {
  const t = CONFIG.TIER[p.tier];
  if (!t) throw new Error('車別錯誤');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.date)) throw new Error('日期格式錯誤');
  const end = p.start + CONFIG.SHIFT_HOURS;
  if (p.start < t.start || end > t.end || (p.start - t.start) % CONFIG.STEP_HOURS !== 0) {
    throw new Error(t.name + '可排時段為 ' + pad_(t.start) + ':00～' + pad_(t.end) + ':00');
  }
  checkEditable_(p.date, p.isAdmin);
  if (p.tier === 2) {
    const [y, m] = p.date.split('-').map(Number);
    if (getLockedT2Dates_(y, m).indexOf(p.date) !== -1) {
      throw new Error('當日二車時段已鎖定未開放排班');
    }
  }
  all.forEach(s => {
    if (s.id === p.excludeId || s.date !== p.date) return;
    const overlap = s.start < end && p.start < s.end;
    if (!overlap) return;
    if (s.tier === p.tier) {
      throw new Error('此時段已被「' + s.name + '」搶先登記（' + t.name + ' ' + pad_(s.start) + ':00-' + pad_(s.end) + ':00），請重新選擇');
    }
    if (s.name === p.name) throw new Error('「' + p.name + '」在同一時間已登記另一車的班');
  });
}

function checkEditable_(date, isAdmin) {
  if (isAdmin) return;
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  if (!CONFIG.ALLOW_PAST_EDIT && date < today) {
    throw new Error('已過去的日期無法再變更');
  }
  const [ty, tm, td] = today.split('-').map(Number);
  const [dy, dm] = date.split('-').map(Number);
  const monthDiff = (dy - ty) * 12 + (dm - tm);
  if (monthDiff === 1 && td < CONFIG.OPEN_NEXT_MONTH_DAY) {
    throw new Error('次月（' + dm + '月）班表將於本月 ' + CONFIG.OPEN_NEXT_MONTH_DAY + ' 日自動開放選填');
  }
  if (monthDiff > 1) {
    throw new Error(dm + '月班表尚未開放選填（每月 ' + CONFIG.OPEN_NEXT_MONTH_DAY + ' 日開放次月班表）');
  }
}

// ======================= 管理員：人員管理 =======================
function apiAdminListMembers(req) {
  requireAdmin_(req.token);
  return readMembers_().map(m => ({ name: m.name, role: m.role, active: m.active, mustChangePin: m.mustChange }));
}

function apiAdminAddMember(req) {
  requireAdmin_(req.token);
  const name = String(req.name || '').trim();
  const role = req.role === 'admin' ? 'admin' : 'member';
  const pin = String(req.pin || CONFIG.DEFAULT_PIN);
  if (!name || name.length > 20) throw new Error('請輸入正確的姓名');
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN 碼須為 4～8 位數字');
  return withLock_(() => {
    const exist = readMembers_().find(m => m.name === name);
    const sh = getSheet_(SHEET_MEMBERS);
    if (exist) {
      if (exist.active) throw new Error('「' + name + '」已在人員清單中');
      sh.getRange(exist.row, 2, 1, 4).setValues([[hashPin_(name, pin), role, '啟用', 'Y']]);
    } else {
      sh.appendRow([name, hashPin_(name, pin), role, '啟用', 'Y', new Date()]);
    }
    return true;
  });
}

function apiAdminRemoveMember(req) {
  const me = requireAdmin_(req.token);
  if (req.name === me.name) throw new Error('不可刪除自己的帳號');
  return withLock_(() => {
    const m = readMembers_().find(x => x.name === req.name);
    if (!m) throw new Error('查無此人員');
    // 軟刪除：保留歷史排班紀錄，僅停用帳號
    getSheet_(SHEET_MEMBERS).getRange(m.row, 4).setValue('停用');
    return true;
  });
}

function apiAdminRestoreMember(req) {
  requireAdmin_(req.token);
  return withLock_(() => {
    const m = readMembers_().find(x => x.name === req.name);
    if (!m) throw new Error('查無此人員');
    getSheet_(SHEET_MEMBERS).getRange(m.row, 4).setValue('啟用');
    return true;
  });
}

function apiAdminResetPin(req) {
  requireAdmin_(req.token);
  const pin = String(req.pin || CONFIG.DEFAULT_PIN);
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN 碼須為 4～8 位數字');
  return withLock_(() => {
    const m = readMembers_().find(x => x.name === req.name);
    if (!m) throw new Error('查無此人員');
    getSheet_(SHEET_MEMBERS).getRange(m.row, 2, 1, 4).setValues([[hashPin_(m.name, pin), m.role, m.active ? '啟用' : '停用', 'Y']]);
    return true;
  });
}

// ======================= 月排班表（Google 試算表呈現） =======================
const COLOR = {
  title: '#66B3FF', weekend: '#F4CCCC', weekday: '#CFE2F3', label: '#DDEBF7',
  t1: '#E2EFDA', t2: '#FFF2CC', na: '#F2F2F2', locked: '#E6E9ED',
  filled1: '#A9D08E', filled2: '#FFD966', white: '#FFFFFF', pass: '#C6EFCE', warn: '#FCE4D6',
};

function monthSheetName_(y, m) { return (y - 1911) + '年' + m + '月'; }

/** 計算該月第一個星期三的日期（1~7） */
function firstWednesdayDay_(y, m) {
  const firstDow = new Date(y, m - 1, 1).getDay(); // 0=日..3=三
  return 1 + ((3 - firstDow + 7) % 7);
}

function safeRender_(date) {
  try {
    const p = date.split('-');
    renderMonthSheet_(Number(p[0]), Number(p[1]));
  } catch (err) {
    console.error('月表產生失敗：' + err);
  }
}

function renderMonthSheet_(y, m) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = monthSheetName_(y, m);
  const sh = ss.getSheetByName(name) || ss.insertSheet(name, 0);
  const prefix = y + '-' + pad_(m) + '-';
  const shifts = readShifts_().filter(s => s.date.indexOf(prefix) === 0);
  const lockedT2 = new Set(getLockedT2Dates_(y, m));
  const trainingDay = firstWednesdayDay_(y, m);

  const SLOT_COUNT = (CONFIG.TIER[1].end - CONFIG.TIER[1].start) / CONFIG.STEP_HOURS; // 14 格（08-09 ~ 21-22）
  const WEEK_BLOCK = 1 + SLOT_COUNT; // 每週 15 列（1 列日期 + 14 列時段）

  const firstDow = new Date(y, m - 1, 1).getDay();
  const days = new Date(y, m, 0).getDate();
  const weeks = Math.ceil((firstDow + days) / 7);
  const NROWS = Math.max(2 + weeks * WEEK_BLOCK, 40), NCOLS = 20;   // A~O 行事曆，Q~T 統計
  const vals = [], bgs = [], bold = [];
  for (let r = 0; r < NROWS; r++) {
    vals.push(new Array(NCOLS).fill(''));
    bgs.push(new Array(NCOLS).fill(COLOR.white));
    bold.push(new Array(NCOLS).fill('normal'));
  }
  const merges = [];   // [row, col, numRows, numCols]（0-based）

  vals[0][0] = CONFIG.UNIT_NAME + '  ' + y + ' 年  ' + m + ' 月 服勤排班表（' + (y - 1911) + '年' + m + '月）';
  merges.push([0, 0, 1, 15]);
  const WD = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'];
  vals[1][0] = '星期'; bgs[1][0] = COLOR.label;
  for (let d = 0; d < 7; d++) {
    const c = 1 + d * 2;
    vals[1][c] = WD[d];
    bgs[1][c + 1] = (d === 0 || d === 6) ? COLOR.weekend : COLOR.weekday;
    merges.push([1, c, 1, 2]);
  }

  const cellOf = {};   // 'yyyy-mm-dd' → {base, col}
  for (let w = 0; w < weeks; w++) {
    const base = 2 + w * WEEK_BLOCK;
    vals[base][0] = '時間/日期'; bgs[base][0] = COLOR.label;
    for (let i = 0; i < SLOT_COUNT; i++) {
      const h = CONFIG.TIER[1].start + i * CONFIG.STEP_HOURS;
      vals[base + 1 + i][0] = pad_(h) + '-' + pad_(h + CONFIG.STEP_HOURS);
      bgs[base + 1 + i][0] = COLOR.label;
    }
    for (let d = 0; d < 7; d++) {
      const c = 1 + d * 2;
      const dayNum = w * 7 + d - firstDow + 1;
      const valid = dayNum >= 1 && dayNum <= days;
      bgs[base][c] = bgs[base][c + 1] = (d === 0 || d === 6) ? COLOR.weekend : COLOR.weekday;
      merges.push([base, c, 1, 2]);
      if (!valid) continue;
      const dateStr = prefix + pad_(dayNum);
      vals[base][c] = dayNum === trainingDay ? dayNum + '(定訓)' : dayNum;
      cellOf[dateStr] = { base, col: c };
      const isT2Locked = lockedT2.has(dateStr);
      for (let i = 0; i < SLOT_COUNT; i++) {
        const h = CONFIG.TIER[1].start + i * CONFIG.STEP_HOURS;
        bgs[base + 1 + i][c] = COLOR.t1;
        const inT2Range = (h + CONFIG.STEP_HOURS <= CONFIG.TIER[2].end);
        bgs[base + 1 + i][c + 1] = !inT2Range ? COLOR.na : (isT2Locked ? COLOR.locked : COLOR.t2);
      }
      if (isT2Locked) {
        vals[base + 1][c + 1] = '二車鎖定';
      }
    }
  }

  shifts.forEach(s => {
    const pos = cellOf[s.date];
    if (!pos) return;
    const r = pos.base + 1 + (s.start - CONFIG.TIER[1].start) / CONFIG.STEP_HOURS;
    const c = pos.col + (s.tier === 2 ? 1 : 0);
    const span = CONFIG.SHIFT_HOURS / CONFIG.STEP_HOURS;
    vals[r][c] = s.adminNote ? (s.name + '\n(' + s.adminNote + ')') : s.name;
    for (let k = 0; k < span; k++) bgs[r + k][c] = s.tier === 1 ? COLOR.filled1 : COLOR.filled2;
    bold[r][c] = 'bold';
    merges.push([r, c, span, 1]);
  });

  // ---- 統計區（Q~T）----
  const stat = {};
  rosterNames_().forEach(n => { stat[n] = { t1: 0, t2: 0 }; });
  shifts.forEach(s => {
    stat[s.name] = stat[s.name] || { t1: 0, t2: 0 };
    stat[s.name][s.tier === 1 ? 't1' : 't2']++;
  });
  const t1Total = shifts.filter(s => s.tier === 1).length;
  const required = days * CONFIG.REQUIRED_T1_PER_DAY;
  const ratePct = Math.round(t1Total / required * 1000) / 10;
  const reached = ratePct >= CONFIG.RATE_TARGET;
  vals[1][16] = '本月協勤排班率（一車）'; merges.push([1, 16, 1, 2]); bgs[1][16] = COLOR.title; bold[1][16] = 'bold';
  vals[1][18] = ratePct + '%（' + (reached ? '已達標' : '未達標') + '）';
  bgs[1][18] = reached ? COLOR.pass : COLOR.warn; bold[1][18] = 'bold';
  vals[1][19] = t1Total + ' / ' + required + ' 班（目標≥' + CONFIG.RATE_TARGET + '%）';
  const hdr = ['人員清單', '一車', '二車', '合計'];
  hdr.forEach((h, i) => { vals[2][16 + i] = h; bgs[2][16 + i] = COLOR.label; bold[2][16 + i] = 'bold'; });
  Object.keys(stat).forEach((n, i) => {
    const r = 3 + i;
    if (r >= vals.length) { vals.push(new Array(NCOLS).fill('')); bgs.push(new Array(NCOLS).fill(COLOR.white)); bold.push(new Array(NCOLS).fill('normal')); }
    vals[r][16] = n; vals[r][17] = stat[n].t1; vals[r][18] = stat[n].t2; vals[r][19] = stat[n].t1 + stat[n].t2;
  });

  // ---- 寫入 ----
  const totalRows = vals.length;
  if (sh.getMaxRows() < totalRows) sh.insertRowsAfter(sh.getMaxRows(), totalRows - sh.getMaxRows());
  if (sh.getMaxColumns() < NCOLS) sh.insertColumnsAfter(sh.getMaxColumns(), NCOLS - sh.getMaxColumns());
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart();
  sh.clear();
  const rng = sh.getRange(1, 1, totalRows, NCOLS);
  rng.setValues(vals).setBackgrounds(bgs).setFontWeights(bold)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).setFontSize(10);
  const cal = sh.getRange(2, 1, 2 + weeks * WEEK_BLOCK - 1, 15);
  cal.setBorder(true, true, true, true, true, true, '#9BBBD9', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(3, 17, Object.keys(stat).length + 1, 4).setBorder(true, true, true, true, true, true, '#9BBBD9', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1B3A5C');
  sh.setRowHeight(1, 40);
  sh.setColumnWidth(1, 80);
  sh.setColumnWidths(2, 14, 72);
  sh.setColumnWidth(16, 20);
  sh.setColumnWidth(17, 110);
  sh.setColumnWidths(18, 3, 95);
  applyMerges_(ss, sh, merges);
  return sh;
}

/** 使用 Sheets 進階服務一次合併（較快）；若未啟用則逐一合併 */
function applyMerges_(ss, sh, merges) {
  if (typeof Sheets !== 'undefined') {
    try {
      SpreadsheetApp.flush();
      const sid = sh.getSheetId();
      Sheets.Spreadsheets.batchUpdate({
        requests: merges.map(g => ({
          mergeCells: {
            mergeType: 'MERGE_ALL',
            range: { sheetId: sid, startRowIndex: g[0], endRowIndex: g[0] + g[2], startColumnIndex: g[1], endColumnIndex: g[1] + g[3] },
          },
        })),
      }, ss.getId());
      return;
    } catch (err) { console.warn('進階服務合併失敗，改用逐一合併：' + err); }
  }
  merges.forEach(g => sh.getRange(g[0] + 1, g[1] + 1, g[2], g[3]).merge());
}

// ======================= 資料存取工具 =======================
function getSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('找不到工作表「' + name + '」，請先執行 setup()');
  return sh;
}

function readMembers_() {
  const sh = getSheet_(SHEET_MEMBERS);
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, MEMBER_HEADERS.length).getValues()
    .map((r, i) => ({
      row: i + 2, name: String(r[0]).trim(), pinHash: String(r[1]), role: String(r[2]) === 'admin' ? 'admin' : 'member',
      active: String(r[3]) !== '停用', mustChange: String(r[4]) === 'Y',
    }))
    .filter(m => m.name);
}

function readShifts_() {
  const sh = getSheet_(SHEET_SHIFTS);
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  const cols = Math.max(sh.getLastColumn(), SHIFT_HEADERS.length);
  return sh.getRange(2, 1, n, cols).getValues()
    .map((r, i) => ({
      row: i + 2, id: String(r[0]),
      date: r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Taipei', 'yyyy-MM-dd') : String(r[1]),
      tier: Number(r[2]), start: Number(r[3]), end: Number(r[4]), name: String(r[5]), createdAt: String(r[6]),
      adminNote: String(r[7] || ''),
    }))
    .filter(s => s.id);
}

function stripRow_(s) {
  return { id: s.id, date: s.date, tier: s.tier, start: s.start, end: s.end, name: s.name, createdAt: s.createdAt, adminNote: s.adminNote || '' };
}

function requireSession_(token) {
  const name = token && CacheService.getScriptCache().get('tok_' + token);
  if (!name) throw new Error('SESSION_EXPIRED：登入已逾時，請重新登入');
  const m = readMembers_().find(x => x.name === name && x.active);
  if (!m) throw new Error('SESSION_EXPIRED：帳號已停用');
  return m;
}

function requireAdmin_(token) {
  const m = requireSession_(token);
  if (m.role !== 'admin') throw new Error('此功能僅限管理員使用');
  return m;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('系統忙碌中，請稍後再試');
  try { return fn(); } finally { lock.releaseLock(); }
}

function hashPin_(name, pin) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + name + '|' + pin, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function pad_(n) { return ('0' + n).slice(-2); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

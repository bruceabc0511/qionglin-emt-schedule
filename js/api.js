/**
 * API 介面層
 *  - RemoteApi：呼叫 Google Apps Script（Google 試算表）
 *  - LocalApi ：未設定 API_URL 時的離線模式（邏輯與 Code.gs 相同，無預載測試班表）
 */
(function () {
  const CFG = window.APP_CONFIG;
  const pad = n => String(n).padStart(2, '0');

  // 清除舊版測試資料
  try { localStorage.removeItem('vems_mock_db_v1'); } catch (_) {}

  // ---------------- Remote（Google Apps Script） ----------------
  class RemoteApi {
    constructor(url) { this.url = url; this.mode = 'google'; }
    async call(action, payload = {}) {
      let res;
      try {
        res = await fetch(this.url, {
          method: 'POST',
          // 使用 text/plain 避免 CORS preflight（Apps Script 不支援 OPTIONS）
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(Object.assign({ action }, payload)),
          redirect: 'follow',
        });
      } catch (e) {
        throw new Error('無法連線到伺服器，請檢查網路或 API_URL 設定');
      }
      const j = await res.json();
      if (!j.ok) throw new Error(j.error || '未知錯誤');
      return j.data;
    }
  }

  // ---------------- Local（離線模式，無測試假資料） ----------------
  const DB_KEY = 'vems_db_v2';
  const INITIAL = ['賴致平', '林心彤', '許世良', '賴志昀', '官正軒', '蔡宗佑', '龍賢銨', '林月娥',
    '陳容清', '張謙加', '陳長鎧', '吳宛諭', '彭育琳', '羅艾婷', '陳佩秀', '陳姿璇',
    '官振瑋', '劉兆華', '劉劭倫', '閻欣慈', '謝桂湘', '鍾桂蓉', '潘麗芬', '洪如意'];

  function hash(s) {
    let h = 2166136261;
    for (const ch of s) { h ^= ch.codePointAt(0); h = Math.imul(h, 16777619) >>> 0; }
    return h.toString(16);
  }
  const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
  const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };

  class LocalApi {
    constructor() { this.mode = 'mock'; this.ensureDb(); }

    ensureDb() {
      const existing = localStorage.getItem(DB_KEY);
      if (existing) {
        try {
          const parsed = JSON.parse(existing);
          if (!Array.isArray(parsed.lockedT2Dates)) { parsed.lockedT2Dates = []; this.save(parsed); }
          return;
        } catch (_) {}
      }
      const members = [{ name: '管理員', pin: hash('管理員|9999'), role: 'admin', active: true, mustChange: false }]
        .concat(INITIAL.map(n => ({ name: n, pin: hash(n + '|0000'), role: 'member', active: true, mustChange: false })));
      this.save({ members, shifts: [], lockedT2Dates: [], sessions: {} });
    }
    load() {
      const db = JSON.parse(localStorage.getItem(DB_KEY));
      if (!Array.isArray(db.lockedT2Dates)) db.lockedT2Dates = [];
      return db;
    }
    save(db) { localStorage.setItem(DB_KEY, JSON.stringify(db)); }

    async call(action, p = {}) {
      await new Promise(r => setTimeout(r, 80));
      const fn = this['_' + action];
      if (!fn) throw new Error('未知的操作：' + action);
      return fn.call(this, p);
    }

    session(token) {
      const db = this.load();
      const name = db.sessions[token];
      const m = name && db.members.find(x => x.name === name && x.active);
      if (!m) throw new Error('SESSION_EXPIRED：登入已逾時，請重新登入');
      return m;
    }
    admin(token) { const m = this.session(token); if (m.role !== 'admin') throw new Error('此功能僅限管理員使用'); return m; }
    roster(db) { return db.members.filter(m => m.active && m.role !== 'admin').map(m => m.name); }

    _members() { return this.load().members.filter(m => m.active).map(m => m.name); }
    _login({ name, pin }) {
      const db = this.load();
      const m = db.members.find(x => x.name === name && x.active);
      if (!m || m.pin !== hash(name + '|' + pin)) throw new Error('姓名或 PIN 碼錯誤');
      const token = uuid();
      db.sessions[token] = name;
      this.save(db);
      return { token, name, role: m.role, mustChangePin: m.mustChange };
    }
    _changePin({ token, oldPin, newPin }) {
      const me = this.session(token);
      if (!/^\d{4,8}$/.test(newPin)) throw new Error('PIN 碼須為 4～8 位數字');
      if (newPin === '0000') throw new Error('請勿使用預設 PIN 碼');
      const db = this.load();
      const m = db.members.find(x => x.name === me.name);
      if (m.pin !== hash(me.name + '|' + oldPin)) throw new Error('原 PIN 碼錯誤');
      m.pin = hash(me.name + '|' + newPin); m.mustChange = false;
      this.save(db);
      return true;
    }
    _getMonth({ token, year, month }) {
      this.session(token);
      const db = this.load();
      const prefix = `${year}-${pad(month)}-`;
      return {
        year, month,
        shifts: db.shifts.filter(s => s.date.startsWith(prefix)),
        lockedT2Dates: db.lockedT2Dates.filter(d => d.startsWith(prefix)),
        members: this.roster(db),
        sheetUrl: '',
      };
    }
    _setTier2Lock({ token, date, dates, locked }) {
      this.admin(token);
      const list = Array.isArray(dates) ? dates.map(String) : [String(date || '')];
      if (!list.length || !/^\d{4}-\d{2}-\d{2}$/.test(list[0])) throw new Error('日期格式錯誤');
      const db = this.load();
      const set = new Set(db.lockedT2Dates);
      list.forEach(d => { if (locked) set.add(d); else set.delete(d); });
      db.lockedT2Dates = Array.from(set).sort();
      this.save(db);
      const prefix = list[0].slice(0, 8);
      return { lockedT2Dates: db.lockedT2Dates.filter(d => d.startsWith(prefix)) };
    }
    validate(db, p) {
      const t = CFG.TIER[p.tier];
      const end = p.start + CFG.SHIFT_HOURS;
      if (!t || p.start < t.start || end > t.end || (p.start - t.start) % CFG.STEP_HOURS) {
        throw new Error(`${t ? t.name : ''}可排時段為 ${pad(t.start)}:00～${pad(t.end)}:00`);
      }
      this.editable(p.date, p.isAdmin);
      if (p.tier === 2 && db.lockedT2Dates.includes(p.date)) {
        throw new Error('當日二車時段已鎖定未開放排班');
      }
      for (const s of db.shifts) {
        if (s.id === p.excludeId || s.date !== p.date) continue;
        if (!(s.start < end && p.start < s.end)) continue;
        if (s.tier === p.tier) throw new Error(`此時段已被「${s.name}」搶先登記（${t.name} ${pad(s.start)}:00-${pad(s.end)}:00），請重新選擇`);
        if (s.name === p.name) throw new Error(`「${p.name}」在同一時間已登記另一車的班`);
      }
    }
    editable(date, isAdmin) {
      if (isAdmin) return;
      const today = todayStr();
      if (!CFG.ALLOW_PAST_EDIT && date < today) throw new Error('已過去的日期無法再變更');
      const now = new Date();
      const ty = now.getFullYear(), tm = now.getMonth() + 1, td = now.getDate(), th = now.getHours();
      const [dy, dm] = date.split('-').map(Number);
      const monthDiff = (dy - ty) * 12 + (dm - tm);
      const openDay = CFG.OPEN_NEXT_MONTH_DAY || 20;
      const openHour = CFG.OPEN_NEXT_MONTH_HOUR != null ? CFG.OPEN_NEXT_MONTH_HOUR : 20;
      const isOpened = td > openDay || (td === openDay && th >= openHour);
      if (monthDiff === 1 && !isOpened) {
        throw new Error(`次月（${dm}月）班表將於本月 ${openDay} 日 ${pad(openHour)}:00 自動開放選填`);
      }
      if (monthDiff > 1) {
        throw new Error(`${dm}月班表尚未開放選填（每月 ${openDay} 日 ${pad(openHour)}:00 開放次月班表）`);
      }
    }
    _addShift({ token, date, tier, start, targetName }) {
      const me = this.session(token);
      const isAdmin = me.role === 'admin';
      const name = (isAdmin && targetName) ? String(targetName).trim() : me.name;
      const adminNote = (isAdmin && name !== me.name) ? '管理員異動' : '';
      const db = this.load();
      this.validate(db, { date, tier, start, name, isAdmin });
      const rec = { id: uuid(), date, tier, start, end: start + CFG.SHIFT_HOURS, name, createdAt: new Date().toISOString(), adminNote };
      db.shifts.push(rec);
      this.save(db);
      return rec;
    }
    _updateShift({ token, id, tier, start, targetName }) {
      const me = this.session(token);
      const isAdmin = me.role === 'admin';
      const db = this.load();
      const s = db.shifts.find(x => x.id === id);
      if (!s) throw new Error('找不到此班次，可能已被刪除');
      if (s.name !== me.name && !isAdmin) throw new Error('僅限本人或管理員可以修改排班');
      const name = (isAdmin && targetName) ? String(targetName).trim() : s.name;
      const isOther = isAdmin && (s.name !== me.name || name !== me.name);
      const adminNote = isOther ? '管理員異動' : (s.adminNote || '');
      this.editable(s.date, isAdmin);
      this.validate(db, { date: s.date, tier, start, name, excludeId: id, isAdmin });
      Object.assign(s, { tier, start, end: start + CFG.SHIFT_HOURS, name, adminNote });
      this.save(db);
      return s;
    }
    _deleteShift({ token, id }) {
      const me = this.session(token);
      const isAdmin = me.role === 'admin';
      const db = this.load();
      const s = db.shifts.find(x => x.id === id);
      if (!s) throw new Error('找不到此班次，可能已被刪除');
      if (s.name !== me.name && !isAdmin) throw new Error('僅限本人或管理員可以取消排班');
      this.editable(s.date, isAdmin);
      db.shifts = db.shifts.filter(x => x.id !== id);
      this.save(db);
      return true;
    }
    _adminListMembers({ token }) {
      this.admin(token);
      return this.load().members.map(m => ({ name: m.name, role: m.role, active: m.active, mustChangePin: m.mustChange }));
    }
    _adminAddMember({ token, name, pin, role }) {
      this.admin(token);
      name = String(name || '').trim(); pin = String(pin || '0000');
      if (!name || name.length > 20) throw new Error('請輸入正確的姓名');
      if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN 碼須為 4～8 位數字');
      const db = this.load();
      const ex = db.members.find(m => m.name === name);
      if (ex && ex.active) throw new Error(`「${name}」已在人員清單中`);
      if (ex) Object.assign(ex, { active: true, pin: hash(name + '|' + pin), role, mustChange: true });
      else db.members.push({ name, pin: hash(name + '|' + pin), role: role === 'admin' ? 'admin' : 'member', active: true, mustChange: true });
      this.save(db);
      return true;
    }
    _adminRemoveMember({ token, name }) {
      const me = this.admin(token);
      if (name === me.name) throw new Error('不可刪除自己的帳號');
      const db = this.load();
      const m = db.members.find(x => x.name === name);
      if (!m) throw new Error('查無此人員');
      m.active = false;
      this.save(db);
      return true;
    }
    _adminRestoreMember({ token, name }) {
      this.admin(token);
      const db = this.load();
      const m = db.members.find(x => x.name === name);
      if (!m) throw new Error('查無此人員');
      m.active = true;
      this.save(db);
      return true;
    }
    _adminResetPin({ token, name, pin }) {
      this.admin(token);
      pin = String(pin || '0000');
      if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN 碼須為 4～8 位數字');
      const db = this.load();
      const m = db.members.find(x => x.name === name);
      if (!m) throw new Error('查無此人員');
      m.pin = hash(name + '|' + pin); m.mustChange = true;
      this.save(db);
      return true;
    }
  }

  window.API = CFG.API_URL ? new RemoteApi(CFG.API_URL) : new LocalApi();
  window.API.resetMock = () => { localStorage.removeItem(DB_KEY); localStorage.removeItem('vems_mock_db_v1'); location.reload(); };
})();

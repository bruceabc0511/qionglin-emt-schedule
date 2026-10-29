/* =========================================================
   協勤排班系統 前端主程式
   ========================================================= */
(function () {
  const CFG = window.APP_CONFIG;
  const API = window.API;
  const $ = s => document.querySelector(s);
  const pad = n => String(n).padStart(2, '0');
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const WD = ['日', '一', '二', '三', '四', '五', '六'];
  const SLOT_ROWS = (CFG.TIER[1].end - CFG.TIER[1].start) / CFG.STEP_HOURS;   // 14 格（08~22，每格 1 小時）
  const SESSION_KEY = 'vems_session';
  const RATE_TARGET = CFG.RATE_TARGET || 90;
  const OPEN_DAY = CFG.OPEN_NEXT_MONTH_DAY || 20;
  const OPEN_HOUR = CFG.OPEN_NEXT_MONTH_HOUR != null ? CFG.OPEN_NEXT_MONTH_HOUR : 20;
  const MAX_MULTI_SHIFTS = 3;

  const now = new Date();
  const state = {
    session: JSON.parse(sessionStorage.getItem(SESSION_KEY) || 'null'),
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    data: { shifts: [], lockedT2Dates: [], members: [], sheetUrl: '' },
    modal: null,          // { date, editing, tier, selections: [{tier, start}], targetName }
    pollTimer: null,
    busy: false,
  };

  // ---------------- 共用工具 ----------------
  const isAdmin = () => state.session && state.session.role === 'admin';
  const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
  const hhmm = h => `${pad(h)}:00`;
  const range = (s, e) => `${hhmm(s)}–${hhmm(e)}`;

  /** 檢查指定年月是否已開放一般成員選填（每月 20 號 20:00 開放次月班表） */
  function getMonthOpenInfo(year, month) {
    const d = new Date();
    const ty = d.getFullYear(), tm = d.getMonth() + 1, td = d.getDate(), th = d.getHours();
    const diff = (year - ty) * 12 + (month - tm);
    const timeStr = `${pad(OPEN_HOUR)}:00`;
    if (diff < 0) return { open: false, reason: '已過去的月份無法再變更', banner: '' };
    if (diff === 0) return { open: true, reason: '', banner: '' };
    if (diff === 1) {
      const isOpened = td > OPEN_DAY || (td === OPEN_DAY && th >= OPEN_HOUR);
      if (isOpened) {
        return {
          open: true,
          reason: '',
          bannerType: 'open',
          banner: `🔓 ${year - 1911}年${month}月班表已開放選填`,
        };
      }
      return {
        open: false,
        reason: `次月（${month}月）班表將於本月 ${OPEN_DAY} 日 ${timeStr} 自動開放選填`,
        bannerType: 'wait',
        banner: `⏰ ${year - 1911}年${month}月班表將於本月 ${OPEN_DAY} 號 ${timeStr} 自動開放選填（目前僅供預覽，管理員可先行排班）`,
      };
    }
    return {
      open: false,
      reason: `${month}月班表尚未開放選填（每月 ${OPEN_DAY} 日 ${timeStr} 開放次月班表）`,
      bannerType: 'wait',
      banner: `🔒 ${year - 1911}年${month}月班表尚未開放選填（將於前一個月 ${OPEN_DAY} 號 ${timeStr} 自動開放）`,
    };
  }

  function checkDateBookable(dateStr) {
    if (isAdmin()) return { ok: true };
    if (!CFG.ALLOW_PAST_EDIT && dateStr < todayStr()) {
      return { ok: false, msg: '已過去的日期無法再變更' };
    }
    const [y, m] = dateStr.split('-').map(Number);
    const info = getMonthOpenInfo(y, m);
    if (!info.open) return { ok: false, msg: info.reason };
    return { ok: true };
  }

  /** 計算該月第一個星期三的日期（1~7） */
  function firstWednesdayDay(year, month) {
    const firstDow = new Date(year, month - 1, 1).getDay(); // 0=日..3=三
    return 1 + ((3 - firstDow + 7) % 7);
  }
  function isTrainingDate(dateStr) {
    const [y, m, d] = dateStr.split('-').map(Number);
    return d === firstWednesdayDay(y, m);
  }
  function isT2Locked(dateStr) {
    return (state.data.lockedT2Dates || []).includes(dateStr);
  }

  function toast(msg, type = '') {
    const el = document.createElement('div');
    el.className = 'toast ' + type;
    el.textContent = msg;
    $('#toastWrap').appendChild(el);
    setTimeout(() => el.remove(), type === 'err' ? 5000 : 3000);
  }
  function loading(on) { $('#loading').classList.toggle('hidden', !on); }

  async function call(action, payload = {}, { silent = false } = {}) {
    if (!silent) loading(true);
    try {
      return await API.call(action, Object.assign({ token: state.session && state.session.token }, payload));
    } catch (e) {
      if (String(e.message).startsWith('SESSION_EXPIRED')) {
        logout();
        toast(e.message.replace(/^SESSION_EXPIRED：/, ''), 'err');
      }
      throw e;
    } finally {
      if (!silent) loading(false);
    }
  }

  // ---------------- 登入 / 登出 ----------------
  async function initLogin() {
    $('#loginView').classList.remove('hidden');
    $('#appView').classList.add('hidden');
    $('#userArea').classList.add('hidden');
    $('#topbarQuick').classList.add('hidden');
    if (API.mode === 'mock') {
      $('#mockHint').innerHTML = `目前尚未設定 Google 試算表連線網址（<code>js/config.js</code> 的 <code>API_URL</code>），資料暫存於本機瀏覽器。`;
    } else {
      $('#mockHint').innerHTML = '';
    }
    try {
      const names = await API.call('members');
      $('#loginName').innerHTML = '<option value="">— 請選擇姓名 —</option>' +
        names.map(n => `<option>${esc(n)}</option>`).join('');
    } catch (e) {
      $('#loginName').innerHTML = '<option value="">無法載入人員清單</option>';
      toast(e.message, 'err');
    }
  }

  $('#loginForm').addEventListener('submit', async e => {
    e.preventDefault();
    const name = $('#loginName').value, pin = $('#loginPin').value;
    if (!name) return toast('請選擇姓名', 'err');
    try {
      const s = await call('login', { name, pin });
      state.session = s;
      sessionStorage.setItem(SESSION_KEY, JSON.stringify(s));   // 每個分頁獨立登入
      $('#loginPin').value = '';
      toast(`歡迎，${s.name}！`, 'ok');
      await enterApp();
      if (s.mustChangePin) openPinModal(true);
    } catch (err) { toast(err.message, 'err'); }
  });

  function logout() {
    sessionStorage.removeItem(SESSION_KEY);
    state.session = null;
    clearInterval(state.pollTimer);
    document.querySelectorAll('.modal').forEach(m => m.classList.add('hidden'));
    initLogin();
  }
  $('#logoutBtn').onclick = logout;

  // ---------------- 主畫面與上方區塊收合控制 ----------------
  function setTopMenuCollapsed(collapsed) {
    const ua = $('#userArea');
    ua.classList.toggle('collapsed', collapsed);
    $('#menuArrow').textContent = collapsed ? '▾' : '▴';
  }

  function setDashboardCollapsed(collapsed) {
    const dp = $('#dashboardPanel');
    dp.classList.toggle('collapsed', collapsed);
    $('#dashArrow').textContent = collapsed ? '▾' : '▴';
  }

  $('#toggleMenuBtn').onclick = () => {
    setTopMenuCollapsed(!$('#userArea').classList.contains('collapsed'));
  };

  $('#toggleDashBtn').onclick = () => {
    setDashboardCollapsed(!$('#dashboardPanel').classList.contains('collapsed'));
  };

  // 點選上方工具列任一功能按鈕後，於手機版自動收合工具列
  $('#userArea').addEventListener('click', e => {
    if (window.innerWidth <= 760 && e.target.closest('.btn')) {
      setTopMenuCollapsed(true);
    }
  });

  // 手機版向下捲動進入月曆時，自動收合展開中的上方工具列與統計區
  let lastScrollY = 0;
  window.addEventListener('scroll', () => {
    const y = window.scrollY || 0;
    if (window.innerWidth <= 760 && y > 140 && y - lastScrollY > 25) {
      if (!$('#userArea').classList.contains('collapsed')) setTopMenuCollapsed(true);
    }
    lastScrollY = y;
  }, { passive: true });

  async function enterApp() {
    $('#loginView').classList.add('hidden');
    $('#appView').classList.remove('hidden');
    $('#userArea').classList.remove('hidden');
    $('#topbarQuick').classList.remove('hidden');
    $('#userChip').textContent = (isAdmin() ? '👮 ' : '👤 ') + state.session.name;
    $('#adminBtn').classList.toggle('hidden', !isAdmin());
    $('#t2ManageBtn').classList.toggle('hidden', !isAdmin());
    const mb = $('#modeBadge');
    mb.textContent = API.mode === 'mock' ? '離線模式' : 'Google試算表已連線';
    mb.classList.toggle('mock', API.mode === 'mock');

    // 手機上預設自動收合上方工具列與統計面板，減少畫面干擾
    const isMobile = window.innerWidth <= 760;
    setTopMenuCollapsed(isMobile);
    setDashboardCollapsed(isMobile);

    await loadMonth();
    clearInterval(state.pollTimer);
    state.pollTimer = setInterval(() => { if (!document.hidden) loadMonth(true); }, CFG.POLL_SECONDS * 1000);
  }

  function allMonthDatesWithMarker(year, month) {
    const days = new Date(year, month, 0).getDate();
    const ym = `${year}-${pad(month)}`;
    const list = Array.from({ length: days }, (_, i) => `${ym}-${pad(i + 1)}`);
    list.push(`${ym}-00`); // 初始化標記（代表此月份已套用二車預設全鎖定）
    return list;
  }

  async function loadMonth(silent = false) {
    try {
      const d = await call('getMonth', { year: state.year, month: state.month }, { silent });
      if (d.year !== state.year || d.month !== state.month) return;  // 使用者已切換月份
      d.lockedT2Dates = d.lockedT2Dates || [];
      const marker = `${d.year}-${pad(d.month)}-00`;
      if (!d.lockedT2Dates.includes(marker)) {
        // 該月份尚未初始化：預設將全月二車時段設為全部鎖定
        d.lockedT2Dates = allMonthDatesWithMarker(d.year, d.month);
        if (isAdmin()) {
          call('setTier2Lock', { dates: d.lockedT2Dates, locked: true }, { silent: true }).catch(() => {});
        }
      }
      state.data = d;
      $('#syncText').textContent = '最後更新 ' + new Date().toLocaleTimeString('zh-TW', { hour12: false });
      render();
      if (state.modal) renderModal();
      if (!$('#t2Modal').classList.contains('hidden')) renderT2Modal();
    } catch (e) { if (!silent) toast(e.message, 'err'); }
  }

  function shiftsOn(date) { return state.data.shifts.filter(s => s.date === date); }

  function render() {
    const { year, month } = state;
    $('#monthLabel').textContent = `${year - 1911}年${month}月（${year}）`;
    const link = $('#sheetLink');
    link.classList.toggle('hidden', !state.data.sheetUrl);
    if (state.data.sheetUrl) link.href = state.data.sheetUrl;

    const openInfo = getMonthOpenInfo(year, month);
    const bannerEl = $('#monthOpenBanner');
    if (openInfo.banner) {
      bannerEl.textContent = openInfo.banner;
      bannerEl.className = 'open-banner ' + (openInfo.bannerType || '');
    } else {
      bannerEl.className = 'open-banner hidden';
    }

    renderCalendar();
    renderStats();
  }

  function renderCalendar() {
    const { year, month } = state;
    const firstDow = new Date(year, month - 1, 1).getDay();
    const days = new Date(year, month, 0).getDate();
    const weeks = Math.ceil((firstDow + days) / 7);
    const trainingDay = firstWednesdayDay(year, month);
    const openInfo = getMonthOpenInfo(year, month);
    const today = todayStr();
    const me = state.session.name;
    const adm = isAdmin();
    let html = '<div class="cal-head corner"></div>' +
      WD.map((w, i) => `<div class="cal-head ${i === 0 || i === 6 ? 'weekend' : ''}">星期${w}</div>`).join('');

    const axis = '<div class="axis"><span></span>' +
      Array.from({ length: SLOT_ROWS }, (_, i) => {
        const h = CFG.TIER[1].start + i * CFG.STEP_HOURS;
        return `<span>${pad(h)}-${pad(h + CFG.STEP_HOURS)}</span>`;
      }).join('') + '</div>';

    for (let w = 0; w < weeks; w++) {
      html += axis;
      for (let d = 0; d < 7; d++) {
        const dayNum = w * 7 + d - firstDow + 1;
        if (dayNum < 1 || dayNum > days) { html += '<div class="day empty"></div>'; continue; }
        const date = `${year}-${pad(month)}-${pad(dayNum)}`;
        const list = shiftsOn(date);
        const t1n = list.filter(s => s.tier === 1).length;
        const locked2 = isT2Locked(date);
        const isTrain = dayNum === trainingDay;
        const notBookable = (date < today || !openInfo.open) && !adm;
        const cls = ['day', d === 0 || d === 6 ? 'weekend' : '', date === today ? 'today' : '', notBookable ? 'past' : '', adm ? 'admin-view' : ''].join(' ');
        let grid = '';
        for (let i = 0; i < SLOT_ROWS; i++) {
          const h = CFG.TIER[1].start + i * CFG.STEP_HOURS;
          grid += `<div class="tl" style="grid-row:${i + 1}">${pad(h)}-${pad(h + CFG.STEP_HOURS)}</div>`;
          grid += `<div class="slot t1" style="grid-row:${i + 1}" data-date="${date}" data-tier="1" data-hour="${h}" title="一車 ${pad(h)}:00 起（4小時）"></div>`;
          if (h + CFG.STEP_HOURS <= CFG.TIER[2].end) {
            grid += locked2
              ? `<div class="slot locked ${i === 0 ? 'first-locked' : ''}" style="grid-row:${i + 1}" data-date="${date}" data-tier="2" data-hour="${h}" title="當日二車時段已鎖定未開放"></div>`
              : `<div class="slot t2" style="grid-row:${i + 1}" data-date="${date}" data-tier="2" data-hour="${h}" title="二車 ${pad(h)}:00 起（4小時）"></div>`;
          } else {
            grid += `<div class="slot na" style="grid-row:${i + 1}" title="二車 18:00 後不排班"></div>`;
          }
        }
        list.forEach(s => {
          const row = (s.start - CFG.TIER[1].start) / CFG.STEP_HOURS + 1;
          const span = CFG.SHIFT_HOURS / CFG.STEP_HOURS;
          const mine = s.name === me;
          const canEdit = (mine && !notBookable) || adm;
          const noteHtml = s.adminNote ? `<span class="adm-tag">(${esc(s.adminNote)})</span>` : '';
          const noteTitle = s.adminNote ? ` [${s.adminNote}]` : '';
          grid += `<div class="shift t${s.tier} ${mine ? 'mine' : ''} ${canEdit ? 'editable' : ''}" style="grid-row:${row} / span ${span}"
            ${canEdit ? `data-id="${s.id}"` : ''} title="${CFG.TIER[s.tier].name} ${range(s.start, s.end)}｜${esc(s.name)}${noteTitle}${canEdit ? '（點擊可修改）' : ''}">
            <span>${esc(s.name)}</span><small>${pad(s.start)}-${pad(s.end)}</small>${noteHtml}</div>`;
        });
        const lockBtn = adm
          ? `<button type="button" class="t2-lock-toggle ${locked2 ? 'locked' : ''}" data-lock-date="${date}" title="點擊切換當日二車開放/鎖定">${locked2 ? '🔒二車' : '🔓二車'}</button>`
          : (locked2 ? `<span class="t2-lock-toggle locked" title="當日二車未開放">🔒二車</span>` : '');
        html += `<div class="${cls}">
          <div class="day-head">
            <span><span class="day-num">${dayNum}</span>${isTrain ? '<span class="day-training">(定訓)</span>' : ''}<span class="day-wd">星期${WD[d]}</span></span>
            <span class="day-head-right">${lockBtn}<span class="badge ${t1n >= CFG.REQUIRED_T1_PER_DAY ? 'full' : ''}" title="一車已排 / 應排">${t1n}/${CFG.REQUIRED_T1_PER_DAY}</span></span>
          </div>
          <div class="day-grid">${grid}</div></div>`;
      }
    }
    $('#calendar').innerHTML = html;
  }

  function renderStats() {
    const { year, month } = state;
    const days = new Date(year, month, 0).getDate();
    const shifts = state.data.shifts;
    const t1 = shifts.filter(s => s.tier === 1);
    const required = days * CFG.REQUIRED_T1_PER_DAY;
    const rate = required ? Math.min(100, t1.length / required * 100) : 0;
    const reached = rate >= RATE_TARGET;
    let gap = 0;
    for (let d = 1; d <= days; d++) {
      const date = `${year}-${pad(month)}-${pad(d)}`;
      gap += Math.max(0, CFG.REQUIRED_T1_PER_DAY - t1.filter(s => s.date === date).length);
    }
    $('#rateCard').classList.toggle('reached', reached);
    const gb = $('#goalBadge');
    gb.textContent = reached ? `✓ 已達標（≥${RATE_TARGET}%）` : `未達標（目標 ${RATE_TARGET}%）`;
    gb.classList.toggle('reached', reached);
    $('#rateValue').textContent = rate.toFixed(1) + '%';
    $('#rateBar').style.width = rate + '%';
    $('#rateSub').textContent = `一車已排 ${t1.length} 班 ／ 應排 ${required} 班（${days} 天 × ${CFG.REQUIRED_T1_PER_DAY} 班）｜達標門檻 ${RATE_TARGET}%`;
    $('#gapValue').textContent = gap;
    $('#gapSub').textContent = gap ? '尚待填補，請踴躍登記' : '本月一車已全數排滿 🎉';
    $('#t2Value').textContent = shifts.length - t1.length;
    const lockedCount = (state.data.lockedT2Dates || []).filter(x => !x.endsWith('-00')).length;
    const openT2Count = Math.max(0, days - lockedCount);
    $('#t2Sub').textContent = openT2Count ? `不列入排班率（本月開放 ${openT2Count} 天）` : '不列入排班率（預設全月鎖定）';
    const me = state.session.name;
    const mine = shifts.filter(s => s.name === me);
    $('#myValue').textContent = mine.length;
    $('#mySub').textContent = `一車 ${mine.filter(s => s.tier === 1).length} 班、二車 ${mine.filter(s => s.tier === 2).length} 班`;
    $('#dashMiniSummary').textContent = `排班率 ${rate.toFixed(1)}%｜我的班 ${mine.length}`;

    const stat = new Map();
    state.data.members.forEach(n => stat.set(n, { t1: 0, t2: 0 }));
    shifts.forEach(s => {
      if (!stat.has(s.name)) stat.set(s.name, { t1: 0, t2: 0 });
      stat.get(s.name)[s.tier === 1 ? 't1' : 't2']++;
    });
    $('#statTable tbody').innerHTML = [...stat.entries()].map(([n, v]) => {
      const total = v.t1 + v.t2;
      return `<tr class="${n === me ? 'me' : ''} ${total ? '' : 'zero'}"><td>${esc(n)}</td><td>${v.t1}</td><td>${v.t2}</td><td><b>${total}</b></td></tr>`;
    }).join('');
  }

  // ---------------- 月份切換 ----------------
  function shiftMonth(delta) {
    const d = new Date(state.year, state.month - 1 + delta, 1);
    state.year = d.getFullYear(); state.month = d.getMonth() + 1;
    state.data = { shifts: [], lockedT2Dates: [], members: state.data.members, sheetUrl: '' };
    render();
    loadMonth();
  }
  $('#prevMonth').onclick = () => shiftMonth(-1);
  $('#nextMonth').onclick = () => shiftMonth(1);
  $('#todayBtn').onclick = () => { const t = new Date(); state.year = t.getFullYear(); state.month = t.getMonth() + 1; render(); loadMonth(); };
  $('#refreshBtn').onclick = () => loadMonth();

  // ---------------- 二車鎖定管理（預設全部鎖定） ----------------
  async function toggleT2DateLock(date, locked) {
    try {
      const [y, m] = date.split('-').map(Number);
      const marker = `${y}-${pad(m)}-00`;
      if (!locked && !(state.data.lockedT2Dates || []).includes(marker)) {
        await call('setTier2Lock', { dates: allMonthDatesWithMarker(y, m), locked: true }, { silent: true });
      }
      const dates = locked ? [date, marker] : [date];
      const res = await call('setTier2Lock', { dates, locked });
      const list = res.lockedT2Dates || [];
      if (!list.includes(marker)) list.push(marker);
      state.data.lockedT2Dates = list;
      toast(`${date} 二車已${locked ? '鎖定' : '開放'}`, 'ok');
      render();
      if (state.modal) renderModal();
      if (!$('#t2Modal').classList.contains('hidden')) renderT2Modal();
    } catch (e) { toast(e.message, 'err'); }
  }

  async function batchSetT2Lock(locked) {
    const { year, month } = state;
    const days = new Date(year, month, 0).getDate();
    const ym = `${year}-${pad(month)}`;
    const marker = `${ym}-00`;
    const dates = Array.from({ length: days }, (_, i) => `${ym}-${pad(i + 1)}`);
    if (locked) dates.push(marker);
    try {
      if (!locked) {
        // 先確保初始化標記存在於後端，再將全月真實日期解鎖
        await call('setTier2Lock', { dates: [dates[0], marker], locked: true }, { silent: true });
      }
      const res = await call('setTier2Lock', { dates, locked });
      const list = res.lockedT2Dates || [];
      if (!list.includes(marker)) list.push(marker);
      state.data.lockedT2Dates = list;
      toast(`本月二車已全數${locked ? '鎖定' : '開放'}`, 'ok');
      render();
      renderT2Modal();
    } catch (e) { toast(e.message, 'err'); }
  }

  $('#t2ManageBtn').onclick = () => {
    renderT2Modal();
    $('#t2Modal').classList.remove('hidden');
  };

  function renderT2Modal() {
    const { year, month } = state;
    const days = new Date(year, month, 0).getDate();
    const trainDay = firstWednesdayDay(year, month);
    $('#t2ModalTitle').textContent = `${year - 1911}年${month}月 二車時段開放／鎖定管理`;
    let html = '';
    for (let d = 1; d <= days; d++) {
      const date = `${year}-${pad(month)}-${pad(d)}`;
      const wd = WD[new Date(year, month - 1, d).getDay()];
      const locked = isT2Locked(date);
      const tLabel = d === trainDay ? '(定訓)' : '';
      html += `<button type="button" class="t2-day-btn ${locked ? 'locked' : ''}" data-t2-date="${date}">
        ${d}${tLabel}（${wd}）<small>${locked ? '🔒 已鎖定' : '🔓 開放中'}</small></button>`;
    }
    $('#t2DayGrid').innerHTML = html;
  }

  $('#t2DayGrid').addEventListener('click', e => {
    const b = e.target.closest('button[data-t2-date]');
    if (!b) return;
    const date = b.dataset.t2Date;
    toggleT2DateLock(date, !isT2Locked(date));
  });
  $('#unlockAllT2Btn').onclick = () => batchSetT2Lock(false);
  $('#lockAllT2Btn').onclick = () => batchSetT2Lock(true);

  // ---------------- 排班對話框（支援單日多選最多 3 班） ----------------
  $('#calendar').addEventListener('click', e => {
    const lockBtn = e.target.closest('button[data-lock-date]');
    if (lockBtn && isAdmin()) {
      const d = lockBtn.dataset.lockDate;
      return toggleT2DateLock(d, !isT2Locked(d));
    }
    const editableShift = e.target.closest('.shift.editable');
    if (editableShift) {
      const s = state.data.shifts.find(x => x.id === editableShift.dataset.id);
      if (!s) return;
      const chk = checkDateBookable(s.date);
      if (!chk.ok) return toast(chk.msg, 'err');
      return openShiftModal(s.date, s, s.tier, s.start);
    }
    const lockedSlot = e.target.closest('.slot.locked');
    if (lockedSlot) {
      if (isAdmin()) {
        return openShiftModal(lockedSlot.dataset.date, null, 2, null);
      }
      return toast('當日二車時段已鎖定未開放排班', 'err');
    }
    const slot = e.target.closest('.slot:not(.na)');
    if (!slot) return;
    const date = slot.dataset.date;
    const chk = checkDateBookable(date);
    if (!chk.ok) return toast(chk.msg, 'err');
    const tier = Number(slot.dataset.tier);
    const t = CFG.TIER[tier];
    let start = Math.min(Number(slot.dataset.hour), t.end - CFG.SHIFT_HOURS);
    const targetName = isAdmin() ? (state.data.members[0] || state.session.name) : state.session.name;
    const opts = slotOptions(date, tier, null, targetName);
    if (!opts.find(o => o.start === start && !o.taken)) {
      const alt = opts.find(o => !o.taken && Math.abs(o.start - start) <= 3);
      start = alt ? alt.start : null;
    }
    openShiftModal(date, null, tier, start);
  });

  /** 計算某日某車別所有可選時段（每 1 小時一格，4 小時一班）與占用狀況 */
  function slotOptions(date, tier, editing, personName) {
    const t = CFG.TIER[tier];
    const who = personName || (editing ? editing.name : state.session.name);
    const list = shiftsOn(date).filter(s => !editing || s.id !== editing.id);
    const locked2 = tier === 2 && isT2Locked(date);
    const opts = [];
    for (let st = t.start; st + CFG.SHIFT_HOURS <= t.end; st += CFG.STEP_HOURS) {
      const en = st + CFG.SHIFT_HOURS;
      const clash = list.filter(s => s.start < en && st < s.end);
      const sameTier = clash.find(s => s.tier === tier);
      const selfOther = clash.find(s => s.tier !== tier && s.name === who);
      let taken = '';
      if (locked2) taken = '二車時段已鎖定';
      else if (sameTier) taken = `已被 ${sameTier.name} 登記`;
      else if (selfOther) taken = `${who} 同時段已有另一車`;
      opts.push({ start: st, end: en, taken });
    }
    return opts;
  }

  function openShiftModal(date, editing, tier, start) {
    const defaultTarget = editing
      ? editing.name
      : (isAdmin() ? (state.data.members[0] || state.session.name) : state.session.name);
    const selections = start != null ? [{ tier, start }] : [];
    state.modal = { date, editing, tier, selections, targetName: defaultTarget };
    renderModal();
    $('#shiftModal').classList.remove('hidden');
  }

  function isIntervalsOverlap(s1, s2) {
    const e1 = s1 + CFG.SHIFT_HOURS;
    const e2 = s2 + CFG.SHIFT_HOURS;
    return s1 < e2 && s2 < e1;
  }

  function renderModal() {
    const m = state.modal;
    if (!m) return;
    if (m.editing && !state.data.shifts.find(s => s.id === m.editing.id)) {
      closeModal('#shiftModal');
      return toast('此班次已不存在', 'err');
    }
    const [y, mo, d] = m.date.split('-').map(Number);
    const wd = WD[new Date(y, mo - 1, d).getDay()];
    const trainTag = isTrainingDate(m.date) ? '(定訓)' : '';
    const adm = isAdmin();
    const editingOther = m.editing && m.editing.name !== state.session.name;
    $('#shiftTitle').textContent = `${y - 1911}年${mo}月${d}日${trainTag}（${wd}）${m.editing ? (editingOther ? ` 管理員修改「${m.editing.name}」的排班` : ' 修改我的排班') : ' 登記協勤'}`;

    const noteInfo = (m.editing && m.editing.adminNote) ? `（異動註記：<b>${esc(m.editing.adminNote)}</b>）` : '';
    $('#shiftExisting').innerHTML = m.editing
      ? `<div class="existing">目前登記：<b>${esc(m.editing.name)} — ${CFG.TIER[m.editing.tier].name} ${range(m.editing.start, m.editing.end)}</b> ${noteInfo}<br>可改選其他時段${adm ? '或更換人員' : ''}後按「儲存修改」，或按「取消此班」。</div>`
      : '';

    const adminPanel = $('#adminShiftPanel');
    adminPanel.classList.toggle('hidden', !adm);
    if (adm) {
      const names = Array.from(new Set([...state.data.members, ...(m.editing ? [m.editing.name] : [])]));
      const sel = $('#shiftTargetMember');
      sel.innerHTML = names.map(n => `<option value="${esc(n)}" ${n === m.targetName ? 'selected' : ''}>${esc(n)}${n !== state.session.name ? '（管理員異動）' : ''}</option>`).join('');
      const locked2 = isT2Locked(m.date);
      $('#t2LockStatusText').textContent = `當日二車狀態：${locked2 ? '🔒 已鎖定（隊員無法選填）' : '🔓 開放排班中'}`;
      const btn = $('#toggleDayT2Btn');
      btn.textContent = locked2 ? '🔓 開放當日二車' : '🔒 鎖定當日二車';
      btn.onclick = () => toggleT2DateLock(m.date, !locked2);
    }

    const locked2 = isT2Locked(m.date);
    $('#tier2TabBtn').classList.toggle('locked-tab', locked2);
    $('#tier2TabSub').textContent = locked2 ? '🔒 已鎖定未開放' : '08:00–18:00（4小時/班）';

    document.querySelectorAll('.tier-tab').forEach(b => b.classList.toggle('active', Number(b.dataset.tier) === m.tier));

    // 移除已被他人搶先或被鎖定的已選時段
    m.selections = m.selections.filter(sel => {
      const tierOpts = slotOptions(m.date, sel.tier, m.editing, m.targetName);
      const found = tierOpts.find(o => o.start === sel.start);
      return found && !found.taken;
    });

    // 渲染多選提示列
    const multiBar = $('#multiSelectBar');
    if (m.editing) {
      multiBar.innerHTML = `<span>請點選下方要更改的新時段（單選）</span>`;
    } else {
      const chips = m.selections.map((sel, idx) =>
        `<span class="multi-chip" data-del-idx="${idx}" title="點擊取消此班">${CFG.TIER[sel.tier].name} ${range(sel.start, sel.start + CFG.SHIFT_HOURS)} ✕</span>`
      ).join('');
      multiBar.innerHTML = `<span>💡 同一天可點選多班（已選 <b>${m.selections.length} / ${MAX_MULTI_SHIFTS}</b> 班）</span>
        <div class="multi-chips">${chips || '<span class="muted">尚未選取時段</span>'}</div>`;
    }

    const opts = slotOptions(m.date, m.tier, m.editing, m.targetName);
    $('#slotList').innerHTML = opts.map(o => {
      const selIdx = m.selections.findIndex(s => s.tier === m.tier && s.start === o.start);
      const isSelected = selIdx !== -1;
      const overlapsSelected = !isSelected && m.selections.some(s => isIntervalsOverlap(s.start, o.start));
      let subText = '可登記';
      if (o.taken) subText = esc(o.taken);
      else if (isSelected) subText = m.editing ? '✓ 已選取' : `✓ 已選（第 ${selIdx + 1} 班）`;
      else if (overlapsSelected) subText = '點擊可改選此時段';
      return `<button type="button" class="slot-opt ${isSelected ? 'selected' : ''} ${overlapsSelected && !o.taken ? 'overlap-selected' : ''}" data-start="${o.start}" ${o.taken ? 'disabled' : ''}>
        ${range(o.start, o.end)}<small>${subText}</small></button>`;
    }).join('');

    const curOne = m.selections[0];
    const unchanged = m.editing && curOne && m.editing.tier === curOne.tier && m.editing.start === curOne.start && m.editing.name === m.targetName;
    $('#saveShiftBtn').disabled = m.selections.length === 0 || unchanged;
    $('#saveShiftBtn').textContent = m.editing
      ? '儲存修改'
      : (m.selections.length > 1 ? `送出登記（${m.selections.length} 班）` : '送出登記');
    $('#deleteShiftBtn').classList.toggle('hidden', !m.editing);
  }

  $('#multiSelectBar').addEventListener('click', e => {
    const chip = e.target.closest('[data-del-idx]');
    if (!chip || !state.modal) return;
    const idx = Number(chip.dataset.delIdx);
    state.modal.selections.splice(idx, 1);
    renderModal();
  });

  $('#shiftTargetMember').addEventListener('change', e => {
    if (!state.modal) return;
    state.modal.targetName = e.target.value;
    renderModal();
  });

  $('#tierTabs').addEventListener('click', e => {
    const b = e.target.closest('.tier-tab');
    if (!b || !state.modal) return;
    state.modal.tier = Number(b.dataset.tier);
    if (state.modal.editing) state.modal.selections = [];
    renderModal();
  });

  $('#slotList').addEventListener('click', e => {
    const b = e.target.closest('.slot-opt');
    if (!b || b.disabled || !state.modal) return;
    const m = state.modal;
    const st = Number(b.dataset.start);

    if (m.editing) {
      m.selections = [{ tier: m.tier, start: st }];
      return renderModal();
    }

    const existingIdx = m.selections.findIndex(s => s.tier === m.tier && s.start === st);
    if (existingIdx !== -1) {
      // 再次點擊已選時段 → 取消選取
      m.selections.splice(existingIdx, 1);
      return renderModal();
    }

    // 檢查是否與已選時段重疊（同一人同時段不可重疊）
    const nonOverlapping = m.selections.filter(s => !isIntervalsOverlap(s.start, st));
    if (nonOverlapping.length < m.selections.length) {
      // 若點擊了與已選時段重疊的時段，直接替換該重疊時段
      nonOverlapping.push({ tier: m.tier, start: st });
      nonOverlapping.sort((a, b) => a.start - b.start || a.tier - b.tier);
      m.selections = nonOverlapping;
      return renderModal();
    }

    if (m.selections.length >= MAX_MULTI_SHIFTS) {
      return toast(`同一天單次最多可選填 ${MAX_MULTI_SHIFTS} 班`, 'err');
    }
    m.selections.push({ tier: m.tier, start: st });
    m.selections.sort((a, b) => a.start - b.start || a.tier - b.tier);
    renderModal();
  });

  $('#saveShiftBtn').onclick = async () => {
    const m = state.modal;
    if (!m || m.selections.length === 0 || state.busy) return;
    state.busy = true;
    try {
      const payloadTarget = isAdmin() ? m.targetName : undefined;
      if (m.editing) {
        const sel = m.selections[0];
        const res = await call('updateShift', { id: m.editing.id, tier: sel.tier, start: sel.start, targetName: payloadTarget });
        toast(res.adminNote ? '已更新排班（已註記：管理員異動）' : '已更新排班', 'ok');
      } else {
        const added = [];
        for (const sel of m.selections) {
          const res = await call('addShift', { date: m.date, tier: sel.tier, start: sel.start, targetName: payloadTarget });
          added.push(`${CFG.TIER[sel.tier].name} ${range(sel.start, sel.start + CFG.SHIFT_HOURS)}`);
        }
        toast(`登記成功（${added.length} 班）：${added.join('、')}`, 'ok');
      }
      closeModal('#shiftModal');
      await loadMonth(true);
    } catch (e) {
      toast(e.message, 'err');
      await loadMonth(true);   // 取得最新狀態（例如已被他人搶先）
    } finally { state.busy = false; }
  };

  $('#deleteShiftBtn').onclick = async () => {
    const m = state.modal;
    if (!m || !m.editing || state.busy) return;
    if (!confirm(`確定取消 ${m.date} ${m.editing.name} ${CFG.TIER[m.editing.tier].name} ${range(m.editing.start, m.editing.end)} 的班？`)) return;
    state.busy = true;
    try {
      await call('deleteShift', { id: m.editing.id });
      toast('已取消此班', 'ok');
      closeModal('#shiftModal');
      await loadMonth(true);
    } catch (e) { toast(e.message, 'err'); await loadMonth(true); }
    finally { state.busy = false; }
  };

  function closeModal(sel) {
    $(sel).classList.add('hidden');
    if (sel === '#shiftModal') state.modal = null;
  }
  document.querySelectorAll('.modal').forEach(md => {
    md.addEventListener('click', e => {
      if (md.dataset.locked) return;
      if (e.target === md || e.target.closest('[data-close]')) closeModal('#' + md.id);
    });
  });
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    document.querySelectorAll('.modal:not(.hidden)').forEach(md => { if (!md.dataset.locked) closeModal('#' + md.id); });
  });

  // ---------------- 變更 PIN ----------------
  function openPinModal(force) {
    const md = $('#pinModal');
    $('#pinForm').reset();
    $('#pinNotice').textContent = force ? '首次登入，請先變更您的個人 PIN 碼（請勿使用預設值）。' : 'PIN 碼為 4～8 位數字，請妥善保管。';
    md.querySelectorAll('[data-close]').forEach(b => b.classList.toggle('hidden', !!force));
    if (force) md.dataset.locked = '1'; else delete md.dataset.locked;
    md.classList.remove('hidden');
  }
  $('#pinBtn').onclick = () => openPinModal(false);
  $('#pinForm').addEventListener('submit', async e => {
    e.preventDefault();
    const oldPin = $('#oldPin').value, newPin = $('#newPin').value;
    if (newPin !== $('#newPin2').value) return toast('兩次輸入的新 PIN 碼不一致', 'err');
    try {
      await call('changePin', { oldPin, newPin });
      delete $('#pinModal').dataset.locked;
      closeModal('#pinModal');
      toast('PIN 碼已變更', 'ok');
    } catch (err) { toast(err.message, 'err'); }
  });

  // ---------------- 人員管理（管理員） ----------------
  $('#adminBtn').onclick = async () => { $('#adminModal').classList.remove('hidden'); await loadMembers(); };

  async function loadMembers() {
    try {
      const list = await call('adminListMembers');
      $('#memberTable tbody').innerHTML = list.map(m => `
        <tr class="${m.active ? '' : 'inactive'}">
          <td>${esc(m.name)}</td>
          <td class="${m.role === 'admin' ? 'role-admin' : ''}">${m.role === 'admin' ? '管理員' : '成員'}</td>
          <td>${m.active ? '啟用' : '已刪除（停用）'}</td>
          <td>
            ${m.active
              ? `<button class="btn btn-light btn-sm" data-act="reset" data-name="${esc(m.name)}">重設 PIN</button>
                 <button class="btn btn-danger btn-sm" data-act="remove" data-name="${esc(m.name)}">刪除</button>`
              : `<button class="btn btn-ghost-dark btn-sm" data-act="restore" data-name="${esc(m.name)}">恢復</button>`}
          </td>
        </tr>`).join('');
    } catch (e) { toast(e.message, 'err'); }
  }

  $('#memberTable').addEventListener('click', async e => {
    const b = e.target.closest('button[data-act]');
    if (!b) return;
    const name = b.dataset.name, act = b.dataset.act;
    try {
      if (act === 'remove') {
        if (!confirm(`確定刪除「${name}」？（帳號停用，歷史排班紀錄保留）`)) return;
        await call('adminRemoveMember', { name });
        toast(`已刪除 ${name}`, 'ok');
      } else if (act === 'restore') {
        await call('adminRestoreMember', { name });
        toast(`已恢復 ${name}`, 'ok');
      } else if (act === 'reset') {
        const pin = prompt(`將「${name}」的 PIN 重設為（4～8 位數字，留空則為 0000）：`, '');
        if (pin === null) return;
        await call('adminResetPin', { name, pin: pin || '0000' });
        toast(`已重設 ${name} 的 PIN，下次登入需變更`, 'ok');
      }
      await loadMembers();
      loadMonth(true);
    } catch (err) { toast(err.message, 'err'); }
  });

  $('#addMemberForm').addEventListener('submit', async e => {
    e.preventDefault();
    const name = $('#newMemberName').value.trim();
    try {
      await call('adminAddMember', { name, pin: $('#newMemberPin').value || '0000', role: $('#newMemberRole').value });
      toast(`已新增 ${name}`, 'ok');
      e.target.reset();
      await loadMembers();
      loadMonth(true);
    } catch (err) { toast(err.message, 'err'); }
  });

  // ---------------- 啟動 ----------------
  $('#unitName').textContent = CFG.UNIT_NAME;
  window.addEventListener('storage', () => { if (state.session) loadMonth(true); });
  if (state.session) enterApp().catch(() => logout());
  else initLogin();
})();

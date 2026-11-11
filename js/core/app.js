/* 셸: 홈 화면, 과제 세션(작업 표시줄·지시문 창·진행 저장), 출력물 5매 미리 보기와 자동 채점, 단축키 표. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const { h } = OA.ui;
  const DATA = G.OA_DATA;
  const KEY = 'oa-sim-v1';
  const APPS = [['excel', 'Excel', 'X', '#217346'], ['access', 'Access', 'A', '#a4373a'], ['ppt', 'PowerPoint', 'P', '#c43e1c']];
  const root = document.getElementById('root');
  OA.apps = OA.apps || {};

  function store(k, v) { try { if (v === undefined) return JSON.parse(localStorage.getItem(KEY + ':' + k) || 'null'); if (v === null) localStorage.removeItem(KEY + ':' + k); else localStorage.setItem(KEY + ':' + k, JSON.stringify(v)); } catch (e) { return null; } return null; }

  // ------------------------------------------------------------------ 시작 상태
  function rawWorkbook(sp) {
    const XL = OA.xl, wb = XL.newWorkbook();
    sp.head.slice(0, 6).forEach((h2, j) => XL.setInput(wb, 0, 4, 2 + j, h2));
    sp.data.forEach((row, i) => row.forEach((v, j) => XL.setInput(wb, 0, 5 + i, 2 + j, String(v))));
    XL.applyStyle(wb, 0, OA.util.parseRange('G5:G16'), { nf: '0%' });
    return wb;
  }
  function startState(set, mode, chapter) {
    const st = { wb: null, db: null, pres: null, start: {} };
    const ch = chapter !== null && chapter !== undefined ? set.chapters[chapter] : null;
    const isCh = n => ch && ch.ch === n;
    if (set.ch15) {
      const XL = OA.xl, wb = XL.newWorkbook();
      XL.renameSheet(wb, 0, '작업표');
      const c = set.chapters[0];
      const t6 = c.task.find(b => b.t === 'table' && b.caption.startsWith('표 15-6.'));
      t6.rows[0].slice(1).forEach((x, j) => XL.setInput(wb, 0, 4, 2 + j, x));
      t6.rows.slice(1).forEach((row, i) => row.slice(1).forEach((v, j) => XL.setInput(wb, 0, 5 + i, 2 + j, v)));
      const t14 = c.task.find(b => b.t === 'table' && b.caption.startsWith('표 15-14.'));
      XL.setInput(wb, 0, 4, 16, '설치일'); XL.setInput(wb, 0, 4, 17, '점검일');
      t14.rows.forEach((row, i) => { XL.setInput(wb, 0, 5 + i, 16, row[2]); XL.setInput(wb, 0, 5 + i, 17, row[3]); });
      st.wb = wb; st.active = 'excel';
      return st;
    }
    if (isCh(11)) { st.wb = OA.solutions.excelSet(set, 10); st.start.excel = '10장 완성 상태(ch11_시작.xlsx)'; }
    else if (mode === 'file' || isCh(10) && mode === 'file') { st.wb = rawWorkbook(set.sp); st.start.excel = '원자료가 입력된 시작 파일'; }
    else { st.wb = OA.xl.newWorkbook(); st.start.excel = '새 통합 문서'; }
    if (isCh(13)) { st.db = OA.solutions.accessSet(set, 12); st.start.access = '12장 완성 데이터베이스(ch13_시작.accdb)'; }
    else if (mode === 'file') { st.db = OA.acc.newDatabase((set.id === 'equip' ? 'ch12' : set.id) + '_작업'); st.start.access = '빈 데이터베이스(시작 파일)'; }
    else st.start.access = '새 데이터베이스부터';
    if (mode === 'file') {
      const P = OA.ppt, pres = P.newPresentation();
      P.setSlideSize(pres, 'standard', 'fit'); P.setLayout(pres, 0, 'blank'); P.addSlide(pres, 0);
      st.pres = pres; st.start.ppt = '4:3 빈 슬라이드 2장(시작 파일)';
    } else st.start.ppt = '새 프레젠테이션부터';
    st.active = ch ? ch.app : 'excel';
    return st;
  }

  // ------------------------------------------------------------------ 홈
  // 첫 화면: 표지와 같은 노랑 머리 띠·「2권 실기」 핵심어, 과제 선택은 표지 2×2 원색 알약(초록 표계산·파랑 자료처리·빨강 시상·보라 모의)
  const chTitle = c => c.title.replace(/^\d+장\s*/, '');
  const chTag = (set, c) => c.ch === 16 ? set.title : 'CHAPTER ' + c.ch;
  function home() {
    document.title = '사무자동화산업기사 2권 실기 · 실기 시뮬레이터';
    root.innerHTML = '';
    const saved = id => store('session:' + id);
    const setOf = id => DATA.sets.find(s => s.id === id);
    const confirmFresh = (s, go) => { if (!s) return go(); OA.ui.msg({ title: '새로 시작', text: '저장된 진행 상태를 지우고 처음부터 시작합니다.', buttons: ['처음부터', '취소'], icon: 'q' }).then(r => { if (r === '처음부터') go(); }); };
    // 시작 단추: 이어서 하기(저장된 상태가 있을 때) · 처음부터 · 실습 파일로 시작
    const acts = (set, chIdx, firstLabel) => {
      const sid = set.id + (chIdx !== null ? ':' + chIdx : '');
      const s = saved(sid);
      const ch = chIdx !== null ? set.chapters[chIdx] : null;
      const fileOk = !ch || ch.startFrom === null || ch.startFrom === undefined;
      return h('div', { class: 'acts' },
        s ? h('button', { class: 'sbtn ink', 'data-start': sid + ':resume', onclick: () => openSession(set, chIdx, null, s) }, '이어서 하기') : null,
        h('button', { class: 'sbtn' + (s ? '' : ' ink'), 'data-start': sid + ':blank', onclick: () => confirmFresh(s, () => openSession(set, chIdx, 'blank')) }, firstLabel || '처음부터'),
        fileOk ? h('button', { class: 'sbtn', 'data-start': sid + ':file', onclick: () => confirmFresh(s, () => openSession(set, chIdx, 'file')) }, '실습 파일로 시작') : null);
    };
    const sheet = k => h('span', { class: 'sheet' }, k);
    const row = (badge, title, subs, set, chIdx, firstLabel) => {
      const sid = set.id + (chIdx !== null ? ':' + chIdx : '');
      return h('div', { class: 'prow' }, h('div', { class: 'bd' }, badge),
        h('div', {}, h('div', { class: 'tt' }, title), h('div', { class: 'ss' }, subs, saved(sid) ? h('span', { class: 'saved' }, '진행 중') : null)),
        acts(set, chIdx, firstLabel));
    };
    const equip = setOf('equip'), c15 = setOf('ch15');
    const chRow = (i, note) => { const c = equip.chapters[i]; return row(String(c.ch), chTitle(c), [sheet(c.sheet), h('span', {}, note)], equip, i, c.startFrom ? '시작' : null); };
    const mocks = DATA.sets.filter(s => /^mock/.test(s.id));
    const block = (color, lab, title, aside, rows) => h('section', { class: 'pblock ' + color },
      h('div', { class: 'pill' }, h('span', { class: 'lab' }, lab), h('span', { class: 'pt' }, title), h('span', { class: 'pa' }, aside)),
      h('div', { class: 'prows' }, rows));
    const pills = h('div', { class: 'pgrid' },
      block('green', '2매', '작업표 · 피벗테이블', 'Excel', [
        chRow(0, '원자료 입력부터'), chRow(1, '10장 완성 상태에서 시작'),
        c15 ? row('15', '유형 과제 A~F', [sheet('수식'), h('span', {}, '출력물 없음')], c15, 0, '과제 시작') : null]),
      block('blue', '2매', '조회화면 · 보고서', 'Access', [chRow(2, '빈 데이터베이스부터'), chRow(3, '12장 완성 상태에서 시작')]),
      block('red', '1매', '슬라이드 유인물', 'PowerPoint', [chRow(4, '4:3 빈 슬라이드부터')]),
      block('violet', '모의', `실기 모의평가 ${mocks.length}세트`, '2시간 · 5매', [
        row('10–14', equip.title, [sheet('5매'), h('span', {}, '10~14장 과제 이어서')], equip, null),
        ...mocks.map((m, i) => row(String(i + 1), m.sub.split(',')[0], [sheet('5매'), h('span', {}, m.title + ' · 2시간')], m, null))]));
    root.appendChild(h('div', { class: 'home' },
      h('header', { class: 'hb' }, h('div', { class: 'hb-in' },
        h('div', { class: 'hb-left' },
          h('div', { class: 'hb-kicker' }, 'Industrial Engineer Office Automation · 2027'),
          h('h1', { class: 'hb-title' }, h('span', { class: 't1' }, '사무자동화산업기사'), h('span', { class: 't2' }, '2권 실기')),
          h('div', { class: 'hb-ribbon' }, h('span', {}, '실기 시뮬레이터 · 원자료 입력부터 5매 출력까지')),
          h('p', { class: 'hb-lead' }, '『2027 사무자동화산업기사 2권 실기』 10~16장 과제를 Excel·Access·PowerPoint 화면에서 처음부터 출력물 5매 미리 보기까지 연습합니다. 결과는 책의 실측 검산값으로 채점합니다.')),
        h('div', { class: 'hb-badges' },
          h('div', { class: 'burst' }, h('div', {}, h('span', {}, '출제기준', h('b', {}, '2026–2028'), '반영'))),
          h('div', { class: 'disc' }, h('span', {}, 'Office 없이', h('b', {}, '5매 출력'), '까지 연습'))))),
      h('main', { class: 'home-main' },
        h('div', { class: 'sechead' }, h('span', { class: 'no' }, '01'), h('h2', {}, '과제 선택'), h('span', { class: 'tchip' }, '작업형'), h('span', { class: 'aside' }, '진행 상태는 이 브라우저에만 저장됩니다')),
        pills,
        h('div', { class: 'sechead' }, h('span', { class: 'no' }, '02'), h('h2', {}, '사용 안내')),
        h('div', { class: 'guide' },
          h('div', { class: 'tagbox' }, h('span', { class: 'tg' }, '연습 방법'), h('ol', {},
            h('li', {}, h('b', {}, '처음부터'), ': 빈 통합 문서, Access 새 데이터베이스 화면, PowerPoint 시작 화면에서 출발합니다.'),
            h('li', {}, h('b', {}, '실습 파일로 시작'), ': 책의 시작 파일과 같은 상태(원자료가 입력된 시트, 빈 데이터베이스, 4:3 빈 슬라이드 2장)에서 출발합니다.'),
            h('li', {}, '작업을 마치면 위 막대의 ', h('b', {}, '출력물 5매 미리 보기·채점'), '으로 점검 항목별 결과를 확인합니다.'))),
          h('div', {},
            h('div', { class: 'tagbox blue' }, h('span', { class: 'tg' }, '채점 기준'), h('p', {}, '부록 표 A-1 출력물 5매 점검표 30항목과 15장 과제 A~F. 값은 책의 실측 검산값과 대조합니다.')),
            h('div', { class: 'tagbox burg' }, h('span', { class: 'tg' }, '주의'), h('p', {}, '탭·그룹·단추 이름과 배치는 Office 2021 한국어판 기준이고, 아이콘은 단순 도형으로 대신했습니다. 시험장 Office의 세부 모양은 다를 수 있으므로 메뉴 위치를 익히는 연습용으로 씁니다.')))),
        h('div', { class: 'sechead' }, h('span', { class: 'no' }, '03'), h('h2', {}, '단축키'), h('span', { class: 'tchip' }, '표 9-4'), h('span', { class: 'aside' }, '이 표의 키만 받습니다')),
        shortcutTable(),
        h('footer', { class: 'home-foot' },
          h('button', { class: 'sbtn burg', onclick: () => OA.ui.msg({ title: '진행 상태 초기화', text: '모든 과제의 저장된 진행 상태를 지웁니다.', buttons: ['지우기', '취소'] }).then(r => { if (r === '지우기') { Object.keys(localStorage).filter(k => k.startsWith(KEY)).forEach(k => localStorage.removeItem(k)); home(); } }) }, '진행 상태 모두 지우기'),
          h('span', {}, '글꼴 Pretendard(SIL OFL 1.1) · SQL 엔진 sql.js(MIT) · 아이콘은 직접 그린 도형')))));
  }
  const SHORTCUTS = [['공통', '저장', 'Ctrl+S'], ['공통', '실행 취소·다시 실행', 'Ctrl+Z·Ctrl+Y'], ['공통', '복사·붙여 넣기', 'Ctrl+C·Ctrl+V'], ['Excel', '셀 서식 대화상자', 'Ctrl+1', 'b'], ['Excel', '셀 편집(수식 입력줄)', 'F2'],
    ['Excel', '참조 유형 전환(상대·절대)', 'F4'], ['Excel', '선택 범위 아래로 채우기', 'Ctrl+D'], ['Excel', '백분율 서식', 'Ctrl+Shift+%'], ['Excel', '통화 서식', 'Ctrl+Shift+$'], ['Excel', '현재 영역 선택', 'Ctrl+A'],
    ['Excel', '새 워크시트', 'Shift+F11'], ['Excel', '다음·이전 시트', 'Ctrl+Page Down·Ctrl+Page Up', 'b'], ['Excel', '선택 범위의 차트(포함)', 'Alt+F1'], ['Access', '인쇄', 'Ctrl+P'], ['Access', '속성 시트 표시·숨기기', 'F4 또는 Alt+Enter'],
    ['Access', '필드 목록 창', 'Alt+F8'], ['Access', '탐색 창 표시·숨기기', 'F11', 'f'], ['Access', '확대/축소 상자(긴 식 입력)', 'Shift+F2'], ['Access', '디자인 보기에서 컨트롤 이동·크기', '화살표 키·Shift+화살표 키'], ['Access', '레코드 저장', 'Shift+Enter'],
    ['PowerPoint', '인쇄', 'Ctrl+P'], ['PowerPoint', '새 슬라이드', 'Ctrl+M'], ['PowerPoint', '개체 복제', 'Ctrl+D'], ['PowerPoint', '그룹·그룹 해제', 'Ctrl+G·Ctrl+Shift+G'], ['PowerPoint', '선택 창', 'Alt+F10'],
    ['PowerPoint', '안내선·눈금선 표시', 'Alt+F9·Shift+F9'], ['PowerPoint', '글꼴 크기 늘리기·줄이기', 'Ctrl+Shift+>·Ctrl+Shift+<'], ['PowerPoint', '개체를 한 단계 앞·뒤로', 'Ctrl+Shift+]·Ctrl+Shift+['], ['PowerPoint', '서식 복사·서식 붙여 넣기', 'Ctrl+Shift+C·Ctrl+Shift+V'], ['PowerPoint', '모든 개체 선택', 'Ctrl+A']];
  function shortcutTable() {
    const keys = k => k.split(/(·| 또는 )/).map(x => x === '·' || x === ' 또는 ' ? x : h('kbd', {}, x));
    return h('div', {},
      h('table', { class: 'stable' }, h('tr', {}, h('th', { style: { width: '120px' } }, '프로그램'), h('th', {}, '동작'), h('th', {}, '키'), h('th', {}, '비고')),
        SHORTCUTS.map(([p, a, k, note]) => h('tr', {}, h('td', {}, p), h('td', {}, a), h('td', {}, keys(k)),
          h('td', {}, note === 'b' ? '웹 브라우저가 탭 전환 키로 먼저 처리할 수 있음. 오른쪽 단추 메뉴(셀 서식)·시트 탭 누르기로 같은 동작' : note === 'f' ? '웹 브라우저 전체 화면 키와 겹칠 수 있음. 탐색 창의 « 단추로 같은 동작' : '')))));
  }

  // ------------------------------------------------------------------ 세션
  let S = null;
  function openSession(set, chIdx, mode, saved) {
    const sid = set.id + (chIdx !== null ? ':' + chIdx : '');
    let st;
    if (saved) {
      st = { wb: saved.wb ? OA.xl.fromJSON(saved.wb) : null, db: saved.db, pres: saved.pres, start: saved.start || {}, active: saved.active || 'excel' };
    } else st = startState(set, mode, chIdx);
    S = { sid, set, chIdx, ch: chIdx !== null ? set.chapters[chIdx] : null, mode: saved ? saved.mode : mode,
      models: { excel: st.wb, access: st.db, ppt: st.pres }, start: st.start,
      status: saved && saved.status ? saved.status : { excel: { dirty: false, version: 0, saved: -1, printed: null }, access: { dirty: false, version: 0, saved: -1, printed: null }, ppt: { dirty: false, version: 0, saved: -1, printed: null } },
      prints: saved && saved.prints ? saved.prints : {}, elapsed: saved ? saved.elapsed || 0 : 0, active: st.active, inst: {} };
    renderSession();
    persist();
  }
  const appsOf = () => S.ch ? [S.ch.app] : (S.set.ch15 ? ['excel'] : ['excel', 'access', 'ppt']);
  function renderSession() {
    root.innerHTML = '';
    document.title = S.set.title + (S.ch ? ' · ' + S.ch.title : '') + ' - 실기 시뮬레이터';
    // 위 막대: 내지 장 머리처럼 버건디 장 꼬리표 + 굵은 과제 이름, 오른쪽에 시험 시계(고정 폭 숫자)
    const isMock = /^mock/.test(S.set.id);
    const tagText = S.ch ? chTag(S.set, S.ch) : isMock ? S.set.title : '10–14장';
    const ttlText = S.ch ? chTitle(S.ch) : isMock ? S.set.sub.split(',')[0] : S.set.title;
    const setName = S.ch ? S.set.title : (isMock ? '2시간 · 출력물 5매' : '출력물 5매');
    const timer = h('span', { class: 'clock', title: '시험 시간 2시간 기준 경과 시간' }, h('span', {}, '경과'), h('span', { class: 't' }, '0:00:00'), h('span', { class: 'lim' }, '/ 2:00:00'));
    const tpBtn = h('button', { class: 'sbtn tptoggle on', title: '지시문·작업 순서·따라 하기 창 보이기·숨기기', onclick: () => { panel.hidden = !panel.hidden; tpBtn.classList.toggle('on', !panel.hidden); } }, '지시문');
    const top = h('div', { class: 'topbar' },
      h('button', { class: 'sbtn ghost', onclick: () => { persist(); stopTimer(); home(); } }, '◀ 홈'),
      h('span', { class: 'chtag' + (isMock || (S.ch && S.ch.ch === 16) ? ' violet' : '') }, tagText),
      h('span', { class: 'ttl' }, ttlText), h('span', { class: 'setname' }, setName),
      tpBtn,
      h('span', { class: 'sp' }), timer,
      h('button', { class: 'sbtn blue', onclick: () => review() }, S.set.ch15 ? '채점' : '출력물 5매 미리 보기·채점'),
      h('button', { class: 'sbtn', onclick: () => OA.ui.dialog({ title: '단축키(표 9-4)', width: 820, body: shortcutTable(), buttons: [{ label: '닫기', primary: true }] }) }, '단축키'),
      h('button', { class: 'sbtn burg', onclick: () => OA.ui.msg({ title: '과제 초기화', text: '이 과제의 진행 상태를 지우고 시작 상태로 되돌립니다.', buttons: ['초기화', '취소'] }).then(r => { if (r === '초기화') { store('session:' + S.sid, null); openSession(S.set, S.chIdx, S.mode || 'blank'); } }) }, '초기화'));
    const panel = taskPanel();
    const appsHost = h('div', { class: 'apps' });
    const work = h('div', { class: 'workarea' }, panel, appsHost);
    const bar = h('div', { class: 'taskbar' });
    root.appendChild(h('div', { class: 'session' }, top, work, bar));
    S.appsHost = appsHost; S.bar = bar;
    for (const a of appsOf()) {
      const factory = OA.apps[a];
      if (!factory) continue;
      const win = h('div', { class: 'appwin', hidden: true });
      appsHost.appendChild(win);
      S.inst[a] = factory.create(win, ctxFor(a));
    }
    if (!appsOf().includes(S.active)) S.active = appsOf()[0];
    switchApp(S.active);
    startTimer(timer);
  }
  function renderBar() {
    if (!S || !S.bar) return;
    S.bar.innerHTML = '';
    for (const [id, label, ch, color] of APPS) {
      if (!S.inst[id]) continue;
      const st = S.status[id];
      const name = S.inst[id].fileName ? S.inst[id].fileName() : label;
      S.bar.appendChild(h('button', { class: S.active === id ? 'on' : '', onclick: () => switchApp(id), title: label },
        h('span', { class: 'ico', style: { background: color } }, ch), h('span', {}, name), st && st.dirty ? h('span', { class: 'dirty', title: '저장하지 않은 변경' }, '●') : null,
        st && st.printed ? h('span', { class: 'printed', title: '출력물을 인쇄함' }, '인쇄함') : null));
    }
    S.bar.appendChild(h('span', { class: 'msg' }, S.lastMsg || ''));
  }
  function switchApp(id) {
    S.active = id;
    for (const [a, inst] of Object.entries(S.inst)) inst.el.hidden = a !== id;
    S.inst[id] && S.inst[id].onShow && S.inst[id].onShow();
    renderBar();
    persist();
  }
  let saveT = null;
  function persist() {
    if (!S) return;
    clearTimeout(saveT);
    saveT = setTimeout(() => {
      const data = { mode: S.mode, wb: S.models.excel, db: S.models.access, pres: S.models.ppt, status: S.status, prints: S.prints, elapsed: S.elapsed, active: S.active, start: S.start };
      store('session:' + S.sid, data);
    }, 400);
  }
  addEventListener('beforeunload', () => { if (S) { clearTimeout(saveT); store('session:' + S.sid, { mode: S.mode, wb: S.models.excel, db: S.models.access, pres: S.models.ppt, status: S.status, prints: S.prints, elapsed: S.elapsed, active: S.active, start: S.start }); } });
  let tick = null;
  function startTimer(el) {
    stopTimer();
    const fmt = s => [Math.floor(s / 3600), Math.floor(s / 60) % 60, s % 60].map((x, i) => i ? String(x).padStart(2, '0') : x).join(':');
    const t = el.querySelector('.t') || el;
    const upd = () => { t.textContent = fmt(S.elapsed); el.classList.toggle('over', S.elapsed > 7200); };
    upd();
    tick = setInterval(() => { if (document.hidden) return; S.elapsed++; upd(); if (S.elapsed % 15 === 0) persist(); }, 1000);
  }
  function stopTimer() { clearInterval(tick); }
  function ctxFor(app) {
    return {
      session: () => S, set: S.set, chapter: S.ch,
      model: () => S.models[app],
      setModel: m => { S.models[app] = m; changed(app); },
      changed: () => changed(app),
      saved: () => { const st = S.status[app]; st.saved = st.version; st.dirty = false; renderBar(); persist(); },
      printed: (key, info) => { const st = S.status[app]; st.printed = st.version; S.prints[key] = Object.assign({ at: Date.now(), version: st.version }, info || {}); S.lastMsg = `${key} 인쇄 완료`; renderBar(); persist(); OA.ui.toast(`인쇄했습니다(출력물 ${key}). 출력물 5매 미리 보기에서 확인할 수 있습니다.`); },
      startNote: S.start[app],
      msg: t => { S.lastMsg = t; renderBar(); },
    };
  }
  function changed(app) {
    const st = S.status[app];
    st.version++;
    st.dirty = st.version !== st.saved;
    if (st.printed !== null && st.printed !== undefined && st.version !== st.printed) st.printedDirty = true;
    renderBar();
    persist();
  }

  // ------------------------------------------------------------------ 지시문 창
  // 내지 형식: 장 머리 띠(버건디, CHAPTER 꼬리표) → 절 번호 상자 → 꼬리표 상자(주의 = 버건디, 검산 = 파랑, 작업 순서 = 회색)
  const tagboxKind = title => /주의/.test(title) ? 'burg' : /검산|핵심|정리|수식/.test(title) ? 'blue' : '';
  function blocksView(blocks) {
    const out = [];
    for (const b of blocks || []) {
      if (b.t === 'p') out.push(h('p', {}, b.text));
      else if (b.t === 'h') out.push(h('h4', {}, b.text));
      else if (b.t === 'box') out.push(h('div', { class: 'tagbox box ' + tagboxKind(b.title) }, h('span', { class: 'tg' }, b.title), b.lines.map(l => h('p', {}, l))));
      else if (b.t === 'table') out.push(h('div', {}, b.caption ? h('div', { class: 'cap' }, b.caption) : null, h('table', {}, h('tr', {}, b.head.map(x => h('th', {}, x))), b.rows.map(r => h('tr', {}, r.map(x => h('td', {}, x)))))));
    }
    return out;
  }
  function chHead(c) {
    const mock = c.ch === 16;
    return h('div', { class: 'chhead' + (mock ? ' violet' : '') }, h('span', { class: 'tag' }, chTag(S.set, c)), h('span', { class: 't' }, chTitle(c)));
  }
  function taskPanel() {
    const set = S.set;
    const tabs = [];
    const chs = S.ch ? [S.ch] : set.chapters;
    tabs.push(['지시문', () => h('div', {}, set.intro ? blocksView(set.intro) : null, chs.map(c => h('div', {}, chHead(c), blocksView(c.task),
      c.app === 'ppt' && set.pt ? set.pt.images.map((src, i) => h('div', {}, h('div', { class: 'cap' }, `제${i + 1} 슬라이드 원안`), h('img', { src, alt: `제${i + 1} 슬라이드 원안` }))) : null,
      c.formulas ? h('div', {}, h('div', { class: 'tagbox blue' }, h('span', { class: 'tg' }, '수식 표(원고 정답)'), h('p', {}, '아래 표는 책의 수식입니다. 먼저 직접 작성한 뒤 확인합니다.')), blocksView(c.formulas)) : null)))]);
    const orders = chs.filter(c => c.order && c.order.length);
    if (orders.length) tabs.push(['작업 순서', () => h('div', {}, orders.map(c => {
      const ids = c.order.map((l, i) => 'o' + c.ch + '_' + i);
      const prog = progressEl(ids);
      return h('div', {}, chHead(c), h('div', { class: 'tagbox' }, h('div', { style: { display: 'flex', alignItems: 'center' } }, h('span', { class: 'tg' }, '작업 순서'), prog.el),
        h('ol', { class: 'chk' }, c.order.map((l, i) => checkItem(ids[i], l, prog.update)))));
    }))]);
    const steps = chs.filter(c => c.steps && c.steps.length);
    if (steps.length) tabs.push(['따라 하기', () => h('div', {}, steps.map(c => {
      const groups = [];
      for (const s of c.steps) { const g = groups[groups.length - 1]; if (g && g.sec === s.sec) g.list.push(s); else groups.push({ sec: s.sec, list: [s] }); }
      return h('div', {}, chHead(c), groups.map(g => {
        const m = /^(\d+\.\d+)\s+(.*)$/.exec(g.sec) || [null, g.sec, ''];
        const ids = g.list.flatMap(s => s.items.map(it => 's' + c.ch + s.sec + s.title + it.n));
        const prog = progressEl(ids);
        return h('div', {}, h('div', { class: 'sechead' }, h('span', { class: 'no' }, m[1]), h('h3', {}, m[2]), prog.el),
          g.list.map(s => h('div', {}, h('h4', {}, '따라 하기: ' + s.title), h('ol', { class: 'chk' }, s.items.map(it => checkItem('s' + c.ch + s.sec + s.title + it.n, it.n + '. ' + it.text, prog.update))))));
      }));
    }))]);
    tabs.push(['시작 상태', () => h('div', {}, h('div', { class: 'tagbox blue' }, h('span', { class: 'tg' }, '시작 상태'),
      h('ul', {}, Object.entries(S.start || {}).map(([k, v]) => h('li', {}, h('b', {}, { excel: 'Excel', access: 'Access', ppt: 'PowerPoint' }[k]), ': ' + v)))),
      h('div', { class: 'tagbox' }, h('span', { class: 'tg' }, '저장'), h('p', {}, '파일 저장(Ctrl+S)은 시뮬레이터 안의 저장 표시이며, 작업 내용은 이 브라우저에 자동으로 보관됩니다.')))]);
    const bar = h('div', { class: 'tp-tabs' });
    const body = h('div', { class: 'tp-body' });
    const show = i => { [...bar.children].forEach((b, k) => b.classList.toggle('on', k === i)); body.innerHTML = ''; body.scrollTop = 0; body.appendChild(tabs[i][1]()); };
    tabs.forEach((t, i) => bar.appendChild(h('button', { onclick: () => show(i) }, t[0])));
    const el = h('div', { class: 'taskpanel' }, bar, body);
    show(0);
    return el;
  }
  // 확인란 진행 표시(완료 수 / 전체)
  function progressEl(ids) {
    const el = h('span', { class: 'progress' });
    const update = () => { el.textContent = `${ids.filter(id => store('chk:' + S.sid + ':' + id)).length} / ${ids.length} 완료`; };
    update();
    return { el, update };
  }
  function checkItem(id, text, onChange) {
    const k = 'chk:' + S.sid + ':' + id;
    const cb = h('input', { type: 'checkbox', checked: !!store(k), onchange: () => { store(k, cb.checked || null); if (onChange) onChange(); } });
    return h('li', {}, h('label', {}, cb, h('span', {}, text)));
  }

  // ------------------------------------------------------------------ 출력물 5매·채점
  const SHEETS = [['5-1', '작업표와 그래프', 'excel'], ['5-2', '피벗테이블', 'excel'], ['5-3', '조회화면', 'access'], ['5-4', '보고서', 'access'], ['5-5', '슬라이드 2장', 'ppt']];
  function sheetsInScope() {
    if (S.set.ch15) return [];
    if (S.ch) return S.ch.sheet.split('·').map(k => k.trim()).filter(k => /^5-\d$/.test(k));
    return SHEETS.map(s => s[0]);
  }
  function review() {
    const scope = sheetsInScope();
    const overlay = h('div', { class: 'review' });
    const close = () => overlay.remove();
    const scale = 0.62; // px/mm
    const pages = h('div', { class: 'sheets5' }, h('p', { class: 'note' }, '편철 순서(답안 표지 → 5-1 → 5-5)대로 놓은 현재 상태의 A4 미리 보기. 점선 상자는 자필 기입 위치(중앙 상단 비번호·수험번호·성명, 중앙 하단 쪽 번호).'));
    for (const [key, title, app] of SHEETS) {
      if (scope.length && !scope.includes(key)) continue;
      const inst = S.inst[app];
      let page;
      try { page = inst && inst.renderSheet ? inst.renderSheet(key, scale) : null; } catch (e) { page = null; console.error(e); }
      if (!page) page = OA.print.blank(scale, '(아직 없음)');
      OA.print.handMarks(page, scale, key);
      const pr = S.prints[key];
      const st = S.status[app];
      const stText = pr ? (st && st.printedDirty ? '인쇄 뒤 수정됨' : '인쇄함') : '인쇄 전';
      pages.appendChild(h('div', { class: 'sheet5' }, h('div', { class: 'lab' }, h('span', { class: 'k' }, key), title), page,
        h('div', { class: 'st' + (pr && !(st && st.printedDirty) ? ' ok' : pr ? ' warn' : '') }, stText)));
    }
    const state = { wb: S.models.excel, db: S.models.access, pres: S.models.ppt, status: S.status };
    let results = [];
    const tableHost = h('div');
    const LABEL = { pass: '통과', fail: '확인 필요', manual: '직접 확인', todo: '해당 없음' };
    const runGrade = () => {
      try {
        if (S.set.ch15) results = OA.gradeCh15 ? OA.gradeCh15(S.set, S.models.excel) : [];
        else results = OA.grade.grade(DATA, S.set, state, scope.concat(S.ch ? [] : ['공통']));
      } catch (e) { results = [{ sheet: '-', item: '채점 오류', sec: '', status: 'fail', notes: [e.message] }]; console.error(e); }
      const n = s => results.filter(r => r.status === s).length;
      const auto = results.length - n('manual');
      // 부록 표 A-1 의 항목 번호(1~30). 15장은 과제 순서 번호
      const noOf = (r, i) => { if (S.set.ch15) return String(i + 1); const k = DATA.checklist.findIndex(c => c.sheet === r.sheet && c.item === r.item); return k >= 0 ? String(k + 1) : '-'; };
      tableHost.innerHTML = '';
      tableHost.appendChild(h('div', { class: 'scorebar' }, h('span', { class: 'big' + (n('pass') < auto ? ' bad' : '') }, `${n('pass')} / ${auto}`), h('span', { class: 'lbl' }, '자동 점검 항목 통과'),
        n('manual') ? h('span', { class: 'rchip manual' }, `직접 확인 ${n('manual')}`) : null,
        h('span', { class: 'basis' }, S.set.ch15 ? '기준: 15.3 검산 상자와 실측값(verify_ch15_p2.json)' : '기준: 부록 표 A-1 출력물 5매 점검표, 값은 책의 실측 검산값')));
      tableHost.appendChild(h('table', { class: 'stable gradetbl' },
        h('tr', {}, h('th', {}, S.set.ch15 ? '번호' : 'A-1'), h('th', {}, '출력물'), h('th', {}, '점검 항목'), h('th', {}, '근거'), h('th', {}, '결과')),
        results.map((r, i) => h('tr', { class: r.status }, h('td', { class: 'no' }, noOf(r, i)), h('td', { class: 'sh' }, r.sheet),
          h('td', {}, r.item, r.notes && r.notes.length ? h('ul', { class: 'notes' }, r.notes.map(x => h('li', {}, x))) : null),
          h('td', { class: 'sec' }, r.sec), h('td', { class: 'res' }, h('span', { class: 'rchip ' + r.status }, LABEL[r.status]))))));
    };
    runGrade();
    const samples = S.set.samples && !S.set.ch15 ? h('div', { class: 'samples' },
      h('div', { class: 'sechead' }, h('span', { class: 'no' }, '견본'), h('h3', {}, '출력 견본(책의 실측 출력)')),
      S.set.samples.filter((src, i) => !scope.length || S.set.samples.length === 1 || scope.includes('5-' + (i + 1))).map(src => h('img', { src, alt: '출력 견본', style: { maxWidth: S.set.samples.length === 1 ? '100%' : '220px', marginRight: '8px' } }))) : null;
    overlay.appendChild(h('div', { class: 'review-h' }, h('span', { class: 'tag' }, S.set.ch15 ? '15.3' : '부록 A-1'),
      h('h2', {}, S.set.ch15 ? '15장 유형 과제 채점' : '출력물 5매 미리 보기와 자동 채점'),
      h('button', { class: 'sbtn on-dark', onclick: runGrade }, '다시 채점'), h('button', { class: 'sbtn on-dark fill', onclick: close }, '닫기')));
    overlay.appendChild(h('div', { class: 'review-b' }, h('div', { class: 'review-grid' + (S.set.ch15 ? ' solo' : '') }, S.set.ch15 ? null : pages, h('div', {}, tableHost, samples))));
    S.appsHost.parentNode.appendChild(overlay);
  }

  // ------------------------------------------------------------------ 키보드
  document.addEventListener('keydown', e => {
    if (!S || OA.ui.anyDialog()) return;
    if (document.querySelector('.review')) { if (e.key === 'Escape') document.querySelector('.review').remove(); return; }
    const inst = S.inst[S.active];
    if (inst && inst.handleKey) inst.handleKey(e);
  });

  OA.shell = { home, openSession, review, store, SHEETS, blocksView };
  addEventListener('DOMContentLoaded', () => {
    const boot = () => home();
    if (G.initSqlJs && G.SQL_WASM_B64 && OA.sqlx) {
      const bin = Uint8Array.from(atob(G.SQL_WASM_B64), c => c.charCodeAt(0));
      G.initSqlJs({ wasmBinary: bin }).then(SQL => OA.sqlx.setEngine(SQL)).catch(err => console.error('sql.js', err));
    }
    boot();
  });
})(typeof window !== 'undefined' ? window : globalThis);

/* 브라우저 화면 조작 도우미(확인용). 개발자 도구 콘솔이나 자동화 도구에서 이 파일 내용을 실행한 뒤 __T 로 화면을 조작한다.
   실제 사용자 조작과 같은 경로(마우스 누르기·입력 상자 입력·키 누르기·단추 누르기)로 이벤트를 보낸다. */
(function () {
  const T = window.__T = {};
  const vis = el => !!(el && el.getClientRects().length && !el.closest('[hidden]'));
  const app = () => [...document.querySelectorAll('.appwin')].find(e => !e.hidden);
  T.app = app;
  T.wait = ms => new Promise(r => setTimeout(r, ms || 30));
  function fire(el, type, opts) {
    const r = el.getBoundingClientRect();
    const o = Object.assign({ bubbles: true, cancelable: true, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2, button: 0 }, opts || {});
    el.dispatchEvent(new MouseEvent(type, o));
  }
  T.clickEl = (el, opts) => { fire(el, 'mousedown', opts); fire(el, 'mouseup', opts); fire(el, 'click', opts); };
  /* 화면에 보이는 단추·메뉴 항목을 글자로 찾아 누른다(정확히 일치 우선, 다음은 시작 일치) */
  T.find = (text, scope) => {
    const root = scope || document;
    const cands = [...root.querySelectorAll('button, .mi, .pset, .xl-tab, .obj, .lst div, label, .ac-doctab, .shapecell, .fd-sec-h')].filter(vis);
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const txt = e => norm([...e.childNodes].filter(n => n.nodeName.toLowerCase() !== 'svg' && !(n.classList && n.classList.contains('ck'))).map(n => n.textContent).join(''));   // 아이콘 SVG 안의 글자는 뺀다
    return cands.find(e => txt(e) === text || e.title === text) || cands.find(e => txt(e).startsWith(text) || (e.title || '').startsWith(text));
  };
  T.click = (text, scope) => { const el = T.find(text, scope); if (!el) throw new Error('찾지 못함: ' + text); T.clickEl(el); return el; };
  /* 리본 탭 이름과 같은 단추(예: 만들기 탭의 '테이블 디자인' 단추와 상황별 탭 '테이블 디자인')를 가리기 위해 리본 안에서만 찾는다 */
  T.rb = (text) => { const el = [...app().querySelectorAll('.ribbon button')].filter(vis).find(b => b.title === text || b.title.startsWith(text + ' (') || b.textContent.replace(/\s+/g, ' ').trim() === text); if (!el) throw new Error('리본 단추 없음: ' + text); T.clickEl(el); return el; };
  T.tab = (text) => { const el = [...app().querySelectorAll('.ribbon-tabs button')].find(b => b.textContent.trim() === text); if (!el) throw new Error('탭 없음: ' + text); T.clickEl(el); return el; };
  T.menu = (...path) => { for (const p of path) T.click(p, document.querySelector('.menu:last-of-type') && p !== path[0] ? document.body : undefined); };
  T.key = (key, mods, target) => {
    const t = target || document.activeElement || document.body;
    const code = { '%': 'Digit5', '$': 'Digit4' }[key] || (key.length === 1 ? (/[a-z]/i.test(key) ? 'Key' + key.toUpperCase() : /\d/.test(key) ? 'Digit' + key : key) : key);
    const ev = new KeyboardEvent('keydown', Object.assign({ key, code, bubbles: true, cancelable: true }, mods || {}));
    t.dispatchEvent(ev);
  };
  // ---------------- Excel
  T.xl = {};
  T.xl.cellCenter = (addr) => {
    const m = /^([A-Z]+)(\d+)$/.exec(addr);
    const w = app();
    const ch = [...w.querySelectorAll('.xl-colhdr')].find(e => e.firstChild && e.firstChild.textContent === m[1]);
    const rh = [...w.querySelectorAll('.xl-rowhdr')].find(e => e.textContent === m[2]);
    const grid = w.querySelector('.xl-grid'), wrap = w.querySelector('.xl-gridwrap');
    const g = grid.getBoundingClientRect();
    return { x: g.left + ch.offsetLeft + ch.offsetWidth / 2, y: g.top + rh.offsetTop + rh.offsetHeight / 2, wrap };
  };
  T.xl.scrollTo = (addr) => { const w = app(); const wrap = w.querySelector('.xl-gridwrap'); const m = /^([A-Z]+)(\d+)$/.exec(addr); const rh = [...w.querySelectorAll('.xl-rowhdr')].find(e => e.textContent === m[2]); if (rh) { const top = rh.offsetTop - 60; if (top < wrap.scrollTop || top > wrap.scrollTop + wrap.clientHeight - 80) wrap.scrollTop = Math.max(0, top); } wrap.dispatchEvent(new Event('scroll')); };
  T.xl.down = (addr, mods) => { T.xl.scrollTo(addr); const c = T.xl.cellCenter(addr); const o = Object.assign({ bubbles: true, cancelable: true, clientX: c.x, clientY: c.y, button: 0 }, mods || {}); c.wrap.dispatchEvent(new MouseEvent('mousedown', o)); window.dispatchEvent(new MouseEvent('mouseup', o)); };
  T.xl.select = (range, mods) => {
    const [a, b] = range.split(':');
    T.xl.scrollTo(a);
    const c1 = T.xl.cellCenter(a);
    const o1 = Object.assign({ bubbles: true, cancelable: true, clientX: c1.x, clientY: c1.y, button: 0 }, mods || {});
    c1.wrap.dispatchEvent(new MouseEvent('mousedown', o1));
    if (b) { const c2 = T.xl.cellCenter(b); window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: c2.x, clientY: c2.y })); window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: c2.x, clientY: c2.y })); }
    else window.dispatchEvent(new MouseEvent('mouseup', o1));
  };
  T.xl.input = () => app().querySelector('.xl-edit');
  /* 셀을 고르고 글자를 입력한 뒤 Enter */
  T.xl.type = (addr, text) => {
    T.xl.down(addr);
    const inp = T.xl.input();
    inp.focus();
    inp.value = text;
    inp.dispatchEvent(new Event('input', { bubbles: true }));
    T.key('Enter', {}, inp);
  };
  T.xl.fillHandle = (toAddr) => {
    const fh = app().querySelector('.xl-fill');
    const r = fh.getBoundingClientRect();
    fh.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX: r.left + 3, clientY: r.top + 3, button: 0 }));
    const c = T.xl.cellCenter(toAddr);
    window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: c.x, clientY: c.y }));
    window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: c.x, clientY: c.y }));
  };
  T.xl.text = addr => { T.xl.down(addr); return app().querySelector('.finput').value; };
  T.dlg = () => [...document.querySelectorAll('.dlg')].pop();
  T.dlgInputNear = (labelText) => {
    const d = T.dlg();
    const tw = document.createTreeWalker(d, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = tw.nextNode())) {
      if (!n.textContent.trim().startsWith(labelText)) continue;
      // 글자 바로 뒤의 입력 칸(같은 부모 안 형제 → 부모의 형제 순)
      for (let el = n.nextSibling; el; el = el.nextSibling) { if (el.nodeType !== 1) continue; if (el.matches('input, select, textarea')) return el; const q = el.querySelector('input, select, textarea'); if (q) return q; }
      let p = n.parentElement;
      const own = p.querySelector('input, select, textarea'); if (own && p.matches('label')) return own;
      for (let el = p.nextElementSibling; el; el = el.nextElementSibling) { if (el.matches('input, select, textarea')) return el; const q = el.querySelector('input, select, textarea'); if (q) return q; }
    }
    throw new Error('라벨 없음: ' + labelText);
  };

  /* 창에 초점이 없으면 focus() 가 focus 이벤트를 보내지 않으므로 직접 보낸다 */
  T.focus = el => { el.focus(); if (!document.hasFocus() || document.activeElement !== el) { el.dispatchEvent(new FocusEvent('focus')); el.dispatchEvent(new FocusEvent('focusin', { bubbles: true })); } };
  T.set = (el, value) => { if (el.type === 'checkbox' || el.type === 'radio') { el.checked = !!value; el.dispatchEvent(new Event('change', { bubbles: true })); return; } T.focus(el); el.value = value; el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); };
  T.ok = () => T.click('확인', T.dlg());
  // ---------------- Access
  T.ac = {};
  /* 테이블 디자인 i번째 행(0부터)에 필드 이름·데이터 형식 입력 */
  T.ac.field = (i, name, type) => {
    const tr = app().querySelectorAll('.ac-tdesign .top tr')[i + 1];
    const [nameIn, typeSel] = [tr.querySelector('input'), tr.querySelector('select')];
    T.set(nameIn, name);
    if (type) { T.focus(typeSel); typeSel.value = type; typeSel.dispatchEvent(new Event('change', { bubbles: true })); }
  };
  T.ac.selectFieldRow = i => T.clickEl(app().querySelectorAll('.ac-tdesign .top tr')[i + 1].querySelector('td.sel'));
  /* 테이블 디자인 필드 속성 */
  T.ac.prop = (label, value) => {
    const cells = [...app().querySelectorAll('.ac-tdesign .props .grid > div')];
    const k = cells.findIndex(c => c.textContent.trim() === label);
    if (k < 0) throw new Error('속성 없음: ' + label);
    const el = cells[k + 1].querySelector('input, select');
    T.focus(el); el.value = value; el.dispatchEvent(new Event('change', { bubbles: true }));
  };
  /* 데이터시트에 여러 행 입력(행마다 마지막 칸에서 Enter) */
  T.ac.rows = async (rows) => {
    for (const vals of rows) {
      const box = app().querySelector('table.ds');
      const trs = [...box.querySelectorAll('tr')];
      const tr = trs[trs.length - 1 - (box.querySelector('tr.tot') ? 1 : 0)];
      const ins = [...tr.querySelectorAll('input')];
      for (let k = 0; k < vals.length; k++) {
        const inp = [...app().querySelector('table.ds').querySelectorAll('tr')][trs.indexOf(tr)].querySelectorAll('input')[k];
        T.focus(inp); inp.value = vals[k] === null ? '' : String(vals[k]); inp.dispatchEvent(new Event('input', { bubbles: true }));
        if (k < vals.length - 1) T.key('Tab', {}, inp);
        else T.key('Enter', {}, inp);
        await T.wait(10);
      }
      void ins;
      await T.wait(20);
    }
  };
  /* 속성 시트 한 칸 입력 */
  T.ac.ps = (label, value) => {
    const cell = [...app().querySelectorAll('.propsheet .pg > div')].find(e => e.title === label);
    if (!cell) throw new Error('속성 시트 항목 없음: ' + label);
    const el = cell.nextElementSibling.querySelector('input, select');
    T.focus(el); el.value = value;
    if (el.tagName === 'SELECT') el.dispatchEvent(new Event('change', { bubbles: true }));
    else T.key('Enter', {}, el);
  };
  /* 폼·보고서 디자인 구역 안에 도구로 사각형을 끌어 그린다(구역 왼쪽 위 기준 화소) */
  T.ac.draw = (tool, x0, y0, x1, y1, sec) => {
    T.rb(tool);
    const body = [...app().querySelectorAll('.fd-sec-b')].find(b => !sec || b.dataset.sec === sec);
    const r = body.getBoundingClientRect();
    body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, clientX: r.left + x0, clientY: r.top + y0 }));
    window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: r.left + x1, clientY: r.top + y1 }));
    window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: r.left + x1, clientY: r.top + y1 }));
  };
  /* 지금 열린 이름 묻기 대화상자에 이름을 넣고 확인 */
  T.nameOk = (name) => { const d = T.dlg(); T.set(d.querySelector('input[type=text], input:not([type])'), name); T.click('확인', d); };
  return 'ready';
})();

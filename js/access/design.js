/* Access: 폼·보고서 디자인 보기(구역·컨트롤·속성 시트·그룹 창·필드 목록), 폼 보기, 보고서 보기, 인쇄 미리 보기, 보고서 마법사, 페이지 설정 */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const { h, icon, menu, dialog, msg, tabs, toast } = OA.ui;
  const U = OA.util, A = OA.acc, R = OA.report;
  const PX = 37.8;                                    // 1cm = 37.8px(96dpi)
  const BORDER = { '투명': 'none', '실선': 'solid', '파선': 'dashed', '짧은 파선': 'dashed', '점선': 'dotted', '드문 점선': 'dotted', '파선 점선': 'dashed', '파선 점선 점선': 'dashed', '이중 실선': 'double' };
  const BW = { '가는 선': 1, '1pt': 1.33, '2pt': 2.67, '3pt': 4, '4pt': 5.33, '5pt': 6.67, '6pt': 8 };
  const ALIGN = { '일반': '', '왼쪽': 'left', '가운데': 'center', '오른쪽': 'right', '균등 분할': 'justify' };
  const WEIGHT = { '얇게': 200, '아주 가늘게': 300, '가늘게': 300, '표준': 400, '보통': 500, '중간 굵게': 600, '굵게': 700, '아주 굵게': 800, '두껍게': 900 };
  const fmtCm = v => (Math.round((+v || 0) * 1000) / 1000) + 'cm';
  const cmOf = s => { const n = parseFloat(String(s).replace(/cm$/i, '').trim()); return isFinite(n) ? Math.round(n * 1000) / 1000 : null; };

  function ctlCss(c, s, fontRatio) {
    const css = { left: c.left * s + 'px', top: c.top * s + 'px', width: Math.max(1, c.width * s) + 'px', height: Math.max(1, c.height * s) + 'px', fontSize: (c.fontSize || 11) * 96 / 72 * fontRatio + 'px',
      fontWeight: WEIGHT[c.fontWeight] || 400, fontStyle: c.italic ? 'italic' : 'normal', textDecoration: c.underline ? 'underline' : 'none', color: c.foreColor || '#000' };
    if (c.type !== 'line') {
      const bs = BORDER[c.borderStyle] || 'none';
      css.border = bs === 'none' ? '1px solid transparent' : `${Math.max(1, (BW[c.borderWidth] || 1) * fontRatio)}px ${bs} ${c.borderColor || '#000'}`;
      css.background = c.backStyle === '보통' ? (c.backColor || '#fff') : 'transparent';
    }
    const al = ALIGN[c.textAlign];
    if (al) css.textAlign = al;
    return css;
  }

  // ===================================================================== 디자인 보기
  OA.acDesign = {};
  OA.acDesign.view = function (api, d, side, propOpen) {
    const db = api.db();
    const o = d.obj;
    const isReport = o.kind === 'report';
    d.sel = d.sel || { kind: isReport ? 'report' : 'form' };
    const canvas = h('div', { class: 'fd-canvas' });
    const wrap = h('div', { class: 'fd' }, h('div', { style: { flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 } }, canvas));
    const W = o.props.width * PX;
    const ruler = h('div', { class: 'fd-ruler', style: { width: W + 40 + 'px' } });
    for (let k = 0; k <= Math.ceil(o.props.width) + 2; k++) ruler.appendChild(h('span', { style: { position: 'absolute', left: k * PX + 'px', top: '3px' } }, String(k)));
    ruler.style.position = 'sticky';
    const selBtn = h('div', { title: isReport ? '보고서 선택기' : '폼 선택기', style: { position: 'absolute', left: 0, top: 0, width: '16px', height: '16px', background: d.sel.kind === 'form' || d.sel.kind === 'report' ? '#222' : '#ccc', cursor: 'pointer' }, onclick: () => { d.sel = { kind: isReport ? 'report' : 'form' }; api.renderNow(); } });
    ruler.appendChild(selBtn);
    canvas.appendChild(ruler);
    const touch = () => { d.dirty = true; };
    const secIds = isReport ? A.SEC_ORDER.filter(id => o.sections[id]) : ['detail'];
    const selected = () => d.sel.kind === 'controls' ? o.controls.filter(c => d.sel.ids.includes(c.id)) : [];
    for (const id of secIds) {
      const sec = o.sections[id];
      const title = isReport ? A.sectionTitle(o, id) : '본문';
      const hdr = h('div', { class: 'fd-sec-h' + (d.sel.kind === 'section' && d.sel.id === id ? ' on' : ''), style: { width: W + 40 + 'px' } }, '▾ ' + title);
      hdr.addEventListener('mousedown', () => { d.sel = { kind: 'section', id }; api.renderNow(); });
      const body = h('div', { class: 'fd-sec-b', style: { width: W + 'px', height: Math.max(0, sec.height * PX) + 'px', backgroundColor: sec.backColor || '#fff' } });
      body.dataset.sec = id;
      for (const c of o.controls.filter(x => x.section === id)) body.appendChild(ctlEl(c));
      const grip = h('div', { class: 'fd-sec-grip', style: { width: W + 'px' }, title: '끌어서 구역 높이 조정' });
      grip.addEventListener('mousedown', e => {
        e.preventDefault();
        const sy = e.clientY, h0 = sec.height;
        api.snapshot(d);
        OA.ui.drag(e, ev => { const nh = Math.max(0, Math.round((h0 + (ev.clientY - sy) / PX) * 100) / 100); const minH = Math.max(0, ...o.controls.filter(x => x.section === id).map(x => x.top + x.height)); sec.height = Math.max(nh, Math.round(minH * 100) / 100); body.style.height = sec.height * PX + 'px'; }, () => { touch(); api.renderNow(); });
      });
      body.addEventListener('mousedown', e => sectionDown(e, id, body));
      canvas.append(hdr, body, grip);
    }
    function ctlEl(c) {
      const isSel = d.sel.kind === 'controls' && d.sel.ids.includes(c.id);
      const el = h('div', { class: 'fd-ctl ' + ({ label: 'lb', textbox: 'tb', listbox: 'lst', line: 'line', rect: 'rect' }[c.type]) + (isSel ? ' sel' : '') });
      Object.assign(el.style, ctlCss(c, PX, 1));
      if (c.type === 'label') el.textContent = c.caption || '';
      if (c.type === 'textbox') { el.textContent = c.controlSource || '언바운드'; if (!c.textAlign || c.textAlign === '일반') el.style.textAlign = 'left'; }
      if (c.type === 'listbox') { el.textContent = c.rowSource ? (c.rowSource.length > 40 ? c.rowSource.slice(0, 40) + '…' : c.rowSource) : '언바운드'; el.style.color = '#555'; }
      if (c.type === 'line') { el.style.border = 'none'; el.style.height = Math.max(1, c.height * PX) + 'px'; el.innerHTML = `<svg width="${Math.max(2, c.width * PX)}" height="${Math.max(2, c.height * PX)}" style="overflow:visible;position:absolute;left:0;top:0"><line x1="0" y1="0" x2="${c.width * PX}" y2="${c.height * PX}" stroke="${c.borderColor || '#000'}" stroke-width="${BW[c.borderWidth] || 1}" stroke-dasharray="${BORDER[c.borderStyle] === 'dashed' ? '6 3' : BORDER[c.borderStyle] === 'dotted' ? '2 2' : ''}"/></svg>`; }
      if (c.type === 'rect') el.textContent = '';
      el.addEventListener('mousedown', e => ctlDown(e, c));
      el.addEventListener('dblclick', e => { e.stopPropagation(); if (c.type === 'label') editLabel(c, el); else { api.showProps(true); } });
      el.addEventListener('contextmenu', e => { e.preventDefault(); e.stopPropagation(); if (!isSel) { d.sel = { kind: 'controls', ids: [c.id] }; api.renderNow(); } menu(el, [{ label: '속성(P)', icon: 'propsheet', onClick: () => api.showProps(true) }, { label: '삭제(D)', icon: 'deleteCells', onClick: () => deleteSel() }], { at: { x: e.clientX, y: e.clientY } }); });
      if (isSel && d.sel.ids.length === 1) {
        for (const [px, py] of [[0, 0], [0.5, 0], [1, 0], [1, 0.5], [1, 1], [0.5, 1], [0, 1], [0, 0.5]]) {
          const hd = h('div', { class: 'hd', style: { left: `calc(${px * 100}% - 3px)`, top: `calc(${py * 100}% - 3px)`, cursor: 'nwse-resize' } });
          hd.addEventListener('mousedown', e => { e.stopPropagation(); e.preventDefault(); resizeDrag(e, c, px, py); });
          el.appendChild(hd);
        }
      }
      if (d.editing === c.id) setTimeout(() => editLabel(c, el), 0);
      return el;
    }
    function ctlDown(e, c) {
      if (e.button !== 0 || d.tool) return;
      e.stopPropagation(); e.preventDefault();
      commitEdit();
      if (e.shiftKey || e.ctrlKey) { if (d.sel.kind !== 'controls') d.sel = { kind: 'controls', ids: [] }; const i = d.sel.ids.indexOf(c.id); if (i >= 0) d.sel.ids.splice(i, 1); else d.sel.ids.push(c.id); api.renderNow(); return; }
      if (!(d.sel.kind === 'controls' && d.sel.ids.includes(c.id))) d.sel = { kind: 'controls', ids: [c.id] };
      const moving = selected();
      const start = moving.map(x => ({ x, l: x.left, t: x.top }));
      const sx = e.clientX, sy = e.clientY;
      let moved = false, snapped = false;
      OA.ui.drag(e, ev => {
        const dx = (ev.clientX - sx) / PX, dy = (ev.clientY - sy) / PX;
        if (!moved && Math.abs(ev.clientX - sx) + Math.abs(ev.clientY - sy) < 3) return;
        if (!snapped) { api.snapshot(d); snapped = true; }
        moved = true;
        for (const s of start) { s.x.left = Math.max(0, Math.round((s.l + dx) * 100) / 100); s.x.top = Math.max(0, Math.round((s.t + dy) * 100) / 100); }
        api.renderNow();
      }, ev => {
        if (moved) {
          // 다른 구역 위에 놓으면 구역 이동
          const tg = document.elementFromPoint(ev.clientX, ev.clientY);
          const body = tg && tg.closest('.fd-sec-b');
          if (body && body.dataset.sec && moving.length === 1 && body.dataset.sec !== moving[0].section) {
            const r = body.getBoundingClientRect();
            moving[0].section = body.dataset.sec;
            moving[0].top = Math.max(0, Math.round((ev.clientY - r.top) / PX * 100) / 100 - 0.1);
            A.growSection(o, moving[0]);
          }
          for (const s of start) A.growSection(o, s.x);
          touch();
        }
        api.renderNow();
      });
    }
    function resizeDrag(e, c, px, py) {
      api.snapshot(d);
      const sx = e.clientX, sy = e.clientY, o0 = { l: c.left, t: c.top, w: c.width, h: c.height };
      OA.ui.drag(e, ev => {
        const dx = (ev.clientX - sx) / PX, dy = (ev.clientY - sy) / PX;
        const r2 = v => Math.round(v * 100) / 100;
        if (px === 0) { c.left = r2(o0.l + dx); c.width = r2(Math.max(0.1, o0.w - dx)); }
        if (px === 1) c.width = r2(Math.max(0.1, o0.w + dx));
        if (py === 0) { c.top = r2(o0.t + dy); c.height = r2(Math.max(c.type === 'line' ? 0 : 0.1, o0.h - dy)); }
        if (py === 1) c.height = r2(Math.max(c.type === 'line' ? 0 : 0.1, o0.h + dy));
        if (c.type === 'line' && ev.shiftKey) c.height = 0;
        api.renderNow();
      }, () => { A.growSection(o, c); touch(); api.renderNow(); });
    }
    function sectionDown(e, id, body) {
      if (e.button !== 0) return;
      commitEdit();
      const r = body.getBoundingClientRect();
      const x0 = (e.clientX - r.left) / PX, y0 = (e.clientY - r.top) / PX;
      if (!d.tool) { d.sel = { kind: 'section', id }; api.renderNow(); return; }
      e.preventDefault();
      const ghost = h('div', { style: { position: 'absolute', border: '1px dashed #333', left: x0 * PX + 'px', top: y0 * PX + 'px', width: 0, height: 0, pointerEvents: 'none' } });
      body.appendChild(ghost);
      let x1 = x0, y1 = y0;
      OA.ui.drag(e, ev => {
        x1 = (ev.clientX - r.left) / PX; y1 = (ev.clientY - r.top) / PX;
        if (d.tool === 'line' && ev.shiftKey) y1 = y0;
        Object.assign(ghost.style, { left: Math.min(x0, x1) * PX + 'px', top: Math.min(y0, y1) * PX + 'px', width: Math.abs(x1 - x0) * PX + 'px', height: Math.abs(y1 - y0) * PX + 'px' });
      }, ev => {
        ghost.remove();
        const r2 = v => Math.round(v * 100) / 100;
        let w = Math.abs(x1 - x0), hh = Math.abs(y1 - y0);
        const kind = d.tool;
        const def = { label: [2.5, 0.6], textbox: [3, 0.6], listbox: [4.5, 3], line: [3, 0], rect: [3, 2] }[kind];
        if (w < 0.15 && hh < 0.15) { w = def[0]; hh = def[1]; }
        if (kind === 'line' && ev.shiftKey) hh = 0;
        api.snapshot(d);
        const props = { section: id, left: r2(Math.min(x0, x1)), top: r2(Math.min(y0, y1)), width: r2(Math.max(0.1, w)), height: r2(kind === 'line' ? (ev.shiftKey ? 0 : hh) : Math.max(0.3, hh)) };
        if (kind === 'label') props.withLabel = false;
        if (kind === 'line' || kind === 'rect') props.withLabel = false;
        const c = A.addControl(o, kind, props);
        if (kind === 'textbox') { const lab = o.controls.find(x => x.attachedTo === c.id); if (lab) lab.caption = c.name + ':'; }
        d.tool = null;
        d.sel = { kind: 'controls', ids: [c.id] };
        touch();
        if (kind === 'label') d.editing = c.id;
        if (kind === 'listbox' && api.app.wizardOn) listWizard(api, d, c);
        api.renderNow();
        api.rib().render();
      });
    }
    function editLabel(c, el) {
      if (d.editingOpen) return;
      d.editingOpen = true;
      d.editing = c.id;
      const ta = h('textarea', { spellcheck: false }, c.caption || '');
      el.textContent = '';
      el.appendChild(ta);
      ta.focus();
      ta.setSelectionRange(ta.value.length, ta.value.length);
      const finish = (ok) => {
        if (!d.editingOpen) return;
        d.editingOpen = false; d.editing = null;
        if (ok) { api.snapshot(d); c.caption = ta.value; }
        if (!c.caption) { o.controls = o.controls.filter(x => x !== c); d.sel = { kind: 'section', id: c.section }; }
        touch();
        api.renderNow();
      };
      ta.addEventListener('keydown', e => {
        e.stopPropagation();
        if (e.key === 'Enter' && e.ctrlKey) { e.preventDefault(); const p = ta.selectionStart; ta.value = ta.value.slice(0, p) + '\n' + ta.value.slice(ta.selectionEnd); ta.setSelectionRange(p + 1, p + 1); return; }
        if (e.key === 'Enter') { e.preventDefault(); finish(true); }
        if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      });
      ta.addEventListener('mousedown', e => e.stopPropagation());
      ta.addEventListener('blur', () => finish(true));
      d.commitLabel = () => finish(true);
    }
    function commitEdit() { if (d.editingOpen && d.commitLabel) d.commitLabel(); }
    function deleteSel() {
      const ids = d.sel.kind === 'controls' ? d.sel.ids.slice() : [];
      if (!ids.length) return;
      api.snapshot(d);
      o.controls = o.controls.filter(c => !ids.includes(c.id) && !ids.includes(c.attachedTo));
      d.sel = { kind: isReport ? 'report' : 'form' };
      touch(); api.renderNow();
    }
    d.keys = e => {
      if (e.target && /INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
      if (e.key === 'Delete') { e.preventDefault(); deleteSel(); return; }
      const arr = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
      if (arr && d.sel.kind === 'controls') {
        e.preventDefault();
        api.snapshot(d);
        const step = 0.05;
        for (const c of selected()) {
          if (e.shiftKey) { c.width = Math.max(0.05, Math.round((c.width + arr[0] * step) * 100) / 100); c.height = Math.max(c.type === 'line' ? 0 : 0.05, Math.round((c.height + arr[1] * step) * 100) / 100); }
          else { c.left = Math.max(0, Math.round((c.left + arr[0] * step) * 100) / 100); c.top = Math.max(0, Math.round((c.top + arr[1] * step) * 100) / 100); }
          A.growSection(o, c);
        }
        touch(); api.renderNow();
      }
      if (e.key === 'Escape') { d.tool = null; api.renderNow(); }
    };
    d.design = {
      setTool(k) { commitEdit(); d.tool = k; api.renderNow(); },
      toggleFieldList() { d.fieldList = !d.fieldList; api.renderNow(); },
      toggleGrouping() { d.grouping = !d.grouping; api.renderNow(); },
      groupingOpen: () => !!d.grouping,
      commitEdit,
      reset() { d.sel = { kind: isReport ? 'report' : 'form' }; },
    };
    canvas.style.cursor = d.tool ? 'crosshair' : 'default';
    if (isReport && d.grouping) wrap.firstChild.appendChild(groupPane(api, d));
    if (d.fieldList) side.appendChild(fieldListPane(api, d));
    if (propOpen) side.appendChild(propSheet(api, d));
    return wrap;
  };

  // ------------------------------------------------------------------ 목록 상자 마법사
  function listWizard(api, d, c) {
    const db = api.db();
    let step = 1, src = null;
    const body = h('div');
    const draw = () => {
      body.innerHTML = '';
      if (step === 1) body.append(h('p', {}, '목록 상자에 표시할 값을 어떻게 가져오시겠습니까?'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'lw', checked: true }), '목록 상자에서 테이블이나 쿼리의 값을 조회합니다.'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'lw', disabled: true }), '원하는 값을 직접 입력합니다.'));
      else body.append(h('p', {}, '목록 상자에 값을 제공할 테이블 또는 쿼리를 선택하십시오.'), h('div', { class: 'lst', style: { height: '160px' } }, Object.keys(db.tables).map(n => '테이블: ' + n).concat(Object.keys(db.queries).map(n => '쿼리: ' + n)).map(n => { const x = h('div', {}, n); x.onclick = () => { src = n.replace(/^(테이블|쿼리): /, ''); [...x.parentNode.children].forEach(y => y.classList.toggle('on', y === x)); }; return x; })));
    };
    draw();
    dialog({ title: '목록 상자 마법사', width: 480, body, buttons: [
      { label: '취소', cancel: true }, { label: '< 뒤로(B)', onClick: () => { step = 1; draw(); return false; } }, { label: '다음(N) >', onClick: () => { step = 2; draw(); return false; } },
      { label: '마침(F)', primary: true, onClick: () => {
        if (!src) return true;
        const cols = A.objFields(db, src);
        api.snapshot(d);
        Object.assign(c, { rowSourceType: '테이블/쿼리', rowSource: `SELECT ${cols.map(x => `[${src}].[${x}]`).join(', ')} FROM [${src}];`, columnCount: cols.length, columnHeads: false });
        d.dirty = true; api.renderNow();
      } }] });
  }

  // ------------------------------------------------------------------ 필드 목록
  function sourceFields(db, src) {
    if (!src) return [];
    try { const r = R.loadSource(db, src); return r.columns || []; } catch (e) { return []; }
  }
  function fieldListPane(api, d) {
    const db = api.db();
    const o = d.obj;
    const fields = sourceFields(db, o.props.recordSource);
    const list = h('div', { class: 'lst', style: { height: '300px' } }, fields.length ? fields.map(f => { const x = h('div', {}, f); x.addEventListener('mousedown', e => dragField(e, f)); return x; }) : h('div', { class: 'note' }, '레코드 원본을 먼저 지정합니다(속성 시트 데이터 탭).'));
    function dragField(e, f) {
      e.preventDefault();
      const ghost = h('div', { class: 'dragghost' }, f);
      document.body.appendChild(ghost);
      OA.ui.drag(e, ev => { ghost.style.left = ev.clientX + 8 + 'px'; ghost.style.top = ev.clientY + 8 + 'px'; }, ev => {
        ghost.remove();
        const tg = document.elementFromPoint(ev.clientX, ev.clientY);
        const body = tg && tg.closest('.fd-sec-b');
        if (!body) return;
        const r = body.getBoundingClientRect();
        api.snapshot(d);
        const c = A.addControl(o, 'textbox', { section: body.dataset.sec, left: Math.round((ev.clientX - r.left) / PX * 100) / 100 + 2.5, top: Math.round((ev.clientY - r.top) / PX * 100) / 100, width: 3, height: 0.6, controlSource: f });
        const lab = o.controls.find(x => x.attachedTo === c.id); if (lab) lab.caption = f;
        d.dirty = true; api.renderNow();
      });
    }
    return h('div', { class: 'pane' }, h('div', { class: 'pane-h' }, h('span', { class: 't' }, '필드 목록'), h('button', { onclick: () => { d.fieldList = false; api.renderNow(); } }, '✕')), h('div', { class: 'pane-b' }, h('div', { class: 'note' }, '현재 레코드 원본에서 사용할 수 있는 필드:'), list));
  }

  // ------------------------------------------------------------------ 그룹, 정렬 및 요약
  function groupPane(api, d) {
    const db = api.db();
    const o = d.obj;
    const fields = sourceFields(db, o.props.recordSource);
    const el = h('div', { class: 'grpane' }, h('div', { class: 'gt' }, '그룹, 정렬 및 요약'));
    const apply = () => { A.setGroupSections(o); d.dirty = true; api.renderNow(); };
    o.grouping.forEach((g, i) => {
      const fsel = h('select', { onchange: () => { api.snapshot(d); g.field = fsel.value; apply(); } }, fields.map(f => h('option', { selected: f === g.field }, f)));
      const osel = h('select', { onchange: () => { api.snapshot(d); g.order = osel.value; apply(); } }, h('option', { value: 'asc', selected: g.order !== 'desc' }, '오름차순'), h('option', { value: 'desc', selected: g.order === 'desc' }, '내림차순'));
      const row = h('div', { class: 'lv' + (g.kind === 'sort' ? ' sort' : '') }, h('span', {}, g.kind === 'group' ? '그룹화 기준' : '정렬 기준'), fsel, osel);
      if (g.kind === 'group') {
        const more = h('button', { class: 'btn', style: { padding: '0 6px', minHeight: '20px' } }, g._open ? '자세히 ◀' : '자세히 ▶');
        more.onclick = () => { g._open = !g._open; api.renderNow(); };
        row.appendChild(more);
        if (g._open) {
          const hs = h('select', { onchange: () => { api.snapshot(d); g.header = hs.value === 'on'; apply(); } }, h('option', { value: 'on', selected: g.header }, '머리글 구역 표시'), h('option', { value: 'off', selected: !g.header }, '머리글 구역 없음'));
          const fs = h('select', { onchange: () => { api.snapshot(d); g.footer = fs.value === 'on'; apply(); } }, h('option', { value: 'on', selected: g.footer }, '바닥글 구역 표시'), h('option', { value: 'off', selected: !g.footer }, '바닥글 구역 없음'));
          const ks = h('select', { onchange: () => { api.snapshot(d); g.keep = ks.value; apply(); } }, ['표시 안 함', '같은 페이지에 그룹 전체 표시', '같은 페이지에 머리글과 첫 번째 레코드 표시'].map(x => h('option', { selected: g.keep === x }, x)));
          row.append(h('span', { class: 'note' }, '요약 없음'), hs, fs, ks);
        }
      }
      row.append(h('button', { class: 'btn', title: '위로', style: { padding: '0 5px', minHeight: '20px' }, onclick: () => { if (i > 0) { api.snapshot(d); [o.grouping[i - 1], o.grouping[i]] = [o.grouping[i], o.grouping[i - 1]]; apply(); } } }, '▲'),
        h('button', { class: 'btn', title: '아래로', style: { padding: '0 5px', minHeight: '20px' }, onclick: () => { if (i < o.grouping.length - 1) { api.snapshot(d); [o.grouping[i + 1], o.grouping[i]] = [o.grouping[i], o.grouping[i + 1]]; apply(); } } }, '▼'),
        h('button', { class: 'btn', title: '삭제', style: { padding: '0 5px', minHeight: '20px' }, onclick: () => { api.snapshot(d); o.grouping.splice(i, 1); apply(); } }, '✕'));
      el.appendChild(row);
    });
    const add = kind => {
      if (!fields.length) return msg({ title: 'Microsoft Access', text: '레코드 원본을 먼저 지정합니다.', icon: 'info' });
      menu(el, fields.map(f => ({ label: f, onClick: () => { api.snapshot(d); const g = { kind, field: f, order: 'asc' }; if (kind === 'group') Object.assign(g, { header: true, footer: false, keep: '표시 안 함' }); const lastGroup = o.grouping.map(x => x.kind).lastIndexOf('group'); if (kind === 'group') o.grouping.splice(lastGroup + 1, 0, g); else o.grouping.push(g); apply(); } })));
    };
    el.appendChild(h('div', { class: 'add' }, h('button', { onclick: () => add('group') }, '☷ 그룹 추가'), h('button', { onclick: () => add('sort') }, '↓ 정렬 추가')));
    return el;
  }

  // ------------------------------------------------------------------ 속성 시트
  function propSheet(api, d) {
    const db = api.db();
    const o = d.obj;
    const isReport = o.kind === 'report';
    d.ptab = d.ptab || '모두';
    const sel = d.sel;
    const objs = sel.kind === 'controls' ? o.controls.filter(c => sel.ids.includes(c.id)) : [];
    const one = objs.length === 1 ? objs[0] : null;
    const typeName = sel.kind === 'controls' ? (objs.length > 1 ? '여러 항목 선택' : { label: '레이블', textbox: '텍스트 상자', listbox: '목록 상자', line: '선', rect: '사각형' }[one.type]) : sel.kind === 'section' ? '구역' : (isReport ? '보고서' : '폼');
    const P = [];   // [tab, label, get, set, kind, opts]
    const yes = v => v ? '예' : '아니요';
    const numProp = (tab, label, key, target) => P.push([tab, label, () => fmtCm(target[key]), v => { const n = cmOf(v); if (n === null || n < 0) throw new Error('입력한 값이 이 필드에 대해 잘못되었습니다.'); target[key] = n; }]);
    const enumProp = (tab, label, key, target, list, map) => P.push([tab, label, () => map ? map.get(target[key]) : target[key], v => { target[key] = map ? map.set(v) : v; }, 'select', list]);
    const textProp = (tab, label, key, target, opts) => P.push([tab, label, () => target[key] === undefined || target[key] === null ? '' : String(target[key]), v => { target[key] = opts && opts.num ? (+v || 0) : v; }, opts && opts.list ? 'combo' : 'text', opts && opts.list]);
    const ynMap = (k) => ({ get: v => yes(!!v), set: v => v === '예' });
    if (sel.kind === 'form' || sel.kind === 'report') {
      const p = o.props;
      textProp('형식', '캡션', 'caption', p);
      if (!isReport) enumProp('형식', '기본 보기', 'defaultView', p, ['단일 폼', '연속 폼', '데이터시트', '분할 표시 폼']);
      numProp('형식', '너비', 'width', p);
      if (!isReport) {
        enumProp('형식', '레코드 선택기', 'recordSelectors', p, ['예', '아니요'], ynMap());
        enumProp('형식', '탐색 단추', 'navButtons', p, ['예', '아니요'], ynMap());
        enumProp('형식', '구분선', 'dividingLines', p, ['예', '아니요'], ynMap());
        enumProp('형식', '스크롤 막대', 'scrollBars', p, ['모두', '가로만', '세로만', '없음']);
      } else {
        enumProp('형식', '페이지 머리글', 'pageHeaderWhere', p, ['모든 페이지', '보고서 머리글 제외', '보고서 바닥글 제외', '보고서 머리글/바닥글 제외']);
      }
      const srcList = Object.keys(db.tables).concat(Object.keys(db.queries));
      textProp('데이터', '레코드 원본', 'recordSource', p, { list: srcList });
      P.push(['기타', '이름', () => d.name || '', () => { throw new Error('개체 이름은 탐색 창에서 바꿉니다.'); }]);
    } else if (sel.kind === 'section') {
      const s = o.sections[sel.id];
      numProp('형식', '높이', 'height', s);
      textProp('형식', '배경색', 'backColor', s);
      if (isReport) textProp('형식', '다른 배경색', 'altBackColor', s);
      enumProp('형식', '표시', 'visible', s, ['예', '아니요'], { get: v => yes(v !== false), set: v => v === '예' });
      if (isReport) { enumProp('형식', '페이지 바꿈', 'pageBreak', s, ['없음', '구역 전', '구역 후', '구역 전/후']); if (/^groupHeader/.test(sel.id)) enumProp('형식', '구역 반복', 'repeat', s, ['예', '아니요'], ynMap()); }
      P.push(['기타', '이름', () => isReport ? A.sectionTitle(o, sel.id) : '본문', () => {}]);
    } else if (objs.length) {
      const t = one || objs[0];
      const all = (key, val) => objs.forEach(c => { c[key] = val; });
      const numAll = (tab, label, key) => P.push([tab, label, () => objs.every(c => c[key] === t[key]) ? fmtCm(t[key]) : '', v => { const n = cmOf(v); if (n === null || n < 0) throw new Error('입력한 값이 이 필드에 대해 잘못되었습니다.'); all(key, n); }]);
      const enumAll = (tab, label, key, list, map) => P.push([tab, label, () => objs.every(c => c[key] === t[key]) ? (map ? map.get(t[key]) : t[key]) : '', v => all(key, map ? map.set(v) : v), 'select', list]);
      const txtAll = (tab, label, key, opts) => P.push([tab, label, () => objs.every(c => c[key] === t[key]) ? (t[key] === undefined || t[key] === null ? '' : String(t[key])) : '', v => all(key, opts && opts.num ? (+v || 0) : v), opts && opts.list ? 'combo' : 'text', opts && opts.list]);
      if (one && one.type === 'label') txtAll('형식', '캡션', 'caption');
      if (one && one.type === 'listbox') { txtAll('형식', '열 개수', 'columnCount', { num: true }); txtAll('형식', '열 너비', 'columnWidths'); enumAll('형식', '열 이름', 'columnHeads', ['예', '아니요'], ynMap()); }
      if (objs.every(c => c.type === 'textbox')) { enumAll('형식', '형식', 'format', ['', '일반 숫자', '통화', '유로', '고정', '표준', '백분율', '지수', '0', '#,##0']); enumAll('형식', '소수 자릿수', 'decimals', ['자동', '0', '1', '2', '3', '4'], { get: v => String(v === undefined ? '자동' : v), set: v => v === '자동' ? '자동' : +v }); }
      enumAll('형식', '표시', 'visible', ['예', '아니요'], { get: v => yes(v !== false), set: v => v === '예' });
      numAll('형식', '너비', 'width'); numAll('형식', '높이', 'height'); numAll('형식', '위쪽', 'top'); numAll('형식', '왼쪽', 'left');
      if (objs.every(c => c.type !== 'line')) { enumAll('형식', '배경 스타일', 'backStyle', ['보통', '투명']); txtAll('형식', '배경색', 'backColor'); }
      enumAll('형식', '테두리 스타일', 'borderStyle', Object.keys(BORDER));
      enumAll('형식', '테두리 두께', 'borderWidth', Object.keys(BW));
      txtAll('형식', '테두리 색', 'borderColor', { list: ['#000000', '텍스트 1, 50% 더 밝게', '배경 1, 25% 더 어둡게'] });
      if (objs.every(c => c.type !== 'line' && c.type !== 'rect')) {
        enumAll('형식', '글꼴 이름', 'fontName', ['맑은 고딕', '굴림', '돋움', '바탕']);
        txtAll('형식', '글꼴 크기', 'fontSize', { num: true, list: ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24'] });
        enumAll('형식', '텍스트 맞춤', 'textAlign', ['일반', '왼쪽', '가운데', '오른쪽', '균등 분할']);
        enumAll('형식', '글꼴 두께', 'fontWeight', Object.keys(WEIGHT));
        enumAll('형식', '밑줄', 'underline', ['예', '아니요'], ynMap());
        enumAll('형식', '기울임꼴', 'italic', ['예', '아니요'], ynMap());
        txtAll('형식', '전경색', 'foreColor');
      }
      if (isReport && objs.every(c => c.type === 'textbox')) { enumAll('형식', '중복 내용 숨기기', 'hideDuplicates', ['예', '아니요'], ynMap()); enumAll('형식', '확장 가능', 'canGrow', ['예', '아니요'], ynMap()); }
      const srcFields = sourceFields(db, o.props.recordSource);
      if (one && one.type === 'textbox') { txtAll('데이터', '컨트롤 원본', 'controlSource', { list: srcFields }); if (isReport) enumAll('데이터', '누적 합계', 'runningSum', ['아니요', '그룹', '모두']); }
      if (one && one.type === 'listbox') {
        txtAll('데이터', '컨트롤 원본', 'controlSource', { list: srcFields });
        txtAll('데이터', '행 원본', 'rowSource', { list: Object.keys(db.queries).concat(Object.keys(db.tables)) });
        enumAll('데이터', '행 원본 유형', 'rowSourceType', ['테이블/쿼리', '값 목록', '필드 목록']);
        txtAll('데이터', '바운드 열', 'boundColumn', { num: true });
      }
      if (one) P.push(['기타', '이름', () => one.name, v => { const n = String(v).trim(); if (!n) throw new Error('이름을 입력하십시오.'); if (o.controls.some(c => c !== one && c.name === n)) throw new Error(`'${n}'은(는) 이미 사용하고 있는 이름입니다. 다른 이름을 입력하십시오.`); one.name = n; }]);
    }
    const tabsOrder = ['형식', '데이터', '이벤트', '기타', '모두'];
    const grid = h('div', { class: 'pg' });
    const rows = P.filter(p => d.ptab === '모두' || p[0] === d.ptab);
    for (const [, label, get, set, kind, opts] of rows) {
      let inp;
      const commit = v => {
        try { api.snapshot(d); set(v); if (sel.kind === 'controls') objs.forEach(c => A.growSection(o, c)); if (sel.kind === 'section') { const minH = Math.max(0, ...o.controls.filter(c => c.section === sel.id).map(c => c.top + c.height)); if (o.sections[sel.id].height < minH - 0.001) { msg({ title: 'Microsoft Access', text: '구역 높이가 컨트롤보다 작을 수 없습니다.', icon: 'warn' }); o.sections[sel.id].height = Math.round(minH * 100) / 100; } } d.dirty = true; api.renderNow(); }
        catch (err) { d.undo && d.undo.pop(); msg({ title: 'Microsoft Access', text: err.message, icon: 'warn' }); api.renderNow(); }
      };
      if (kind === 'select') { const cur = get(); inp = h('select', { onchange: () => commit(inp.value) }, (cur === '' ? [h('option', { value: '' }, '')] : []).concat(opts.map(x => h('option', { value: x, selected: String(x) === String(cur) }, x)))); }
      else {
        inp = h('input', { value: get(), list: kind === 'combo' ? 'ps_' + label : null });
        inp.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter' || e.key === 'Tab') { if (e.key === 'Enter') e.preventDefault(); if (inp.value !== get()) commit(inp.value); } });
        inp.addEventListener('change', () => { if (inp.value !== get()) commit(inp.value); });
      }
      grid.append(h('div', { title: label }, label), h('div', {}, inp, kind === 'combo' ? h('datalist', { id: 'ps_' + label }, (opts || []).map(x => h('option', { value: x }))) : null));
    }
    if (!rows.length) grid.append(h('div', { style: { gridColumn: '1 / 3', color: '#888' } }, '이 탭에 표시할 속성이 없습니다.'));
    const objSel = h('select', { onchange: () => {
      const v = objSel.value;
      if (v === '#obj') d.sel = { kind: isReport ? 'report' : 'form' };
      else if (v.startsWith('#sec:')) d.sel = { kind: 'section', id: v.slice(5) };
      else d.sel = { kind: 'controls', ids: [v] };
      api.renderNow();
    } }, h('option', { value: '#obj', selected: sel.kind === 'form' || sel.kind === 'report' }, isReport ? '보고서' : '폼'),
      Object.keys(o.sections).map(id => h('option', { value: '#sec:' + id, selected: sel.kind === 'section' && sel.id === id }, isReport ? A.sectionTitle(o, id) : '본문')),
      o.controls.map(c => h('option', { value: c.id, selected: one && one.id === c.id }, c.name)));
    return h('div', { class: 'propsheet' },
      h('div', { class: 'ph' }, h('span', { class: 't' }, '속성 시트'), h('button', { title: '닫기', onclick: () => api.showProps(false) }, '✕')),
      h('div', { class: 'sel' }, h('div', { class: 'note' }, '선택 유형: ' + typeName), objSel),
      h('div', { class: 'ptabs' }, tabsOrder.map(t => h('button', { class: d.ptab === t ? 'on' : '', onclick: () => { d.ptab = t; api.renderNow(); } }, t))),
      grid);
  }

  // ===================================================================== 폼 보기·보고서 보기
  function formContent(db, f, s, fontRatio) {
    const box = h('div', { style: { position: 'relative', width: f.props.width * s + 'px', height: f.sections.detail.height * s + 'px', background: f.sections.detail.backColor || '#fff' } });
    for (const c of f.controls.filter(x => x.visible !== false && x.section === 'detail')) {
      const el = h('div', { style: Object.assign({ position: 'absolute', overflow: 'hidden', whiteSpace: 'pre-wrap', lineHeight: '1.25', padding: '0 ' + 2 * fontRatio + 'px', fontFamily: 'Malgun Gothic, sans-serif' }, ctlCss(c, s, fontRatio)) });
      if (c.type === 'label') el.textContent = c.caption || '';
      else if (c.type === 'textbox') { const v = R.controlValue(c, { row: {}, scope: [] }, {}); el.textContent = v.text; }
      else if (c.type === 'line') { el.style.border = 'none'; el.innerHTML = `<svg width="${Math.max(2, c.width * s)}" height="${Math.max(2, c.height * s)}" style="overflow:visible"><line x1="0" y1="0" x2="${c.width * s}" y2="${c.height * s}" stroke="#000" stroke-width="${fontRatio}"/></svg>`; }
      else if (c.type === 'listbox') {
        const data = R.listboxData(db, c);
        const n = Math.max(1, +c.columnCount || 1);
        const widths = R.parseWidths(c.columnWidths, n, c.width);
        const tbl = h('table', { style: { borderCollapse: 'collapse', tableLayout: 'fixed', fontSize: 'inherit', whiteSpace: 'nowrap' } });
        const colgroup = h('colgroup', {}, widths.map(w => h('col', { style: { width: w * s + 'px' } })));
        tbl.appendChild(colgroup);
        if (c.columnHeads) tbl.appendChild(h('tr', {}, [...Array(n).keys()].map(k => h('td', { style: { padding: '0 ' + 2 * fontRatio + 'px', overflow: 'hidden', borderBottom: '1px solid #ccc' } }, data.columns[k] || ''))));
        data.rows.forEach(row => tbl.appendChild(h('tr', {}, [...Array(n).keys()].map(k => { const v = row[k]; const t = OA.acc.formatValue(v, data.info && data.info[k]); return h('td', { style: { padding: '0 ' + 2 * fontRatio + 'px', overflow: 'hidden', textAlign: typeof v === 'number' ? 'right' : 'left' } }, t); }))));
        if (data.error) el.textContent = data.error;
        else el.appendChild(tbl);
        const sumW = widths.reduce((a, b) => a + b, 0);
        if (sumW > c.width + 0.01) el.appendChild(h('div', { style: { position: 'absolute', left: 0, right: 0, bottom: 0, height: 10 * fontRatio + 'px', background: '#e6e6e6', borderTop: '1px solid #bbb' } }));
      }
      box.appendChild(el);
    }
    return box;
  }
  OA.acDesign.formView = function (api, d) {
    const db = api.db();
    const f = d.obj;
    const wrap = h('div', { style: { padding: '0', background: '#fff', minHeight: '100%' } });
    const inner = h('div', { style: { display: 'flex' } });
    if (f.props.recordSelectors) inner.appendChild(h('div', { style: { width: '18px', background: '#e8e8e8', borderRight: '1px solid #ccc' } }, '▶'));
    inner.appendChild(formContent(db, f, PX, 1));
    wrap.appendChild(inner);
    if (f.props.navButtons) wrap.appendChild(h('div', { class: 'recnav' }, '레코드: |◀ ◀ ', h('input', { value: '1/1', readOnly: true }), ' ▶ ▶| ▶*  필터 없음'));
    d.keys = null;
    return wrap;
  };
  function reportPageEl(db, rpt, out, pi, s, mmScale) {
    // s: px per cm, page 는 A4
    const box = out.box;
    const pageW = box.paperW, pageH = box.paperH;
    const pg = h('div', { class: 'a4', style: { width: pageW * s + 'px', height: pageH * s + 'px' } });
    pg.mm = { W: pageW * 10, H: pageH * 10, scale: s / 10 };
    const hp = out.hPages;
    const logical = Math.floor(pi / hp), hOff = (pi % hp) * box.availW;
    const page = out.pages[logical];
    if (!page) return pg;
    const content = h('div', { style: { position: 'absolute', left: rpt.page.left / 10 * s + 'px', top: rpt.page.top / 10 * s + 'px', width: box.availW * s + 'px', height: box.availH * s + 'px', overflow: 'hidden' } });
    const inner = h('div', { style: { position: 'absolute', left: -hOff * s + 'px', top: 0, width: rpt.props.width * s + 'px' } });
    const fr = s / PX;
    for (const it of page.items) {
      const sec = rpt.sections[it.inst.section];
      if (!sec || sec.height <= 0) continue;
      const bg = it.inst.alt ? (sec.altBackColor || sec.backColor) : sec.backColor;
      const secEl = h('div', { style: { position: 'absolute', left: 0, top: it.y * s + 'px', width: rpt.props.width * s + 'px', height: sec.height * s + 'px', background: rpt.page.dataOnly ? 'transparent' : (bg || '#fff') } });
      for (const x of it.controls) {
        const c = x.c;
        if (c.visible === false) continue;
        if (rpt.page.dataOnly && c.type === 'label') continue;
        const el = h('div', { style: Object.assign({ position: 'absolute', overflow: 'hidden', whiteSpace: 'nowrap', lineHeight: '1.25', padding: '0 ' + 2 * fr + 'px', fontFamily: 'Malgun Gothic, sans-serif' }, ctlCss(c, s, fr)) });
        if (c.type === 'line') { el.style.border = 'none'; el.innerHTML = `<svg width="${Math.max(2, c.width * s)}" height="${Math.max(2, c.height * s + 2)}" style="overflow:visible;position:absolute;top:0;left:0"><line x1="0" y1="0" x2="${c.width * s}" y2="${c.height * s}" stroke="${c.borderColor || '#000'}" stroke-width="${Math.max(0.8, (BW[c.borderWidth] || 1) * fr)}"/></svg>`; }
        else { el.textContent = x.text; if ((!c.textAlign || c.textAlign === '일반') && typeof x.value === 'number') el.style.textAlign = 'right'; }
        secEl.appendChild(el);
      }
      inner.appendChild(secEl);
    }
    content.appendChild(inner);
    pg.appendChild(content);
    return pg;
  }
  OA.acDesign.reportView = function (api, d) {
    const db = api.db();
    let out;
    try { out = R.renderReport(db, d.obj); } catch (e) { return h('div', { class: 'ac-empty' }, '보고서를 열 수 없습니다: ' + e.message); }
    if (out.error || out.needParam) return h('div', { class: 'ac-empty' }, out.error || ('매개 변수 값이 필요합니다: ' + out.needParam));
    const wrap = h('div', { style: { padding: '14px', background: '#ddd', minHeight: '100%', display: 'flex', flexDirection: 'column', gap: '12px', alignItems: 'center' } });
    for (let i = 0; i < out.pageCount; i++) wrap.appendChild(reportPageEl(db, d.obj, out, i, PX));
    d.keys = null;
    return wrap;
  };
  function formPageEl(db, f, s, pi) {
    const box = R.pageBox(f.page);
    const pg = h('div', { class: 'a4', style: { width: box.paperW * s + 'px', height: box.paperH * s + 'px' } });
    pg.mm = { W: box.paperW * 10, H: box.paperH * 10, scale: s / 10 };
    const r = R.renderForm(db, f);
    const col = pi % r.hPages, row = Math.floor(pi / r.hPages);
    const content = h('div', { style: { position: 'absolute', left: f.page.left / 10 * s + 'px', top: f.page.top / 10 * s + 'px', width: box.availW * s + 'px', height: box.availH * s + 'px', overflow: 'hidden' } });
    const fc = formContent(db, f, s, s / PX);
    fc.style.position = 'absolute'; fc.style.left = -col * box.availW * s + 'px'; fc.style.top = -row * box.availH * s + 'px';
    if (f.page.dataOnly) fc.querySelectorAll('div').forEach(x => { if (x.childElementCount === 0 && !x.querySelector('table')) x.style.border = 'none'; });
    content.appendChild(fc);
    pg.appendChild(content);
    return pg;
  }
  OA.acDesign.renderFormPage = (db, f, mmScale, pi) => formPageEl(db, f, mmScale * 10, pi || 0);
  OA.acDesign.renderReportPage = (db, rpt, mmScale, pi) => {
    let out;
    try { out = R.renderReport(db, rpt); } catch (e) { return null; }
    if (out.error || out.needParam) return null;
    return reportPageEl(db, rpt, out, pi || 0, mmScale * 10);
  };

  // ===================================================================== 인쇄 미리 보기
  OA.acDesign.preview = function (api, d) {
    const db = api.db();
    const src = d.src;
    const o = src.obj;
    let count = 1;
    if (o.kind === 'form') count = R.renderForm(db, o).pageCount;
    else { const out = R.renderReport(db, o); if (out.error || out.needParam) return h('div', { class: 'ac-empty' }, out.error || ('매개 변수 값이 필요합니다: ' + out.needParam)); count = out.pageCount; }
    d.page = Math.max(0, Math.min(d.page || 0, count - 1));
    const area = h('div', { style: { flex: 1, overflow: 'auto', background: '#9a9a9a', display: 'flex', justifyContent: 'center', alignItems: 'flex-start', gap: '16px', padding: '16px' } });
    const fit = d.zoom ? d.zoom * PX : Math.max(8, Math.min(26, (api.app.el.querySelector('.ac-doc').clientHeight - 80) / 29.7));
    const mk = i => o.kind === 'form' ? formPageEl(db, o, fit, i) : reportPageEl(db, o, R.renderReport(db, o), i, fit);
    area.appendChild(mk(d.page));
    if (d.two && d.page + 1 < count) area.appendChild(mk(d.page + 1));
    const pageIn = h('input', { value: String(d.page + 1), style: { width: '40px', textAlign: 'center' }, onkeydown: e => { e.stopPropagation(); if (e.key === 'Enter') { d.page = Math.max(0, Math.min(count - 1, (parseInt(pageIn.value, 10) || 1) - 1)); api.renderNow(); } } });
    const btn = (t, dis, fn, tip) => h('button', { disabled: dis, title: tip, onclick: fn, style: dis ? { color: '#bbb' } : {} }, t);
    const nav = h('div', { class: 'recnav' }, '페이지: ', btn('|◀', d.page === 0, () => { d.page = 0; api.renderNow(); }, '첫 페이지'), btn('◀', d.page === 0, () => { d.page--; api.renderNow(); }, '이전 페이지'), pageIn,
      btn('▶', d.page >= count - 1, () => { d.page++; api.renderNow(); }, '다음 페이지'), btn('▶|', d.page >= count - 1, () => { d.page = count - 1; api.renderNow(); }, '마지막 페이지'), h('span', { style: { marginLeft: '10px', color: '#555' } }, `총 ${count}쪽`), h('span', { class: 'sp', style: { flex: 1 } }), h('span', {}, d.zoom ? Math.round(d.zoom * 100) + '%' : '페이지에 맞춤'));
    d.keys = null;
    return h('div', { style: { display: 'flex', flexDirection: 'column', height: '100%' } }, area, nav);
  };

  // ===================================================================== 페이지 설정
  OA.acPageSetup = function (api, src, after) {
    const pg = src.obj.page;
    const n = v => h('input', { type: 'text', value: String(v), style: { width: '70px' } });
    const t = n(pg.top), b = n(pg.bottom), l = n(pg.left), r = n(pg.right);
    const dataOnly = h('input', { type: 'checkbox', checked: !!pg.dataOnly });
    const pP = h('input', { type: 'radio', name: 'ao', checked: pg.orient !== 'landscape' }), pL = h('input', { type: 'radio', name: 'ao', checked: pg.orient === 'landscape' });
    const paper = h('select', {}, ['A4', 'A3', 'B5'].map(x => h('option', { selected: pg.paper === x }, x)));
    const sample = h('div', { style: { width: '90px', height: '120px', border: '1px solid #555', position: 'relative', background: '#fff' } });
    const drawS = () => { sample.innerHTML = ''; const k = 120 / 297; sample.appendChild(h('div', { style: { position: 'absolute', left: (+l.value || 0) * k + 'px', right: (+r.value || 0) * k + 'px', top: (+t.value || 0) * k + 'px', bottom: (+b.value || 0) * k + 'px', background: 'repeating-linear-gradient(#fff 0 5px, #ccc 5px 6px)' } })); };
    [t, b, l, r].forEach(x => x.addEventListener('input', drawS));
    drawS();
    const tb = tabs([
      { id: 'opt', label: '인쇄 옵션', body: h('div', { style: { minHeight: '230px', display: 'flex', gap: '20px' } },
        h('fieldset', { style: { flex: 1 } }, h('legend', {}, '여백(밀리미터)'), h('div', { class: 'grid2' }, '위쪽(T):', t, '아래쪽(B):', b, '왼쪽(F):', l, '오른쪽(G):', r)),
        h('div', {}, h('div', {}, '샘플'), sample, h('label', { class: 'row' }, dataOnly, '데이터만 인쇄(Y)'))) },
      { id: 'page', label: '페이지', body: h('div', { style: { minHeight: '230px' } }, h('fieldset', {}, h('legend', {}, '방향'), h('label', { class: 'row' }, pP, icon('paper', 26), '세로(R)'), h('label', { class: 'row' }, pL, icon('orient', 26), '가로(L)')),
        h('fieldset', {}, h('legend', {}, '용지'), h('div', { class: 'grid2' }, '크기(E):', paper, '원본(S):', h('select', {}, h('option', {}, '자동 선택')))),
        h('fieldset', {}, h('legend', {}, src.kind === 'form' ? '폼 프린터' : '보고서 프린터'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'pr', checked: true }), '기본 프린터(D)'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'pr', disabled: true }), '특정 프린터 사용(N)'))) },
      { id: 'cols', label: '열', body: h('div', { style: { minHeight: '230px' } }, h('fieldset', {}, h('legend', {}, '그리드 설정'), h('div', { class: 'grid2' }, '열 개수(C):', h('input', { type: 'text', value: '1', disabled: true }), '행 간격(W):', h('input', { type: 'text', value: '0cm', disabled: true })))) },
    ], 'opt');
    dialog({ title: '페이지 설정', width: 460, body: tb, buttons: [{ label: '확인', primary: true, onClick: () => {
      const v = x => { const k = parseFloat(x.value); return isFinite(k) && k >= 0 ? k : null; };
      if ([t, b, l, r].some(x => v(x) === null)) { msg({ title: 'Microsoft Access', text: '여백 값이 잘못되었습니다.', icon: 'warn' }); return false; }
      Object.assign(pg, { top: v(t), bottom: v(b), left: v(l), right: v(r), dataOnly: dataOnly.checked, orient: pL.checked ? 'landscape' : 'portrait', paper: paper.value });
      after && after();
    } }, { label: '취소', cancel: true }] });
  };

  // ===================================================================== 보고서 마법사
  OA.acWizard = function (api, onDone) {
    const db = api.db();
    const sources = Object.keys(db.tables).map(n => ['테이블: ' + n, n]).concat(Object.keys(db.queries).map(n => ['쿼리: ' + n, n]));
    let step = 1, source = sources.length ? sources[sources.length - 1][1] : null;
    let avail = source ? A.objFields(db, source) : [], chosen = [], groups = [], sorts = [['', 'asc'], ['', 'asc'], ['', 'asc'], ['', 'asc']], sums = {}, showMode = '자세히 및 요약', layout = '단계', orient = '세로', title = '', after = 'preview';   // Access 기본값: 보고서 미리 보기(13.2-4에서 디자인 수정을 고른다)
    let pick = null, pickC = null, pickG = null;
    const body = h('div', { class: 'wiz' });
    const numeric = f => { try { const r = A.runSaved(db, source); const k = r.columns.indexOf(f); return k >= 0 && r.rows.some(x => typeof x[k] === 'number') && r.rows.every(x => x[k] === null || typeof x[k] === 'number'); } catch (e) { return false; } };
    const listEl = (items, curv, set) => h('div', { class: 'lst' }, items.map(x => { const dv = h('div', { class: x === curv ? 'on' : '' }, x); dv.onclick = () => { set(x); draw(); }; dv.ondblclick = () => { set(x); dbl(x); }; return dv; }));
    let dbl = () => {};
    const draw = () => {
      body.innerHTML = '';
      if (step === 1) {
        dbl = x => { if (avail.includes(x)) { chosen.push(x); avail = avail.filter(y => y !== x); pick = null; draw(); } };
        const srcSel = h('select', { style: { width: '260px' }, onchange: () => { source = srcSel.value; avail = A.objFields(db, source); chosen = []; draw(); } }, sources.map(([l, n]) => h('option', { value: n, selected: n === source }, l)));
        body.append(h('p', {}, '보고서에 사용할 필드를 선택하십시오. 여러 테이블이나 쿼리에서 필드를 선택할 수 있습니다.'), h('div', { class: 'row' }, '테이블/쿼리(T)', srcSel),
          h('div', { class: 'two' }, h('div', {}, '사용 가능한 필드(A):', listEl(avail, pick, x => { pick = x; })),
            h('div', { class: 'mid' }, h('button', { class: 'btn', onclick: () => { if (pick && avail.includes(pick)) { chosen.push(pick); avail = avail.filter(y => y !== pick); pick = null; draw(); } } }, '>'), h('button', { class: 'btn', onclick: () => { chosen = chosen.concat(avail); avail = []; draw(); } }, '>>'),
              h('button', { class: 'btn', onclick: () => { if (pickC) { avail.push(pickC); chosen = chosen.filter(y => y !== pickC); pickC = null; draw(); } } }, '<'), h('button', { class: 'btn', onclick: () => { avail = A.objFields(db, source); chosen = []; draw(); } }, '<<')),
            h('div', {}, '선택한 필드(S):', listEl(chosen, pickC, x => { pickC = x; }))));
      } else if (step === 2) {
        const pool = chosen.filter(f => !groups.includes(f));
        dbl = x => { if (pool.includes(x)) { groups.push(x); draw(); } };
        body.append(h('p', {}, '그룹 수준을 추가하시겠습니까?'), h('div', { class: 'two' }, listEl(pool, pick, x => { pick = x; }),
          h('div', { class: 'mid' }, h('button', { class: 'btn', onclick: () => { if (pick && pool.includes(pick)) { groups.push(pick); pick = null; draw(); } } }, '>'), h('button', { class: 'btn', onclick: () => { if (pickG) { groups = groups.filter(g => g !== pickG); pickG = null; draw(); } } }, '<'),
            h('div', { class: 'note' }, '우선 순위'), h('button', { class: 'btn', onclick: () => { const i = groups.indexOf(pickG); if (i > 0) { [groups[i - 1], groups[i]] = [groups[i], groups[i - 1]]; draw(); } } }, '↑'), h('button', { class: 'btn', onclick: () => { const i = groups.indexOf(pickG); if (i >= 0 && i < groups.length - 1) { [groups[i + 1], groups[i]] = [groups[i], groups[i + 1]]; draw(); } } }, '↓')),
          h('div', { class: 'lst', style: { width: '260px', height: '220px', background: '#fafafa' } }, groups.map((g, i) => { const dv = h('div', { class: g === pickG ? 'on' : '', style: { paddingLeft: 6 + i * 14 + 'px', fontWeight: 700 } }, g); dv.onclick = () => { pickG = g; draw(); }; return dv; }), chosen.filter(f => !groups.includes(f)).map(f => h('div', { style: { paddingLeft: 6 + groups.length * 14 + 'px', color: '#555' } }, f)))),
          h('button', { class: 'btn', disabled: true }, '그룹화 옵션(O)...'));
      } else if (step === 3) {
        const pool = [''].concat(chosen.filter(f => !groups.includes(f)));
        body.append(h('p', {}, '레코드를 정렬할 순서를 지정하십시오. 최대 4개의 필드를 기준으로 오름차순이나 내림차순으로 정렬할 수 있습니다.'),
          ...sorts.map((s, i) => h('div', { class: 'row' }, String(i + 1), h('select', { style: { width: '240px' }, onchange: e => { s[0] = e.target.value; } }, pool.map(f => h('option', { selected: f === s[0] }, f))), h('button', { class: 'btn', onclick: () => { s[1] = s[1] === 'asc' ? 'desc' : 'asc'; draw(); } }, s[1] === 'asc' ? '오름차순' : '내림차순'))),
          h('button', { class: 'btn', disabled: !groups.length, onclick: () => summaryDlg() }, '요약 옵션(O)...'));
      } else if (step === 4) {
        const rad = (name, val, cur, set) => { const i = h('input', { type: 'radio', name, checked: val === cur }); i.onchange = () => set(val); return h('label', { class: 'row' }, i, val); };
        body.append(h('p', {}, '보고서 모양을 선택하십시오.'), h('div', { style: { display: 'flex', gap: '40px' } },
          h('fieldset', {}, h('legend', {}, '모양'), ['단계', '블록', '외곽선'].map(v => rad('lay', v, layout, x => { layout = x; }))),
          h('fieldset', {}, h('legend', {}, '용지 방향'), ['세로', '가로'].map(v => rad('ori', v, orient, x => { orient = x; })))),
          h('label', { class: 'row' }, h('input', { type: 'checkbox', checked: true }), '모든 필드를 한 페이지에 맞게 필드 너비 조정(W)'));
      } else {
        if (!title) title = source;
        const ti = h('input', { type: 'text', value: title, style: { width: '300px' }, oninput: () => { title = ti.value; } });
        const r1 = h('input', { type: 'radio', name: 'aft', checked: after === 'preview', onchange: () => { after = 'preview'; } }), r2 = h('input', { type: 'radio', name: 'aft', checked: after === 'design', onchange: () => { after = 'design'; } });
        body.append(h('p', {}, '보고서 제목을 지정하십시오.'), ti, h('p', {}, '보고서를 미리 보시겠습니까, 아니면 보고서 디자인을 수정하시겠습니까?'), h('label', { class: 'row' }, r1, '보고서 미리 보기(P)'), h('label', { class: 'row' }, r2, '보고서 디자인 수정(M)'));
      }
      dlgApi && dlgApi.buttons && (dlgApi.buttons['< 뒤로(B)'].disabled = step === 1, dlgApi.buttons['다음(N) >'].disabled = step === 5);
    };
    function summaryDlg() {
      const nums = chosen.filter(f => !groups.includes(f) && numeric(f));
      const grid = h('div', { style: { display: 'grid', gridTemplateColumns: '120px repeat(4, 50px)', gap: '4px', alignItems: 'center' } }, h('b', {}, '필드'), h('b', {}, '합계'), h('b', {}, '평균'), h('b', {}, '최소'), h('b', {}, '최대'));
      for (const f of nums) {
        sums[f] = sums[f] || {};
        grid.append(h('span', {}, f), ...['sum', 'avg', 'min', 'max'].map(k => { const cb = h('input', { type: 'checkbox', checked: !!sums[f][k] }); cb.onchange = () => { sums[f][k] = cb.checked; }; return h('span', { style: { textAlign: 'center' } }, cb); }));
      }
      const r1 = h('input', { type: 'radio', name: 'sm', checked: showMode === '자세히 및 요약' }), r2 = h('input', { type: 'radio', name: 'sm', checked: showMode === '요약만' });
      dialog({ title: '요약 옵션', width: 520, body: h('div', { style: { display: 'flex', gap: '20px' } }, h('div', {}, h('p', {}, '어떤 요약 값을 계산하시겠습니까?'), grid), h('fieldset', {}, h('legend', {}, '표시'), h('label', { class: 'row' }, r1, '자세히 및 요약(D)'), h('label', { class: 'row' }, r2, '요약만(S)'), h('label', { class: 'row' }, h('input', { type: 'checkbox', disabled: true }), '합계에 대한 백분율 계산(P)'))),
        buttons: [{ label: '확인', primary: true, onClick: () => { showMode = r2.checked ? '요약만' : '자세히 및 요약'; } }, { label: '취소', cancel: true }] });
    }
    let dlgApi = null;
    dlgApi = dialog({ title: '보고서 마법사', width: 680, body, buttons: [
      { label: '취소', cancel: true },
      { label: '< 뒤로(B)', onClick: () => { step = Math.max(1, step - 1); draw(); return false; } },
      { label: '다음(N) >', onClick: () => { if (step === 1 && !chosen.length) { msg({ title: 'Microsoft Access', text: '필드를 하나 이상 선택해야 합니다.', icon: 'warn' }); return false; } step = Math.min(5, step + 1); draw(); return false; } },
      { label: '마침(F)', primary: true, onClick: () => {
        if (!chosen.length) { msg({ title: 'Microsoft Access', text: '필드를 하나 이상 선택해야 합니다.', icon: 'warn' }); return false; }
        const sumFields = Object.keys(sums).filter(f => sums[f].sum);
        const rpt = A.reportWizard(db, { source, fields: chosen, groups, sorts: sorts.filter(s => s[0]), sums: groups.length ? sumFields : [], layout, orient, title: title || source });
        const name = (title || source).trim();
        let nm = name, k = 1;
        while (db.reports[nm]) nm = name + (k++);
        A.saveReport(db, nm, rpt);
        api.changed();
        const d = api.openDoc({ kind: 'report', name: nm, obj: U.clone(db.reports[nm]), type: after === 'design' ? 'reportDesign' : 'reportView' });
        if (after === 'preview') api.openPreview(d);
        void onDone;
      } }] });
    draw();
  };
})(typeof window !== 'undefined' ? window : globalThis);

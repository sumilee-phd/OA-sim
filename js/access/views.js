/* Access: 관계 창, 쿼리 디자인·SQL·결과, 테이블 추가 창 */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const { h, icon, menu, dialog, msg, tabs } = OA.ui;
  const U = OA.util, A = OA.acc;

  function addPane(api, title, onAdd, onClose, withQueries) {
    const db = api.db();
    let mode = 'tables';
    const list = h('div', { class: 'lst', style: { height: '260px' } });
    let picked = null;
    const fill = () => {
      list.innerHTML = '';
      const names = (mode === 'tables' ? Object.keys(db.tables) : mode === 'queries' ? Object.keys(db.queries) : Object.keys(db.tables).concat(Object.keys(db.queries)));
      for (const n of names) {
        const d = h('div', {}, n);
        d.addEventListener('click', () => { picked = n; [...list.children].forEach(c => c.classList.toggle('on', c === d)); });
        d.addEventListener('dblclick', () => onAdd(n));
        list.appendChild(d);
      }
      if (!names.length) list.appendChild(h('div', { class: 'note' }, '항목이 없습니다.'));
    };
    const tabsEl = h('div', { class: 'dtabs' }, [['tables', '테이블'], ['links', '링크'], ['queries', '쿼리'], ['all', '모두']].map(([k, l]) => h('button', { class: k === mode ? 'on' : '', disabled: k === 'links' || (!withQueries && k === 'queries'), onclick: e => { mode = k; [...tabsEl.children].forEach(b => b.classList.toggle('on', b === e.target)); fill(); } }, l)));
    fill();
    return h('div', { class: 'pane' }, h('div', { class: 'pane-h' }, h('span', { class: 't' }, title), h('button', { title: '닫기', onclick: onClose }, '✕')),
      h('div', { class: 'pane-b' }, h('input', { placeholder: '검색', style: { width: '100%', height: '24px', marginBottom: '6px' }, onkeydown: e => e.stopPropagation() }), tabsEl, list,
        h('button', { class: 'btn', style: { marginTop: '8px' }, onclick: () => picked && onAdd(picked) }, '선택한 테이블 추가')));
  }
  function fieldsOf(db, name) { return A.objFields(db, name); }
  function tableBox(db, name, pos, opts) {
    const t = db.tables[name];
    const fl = h('div', { class: 'fl' }, opts.star ? h('div', { 'data-f': '*' }, '*') : null, fieldsOf(db, name).map(f => h('div', { class: t && t.pk.includes(f) ? 'pk' : '', 'data-f': f }, f)));
    const box = h('div', { class: 'tbox', style: { left: pos.x + 'px', top: pos.y + 'px' } }, h('div', { class: 'h' }, name), fl);
    box.dataset.t = name;
    box.querySelector('.h').addEventListener('mousedown', e => {
      e.preventDefault();
      const sx = e.clientX - pos.x, sy = e.clientY - pos.y;
      OA.ui.drag(e, ev => { pos.x = Math.max(0, ev.clientX - sx); pos.y = Math.max(0, ev.clientY - sy); box.style.left = pos.x + 'px'; box.style.top = pos.y + 'px'; opts.onMove && opts.onMove(); }, () => opts.onMoved && opts.onMoved());
    });
    box.querySelector('.h').addEventListener('contextmenu', e => { e.preventDefault(); if (opts.onRemove) menu(box, [{ label: '테이블 숨기기(H)', onClick: opts.onRemove }], { at: { x: e.clientX, y: e.clientY } }); });
    return box;
  }
  function fieldY(box, f) {
    const el = box.querySelector(`[data-f="${CSS.escape(f)}"]`);
    if (!el) return box.offsetTop + 20;
    return box.offsetTop + el.offsetTop + el.offsetHeight / 2 + box.querySelector('.h').offsetHeight - box.querySelector('.fl').scrollTop + 2;
  }
  function drawLines(svg, desk, pairs) {
    svg.innerHTML = '';
    for (const p of pairs) {
      const a = desk.querySelector(`.tbox[data-t="${CSS.escape(p.t1)}"]`), b = desk.querySelector(`.tbox[data-t="${CSS.escape(p.t2)}"]`);
      if (!a || !b) continue;
      const aR = a.offsetLeft + a.offsetWidth, bL = b.offsetLeft;
      const leftToRight = a.offsetLeft <= b.offsetLeft;
      const x1 = leftToRight ? aR : a.offsetLeft, x2 = leftToRight ? bL : b.offsetLeft + b.offsetWidth;
      const y1 = fieldY(a, p.f1), y2 = fieldY(b, p.f2);
      const d1 = leftToRight ? 12 : -12;
      const path = `M${x1} ${y1} H${x1 + d1} L${x2 - d1} ${y2} H${x2}`;
      const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      g.innerHTML = `<path d="${path}" stroke="transparent" stroke-width="10" fill="none" style="cursor:pointer;pointer-events:stroke"/><path d="${path}" stroke="${p.sel ? '#000' : '#333'}" stroke-width="${p.sel ? 3 : 1.4}" fill="none" style="pointer-events:none"/>` +
        (p.ri ? `<text x="${x1 + d1 / 2}" y="${y1 - 4}" font-size="12" text-anchor="middle" font-weight="700">1</text><text x="${x2 - d1 / 2}" y="${y2 - 4}" font-size="13" text-anchor="middle" font-weight="700">∞</text>` : '');
      g.style.pointerEvents = 'auto';
      if (p.onClick) g.addEventListener('click', p.onClick);
      if (p.onDbl) g.addEventListener('dblclick', p.onDbl);
      svg.appendChild(g);
    }
  }

  // ===================================================================== 관계
  OA.acRel = {
    view(api, d, side) {
      const db = api.db();
      const desk = h('div', { class: 'reldesk', style: { height: '100%', minHeight: '520px' } });
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('style', 'position:absolute;inset:0;width:100%;height:100%;pointer-events:none;overflow:visible');
      desk.appendChild(svg);
      const lay = d.layout;
      lay.tables = lay.tables.filter(t => db.tables[t.name]);
      let selRel = d.selRel || null;
      const redraw = () => drawLines(svg, desk, db.relationships.filter(r => lay.tables.some(t => t.name === r.t1) && lay.tables.some(t => t.name === r.t2)).map(r => ({ ...r, sel: r === selRel, onClick: () => { selRel = r; d.selRel = r; redraw(); }, onDbl: () => editRel(r) })));
      for (const t of lay.tables) {
        const box = tableBox(db, t.name, t, { onMove: redraw, onMoved: () => { d.layoutDirty = true; }, onRemove: () => { lay.tables = lay.tables.filter(x => x !== t); d.layoutDirty = true; api.renderNow(); } });
        box.querySelectorAll('.fl div').forEach(fd => fd.addEventListener('mousedown', e => {
          e.preventDefault();
          const f1 = fd.dataset.f, t1 = t.name;
          fd.classList.add('dragsrc');
          const ghost = h('div', { class: 'dragghost' }, f1);
          document.body.appendChild(ghost);
          OA.ui.drag(e, ev => { ghost.style.left = ev.clientX + 8 + 'px'; ghost.style.top = ev.clientY + 8 + 'px'; }, ev => {
            ghost.remove(); fd.classList.remove('dragsrc');
            const tgt = document.elementFromPoint(ev.clientX, ev.clientY);
            const tb = tgt && tgt.closest('.tbox'), ff = tgt && tgt.closest('.fl div');
            if (!tb || !ff || tb.dataset.t === t1) return;
            relDialog(api, { t1, f1, t2: tb.dataset.t, f2: ff.dataset.f, ri: false }, true, () => { redraw(); api.renderNow(); });
          });
        }));
        box.querySelector('.fl').addEventListener('scroll', redraw);
        desk.appendChild(box);
      }
      setTimeout(redraw, 0);
      function editRel(r) { relDialog(api, Object.assign({}, r, { existing: r }), false, () => { redraw(); api.renderNow(); }); }
      d.rel = {
        showAdd() { d.showAdd = true; api.renderNow(); },
        editSelected() { if (selRel) editRel(selRel); else relDialog(api, { t1: lay.tables[0] && lay.tables[0].name, f1: '', t2: lay.tables[1] && lay.tables[1].name, f2: '', ri: false }, true, () => api.renderNow()); },
        clearLayout() { msg({ title: 'Microsoft Access', text: '관계 창에서 레이아웃을 지웁니다. 관계는 삭제되지 않습니다. 계속하시겠습니까?', buttons: ['예', '아니요'], icon: 'q' }).then(r => { if (r === '예') { lay.tables = []; d.layoutDirty = true; api.renderNow(); } }); },
        showAll() { const names = new Set(db.relationships.flatMap(r => [r.t1, r.t2])); for (const n of names) if (!lay.tables.some(t => t.name === n)) lay.tables.push({ name: n, x: 30 + lay.tables.length * 230, y: 30 }); d.layoutDirty = true; api.renderNow(); },
      };
      d.keys = e => {
        if (e.key === 'Delete' && selRel) {
          e.preventDefault();
          msg({ title: 'Microsoft Access', text: '선택한 관계를 데이터베이스에서 영구적으로 삭제하시겠습니까?', buttons: ['예', '아니요'], icon: 'q' }).then(r => { if (r === '예') { A.removeRelationship(db, selRel); selRel = null; d.selRel = null; api.changed(); api.renderNow(); } });
        }
      };
      if (d.showAdd) side.appendChild(addPane(api, '테이블 추가', n => { if (!lay.tables.some(t => t.name === n)) { lay.tables.push({ name: n, x: 30 + lay.tables.length * 230, y: 30 }); d.layoutDirty = true; api.renderNow(); } }, () => { d.showAdd = false; api.renderNow(); }, false));
      return desk;
    },
  };
  function relDialog(api, r, isNew, done) {
    const db = api.db();
    const tA = h('select', {}, Object.keys(db.tables).map(n => h('option', { selected: n === r.t1 }, n)));
    const tB = h('select', {}, Object.keys(db.tables).map(n => h('option', { selected: n === r.t2 }, n)));
    const fA = h('select', {}), fB = h('select', {});
    const fillF = () => { fA.innerHTML = ''; fB.innerHTML = ''; fieldsOf(db, tA.value).forEach(f => fA.appendChild(h('option', { selected: f === r.f1 }, f))); fieldsOf(db, tB.value).forEach(f => fB.appendChild(h('option', { selected: f === r.f2 }, f))); };
    tA.disabled = !isNew; tB.disabled = !isNew;
    fillF();
    const ri = h('input', { type: 'checkbox', checked: !!r.ri });
    const cu = h('input', { type: 'checkbox', checked: !!r.cascadeUpd, disabled: !r.ri }), cd = h('input', { type: 'checkbox', checked: !!r.cascadeDel, disabled: !r.ri });
    ri.addEventListener('change', () => { cu.disabled = cd.disabled = !ri.checked; });
    const kind = h('span', {}, '');
    const updKind = () => { const a = db.tables[tA.value], b = db.tables[tB.value]; kind.textContent = (A.isUnique(a, fA.value) || A.isUnique(b, fB.value)) ? (A.isUnique(a, fA.value) && A.isUnique(b, fB.value) ? '일대일' : '일대다') : '정의되지 않음'; };
    [fA, fB].forEach(x => x.addEventListener('change', updKind));
    updKind();
    dialog({ title: '관계 편집', width: 520, body: h('div', {},
      h('div', { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '4px 10px' } }, '테이블/쿼리(T):', '관련 테이블/쿼리(R):', tA, tB, fA, fB),
      h('label', { class: 'row', style: { marginTop: '10px' } }, ri, '항상 참조 무결성 유지(E)'), h('label', { class: 'row' }, cu, '관련 필드 모두 업데이트(U)'), h('label', { class: 'row' }, cd, '관련 레코드 모두 삭제(D)'),
      h('div', { class: 'row' }, '관계 종류:', kind)),
      buttons: [{ label: '조인 유형(J)...', left: true, onClick: () => { dialog({ title: '조인 속성', width: 460, body: h('div', {}, h('label', { class: 'row' }, h('input', { type: 'radio', name: 'jt', checked: true }), '1: 두 테이블의 조인된 필드가 같은 행만 포함'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'jt', disabled: true }), '2: 왼쪽 테이블의 모든 레코드 포함'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'jt', disabled: true }), '3: 오른쪽 테이블의 모든 레코드 포함')) }); return false; } },
        { label: isNew ? '만들기(C)' : '확인', primary: true, onClick: () => {
          if (r.existing) A.removeRelationship(db, r.existing);
          const res = A.addRelationship(db, { t1: tA.value, f1: fA.value, t2: tB.value, f2: fB.value, ri: ri.checked, cascadeUpd: cu.checked, cascadeDel: cd.checked });
          if (res.error) { if (r.existing) db.relationships.push(r.existing); msg({ title: 'Microsoft Access', text: res.error, icon: 'warn' }); return false; }
          api.changed(); done && done();
        } }, { label: '취소', cancel: true }, { label: '새로 만들기(N)...', disabled: true }] });
  }

  // ===================================================================== 쿼리
  const SORTS = ['', '오름차순', '내림차순', '(정렬 안 함)'];
  OA.acQuery = {
    design(api, d, side) {
      const db = api.db();
      const q = d.obj;
      const wrap = h('div', { style: { display: 'flex', flexDirection: 'column', height: '100%' } });
      const top = h('div', { class: 'qtop' });
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('style', 'position:absolute;inset:0;width:100%;height:100%;pointer-events:none;overflow:visible');
      top.appendChild(svg);
      const touch = () => { if (q.mode === 'sql') { q.mode = 'design'; } if (!d.dirty) { d.dirty = true; api.renderNow(); } };
      const redraw = () => drawLines(svg, top, q.joins.map(j => ({ t1: j.t1, f1: j.f1, t2: j.t2, f2: j.f2, ri: !!db.relationships.find(r => r.t1 === j.t1 && r.t2 === j.t2 && r.ri) })));
      for (const t of q.tables) {
        const box = tableBox(db, t.name, t, { star: true, onMove: redraw, onRemove: () => { api.snapshot(d); q.tables = q.tables.filter(x => x !== t); q.joins = q.joins.filter(j => j.t1 !== t.name && j.t2 !== t.name); q.cols = q.cols.filter(c => c.table !== t.name); touch(); api.renderNow(); } });
        box.querySelectorAll('.fl div').forEach(fd => {
          fd.addEventListener('dblclick', () => { api.snapshot(d); const f = fd.dataset.f; q.cols.push({ field: f === '*' ? `${t.name}.*` : f, table: t.name, sort: '', show: true, crit: [] }); touch(); api.renderNow(); });
          fd.addEventListener('mousedown', e => {
            e.preventDefault();
            const ghost = h('div', { class: 'dragghost' }, fd.dataset.f);
            document.body.appendChild(ghost);
            OA.ui.drag(e, ev => { ghost.style.left = ev.clientX + 8 + 'px'; ghost.style.top = ev.clientY + 8 + 'px'; }, ev => {
              ghost.remove();
              const tg = document.elementFromPoint(ev.clientX, ev.clientY);
              const cell = tg && tg.closest('td[data-col]');
              const tb = tg && tg.closest('.tbox'), ff = tg && tg.closest('.fl div');
              if (cell) { api.snapshot(d); const ci = +cell.dataset.col; const f = fd.dataset.f; q.cols.splice(Math.min(ci, q.cols.length), 0, { field: f === '*' ? `${t.name}.*` : f, table: t.name, sort: '', show: true, crit: [] }); touch(); api.renderNow(); }
              else if (tb && ff && tb.dataset.t !== t.name) { api.snapshot(d); q.joins.push({ t1: t.name, f1: fd.dataset.f, t2: tb.dataset.t, f2: ff.dataset.f, kind: 'INNER' }); touch(); api.renderNow(); }
            });
          });
        });
        top.appendChild(box);
      }
      setTimeout(redraw, 0);
      // 디자인 눈금
      const ncol = Math.max(q.cols.length + 3, 8);
      const grid = h('table', { class: 'qgrid' });
      const head = h('tr', { class: 'colhead' }, h('th', {}));
      for (let i = 0; i < ncol; i++) head.appendChild(h('td', { class: d.selCol === i ? 'on' : '', onclick: () => { d.selCol = i; api.renderNow(); } }));
      grid.appendChild(head);
      const allFields = [];
      for (const t of q.tables) { allFields.push(`${t.name}.*`); for (const f of fieldsOf(db, t.name)) allFields.push(f); }
      const dl = h('datalist', { id: d.id + '_fl' }, [...new Set(allFields)].map(f => h('option', { value: f })));
      const zoomBox = (inp, apply) => inp.addEventListener('keydown', e => {
        e.stopPropagation();
        if (e.key === 'F2' && e.shiftKey) {
          e.preventDefault();
          const ta = h('textarea', { style: { width: '100%', height: '160px', fontSize: '13px' } }, inp.value);
          dialog({ title: '확대/축소', width: 560, help: false, body: ta, buttons: [{ label: '확인', primary: true, onClick: () => { inp.value = ta.value; apply(ta.value); } }, { label: '취소', cancel: true }, { label: '글꼴(F)...', left: true, disabled: true }] });
        }
      });
      const rowsDef = [['필드:', 'field'], ['테이블:', 'table'], ['정렬:', 'sort'], ['표시:', 'show'], ['조건:', 'c0'], ['또는:', 'c1'], ['', 'c2'], ['', 'c3'], ['', 'c4']];
      for (const [label, key] of rowsDef) {
        const tr = h('tr', {}, h('th', {}, label));
        for (let i = 0; i < ncol; i++) {
          const c = q.cols[i];
          const ensure = () => { while (q.cols.length <= i) q.cols.push({ field: '', table: '', sort: '', show: true, crit: [] }); return q.cols[i]; };
          let el;
          if (key === 'field') {
            el = h('input', { value: c ? c.field : '', list: d.id + '_fl', onchange: () => { api.snapshot(d); const col = ensure(); const v = el.value.trim(); col.field = v; if (!A.parseCalc(v) && v && !v.endsWith('.*')) { const tb = A.fieldTable(db, q, v); col.table = tb || col.table; } if (A.parseCalc(v)) col.table = ''; if (!v) q.cols.splice(i, 1); touch(); api.renderNow(); } });
            zoomBox(el, v => { api.snapshot(d); const col = ensure(); col.field = v; if (A.parseCalc(v)) col.table = ''; touch(); api.renderNow(); });
          } else if (key === 'table') {
            el = h('select', { onchange: () => { api.snapshot(d); ensure().table = el.value; touch(); } }, h('option', { value: '' }, ''), q.tables.map(t => h('option', { selected: c && c.table === t.name }, t.name)));
            if (!c) el.disabled = true;
          } else if (key === 'sort') {
            el = h('select', { onchange: () => { api.snapshot(d); ensure().sort = el.value === '(정렬 안 함)' ? '' : el.value; touch(); } }, SORTS.map(s => h('option', { selected: c && c.sort === s }, s)));
            if (!c) el.disabled = true;
          } else if (key === 'show') {
            el = h('input', { type: 'checkbox', checked: c ? c.show !== false : false, disabled: !c, onchange: () => { api.snapshot(d); ensure().show = el.checked; touch(); } });
          } else {
            const k = +key.slice(1);
            el = h('input', { value: c && c.crit[k] ? c.crit[k] : '', onchange: () => { api.snapshot(d); const col = ensure(); col.crit[k] = el.value; touch(); } });
            zoomBox(el, v => { api.snapshot(d); ensure().crit[k] = v; touch(); api.renderNow(); });
          }
          if (el.tagName === 'INPUT' && el.type !== 'checkbox') el.addEventListener('keydown', e => e.stopPropagation());
          tr.appendChild(h('td', { class: key === 'show' ? 'cb' : '', 'data-col': i }, el));
        }
        grid.appendChild(tr);
      }
      const bottom = h('div', { style: { flex: 1, overflow: 'auto', background: '#fff' } }, dl, grid);
      wrap.append(top, bottom);
      d.qd = {
        showAdd() { d.showAdd = true; api.renderNow(); },
        insertCol() { api.snapshot(d); q.cols.splice(d.selCol || 0, 0, { field: '', table: '', sort: '', show: true, crit: [] }); touch(); api.renderNow(); },
        deleteCol() { if (d.selCol === undefined || !q.cols[d.selCol]) return; api.snapshot(d); q.cols.splice(d.selCol, 1); touch(); api.renderNow(); },
        sync() { q.cols = q.cols.filter(c => c.field); },
      };
      d.keys = e => { if (e.key === 'Delete' && d.selCol !== undefined && q.cols[d.selCol] && !(e.target && /INPUT|SELECT/.test(e.target.tagName))) { e.preventDefault(); d.qd.deleteCol(); } };
      if (d.showAdd) side.appendChild(addPane(api, '테이블 추가', n => { api.snapshot(d); A.addQueryTable(db, q, n); touch(); api.renderNow(); }, () => { d.showAdd = false; api.renderNow(); }, true));
      return wrap;
    },
    sql(api, d) {
      const db = api.db();
      if (!d.obj.sql || d.obj.mode !== 'sql') { try { d.obj.sql = A.designToSQL(db, d.obj); } catch (e) { d.obj.sql = 'SELECT;'; } }
      const ta = h('textarea', { class: 'sqlview', spellcheck: false }, d.obj.sql);
      ta.addEventListener('input', () => { d.obj.sql = ta.value; d.obj.mode = 'sql'; if (!d.dirty) { d.dirty = true; } });
      ta.addEventListener('keydown', e => { if (e.ctrlKey && (e.code === 'KeyS')) return; e.stopPropagation(); });
      setTimeout(() => ta.focus(), 0);
      d.keys = null;
      return h('div', { style: { height: '100%' } }, ta);
    },
    result(api, d) {
      const db = api.db();
      const r = d.result || { columns: [], rows: [], fieldInfo: [] };
      const wrap = h('div', { style: { display: 'flex', flexDirection: 'column', height: '100%' } });
      const tbl = h('table', { class: 'ds' }, h('tr', {}, h('th', {}), r.columns.map(c => h('th', {}, c))));
      r.rows.forEach((row, i) => tbl.appendChild(h('tr', { class: i === 0 ? 'cur' : '' }, h('td', { class: 'rs' }, i === 0 ? '▶' : ''), row.map((v, k) => h('td', { style: { padding: '0 6px', textAlign: typeof v === 'number' ? 'right' : 'left' } }, A.formatValue(v, r.fieldInfo && r.fieldInfo[k]))))));
      tbl.appendChild(h('tr', {}, h('td', { class: 'rs' }, '*'), r.columns.map(() => h('td'))));
      if (d.totals) {
        d.totFn = d.totFn || {};
        tbl.appendChild(h('tr', { class: 'tot' }, h('td', { class: 'rs' }, 'Σ'), r.columns.map((c, k) => {
          const vals = r.rows.map(x => x[k]).filter(v => typeof v === 'number');
          const fn = d.totFn[c] || '';
          const sel = h('select', { onchange: () => { d.totFn[c] = sel.value; api.renderNow(); } }, ['', '합계', '평균', '개수'].map(o => h('option', { value: o, selected: o === fn }, o || '없음')));
          const val = fn === '합계' ? vals.reduce((s, x) => s + x, 0) : fn === '평균' ? vals.reduce((s, x) => s + x, 0) / (vals.length || 1) : fn === '개수' ? r.rows.length : null;
          return h('td', {}, k === 0 && !fn ? '요약 ' : '', val !== null ? A.formatValue(val, fn === '개수' ? {} : r.fieldInfo[k]) + ' ' : '', vals.length || k > 0 ? sel : null);
        })));
      }
      wrap.append(h('div', { style: { flex: 1, overflow: 'auto' } }, tbl), h('div', { class: 'recnav' }, '레코드: |◀ ◀ ', h('input', { value: `${r.rows.length ? 1 : 0}/${r.rows.length}`, readOnly: true }), ' ▶ ▶| ▶*'));
      d.keys = null;
      return wrap;
    },
  };
})(typeof window !== 'undefined' ? window : globalThis);

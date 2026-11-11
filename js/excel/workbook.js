/* Excel 통합 문서 모델(DOM 없음). 상태는 JSON 으로 저장할 수 있는 순수 객체, 계산 캐시는 열거되지 않는 속성. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util, NF = OA.numfmt, FX = OA.formula;

  const DEF_COL = 72, DEF_ROW = 22;          // 한국어 Excel 기본 열 너비 8.38(72픽셀), 행 높이 16.5pt(22픽셀)
  const MAX_ROW = 200, MAX_COL = 30;
  const PX_MM = 25.4 / 96;
  const DEFAULT_PAGE = { orient: 'portrait', paper: 'A4', top: 1.9, bottom: 1.9, left: 1.8, right: 1.8, header: 0.8, footer: 0.8,
    centerH: false, centerV: false, scaleMode: 'zoom', zoom: 100, fitW: 1, fitH: 1 };

  function newSheet(name) {
    return { name, cells: {}, colW: {}, rowH: {}, merges: [], charts: [], pivots: [], printArea: null, page: U.clone(DEFAULT_PAGE) };
  }
  function newWorkbook() {
    const wb = { sheets: [newSheet('Sheet1')], active: 0, caches: [], seq: { chart: 1, pivot: 1, cache: 1 } };
    return attach(wb);
  }
  function attach(wb) {
    Object.defineProperty(wb, '__calc', { value: new Map(), writable: true, enumerable: false, configurable: true });
    Object.defineProperty(wb, '__busy', { value: new Set(), writable: true, enumerable: false, configurable: true });
    Object.defineProperty(wb, '__pv', { value: new Map(), writable: true, enumerable: false, configurable: true });
    return wb;
  }
  function fromJSON(o) { return attach(typeof o === 'string' ? JSON.parse(o) : U.clone(o)); }
  function touch(wb) { wb.__calc = new Map(); wb.__pv = new Map(); }
  const key = (r, c) => r + ',' + c;

  function sheetOf(wb, s) {
    if (s === undefined || s === null) return wb.sheets[wb.active];
    if (typeof s === 'object') return s;
    if (typeof s === 'number') return wb.sheets[s];
    return wb.sheets.find(x => x.name === s) || null;
  }
  function getCell(sh, r, c) { return sh.cells[key(r, c)] || null; }
  function ensureCell(sh, r, c) { return sh.cells[key(r, c)] || (sh.cells[key(r, c)] = {}); }
  function cleanCell(sh, r, c) {
    const k = key(r, c), cl = sh.cells[k];
    if (cl && (cl.v === undefined || cl.v === null) && !cl.f && (!cl.s || Object.keys(cl.s).length === 0)) delete sh.cells[k];
  }

  // ------------------------------------------------------------------ 계산
  function pivotCellAt(wb, sh, r, c) {
    if (!sh.pivots.length || !OA.pivot) return null;
    for (const pv of sh.pivots) {
      const g = OA.pivot.grid(wb, sh, pv);
      if (!g) continue;
      if (r >= g.r0 && r < g.r0 + g.rows.length && c >= g.c0 && c < g.c0 + g.ncol) {
        const cell = g.rows[r - g.r0][c - g.c0];
        return cell || { v: null };
      }
    }
    return null;
  }
  function getValue(wb, sName, r, c) {
    const sh = sheetOf(wb, sName);
    if (!sh) return FX.E.REF;
    const pc = pivotCellAt(wb, sh, r, c);
    if (pc) return pc.v === undefined ? null : pc.v;
    const cl = getCell(sh, r, c);
    if (!cl) return null;
    if (!cl.f) return cl.v === undefined ? null : cl.v;
    const ck = sh.name + '|' + key(r, c);
    if (wb.__calc.has(ck)) return wb.__calc.get(ck);
    if (wb.__busy.has(ck)) return 0;                   // 순환 참조
    wb.__busy.add(ck);
    const v = FX.evaluate(cl.f.slice(1), { get: (s2, r2, c2) => getValue(wb, s2, r2, c2), sheet: sh.name, row: r, col: c });
    wb.__busy.delete(ck);
    wb.__calc.set(ck, v);
    return v;
  }
  function cellFormat(wb, sh, r, c) {
    const pc = pivotCellAt(wb, sh, r, c);
    if (pc) return pc.nf || NF.GENERAL;
    const cl = getCell(sh, r, c);
    return (cl && cl.s && cl.s.nf) || NF.GENERAL;
  }
  function display(wb, sName, r, c) {
    const sh = sheetOf(wb, sName);
    const pc = pivotCellAt(wb, sh, r, c);
    if (pc) {
      if (pc.text !== undefined) return { text: pc.text, align: pc.align || 'left' };
      return NF.format(pc.v, pc.nf);
    }
    const cl = getCell(sh, r, c);
    if (!cl) return { text: '', align: 'left' };
    const v = getValue(wb, sh, r, c);
    const f = NF.format(v, (cl.s && cl.s.nf) || NF.GENERAL);
    return f;
  }
  /* 수식 입력줄에 보이는 내용 */
  function editText(wb, sName, r, c) {
    const sh = sheetOf(wb, sName);
    const pc = pivotCellAt(wb, sh, r, c);
    if (pc) return pc.text !== undefined ? pc.text : (pc.v === null || pc.v === undefined ? '' : String(pc.v));
    const cl = getCell(sh, r, c);
    if (!cl) return '';
    if (cl.f) return cl.f;
    if (cl.v === null || cl.v === undefined) return '';
    if (cl.q) return "'" + cl.v;
    if (typeof cl.v === 'number') {
      const nf = cl.s && cl.s.nf;
      if (nf && NF.isDateFmt(nf)) {
        if (/[yd]/i.test(nf)) { const p = NF.serialParts(cl.v); return `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`; }
        const p = NF.serialParts(cl.v); return `${p.hh}:${String(p.mi).padStart(2, '0')}:00`.replace(/:00$/, ':00');
      }
      if (nf && /%/.test(nf)) return NF.generalNum(+(cl.v * 100).toPrecision(12)) + '%';
      return NF.generalNum(cl.v);
    }
    if (typeof cl.v === 'boolean') return cl.v ? 'TRUE' : 'FALSE';
    return String(cl.v);
  }

  // ------------------------------------------------------------------ 병합
  function mergeAt(sh, r, c) { return sh.merges.find(m => U.inRange(m, r, c)) || null; }
  function merge(wb, sName, rg, mode) {
    const sh = sheetOf(wb, sName);
    sh.merges = sh.merges.filter(m => !(m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1));
    if (mode === 'unmerge') { touch(wb); return; }
    const rows = mode === 'across' ? [...Array(rg.r2 - rg.r1 + 1).keys()].map(k => ({ r1: rg.r1 + k, r2: rg.r1 + k })) : [{ r1: rg.r1, r2: rg.r2 }];
    for (const rr of rows) {
      const m = { r1: rr.r1, c1: rg.c1, r2: rr.r2, c2: rg.c2 };
      if (m.r1 === m.r2 && m.c1 === m.c2) continue;
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) {
        if (r === m.r1 && c === m.c1) continue;
        const cl = getCell(sh, r, c);
        if (cl) { delete cl.v; delete cl.f; delete cl.q; cleanCell(sh, r, c); }
      }
      sh.merges.push(m);
      if (mode === 'center') { const cl = ensureCell(sh, m.r1, m.c1); cl.s = Object.assign({}, cl.s, { ha: 'center' }); }
    }
    touch(wb);
  }
  function lostByMerge(sh, rg) {
    let n = 0;
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      if (r === rg.r1 && c === rg.c1) continue;
      const cl = getCell(sh, r, c);
      if (cl && (cl.f || (cl.v !== null && cl.v !== undefined && cl.v !== ''))) n++;
    }
    return n;
  }

  // ------------------------------------------------------------------ 입력
  function setInput(wb, sName, r, c, text) {
    const sh = sheetOf(wb, sName);
    if (OA.pivot) {
      const hit = OA.pivot.hitTest(wb, sh, r, c);
      if (hit) return OA.pivot.typeInto(wb, sh, hit, text);
    }
    const m = mergeAt(sh, r, c);
    if (m && (m.r1 !== r || m.c1 !== c)) { r = m.r1; c = m.c1; }
    const cl = ensureCell(sh, r, c);
    delete cl.f; delete cl.q; delete cl.v;
    const s = text == null ? '' : String(text);
    if (s.startsWith('=') && s.length > 1) {
      const err = FX.check(s);
      if (err) { cleanCell(sh, r, c); return { error: '입력한 수식에 문제가 있습니다.\n\n' + s }; }
      cl.f = normalizeFormula(s);
      // 함수 결과의 서식 상속(시간 빼기·더하기는 h:mm, 날짜 함수는 날짜)
      if (!cl.s || !cl.s.nf) {
        const inh = inheritFormat(wb, sh, cl.f);
        if (inh) cl.s = Object.assign({}, cl.s, { nf: inh });
      }
    } else {
      const p = NF.parseInput(s);
      if (p.v === null || p.v === undefined) { cleanCell(sh, r, c); touch(wb); return {}; }
      cl.v = p.v;
      if (p.q) cl.q = true;
      if (p.nf && (!cl.s || !cl.s.nf || cl.s.nf === NF.GENERAL)) cl.s = Object.assign({}, cl.s, { nf: p.nf });
    }
    touch(wb);
    return {};
  }
  function normalizeFormula(s) {
    // 함수 이름은 대문자로, 나머지는 입력 그대로
    const body = s.slice(1);
    const toks = FX.tokenize(body);
    let out = '', pos = 0;
    for (const t of toks) {
      if (t.k === 'fn' || t.k === 'range' || t.k === 'ref') {
        out += body.slice(pos, t.p) + body.slice(t.p, t.e).toUpperCase();
        pos = t.e;
      }
    }
    out += body.slice(pos);
    return '=' + out;
  }
  function inheritFormat(wb, sh, f) {
    const body = f.slice(1);
    const toks = FX.tokenize(body);
    const fns = toks.filter(t => t.k === 'fn').map(t => t.v);
    if (fns.some(n => n === 'DATE' || n === 'TODAY')) return 'yyyy-mm-dd';
    if (fns.length) return null;
    const refs = toks.filter(t => t.k === 'ref');
    const ops = toks.filter(t => t.k === 'op').map(t => t.v);
    if (refs.length && ops.length && ops.every(o => o === '+' || o === '-')) {
      for (const t of refs) {
        const s2 = t.sheet ? sheetOf(wb, t.sheet) : sh;
        const cl = s2 && getCell(s2, t.a.r, t.a.c);
        const nf = cl && cl.s && cl.s.nf;
        if (nf && NF.isDateFmt(nf)) return nf;
      }
    }
    return null;
  }
  function clearRange(wb, sName, rg, what) {
    const sh = sheetOf(wb, sName);
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      const cl = getCell(sh, r, c);
      if (!cl) continue;
      if (what === 'all') { delete sh.cells[key(r, c)]; continue; }
      if (what === 'format') { delete cl.s; cleanCell(sh, r, c); continue; }
      delete cl.v; delete cl.f; delete cl.q;
      cleanCell(sh, r, c);
    }
    touch(wb);
  }

  // ------------------------------------------------------------------ 서식
  function eachCell(rg, fn) { for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) fn(r, c); }
  function applyStyle(wb, sName, ranges, patch) {
    const sh = sheetOf(wb, sName);
    for (const rg of [].concat(ranges)) eachCell(rg, (r, c) => {
      const cl = ensureCell(sh, r, c);
      cl.s = Object.assign({}, cl.s, patch);
      for (const k of Object.keys(cl.s)) if (cl.s[k] === null) delete cl.s[k];
      cleanCell(sh, r, c);
    });
    touch(wb);
  }
  function styleOf(sh, r, c) { const cl = getCell(sh, r, c); return (cl && cl.s) || {}; }
  function setBorder(wb, sName, rg, kind) {
    const sh = sheetOf(wb, sName);
    eachCell(rg, (r, c) => {
      const cl = ensureCell(sh, r, c);
      const bd = Object.assign({}, (cl.s && cl.s.bd) || {});
      if (kind === 'none') { delete bd.t; delete bd.b; delete bd.l; delete bd.r; }
      if (kind === 'all') { bd.t = bd.b = bd.l = bd.r = 1; }
      if (kind === 'outside' || kind === 'thick') {
        const w = kind === 'thick' ? 2 : 1;
        if (r === rg.r1) bd.t = w; if (r === rg.r2) bd.b = w; if (c === rg.c1) bd.l = w; if (c === rg.c2) bd.r = w;
      }
      if (kind === 'bottom' && r === rg.r2) bd.b = 1;
      if (kind === 'top' && r === rg.r1) bd.t = 1;
      if (kind === 'left' && c === rg.c1) bd.l = 1;
      if (kind === 'right' && c === rg.c2) bd.r = 1;
      cl.s = Object.assign({}, cl.s, { bd });
      if (!Object.keys(bd).length) delete cl.s.bd;
      cleanCell(sh, r, c);
    });
    touch(wb);
  }
  /* 셀 테두리 실제 표시(이웃 셀의 반대쪽 테두리 포함) */
  function borderAt(sh, r, c) {
    const s = styleOf(sh, r, c).bd || {};
    const up = (styleOf(sh, r - 1, c).bd || {}).b, dn = (styleOf(sh, r + 1, c).bd || {}).t;
    const lf = (styleOf(sh, r, c - 1).bd || {}).r, rt = (styleOf(sh, r, c + 1).bd || {}).l;
    return { t: s.t || up || 0, b: s.b || dn || 0, l: s.l || lf || 0, r: s.r || rt || 0 };
  }

  // ------------------------------------------------------------------ 복사·채우기
  function copy(wb, sName, rg) {
    const sh = sheetOf(wb, sName);
    const cells = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      const row = [];
      for (let c = rg.c1; c <= rg.c2; c++) {
        const pc = pivotCellAt(wb, sh, r, c);
        if (pc) { row.push({ v: pc.v, s: pc.nf ? { nf: pc.nf } : undefined }); continue; }
        row.push(U.clone(getCell(sh, r, c)));
      }
      cells.push(row);
    }
    const merges = sh.merges.filter(m => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2).map(m => ({ r1: m.r1 - rg.r1, c1: m.c1 - rg.c1, r2: m.r2 - rg.r1, c2: m.c2 - rg.c1 }));
    const pivot = OA.pivot ? OA.pivot.coveredBy(wb, sh, rg) : null;
    return { sheet: sh.name, rg: U.clone(rg), cells, merges, pivot };
  }
  function paste(wb, sName, clip, r0, c0, opts) {
    const sh = sheetOf(wb, sName);
    if (clip.pivot && OA.pivot) { OA.pivot.pasteCopy(wb, sh, clip.pivot, r0, c0, clip.rg); touch(wb); return {}; }
    const R = clip.cells.length, C = clip.cells[0].length;
    const dr = r0 - clip.rg.r1, dc = c0 - clip.rg.c1;
    // 붙여 넣을 범위가 선택 범위의 배수면 반복(채우기와 같게)
    let reps = { r: 1, c: 1 };
    if (opts && opts.dest) { const d = opts.dest; const h = d.r2 - d.r1 + 1, w = d.c2 - d.c1 + 1; if (h % R === 0 && w % C === 0) reps = { r: h / R, c: w / C }; }
    for (let ir = 0; ir < reps.r; ir++) for (let ic = 0; ic < reps.c; ic++) {
      for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
        const src = clip.cells[i][j];
        const r = r0 + ir * R + i, c = c0 + ic * C + j;
        if (r > MAX_ROW || c > MAX_COL) continue;
        if (!src) { delete sh.cells[key(r, c)]; continue; }
        const n = U.clone(src);
        if (n.f) n.f = FX.shift(n.f, dr + ir * R, dc + ic * C);
        sh.cells[key(r, c)] = n;
      }
      for (const m of clip.merges) {
        const nm = { r1: r0 + ir * R + m.r1, c1: c0 + ic * C + m.c1, r2: r0 + ir * R + m.r2, c2: c0 + ic * C + m.c2 };
        sh.merges = sh.merges.filter(x => !(x.r1 <= nm.r2 && x.r2 >= nm.r1 && x.c1 <= nm.c2 && x.c2 >= nm.c1));
        sh.merges.push(nm);
      }
    }
    touch(wb);
    return {};
  }
  /* 채우기: src(원본 범위)를 dest(원본 포함 확장 범위)로. 아래·오른쪽·위·왼쪽 모두 */
  function fill(wb, sName, src, dest) {
    const sh = sheetOf(wb, sName);
    const R = src.r2 - src.r1 + 1, C = src.c2 - src.c1 + 1;
    for (let r = dest.r1; r <= dest.r2; r++) for (let c = dest.c1; c <= dest.c2; c++) {
      if (U.inRange(src, r, c)) continue;
      const sr = src.r1 + (((r - src.r1) % R) + R) % R, sc = src.c1 + (((c - src.c1) % C) + C) % C;
      const s = getCell(sh, sr, sc);
      if (!s) { delete sh.cells[key(r, c)]; continue; }
      const n = U.clone(s);
      if (n.f) n.f = FX.shift(n.f, r - sr, c - sc);
      sh.cells[key(r, c)] = n;
    }
    touch(wb);
  }

  // ------------------------------------------------------------------ 정렬
  function sortRange(wb, sName, rg, keys, header) {
    const sh = sheetOf(wb, sName);
    const top = header ? rg.r1 + 1 : rg.r1;
    const rows = [];
    for (let r = top; r <= rg.r2; r++) {
      const vals = keys.map(k => getValue(wb, sh, r, k.col));
      const cells = [];
      for (let c = rg.c1; c <= rg.c2; c++) cells.push(getCell(sh, r, c));
      rows.push({ r, vals, cells });
    }
    const cmpVal = (a, b) => {
      const ea = a === null || a === undefined || a === '', eb = b === null || b === undefined || b === '';
      if (ea && eb) return 0; if (ea) return 1; if (eb) return -1;      // 빈 셀은 항상 맨 뒤
      return FX.cmp(a, b);
    };
    rows.sort((x, y) => {
      for (let i = 0; i < keys.length; i++) {
        const ea = x.vals[i] === null || x.vals[i] === '', eb = y.vals[i] === null || y.vals[i] === '';
        let d = cmpVal(x.vals[i], y.vals[i]);
        if (keys[i].order === 'desc' && !ea && !eb) d = -d;
        if (d) return d;
      }
      return x.r - y.r;
    });
    rows.forEach((row, i) => {
      const nr = top + i;
      row.cells.forEach((cl, j) => {
        const c = rg.c1 + j;
        if (!cl) { delete sh.cells[key(nr, c)]; return; }
        const n = U.clone(cl);
        if (n.f) n.f = FX.shift(n.f, nr - row.r, 0);
        sh.cells[key(nr, c)] = n;
      });
    });
    touch(wb);
  }
  /* 현재 영역(Ctrl+A, 정렬 범위 자동 확장) */
  function currentRegion(wb, sName, r, c) {
    const sh = sheetOf(wb, sName);
    const filled = (rr, cc) => { const cl = getCell(sh, rr, cc); return !!(cl && (cl.f || (cl.v !== null && cl.v !== undefined && cl.v !== ''))) || !!mergeAt(sh, rr, cc); };
    let g = { r1: r, c1: c, r2: r, c2: c };
    if (!filled(r, c)) {
      let any = false;
      for (let dr = -1; dr <= 1; dr++) for (let dc = -1; dc <= 1; dc++) if (r + dr >= 1 && c + dc >= 1 && filled(r + dr, c + dc)) any = true;
      if (!any) return g;
    }
    let grow = true;
    while (grow) {
      grow = false;
      const test = (rr, cc) => rr >= 1 && cc >= 1 && rr <= MAX_ROW && cc <= MAX_COL && filled(rr, cc);
      for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) { if (test(g.r1 - 1, cc)) { g.r1--; grow = true; break; } }
      for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) { if (test(g.r2 + 1, cc)) { g.r2++; grow = true; break; } }
      for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) { if (test(rr, g.c1 - 1)) { g.c1--; grow = true; break; } }
      for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) { if (test(rr, g.c2 + 1)) { g.c2++; grow = true; break; } }
    }
    return g;
  }

  // ------------------------------------------------------------------ 시트
  function addSheet(wb, name, at) {
    let n = wb.sheets.length + 1;
    while (!name && wb.sheets.some(s => s.name === 'Sheet' + n)) n++;
    const sh = newSheet(name || 'Sheet' + n);
    const idx = at === undefined ? wb.active + 1 : at;
    wb.sheets.splice(idx, 0, sh);
    wb.active = idx;
    touch(wb);
    return sh;
  }
  function renameSheet(wb, idx, name) {
    name = String(name || '').trim();
    if (!name) return { error: '시트 이름을 입력하십시오.' };
    if (/[\\/?*[\]:]/.test(name) || name.length > 31) return { error: '시트 이름에 사용할 수 없는 문자가 있거나 너무 깁니다.' };
    if (wb.sheets.some((s, i) => i !== idx && s.name === name)) return { error: '이미 사용 중인 이름입니다. 다른 이름을 입력하십시오.' };
    const old = wb.sheets[idx].name;
    wb.sheets[idx].name = name;
    // 다른 시트 수식·차트·피벗 원본의 시트 이름 바꾸기
    const rep = s => s.split(old + '!').join(name + '!').split("'" + old + "'!").join(name + '!');
    for (const sh of wb.sheets) {
      for (const k of Object.keys(sh.cells)) { const cl = sh.cells[k]; if (cl.f) cl.f = rep(cl.f); }
      for (const ch of sh.charts) for (const s of ch.series) for (const p of ['name', 'val', 'cat']) if (s[p]) s[p] = rep(s[p]);
    }
    for (const c of wb.caches) if (c.sheet === old) c.sheet = name;
    touch(wb);
    return {};
  }
  function deleteSheet(wb, idx) {
    if (wb.sheets.length <= 1) return { error: '통합 문서에는 시트가 하나 이상 있어야 합니다.' };
    wb.sheets.splice(idx, 1);
    wb.active = Math.max(0, Math.min(wb.active, wb.sheets.length - 1));
    touch(wb);
    return {};
  }
  const colW = (sh, c) => sh.colW[c] || DEF_COL;
  const rowH = (sh, r) => sh.rowH[r] || DEF_ROW;
  function colLeft(sh, c) { let x = 0; for (let k = 1; k < c; k++) x += colW(sh, k); return x; }
  function rowTop(sh, r) { let y = 0; for (let k = 1; k < r; k++) y += rowH(sh, k); return y; }
  function cellAtXY(sh, x, y) {
    let c = 1, acc = 0; while (c < MAX_COL && acc + colW(sh, c) <= x) { acc += colW(sh, c); c++; }
    let r = 1; acc = 0; while (r < MAX_ROW && acc + rowH(sh, r) <= y) { acc += rowH(sh, r); r++; }
    return { r, c };
  }

  function usedRange(wb, sh) {
    let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
    for (const k of Object.keys(sh.cells)) {
      const cl = sh.cells[k];
      if (!(cl.f || (cl.v !== null && cl.v !== undefined && cl.v !== '') || (cl.s && cl.s.bd))) continue;
      const [r, c] = k.split(',').map(Number);
      r1 = Math.min(r1, r); c1 = Math.min(c1, c); r2 = Math.max(r2, r); c2 = Math.max(c2, c);
    }
    for (const m of sh.merges) { r1 = Math.min(r1, m.r1); c1 = Math.min(c1, m.c1); r2 = Math.max(r2, m.r2); c2 = Math.max(c2, m.c2); }
    if (OA.pivot) for (const pv of sh.pivots) { const g = OA.pivot.grid(wb, sh, pv); if (g) { r1 = Math.min(r1, g.r0); c1 = Math.min(c1, g.c0); r2 = Math.max(r2, g.r0 + g.rows.length - 1); c2 = Math.max(c2, g.c0 + g.ncol - 1); } }
    for (const ch of sh.charts) {
      const a = cellAtXY(sh, ch.x, ch.y), b = cellAtXY(sh, ch.x + ch.w - 1, ch.y + ch.h - 1);
      r1 = Math.min(r1, a.r); c1 = Math.min(c1, a.c); r2 = Math.max(r2, b.r); c2 = Math.max(c2, b.c);
    }
    if (!isFinite(r1)) return null;
    return { r1, c1, r2, c2 };
  }

  // ------------------------------------------------------------------ 인쇄 배치
  function printLayout(wb, sName) {
    const sh = sheetOf(wb, sName);
    const pg = sh.page;
    const area = sh.printArea ? U.parseRange(sh.printArea) : usedRange(wb, sh);
    const paper = { w: 210, h: 297 };
    const P = pg.orient === 'landscape' ? { w: paper.h, h: paper.w } : paper;
    const avail = { w: P.w - (pg.left + pg.right) * 10, h: P.h - (pg.top + pg.bottom) * 10 };
    if (!area) return { sh, area: null, pages: [{ empty: true }], scale: 1, paper: P, avail, pg };
    const cw = [], rh = [];
    for (let c = area.c1; c <= area.c2; c++) cw.push(colW(sh, c) * PX_MM);
    for (let r = area.r1; r <= area.r2; r++) rh.push(rowH(sh, r) * PX_MM);
    const W = cw.reduce((s, x) => s + x, 0), H = rh.reduce((s, x) => s + x, 0);
    let scale = pg.scaleMode === 'fit' ? 1 : (pg.zoom || 100) / 100;
    if (pg.scaleMode === 'fit') {                      // 자동 맞춤은 줄이기만 한다(최소 10%)
      const sw = pg.fitW ? avail.w * pg.fitW / W : Infinity;
      const shh = pg.fitH ? avail.h * pg.fitH / H : Infinity;
      scale = Math.max(0.1, Math.min(1, sw, shh));
    }
    const split = (sizes, lim) => {
      const out = []; let start = 0, acc = 0;
      sizes.forEach((s, i) => { if (acc + s * scale > lim + 0.01 && i > start) { out.push([start, i - 1]); start = i; acc = 0; } acc += s * scale; });
      out.push([start, sizes.length - 1]);
      return out;
    };
    const cs = split(cw, avail.w), rs = split(rh, avail.h);
    const pages = [];
    for (const [a, b] of cs) for (const [p, q] of rs) {
      pages.push({ c1: area.c1 + a, c2: area.c1 + b, r1: area.r1 + p, r2: area.r1 + q,
        w: cw.slice(a, b + 1).reduce((s, x) => s + x, 0) * scale, h: rh.slice(p, q + 1).reduce((s, x) => s + x, 0) * scale });
    }
    // Excel 순서: 아래로 먼저
    pages.sort((x, y) => x.c1 - y.c1 || x.r1 - y.r1);
    return { sh, area, pages, scale, paper: P, avail, pg, W: W * scale, H: H * scale, PX_MM };
  }

  // ------------------------------------------------------------------ 차트 데이터
  function refValues(wb, ref) {
    const g = U.parseRange(ref);
    if (!g) return [];
    const out = [];
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) out.push(getValue(wb, g.sheet, r, c));
    return out;
  }
  function refFormat(wb, ref) {
    const g = U.parseRange(ref);
    if (!g) return NF.GENERAL;
    const sh = sheetOf(wb, g.sheet);
    return sh ? cellFormat(wb, sh, g.r1, g.c1) : NF.GENERAL;
  }
  /* 선택 영역(여러 범위) → 계열 정의. 첫 열이 문자면 항목, 첫 행이 문자면 계열 이름 */
  function seriesFromSelection(wb, sName, ranges) {
    const sh = sheetOf(wb, sName);
    const rr = ranges.map(g => U.clone(g)).sort((a, b) => a.c1 - b.c1);
    const r1 = Math.min(...rr.map(g => g.r1)), r2 = Math.max(...rr.map(g => g.r2));
    const cols = [];
    for (const g of rr) for (let c = g.c1; c <= g.c2; c++) cols.push(c);
    const isText = (r, c) => typeof getValue(wb, sh, r, c) === 'string';
    const isNumOrBlank = (r, c) => { const v = getValue(wb, sh, r, c); return v === null || typeof v === 'number'; };
    const hasHeader = cols.some(c => isText(r1, c)) && cols.slice(cols.length > 1 ? 1 : 0).every(c => { for (let r = r1 + 1; r <= r2; r++) if (!isNumOrBlank(r, c)) return false; return true; });
    const dataTop = hasHeader ? r1 + 1 : r1;
    let catCol = null;
    if (cols.length > 1) { let txt = 0; for (let r = dataTop; r <= r2; r++) if (isText(r, cols[0])) txt++; if (txt > 0) catCol = cols[0]; }
    const valCols = catCol === null ? cols : cols.slice(1);
    const S = (r, c, r2b, c2b) => U.rangeStr({ r1: r, c1: c, r2: r2b, c2: c2b }, true, sh.name);
    return valCols.map(c => ({
      name: hasHeader ? S(r1, c, r1, c) : null,
      val: S(dataTop, c, r2, c),
      cat: catCol !== null ? S(dataTop, catCol, r2, catCol) : null,
      type: 'col', secondary: false, labels: null,
    }));
  }
  function seriesName(wb, s, i) {
    if (!s.name) return '계열' + (i + 1);
    if (/^[^!]+![$A-Z0-9:]+$/i.test(s.name)) { const v = refValues(wb, s.name)[0]; return v === null || v === undefined ? '' : FX.toStr(v); }
    return s.name;
  }
  function newChart(wb, sName, series, kind, at) {
    const sh = sheetOf(wb, sName);
    const id = 'chart' + (wb.seq.chart++);
    const ch = { id, name: '차트 ' + (sh.charts.length + 1), x: at ? at.x : 300, y: at ? at.y : 40, w: 480, h: 288, kind: kind || 'col',
      series: series.map(s => Object.assign({}, s, { type: s.type || 'col' })),
      title: { show: true, text: null }, axisTitle: { x: { show: false, text: '축 제목' }, y: { show: false, text: '축 제목' } },
      legend: { show: true, pos: 'bottom' }, gridlines: true, axes: true };
    if (ch.series.length === 1 && ch.title.text === null) ch.title.text = null;
    sh.charts.push(ch);
    touch(wb);
    return ch;
  }
  function chartTitleText(wb, ch) {
    if (ch.title.text !== null && ch.title.text !== undefined) return ch.title.text;
    return ch.series.length === 1 ? seriesName(wb, ch.series[0], 0) : '차트 제목';
  }
  function seriesFormula(s, i) { return `=SERIES(${s.name || ''},${s.cat || ''},${s.val},${i + 1})`; }

  OA.xl = { DEF_COL, DEF_ROW, MAX_ROW, MAX_COL, PX_MM, DEFAULT_PAGE, newWorkbook, newSheet, fromJSON, attach, touch, sheetOf, getCell, ensureCell,
    getValue, display, editText, cellFormat, merge, mergeAt, lostByMerge, setInput, clearRange, applyStyle, styleOf, setBorder, borderAt,
    copy, paste, fill, sortRange, currentRegion, addSheet, renameSheet, deleteSheet, colW, rowH, colLeft, rowTop, cellAtXY, usedRange,
    printLayout, refValues, refFormat, seriesFromSelection, seriesName, newChart, chartTitleText, seriesFormula, pivotCellAt };
})(typeof window !== 'undefined' ? window : globalThis);

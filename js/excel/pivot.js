/* 피벗테이블 엔진(DOM 없음): 캐시·필드 배치·선택 항목 그룹·숫자 구간 그룹·압축/개요/테이블 형식·부분합·총합계·확장/축소·이름 바꾸기·복사. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util, NF = OA.numfmt, FX = OA.formula, XL = () => OA.xl;

  const FUNC = { sum: '합계', count: '개수', avg: '평균', max: '최대', min: '최소', product: '곱' };
  const FUNC_LIST = ['sum', 'count', 'avg', 'max', 'min', 'product'];

  const rawKey = v => v === null || v === undefined || v === '' ? 'b:' : typeof v === 'number' ? 'n:' + v : typeof v === 'boolean' ? 'l:' + v : 's:' + v;
  const keyVal = k => k.startsWith('n:') ? Number(k.slice(2)) : k.startsWith('s:') ? k.slice(2) : k.startsWith('l:') ? k.slice(2) === 'true' : null;

  function cacheOf(wb, pv) { return wb.caches.find(c => c.id === pv.cacheId); }

  // ------------------------------------------------------------------ 캐시
  function readSource(wb, ref) {
    const g = U.parseRange(ref);
    if (!g) return { error: '참조가 유효하지 않습니다.' };
    const sh = XL().sheetOf(wb, g.sheet);
    if (!sh) return { error: '참조가 유효하지 않습니다.' };
    if (g.r2 - g.r1 < 1) return { error: '피벗 테이블을 만들려면 두 행 이상의 데이터 범위를 선택하십시오.' };
    const fields = [], nfs = [];
    for (let c = g.c1; c <= g.c2; c++) {
      const v = XL().getValue(wb, sh, g.r1, c);
      if (v === null || v === '' || typeof v === 'object') return { error: '피벗 테이블 필드 이름이 잘못되었습니다. 피벗 테이블 보고서를 만들려면 레이블이 지정된 열이 있는 목록으로 구성된 데이터를 사용해야 합니다.' };
      fields.push(FX.toStr(v));
      nfs.push(XL().cellFormat(wb, sh, g.r1 + 1, c));
    }
    const records = [];
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      const row = [];
      for (let c = g.c1; c <= g.c2; c++) { const v = XL().getValue(wb, sh, r, c); row.push(FX.isErr(v) ? null : v); }
      records.push(row);
    }
    // 같은 이름 필드는 뒤에 숫자를 붙인다
    const seen = {};
    fields.forEach((f, i) => { if (seen[f]) { let n = 2; while (fields.includes(f + n)) n++; fields[i] = f + n; } seen[f] = 1; });
    return { sheet: sh.name, ref: U.rangeStr(g, true), fields, nfs, records };
  }
  function createCache(wb, ref) {
    const src = readSource(wb, ref);
    if (src.error) return src;
    const c = Object.assign({ id: 'cache' + (wb.seq.cache++), groups: [], ranges: {} }, src);
    wb.caches.push(c);
    return c;
  }
  function refresh(wb, pv) {
    const c = cacheOf(wb, pv);
    const src = readSource(wb, c.sheet + '!' + c.ref);
    if (src.error) return src;
    Object.assign(c, { fields: src.fields.concat(c.groups.map(g => g.field)).filter((f, i, a) => a.indexOf(f) === i && (src.fields.includes(f) || c.groups.some(g => g.field === f))), nfs: src.nfs, records: src.records });
    XL().touch(wb);
    return {};
  }
  function allFields(c) { return c.fields.slice(0, c.records.length ? c.records[0].length : c.fields.length).concat(c.groups.map(g => g.field)); }
  function baseFields(c) { return c.fields.slice(0, c.nfs.length); }
  function isNumericField(c, f) {
    const i = baseFields(c).indexOf(f);
    if (i < 0) return false;
    let any = false;
    for (const r of c.records) { const v = r[i]; if (v === null) continue; if (typeof v !== 'number') return false; any = true; }
    return any;
  }

  // ------------------------------------------------------------------ 피벗 만들기
  function create(wb, ref, dest) {
    const c = createCache(wb, ref);
    if (c.error) return c;
    let sh, r, col;
    if (dest.newSheet) {
      sh = XL().addSheet(wb, null, wb.active);
      r = 3; col = 1;
    } else {
      const g = U.parseRange(dest.ref);
      if (!g) return { error: '대상 참조가 유효하지 않습니다.' };
      sh = XL().sheetOf(wb, g.sheet || dest.sheet);
      if (!sh) return { error: '대상 참조가 유효하지 않습니다.' };
      r = g.r1; col = g.c1;
    }
    let n = 1;
    const names = wb.sheets.flatMap(s => s.pivots.map(p => p.name));
    while (names.includes('피벗 테이블' + n)) n++;
    const pv = { id: 'pv' + (wb.seq.pivot++), name: '피벗 테이블' + n, cacheId: c.id, r, c: col, rows: [], cols: [], filters: [], values: [],
      layout: 'compact', subtotals: 'top', grandRow: true, grandCol: true, collapsed: {}, itemNames: {}, captions: {},
      showButtons: true, showHeaders: true, printButtons: false, repeat: false };
    sh.pivots.push(pv);
    wb.active = wb.sheets.indexOf(sh);
    XL().touch(wb);
    return { pv, sheet: sh };
  }
  function find(wb, id) {
    for (const sh of wb.sheets) for (const pv of sh.pivots) if (pv.id === id) return { sh, pv };
    return null;
  }

  // ------------------------------------------------------------------ 필드 배치
  function valueName(func, field, pv) {
    let base = `${FUNC[func]} : ${field}`, name = base, n = 2;
    while (pv && pv.values.some(v => v.name === name)) name = base + n++;
    return name;
  }
  function placed(pv, f) { return pv.rows.includes(f) ? 'rows' : pv.cols.includes(f) ? 'cols' : pv.filters.includes(f) ? 'filters' : pv.values.some(v => v.field === f) ? 'values' : null; }
  function addField(wb, pv, field, area, pos) {
    const c = cacheOf(wb, pv);
    if (!area) area = isNumericField(c, field) ? 'values' : 'rows';
    if (area !== 'values') removeField(wb, pv, field, null, true);
    if (area === 'values') {
      const v = { field, func: isNumericField(c, field) ? 'sum' : 'count', nf: null };
      v.name = valueName(v.func, field, pv);
      if (pos === undefined) pv.values.push(v); else pv.values.splice(pos, 0, v);
    } else {
      const list = pv[area];
      if (pos === undefined) list.push(field); else list.splice(pos, 0, field);
    }
    syncValuesField(pv);
    XL().touch(wb);
  }
  function removeField(wb, pv, field, area, keepValues) {
    for (const a of ['rows', 'cols', 'filters']) if ((!area || area === a)) pv[a] = pv[a].filter(f => f !== field);
    if (!keepValues && (!area || area === 'values')) pv.values = pv.values.filter(v => v.field !== field);
    syncValuesField(pv);
    XL().touch(wb);
  }
  function removeValue(wb, pv, idx) { pv.values.splice(idx, 1); syncValuesField(pv); XL().touch(wb); }
  function moveInArea(wb, pv, area, from, to) {
    const list = area === 'values' ? pv.values : pv[area];
    const [x] = list.splice(from, 1);
    list.splice(Math.max(0, Math.min(list.length, to)), 0, x);
    XL().touch(wb);
  }
  function syncValuesField(pv) {
    // 값 필드가 2개 이상이면 열 영역에 Σ 값
    const has = pv.cols.includes('Σ값') || pv.rows.includes('Σ값');
    if (pv.values.length >= 2 && !has) pv.cols.push('Σ값');
    if (pv.values.length < 2 && has) { pv.cols = pv.cols.filter(f => f !== 'Σ값'); pv.rows = pv.rows.filter(f => f !== 'Σ값'); }
  }
  function setValueField(wb, pv, idx, patch) {
    const v = pv.values[idx];
    if (patch.name !== undefined) {
      const nm = String(patch.name).trim();
      const c = cacheOf(wb, pv);
      if (!nm) return { error: '이름을 입력하십시오.' };
      if (allFields(c).includes(nm) || pv.values.some((x, i) => i !== idx && x.name === nm)) return { error: '피벗 테이블 필드 이름이 이미 있습니다.' };
    }
    const oldAuto = v.name === valueName(v.func, v.field, null) || /^(합계|개수|평균|최대|최소|곱) : /.test(v.name);
    Object.assign(v, patch);
    if (patch.func && patch.name === undefined && oldAuto) v.name = valueName(v.func, v.field, Object.assign({}, pv, { values: pv.values.filter((x, i) => i !== idx) }));
    XL().touch(wb);
    return {};
  }

  // ------------------------------------------------------------------ 항목
  function groupDef(c, f) { return c.groups.find(g => g.field === f) || null; }
  function groupFor(c, base) { return c.groups.find(g => g.base === base) || null; }
  function itemKeyOf(c, f, rec) {
    const gd = groupDef(c, f);
    if (gd) {
      const bi = c.fields.indexOf(gd.base);
      const bk = itemKeyOf(c, gd.base, rec);
      const grp = gd.items.find(it => it.members.includes(bk));
      return grp ? 'g:' + grp.id : bk;
    }
    const i = c.fields.indexOf(f);
    const v = rec[i];
    const rg = c.ranges && c.ranges[f];
    if (rg && typeof v === 'number') {
      if (v < rg.start) return 'r:<' + rg.start;
      if (v > rg.end) return 'r:>' + rg.end;
      const k = Math.floor((v - rg.start) / rg.by);
      const a = rg.start + k * rg.by, b = a + rg.by - 1;
      return 'r:' + a + '-' + b;
    }
    return rawKey(v);
  }
  function baseOrder(c, f) {
    const i = c.fields.indexOf(f);
    const vals = [];
    const seen = new Set();
    for (const r of c.records) { const k = itemKeyOf(c, f, r); if (!seen.has(k)) { seen.add(k); vals.push(k); } }
    return sortKeys(c, f, vals);
  }
  function sortKeys(c, f, keys) {
    const gd = groupDef(c, f);
    if (gd) {
      const bo = baseOrder(c, gd.base);
      const pos = k => k.startsWith('g:') ? Math.min(...gd.items.find(it => 'g:' + it.id === k).members.map(m => bo.indexOf(m)).filter(x => x >= 0)) : bo.indexOf(k);
      return keys.slice().sort((a, b) => pos(a) - pos(b));
    }
    const rank = k => k.startsWith('b:') ? 3 : k.startsWith('n:') || k.startsWith('r:') ? 0 : k.startsWith('l:') ? 2 : 1;
    const numOf = k => k.startsWith('r:') ? (k[2] === '<' ? -Infinity : k[2] === '>' ? Infinity : parseFloat(k.slice(2))) : Number(k.slice(2));
    return keys.slice().sort((a, b) => {
      const ra = rank(a), rb = rank(b);
      if (ra !== rb) return ra - rb;
      if (ra === 0) return numOf(a) - numOf(b);
      return U.koCompare(a.slice(2), b.slice(2));
    });
  }
  function fieldNf(c, f) {
    const gd = groupDef(c, f);
    if (gd) return null;
    const i = c.fields.indexOf(f);
    return c.nfs[i] || NF.GENERAL;
  }
  function itemLabel(c, pv, f, k) {
    const custom = pv.itemNames[f] && pv.itemNames[f][k];
    if (custom !== undefined) return custom;
    if (k.startsWith('g:')) { const gd = groupDef(c, f); const it = gd && gd.items.find(x => 'g:' + x.id === k); return it ? it.name : k; }
    if (k.startsWith('r:')) return k.slice(2);
    if (k === 'b:') return '(비어 있음)';
    const v = keyVal(k);
    if (groupDef(c, f)) return FX.toStr(v);                    // 그룹 필드에 남은 원래 항목: 일반 표시(0, 0.05)
    if (typeof v === 'number') return NF.format(v, fieldNf(c, f)).text;
    return FX.toStr(v);
  }
  function caption(c, pv, f) { return (pv.captions && pv.captions[f]) || f; }

  // ------------------------------------------------------------------ 집계
  function aggregate(c, recs, vf) {
    const i = c.fields.indexOf(vf.field);
    const xs = recs.map(r => r[i]);
    const nums = xs.filter(x => typeof x === 'number');
    switch (vf.func) {
      case 'sum': return nums.reduce((s, x) => s + x, 0);
      case 'count': return xs.filter(x => x !== null && x !== '').length;
      case 'avg': return nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : FX.E.DIV0;
      case 'max': return nums.length ? Math.max(...nums) : 0;
      case 'min': return nums.length ? Math.min(...nums) : 0;
      case 'product': return nums.reduce((s, x) => s * x, 1);
    }
    return 0;
  }
  function buildTree(c, pv, fields, recs, level) {
    if (level >= fields.length) return [];
    const f = fields[level];
    const groups = new Map();
    for (const r of recs) { const k = itemKeyOf(c, f, r); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); }
    return sortKeys(c, f, [...groups.keys()]).map(k => ({ field: f, key: k, recs: groups.get(k), children: buildTree(c, pv, fields, groups.get(k), level + 1) }));
  }
  function isCollapsed(pv, f, k) { return !!(pv.collapsed[f] && pv.collapsed[f].includes(k)); }

  // ------------------------------------------------------------------ 격자
  function grid(wb, sh, pv) {
    const ck = pv.id;
    if (wb.__pv.has(ck)) return wb.__pv.get(ck);
    const g = computeGrid(wb, sh, pv);
    wb.__pv.set(ck, g);
    return g;
  }
  function computeGrid(wb, sh, pv) {
    const c = cacheOf(wb, pv);
    if (!c) return null;
    const rows = pv.rows.filter(f => f !== 'Σ값');
    const colFields = pv.cols.filter(f => f !== 'Σ값');
    const out = [];
    let r0 = pv.r;
    const filterRows = [];
    for (const f of pv.filters) filterRows.push([{ v: caption(c, pv, f), kind: 'filterhdr', field: f, bold: true }, { v: '(모두)', kind: 'filterval', field: f, dd: true }]);
    if (filterRows.length) filterRows.push([]);
    const vals = pv.values;
    if (!rows.length && !colFields.length && !vals.length) {
      const empty = [];
      for (let i = 0; i < 18; i++) empty.push([{ v: null, kind: 'empty' }, { v: null, kind: 'empty' }, { v: null, kind: 'empty' }]);
      empty[1][0] = { v: null, text: pv.name, kind: 'empty', placeholder: true };
      empty[3][0] = { v: null, text: '보고서를 작성하려면 피벗 테이블 필드 목록에서 필드를 선택하십시오.', kind: 'empty', placeholder: true };
      return { r0: pv.r, c0: pv.c, ncol: 3, rows: empty, empty: true, pv };
    }
    const tabular = pv.layout === 'tabular', outline = pv.layout === 'outline', compact = pv.layout === 'compact';
    const sub = pv.subtotals;
    // 열 항목
    let colItems = [null];
    if (colFields.length) {
      const f = colFields[0];
      const keys = sortKeys(c, f, [...new Set(c.records.map(r => itemKeyOf(c, f, r)))]);
      colItems = keys;
    }
    const valCols = [];
    for (const ci of colItems) for (let vi = 0; vi < Math.max(1, vals.length); vi++) valCols.push({ ci, vi });
    if (colFields.length && pv.grandCol) for (let vi = 0; vi < Math.max(1, vals.length); vi++) valCols.push({ ci: '__grand', vi });
    const nLabelCols = compact ? (rows.length ? 1 : 0) : rows.length;
    const ncol = Math.max(1, nLabelCols) + valCols.length;
    const valOf = (recs, col) => {
      if (!vals.length) return null;
      let rs = recs;
      if (col.ci !== null && col.ci !== '__grand') rs = recs.filter(r => itemKeyOf(c, colFields[0], r) === col.ci);
      if (!rs.length && col.ci !== null && col.ci !== '__grand') return null;
      return aggregate(c, rs, vals[col.vi]);
    };
    const nfOf = col => vals.length ? (vals[col.vi].nf || NF.GENERAL) : NF.GENERAL;
    const hdrFill = true;
    // 머리글
    if (colFields.length) {
      const h1 = [];
      for (let i = 0; i < Math.max(1, nLabelCols); i++) h1.push({ v: null, kind: 'blank', hdr: true });
      h1.push({ v: compact ? '열 레이블' : caption(c, pv, colFields[0]), kind: 'fieldhdr', field: colFields[0], dd: pv.showHeaders, hdr: true, bold: true });
      while (h1.length < ncol) h1.push({ v: null, kind: 'blank', hdr: true });
      out.push(h1);
    }
    const hdr = [];
    if (compact) hdr.push(rows.length ? { v: pv.showHeaders ? '행 레이블' : null, kind: 'fieldhdr', field: rows[0] || null, dd: pv.showHeaders, bold: true, hdr: true } : { v: null, kind: 'blank', hdr: true });
    else if (rows.length) rows.forEach(f => hdr.push({ v: pv.showHeaders ? caption(c, pv, f) : null, kind: 'fieldhdr', field: f, dd: pv.showHeaders, bold: true, hdr: true }));
    else hdr.push({ v: null, kind: 'blank', hdr: true });
    for (const col of valCols) {
      let text;
      if (colFields.length) text = col.ci === '__grand' ? (vals.length > 1 ? `전체 ${vals[col.vi].name}` : '총합계') : itemLabel(c, pv, colFields[0], col.ci) + (vals.length > 1 ? ' ' + vals[col.vi].name : '');
      else text = vals.length ? vals[col.vi].name : null;
      hdr.push({ v: text, kind: colFields.length ? 'colitem' : 'valhdr', vi: col.vi, field: colFields[0], key: col.ci, bold: true, hdr: true, align: 'left' });
    }
    out.push(hdr);
    // 본문
    const tree = buildTree(c, pv, rows, c.records, 0);
    const mkValues = (recs, bold) => valCols.map(col => ({ v: valOf(recs, col), nf: nfOf(col), kind: 'value', bold }));
    const labelCell = (node, level, showBtn) => ({
      v: null, text: itemLabel(c, pv, node.field, node.key), kind: 'label', field: node.field, key: node.key, level,
      btn: showBtn && pv.showButtons && level < rows.length - 1 ? (isCollapsed(pv, node.field, node.key) ? '+' : '-') : null, align: 'left',
      indent: compact ? level : 0, bold: level < rows.length - 1 && (compact || outline || tabular) && (compact || outline),
    });
    // 테이블 형식: 바깥 필드 레이블은 그 묶음의 첫 행에만(항목 레이블 반복이면 모든 행)
    const cur = [], pend = [];
    const blank = () => ({ v: null, kind: 'blank' });
    function rowPrefix(level) {
      const row = [];
      for (let i = 0; i <= level; i++) { row.push(pend[i] || pv.repeat ? Object.assign({}, cur[i]) : blank()); pend[i] = false; }
      return row;
    }
    function emitTabular(nodes, level) {
      for (const node of nodes) {
        cur[level] = labelCell(node, level, true);
        pend[level] = true;
        const leafish = level === rows.length - 1 || isCollapsed(pv, node.field, node.key);
        if (leafish) {
          const row = rowPrefix(level);
          while (row.length < rows.length) row.push(blank());
          out.push(row.concat(mkValues(node.recs, false)));
          continue;
        }
        emitTabular(node.children, level + 1);
        if (sub !== 'none') {
          const row = [];
          for (let i = 0; i < level; i++) row.push(pv.repeat ? Object.assign({}, cur[i]) : blank());
          row.push({ v: null, text: itemLabel(c, pv, node.field, node.key) + ' 요약', kind: 'subtotal', field: node.field, key: node.key, bold: true, align: 'left' });
          while (row.length < rows.length) row.push({ v: null, kind: 'blank', bold: true });
          out.push(row.concat(mkValues(node.recs, true)));
        }
      }
    }
    function emitCompact(nodes, level) {
      for (const node of nodes) {
        const lab = labelCell(node, level, true);
        const leafish = level === rows.length - 1 || isCollapsed(pv, node.field, node.key);
        if (outline) {
          const row = [];
          for (let i = 0; i < level; i++) row.push({ v: null, kind: 'blank' });
          row.push(lab);
          while (row.length < rows.length) row.push({ v: null, kind: 'blank' });
          if (leafish) { out.push(row.concat(mkValues(node.recs, false))); continue; }
          out.push(row.concat(sub === 'top' ? mkValues(node.recs, true) : valCols.map(() => ({ v: null, kind: 'blank' }))));
        } else {
          if (leafish) { out.push([lab].concat(mkValues(node.recs, level < rows.length - 1))); continue; }
          out.push([lab].concat(sub === 'top' ? mkValues(node.recs, true) : valCols.map(() => ({ v: null, kind: 'blank' }))));
        }
        emitCompact(node.children, level + 1);
        if (sub === 'bottom') {
          const row = [];
          if (outline) { for (let i = 0; i < level; i++) row.push({ v: null, kind: 'blank' }); }
          row.push({ v: null, text: itemLabel(c, pv, node.field, node.key) + ' 요약', kind: 'subtotal', bold: true, indent: compact ? level : 0, align: 'left' });
          while (row.length < (outline ? rows.length : 1)) row.push({ v: null, kind: 'blank' });
          out.push(row.concat(mkValues(node.recs, true)));
        }
      }
    }
    if (rows.length) { if (tabular) emitTabular(tree, 0); else emitCompact(tree, 0); }
    // 총합계
    if (rows.length && pv.grandRow && vals.length) {
      const row = [{ v: null, text: '총합계', kind: 'grand', bold: true, align: 'left' }];
      while (row.length < Math.max(1, nLabelCols)) row.push({ v: null, kind: 'grand', bold: true });
      out.push(row.concat(valCols.map(col => ({ v: valOf(c.records, col), nf: nfOf(col), kind: 'value', bold: true, grand: true }))));
    } else if (!rows.length && vals.length) {
      const row = [{ v: null, kind: 'blank' }].concat(valCols.map(col => ({ v: valOf(c.records, col), nf: nfOf(col), kind: 'value' })));
      out.push(row);
    }
    for (const row of out) while (row.length < ncol) row.push({ v: null, kind: 'blank' });
    const all = filterRows.map(fr => { while (fr.length < ncol) fr.push({ v: null, kind: 'blank' }); return fr; }).concat(out);
    return { r0, c0: pv.c, ncol, rows: all, hdrRow: filterRows.length + (colFields.length ? 1 : 0), pv };
  }
  function tableRange(wb, sh, pv) { const g = grid(wb, sh, pv); return g ? { r1: g.r0, c1: g.c0, r2: g.r0 + g.rows.length - 1, c2: g.c0 + g.ncol - 1 } : null; }

  function hitTest(wb, sh, r, c) {
    for (const pv of sh.pivots) {
      const g = grid(wb, sh, pv);
      if (!g) continue;
      if (r >= g.r0 && r < g.r0 + g.rows.length && c >= g.c0 && c < g.c0 + g.ncol) return { pv, cell: g.rows[r - g.r0][c - g.c0] || { kind: 'blank' }, r, c };
    }
    return null;
  }
  function typeInto(wb, sh, hit, text) {
    const { pv, cell } = hit;
    const cc = cacheOf(wb, pv);
    const t = String(text == null ? '' : text).trim();
    if (cell.kind === 'label' || cell.kind === 'colitem' && cell.key) {
      if (!t) return { error: '피벗 테이블의 이 부분은 변경할 수 없습니다.' };
      const f = cell.field;
      const others = baseOrderForField(cc, pv, f).filter(k => k !== cell.key).map(k => itemLabel(cc, pv, f, k));
      if (others.includes(t)) return { error: '피벗 테이블 필드 이름이 이미 있습니다.' };
      pv.itemNames[f] = pv.itemNames[f] || {};
      pv.itemNames[f][cell.key] = t;
      XL().touch(wb); return {};
    }
    if (cell.kind === 'fieldhdr' && cell.field && pv.layout !== 'compact') {
      if (!t) return { error: '피벗 테이블의 이 부분은 변경할 수 없습니다.' };
      const used = allFields(cc).filter(f => f !== cell.field).map(f => caption(cc, pv, f));
      if (used.includes(t) || allFields(cc).filter(f => f !== cell.field).includes(t)) return { error: '피벗 테이블 필드 이름이 이미 있습니다.' };
      pv.captions[cell.field] = t;
      XL().touch(wb); return {};
    }
    if (cell.kind === 'valhdr') return setValueField(wb, pv, cell.vi, { name: t });
    return { error: '피벗 테이블의 이 부분은 변경할 수 없습니다. 이 부분을 변경하려면 피벗 테이블 필드 목록을 사용하십시오.' };
  }
  function baseOrderForField(c, pv, f) { return [...new Set(c.records.map(r => itemKeyOf(c, f, r)))]; }

  // ------------------------------------------------------------------ 확장·축소
  function toggle(wb, pv, field, key) {
    const list = pv.collapsed[field] = pv.collapsed[field] || [];
    const i = list.indexOf(key);
    if (i >= 0) list.splice(i, 1); else list.push(key);
    XL().touch(wb);
  }
  function setCollapsed(wb, pv, field, key, on) {
    const list = pv.collapsed[field] = pv.collapsed[field] || [];
    const i = list.indexOf(key);
    if (on && i < 0) list.push(key);
    if (!on && i >= 0) list.splice(i, 1);
    XL().touch(wb);
  }
  function collapseField(wb, pv, field, on) {
    const c = cacheOf(wb, pv);
    pv.collapsed[field] = on ? baseOrderForField(c, pv, field) : [];
    XL().touch(wb);
  }

  // ------------------------------------------------------------------ 그룹
  function group(wb, sh, pv, field, keys) {
    const c = cacheOf(wb, pv);
    if (!keys.length) return { error: '선택 항목을 그룹으로 묶을 수 없습니다.' };
    if (groupDef(c, field)) return { error: '선택 항목을 그룹으로 묶을 수 없습니다.' };
    if (keys.some(k => k.startsWith('r:'))) return { error: '선택 항목을 그룹으로 묶을 수 없습니다.' };
    let gd = groupFor(c, field);
    if (!gd) {
      let name = field + '2', n = 3;
      while (allFields(c).includes(name)) name = field + n++;
      gd = { field: name, base: field, items: [] };
      c.groups.push(gd);
      // 이 피벗의 행 영역에서 기준 필드 바로 앞에
      for (const area of ['rows', 'cols']) {
        const i = pv[area].indexOf(field);
        if (i >= 0) pv[area].splice(i, 0, name);
      }
    }
    for (const it of gd.items) it.members = it.members.filter(m => !keys.includes(m));
    gd.items = gd.items.filter(it => it.members.length);
    let n = 1;
    while (gd.items.some(it => it.name === '그룹' + n)) n++;
    gd.items.push({ id: 'G' + n + '_' + Date.now().toString(36).slice(-3) + gd.items.length, name: '그룹' + n, members: keys.slice() });
    XL().touch(wb);
    return { field: gd.field };
  }
  function rangeGroup(wb, pv, field, start, end, by) {
    const c = cacheOf(wb, pv);
    if (!(by > 0) || !(end >= start)) return { error: '끝 값은 시작 값보다 커야 합니다.' };
    c.ranges = c.ranges || {};
    c.ranges[field] = { start, end, by };
    XL().touch(wb);
    return {};
  }
  function ungroup(wb, sh, pv, field, keys) {
    const c = cacheOf(wb, pv);
    const gd = groupDef(c, field);
    if (gd) {
      const gk = keys.filter(k => k.startsWith('g:')).map(k => k.slice(2));
      gd.items = gd.items.filter(it => !gk.includes(it.id));
      if (!gd.items.length || !gk.length) {
        c.groups = c.groups.filter(x => x !== gd);
        for (const s of wb.sheets) for (const p of s.pivots) if (p.cacheId === c.id) for (const a of ['rows', 'cols', 'filters']) p[a] = p[a].filter(f => f !== gd.field);
      }
      XL().touch(wb);
      return {};
    }
    if (c.ranges && c.ranges[field]) { delete c.ranges[field]; XL().touch(wb); return {}; }
    return { error: '선택 항목의 그룹을 해제할 수 없습니다.' };
  }

  // ------------------------------------------------------------------ 복사
  function coveredBy(wb, sh, rg) {
    for (const pv of sh.pivots) {
      const t = tableRange(wb, sh, pv);
      if (t && rg.r1 <= t.r1 && rg.c1 <= t.c1 && rg.r2 >= t.r2 && rg.c2 >= t.c2) return { id: pv.id, dr: t.r1 - rg.r1, dc: t.c1 - rg.c1 };
    }
    return null;
  }
  function pasteCopy(wb, sh, info, r0, c0) {
    const src = find(wb, info.id);
    if (!src) return;
    const c = cacheOf(wb, src.pv);
    const pv = U.clone(src.pv);
    let n = 1;
    const names = wb.sheets.flatMap(s => s.pivots.map(p => p.name));
    while (names.includes('피벗 테이블' + n)) n++;
    pv.id = 'pv' + (wb.seq.pivot++);
    pv.name = '피벗 테이블' + n;
    pv.r = r0 + info.dr; pv.c = c0 + info.dc;
    // 그룹 필드의 항목 이름·필드 머리글은 피벗마다 따로 적용되므로 붙여 넣은 피벗은 초기 이름으로 시작한다(11.2 주의)
    for (const gd of c.groups) { delete pv.itemNames[gd.field]; delete pv.captions[gd.field]; }
    sh.pivots.push(pv);
    XL().touch(wb);
    return pv;
  }
  function removePivot(wb, sh, pv) { sh.pivots = sh.pivots.filter(p => p !== pv); XL().touch(wb); }
  /* 첫 행 필드 항목별 합계와 총합계(채점·검산용) */
  function summary(wb, pv, level) {
    const c = cacheOf(wb, pv);
    const rows = pv.rows.filter(f => f !== 'Σ값');
    const f = rows[level || 0];
    const out = {};
    if (!f) return { items: out, total: pv.values.map(v => aggregate(c, c.records, v)) };
    const groups = new Map();
    for (const r of c.records) { const k = itemKeyOf(c, f, r); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); }
    for (const [k, recs] of groups) out[itemLabel(c, pv, f, k)] = { key: k, values: pv.values.map(v => aggregate(c, recs, v)), recs };
    return { field: f, items: out, total: pv.values.map(v => aggregate(c, c.records, v)) };
  }
  /* 격자에서 보이는 레이블(필드별) */
  function visibleLabels(wb, sh, pv, field) {
    const g = grid(wb, sh, pv);
    const out = [];
    for (const row of g.rows) for (const cell of row) if (cell && cell.kind === 'label' && cell.field === field) out.push(cell.text);
    return out;
  }

  OA.pivot = { FUNC, FUNC_LIST, rawKey, keyVal, create, find, refresh, cacheOf, allFields, baseFields, isNumericField, addField, removeField, removeValue,
    moveInArea, setValueField, placed, valueName, grid, tableRange, hitTest, typeInto, toggle, setCollapsed, collapseField, group, rangeGroup, ungroup,
    coveredBy, pasteCopy, removePivot, itemLabel, caption, groupDef, isCollapsed, baseOrderForField, itemKeyOf, aggregate, summary, visibleLabels };
})(typeof window !== 'undefined' ? window : globalThis);

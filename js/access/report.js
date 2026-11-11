/* Access 식 계산기와 보고서·폼 출력 엔진(DOM 없음).
   식: [필드], "문자", 숫자, & + - * / \ Mod ^, 비교, And Or Not, Like, Is Null, 함수(Sum Avg Count Min Max IIf Nz Format Date Now Year Month Day Left Right Mid Len Int Round Val UCase LCase Trim), [Page] [Pages].
   출력: 구역 인스턴스(보고서 머리글, 페이지 머리글, 그룹 머리글, 본문, 그룹 바닥글, 페이지 바닥글, 보고서 바닥글)를 쪽으로 나눈 결과. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util;

  // ------------------------------------------------------------------ 식
  function tokenize(s) {
    const t = [];
    let i = 0;
    s = String(s || '');
    while (i < s.length) {
      const ch = s[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"' || ch === "'") { let j = i + 1, v = ''; while (j < s.length) { if (s[j] === ch) { if (s[j + 1] === ch) { v += ch; j += 2; continue; } break; } v += s[j]; j++; } t.push({ k: 'str', v }); i = j + 1; continue; }
      if (ch === '[') { const j = s.indexOf(']', i); t.push({ k: 'id', v: s.slice(i + 1, j < 0 ? s.length : j) }); i = j < 0 ? s.length : j + 1; continue; }
      if (ch === '#') { const j = s.indexOf('#', i + 1); if (j > 0) { t.push({ k: 'date', v: s.slice(i + 1, j) }); i = j + 1; continue; } }
      let m = /^\d+(\.\d+)?/.exec(s.slice(i));
      if (m) { t.push({ k: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      m = /^[A-Za-z_ㄱ-힝][A-Za-z0-9_ㄱ-힝]*/.exec(s.slice(i));
      if (m) { const up = m[0].toUpperCase(); t.push(['AND', 'OR', 'NOT', 'MOD', 'LIKE', 'IS', 'NULL', 'TRUE', 'FALSE'].includes(up) ? { k: 'kw', v: up } : { k: 'id', v: m[0], bare: true }); i += m[0].length; continue; }
      const two = s.slice(i, i + 2);
      if (two === '<=' || two === '>=' || two === '<>') { t.push({ k: 'op', v: two }); i += 2; continue; }
      if ('=<>+-*/\\&^(),!.'.includes(ch)) { t.push({ k: 'op', v: ch }); i++; continue; }
      if (ch === '%') throw new Error('구문 오류');
      t.push({ k: 'op', v: ch }); i++;
    }
    return t;
  }
  function parse(src) {
    const toks = tokenize(src);
    let i = 0;
    const peek = () => toks[i], next = () => toks[i++];
    const BP = { OR: 1, AND: 2, '=': 4, '<>': 4, '<': 4, '>': 4, '<=': 4, '>=': 4, LIKE: 4, IS: 4, '&': 5, '+': 6, '-': 6, MOD: 7, '\\': 8, '*': 9, '/': 9, '^': 10 };
    function nud(t) {
      if (!t) throw new Error('식이 끝나지 않았습니다');
      if (t.k === 'num') return { t: 'num', v: t.v };
      if (t.k === 'str') return { t: 'str', v: t.v };
      if (t.k === 'date') return { t: 'str', v: t.v };
      if (t.k === 'kw' && t.v === 'NULL') return { t: 'null' };
      if (t.k === 'kw' && (t.v === 'TRUE' || t.v === 'FALSE')) return { t: 'num', v: t.v === 'TRUE' ? -1 : 0 };
      if (t.k === 'kw' && t.v === 'NOT') return { t: 'not', a: expr(3) };
      if (t.k === 'op' && t.v === '-') return { t: 'neg', a: expr(11) };
      if (t.k === 'op' && t.v === '+') return expr(11);
      if (t.k === 'op' && t.v === '(') { const e = expr(0); next(); return e; }
      if (t.k === 'id') {
        if (peek() && peek().k === 'op' && peek().v === '(') {
          next();
          const args = [];
          if (peek() && peek().k === 'op' && peek().v === ')') { next(); return { t: 'fn', name: t.v.toUpperCase(), args }; }
          for (;;) {
            if (peek() && peek().k === 'op' && peek().v === '*') { next(); args.push({ t: 'star' }); } else args.push(expr(0));
            const s = next();
            if (!s || s.v === ')') break;
          }
          return { t: 'fn', name: t.v.toUpperCase(), args };
        }
        let name = t.v;
        while (peek() && peek().k === 'op' && (peek().v === '.' || peek().v === '!') && toks[i + 1] && toks[i + 1].k === 'id') { next(); name = next().v; }
        return { t: 'id', v: name };
      }
      throw new Error('구문 오류');
    }
    function expr(rbp) {
      let left = nud(next());
      for (;;) {
        const t = peek();
        if (!t) break;
        const op = t.k === 'kw' ? t.v : t.k === 'op' ? t.v : null;
        const bp = BP[op];
        if (!bp || bp <= rbp) break;
        next();
        if (op === 'IS') { let neg = false; if (peek() && peek().v === 'NOT') { next(); neg = true; } next(); left = { t: 'isnull', a: left, neg }; continue; }
        left = { t: 'bin', op, a: left, b: expr(op === '^' ? bp - 1 : bp) };
      }
      return left;
    }
    const ast = expr(0);
    if (i < toks.length) throw new Error('구문 오류');
    return ast;
  }
  const isNull = v => v === null || v === undefined;
  const num = v => isNull(v) ? null : typeof v === 'number' ? v : (v instanceof Date ? v.getTime() : (isFinite(Number(v)) && String(v).trim() !== '' ? Number(v) : NaN));
  function str(v) {
    if (isNull(v)) return '';
    if (v instanceof Date) return fmtDate(v, v.__time ? 'yyyy-mm-dd ampm h:nn:ss' : 'yyyy-mm-dd');
    if (typeof v === 'number') return generalNumber(v);
    return String(v);
  }
  function generalNumber(v) {
    if (Number.isInteger(v)) return String(v);
    let s = U.roundHalfAway(v, 6).toFixed(6).replace(/0+$/, '').replace(/\.$/, '');
    return s;
  }
  function fmtDate(d, f) {
    const p = n => String(n).padStart(2, '0');
    const h = d.getHours();
    return f.replace('yyyy', d.getFullYear()).replace('mm', p(d.getMonth() + 1)).replace('dd', p(d.getDate()))
      .replace('ampm', h < 12 ? '오전' : '오후').replace('h:nn:ss', `${(h % 12) || 12}:${p(d.getMinutes())}:${p(d.getSeconds())}`)
      .replace('hh', p(h)).replace('nn', p(d.getMinutes())).replace('ss', p(d.getSeconds()));
  }
  function formatValue(v, f) {
    if (isNull(v)) return '';
    const fs = String(f || '');
    if (v instanceof Date || (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v) && /[ymd]/i.test(fs))) {
      const d = v instanceof Date ? v : new Date(v.length <= 10 ? v + 'T00:00:00' : v.replace(' ', 'T'));
      if (/^(기본 날짜|General Date)$/i.test(fs) || !fs) return fmtDate(d, 'yyyy-mm-dd');
      return fmtDate(d, fs.replace(/YYYY/g, 'yyyy').replace(/DD/g, 'dd').replace(/(h+):(m+)/, (m0, a, b) => a + ':nn'));
    }
    if (typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && isFinite(Number(v)))) {
      const n = Number(v);
      if (/^(통화|Currency)$/i.test(fs)) return OA.acc.formatValue(n, { format: '통화', decimals: 0 });
      if (/^(표준|Standard)$/i.test(fs)) return OA.acc.formatValue(n, { format: '표준' });
      if (/^(고정|Fixed)$/i.test(fs)) return OA.acc.formatValue(n, { format: '고정' });
      if (/^(백분율|Percent)$/i.test(fs)) return OA.acc.formatValue(n, { format: '백분율' });
      if (fs) return OA.numfmt.format(n, fs).text;
      return generalNumber(n);
    }
    return String(v);
  }
  function likeRe(p) { return new RegExp('^' + String(p).replace(/[.+^${}()|\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/#/g, '\\d') + '$', 'i'); }

  function ev(n, ctx) {
    switch (n.t) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'null': return null;
      case 'star': return 1;
      case 'id': {
        const k = n.v;
        if (/^page$/i.test(k)) return ctx.page || 1;
        if (/^pages$/i.test(k)) return ctx.pages || 1;
        if (ctx.row && Object.prototype.hasOwnProperty.call(ctx.row, k)) return ctx.row[k];
        if (ctx.controls && ctx.controls[k] !== undefined) return ctx.controls[k];
        if (ctx.row) { const hit = Object.keys(ctx.row).find(x => x.toLowerCase() === k.toLowerCase()); if (hit) return ctx.row[hit]; }
        throw Object.assign(new Error('#이름?'), { name404: k });
      }
      case 'neg': { const a = num(ev(n.a, ctx)); return a === null ? null : -a; }
      case 'not': { const a = ev(n.a, ctx); return isNull(a) ? null : (a ? 0 : -1); }
      case 'isnull': { const a = isNull(ev(n.a, ctx)); return (n.neg ? !a : a) ? -1 : 0; }
      case 'bin': {
        if (n.op === 'AND' || n.op === 'OR') { const a = ev(n.a, ctx), b = ev(n.b, ctx); const x = !!(a && a !== 0), y = !!(b && b !== 0); return (n.op === 'AND' ? x && y : x || y) ? -1 : 0; }
        const a = ev(n.a, ctx), b = ev(n.b, ctx);
        if (n.op === '&') return str(a) + str(b);
        if (isNull(a) || isNull(b)) return null;
        if (n.op === 'LIKE') return likeRe(b).test(str(a)) ? -1 : 0;
        if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) {
          const x = typeof a === 'number' && typeof b === 'number' ? a : str(a), y = typeof a === 'number' && typeof b === 'number' ? b : str(b);
          const c = x === y ? 0 : x < y ? -1 : 1;
          const r = n.op === '=' ? c === 0 : n.op === '<>' ? c !== 0 : n.op === '<' ? c < 0 : n.op === '>' ? c > 0 : n.op === '<=' ? c <= 0 : c >= 0;
          return r ? -1 : 0;
        }
        const x = num(a), y = num(b);
        if (n.op === '+' && (typeof a === 'string' || typeof b === 'string') && (isNaN(x) || isNaN(y))) return str(a) + str(b);
        if (isNaN(x) || isNaN(y)) throw new Error('#오류');
        switch (n.op) {
          case '+': return x + y; case '-': return x - y; case '*': return x * y;
          case '/': if (y === 0) throw new Error('#0으로 나누기'); return x / y;
          case '\\': return Math.trunc(Math.round(x) / Math.round(y));
          case 'MOD': return Math.round(x) % Math.round(y);
          case '^': return Math.pow(x, y);
        }
        return null;
      }
      case 'fn': return callFn(n, ctx);
    }
    return null;
  }
  function aggregate(n, ctx, kind) {
    const rows = ctx.scope || [];
    const arg = n.args[0];
    if (kind === 'COUNT' && (!arg || arg.t === 'star')) return rows.length;
    const vals = rows.map(r => ev(arg, Object.assign({}, ctx, { row: r }))).filter(v => !isNull(v));
    if (kind === 'COUNT') return vals.length;
    const ns = vals.map(Number);
    if (kind === 'SUM') return ns.length ? ns.reduce((s, x) => s + x, 0) : null;
    if (kind === 'AVG') return ns.length ? ns.reduce((s, x) => s + x, 0) / ns.length : null;
    if (kind === 'MIN') return ns.length ? Math.min(...ns) : null;
    if (kind === 'MAX') return ns.length ? Math.max(...ns) : null;
    return null;
  }
  function callFn(n, ctx) {
    const A = i => ev(n.args[i], ctx);
    switch (n.name) {
      case 'SUM': case 'AVG': case 'COUNT': case 'MIN': case 'MAX': return aggregate(n, ctx, n.name);
      case 'IIF': { const c = A(0); return (c && c !== 0) ? A(1) : A(2); }
      case 'NZ': { const v = A(0); return isNull(v) ? (n.args[1] ? A(1) : 0) : v; }
      case 'FORMAT': return formatValue(A(0), n.args[1] ? A(1) : '');
      case 'DATE': { const d = ctx.now ? new Date(ctx.now) : new Date(); d.setHours(0, 0, 0, 0); return d; }
      case 'NOW': { const d = ctx.now ? new Date(ctx.now) : new Date(); d.__time = true; return d; }
      case 'YEAR': { const v = A(0); return v instanceof Date ? v.getFullYear() : +String(v).slice(0, 4); }
      case 'MONTH': { const v = A(0); return v instanceof Date ? v.getMonth() + 1 : +String(v).slice(5, 7); }
      case 'DAY': { const v = A(0); return v instanceof Date ? v.getDate() : +String(v).slice(8, 10); }
      case 'LEFT': return str(A(0)).slice(0, A(1));
      case 'RIGHT': { const s = str(A(0)), k = A(1); return k <= 0 ? '' : s.slice(-k); }
      case 'MID': return n.args.length > 2 ? str(A(0)).substr(A(1) - 1, A(2)) : str(A(0)).slice(A(1) - 1);
      case 'LEN': return str(A(0)).length;
      case 'INT': return Math.floor(num(A(0)));
      case 'ROUND': { const x = num(A(0)), k = n.args[1] ? A(1) : 0; return U.roundHalfAway(x, k); }
      case 'VAL': { const m = /^\s*-?\d+(\.\d+)?/.exec(str(A(0))); return m ? +m[0] : 0; }
      case 'UCASE': return str(A(0)).toUpperCase();
      case 'LCASE': return str(A(0)).toLowerCase();
      case 'TRIM': return str(A(0)).trim();
      case 'ABS': return Math.abs(num(A(0)));
    }
    throw Object.assign(new Error('#이름?'), { fn404: n.name });
  }
  const cacheP = new Map();
  function evaluate(src, ctx) {
    let ast = cacheP.get(src);
    if (!ast) { ast = parse(src); cacheP.set(src, ast); }
    return ev(ast, ctx || {});
  }
  function isAggregate(src) { return /\b(Sum|Avg|Count|Min|Max)\s*\(/i.test(src || ''); }

  // ------------------------------------------------------------------ 컨트롤 값
  function controlValue(c, ctx, fieldInfo) {
    if (c.type === 'label') return { text: c.caption || '' };
    if (c.type === 'line' || c.type === 'rect') return { text: '' };
    const src = (c.controlSource || '').trim();
    if (!src) return { text: '', value: null };
    let value, info = null;
    try {
      if (src.startsWith('=')) value = evaluate(src.slice(1), ctx);
      else {
        if (ctx.row && Object.prototype.hasOwnProperty.call(ctx.row, src)) value = ctx.row[src];
        else return { text: '#이름?', value: null, error: '#이름?' };
        info = fieldInfo && fieldInfo[src];
      }
    } catch (e) { return { text: e.message.startsWith('#') ? e.message : '#오류', value: null, error: e.message }; }
    const fmt = c.format || (info && info.format) || '';
    const dec = c.decimals !== undefined && c.decimals !== '자동' ? c.decimals : (info && info.decimals !== undefined ? info.decimals : '자동');
    let text;
    if (typeof value === 'number' && fmt) text = OA.acc.formatValue(value, { format: fmt, decimals: dec });
    else if (value instanceof Date) text = formatValue(value, fmt || (value.__time ? '' : ''));
    else text = typeof value === 'number' ? generalNumber(value) : str(value);
    if (value instanceof Date && !fmt) text = value.__time ? fmtDate(value, 'yyyy-mm-dd ampm h:nn:ss') : fmtDate(value, 'yyyy-mm-dd');
    return { text, value };
  }

  // ------------------------------------------------------------------ 데이터
  function loadSource(d, src) {
    if (!src) return { rows: [], columns: [], info: {} };
    const s = src.trim();
    let r;
    if (/^select\s/i.test(s)) r = OA.acc.runSQL(d, s);
    else if (d.tables[s] || d.queries[s]) r = OA.acc.runSaved(d, s);
    else return { error: `Microsoft Access 데이터베이스 엔진에서 입력 테이블 또는 쿼리 '${s}'을(를) 찾을 수 없습니다.` };
    if (r.needParam) return { needParam: r.needParam };
    const info = {};
    (r.columns || []).forEach((c, i) => { info[c] = (r.fieldInfo && r.fieldInfo[i]) || {}; });
    return { columns: r.columns, rows: r.rows.map(row => Object.fromEntries(r.columns.map((c, i) => [c, row[i]]))), info };
  }
  function cmpVals(a, b) {
    if (isNull(a) && isNull(b)) return 0;
    if (isNull(a)) return -1; if (isNull(b)) return 1;
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    return U.koCompare(String(a), String(b));
  }

  // ------------------------------------------------------------------ 보고서
  function pageBox(page) {
    const W = page.paper === 'A3' ? 297 : 210, H = page.paper === 'A3' ? 420 : 297;
    const P = page.orient === 'landscape' ? { w: H, h: W } : { w: W, h: H };
    return { paperW: P.w / 10, paperH: P.h / 10, availW: (P.w - page.left - page.right) / 10, availH: (P.h - page.top - page.bottom) / 10 };
  }
  function renderReport(d, rpt, opts) {
    const src = loadSource(d, rpt.props.recordSource);
    if (src.error || src.needParam) return src;
    const levels = rpt.grouping || [];
    const rows = src.rows.slice().sort((x, y) => { for (const g of levels) { const c = cmpVals(x[g.field], y[g.field]); if (c) return g.order === 'desc' ? -c : c; } return 0; });
    const groups = levels.filter(g => g.kind === 'group');
    const secs = rpt.sections;
    const ctrls = id => rpt.controls.filter(c => c.section === id);
    const inst = [];
    const now = opts && opts.now;
    const mk = (id, ctx) => ({ section: id, height: (secs[id] && secs[id].height) || 0, ctx });
    if (secs.reportHeader) inst.push(mk('reportHeader', { scope: rows, row: rows[0] || {} }));
    const open = []; // 각 그룹 수준의 현재 행 목록
    const flushFooters = (from) => {
      for (let k = groups.length - 1; k >= from; k--) {
        if (open[k] && secs['groupFooter' + k]) inst.push(mk('groupFooter' + k, { scope: open[k], row: open[k][open[k].length - 1] || {} }));
        open[k] = null;
      }
    };
    let prev = null, detailIdx = 0;
    rows.forEach((row, ri) => {
      let brk = -1;
      for (let k = 0; k < groups.length; k++) if (!prev || cmpVals(prev[groups[k].field], row[groups[k].field]) !== 0) { brk = k; break; }
      if (brk >= 0) {
        flushFooters(brk);
        for (let k = brk; k < groups.length; k++) {
          const scope = rows.slice(ri).filter((r, idx, arr) => { for (let j = 0; j <= k; j++) if (cmpVals(r[groups[j].field], row[groups[j].field]) !== 0) return false; return true; });
          open[k] = scope;
          if (secs['groupHeader' + k]) inst.push(mk('groupHeader' + k, { scope, row }));
        }
        detailIdx = 0;
      }
      inst.push(Object.assign(mk('detail', { scope: [row], row }), { alt: detailIdx++ % 2 === 1, prev }));
      prev = row;
    });
    flushFooters(0);
    if (secs.reportFooter) inst.push(mk('reportFooter', { scope: rows, row: rows[rows.length - 1] || {} }));
    // 쪽 나누기
    const box = pageBox(rpt.page);
    const phH = secs.pageHeader ? secs.pageHeader.height : 0, pfH = secs.pageFooter ? secs.pageFooter.height : 0;
    const pages = [];
    let page = null;
    const newPage = (first) => { page = { items: [], y: 0 }; pages.push(page); if (secs.pageHeader && !(first && inst[0] && inst[0].section === 'reportHeader' && false)) { page.items.push({ inst: mk('pageHeader', { scope: rows, row: rows[0] || {} }), y: 0 }); page.y = phH; } };
    // 보고서 머리글은 첫 쪽 맨 위, 그다음 페이지 머리글
    page = { items: [], y: 0 }; pages.push(page);
    let k0 = 0;
    if (inst[0] && inst[0].section === 'reportHeader') { page.items.push({ inst: inst[0], y: 0 }); page.y = inst[0].height; k0 = 1; }
    if (secs.pageHeader) { page.items.push({ inst: mk('pageHeader', { scope: rows, row: rows[0] || {} }), y: page.y }); page.y += phH; }
    for (let k = k0; k < inst.length; k++) {
      const it = inst[k];
      const limit = box.availH - pfH;
      if (page.y + it.height > limit + 1e-6 && page.y > phH + 1e-6) newPage();
      page.items.push({ inst: it, y: page.y });
      page.y += it.height;
    }
    const hPages = Math.max(1, Math.ceil((rpt.props.width - 1e-6) / box.availW));
    // 컨트롤 계산(쪽 번호 포함)
    const lastValue = {};
    pages.forEach((pg, pi) => {
      if (secs.pageFooter) pg.items.push({ inst: mk('pageFooter', { scope: rows, row: {} }), y: box.availH - pfH });
      for (const item of pg.items) {
        const id = item.inst.section;
        item.controls = ctrls(id).map(c => {
          const ctx = Object.assign({}, item.inst.ctx, { page: pi + 1, pages: pages.length * hPages, now });
          const val = controlValue(c, ctx, src.info);
          let hidden = false;
          if (c.hideDuplicates && id === 'detail' && c.type === 'textbox') {
            const keyc = c.id;
            if (item.inst.prev !== null && item.inst.prev !== undefined && lastValue[keyc] !== undefined && lastValue[keyc] === val.text) hidden = true;
            lastValue[keyc] = val.text;
          }
          return { c, text: hidden ? '' : val.text, value: val.value, error: val.error, hidden };
        });
        if (id !== 'detail') for (const c of ctrls(id)) if (c.hideDuplicates) delete lastValue[c.id];
      }
    });
    return { pages, hPages, pageCount: pages.length * hPages, box, rows, info: src.info, instances: inst, width: rpt.props.width };
  }

  // ------------------------------------------------------------------ 폼
  function listboxData(d, c) {
    const src = (c.rowSource || '').trim();
    if (!src) return { columns: [], rows: [] };
    if (c.rowSourceType === '값 목록') {
      const vals = src.split(';').map(s => s.replace(/^"|"$/g, ''));
      const n = Math.max(1, +c.columnCount || 1);
      const rows = []; for (let i = 0; i < vals.length; i += n) rows.push(vals.slice(i, i + n));
      return { columns: [], rows, info: [] };
    }
    if (c.rowSourceType === '필드 목록') { const r = loadSource(d, src); return { columns: [], rows: (r.columns || []).map(x => [x]) }; }
    const r = loadSource(d, src);
    if (r.error || r.needParam) return { error: r.error || ('매개 변수: ' + r.needParam), columns: [], rows: [] };
    return { columns: r.columns, rows: r.rows.map(o => r.columns.map(k => o[k])), info: r.columns.map(k => r.info[k]) };
  }
  function parseWidths(s, n, total) {
    const parts = String(s || '').split(';').map(x => x.trim()).filter(x => x !== '');
    const w = parts.map(x => { const m = /^([\d.]+)\s*(cm|mm|in|")?$/i.exec(x); if (!m) return null; const v = +m[1]; return m[2] === 'mm' ? v / 10 : (m[2] === 'in' || m[2] === '"') ? v * 2.54 : v; });
    const out = [];
    const known = w.filter(x => x !== null).reduce((s, x) => s + x, 0);
    const rest = Math.max(0, n - w.length);
    for (let i = 0; i < n; i++) out.push(i < w.length && w[i] !== null ? w[i] : (rest ? Math.max(0.5, (total - known) / rest) : 2.54));
    return out;
  }
  function renderForm(d, f) {
    const box = pageBox(f.page);
    const h = f.sections.detail.height;
    const vPages = Math.max(1, Math.ceil((h - 1e-6) / box.availH));
    const hPages = Math.max(1, Math.ceil((f.props.width - 1e-6) / box.availW));
    const items = f.controls.filter(c => c.visible !== false).map(c => {
      if (c.type === 'listbox') return { c, list: listboxData(d, c) };
      if (c.type === 'textbox') { const v = controlValue(c, { row: {}, scope: [] }, {}); return { c, text: v.text }; }
      return { c, text: c.type === 'label' ? c.caption : '' };
    });
    return { box, items, pageCount: vPages * hPages, hPages, vPages, width: f.props.width, height: h };
  }

  OA.accExpr = { tokenize, parse, evaluate, formatValue, generalNumber, isAggregate, str };
  OA.report = { renderReport, renderForm, listboxData, parseWidths, pageBox, loadSource, controlValue };
})(typeof window !== 'undefined' ? window : globalThis);

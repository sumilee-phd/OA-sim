/* Excel 수식 엔진(DOM 없음): 토큰 분석 → 파서(AST) → 계산, 참조 이동(복사·채우기·정렬). 외부 라이브러리 없이 작성. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util;
  const NF = OA.numfmt;

  const ERR = (e) => ({ e });
  const E = { VALUE: ERR('#VALUE!'), REF: ERR('#REF!'), DIV0: ERR('#DIV/0!'), NA: ERR('#N/A'), NUM: ERR('#NUM!'), NAME: ERR('#NAME?') };
  const isErr = v => v && typeof v === 'object' && typeof v.e === 'string';
  const isRange = v => v && typeof v === 'object' && v.kind === 'range';
  const isArr = v => v && typeof v === 'object' && v.kind === 'arr';

  // ------------------------------------------------------------------ 토큰
  const REF1 = /^\$?[A-Za-z]{1,3}\$?[0-9]+/;
  function tokenize(src) {
    const t = [];
    let i = 0;
    const s = src;
    while (i < s.length) {
      const ch = s[i];
      if (ch === ' ' || ch === '\n' || ch === '\t' || ch === '\r') { i++; continue; }
      const start = i;
      if (ch === '"') {
        let j = i + 1, v = '';
        while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { v += '"'; j += 2; continue; } break; } v += s[j]; j++; }
        t.push({ k: 'str', v, p: start, e: j + 1 }); i = j + 1; continue;
      }
      if (ch === '#') {
        const m = /^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!)/i.exec(s.slice(i));
        if (m) { t.push({ k: 'err', v: m[0].toUpperCase(), p: start, e: i + m[0].length }); i += m[0].length; continue; }
      }
      // 시트 이름 붙은 참조
      let m = /^'((?:[^']|'')+)'!/.exec(s.slice(i)) || /^([A-Za-z0-9_.ㄱ-힝]+)!/.exec(s.slice(i));
      if (m) {
        const sheet = m[1].replace(/''/g, "'");
        const rest = s.slice(i + m[0].length);
        const r1 = REF1.exec(rest);
        if (r1) {
          let len = r1[0].length, text2 = null;
          const r2 = /^:(\$?[A-Za-z]{1,3}\$?[0-9]+)/.exec(rest.slice(len));
          if (r2) { text2 = r2[1]; len += r2[0].length; }
          t.push(mkRef(sheet, r1[0], text2, start, i + m[0].length + len)); i += m[0].length + len; continue;
        }
      }
      m = REF1.exec(s.slice(i));
      if (m && !/^[A-Za-z0-9_.(]/.test(s.slice(i + m[0].length).replace(/^:\$?[A-Za-z]{1,3}\$?[0-9]+/, '')) && !(i > 0 && /[A-Za-z0-9_.]/.test(s[i - 1]))) {
        let len = m[0].length, text2 = null;
        const r2 = /^:(\$?[A-Za-z]{1,3}\$?[0-9]+)/.exec(s.slice(i + len));
        if (r2) { text2 = r2[1]; len += r2[0].length; }
        t.push(mkRef(null, m[0], text2, start, i + len)); i += len; continue;
      }
      m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(s.slice(i));
      if (m) { t.push({ k: 'num', v: parseFloat(m[0]), p: start, e: i + m[0].length }); i += m[0].length; continue; }
      m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i));
      if (m) {
        const word = m[0];
        const after = s.slice(i + word.length).replace(/^\s+/, '');
        if (after[0] === '(') t.push({ k: 'fn', v: word.toUpperCase(), p: start, e: i + word.length });
        else if (/^(TRUE|FALSE)$/i.test(word)) t.push({ k: 'bool', v: /^TRUE$/i.test(word), p: start, e: i + word.length });
        else t.push({ k: 'name', v: word, p: start, e: i + word.length });
        i += word.length; continue;
      }
      // 한글 등 이름(정의되지 않은 이름)
      m = /^[ㄱ-힝][ㄱ-힝0-9_]*/.exec(s.slice(i));
      if (m) { t.push({ k: 'name', v: m[0], p: start, e: i + m[0].length }); i += m[0].length; continue; }
      const two = s.slice(i, i + 2);
      if (two === '<=' || two === '>=' || two === '<>') { t.push({ k: 'op', v: two, p: start, e: i + 2 }); i += 2; continue; }
      if ('+-*/^&=<>%'.includes(ch)) { t.push({ k: 'op', v: ch, p: start, e: i + 1 }); i++; continue; }
      if (ch === '(' || ch === ')' || ch === ',' || ch === ';') { t.push({ k: ch === ';' ? ',' : ch, v: ch, p: start, e: i + 1 }); i++; continue; }
      if (ch === ':') { t.push({ k: 'op', v: ':', p: start, e: i + 1 }); i++; continue; }
      t.push({ k: 'bad', v: ch, p: start, e: i + 1 }); i++;
    }
    return t;
  }
  function mkRef(sheet, a, b, p, e) {
    const A = U.parseAddr(a);
    if (b) {
      const B = U.parseAddr(b);
      return { k: 'range', sheet, a: A, b: B, p, e };
    }
    return { k: 'ref', sheet, a: A, p, e };
  }

  // ------------------------------------------------------------------ 파서
  function parse(src) {
    const toks = tokenize(src);
    let i = 0;
    const peek = () => toks[i];
    const next = () => toks[i++];
    function expect(k) { const t = next(); if (!t || t.k !== k) throw new SyntaxError('수식 오류'); return t; }
    const INFIX = { '=': 10, '<>': 10, '<': 10, '>': 10, '<=': 10, '>=': 10, '&': 20, '+': 30, '-': 30, '*': 40, '/': 40, '^': 50 };
    function nud(t) {
      if (!t) throw new SyntaxError('수식이 끝나지 않았습니다');
      switch (t.k) {
        case 'num': return { t: 'num', v: t.v };
        case 'str': return { t: 'str', v: t.v };
        case 'bool': return { t: 'bool', v: t.v };
        case 'err': return { t: 'err', v: t.v };
        case 'ref': return { t: 'ref', sheet: t.sheet, r: t.a.r, c: t.a.c };
        case 'range': return { t: 'range', sheet: t.sheet, r1: Math.min(t.a.r, t.b.r), c1: Math.min(t.a.c, t.b.c), r2: Math.max(t.a.r, t.b.r), c2: Math.max(t.a.c, t.b.c) };
        case 'name': return { t: 'name', v: t.v };
        case 'fn': {
          expect('(');
          const args = [];
          if (peek() && peek().k === ')') { next(); return { t: 'fn', name: t.v, args }; }
          for (;;) {
            if (peek() && (peek().k === ',' || peek().k === ')')) args.push({ t: 'missing' });
            else args.push(expr(0));
            const s = next();
            if (!s) throw new SyntaxError('괄호가 닫히지 않았습니다');
            if (s.k === ')') break;
            if (s.k !== ',') throw new SyntaxError('수식 오류');
          }
          return { t: 'fn', name: t.v, args };
        }
        case '(': { const e = expr(0); expect(')'); return e; }
        case 'op':
          if (t.v === '-') return { t: 'neg', a: expr(55) };
          if (t.v === '+') return expr(55);
          break;
      }
      throw new SyntaxError('수식 오류');
    }
    function expr(rbp) {
      let left = nud(next());
      for (;;) {
        const t = peek();
        if (!t || t.k !== 'op') break;
        if (t.v === '%') { next(); left = { t: 'pct', a: left }; continue; }
        const bp = INFIX[t.v];
        if (bp === undefined || bp <= rbp) break;
        next();
        const right = expr(t.v === '^' ? bp : bp);
        left = { t: 'bin', op: t.v, a: left, b: right };
      }
      return left;
    }
    const ast = expr(0);
    if (i < toks.length) throw new SyntaxError('수식 오류');
    return ast;
  }

  /* 상대 참조 이동: 수식 문자열(맨 앞 = 포함)과 행·열 이동량 */
  function shift(formula, dr, dc) {
    const body = formula.startsWith('=') ? formula.slice(1) : formula;
    const toks = tokenize(body);
    let out = '', pos = 0;
    for (const t of toks) {
      if (t.k !== 'ref' && t.k !== 'range') continue;
      out += body.slice(pos, t.p);
      const mv = (a) => {
        const r = a.absR ? a.r : a.r + dr, c = a.absC ? a.c : a.c + dc;
        if (r < 1 || c < 1) return null;
        return U.addr(r, c, a.absR, a.absC);
      };
      const sheetPart = t.sheet ? (/[^\wㄱ-힝]/.test(t.sheet) ? "'" + t.sheet + "'" : t.sheet) + '!' : '';
      const a = mv(t.a), b = t.k === 'range' ? mv(t.b) : '';
      if (a === null || b === null) out += '#REF!';
      else out += sheetPart + a + (t.k === 'range' ? ':' + b : '');
      pos = t.e;
    }
    out += body.slice(pos);
    return (formula.startsWith('=') ? '=' : '') + out;
  }

  /* F4: 커서 위치의 참조를 $A$1 → A$1 → $A1 → A1 순으로 */
  function cycleRef(text, caret) {
    const body = text;
    const toks = tokenize(body.startsWith('=') ? body.slice(1) : body);
    const off = body.startsWith('=') ? 1 : 0;
    for (const t of toks) {
      if ((t.k === 'ref' || t.k === 'range') && caret >= t.p + off && caret <= t.e + off) {
        // A1 → $A$1 → A$1 → $A1 → A1
        const step = (a) => {
          let absR, absC;
          if (!a.absR && !a.absC) { absR = true; absC = true; } else if (a.absR && a.absC) { absR = true; absC = false; } else if (a.absR && !a.absC) { absR = false; absC = true; } else { absR = false; absC = false; }
          return { r: a.r, c: a.c, absR, absC };
        };
        const a = step(t.a);
        let txt = U.addr(a.r, a.c, a.absR, a.absC);
        if (t.k === 'range') { const b = step(t.b); txt += ':' + U.addr(b.r, b.c, b.absR, b.absC); }
        const sheetPart = t.sheet ? body.slice(t.p + off, t.e + off).split('!')[0] + '!' : '';
        const nt = body.slice(0, t.p + off) + sheetPart + txt + body.slice(t.e + off);
        return { text: nt, caret: t.p + off + sheetPart.length + txt.length };
      }
    }
    return { text, caret };
  }

  // ------------------------------------------------------------------ 값 변환
  function toNum(v) {
    if (isErr(v)) return v;
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      const p = NF.parseInput(v.trim());
      if (typeof p.v === 'number') return p.v;
      return E.VALUE;
    }
    return E.VALUE;
  }
  function toStr(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return NF.generalNum(v);
    return String(v);
  }
  function toBool(v) {
    if (isErr(v)) return v;
    if (typeof v === 'boolean') return v;
    if (v === null || v === undefined || v === '') return false;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') { if (/^true$/i.test(v)) return true; if (/^false$/i.test(v)) return false; return E.VALUE; }
    return E.VALUE;
  }
  function cmp(a, b) { // Excel 비교: 숫자 < 문자 < 논리
    const rank = v => (v === null || v === undefined) ? -1 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
    if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra - rb;
    if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x === y ? 0 : U.koCompare(x, y) < 0 ? -1 : 1; }
    if (typeof a === 'boolean') return (a ? 1 : 0) - (b ? 1 : 0);
    return a === b ? 0 : a < b ? -1 : 1;
  }

  // ------------------------------------------------------------------ 계산
  function rangeValues(ctx, rg) {
    const out = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      const row = [];
      for (let c = rg.c1; c <= rg.c2; c++) row.push(ctx.get(rg.sheet || ctx.sheet, r, c));
      out.push(row);
    }
    return out;
  }
  function toArray(ctx, v) {
    if (isRange(v)) return rangeValues(ctx, v);
    if (isArr(v)) return v.a;
    return [[v]];
  }
  function implicit(ctx, v) {
    if (isRange(v)) {
      if (v.r1 === v.r2 && v.c1 === v.c2) return ctx.get(v.sheet || ctx.sheet, v.r1, v.c1);
      if (v.c1 === v.c2 && ctx.row >= v.r1 && ctx.row <= v.r2) return ctx.get(v.sheet || ctx.sheet, ctx.row, v.c1);
      if (v.r1 === v.r2 && ctx.col >= v.c1 && ctx.col <= v.c2) return ctx.get(v.sheet || ctx.sheet, v.r1, ctx.col);
      return E.VALUE;
    }
    if (isArr(v)) return v.a[0][0];
    return v;
  }
  function mapArr(ctx, vals, fn) { // 배열 모드에서 원소별 계산
    const arrs = vals.map(v => (isRange(v) || isArr(v)) ? toArray(ctx, v) : null);
    let R = 1, C = 1;
    arrs.forEach(a => { if (a) { R = Math.max(R, a.length); C = Math.max(C, a[0].length); } });
    const out = [];
    for (let r = 0; r < R; r++) {
      const row = [];
      for (let c = 0; c < C; c++) {
        const args = vals.map((v, k) => {
          const a = arrs[k];
          if (!a) return v;
          const rr = a.length === 1 ? 0 : r, cc = a[0].length === 1 ? 0 : c;
          return (a[rr] && a[rr][cc] !== undefined) ? a[rr][cc] : E.NA;
        });
        row.push(fn(...args));
      }
      out.push(row);
    }
    return { kind: 'arr', a: out };
  }

  function binop(op, a, b) {
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    if (op === '&') return toStr(a) + toStr(b);
    if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
      const c = cmp(a, b);
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    }
    const x = toNum(a), y = toNum(b);
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    switch (op) {
      case '+': return x + y;
      case '-': return x - y;
      case '*': return x * y;
      case '/': return y === 0 ? E.DIV0 : x / y;
      case '^': return Math.pow(x, y);
    }
    return E.VALUE;
  }

  function ev(n, ctx) {
    switch (n.t) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'bool': return n.v;
      case 'err': return ERR(n.v);
      case 'missing': return null;
      case 'name': return E.NAME;
      case 'ref': return { kind: 'range', sheet: n.sheet, r1: n.r, c1: n.c, r2: n.r, c2: n.c };
      case 'range': return { kind: 'range', sheet: n.sheet, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 };
      case 'neg': {
        const a = ev(n.a, ctx);
        if (ctx.array && (isRange(a) || isArr(a))) return mapArr(ctx, [a], x => binop('-', 0, x));
        return binop('-', 0, implicit(ctx, a));
      }
      case 'pct': {
        const a = ev(n.a, ctx);
        if (ctx.array && (isRange(a) || isArr(a))) return mapArr(ctx, [a], x => binop('/', x, 100));
        return binop('/', implicit(ctx, a), 100);
      }
      case 'bin': {
        const a = ev(n.a, ctx), b = ev(n.b, ctx);
        if (ctx.array && (isRange(a) || isArr(a) || isRange(b) || isArr(b))) return mapArr(ctx, [a, b], (x, y) => binop(n.op, x, y));
        return binop(n.op, implicit(ctx, a), implicit(ctx, b));
      }
      case 'fn': {
        const f = FN[n.name];
        if (!f) return E.NAME;
        return f(n.args, ctx);
      }
    }
    return E.VALUE;
  }

  // ------------------------------------------------------------------ 함수
  const FN = {};
  const argVal = (a, ctx) => implicit(ctx, ev(a, ctx));
  const argRaw = (a, ctx) => ev(a, ctx);
  function firstErr(list) { for (const v of list) if (isErr(v)) return v; return null; }

  function numbersOf(args, ctx, opts) { // SUM 류: 범위 안 문자·논리는 무시, 직접 인수는 변환
    const out = [];
    for (const a of args) {
      if (a.t === 'missing') continue;
      const v = ev(a, ctx);
      if (isRange(v) || isArr(v)) {
        for (const row of toArray(ctx, v)) for (const x of row) {
          if (isErr(x)) return x;
          if (typeof x === 'number') out.push(x);
          else if (opts && opts.counta && x !== null && x !== '') out.push(x);
        }
      } else {
        if (isErr(v)) return v;
        if (opts && opts.counta) { if (v !== null) out.push(v); continue; }
        if (opts && opts.count) { if (typeof v === 'number' || (typeof v === 'string' && typeof NF.parseInput(v).v === 'number')) out.push(1); continue; }
        const x = toNum(v);
        if (isErr(x)) return x;
        out.push(x);
      }
    }
    return out;
  }
  FN.SUM = (args, ctx) => { const v = numbersOf(args, ctx); return isErr(v) ? v : v.reduce((s, x) => s + x, 0); };
  FN.AVERAGE = (args, ctx) => { const v = numbersOf(args, ctx); if (isErr(v)) return v; return v.length ? v.reduce((s, x) => s + x, 0) / v.length : E.DIV0; };
  FN.MIN = (args, ctx) => { const v = numbersOf(args, ctx); return isErr(v) ? v : v.length ? Math.min(...v) : 0; };
  FN.MAX = (args, ctx) => { const v = numbersOf(args, ctx); return isErr(v) ? v : v.length ? Math.max(...v) : 0; };
  FN.COUNT = (args, ctx) => { const v = numbersOf(args, ctx, { count: true }); return isErr(v) ? 0 : v.length; };
  FN.COUNTA = (args, ctx) => { const v = numbersOf(args, ctx, { counta: true }); return isErr(v) ? 0 : v.length; };
  FN.PRODUCT = (args, ctx) => { const v = numbersOf(args, ctx); return isErr(v) ? v : v.reduce((s, x) => s * x, 1); };
  FN.MEDIAN = (args, ctx) => { const v = numbersOf(args, ctx); if (isErr(v)) return v; if (!v.length) return E.NUM; v.sort((a, b) => a - b); const m = v.length >> 1; return v.length % 2 ? v[m] : (v[m - 1] + v[m]) / 2; };
  FN.LARGE = (args, ctx) => { const v = numbersOf([args[0]], ctx); const k = toNum(argVal(args[1], ctx)); if (isErr(v)) return v; v.sort((a, b) => b - a); return v[k - 1] === undefined ? E.NUM : v[k - 1]; };
  FN.SMALL = (args, ctx) => { const v = numbersOf([args[0]], ctx); const k = toNum(argVal(args[1], ctx)); if (isErr(v)) return v; v.sort((a, b) => a - b); return v[k - 1] === undefined ? E.NUM : v[k - 1]; };
  FN.COUNTBLANK = (args, ctx) => { let n = 0; for (const row of toArray(ctx, argRaw(args[0], ctx))) for (const x of row) if (x === null || x === '') n++; return n; };

  FN.IF = (args, ctx) => {
    const c = toBool(argVal(args[0], ctx));
    if (isErr(c)) return c;
    if (c) return args.length > 1 ? (args[1].t === 'missing' ? 0 : implicit(ctx, ev(args[1], ctx))) : true;
    return args.length > 2 ? (args[2].t === 'missing' ? 0 : implicit(ctx, ev(args[2], ctx))) : false;
  };
  FN.IFERROR = (args, ctx) => { const v = argVal(args[0], ctx); return isErr(v) ? argVal(args[1], ctx) : v; };
  FN.AND = (args, ctx) => { let r = true; for (const a of args) { const v = ev(a, ctx); for (const row of toArray(ctx, v)) for (const x of row) { if (x === null && isRange(v)) continue; const b = toBool(x); if (isErr(b)) return b; r = r && b; } } return r; };
  FN.OR = (args, ctx) => { let r = false; for (const a of args) { const v = ev(a, ctx); for (const row of toArray(ctx, v)) for (const x of row) { if (x === null && isRange(v)) continue; const b = toBool(x); if (isErr(b)) return b; r = r || b; } } return r; };
  FN.NOT = (args, ctx) => { const b = toBool(argVal(args[0], ctx)); return isErr(b) ? b : !b; };

  // 조건(criteria)
  function wildRe(s) {
    let re = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '~' && i + 1 < s.length) { re += s[i + 1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); i++; }
      else if (ch === '*') re += '[\\s\\S]*';
      else if (ch === '?') re += '[\\s\\S]';
      else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + re + '$', 'i');
  }
  function makeCrit(c) {
    if (isErr(c)) return () => false;
    if (typeof c === 'number') return x => typeof x === 'number' ? x === c : (typeof x === 'string' && x.trim() !== '' && Number(x) === c);
    if (typeof c === 'boolean') return x => x === c;
    if (c === null) return x => x === null || x === '';
    let s = String(c), op = '=';
    const m = /^(<=|>=|<>|=|<|>)([\s\S]*)$/.exec(s);
    if (m) { op = m[1]; s = m[2]; }
    let target = s;
    const p = NF.parseInput(s);
    const numT = (s.trim() !== '' && typeof p.v === 'number') ? p.v : null;
    if (op === '=' || op === '<>') {
      if (s === '') return op === '=' ? (x => x === null || x === '') : (x => !(x === null || x === ''));
      let test;
      if (numT !== null) test = x => (typeof x === 'number' && Math.abs(x - numT) < 1e-12) || (typeof x === 'string' && x === s);
      else if (/^(true|false)$/i.test(s)) test = x => typeof x === 'boolean' && x === /^true$/i.test(s);
      else { const re = wildRe(target); test = x => typeof x === 'string' && re.test(x); }
      return op === '=' ? test : (x => !test(x));
    }
    if (numT !== null) return x => typeof x === 'number' && (op === '<' ? x < numT : op === '>' ? x > numT : op === '<=' ? x <= numT : x >= numT);
    return x => typeof x === 'string' && (() => { const k = U.koCompare(x.toLowerCase(), s.toLowerCase()); return op === '<' ? k < 0 : op === '>' ? k > 0 : op === '<=' ? k <= 0 : k >= 0; })();
  }
  function critPairs(args, ctx, from) {
    const pairs = [];
    for (let k = from; k + 1 < args.length; k += 2) {
      const rg = argRaw(args[k], ctx);
      if (!isRange(rg)) return E.VALUE;
      pairs.push({ vals: rangeValues(ctx, rg), test: makeCrit(argVal(args[k + 1], ctx)), rg });
    }
    return pairs;
  }
  function sameShape(rg, base) {
    return { kind: 'range', sheet: rg.sheet, r1: rg.r1, c1: rg.c1, r2: rg.r1 + (base.r2 - base.r1), c2: rg.c1 + (base.c2 - base.c1) };
  }
  function condAgg(ctx, sumRg, pairs, mode) {
    if (isErr(pairs)) return pairs;
    const sv = sumRg ? rangeValues(ctx, sumRg) : null;
    let s = 0, n = 0;
    const R = pairs[0].vals.length, C = pairs[0].vals[0].length;
    for (const p of pairs) if (p.vals.length !== R || p.vals[0].length !== C) return E.VALUE;
    for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
      if (!pairs.every(p => p.test(p.vals[r][c]))) continue;
      if (mode === 'count') { n++; continue; }
      const x = sv[r] ? sv[r][c] : null;
      if (isErr(x)) return x;
      if (typeof x === 'number') { s += x; n++; }
    }
    if (mode === 'count') return n;
    if (mode === 'avg') return n ? s / n : E.DIV0;
    return s;
  }
  FN.SUMIF = (args, ctx) => {
    const rg = argRaw(args[0], ctx);
    if (!isRange(rg)) return E.VALUE;
    const pairs = [{ vals: rangeValues(ctx, rg), test: makeCrit(argVal(args[1], ctx)) }];
    let sr = args[2] && args[2].t !== 'missing' ? argRaw(args[2], ctx) : rg;
    if (!isRange(sr)) return E.VALUE;
    return condAgg(ctx, sameShape(sr, rg), pairs, 'sum');
  };
  FN.AVERAGEIF = (args, ctx) => {
    const rg = argRaw(args[0], ctx);
    if (!isRange(rg)) return E.VALUE;
    const pairs = [{ vals: rangeValues(ctx, rg), test: makeCrit(argVal(args[1], ctx)) }];
    let sr = args[2] && args[2].t !== 'missing' ? argRaw(args[2], ctx) : rg;
    return condAgg(ctx, sameShape(sr, rg), pairs, 'avg');
  };
  FN.COUNTIF = (args, ctx) => {
    const rg = argRaw(args[0], ctx);
    if (!isRange(rg)) return E.VALUE;
    return condAgg(ctx, null, [{ vals: rangeValues(ctx, rg), test: makeCrit(argVal(args[1], ctx)) }], 'count');
  };
  FN.SUMIFS = (args, ctx) => { const sr = argRaw(args[0], ctx); if (!isRange(sr)) return E.VALUE; return condAgg(ctx, sr, critPairs(args, ctx, 1), 'sum'); };
  FN.AVERAGEIFS = (args, ctx) => { const sr = argRaw(args[0], ctx); if (!isRange(sr)) return E.VALUE; return condAgg(ctx, sr, critPairs(args, ctx, 1), 'avg'); };
  FN.COUNTIFS = (args, ctx) => condAgg(ctx, null, critPairs(args, ctx, 0), 'count');

  FN.SUMPRODUCT = (args, ctx) => {
    const sub = Object.assign({}, ctx, { array: true });
    const arrs = args.map(a => toArray(sub, ev(a, sub)));
    const R = arrs[0].length, C = arrs[0][0].length;
    for (const a of arrs) if (a.length !== R || a[0].length !== C) return E.VALUE;
    let s = 0;
    for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
      let p = 1;
      for (const a of arrs) { const x = a[r][c]; if (isErr(x)) return x; p *= (typeof x === 'number') ? x : 0; }
      s += p;
    }
    return s;
  };

  function scalarFn(fn, nargs) { // 배열 모드에서 원소별로
    return (args, ctx) => {
      const vals = args.map(a => a.t === 'missing' ? undefined : ev(a, ctx));
      if (ctx.array && vals.some(v => isRange(v) || isArr(v))) return mapArr(ctx, vals, (...xs) => fn(...xs));
      return fn(...vals.map(v => v === undefined ? undefined : implicit(ctx, v)));
    };
  }
  const num1 = (f) => scalarFn((x, ...r) => { const n = toNum(x); return isErr(n) ? n : f(n, ...r); });

  FN.ISNUMBER = scalarFn(v => typeof v === 'number');
  FN.ISTEXT = scalarFn(v => typeof v === 'string');
  FN.ISBLANK = scalarFn(v => v === null || v === undefined);
  FN.ISERROR = scalarFn(v => isErr(v));
  FN.FIND = scalarFn((find, within, start) => {
    const e = firstErr([find, within, start]); if (e) return e;
    const s = toStr(within), f = toStr(find), st = start === undefined ? 1 : toNum(start);
    if (isErr(st) || st < 1) return E.VALUE;
    const k = s.indexOf(f, st - 1);
    return k < 0 ? E.VALUE : k + 1;
  });
  FN.SEARCH = scalarFn((find, within, start) => {
    const e = firstErr([find, within, start]); if (e) return e;
    const s = toStr(within), st = start === undefined ? 1 : toNum(start);
    const re = new RegExp(wildRe(toStr(find)).source.slice(1, -1), 'i');
    const m = re.exec(s.slice(st - 1));
    return m ? m.index + st : E.VALUE;
  });
  FN.LEFT = scalarFn((t, n) => { const e = firstErr([t, n]); if (e) return e; const k = n === undefined ? 1 : toNum(n); if (isErr(k) || k < 0) return E.VALUE; return toStr(t).slice(0, Math.floor(k)); });
  FN.RIGHT = scalarFn((t, n) => { const e = firstErr([t, n]); if (e) return e; const k = n === undefined ? 1 : toNum(n); if (isErr(k) || k < 0) return E.VALUE; const s = toStr(t); return k === 0 ? '' : s.slice(-Math.floor(k)); });
  FN.MID = scalarFn((t, st, n) => { const e = firstErr([t, st, n]); if (e) return e; const a = toNum(st), b = toNum(n); if (isErr(a) || isErr(b) || a < 1 || b < 0) return E.VALUE; return toStr(t).substr(Math.floor(a) - 1, Math.floor(b)); });
  FN.LEN = scalarFn(t => isErr(t) ? t : toStr(t).length);
  FN.TRIM = scalarFn(t => isErr(t) ? t : toStr(t).trim().replace(/ +/g, ' '));
  FN.VALUE = scalarFn(t => isErr(t) ? t : toNum(t));
  FN.TEXT = scalarFn((v, f) => { const e = firstErr([v, f]); if (e) return e; return NF.format(typeof v === 'string' ? toNum(v) : v, toStr(f)).text; });
  FN.CONCATENATE = (args, ctx) => { let s = ''; for (const a of args) { const v = argVal(a, ctx); if (isErr(v)) return v; s += toStr(v); } return s; };
  FN.CONCAT = FN.CONCATENATE;
  FN.ABS = num1(x => Math.abs(x));
  FN.INT = num1(x => Math.floor(x));
  FN.TRUNC = scalarFn((x, d) => { const n = toNum(x), k = d === undefined ? 0 : toNum(d); if (isErr(n)) return n; const p = Math.pow(10, k); return Math.trunc(n * p) / p; });
  FN.MOD = scalarFn((a, b) => { const x = toNum(a), y = toNum(b); if (isErr(x)) return x; if (isErr(y)) return y; if (y === 0) return E.DIV0; return x - y * Math.floor(x / y); });
  function roundFn(mode) {
    return scalarFn((a, d) => {
      const x = toNum(a), k = d === undefined ? 0 : toNum(d);
      if (isErr(x)) return x; if (isErr(k)) return k;
      const p = Math.pow(10, Math.trunc(k));
      const v = Math.abs(x) * p;
      let r;
      if (mode === 'round') r = Math.floor(v * (1 + 1e-14) + 0.5);
      else if (mode === 'up') r = Math.ceil(v * (1 - 1e-14));
      else r = Math.floor(v * (1 + 1e-14));
      return Math.sign(x) * r / p;
    });
  }
  FN.ROUND = roundFn('round');
  FN.ROUNDUP = roundFn('up');
  FN.ROUNDDOWN = roundFn('down');
  FN.CHOOSE = (args, ctx) => {
    const k = toNum(argVal(args[0], ctx));
    if (isErr(k)) return k;
    const i = Math.floor(k);
    if (i < 1 || i >= args.length) return E.VALUE;
    return argVal(args[i], ctx);
  };
  function lookupIndex(col, key, approxMode) {
    if (!approxMode) {
      for (let i = 0; i < col.length; i++) if (col[i] !== null && cmp(col[i], key) === 0 && typeof col[i] === typeof key) return i;
      for (let i = 0; i < col.length; i++) if (col[i] !== null && cmp(col[i], key) === 0) return i;
      return -1;
    }
    let lo = 0, hi = col.length - 1, ans = -1; // Excel 근사 일치는 이진 탐색
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const v = col[mid];
      if (v !== null && (typeof v === typeof key) && cmp(v, key) <= 0) { ans = mid; lo = mid + 1; } else hi = mid - 1;
    }
    return ans;
  }
  FN.VLOOKUP = (args, ctx) => {
    const key = argVal(args[0], ctx), tb = argRaw(args[1], ctx), ci = toNum(argVal(args[2], ctx));
    const approxMode = args.length < 4 || args[3].t === 'missing' ? true : toBool(argVal(args[3], ctx));
    if (isErr(key)) return key;
    if (!isRange(tb) && !isArr(tb)) return E.VALUE;
    const a = toArray(ctx, tb);
    if (isErr(ci) || ci < 1) return E.VALUE;
    if (ci > a[0].length) return E.REF;
    const i = lookupIndex(a.map(r => r[0]), key, approxMode === true);
    return i < 0 ? E.NA : (a[i][ci - 1] === null ? 0 : a[i][ci - 1]);
  };
  FN.HLOOKUP = (args, ctx) => {
    const key = argVal(args[0], ctx), tb = argRaw(args[1], ctx), ri = toNum(argVal(args[2], ctx));
    const approxMode = args.length < 4 || args[3].t === 'missing' ? true : toBool(argVal(args[3], ctx));
    const a = toArray(ctx, tb);
    if (ri > a.length) return E.REF;
    const i = lookupIndex(a[0], key, approxMode === true);
    return i < 0 ? E.NA : (a[ri - 1][i] === null ? 0 : a[ri - 1][i]);
  };
  FN.INDEX = (args, ctx) => {
    const a = toArray(ctx, argRaw(args[0], ctx));
    const r = toNum(argVal(args[1], ctx)), c = args[2] ? toNum(argVal(args[2], ctx)) : 1;
    const rr = a.length === 1 && args.length === 2 ? 1 : r, cc = a.length === 1 && args.length === 2 ? r : c;
    if (!a[rr - 1] || a[rr - 1][cc - 1] === undefined) return E.REF;
    return a[rr - 1][cc - 1] === null ? 0 : a[rr - 1][cc - 1];
  };
  FN.MATCH = (args, ctx) => {
    const key = argVal(args[0], ctx), a = toArray(ctx, argRaw(args[1], ctx));
    const type = args[2] ? toNum(argVal(args[2], ctx)) : 1;
    const flat = a.length === 1 ? a[0] : a.map(r => r[0]);
    let i;
    if (type === 0) { i = -1; const re = typeof key === 'string' ? wildRe(key) : null; for (let k = 0; k < flat.length; k++) if (re ? (typeof flat[k] === 'string' && re.test(flat[k])) : cmp(flat[k], key) === 0) { i = k; break; } } else i = lookupIndex(flat, key, true);
    return i < 0 ? E.NA : i + 1;
  };
  FN.RANK = (args, ctx) => {
    const x = toNum(argVal(args[0], ctx));
    const rg = argRaw(args[1], ctx);
    if (isErr(x)) return x;
    const order = args[2] && args[2].t !== 'missing' ? toNum(argVal(args[2], ctx)) : 0;
    const vals = toArray(ctx, rg).flat().filter(v => typeof v === 'number');
    if (!vals.some(v => v === x)) return E.NA;
    return 1 + vals.filter(v => order ? v < x : v > x).length;
  };
  FN['RANK.EQ'] = FN.RANK;

  // 날짜·시간
  FN.DATE = scalarFn((y, m, d) => {
    const Y = toNum(y), M = toNum(m), D = toNum(d);
    const e = firstErr([Y, M, D]); if (e) return e;
    const yy = Y < 1900 ? Y + 1900 : Y;
    const ms = Date.UTC(yy, Math.floor(M) - 1, Math.floor(D));
    let days = Math.round((ms - Date.UTC(1899, 11, 31)) / 86400000);
    if (days >= 60) days += 1;
    return days;
  });
  const part = (k) => num1(x => NF.serialParts(x)[k]);
  FN.YEAR = part('y'); FN.MONTH = part('m'); FN.DAY = part('d');
  FN.HOUR = part('hh'); FN.MINUTE = part('mi'); FN.SECOND = part('ss');
  FN.TIME = scalarFn((h, m, s) => { const H = toNum(h), M = toNum(m), S = toNum(s); const e = firstErr([H, M, S]); if (e) return e; return ((H * 3600 + M * 60 + S) / 86400) % 1; });
  FN.TODAY = () => { const d = new Date(); return NF.dateSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); };
  FN.NOW = () => { const d = new Date(); return NF.dateSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; };
  FN.WEEKDAY = num1(x => { const p = NF.serialParts(x); return new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay() + 1; });
  FN.DATEDIF = scalarFn((a, b, u) => {
    const s = toNum(a), t = toNum(b);
    const e = firstErr([s, t, u]); if (e) return e;
    if (s > t) return E.NUM;
    const p = NF.serialParts(s), q = NF.serialParts(t);
    let months = (q.y - p.y) * 12 + (q.m - p.m);
    if (q.d < p.d) months -= 1;
    switch (String(u).toLowerCase()) {
      case 'm': return months;
      case 'y': return Math.floor(months / 12);
      case 'd': return Math.floor(t) - Math.floor(s);
      case 'ym': return months % 12;
      case 'md': { let d = q.d - p.d; if (d < 0) { const prev = new Date(Date.UTC(q.y, q.m - 1, 0)).getUTCDate(); d += prev; } return d; }
      case 'yd': { const ys = NF.dateSerial(q.y, p.m, p.d); return ys <= t ? Math.floor(t) - ys : Math.floor(t) - NF.dateSerial(q.y - 1, p.m, p.d); }
    }
    return E.NUM;
  });

  // ------------------------------------------------------------------ 공개
  const cache = new Map();
  function compile(src) {
    if (cache.has(src)) return cache.get(src);
    let r;
    try { r = { ast: parse(src) }; } catch (e) { r = { error: e.message }; }
    cache.set(src, r);
    return r;
  }
  /* 수식(맨 앞 = 제외) 계산. ctx = {get(sheet,r,c), sheet, row, col} */
  function evaluate(src, ctx) {
    const c = compile(src);
    if (c.error) return E.NAME;
    let v;
    try { v = ev(c.ast, ctx); } catch (e) { return E.VALUE; }
    v = implicit(ctx, v);
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number' && !isFinite(v)) return E.NUM;
    return v;
  }
  /* 수식 안의 함수 이름 목록(채점용) */
  function functionsUsed(src) {
    return tokenize(src.startsWith('=') ? src.slice(1) : src).filter(t => t.k === 'fn').map(t => t.v);
  }
  function refsOf(src) {
    return tokenize(src.startsWith('=') ? src.slice(1) : src).filter(t => t.k === 'ref' || t.k === 'range');
  }
  function check(src) { const c = compile(src.startsWith('=') ? src.slice(1) : src); return c.error || null; }

  OA.formula = { tokenize, parse, evaluate, shift, cycleRef, functionsUsed, refsOf, check, isErr, E, toStr, cmp, makeCrit, FUNCTIONS: Object.keys(FN) };
})(typeof window !== 'undefined' ? window : globalThis);

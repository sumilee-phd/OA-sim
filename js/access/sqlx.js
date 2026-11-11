/* Access SQL → SQLite 변환과 sql.js 실행(DOM 없음).
   차이 처리: 큰따옴표 문자열, & 연결, Like 의 * ? , IIf·Nz·Format 등 VBA 함수(사용자 함수로 등록), 나눗셈은 실수, Mod,
   같은 SELECT 안의 앞 계산 필드 별칭 참조(식으로 펼침), True/False, #날짜#, 식 이름 없는 계산 열의 식1·식2, 매개 변수 입력. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  let SQL = null;

  class AccessError extends Error { constructor(msg, extra) { super(msg); this.access = true; Object.assign(this, extra || {}); } }

  const KW = new Set(['SELECT', 'FROM', 'WHERE', 'ORDER', 'BY', 'GROUP', 'HAVING', 'AS', 'AND', 'OR', 'NOT', 'INNER', 'LEFT', 'RIGHT', 'JOIN', 'ON', 'LIKE',
    'BETWEEN', 'IN', 'IS', 'NULL', 'ASC', 'DESC', 'DISTINCT', 'DISTINCTROW', 'TOP', 'MOD', 'TRUE', 'FALSE', 'YES', 'NO', 'OUTER', 'UNION', 'ALL']);

  function tokenize(s) {
    const t = [];
    let i = 0;
    while (i < s.length) {
      const ch = s[i];
      if (/\s/.test(ch)) { let j = i; while (j < s.length && /\s/.test(s[j])) j++; t.push({ k: 'ws', v: s.slice(i, j) }); i = j; continue; }
      if (ch === "'" || ch === '"') {
        let j = i + 1, v = '';
        while (j < s.length) { if (s[j] === ch) { if (s[j + 1] === ch) { v += ch; j += 2; continue; } break; } v += s[j]; j++; }
        if (j >= s.length) throw new AccessError('문자열이 닫히지 않았습니다. 따옴표를 확인하십시오.');
        t.push({ k: 'str', v, raw: s.slice(i, j + 1) }); i = j + 1; continue;
      }
      if (ch === '[') { const j = s.indexOf(']', i); if (j < 0) throw new AccessError('대괄호가 닫히지 않았습니다.'); t.push({ k: 'id', v: s.slice(i + 1, j), br: true, raw: s.slice(i, j + 1) }); i = j + 1; continue; }
      if (ch === '#') { const j = s.indexOf('#', i + 1); if (j > 0 && /^[\d\-/: ]+$/.test(s.slice(i + 1, j))) { t.push({ k: 'date', v: s.slice(i + 1, j), raw: s.slice(i, j + 1) }); i = j + 1; continue; } }
      let m = /^\d+(\.\d+)?([eE][+-]?\d+)?/.exec(s.slice(i)) || /^\.\d+/.exec(s.slice(i));
      if (m) { t.push({ k: 'num', v: m[0], raw: m[0] }); i += m[0].length; continue; }
      m = /^[A-Za-z_ㄱ-힝][A-Za-z0-9_ㄱ-힝]*/.exec(s.slice(i));
      if (m) { const up = m[0].toUpperCase(); t.push({ k: KW.has(up) ? 'kw' : 'id', v: m[0], up, raw: m[0] }); i += m[0].length; continue; }
      const two = s.slice(i, i + 2);
      if (two === '<=' || two === '>=' || two === '<>') { t.push({ k: 'op', v: two, raw: two }); i += 2; continue; }
      if ('=<>+-*/\\&^(),;.!%'.includes(ch)) { t.push({ k: 'op', v: ch, raw: ch }); i++; continue; }
      t.push({ k: 'bad', v: ch, raw: ch }); i++;
    }
    return t;
  }
  const isWs = t => t.k === 'ws';
  const textOf = toks => toks.map(t => t.raw !== undefined ? t.raw : t.v).join('');
  function splitTop(toks, pred) { // 깊이 0 에서 pred 인 토큰으로 나눔
    const out = [[]];
    let d = 0;
    for (const t of toks) {
      if (t.k === 'op' && t.v === '(') d++;
      if (t.k === 'op' && t.v === ')') d--;
      if (d === 0 && pred(t)) { out.push([]); continue; }
      out[out.length - 1].push(t);
    }
    return out;
  }
  function trimWs(toks) { let a = 0, b = toks.length; while (a < b && isWs(toks[a])) a++; while (b > a && isWs(toks[b - 1])) b--; return toks.slice(a, b); }

  /* SELECT 문을 절로 나눔: {select, from, where, group, having, order, distinct, top} */
  function clauses(toks) {
    const at = [];
    let d = 0;
    toks.forEach((t, i) => {
      if (t.k === 'op' && t.v === '(') d++;
      if (t.k === 'op' && t.v === ')') d--;
      if (d === 0 && t.k === 'kw' && ['SELECT', 'FROM', 'WHERE', 'HAVING'].includes(t.up)) at.push([t.up, i]);
      if (d === 0 && t.k === 'kw' && (t.up === 'ORDER' || t.up === 'GROUP')) {
        let j = i + 1; while (j < toks.length && isWs(toks[j])) j++;
        if (toks[j] && toks[j].up === 'BY') at.push([t.up, i]);
      }
    });
    const out = {};
    at.forEach(([name, i], k) => {
      const end = k + 1 < at.length ? at[k + 1][1] : toks.length;
      let body = toks.slice(i + 1, end);
      if (name === 'ORDER' || name === 'GROUP') { body = trimWs(body); body = body.slice(1); }
      out[name] = trimWs(body.filter((x, idx, arr) => !(idx === arr.length - 1 && x.k === 'op' && x.v === ';')));
    });
    if (out.SELECT) {
      let s = out.SELECT;
      if (s[0] && (s[0].up === 'DISTINCT' || s[0].up === 'DISTINCTROW')) { out.distinct = s[0].up === 'DISTINCT'; s = trimWs(s.slice(1)); }
      if (s[0] && s[0].up === 'TOP') { const n = s.find(x => x.k === 'num'); out.top = n ? +n.v : null; s = trimWs(s.slice(s.indexOf(n) + 1)); }
      out.SELECT = s;
    }
    for (const k of Object.keys(out)) if (Array.isArray(out[k]) && out[k].length && out[k][out[k].length - 1].k === 'op' && out[k][out[k].length - 1].v === ';') out[k] = trimWs(out[k].slice(0, -1));
    return out;
  }
  function selectItems(selToks) {
    return splitTop(selToks, t => t.k === 'op' && t.v === ',').map(trimWs).map(item => {
      // 마지막 깊이 0 의 AS
      let d = 0, asAt = -1;
      item.forEach((t, i) => { if (t.k === 'op' && t.v === '(') d++; if (t.k === 'op' && t.v === ')') d--; if (d === 0 && t.k === 'kw' && t.up === 'AS') asAt = i; });
      if (asAt >= 0) {
        const alias = trimWs(item.slice(asAt + 1));
        return { expr: trimWs(item.slice(0, asAt)), alias: alias.map(x => x.v).join('').trim() };
      }
      return { expr: item, alias: null };
    });
  }
  /* 표현식 토큰이 단순 열 참조이면 {table, field} */
  function simpleColumn(expr) {
    const e = expr.filter(t => !isWs(t));
    if (e.length === 1 && e[0].k === 'id') return { table: null, field: e[0].v };
    if (e.length === 3 && e[0].k === 'id' && e[1].k === 'op' && (e[1].v === '.' || e[1].v === '!') && (e[2].k === 'id' || (e[2].k === 'op' && e[2].v === '*'))) return { table: e[0].v, field: e[2].v };
    if (e.length === 3 && e[0].k === 'op' && e[0].v === '(' && e[2].k === 'op') return simpleColumn([e[1]]);
    const strip = e.length >= 2 && e[0].v === '(' && e[e.length - 1].v === ')' ? e.slice(1, -1) : null;
    if (strip) return simpleColumn(strip);
    return null;
  }

  /* 식 토큰 → SQLite 문자열 */
  function exprToSqlite(toks, ctx) {
    let out = '';
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      const prev = (() => { for (let j = i - 1; j >= 0; j--) if (!isWs(toks[j])) return toks[j]; return null; })();
      const nextT = (() => { for (let j = i + 1; j < toks.length; j++) if (!isWs(toks[j])) return toks[j]; return null; })();
      if (t.k === 'ws') { out += ' '; continue; }
      if (t.k === 'str') {
        let v = t.v;
        if (prev && prev.k === 'kw' && prev.up === 'LIKE') v = v.replace(/\*/g, '%').replace(/\?/g, '_');
        out += "'" + v.replace(/'/g, "''") + "'";
        continue;
      }
      if (t.k === 'date') { const m = /^(\d{1,4})[-/](\d{1,2})[-/](\d{1,4})/.exec(t.v); out += m ? (m[1].length === 4 ? `'${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}'` : `'${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}'`) : "'" + t.v + "'"; continue; }
      if (t.k === 'num') {
        if (nextT && nextT.k === 'op' && nextT.v === '%') throw new AccessError(`'${ctx.exprText}' 쿼리식의 구문 오류입니다.`, { syntax: true });
        out += t.v; continue;
      }
      if (t.k === 'op') {
        if (t.v === '&') { out += '||'; continue; }
        if (t.v === '/') { out += '*1.0/'; continue; }
        if (t.v === '\\') { out += '/'; continue; }
        if (t.v === '!') { out += '.'; continue; }
        if (t.v === '%') throw new AccessError(`'${ctx.exprText}' 쿼리식의 구문 오류입니다.`, { syntax: true });
        out += t.v; continue;
      }
      if (t.k === 'kw') {
        if (t.up === 'MOD') { out += '%'; continue; }
        if (t.up === 'TRUE' || t.up === 'YES') { out += '1'; continue; }
        if (t.up === 'FALSE' || t.up === 'NO') { out += '0'; continue; }
        out += t.up; continue;
      }
      if (t.k === 'id') {
        const afterDot = prev && prev.k === 'op' && (prev.v === '.' || prev.v === '!');
        const isFn = nextT && nextT.k === 'op' && nextT.v === '(';
        const nm = t.v;
        if (!afterDot && !isFn && ctx.params && Object.prototype.hasOwnProperty.call(ctx.params, nm)) { out += lit(ctx.params[nm]); continue; }
        if (!afterDot && !isFn && ctx.aliases && ctx.aliases[nm] && !ctx.columns.has(nm)) { out += '(' + ctx.aliases[nm] + ')'; continue; }
        const nextIsDot = nextT && nextT.k === 'op' && (nextT.v === '.' || nextT.v === '!');
        if (!afterDot && !isFn && !nextIsDot && ctx.fromTables && ctx.tableColumns) {
          const owners = ctx.fromTables.filter(tb => (ctx.tableColumns[tb] || []).includes(nm));
          if (owners.length > 1) {
            // 출력 열에 같은 이름의 열이 있으면 그 테이블로(Access 와 같게), 없으면 Access 오류
            const own = ctx.outCols && ctx.outCols[nm];
            if (own) { out += '[' + own + '].[' + nm + ']'; continue; }
            throw new AccessError(`지정한 필드 '[${nm}]'은(는) SQL 문의 FROM 절에 있는 여러 테이블을 참조할 수 있습니다.`);
          }
        }
        out += '[' + nm + ']';
        continue;
      }
      if (t.k === 'bad') throw new AccessError(`'${ctx.exprText}' 쿼리식의 구문 오류입니다.`, { syntax: true });
    }
    return out;
  }
  function lit(v) {
    if (v === null || v === undefined || v === '') return 'NULL';
    if (typeof v === 'number' || /^-?\d+(\.\d+)?$/.test(String(v))) return String(v);
    return "'" + String(v).replace(/'/g, "''") + "'";
  }

  /* Access SQL 전체 → {sql, items:[{name, column:{table,field}|null}]} */
  function translate(accessSql, opts) {
    const toks = tokenize(accessSql.trim());
    const cl = clauses(toks);
    if (!cl.SELECT || !cl.FROM) throw new AccessError('SQL 문이 잘못되었습니다. \'DELETE\', \'INSERT\', \'PROCEDURE\', \'SELECT\' 또는 \'UPDATE\'가 필요합니다.', { syntax: true });
    const columns = new Set(opts && opts.columns ? opts.columns : []);
    const items = selectItems(cl.SELECT);
    const aliases = {};
    const outItems = [];
    let auto = 1;
    // FROM 의 테이블 이름과 출력 열(모호한 열 이름 해석용)
    const ft = cl.FROM.filter(t => !isWs(t));
    const fromTables = ft.filter((t, i) => t.k === 'id' && !(ft[i + 1] && ft[i + 1].k === 'op' && (ft[i + 1].v === '.' || ft[i + 1].v === '!')) && !(ft[i - 1] && ft[i - 1].k === 'op' && (ft[i - 1].v === '.' || ft[i - 1].v === '!'))).map(t => t.v);
    const outCols = {};
    for (const it of items) { const sc = simpleColumn(it.expr); if (sc && sc.table && sc.field !== '*' && !it.alias && !outCols[sc.field]) outCols[sc.field] = sc.table; }
    const tableColumns = opts && opts.tableColumns;
    const parts = items.map(it => {
      const exprText = textOf(it.expr);
      const col = simpleColumn(it.expr);
      const ctx = { exprText, aliases, columns, params: opts && opts.params, fromTables, tableColumns, outCols };
      let sq = exprToSqlite(it.expr, ctx);
      let name = it.alias;
      if (!name) {
        if (col) name = col.field;
        else name = '식' + (auto++);
      }
      if (it.alias) aliases[it.alias] = sq;
      outItems.push({ name, column: col && col.field !== '*' ? col : null, star: col && col.field === '*' ? col.table : null, expr: exprText });
      return col && col.field === '*' ? sq : sq + ' AS [' + name + ']';
    });
    const ctxW = { aliases, columns, params: opts && opts.params, fromTables, tableColumns, outCols };
    let sql = 'SELECT ' + (cl.distinct ? 'DISTINCT ' : '') + parts.join(', ');
    sql += ' FROM ' + exprToSqlite(cl.FROM, Object.assign({ exprText: textOf(cl.FROM) }, ctxW, { aliases: {} }));
    if (cl.WHERE) sql += ' WHERE ' + exprToSqlite(cl.WHERE, Object.assign({ exprText: textOf(cl.WHERE) }, ctxW));
    if (cl.GROUP) sql += ' GROUP BY ' + exprToSqlite(cl.GROUP, Object.assign({ exprText: textOf(cl.GROUP) }, ctxW));
    if (cl.HAVING) sql += ' HAVING ' + exprToSqlite(cl.HAVING, Object.assign({ exprText: textOf(cl.HAVING) }, ctxW));
    if (cl.ORDER) sql += ' ORDER BY ' + exprToSqlite(cl.ORDER, Object.assign({ exprText: textOf(cl.ORDER) }, ctxW));
    if (cl.top) sql += ' LIMIT ' + cl.top;
    return { sql, items: outItems, clauses: cl };
  }

  // ------------------------------------------------------------------ 실행
  function setEngine(lib) { SQL = lib; }
  function ready() { return !!SQL; }
  const vbaDate = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
  function registerFunctions(db) {
    db.create_function('IIf', (c, a, b) => (c && c !== '0' ? a : b));
    db.create_function('Nz', (v, d) => (v === null || v === undefined ? (d === undefined || d === null ? 0 : d) : v));
    db.create_function('Left', (s, n) => s === null ? null : String(s).slice(0, n));
    db.create_function('Right', (s, n) => s === null ? null : (n <= 0 ? '' : String(s).slice(-n)));
    db.create_function('Mid', (s, a, n) => s === null ? null : (n === undefined || n === null ? String(s).slice(a - 1) : String(s).substr(a - 1, n)));
    db.create_function('Len', s => s === null ? null : String(s).length);
    db.create_function('Int', x => x === null ? null : Math.floor(x));
    db.create_function('Val', s => { const m = /^\s*-?\d+(\.\d+)?/.exec(String(s)); return m ? +m[0] : 0; });
    db.create_function('UCase', s => s === null ? null : String(s).toUpperCase());
    db.create_function('LCase', s => s === null ? null : String(s).toLowerCase());
    db.create_function('Date', () => vbaDate());
    db.create_function('Now', () => new Date().toISOString().slice(0, 19).replace('T', ' '));
    db.create_function('Year', s => s === null ? null : +String(s).slice(0, 4));
    db.create_function('Month', s => s === null ? null : +String(s).slice(5, 7));
    db.create_function('Day', s => s === null ? null : +String(s).slice(8, 10));
    db.create_function('Format', (v, f) => (OA.accExpr ? OA.accExpr.formatValue(v, f) : String(v)));
  }
  const sqlType = f => ({ '짧은 텍스트': 'TEXT', '긴 텍스트': 'TEXT', '숫자': (/실수/.test(f.size || '') ? 'REAL' : 'INTEGER'), '큰 번호': 'INTEGER',
    '통화': 'REAL', '일련 번호': 'INTEGER', 'Yes/No': 'INTEGER', '날짜/시간': 'TEXT', '하이퍼링크': 'TEXT' })[f.type] || 'TEXT';
  function buildDb(acc) {
    if (!SQL) throw new AccessError('SQL 엔진을 불러오는 중입니다. 잠시 뒤 다시 실행하십시오.');
    const db = new SQL.Database();
    registerFunctions(db);
    for (const tn of Object.keys(acc.tables)) {
      const t = acc.tables[tn];
      if (!t.fields.length) continue;
      db.run(`CREATE TABLE [${tn}] (${t.fields.map(f => `[${f.name}] ${sqlType(f)}`).join(', ')})`);
      const st = db.prepare(`INSERT INTO [${tn}] VALUES (${t.fields.map(() => '?').join(', ')})`);
      for (const row of t.rows) st.run(t.fields.map(f => row[f.name] === undefined ? null : row[f.name]));
      st.free();
    }
    return db;
  }
  /* 실행 결과 {columns:[이름], rows:[[값]], items} 또는 {needParam: 이름} 또는 throw AccessError */
  function run(acc, accessSql, params) {
    const columns = [], tableColumns = {};
    for (const tn of Object.keys(acc.tables)) { tableColumns[tn] = acc.tables[tn].fields.map(f => f.name); columns.push(...tableColumns[tn]); }
    const tr = translate(accessSql, { columns, params: params || {}, tableColumns });
    const db = buildDb(acc);
    try {
      const r = db.exec(tr.sql);
      const res = r[0] || { columns: tr.items.filter(x => !x.star).map(x => x.name), values: [] };
      return { columns: res.columns, rows: res.values, items: tr.items, sql: tr.sql };
    } catch (e) {
      const m = /no such column: (.+)$/.exec(e.message);
      if (m) return { needParam: m[1].replace(/^\[|\]$/g, '').replace(/\]\.\[/g, '.') };
      const t = /no such table: (.+)$/.exec(e.message);
      if (t) throw new AccessError(`Microsoft Access 데이터베이스 엔진에서 입력 테이블 또는 쿼리 '${t[1]}'을(를) 찾을 수 없습니다. 해당 개체가 있는지, 이름의 철자가 올바른지 확인하십시오.`);
      throw new AccessError('쿼리 식의 구문 오류입니다. (' + e.message.replace(/^near "(.*)": syntax error$/, "'$1' 근처") + ')', { syntax: true });
    } finally { db.close(); }
  }

  OA.sqlx = { tokenize, clauses, selectItems, simpleColumn, translate, run, setEngine, ready, AccessError, textOf, trimWs, splitTop, isWs, lit, KW };
})(typeof window !== 'undefined' ? window : globalThis);

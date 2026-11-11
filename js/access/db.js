/* Access 데이터베이스 모델(DOM 없음): 테이블·필드 속성·기본 키·레코드 검사·관계(참조 무결성)·쿼리 디자인 눈금 ⇄ SQL·폼·보고서 마법사. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util;
  const X = () => OA.sqlx;

  const TYPES = ['짧은 텍스트', '긴 텍스트', '숫자', '큰 번호', '날짜/시간', '통화', '일련 번호', 'Yes/No', 'OLE 개체', '하이퍼링크', '첨부 파일', '계산'];
  const NUM_SIZES = ['바이트', '정수', '정수(Long)', '실수(single)', '실수(double)', '복제 ID', '10진수'];
  const FORMATS = { '숫자': ['', '일반 숫자', '통화', '유로', '고정', '표준', '백분율', '지수'], '통화': ['', '일반 숫자', '통화', '유로', '고정', '표준', '백분율', '지수'], '날짜/시간': ['', '기본 날짜', '자세한 날짜', '보통 날짜', '간단한 날짜', '자세한 시간', '보통 시간', '간단한 시간'] };

  function newDatabase(name) {
    return { name: name || 'Database1', folder: '바탕 화면', tables: {}, queries: {}, forms: {}, reports: {}, relationships: [], relLayout: null, order: [], seq: { table: 1 } };
  }
  function defaultField(name, type) {
    const f = { name, type: type || '짧은 텍스트', desc: '', format: '', caption: '', defaultValue: '', required: false, indexed: '아니요' };
    applyTypeDefaults(f);
    return f;
  }
  function applyTypeDefaults(f) {
    if (f.type === '짧은 텍스트') { if (typeof f.size !== 'number') f.size = 255; f.allowZero = true; }
    else if (f.type === '숫자') { if (!NUM_SIZES.includes(f.size)) f.size = '정수(Long)'; if (f.decimals === undefined) f.decimals = '자동'; if (f.defaultValue === '') f.defaultValue = '0'; }
    else if (f.type === '통화') { delete f.size; if (!f.format) f.format = '통화'; if (f.decimals === undefined) f.decimals = '자동'; if (f.defaultValue === '') f.defaultValue = '0'; }
    else if (f.type === '일련 번호') { f.size = '정수(Long)'; f.indexed = f.indexed === '아니요' ? '예(중복 불가능)' : f.indexed; }
    else delete f.size;
    return f;
  }
  function createTable(d, name, fields, pk) {
    const t = { name, fields: fields.map(f => Object.assign(defaultField(f.name, f.type), f)), pk: pk ? pk.slice() : [], rows: [] };
    t.fields.forEach(applyTypeDefaults);
    for (const k of t.pk) { const f = t.fields.find(x => x.name === k); if (f) { f.required = true; f.indexed = '예(중복 불가능)'; } }
    d.tables[name] = t;
    if (!d.order.includes('t:' + name)) d.order.push('t:' + name);
    return t;
  }
  function setPrimaryKey(t, names) {
    for (const f of t.fields) if (t.pk.includes(f.name) && !names.includes(f.name)) { f.indexed = '아니요'; f.required = false; }
    t.pk = names.slice();
    for (const k of names) { const f = t.fields.find(x => x.name === k); if (f) { f.required = true; f.indexed = '예(중복 불가능)'; } }
  }

  // ------------------------------------------------------------------ 레코드
  function coerce(f, raw) {
    if (raw === null || raw === undefined || raw === '') return { v: null };
    const s = String(raw).trim();
    switch (f.type) {
      case '숫자': case '큰 번호': case '일련 번호': case '통화': {
        const n = Number(s.replace(/[₩,\s]/g, ''));
        if (!isFinite(n) || s.replace(/[₩,\s]/g, '') === '') return { error: '입력한 값이 이 필드에 대해 잘못되었습니다.\n\n예를 들어 숫자 필드에 텍스트를 입력했습니다.' };
        if (f.type === '숫자' && /정수|바이트/.test(f.size || '') && !Number.isInteger(n)) return { v: Math.round(n) };
        if (f.size === '바이트' && (n < 0 || n > 255)) return { error: '입력한 값이 이 필드에 대해 잘못되었습니다.' };
        return { v: n };
      }
      case 'Yes/No': return { v: /^(예|yes|true|-1|1)$/i.test(s) ? -1 : 0 };
      case '짧은 텍스트': if (s.length > (f.size || 255)) return { error: `입력한 값이 필드 크기(${f.size})보다 깁니다.` }; return { v: s };
      default: return { v: s };
    }
  }
  function checkRecord(d, tn, rec, skipIndex) {
    const t = d.tables[tn];
    for (const f of t.fields) {
      if ((f.required || t.pk.includes(f.name)) && (rec[f.name] === null || rec[f.name] === undefined || rec[f.name] === '')) {
        if (t.pk.includes(f.name)) return { error: '인덱스 또는 기본 키는 Null 값을 포함할 수 없습니다.', field: f.name };
        return { error: `필드 '${tn}.${f.name}'은(는) Null 값을 포함할 수 없습니다. 이 필드의 Required 속성이 True로 설정되어 있습니다. 이 필드에 값을 입력하십시오.`, field: f.name };
      }
    }
    if (t.pk.length) {
      const dup = t.rows.some((r, i) => i !== skipIndex && t.pk.every(k => String(r[k]) === String(rec[k])));
      if (dup) return { error: '요청한 내용을 테이블에 적용하지 못했습니다. 인덱스, 기본 키 또는 관계에서 값이 중복되었습니다. 중복된 값이 있는 필드의 데이터를 변경하거나, 인덱스를 제거하거나, 중복된 값이 허용되도록 인덱스를 다시 정의한 후 다시 시도하십시오.' };
    }
    for (const rel of d.relationships) {
      if (!rel.ri || rel.t2 !== tn) continue;
      const v = rec[rel.f2];
      if (v === null || v === undefined || v === '') continue;
      if (!d.tables[rel.t1].rows.some(r => String(r[rel.f1]) === String(v))) return { error: `'${rel.t1}' 테이블에 관련 레코드가 있어야 하므로 레코드를 추가하거나 변경할 수 없습니다.` };
    }
    return {};
  }
  function insertRow(d, tn, rec) {
    const t = d.tables[tn];
    const row = {};
    for (const f of t.fields) {
      const c = coerce(f, rec[f.name]);
      if (c.error) return { error: c.error, field: f.name };
      row[f.name] = c.v;
    }
    const chk = checkRecord(d, tn, row, -1);
    if (chk.error) return chk;
    t.rows.push(row);
    return { row };
  }
  function updateRow(d, tn, idx, rec) {
    const t = d.tables[tn];
    const row = {};
    for (const f of t.fields) {
      const c = coerce(f, rec[f.name]);
      if (c.error) return { error: c.error, field: f.name };
      row[f.name] = c.v;
    }
    const chk = checkRecord(d, tn, row, idx);
    if (chk.error) return chk;
    // 참조되는 쪽 값 변경 검사
    for (const rel of d.relationships) if (rel.ri && rel.t1 === tn && String(t.rows[idx][rel.f1]) !== String(row[rel.f1])) {
      if (d.tables[rel.t2].rows.some(r => String(r[rel.f2]) === String(t.rows[idx][rel.f1]))) return { error: `'${rel.t2}' 테이블에 관련 레코드가 있으므로 레코드를 삭제하거나 변경할 수 없습니다.` };
    }
    t.rows[idx] = row;
    return { row };
  }
  function deleteRow(d, tn, idx) {
    const t = d.tables[tn];
    for (const rel of d.relationships) if (rel.ri && rel.t1 === tn) {
      if (d.tables[rel.t2].rows.some(r => String(r[rel.f2]) === String(t.rows[idx][rel.f1]))) return { error: `'${rel.t2}' 테이블에 관련 레코드가 있으므로 레코드를 삭제하거나 변경할 수 없습니다.` };
    }
    t.rows.splice(idx, 1);
    return {};
  }

  // ------------------------------------------------------------------ 관계
  function isUnique(t, f) { return (t.pk.length === 1 && t.pk[0] === f) || (t.fields.find(x => x.name === f) || {}).indexed === '예(중복 불가능)'; }
  function addRelationship(d, rel) {
    let { t1, f1, t2, f2 } = rel;
    const A = d.tables[t1], B = d.tables[t2];
    if (!A || !B) return { error: '테이블을 찾을 수 없습니다.' };
    if (!isUnique(A, f1) && isUnique(B, f2)) { [t1, f1, t2, f2] = [t2, f2, t1, f1]; }
    const a = d.tables[t1], b = d.tables[t2];
    const fa = a.fields.find(x => x.name === f1), fb = b.fields.find(x => x.name === f2);
    const typeKey = f => f.type === '일련 번호' ? '숫자:정수(Long)' : f.type === '숫자' ? '숫자:' + f.size : f.type;
    const kind = isUnique(a, f1) ? (isUnique(b, f2) ? '일대일' : '일대다') : '정의되지 않음';
    if (rel.ri) {
      if (kind === '정의되지 않음') return { error: '이 관계에 대해 참조 무결성을 적용할 수 없습니다. 기본 테이블의 관련 필드가 기본 키이거나 고유 인덱스를 가지고 있어야 합니다.' };
      if (typeKey(fa) !== typeKey(fb)) return { error: '관계를 만들 수 없으며 참조 무결성을 적용할 수 없습니다. 관련 필드의 데이터 형식이 다릅니다.' };
      const keys = new Set(a.rows.map(r => String(r[f1])));
      if (b.rows.some(r => r[f2] !== null && r[f2] !== undefined && r[f2] !== '' && !keys.has(String(r[f2])))) return { error: `참조 무결성을 적용할 수 없습니다. '${t2}' 테이블의 데이터가 참조 무결성 규칙을 위반합니다.\n\n예를 들어 관련 테이블에 직원에 해당하는 레코드가 있지만 기본 테이블에는 해당 레코드가 없습니다.` };
    }
    d.relationships = d.relationships.filter(r => !(r.t1 === t1 && r.f1 === f1 && r.t2 === t2 && r.f2 === f2));
    const r = { t1, f1, t2, f2, ri: !!rel.ri, cascadeUpd: !!rel.cascadeUpd, cascadeDel: !!rel.cascadeDel, kind, join: rel.join || 1 };
    d.relationships.push(r);
    return { rel: r };
  }
  function removeRelationship(d, r) { d.relationships = d.relationships.filter(x => x !== r); }

  // ------------------------------------------------------------------ 쿼리 디자인 ⇄ SQL
  function newQueryDesign(d, tables) {
    const q = { mode: 'design', tables: [], joins: [], cols: [] };
    (tables || []).forEach(t => addQueryTable(d, q, t));
    return q;
  }
  function addQueryTable(d, q, tn) {
    if (q.tables.some(t => t.name === tn)) return;
    q.tables.push({ name: tn, x: 20 + q.tables.length * 220, y: 16 });
    for (const r of d.relationships) {
      const names = q.tables.map(t => t.name);
      if (names.includes(r.t1) && names.includes(r.t2) && !q.joins.some(j => j.t1 === r.t1 && j.t2 === r.t2 && j.f1 === r.f1)) q.joins.push({ t1: r.t1, f1: r.f1, t2: r.t2, f2: r.f2, kind: 'INNER' });
    }
  }
  function fieldTable(d, q, field) {
    for (const t of q.tables) { const src = d.tables[t.name] || d.queries[t.name]; if (src && objFields(d, t.name).includes(field)) return t.name; }
    return '';
  }
  function objFields(d, name) {
    if (d.tables[name]) return d.tables[name].fields.map(f => f.name);
    if (d.queries[name]) { try { const r = runSaved(d, name); return r.columns || []; } catch (e) { return []; } }
    return [];
  }
  function parseCalc(field) {
    const m = /^\s*([^:[\]]+?)\s*:\s*([\s\S]+)$/.exec(field || '');
    if (m && !/^\s*\[/.test(field)) return { alias: m[1].trim(), expr: m[2].trim() };
    return null;
  }
  const qualify = (table, field) => table ? `${table}.${field}` : field;

  /* 조건 칸 문자열 → Access 조건 식(필드 식 f 를 붙인다) */
  function critToSql(text, f) {
    const T = X();
    const toks = T.tokenize(String(text).trim());
    // 깊이 0 의 Or / And 로 나눔(Between ... And 는 하나로)
    const atoms = [[]], joins = [];
    let d = 0, between = false;
    for (const t of toks) {
      if (t.k === 'op' && t.v === '(') d++;
      if (t.k === 'op' && t.v === ')') d--;
      if (d === 0 && t.k === 'kw' && t.up === 'BETWEEN') between = true;
      if (d === 0 && t.k === 'kw' && (t.up === 'OR' || t.up === 'AND')) {
        if (t.up === 'AND' && between) { between = false; atoms[atoms.length - 1].push(t); continue; }
        joins.push(t.up === 'OR' ? 'Or' : 'And'); atoms.push([]); continue;
      }
      atoms[atoms.length - 1].push(t);
    }
    const conds = atoms.map(a => {
      const tk = T.trimWs(a);
      if (!tk.length) return null;
      const first = tk[0];
      const rest = T.textOf(tk.slice(1)).trim();
      const q = tok => (tok.k === 'id' && !tok.br && tk.length === 1) ? `"${tok.v}"` : T.textOf([tok]);
      if (first.k === 'op' && ['=', '<', '>', '<=', '>=', '<>'].includes(first.v)) {
        const r = T.trimWs(tk.slice(1));
        const val = r.length === 1 && r[0].k === 'id' && !r[0].br ? `"${r[0].v}"` : T.textOf(r).trim();
        return `(${f})${first.v}${val}`;
      }
      if (first.k === 'kw' && ['LIKE', 'BETWEEN', 'IN', 'IS'].includes(first.up)) return `(${f}) ${first.v[0].toUpperCase() + first.v.slice(1).toLowerCase()} ${rest}`;
      if (first.k === 'kw' && first.up === 'NOT') {
        const r = T.trimWs(tk.slice(1));
        if (r[0] && r[0].k === 'kw' && ['LIKE', 'IN', 'BETWEEN'].includes(r[0].up)) return `(${f}) Not ${T.textOf(r).trim()}`;
        return `Not (${f})=${r.length === 1 ? q(r[0]) : T.textOf(r).trim()}`;
      }
      if (tk.length === 1) return `(${f})=${q(first)}`;
      return `(${f})=${T.textOf(tk).trim()}`;
    }).filter(Boolean);
    if (!conds.length) return null;
    let s = conds[0];
    for (let i = 1; i < conds.length; i++) s += ' ' + (joins[i - 1] || 'Or') + ' ' + conds[i];
    return '(' + s + ')';
  }
  function designToSQL(d, q) {
    const sel = [], order = [], rowsCount = Math.max(0, ...q.cols.map(c => (c.crit || []).length));
    const condRows = [...Array(rowsCount)].map(() => []);
    for (const c of q.cols) {
      if (!c.field) continue;
      const calc = parseCalc(c.field);
      let expr, item;
      if (calc) { expr = calc.expr; item = `${calc.expr} AS ${calc.alias}`; }
      else if (c.field === '*' || /\.\*$/.test(c.field)) { expr = null; item = c.field === '*' ? `${c.table || q.tables[0].name}.*` : c.field; }
      else { const tb = c.table || fieldTable(d, q, c.field); expr = qualify(tb, c.field); item = expr; }
      if (c.show !== false) sel.push(item);
      if (c.sort === '오름차순') order.push(expr || item);
      if (c.sort === '내림차순') order.push((expr || item) + ' DESC');
      (c.crit || []).forEach((t, i) => { if (t && String(t).trim()) { const s = critToSql(t, expr || item); if (s) condRows[i].push(s); } });
    }
    if (!sel.length) throw new X().AccessError('쿼리 출력 대상으로 선택된 필드가 적어도 하나 있어야 합니다.');
    let sql = 'SELECT ' + sel.join(', ') + '\nFROM ' + fromClause(d, q);
    const rows = condRows.filter(r => r.length).map(r => '(' + r.join(' AND ') + ')');
    if (rows.length === 1) sql += '\nWHERE ' + rows[0];
    else if (rows.length > 1) sql += '\nWHERE ' + rows.map(r => '(' + r + ')').join(' OR ');
    if (order.length) sql += '\nORDER BY ' + order.join(', ');
    return sql + ';';
  }
  function fromClause(d, q) {
    const names = q.tables.map(t => t.name);
    if (!names.length) return '';
    const used = new Set();
    let s = '';
    const joins = q.joins.filter(j => names.includes(j.t1) && names.includes(j.t2));
    if (joins.length) {
      const j0 = joins[0];
      s = `${j0.t1} ${j0.kind || 'INNER'} JOIN ${j0.t2} ON ${j0.t1}.${j0.f1} = ${j0.t2}.${j0.f2}`;
      used.add(j0.t1); used.add(j0.t2);
      for (const j of joins.slice(1)) {
        const other = used.has(j.t1) ? j.t2 : j.t1;
        if (used.has(j.t1) && used.has(j.t2)) { s = `(${s}) AND ${j.t1}.${j.f1} = ${j.t2}.${j.f2}`; continue; }
        s = `(${s}) ${j.kind || 'INNER'} JOIN ${other} ON ${j.t1}.${j.f1} = ${j.t2}.${j.f2}`;
        used.add(other);
      }
    }
    const rest = names.filter(n => !used.has(n));
    return [s].concat(rest).filter(Boolean).join(', ');
  }

  /* SQL → 디자인 눈금(가능한 범위). 실패하면 null */
  function sqlToDesign(d, sql) {
    const T = X();
    let toks;
    try { toks = T.tokenize(sql); } catch (e) { return null; }
    const cl = T.clauses(toks);
    if (!cl.SELECT || !cl.FROM) return null;
    const q = newQueryDesign(d, []);
    // FROM: 테이블 이름과 조인
    const fromToks = cl.FROM.filter(t => !T.isWs(t));
    const tables = [];
    for (let i = 0; i < fromToks.length; i++) {
      const t = fromToks[i], prev = fromToks[i - 1], next = fromToks[i + 1];
      if (t.k === 'id' && !(next && next.k === 'op' && (next.v === '.' || next.v === '!')) && !(prev && prev.k === 'op' && (prev.v === '.' || prev.v === '!'))) tables.push(t.v);
    }
    tables.forEach(tn => { if (!q.tables.some(x => x.name === tn)) q.tables.push({ name: tn, x: 20 + q.tables.length * 220, y: 16 }); });
    const onRe = /([^\s().=]+)\.([^\s().=]+)\s*=\s*([^\s().=]+)\.([^\s().=]+)/g;
    let m;
    const fromText = T.textOf(cl.FROM);
    while ((m = onRe.exec(fromText))) q.joins.push({ t1: m[1], f1: m[2], t2: m[3], f2: m[4], kind: /LEFT JOIN/i.test(fromText) ? 'LEFT' : /RIGHT JOIN/i.test(fromText) ? 'RIGHT' : 'INNER' });
    for (const it of T.selectItems(cl.SELECT)) {
      const col = T.simpleColumn(it.expr);
      if (col && !it.alias) q.cols.push({ field: col.field === '*' ? `${col.table}.*` : col.field, table: col.field === '*' ? col.table : (col.table || fieldTable(d, q, col.field)), sort: '', show: true, crit: [] });
      else q.cols.push({ field: `${it.alias || '식' + (q.cols.length + 1)}: ${T.textOf(it.expr).trim()}`, table: '', sort: '', show: true, crit: [] });
    }
    const findCol = (table, field) => {
      let c = q.cols.find(x => x.field === field && (!table || x.table === table));
      if (!c) { c = { field, table: table || fieldTable(d, q, field), sort: '', show: false, crit: [] }; q.cols.push(c); }
      return c;
    };
    if (cl.WHERE) {
      const conj = T.splitTop(cl.WHERE, t => t.k === 'kw' && t.up === 'AND').map(T.trimWs);
      for (let part of conj) {
        while (part.length && part[0].v === '(' && part[part.length - 1].v === ')' && balanced(part.slice(1, -1))) part = T.trimWs(part.slice(1, -1));
        const disj = T.splitTop(part, t => t.k === 'kw' && t.up === 'OR').map(T.trimWs);
        let target = null; const vals = [];
        for (let a of disj) {
          while (a.length && a[0].v === '(' && a[a.length - 1].v === ')' && balanced(a.slice(1, -1))) a = T.trimWs(a.slice(1, -1));
          const opIdx = a.findIndex(t => (t.k === 'op' && ['=', '<', '>', '<=', '>=', '<>'].includes(t.v)) || (t.k === 'kw' && ['LIKE', 'BETWEEN', 'IN', 'IS'].includes(t.up)));
          if (opIdx < 1) return Object.assign(q, { partial: true });
          const lhs = T.simpleColumn(T.trimWs(a.slice(0, opIdx)));
          if (!lhs) return Object.assign(q, { partial: true });
          const k = (lhs.table || '') + '.' + lhs.field;
          if (target && target.k !== k) return Object.assign(q, { partial: true });
          target = { k, lhs };
          const op = a[opIdx];
          const rhs = T.textOf(T.trimWs(a.slice(opIdx + 1))).trim();
          vals.push(op.k === 'op' && op.v === '=' ? rhs : (op.k === 'op' ? op.v + rhs : op.v[0].toUpperCase() + op.v.slice(1).toLowerCase() + ' ' + rhs));
        }
        if (target) { const c = findCol(target.lhs.table, target.lhs.field); c.crit[0] = (c.crit[0] ? c.crit[0] + ' And ' : '') + vals.join(' Or '); }
      }
    }
    if (cl.ORDER) {
      for (const it of T.splitTop(cl.ORDER, t => t.k === 'op' && t.v === ',').map(T.trimWs)) {
        const desc = it.length && it[it.length - 1].up === 'DESC';
        const core = T.trimWs(it.filter(t => !(t.k === 'kw' && (t.up === 'ASC' || t.up === 'DESC'))));
        const col = T.simpleColumn(core);
        if (!col) { q.partial = true; continue; }
        findCol(col.table, col.field).sort = desc ? '내림차순' : '오름차순';
      }
    }
    return q;
  }
  function balanced(toks) { let d = 0; for (const t of toks) { if (t.v === '(') d++; if (t.v === ')') { d--; if (d < 0) return false; } } return d === 0; }

  function saveQuery(d, name, q) {
    const copy = U.clone(q);
    if (copy.mode !== 'sql') copy.sql = designToSQL(d, copy);
    d.queries[name] = Object.assign({ name }, copy);
    if (!d.order.includes('q:' + name)) d.order.push('q:' + name);
    return d.queries[name];
  }
  function querySQL(d, name) {
    const q = d.queries[name];
    if (!q) return '';
    return q.mode === 'sql' ? q.sql : designToSQL(d, q);
  }
  /* 테이블·저장된 쿼리를 하위 쿼리로 펼친 실행용 DB 구성 */
  function runSQL(d, sql, params) {
    const acc = { tables: Object.assign({}, d.tables) };
    // 저장된 쿼리를 테이블처럼 쓰는 경우: 결과를 임시 테이블로 만든다
    const names = Object.keys(d.queries);
    for (const n of names) {
      if (!new RegExp('(^|[\\s,(\\[])' + n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '([\\s,)\\].;]|$)').test(sql)) continue;
      if (acc.tables[n]) continue;
      const r = runSaved(d, n, params);
      if (r.needParam || !r.columns) continue;
      acc.tables[n] = { name: n, fields: r.columns.map((c, i) => ({ name: c, type: (r.fieldInfo && r.fieldInfo[i] && r.fieldInfo[i].type) || guessType(r.rows, i) })), rows: r.rows.map(row => Object.fromEntries(r.columns.map((c, i) => [c, row[i]]))) };
    }
    const res = X().run(acc, sql, params);
    if (res.columns) res.fieldInfo = res.items ? fieldInfo(d, res) : [];
    return res;
  }
  function guessType(rows, i) { return rows.every(r => r[i] === null || typeof r[i] === 'number') ? '숫자' : '짧은 텍스트'; }
  function fieldInfo(d, res) {
    const items = res.items.filter(it => !it.star);
    return res.columns.map((name, i) => {
      const it = items[i];
      if (it && it.column) {
        const tn = it.column.table || Object.keys(d.tables).find(t => d.tables[t].fields.some(f => f.name === it.column.field));
        const t = d.tables[tn];
        const f = t && t.fields.find(x => x.name === it.column.field);
        if (f) return { type: f.type, format: f.format, decimals: f.decimals };
        const qq = d.queries[tn];
        if (qq) { try { const r = runSaved(d, tn); const k = r.columns.indexOf(it.column.field); if (k >= 0 && r.fieldInfo) return r.fieldInfo[k]; } catch (e) { /* 무시 */ } }
      }
      for (const tn of Object.keys(d.tables)) { const f = d.tables[tn].fields.find(x => x.name === name); if (f && it && it.column) return { type: f.type, format: f.format, decimals: f.decimals }; }
      return { type: null, format: '', decimals: '자동' };
    });
  }
  function runSaved(d, name, params) {
    if (d.tables[name]) { const t = d.tables[name]; return { columns: t.fields.map(f => f.name), rows: t.rows.map(r => t.fields.map(f => r[f.name])), fieldInfo: t.fields.map(f => ({ type: f.type, format: f.format, decimals: f.decimals })) }; }
    const sql = querySQL(d, name);
    return runSQL(d, sql, params);
  }

  /* 값 표시(필드 형식) */
  function formatValue(v, info) {
    if (v === null || v === undefined) return '';
    const fmt = (info && info.format) || '';
    if (typeof v === 'number') {
      const dec = info && info.decimals !== undefined && info.decimals !== '자동' ? +info.decimals : null;
      if (fmt === '통화' || fmt === 'Currency') return (v < 0 ? '-' : '') + '₩' + U.thousands(U.roundHalfAway(Math.abs(v), dec === null ? 0 : dec).toFixed(dec === null ? 0 : dec));
      if (fmt === '표준' || fmt === 'Standard') { const k = dec === null ? 2 : dec; return (v < 0 ? '-' : '') + U.thousands(U.roundHalfAway(Math.abs(v), k).toFixed(k)); }
      if (fmt === '고정' || fmt === 'Fixed') { const k = dec === null ? 2 : dec; return U.roundHalfAway(v, k).toFixed(k); }
      if (fmt === '백분율' || fmt === 'Percent') { const k = dec === null ? 2 : dec; return U.roundHalfAway(v * 100, k).toFixed(k) + '%'; }
      if (fmt && /[0#]/.test(fmt)) return OA.numfmt.format(v, fmt).text;
      if (dec !== null && fmt === '일반 숫자') return U.roundHalfAway(v, dec).toFixed(dec);
      return OA.numfmt.generalNum(v);
    }
    if (v === true) return '-1';
    if (v === false) return '0';
    return String(v);
  }

  // ------------------------------------------------------------------ 폼
  const NAME_PREFIX = { label: '레이블', textbox: '텍스트', listbox: '목록', combobox: '콤보', line: '선', rect: '상자', button: '명령' };
  function newForm(d) {
    return { kind: 'form', props: { caption: '', width: 13.6, recordSource: '', recordSelectors: true, navButtons: true, dividingLines: false, scrollBars: '모두', defaultView: '단일 폼' },
      sections: { detail: { height: 6.35, backColor: '#FFFFFF' } }, controls: [], page: defaultPage(), seq: 0 };
  }
  function defaultPage() { return { top: 6.35, bottom: 6.35, left: 6.35, right: 6.35, orient: 'portrait', paper: 'A4', dataOnly: false, cols: 1 }; }
  function controlDefaults(type) {
    const base = { fontName: '맑은 고딕', fontSize: 11, fontWeight: '표준', italic: false, underline: false, textAlign: '일반', foreColor: '#000000',
      backStyle: '투명', backColor: '#FFFFFF', borderStyle: '투명', borderWidth: '가는 선', borderColor: '#000000', visible: true };
    if (type === 'textbox') Object.assign(base, { controlSource: '', format: '', decimals: '자동', hideDuplicates: false, borderStyle: '실선', backStyle: '보통', canGrow: false, runningSum: '아니요' });
    if (type === 'listbox') Object.assign(base, { controlSource: '', rowSourceType: '테이블/쿼리', rowSource: '', boundColumn: 1, columnCount: 1, columnHeads: false, columnWidths: '', borderStyle: '실선', backStyle: '보통' });
    if (type === 'line') Object.assign(base, { borderStyle: '실선', borderWidth: '가는 선' });
    if (type === 'rect') Object.assign(base, { borderStyle: '실선' });
    if (type === 'label') Object.assign(base, { caption: '' });
    return base;
  }
  function addControl(f, type, props) {
    const p = Object.assign({}, props || {});
    const withLabel = p.withLabel !== undefined ? p.withLabel : (type === 'listbox' || type === 'textbox' || type === 'combobox');
    delete p.withLabel;
    const c = Object.assign({ id: U.uid('c'), type, section: 'detail', left: 1, top: 1, width: 3, height: 0.6 }, controlDefaults(type));
    if (!p.name) c.name = NAME_PREFIX[type] + (f.seq++);
    Object.assign(c, p);
    f.controls.push(c);
    let lab = null;
    if (withLabel) {
      lab = Object.assign({ id: U.uid('c'), type: 'label', section: c.section }, controlDefaults('label'), {
        name: NAME_PREFIX.label + (f.seq++), caption: c.name, left: Math.max(0, c.left - 2.5), top: c.top, width: 2.3, height: 0.6, attachedTo: c.id });
      f.controls.push(lab);
    }
    growSection(f, c);
    return c;
  }
  function growSection(f, c) {
    const s = f.sections[c.section];
    if (s && c.top + c.height > s.height) s.height = Math.ceil((c.top + c.height) * 10) / 10;
  }
  function saveForm(d, name, f) { f.name = name; d.forms[name] = U.clone(f); if (!d.order.includes('f:' + name)) d.order.push('f:' + name); return d.forms[name]; }
  function saveReport(d, name, r) { r.name = name; d.reports[name] = U.clone(r); if (!d.order.includes('r:' + name)) d.order.push('r:' + name); return d.reports[name]; }

  // ------------------------------------------------------------------ 보고서 마법사
  const SEC_ORDER = ['reportHeader', 'pageHeader', 'groupHeader0', 'groupHeader1', 'detail', 'groupFooter1', 'groupFooter0', 'pageFooter', 'reportFooter'];
  function sectionTitle(r, id) {
    const g = id.match(/^group(Header|Footer)(\d)$/);
    if (g) { const lv = (r.grouping.filter(x => x.kind === 'group')[+g[2]] || {}).field || '그룹'; return lv + (g[1] === 'Header' ? ' 머리글' : ' 바닥글'); }
    return { reportHeader: '보고서 머리글', pageHeader: '페이지 머리글', detail: '본문', pageFooter: '페이지 바닥글', reportFooter: '보고서 바닥글' }[id];
  }
  function newReport(d, source) {
    return { kind: 'report', props: { recordSource: source || '', width: 16, caption: '' }, grouping: [], page: defaultPage(), seq: 0,
      sections: { reportHeader: { height: 1.5, backColor: '#FFFFFF', altBackColor: '#FFFFFF' }, pageHeader: { height: 0.8, backColor: '#FFFFFF', altBackColor: '#FFFFFF' },
        detail: { height: 0.6, backColor: '#FFFFFF', altBackColor: '#FFFFFF' }, pageFooter: { height: 0.8, backColor: '#FFFFFF', altBackColor: '#FFFFFF' },
        reportFooter: { height: 0.8, backColor: '#FFFFFF', altBackColor: '#FFFFFF' } }, controls: [] };
  }
  function setGroupSections(r) {
    const groups = r.grouping.filter(x => x.kind === 'group');
    for (const id of Object.keys(r.sections)) if (/^group(Header|Footer)\d$/.test(id)) {
      const k = +id.slice(-1);
      const g = groups[k];
      if (!g || (id.startsWith('groupHeader') && !g.header) || (id.startsWith('groupFooter') && !g.footer)) { delete r.sections[id]; r.controls = r.controls.filter(c => c.section !== id); }
    }
    groups.forEach((g, k) => {
      if (g.header && !r.sections['groupHeader' + k]) r.sections['groupHeader' + k] = { height: 0.8, backColor: '#FFFFFF', altBackColor: '#FFFFFF', repeat: false, pageBreak: '없음' };
      if (g.footer && !r.sections['groupFooter' + k]) r.sections['groupFooter' + k] = { height: 0.8, backColor: '#FFFFFF', altBackColor: '#FFFFFF', pageBreak: '없음' };
    });
  }
  function reportWizard(d, o) {
    const r = newReport(d, o.source);
    const W = o.orient === '가로' ? 27.4 : 19.0;
    r.props.width = W;
    r.page.orient = o.orient === '가로' ? 'landscape' : 'portrait';
    r.grouping = o.groups.map(f => ({ kind: 'group', field: f, order: 'asc', header: true, footer: !!(o.sums && o.sums.length), keep: '표시 안 함', by: '각 값' }))
      .concat((o.sorts || []).map(([f, ord]) => ({ kind: 'sort', field: f, order: ord || 'asc' })));
    setGroupSections(r);
    Object.assign(r.sections.reportHeader, { height: 1.6, backColor: '#DDE7F3', altBackColor: '#DDE7F3' });
    Object.assign(r.sections.detail, { altBackColor: '#F2F2F2' });
    const t = addControl(r, 'label', { section: 'reportHeader', left: 0, top: 0, width: 9, height: 1.1, caption: o.title || o.source, fontSize: 20, fontWeight: '보통', withLabel: false });
    const cols = o.fields;
    const nonGroup = cols.filter(f => !o.groups.includes(f));
    const step = W / cols.length;
    cols.forEach((f, i) => {
      addControl(r, 'label', { section: 'pageHeader', left: i * step, top: 0.1, width: step - 0.1, height: 0.6, caption: f, fontWeight: '굵게', withLabel: false, underline: false });
    });
    addControl(r, 'line', { section: 'pageHeader', left: 0, top: 0.75, width: W, height: 0, withLabel: false });
    o.groups.forEach((f, k) => {
      const i = cols.indexOf(f);
      addControl(r, 'textbox', { section: 'groupHeader' + k, left: i * step, top: 0.1, width: step - 0.1, height: 0.6, controlSource: f, withLabel: false, borderStyle: '투명' });
    });
    nonGroup.forEach(f => {
      const i = cols.indexOf(f);
      addControl(r, 'textbox', { section: 'detail', left: i * step, top: 0, width: step - 0.1, height: 0.6, controlSource: f, withLabel: false, borderStyle: '투명' });
    });
    if (o.sums && o.sums.length) {
      o.groups.forEach((g, k) => {
        const sec = 'groupFooter' + k;
        r.sections[sec].height = 1.4;
        addControl(r, 'textbox', { section: sec, left: 0, top: 0.1, width: W * 0.6, height: 0.6, borderStyle: '투명', withLabel: false, name: 'AccessTotals' + g,
          controlSource: `="요약 '${g}' = " & " " & [${g}] & " (" & Count(*) & " " & IIf(Count(*)=1,"세부 레코드","세부 레코드") & ")"` });
        addControl(r, 'label', { section: sec, left: 0, top: 0.75, width: 2.5, height: 0.6, caption: '합계', withLabel: false });
        o.sums.forEach(f => { const i = cols.indexOf(f); addControl(r, 'textbox', { section: sec, left: i * step, top: 0.75, width: step - 0.1, height: 0.6, controlSource: `=Sum([${f}])`, borderStyle: '투명', withLabel: false }); });
      });
      addControl(r, 'label', { section: 'reportFooter', left: 0, top: 0.1, width: 2.5, height: 0.6, caption: '총합계', withLabel: false });
      o.sums.forEach(f => { const i = cols.indexOf(f); addControl(r, 'textbox', { section: 'reportFooter', left: i * step, top: 0.1, width: step - 0.1, height: 0.6, controlSource: `=Sum([${f}])`, borderStyle: '투명', withLabel: false }); });
    }
    addControl(r, 'textbox', { section: 'pageFooter', left: 0, top: 0.1, width: 6, height: 0.6, controlSource: '=Now()', borderStyle: '투명', withLabel: false });
    addControl(r, 'textbox', { section: 'pageFooter', left: W - 6, top: 0.1, width: 6, height: 0.6, controlSource: '="페이지 " & [Page] & "/" & [Pages]', textAlign: '오른쪽', borderStyle: '투명', withLabel: false });
    // 바인딩된 텍스트 상자의 형식은 필드 형식을 따른다
    return r;
  }
  /* 13.2~13.3 따라 하기의 디자인 정리(모범 풀이용) */
  function reportTidy(d, r, o) {
    const X0 = 0.3;
    const pos = { [o.cols[0]]: [0.3, 1.6, '가운데'], [o.cols[1]]: [1.9, 1.8, '가운데'], [o.cols[2]]: [3.7, 1.8, '가운데'], [o.cols[3]]: [5.5, 2.6, '오른쪽'],
      [o.cols[4]]: [8.1, 2.6, '오른쪽'], [o.cols[5]]: [10.7, 2.4, '오른쪽'], [o.cols[6]]: [13.1, 2.7, '오른쪽'] };
    r.controls = r.controls.filter(c => c.section !== 'pageFooter');
    r.sections.pageFooter.height = 0;
    const gtb = r.controls.find(c => c.section === 'groupHeader0' && c.type === 'textbox');
    gtb.section = 'detail';
    r.sections.groupHeader0.height = 0;
    for (const s of Object.values(r.sections)) { s.backColor = '#FFFFFF'; s.altBackColor = '#FFFFFF'; }
    for (const c of r.controls.filter(x => x.section === 'detail' && x.type === 'textbox')) {
      const p = pos[c.controlSource];
      Object.assign(c, { left: p[0], width: p[1], textAlign: p[2], top: 0.04, height: 0.48, fontSize: 10, borderStyle: '투명' });
      if (c.controlSource === o.cols[0]) c.hideDuplicates = true;
      if (o.money.includes(c.controlSource)) { c.format = '통화'; c.decimals = 0; }
      if (c.controlSource === o.days) c.format = '0';
    }
    for (const c of r.controls.filter(x => x.section === 'pageHeader' && x.type === 'label')) {
      const p = pos[c.caption];
      Object.assign(c, { left: p[0], width: p[1], textAlign: '가운데', top: 0.1, height: 0.5, fontSize: 10 });
    }
    r.controls = r.controls.filter(c => !(c.section === 'pageHeader' && c.type === 'line'));
    addControl(r, 'line', { section: 'pageHeader', left: X0, top: 0.02, width: 15.5, height: 0, withLabel: false });
    addControl(r, 'line', { section: 'pageHeader', left: X0, top: 0.68, width: 15.5, height: 0, withLabel: false });
    r.sections.pageHeader.height = 0.72;
    r.sections.detail.height = 0.56;
    // 그룹 바닥글
    r.controls = r.controls.filter(c => !(c.section === 'groupFooter0' && (c.name.startsWith('AccessTotals') || (c.type === 'label' && c.caption === '합계'))));
    addControl(r, 'textbox', { section: 'groupFooter0', left: 1.9, top: 0.08, width: 5.6, height: 0.5, controlSource: `=[${o.gname}] & " 합계"`, textAlign: '가운데', fontSize: 10, borderStyle: '투명', withLabel: false });
    for (const c of r.controls.filter(x => x.section === 'groupFooter0' && /^=Sum\(/.test(x.controlSource || ''))) {
      const f = /\[(.+)\]/.exec(c.controlSource)[1], p = pos[f];
      Object.assign(c, { left: p[0], width: p[1], top: 0.08, height: 0.5, format: '통화', decimals: 0, textAlign: '오른쪽', fontSize: 10 });
    }
    addControl(r, 'line', { section: 'groupFooter0', left: X0, top: 0.02, width: 15.5, height: 0, withLabel: false });
    addControl(r, 'line', { section: 'groupFooter0', left: X0, top: 0.64, width: 15.5, height: 0, withLabel: false });
    r.sections.groupFooter0.height = 0.68;
    // 보고서 바닥글: 총평균
    for (const c of r.controls.filter(x => x.section === 'reportFooter')) {
      if (c.type === 'label') Object.assign(c, { caption: '총평균', left: 1.9, width: 5.6, top: 0.08, height: 0.5, textAlign: '가운데', fontSize: 10 });
      if (c.type === 'textbox') { const f = /\[(.+)\]/.exec(c.controlSource)[1], p = pos[f]; Object.assign(c, { controlSource: `=Avg([${f}])`, left: p[0], width: p[1], top: 0.08, height: 0.5, format: '통화', decimals: 0, textAlign: '오른쪽', fontSize: 10 }); }
    }
    addControl(r, 'line', { section: 'reportFooter', left: X0, top: 0.64, width: 15.5, height: 0, withLabel: false });
    r.sections.reportFooter.height = 0.7;
    // 보고서 머리글
    r.controls = r.controls.filter(c => c.section !== 'reportHeader');
    addControl(r, 'label', { section: 'reportHeader', left: X0, top: 0.15, width: 15.5, height: 0.9, caption: o.title, fontSize: 16, fontWeight: '굵게', textAlign: '가운데', withLabel: false });
    addControl(r, 'textbox', { section: 'reportHeader', left: 11.3, top: 1.25, width: 4.5, height: 0.5, controlSource: '="작성일자 : " & Format(Date(),"yyyy-mm-dd")', fontSize: 9, textAlign: '오른쪽', borderStyle: '투명', withLabel: false });
    r.sections.reportHeader.height = 1.9;
    r.props.width = 16.5;
    return r;
  }

  OA.acc = { TYPES, NUM_SIZES, FORMATS, newDatabase, defaultField, applyTypeDefaults, createTable, setPrimaryKey, coerce, checkRecord, insertRow, updateRow, deleteRow,
    isUnique, addRelationship, removeRelationship, newQueryDesign, addQueryTable, parseCalc, critToSql, designToSQL, fromClause, sqlToDesign, saveQuery, querySQL,
    runSQL, runSaved, objFields, formatValue, newForm, defaultPage, controlDefaults, addControl, growSection, saveForm, saveReport, newReport, setGroupSections,
    reportWizard, reportTidy, SEC_ORDER, sectionTitle, NAME_PREFIX, fieldTable };
})(typeof window !== 'undefined' ? window : globalThis);

/* 공통 유틸(DOM 없음). 브라우저와 Node 테스트가 같이 읽는다. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};

  function colName(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function colNum(s) {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function addr(r, c, absR, absC) { return (absC ? '$' : '') + colName(c) + (absR ? '$' : '') + r; }
  const ADDR_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
  function parseAddr(s) {
    const m = ADDR_RE.exec(String(s).trim());
    if (!m) return null;
    return { r: +m[4], c: colNum(m[2]), absC: !!m[1], absR: !!m[3] };
  }
  /* 'B4:K16', '작업표!$B$4:$K$16', "'작업 표'!A1" → {sheet, r1,c1,r2,c2} */
  function parseRange(s) {
    s = String(s).trim();
    let sheet = null;
    const bang = s.lastIndexOf('!');
    if (bang >= 0) {
      sheet = s.slice(0, bang);
      if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
      s = s.slice(bang + 1);
    }
    const parts = s.split(':');
    const a = parseAddr(parts[0]);
    const b = parts.length > 1 ? parseAddr(parts[1]) : a;
    if (!a || !b) return null;
    return { sheet, r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
  }
  function rangeStr(g, abs, sheet) {
    const a = addr(g.r1, g.c1, abs, abs), b = addr(g.r2, g.c2, abs, abs);
    const core = (g.r1 === g.r2 && g.c1 === g.c2) ? a : a + ':' + b;
    if (!sheet) return core;
    return (/[^\w가-힣]/.test(sheet) ? "'" + sheet.replace(/'/g, "''") + "'" : sheet) + '!' + core;
  }
  function inRange(g, r, c) { return r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2; }
  function sameRange(a, b) { return a && b && a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2; }

  /* Excel 표시용 반올림: 0.5 는 0 에서 먼 쪽 */
  function roundHalfAway(x, d) {
    if (!isFinite(x)) return x;
    const p = Math.pow(10, d || 0);
    const v = Math.abs(x) * p;
    const r = Math.floor(v * (1 + 4e-15) + 0.5);         // 부동소수 표현 오차(267.49999…) 보정, 몇 ulp 만큼만
    return Math.sign(x) * r / p;
  }
  function thousands(intStr) { return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
  function approx(a, b, eps) { return Math.abs(a - b) <= (eps == null ? 1e-6 : eps); }
  function koCompare(a, b) { return String(a).localeCompare(String(b), 'ko'); }
  function clone(o) { return o == null ? o : JSON.parse(JSON.stringify(o)); }
  let seq = 0;
  function uid(p) { seq += 1; return (p || 'id') + seq + '_' + Math.random().toString(36).slice(2, 6); }
  function normText(s) { return String(s == null ? '' : s).replace(/[\s ]+/g, '').trim(); }

  OA.util = { colName, colNum, addr, parseAddr, parseRange, rangeStr, inRange, sameRange, roundHalfAway, thousands,
    approx, koCompare, clone, uid, normText };
})(typeof window !== 'undefined' ? window : globalThis);

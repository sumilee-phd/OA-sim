/* Excel 표시 형식 엔진(DOM 없음). 형식 코드: 일반(G/표준), 0·#·,·.·%, 구역 ;, "문자", \문자, [h], yyyy-mm-dd, h:mm 등. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util;

  const GENERAL = 'G/표준';
  function isGeneral(f) { return !f || f === 'General' || f === GENERAL; }

  function splitSections(fmt) {
    const out = [];
    let cur = '', q = false;
    for (let i = 0; i < fmt.length; i++) {
      const ch = fmt[i];
      if (ch === '"') q = !q;
      if (ch === '\\' && !q) { cur += ch + (fmt[i + 1] || ''); i++; continue; }
      if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }

  /* 형식 구역을 토큰으로: {k:'lit',v} {k:'num',v} {k:'date',v} {k:'text'} {k:'elapsed',v} */
  function tokenize(sec) {
    const t = [];
    let i = 0;
    while (i < sec.length) {
      const ch = sec[i];
      if (ch === '"') { const j = sec.indexOf('"', i + 1); t.push({ k: 'lit', v: sec.slice(i + 1, j < 0 ? sec.length : j) }); i = j < 0 ? sec.length : j + 1; continue; }
      if (ch === '\\') { t.push({ k: 'lit', v: sec[i + 1] || '' }); i += 2; continue; }
      if (ch === '_') { t.push({ k: 'lit', v: ' ' }); i += 2; continue; }
      if (ch === '*') { i += 2; continue; }
      if (ch === '[') {
        const j = sec.indexOf(']', i);
        const inner = sec.slice(i + 1, j);
        if (/^h+$/i.test(inner) || /^m+$/i.test(inner) || /^s+$/i.test(inner)) t.push({ k: 'elapsed', v: inner.toLowerCase() });
        else if (inner.startsWith('$')) { const sym = inner.slice(1).split('-')[0]; if (sym) t.push({ k: 'lit', v: sym }); }
        i = j + 1; continue;
      }
      if (/[0#?.,%]/.test(ch)) { t.push({ k: 'num', v: ch }); i++; continue; }
      if (ch === '@') { t.push({ k: 'text' }); i++; continue; }
      if (/[yYmMdDhHsS]/.test(ch)) {
        let j = i; while (j < sec.length && sec[j].toLowerCase() === ch.toLowerCase()) j++;
        t.push({ k: 'date', v: sec.slice(i, j).toLowerCase() }); i = j; continue;
      }
      if (/^(AM\/PM|오전\/오후)/i.test(sec.slice(i))) { const m = /^(AM\/PM|오전\/오후)/i.exec(sec.slice(i))[0]; t.push({ k: 'ampm', v: m }); i += m.length; continue; }
      t.push({ k: 'lit', v: ch }); i++;
    }
    return t;
  }

  /* 일련번호 → {y,m,d,hh,mi,ss}. 1900 체계(60 = 1900-02-29 가상) */
  function serialParts(v) {
    let days = Math.floor(v);
    let frac = v - days;
    let secs = Math.round(frac * 86400);
    if (secs >= 86400) { secs -= 86400; days += 1; }
    let y, m, d;
    if (days === 60) { y = 1900; m = 2; d = 29; } else {
      const base = Date.UTC(1899, 11, 31) + (days > 60 ? days - 1 : days) * 86400000;
      const dt = new Date(base);
      y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate();
    }
    return { y, m, d, hh: Math.floor(secs / 3600), mi: Math.floor((secs % 3600) / 60), ss: secs % 60, total: v };
  }
  function dateSerial(y, m, d) {
    const ms = Date.UTC(y, m - 1, d);
    let days = Math.round((ms - Date.UTC(1899, 11, 31)) / 86400000);
    if (days >= 60) days += 1;
    return days;
  }
  const pad = (n, w) => String(n).padStart(w, '0');

  function fmtDate(v, toks) {
    const p = serialParts(v);
    const hasH = toks.some(t => t.k === 'date' && t.v[0] === 'h') || toks.some(t => t.k === 'elapsed');
    let out = '';
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.k === 'lit') { out += t.v; continue; }
      if (t.k === 'elapsed') {
        const totalSec = Math.round(v * 86400);
        if (t.v[0] === 'h') out += Math.floor(totalSec / 3600);
        else if (t.v[0] === 'm') out += Math.floor(totalSec / 60);
        else out += totalSec;
        continue;
      }
      if (t.k === 'ampm') { out += p.hh < 12 ? (t.v.startsWith('오') ? '오전' : 'AM') : (t.v.startsWith('오') ? '오후' : 'PM'); continue; }
      if (t.k !== 'date') { if (t.k === 'num') out += t.v; continue; }
      const c = t.v;
      if (c[0] === 'y') out += c.length <= 2 ? pad(p.y % 100, 2) : String(p.y);
      else if (c[0] === 'd') out += c.length === 1 ? p.d : c.length === 2 ? pad(p.d, 2) : c.length === 3 ? '일월화수목금토'[new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay()] : '일월화수목금토'[new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay()] + '요일';
      else if (c[0] === 'h') {
        const ampm = toks.some(x => x.k === 'ampm');
        const h = ampm ? ((p.hh % 12) || 12) : p.hh;
        out += c.length === 1 ? h : pad(h, 2);
      } else if (c[0] === 's') out += c.length === 1 ? p.ss : pad(p.ss, 2);
      else if (c[0] === 'm') {
        // h 뒤 또는 s 앞의 m 은 분
        let prev = null, next = null;
        for (let j = i - 1; j >= 0; j--) if (toks[j].k === 'date' || toks[j].k === 'elapsed') { prev = toks[j].v; break; }
        for (let j = i + 1; j < toks.length; j++) if (toks[j].k === 'date') { next = toks[j].v; break; }
        const minute = (prev && prev[0] === 'h') || (next && next[0] === 's') || (!toks.some(x => x.k === 'date' && (x.v[0] === 'y' || x.v[0] === 'd')) && hasH);
        if (minute) out += c.length === 1 ? p.mi : pad(p.mi, 2);
        else out += c.length === 1 ? p.m : pad(p.m, 2);
      }
    }
    return out;
  }

  function fmtNumber(v, toks, neg) {
    const numToks = toks.filter(t => t.k === 'num');
    let pct = numToks.filter(t => t.v === '%').length;
    let x = Math.abs(v) * Math.pow(100, pct);
    const ph = toks.map(t => t.k === 'num' && /[0#?]/.test(t.v));
    const firstPh = ph.indexOf(true), lastPh = ph.lastIndexOf(true);
    if (firstPh < 0) { // 자리 표시 없음: 리터럴만
      return (neg ? '-' : '') + toks.map(t => t.k === 'lit' ? t.v : (t.k === 'num' && t.v === '%' ? '%' : '')).join('');
    }
    let dotIdx = -1;
    for (let i = firstPh; i <= lastPh + 1 && i < toks.length; i++) if (toks[i].k === 'num' && toks[i].v === '.') { dotIdx = i; break; }
    let intZeros = 0, decZeros = 0, decHash = 0, comma = false;
    for (let i = firstPh; i <= lastPh; i++) {
      const t = toks[i];
      if (t.k !== 'num') continue;
      if (dotIdx >= 0 && i > dotIdx) { if (t.v === '0') decZeros++; else if (t.v === '#' || t.v === '?') decHash++; }
      else { if (t.v === '0') intZeros++; if (t.v === ',') comma = true; }
    }
    // 끝 쉼표(배율) 처리: 마지막 자리 표시 뒤의 , 는 1000 나누기
    for (let i = lastPh + 1; i < toks.length && toks[i].k === 'num' && toks[i].v === ','; i++) x /= 1000;
    const decs = decZeros + decHash;
    const r = U.roundHalfAway(x, decs);
    let [ip, dp] = r.toFixed(decs).split('.');
    if (ip === '0' && intZeros === 0) ip = '';
    ip = ip.padStart(intZeros, '0');
    if (comma) ip = U.thousands(ip);
    if (dp) { // 뒤쪽 # 자리의 0 제거
      let keep = dp.length;
      while (keep > decZeros && dp[keep - 1] === '0') keep--;
      dp = dp.slice(0, keep);
    }
    const numStr = ip + (dotIdx >= 0 ? '.' + (dp || '') : '');
    let pre = '', post = '';
    for (let i = 0; i < firstPh; i++) { const t = toks[i]; pre += t.k === 'lit' ? t.v : (t.k === 'num' && t.v === '%' ? '%' : ''); }
    for (let i = lastPh + 1; i < toks.length; i++) { const t = toks[i]; if (t.k === 'num' && (t.v === '.' || t.v === ',')) continue; post += t.k === 'lit' ? t.v : (t.k === 'num' && t.v === '%' ? '%' : ''); }
    return (neg ? '-' : '') + pre + numStr + post;
  }

  function generalNum(v) {
    if (!isFinite(v)) return '#NUM!';
    if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
    const a = Math.abs(v);
    if (a !== 0 && (a >= 1e11 || a < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E').replace(/E\+?/, 'E+').replace('E+-', 'E-');
    const intDigits = a >= 1 ? Math.floor(Math.log10(a)) + 1 : 1;
    const decs = Math.max(0, 10 - intDigits);
    let s = U.roundHalfAway(v, decs).toFixed(decs);
    if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
    return s;
  }

  /* G/표준 숫자가 열 너비보다 넓을 때: Excel 처럼 소수 자릿수를 줄이고, 정수부도 넘치면 지수 표기로 줄인다.
     fits(text) 가 참인 가장 긴 표기를 돌려준다. 맞는 표기가 없으면 null(호출 쪽에서 ### 표시) */
  function fitGeneral(v, fits) {
    const full = generalNum(v);
    if (fits(full)) return full;
    const a = Math.abs(v);
    if (a >= 1 || a === 0) {
      const m = /\.(\d+)$/.exec(full);
      for (let d = m ? m[1].length - 1 : -1; d >= 0; d--) {
        let t = U.roundHalfAway(v, d).toFixed(d);
        if (t.includes('.')) t = t.replace(/0+$/, '').replace(/\.$/, '');
        if (fits(t)) return t;
      }
    } else {
      for (let d = 9; d >= 1; d--) { let t = U.roundHalfAway(v, d).toFixed(d).replace(/0+$/, '').replace(/\.$/, ''); if (t !== '0' && t !== '-0' && fits(t)) return t; }
    }
    for (let p = 5; p >= 0; p--) {
      const t = v.toExponential(p).replace(/\.?0+e/, 'e').replace(/e\+?/, 'E+').replace('E+-', 'E-').replace(/E([+-])(\d)$/, 'E$10$2');
      if (fits(t)) return t;
    }
    return null;
  }

  function isError(v) { return v && typeof v === 'object' && v.e; }

  /* value 를 형식 fmt 로 표시. 결과 {text, align:'left'|'right'|'center'} */
  function format(v, fmt) {
    if (v === null || v === undefined || v === '') return { text: '', align: 'left' };
    if (isError(v)) return { text: v.e, align: 'center' };
    if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', align: 'center' };
    if (typeof v === 'string') {
      if (!isGeneral(fmt)) {
        const secs = splitSections(fmt);
        if (secs.length >= 4) return { text: tokenize(secs[3]).map(t => t.k === 'text' ? v : t.k === 'lit' ? t.v : '').join(''), align: 'left' };
        if (secs[0].includes('@')) return { text: tokenize(secs[0]).map(t => t.k === 'text' ? v : t.k === 'lit' ? t.v : '').join(''), align: 'left' };
      }
      return { text: v, align: 'left' };
    }
    if (isGeneral(fmt)) return { text: generalNum(v), align: 'right' };
    if (fmt === '@') return { text: generalNum(v), align: 'left' };
    const secs = splitSections(fmt);
    let sec = secs[0], neg = v < 0;
    if (secs.length >= 2 && v < 0) { sec = secs[1]; neg = false; }
    if (secs.length >= 3 && v === 0) sec = secs[2];
    if (secs.length === 1 || (secs.length >= 2 && v >= 0)) neg = v < 0;
    sec = sec.replace(/\[(?:빨강|파랑|검정|녹색|Red|Blue|Black|Green|Color\d+)\]/gi, '');
    if (/\[(?:[<>=])/.test(sec)) sec = sec.replace(/\[[<>=][^\]]*\]/g, '');
    const toks = tokenize(sec);
    if (toks.some(t => t.k === 'date' || t.k === 'elapsed')) {
      if (v < 0) return { text: '#'.repeat(8), align: 'right' };
      return { text: fmtDate(v, toks), align: 'right' };
    }
    return { text: fmtNumber(v, toks, neg), align: 'right' };
  }

  function isDateFmt(fmt) {
    if (isGeneral(fmt)) return false;
    const toks = tokenize(splitSections(fmt)[0]);
    return toks.some(t => t.k === 'date' || t.k === 'elapsed');
  }

  /* 셀 입력 문자열 해석 → {v, nf?, q?}. Excel 의 자동 형식 지정 흉내 */
  function parseInput(s) {
    if (s === null || s === undefined) return { v: null };
    const raw = String(s);
    if (raw === '') return { v: null };
    if (raw[0] === "'") return { v: raw.slice(1), q: true };
    const t = raw.trim();
    if (/^(TRUE|FALSE)$/i.test(t)) return { v: /^TRUE$/i.test(t) };
    let m;
    if ((m = /^([+-]?)(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.exec(t))) {
      const v = parseFloat(t.replace(/,/g, ''));
      return { v, nf: m[3] ? (m[4] ? '#,##0.00' : '#,##0') : null };
    }
    if ((m = /^([+-]?)(\d*\.?\d+)%$/.exec(t))) return { v: parseFloat(t) / 100, nf: t.includes('.') ? '0.00%' : '0%' };
    if ((m = /^([+-]?)₩\s?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?$/.exec(t))) return { v: parseFloat(t.replace(/[₩,\s]/g, '')), nf: '₩#,##0;-₩#,##0' };
    if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t))) {
      const v = (+m[1] * 3600 + +m[2] * 60 + (+m[3] || 0)) / 86400;
      return { v, nf: m[3] ? 'h:mm:ss' : 'h:mm' };
    }
    if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(t))) {
      const y = +m[1], mo = +m[2], d = +m[3];
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return { v: dateSerial(y, mo, d), nf: 'yyyy-mm-dd' };
    }
    if (/^[+-]?\d+(\.\d+)?[eE][+-]?\d+$/.test(t)) return { v: parseFloat(t) };
    return { v: raw };
  }

  /* 셀 서식 대화상자 범주 목록 */
  const CATEGORIES = [
    { id: '일반', sample: GENERAL },
    { id: '숫자', sample: '0' },
    { id: '통화', sample: '₩#,##0;-₩#,##0' },
    { id: '회계', sample: '_-₩* #,##0_-;-₩* #,##0_-;_-₩* "-"_-;_-@_-' },
    { id: '날짜', sample: 'yyyy-mm-dd' },
    { id: '시간', sample: 'h:mm' },
    { id: '백분율', sample: '0%' },
    { id: '분수', sample: '# ?/?' },
    { id: '지수', sample: '0.00E+00' },
    { id: '텍스트', sample: '@' },
    { id: '기타', sample: GENERAL },
    { id: '사용자 지정', sample: null },
  ];
  const CUSTOM_LIST = [GENERAL, '0', '0.00', '#,##0', '#,##0.00', '#,##0;-#,##0', '#,##0;[빨강]-#,##0', '₩#,##0;-₩#,##0',
    '₩#,##0;[빨강]-₩#,##0', '0%', '0.00%', '0.00E+00', '# ?/?', 'yyyy-mm-dd', 'yyyy"년" m"월" d"일"', 'm/d', 'h:mm', 'h:mm:ss',
    '[h]:mm', 'h:mm AM/PM', '@'];

  OA.numfmt = { GENERAL, isGeneral, format, parseInput, dateSerial, serialParts, isDateFmt, CATEGORIES, CUSTOM_LIST, generalNum, fitGeneral };
})(typeof window !== 'undefined' ? window : globalThis);

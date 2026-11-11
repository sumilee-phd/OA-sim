/* 차트 그리기(SVG 문자열, DOM 없음): 묶은 세로 막대형·꺾은선형·표식이 있는 꺾은선형·혼합, 기본/보조 축, 축 제목, 범례, 데이터 레이블. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const COLORS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
  const TYPE_NAMES = { col: '묶은 세로 막대형', stacked: '누적 세로 막대형', line: '꺾은선형', lineMarkers: '표식이 있는 꺾은선형', area: '영역형' };
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

  function nice(max, min) {
    const span = Math.max(1e-9, max - min);
    const raw = span / 6;
    const p = Math.pow(10, Math.floor(Math.log10(raw)));
    const m = raw / p;
    const step = (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p;
    const lo = Math.floor(min / step) * step;
    let hi = Math.ceil(max / step) * step;
    if (hi === max && max !== 0) hi += 0;
    if (hi <= lo) hi = lo + step;
    return { lo, hi, step };
  }
  const tw = (s, px) => String(s).length * px * 0.62 + [...String(s)].filter(ch => ch.charCodeAt(0) > 0x2fff).length * px * 0.38;

  function data(wb, ch) {
    const XL = OA.xl;
    const series = ch.series.map((s, i) => ({
      s, i, name: XL.seriesName(wb, s, i), vals: XL.refValues(wb, s.val).map(v => typeof v === 'number' ? v : null), nf: XL.refFormat(wb, s.val),
      color: s.color || COLORS[i % COLORS.length],
    }));
    let cats = [];
    const first = ch.series.find(s => s.cat);
    if (first) cats = XL.refValues(wb, first.cat).map(v => v === null ? '' : OA.formula.toStr(v));
    const n = Math.max(0, ...series.map(s => s.vals.length));
    while (cats.length < n) cats.push(String(cats.length + 1));
    return { series, cats: cats.slice(0, n), n };
  }

  /* opts: {w, h, selSeries} */
  function svg(wb, ch, opts) {
    const W = opts.w, H = opts.h;
    const d = data(wb, ch);
    const NF = OA.numfmt;
    const out = [];
    out.push(`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="Malgun Gothic, sans-serif">`);
    out.push(`<rect x="0" y="0" width="${W}" height="${H}" fill="#fff" data-part="area"/>`);
    let top = 10, bottom = H - 8, left = 10, right = W - 12;
    const title = ch.title.show ? OA.xl.chartTitleText(wb, ch) : null;
    if (title !== null) { out.push(`<text x="${W / 2}" y="${top + 16}" font-size="18" fill="#595959" text-anchor="middle" font-weight="700" data-part="title" style="cursor:text">${esc(title)}</text>`); top += 30; }
    if (ch.legend.show && ch.legend.pos === 'bottom') bottom -= 22;
    if (ch.legend.show && ch.legend.pos === 'top') top += 20;
    if (ch.legend.show && ch.legend.pos === 'right') right -= 90;
    if (ch.legend.show && ch.legend.pos === 'left') left += 90;
    const hasX = ch.axisTitle.x.show, hasY = ch.axisTitle.y.show;
    if (hasX) bottom -= 22;
    if (hasY) left += 20;
    const prim = d.series.filter(s => !s.s.secondary), sec = d.series.filter(s => s.s.secondary);
    const scaleOf = list => {
      const vals = list.flatMap(s => s.vals).filter(v => v !== null);
      const mx = vals.length ? Math.max(0, ...vals) : 1, mn = vals.length ? Math.min(0, ...vals) : 0;
      return nice(mx, mn);
    };
    const sp = scaleOf(prim.length ? prim : d.series), ss = sec.length ? scaleOf(sec) : null;
    const nf = (prim[0] || d.series[0] || {}).nf || NF.GENERAL;
    const nfs = ss ? sec[0].nf : null;
    const lab = (v, f) => NF.format(v, f).text;
    const ticks = sc => { const t = []; for (let v = sc.lo; v <= sc.hi + sc.step * 1e-6; v += sc.step) t.push(+v.toPrecision(12)); return t; };
    const tp = ticks(sp), ts = ss ? ticks(ss) : [];
    const yLabW = Math.max(...tp.map(v => tw(lab(v, nf), 11))) + 8;
    left += ch.axes ? yLabW : 0;
    if (ss) right -= Math.max(...ts.map(v => tw(lab(v, nfs), 11))) + 8;
    const catH = 18;
    const plotB = bottom - (ch.axes ? catH : 0);
    const plotT = top + 6;
    const y = (v, sc) => plotB - (v - sc.lo) / (sc.hi - sc.lo) * (plotB - plotT);
    // 눈금선·축
    if (ch.gridlines) for (const v of tp) out.push(`<line x1="${left}" x2="${right}" y1="${y(v, sp)}" y2="${y(v, sp)}" stroke="#d9d9d9"/>`);
    if (ch.axes) {
      for (const v of tp) out.push(`<text x="${left - 5}" y="${y(v, sp) + 4}" font-size="11" fill="#595959" text-anchor="end">${esc(lab(v, nf))}</text>`);
      if (ss) for (const v of ts) out.push(`<text x="${right + 5}" y="${y(v, ss) + 4}" font-size="11" fill="#595959">${esc(lab(v, nfs))}</text>`);
    }
    out.push(`<line x1="${left}" x2="${right}" y1="${y(Math.max(sp.lo, 0), sp)}" y2="${y(Math.max(sp.lo, 0), sp)}" stroke="#bfbfbf"/>`);
    const n = Math.max(1, d.n);
    const band = (right - left) / n;
    if (ch.axes) d.cats.forEach((c, i) => out.push(`<text x="${left + band * (i + 0.5)}" y="${plotB + 14}" font-size="11" fill="#595959" text-anchor="middle">${esc(c)}</text>`));
    for (let i = 0; i <= n; i++) out.push(`<line x1="${left + band * i}" x2="${left + band * i}" y1="${plotB}" y2="${plotB + 4}" stroke="#bfbfbf"/>`);
    // 계열
    const cols = d.series.filter(s => s.s.type === 'col' || s.s.type === 'stacked');
    const bw = band * 0.62 / Math.max(1, cols.length);
    const labels = [];
    cols.forEach((s, k) => {
      const sc = s.s.secondary && ss ? ss : sp;
      s.vals.forEach((v, i) => {
        if (v === null) return;
        const x = left + band * i + band * 0.19 + bw * k;
        const y0 = y(Math.max(sc.lo, 0), sc), y1 = y(v, sc);
        out.push(`<rect x="${x}" y="${Math.min(y0, y1)}" width="${bw}" height="${Math.abs(y0 - y1)}" fill="${s.color}" data-part="series-${s.i}" style="cursor:pointer"/>`);
        if (s.s.labels) { const pos = s.s.labels; const ly = pos === 'inEnd' ? y1 + 14 : pos === 'center' ? (y0 + y1) / 2 + 4 : pos === 'inBase' ? y0 - 4 : y1 - 4; labels.push(`<text x="${x + bw / 2}" y="${ly}" font-size="11" fill="#404040" text-anchor="middle">${esc(lab(v, s.nf))}</text>`); }
      });
      if (opts.selSeries === s.i) s.vals.forEach((v, i) => { if (v !== null) { const x = left + band * i + band * 0.19 + bw * k + bw / 2, yy = y(v, sc); out.push(`<circle cx="${x}" cy="${yy}" r="3.5" fill="#fff" stroke="#555"/>`); } });
    });
    d.series.filter(s => s.s.type === 'line' || s.s.type === 'lineMarkers').forEach(s => {
      const sc = s.s.secondary && ss ? ss : sp;
      const pts = s.vals.map((v, i) => v === null ? null : [left + band * (i + 0.5), y(v, sc)]);
      const path = pts.filter(Boolean).map((p, i) => (i ? 'L' : 'M') + p[0].toFixed(1) + ' ' + p[1].toFixed(1)).join(' ');
      out.push(`<path d="${path}" stroke="${s.color}" stroke-width="2.6" fill="none" data-part="series-${s.i}" style="cursor:pointer"/>`);
      out.push(`<path d="${path}" stroke="transparent" stroke-width="10" fill="none" data-part="series-${s.i}" style="cursor:pointer"/>`);
      if (s.s.type === 'lineMarkers') pts.forEach(p => { if (p) out.push(`<rect x="${p[0] - 4}" y="${p[1] - 4}" width="8" height="8" fill="${s.color}" data-part="series-${s.i}" style="cursor:pointer"/>`); });
      if (s.s.labels) pts.forEach((p, i) => {
        if (!p) return;
        const pos = s.s.labels;
        const lx = pos === 'left' ? p[0] - 8 : pos === 'right' ? p[0] + 8 : p[0], ly = pos === 'top' ? p[1] - 9 : pos === 'bottom' ? p[1] + 17 : p[1] + 4;
        labels.push(`<text x="${lx}" y="${ly}" font-size="11" fill="#404040" text-anchor="${pos === 'left' ? 'end' : pos === 'right' ? 'start' : 'middle'}">${esc(lab(s.vals[i], s.nf))}</text>`);
      });
      if (opts.selSeries === s.i) pts.forEach(p => { if (p) out.push(`<circle cx="${p[0]}" cy="${p[1]}" r="4" fill="#fff" stroke="#555"/>`); });
    });
    out.push(...labels);
    if (hasY) out.push(`<text transform="translate(${14} ${(plotT + plotB) / 2}) rotate(-90)" font-size="12" fill="#595959" text-anchor="middle" font-weight="700" data-part="axisY" style="cursor:text">${esc(ch.axisTitle.y.text)}</text>`);
    if (hasX) out.push(`<text x="${(left + right) / 2}" y="${bottom + 18}" font-size="12" fill="#595959" text-anchor="middle" font-weight="700" data-part="axisX" style="cursor:text">${esc(ch.axisTitle.x.text)}</text>`);
    if (ch.legend.show) {
      const items = d.series.map(s => ({ s, w: tw(s.name, 11) + 26 }));
      const total = items.reduce((a, b) => a + b.w, 0);
      let lx = ch.legend.pos === 'bottom' || ch.legend.pos === 'top' ? (W - total) / 2 : ch.legend.pos === 'right' ? W - 86 : 8;
      let ly = ch.legend.pos === 'bottom' ? H - 14 : ch.legend.pos === 'top' ? top - 8 : (plotT + plotB) / 2 - items.length * 9;
      const g = [`<g data-part="legend" style="cursor:move">`];
      for (const it of items) {
        const isLine = it.s.s.type === 'line' || it.s.s.type === 'lineMarkers';
        if (isLine) g.push(`<line x1="${lx}" x2="${lx + 18}" y1="${ly - 4}" y2="${ly - 4}" stroke="${it.s.color}" stroke-width="2.5"/>` + (it.s.s.type === 'lineMarkers' ? `<rect x="${lx + 5}" y="${ly - 8}" width="8" height="8" fill="${it.s.color}"/>` : ''));
        else g.push(`<rect x="${lx + 3}" y="${ly - 9}" width="11" height="11" fill="${it.s.color}"/>`);
        g.push(`<text x="${lx + 22}" y="${ly}" font-size="11" fill="#595959">${esc(it.s.name)}</text>`);
        if (ch.legend.pos === 'bottom' || ch.legend.pos === 'top') lx += it.w; else ly += 18;
      }
      g.push('</g>');
      out.push(g.join(''));
    }
    out.push('</svg>');
    return out.join('');
  }
  OA.chart = { svg, data, COLORS, TYPE_NAMES };
})(typeof window !== 'undefined' ? window : globalThis);

/* A4 쪽 그리기 도우미(미리 보기·5매 축소판 공용). 단위: mm, scale = 화면 px/mm */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const { h } = OA.ui;
  function page(scale, opts) {
    const o = opts || {};
    const W = o.landscape ? 297 : 210, H = o.landscape ? 210 : 297;
    const el = h('div', { class: 'a4' + (o.gray ? ' gray' : ''), style: { width: W * scale + 'px', height: H * scale + 'px' } });
    el.mm = { W, H, scale };
    return el;
  }
  function blank(scale, text) {
    const p = page(scale);
    p.appendChild(h('div', { style: { position: 'absolute', inset: 0, display: 'grid', placeItems: 'center', color: '#aaa', fontSize: Math.max(10, 6 * scale * 2) + 'px' } }, text || ''));
    return p;
  }
  function handMarks(p, scale, key) {
    if (!p || !p.mm) return p;
    const s = p.mm.scale;
    p.appendChild(h('div', { class: 'hand', style: { left: (p.mm.W / 2 - 45) * s + 'px', top: 14 * s + 'px', width: 90 * s + 'px', height: 16 * s + 'px', fontSize: Math.max(8, 3.4 * s) + 'px' } }, '비번호 · 수험번호 · 성명(자필)'));
    p.appendChild(h('div', { class: 'hand', style: { left: (p.mm.W / 2 - 14) * s + 'px', top: (p.mm.H - 14) * s + 'px', width: 28 * s + 'px', height: 9 * s + 'px', fontSize: Math.max(8, 3.4 * s) + 'px' } }, key || '쪽 번호'));
    return p;
  }
  OA.print = { page, blank, handMarks };
})(typeof window !== 'undefined' ? window : globalThis);

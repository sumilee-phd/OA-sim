/* 공통 화면 부품: DOM 도우미, 직접 그린 아이콘, 리본, 메뉴, 대화상자, 메시지 상자, 범위 선택 상자, 알림. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};

  function h(tag, props, ...kids) {
    const el = tag === 'svg' || props && props.svg ? document.createElementNS('http://www.w3.org/2000/svg', tag) : document.createElement(tag);
    if (props) for (const [k, v] of Object.entries(props)) {
      if (v === undefined || v === null || v === false || k === 'svg') continue;
      if (k === 'class') el.setAttribute('class', v);
      else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'html') el.innerHTML = v;
      else if (k === 'text') el.textContent = v;
      else if (k in el && typeof v !== 'string' && k !== 'list') el[k] = v;
      else el.setAttribute(k, v === true ? '' : v);
    }
    for (const kid of kids.flat(Infinity)) {
      if (kid === null || kid === undefined || kid === false) continue;
      el.appendChild(typeof kid === 'string' || typeof kid === 'number' ? document.createTextNode(String(kid)) : kid);
    }
    return el;
  }
  const svgNS = 'http://www.w3.org/2000/svg';

  // ------------------------------------------------------------------ 아이콘(32×32 좌표, 직접 그린 단순 도형)
  const A = '#217346', B = '#2b579a', R = '#c43e1c', K = '#444', Gy = '#9a9a9a';
  const I = {
    paste: `<rect x="7" y="6" width="18" height="22" rx="2" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="11" y="3" width="10" height="6" rx="1" fill="${Gy}"/><rect x="14" y="13" width="14" height="16" fill="#fff" stroke="${B}" stroke-width="2"/>`,
    cut: `<circle cx="10" cy="24" r="4" fill="none" stroke="${K}" stroke-width="2"/><circle cx="22" cy="24" r="4" fill="none" stroke="${K}" stroke-width="2"/><path d="M12 21 L23 5 M20 21 L9 5" stroke="${K}" stroke-width="2"/>`,
    copy: `<rect x="5" y="5" width="14" height="17" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="12" y="11" width="14" height="17" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    brush: `<rect x="6" y="4" width="18" height="9" rx="1" fill="#f3c86b" stroke="${K}" stroke-width="2"/><path d="M15 13 V18 M15 18 V28" stroke="${K}" stroke-width="3"/>`,
    bold: `<text x="16" y="24" font-size="20" font-weight="900" text-anchor="middle" fill="${K}">가</text>`,
    italic: `<text x="16" y="24" font-size="20" font-style="italic" text-anchor="middle" fill="${K}">가</text>`,
    underline: `<text x="16" y="21" font-size="18" text-anchor="middle" fill="${K}">가</text><path d="M7 27 H25" stroke="${K}" stroke-width="2"/>`,
    border: `<rect x="5" y="5" width="22" height="22" fill="none" stroke="${Gy}" stroke-dasharray="2 2"/><path d="M5 27 H27" stroke="${K}" stroke-width="3"/>`,
    borderAll: `<rect x="5" y="5" width="22" height="22" fill="none" stroke="${K}" stroke-width="2"/><path d="M16 5 V27 M5 16 H27" stroke="${K}" stroke-width="2"/>`,
    fill: `<path d="M8 14 L16 6 L25 15 L17 23 Z" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="25" width="24" height="5" fill="#ffd400"/>`,
    fontcolor: `<text x="16" y="21" font-size="19" text-anchor="middle" fill="${K}">가</text><rect x="4" y="25" width="24" height="5" fill="#c00000"/>`,
    alignL: `<path d="M5 8 H27 M5 14 H19 M5 20 H27 M5 26 H19" stroke="${K}" stroke-width="2"/>`,
    alignC: `<path d="M5 8 H27 M9 14 H23 M5 20 H27 M9 26 H23" stroke="${K}" stroke-width="2"/>`,
    alignR: `<path d="M5 8 H27 M13 14 H27 M5 20 H27 M13 26 H27" stroke="${K}" stroke-width="2"/>`,
    alignT: `<path d="M5 5 H27" stroke="${K}" stroke-width="2"/><path d="M10 8 H22 M10 13 H22" stroke="${B}" stroke-width="3"/>`,
    alignM: `<path d="M5 16 H27" stroke="${Gy}"/><path d="M10 12 H22 M10 20 H22" stroke="${B}" stroke-width="3"/>`,
    alignB: `<path d="M5 27 H27" stroke="${K}" stroke-width="2"/><path d="M10 19 H22 M10 24 H22" stroke="${B}" stroke-width="3"/>`,
    wrap: `<path d="M5 9 H27 M5 16 H23 Q27 16 27 20 Q27 24 23 24 H17" stroke="${K}" stroke-width="2" fill="none"/><path d="M19 21 L16 24 L19 27" stroke="${K}" stroke-width="2" fill="none"/>`,
    merge: `<rect x="4" y="8" width="24" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M8 16 H24 M11 13 L8 16 L11 19 M21 13 L24 16 L21 19" stroke="${B}" stroke-width="2" fill="none"/>`,
    percent: `<text x="16" y="23" font-size="18" font-weight="700" text-anchor="middle" fill="${K}">%</text>`,
    comma: `<text x="16" y="24" font-size="22" font-weight="700" text-anchor="middle" fill="${K}">,</text>`,
    won: `<text x="16" y="23" font-size="17" font-weight="700" text-anchor="middle" fill="${K}">₩</text>`,
    decInc: `<text x="16" y="22" font-size="11" text-anchor="middle" fill="${K}">.0→.00</text>`,
    decDec: `<text x="16" y="22" font-size="11" text-anchor="middle" fill="${K}">.00→.0</text>`,
    sortAZ: `<text x="11" y="15" font-size="11" text-anchor="middle" fill="${K}">가</text><text x="11" y="28" font-size="11" text-anchor="middle" fill="${K}">하</text><path d="M24 5 V27 M20 23 L24 27 L28 23" stroke="${B}" stroke-width="2" fill="none"/>`,
    sortZA: `<text x="11" y="15" font-size="11" text-anchor="middle" fill="${K}">하</text><text x="11" y="28" font-size="11" text-anchor="middle" fill="${K}">가</text><path d="M24 5 V27 M20 23 L24 27 L28 23" stroke="${B}" stroke-width="2" fill="none"/>`,
    sort: `<rect x="4" y="5" width="14" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 12 H18 M4 19 H18" stroke="${K}"/><path d="M24 5 V27 M20 23 L24 27 L28 23" stroke="${B}" stroke-width="2" fill="none"/>`,
    filter: `<path d="M5 6 H27 L18 17 V27 L14 24 V17 Z" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    pivot: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 11 H28 M11 4 V28" stroke="${K}" stroke-width="2"/><path d="M15 20 Q20 14 24 18" stroke="${A}" stroke-width="2" fill="none"/><path d="M22 15 L24 18 L21 19" stroke="${A}" stroke-width="2" fill="none"/>`,
    table: `<rect x="4" y="6" width="24" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="6" width="24" height="5" fill="${B}"/><path d="M4 16 H28 M4 21 H28 M12 6 V26 M20 6 V26" stroke="${K}"/>`,
    recchart: `<rect x="4" y="16" width="5" height="12" fill="${B}"/><rect x="12" y="9" width="5" height="19" fill="${B}"/><rect x="20" y="13" width="5" height="15" fill="${B}"/><path d="M22 3 l2 4 4 1 -3 3 1 4 -4 -2 -4 2 1 -4 -3 -3 4 -1 z" fill="#f0a30a"/>`,
    colchart: `<path d="M4 28 H28" stroke="${K}"/><rect x="6" y="14" width="5" height="14" fill="${B}"/><rect x="13" y="7" width="5" height="21" fill="${B}"/><rect x="20" y="17" width="5" height="11" fill="${B}"/>`,
    linechart: `<path d="M4 28 H28 M4 4 V28" stroke="${K}"/><path d="M6 22 L13 13 L19 18 L27 7" stroke="${B}" stroke-width="2.5" fill="none"/>`,
    combochart: `<path d="M4 28 H28" stroke="${K}"/><rect x="6" y="16" width="5" height="12" fill="${B}"/><rect x="14" y="10" width="5" height="18" fill="${B}"/><rect x="22" y="18" width="5" height="10" fill="${B}"/><path d="M5 12 L14 6 L24 11" stroke="#ed7d31" stroke-width="2.5" fill="none"/>`,
    piechart: `<circle cx="16" cy="16" r="11" fill="${B}"/><path d="M16 16 V5 A11 11 0 0 1 27 16 Z" fill="#ed7d31"/>`,
    margins: `<rect x="6" y="3" width="20" height="26" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="10" y="9" width="12" height="15" fill="none" stroke="${B}" stroke-dasharray="2 2"/>`,
    orient: `<rect x="5" y="9" width="22" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="11" y="3" width="10" height="26" fill="none" stroke="${B}" stroke-width="2"/>`,
    paper: `<path d="M7 3 H20 L26 9 V29 H7 Z" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M20 3 V9 H26" fill="none" stroke="${K}" stroke-width="2"/>`,
    printarea: `<rect x="5" y="5" width="22" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="9" y="9" width="12" height="12" fill="none" stroke="${B}" stroke-width="2" stroke-dasharray="3 2"/>`,
    breaks: `<rect x="6" y="3" width="20" height="26" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M3 16 H29" stroke="${B}" stroke-width="2" stroke-dasharray="3 2"/>`,
    printtitle: `<rect x="5" y="4" width="22" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="5" y="4" width="22" height="6" fill="${B}"/>`,
    save: `<rect x="5" y="5" width="22" height="22" fill="#fff" stroke="currentColor" stroke-width="2"/><rect x="10" y="5" width="12" height="7" fill="currentColor"/><rect x="9" y="17" width="14" height="10" fill="none" stroke="currentColor" stroke-width="2"/>`,
    undo: `<path d="M9 13 H21 A6 6 0 0 1 21 25 H12" stroke="currentColor" stroke-width="2.5" fill="none"/><path d="M13 8 L8 13 L13 18" stroke="currentColor" stroke-width="2.5" fill="none"/>`,
    redo: `<path d="M23 13 H11 A6 6 0 0 0 11 25 H20" stroke="currentColor" stroke-width="2.5" fill="none"/><path d="M19 8 L24 13 L19 18" stroke="currentColor" stroke-width="2.5" fill="none"/>`,
    sum: `<text x="16" y="25" font-size="24" text-anchor="middle" fill="${K}">Σ</text>`,
    insertCells: `<rect x="5" y="9" width="22" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M16 9 V25 M5 17 H27" stroke="${K}"/><path d="M22 3 V9 M19 6 H25" stroke="${A}" stroke-width="2"/>`,
    deleteCells: `<rect x="5" y="9" width="22" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M16 9 V25 M5 17 H27" stroke="${K}"/><path d="M19 3 L25 9 M25 3 L19 9" stroke="#c00000" stroke-width="2"/>`,
    formatCells: `<rect x="5" y="9" width="22" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M16 9 V25 M5 17 H27" stroke="${K}"/><rect x="18" y="2" width="9" height="7" fill="#f3c86b"/>`,
    find: `<circle cx="13" cy="13" r="7" fill="none" stroke="${K}" stroke-width="2.5"/><path d="M18 18 L27 27" stroke="${K}" stroke-width="3"/>`,
    condfmt: `<rect x="4" y="5" width="24" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="7" y="9" width="10" height="4" fill="#c00000"/><rect x="7" y="15" width="16" height="4" fill="#f0a30a"/><rect x="7" y="21" width="6" height="4" fill="${A}"/>`,
    cellstyle: `<rect x="4" y="9" width="24" height="14" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="9" width="12" height="14" fill="#ddebf7"/>`,
    refresh: `<path d="M24 12 A9 9 0 1 0 25 19" stroke="${A}" stroke-width="2.5" fill="none"/><path d="M25 5 V12 H18" stroke="${A}" stroke-width="2.5" fill="none"/>`,
    group: `<rect x="4" y="4" width="11" height="11" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="17" y="17" width="11" height="11" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="2" y="2" width="28" height="28" fill="none" stroke="${B}" stroke-dasharray="3 2"/>`,
    ungroup: `<rect x="4" y="4" width="11" height="11" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="17" y="17" width="11" height="11" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    fieldlist: `<rect x="5" y="3" width="22" height="26" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M9 9 H23 M9 15 H23 M9 21 H23" stroke="${B}" stroke-width="2"/>`,
    plusminus: `<rect x="4" y="6" width="10" height="10" fill="#fff" stroke="${K}"/><path d="M6 11 H12 M9 8 V14" stroke="${K}" stroke-width="2"/><rect x="18" y="16" width="10" height="10" fill="#fff" stroke="${K}"/><path d="M20 21 H26" stroke="${K}" stroke-width="2"/>`,
    headers: `<rect x="4" y="6" width="24" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="6" width="24" height="6" fill="#ddebf7"/>`,
    options: `<circle cx="16" cy="16" r="6" fill="none" stroke="${K}" stroke-width="3"/><path d="M16 3 V8 M16 24 V29 M3 16 H8 M24 16 H29" stroke="${K}" stroke-width="3"/>`,
    expand: `<rect x="6" y="6" width="20" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M10 16 H22 M16 10 V22" stroke="${A}" stroke-width="3"/>`,
    collapse: `<rect x="6" y="6" width="20" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M10 16 H22" stroke="${A}" stroke-width="3"/>`,
    layout: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 11 H28 M12 4 V28" stroke="${K}"/><path d="M15 16 H25 M15 21 H25" stroke="${B}" stroke-width="2"/>`,
    subtotal: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 18 H28" stroke="${B}" stroke-width="3"/>`,
    grand: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 24 H28 M22 4 V28" stroke="${B}" stroke-width="3"/>`,
    blankrow: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 12 H28 M4 20 H28" stroke="${K}" stroke-dasharray="2 2"/>`,
    plus: `<path d="M16 6 V26 M6 16 H26" stroke="${A}" stroke-width="3"/>`,
    quicklayout: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="8" y="16" width="4" height="9" fill="${B}"/><rect x="14" y="11" width="4" height="14" fill="${B}"/><rect x="20" y="8" width="5" height="4" fill="${Gy}"/>`,
    colors: `<circle cx="11" cy="12" r="6" fill="${B}"/><circle cx="21" cy="12" r="6" fill="#ed7d31"/><circle cx="16" cy="21" r="6" fill="${A}"/>`,
    switchrc: `<path d="M6 10 H24 M20 6 L24 10 L20 14 M26 22 H8 M12 18 L8 22 L12 26" stroke="${K}" stroke-width="2" fill="none"/>`,
    selectdata: `<rect x="4" y="5" width="18" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 12 H22 M4 19 H22 M13 5 V27" stroke="${K}"/><rect x="18" y="16" width="10" height="12" fill="${B}"/>`,
    changetype: `<rect x="6" y="14" width="5" height="14" fill="${B}"/><rect x="13" y="8" width="5" height="20" fill="${B}"/><path d="M20 10 A7 7 0 1 1 27 18" stroke="${A}" stroke-width="2" fill="none"/>`,
    movechart: `<rect x="4" y="10" width="18" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M18 6 H27 V15 M27 6 L19 14" stroke="${B}" stroke-width="2" fill="none"/>`,
    tableA: `<rect x="4" y="6" width="24" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="6" width="24" height="5" fill="#a4373a"/><path d="M4 16 H28 M4 21 H28 M14 11 V26" stroke="${K}"/>`,
    query: `<rect x="4" y="6" width="17" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 12 H21 M4 18 H21" stroke="${K}"/><circle cx="21" cy="20" r="5" fill="#fff" stroke="${B}" stroke-width="2.5"/><path d="M24 23 L28 27" stroke="${B}" stroke-width="3"/>`,
    form: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="8" y="9" width="7" height="3" fill="${K}"/><rect x="16" y="8" width="9" height="5" fill="none" stroke="${K}"/><rect x="8" y="17" width="7" height="3" fill="${K}"/><rect x="16" y="16" width="9" height="5" fill="none" stroke="${K}"/>`,
    report: `<path d="M6 3 H21 L26 8 V29 H6 Z" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M10 12 H22 M10 17 H22 M10 22 H18" stroke="#a4373a" stroke-width="2"/>`,
    relations: `<rect x="3" y="5" width="11" height="13" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="18" y="14" width="11" height="13" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M14 11 H16 V20 H18" stroke="#a4373a" stroke-width="2" fill="none"/>`,
    key: `<circle cx="11" cy="12" r="6" fill="#f0c419" stroke="${K}" stroke-width="2"/><path d="M15 16 L27 28 M22 23 L25 20 M25 26 L28 23" stroke="${K}" stroke-width="2.5"/>`,
    run: `<text x="16" y="27" font-size="28" font-weight="900" text-anchor="middle" fill="#c00000">!</text>`,
    view: `<rect x="4" y="5" width="24" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 11 H28 M12 11 V27" stroke="${K}"/><path d="M15 17 H25 M15 22 H25" stroke="${B}" stroke-width="2"/>`,
    sql: `<text x="16" y="21" font-size="11" font-weight="700" text-anchor="middle" fill="${K}">SQL</text><rect x="3" y="7" width="26" height="18" fill="none" stroke="${K}" stroke-width="1.5"/>`,
    propsheet: `<rect x="5" y="3" width="22" height="26" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M14 3 V29 M5 10 H27 M5 16 H27 M5 22 H27" stroke="${K}"/>`,
    label: `<text x="16" y="23" font-size="17" text-anchor="middle" fill="${K}">Aa</text>`,
    textboxC: `<rect x="3" y="8" width="26" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><text x="13" y="21" font-size="11" text-anchor="middle" fill="${K}">ab</text><path d="M23 11 V21" stroke="${K}"/>`,
    listbox: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M8 10 H20 M8 16 H20 M8 22 H20" stroke="${K}" stroke-width="2"/><rect x="22" y="4" width="6" height="24" fill="#ddd"/>`,
    lineC: `<path d="M5 26 L27 6" stroke="${K}" stroke-width="2.5"/>`,
    rectC: `<rect x="5" y="8" width="22" height="16" fill="none" stroke="${K}" stroke-width="2.5"/>`,
    selectArrow: `<path d="M9 4 L9 26 L14 21 L18 29 L21 28 L17 20 L24 20 Z" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    wand: `<path d="M5 27 L22 10" stroke="${K}" stroke-width="3"/><path d="M24 3 l1.5 3 3 1 -3 1.5 -1.5 3 -1.5 -3 -3 -1.5 3 -1 z" fill="#f0a30a"/>`,
    grouping: `<rect x="4" y="5" width="24" height="6" fill="#fff3d6" stroke="${K}"/><rect x="9" y="14" width="19" height="5" fill="#fff" stroke="${K}"/><rect x="9" y="22" width="19" height="5" fill="#fff" stroke="${K}"/>`,
    totals: `<text x="16" y="25" font-size="24" text-anchor="middle" fill="#a4373a">Σ</text>`,
    preview: `<path d="M6 3 H20 L26 9 V29 H6 Z" fill="#fff" stroke="${K}" stroke-width="2"/><circle cx="15" cy="17" r="5" fill="none" stroke="${B}" stroke-width="2"/><path d="M19 21 L24 26" stroke="${B}" stroke-width="2.5"/>`,
    pagesetup: `<path d="M6 3 H20 L26 9 V29 H6 Z" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="10" y="12" width="12" height="12" fill="none" stroke="${B}" stroke-dasharray="2 2"/>`,
    printer: `<rect x="8" y="3" width="16" height="9" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="3" y="12" width="26" height="11" rx="2" fill="#ddd" stroke="${K}" stroke-width="2"/><rect x="8" y="19" width="16" height="10" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    close: `<path d="M8 8 L24 24 M24 8 L8 24" stroke="#c00000" stroke-width="3"/>`,
    zoom: `<circle cx="13" cy="13" r="8" fill="none" stroke="${K}" stroke-width="2.5"/><path d="M19 19 L27 27" stroke="${K}" stroke-width="3"/>`,
    onepage: `<rect x="9" y="4" width="14" height="24" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    twopage: `<rect x="3" y="6" width="12" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="17" y="6" width="12" height="20" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    newslide: `<rect x="3" y="7" width="22" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M25 18 V28 M20 23 H30" stroke="${R}" stroke-width="2.5"/>`,
    slidelayout: `<rect x="3" y="6" width="26" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="6" y="9" width="20" height="4" fill="${Gy}"/><rect x="6" y="15" width="9" height="8" fill="${Gy}"/><rect x="17" y="15" width="9" height="8" fill="${Gy}"/>`,
    shapes: `<rect x="3" y="14" width="13" height="13" fill="${R}"/><circle cx="21" cy="11" r="8" fill="#ffd166" stroke="${K}"/><path d="M14 4 L20 14 H8 Z" fill="#fff" stroke="${K}"/>`,
    textboxP: `<rect x="4" y="6" width="24" height="20" fill="#fff" stroke="${K}" stroke-width="2" stroke-dasharray="3 2"/><text x="16" y="22" font-size="13" text-anchor="middle" fill="${K}">가</text>`,
    arrange: `<rect x="4" y="4" width="15" height="15" fill="${Gy}"/><rect x="13" y="13" width="15" height="15" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    shapeFill: `<path d="M8 14 L16 6 L25 15 L17 23 Z" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="25" width="24" height="5" fill="#ffd400"/>`,
    shapeLine: `<path d="M6 22 L20 8 L24 12 L10 26 Z" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="27" width="24" height="3" fill="${K}"/>`,
    shapeEffect: `<rect x="8" y="8" width="16" height="16" fill="${Gy}"/><rect x="5" y="5" width="16" height="16" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    alignObj: `<path d="M5 4 V28" stroke="${K}" stroke-width="2"/><rect x="8" y="7" width="16" height="6" fill="${B}"/><rect x="8" y="18" width="10" height="6" fill="${B}"/>`,
    front: `<rect x="4" y="4" width="15" height="15" fill="#ddd" stroke="${K}"/><rect x="12" y="12" width="15" height="15" fill="${B}"/>`,
    back: `<rect x="12" y="12" width="15" height="15" fill="#ddd" stroke="${K}"/><rect x="4" y="4" width="15" height="15" fill="${B}" opacity=".55"/>`,
    selpane: `<rect x="5" y="3" width="22" height="26" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M9 9 H19 M9 15 H19 M9 21 H19" stroke="${K}" stroke-width="2"/><circle cx="23" cy="9" r="1.5" fill="${K}"/><circle cx="23" cy="15" r="1.5" fill="${K}"/><circle cx="23" cy="21" r="1.5" fill="${K}"/>`,
    slidesize: `<rect x="3" y="8" width="26" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="8" y="5" width="16" height="22" fill="none" stroke="${R}" stroke-width="2"/>`,
    handout: `<rect x="7" y="2" width="18" height="28" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="10" y="5" width="12" height="9" fill="${Gy}"/><rect x="10" y="17" width="12" height="9" fill="${Gy}"/>`,
    grayscale: `<rect x="4" y="6" width="8" height="20" fill="#333"/><rect x="12" y="6" width="8" height="20" fill="#888"/><rect x="20" y="6" width="8" height="20" fill="#ddd" stroke="${K}"/>`,
    bullets: `<circle cx="7" cy="9" r="2.5" fill="${K}"/><circle cx="7" cy="16" r="2.5" fill="${K}"/><circle cx="7" cy="23" r="2.5" fill="${K}"/><path d="M12 9 H27 M12 16 H27 M12 23 H27" stroke="${K}" stroke-width="2"/>`,
    numbering: `<text x="6" y="12" font-size="8" fill="${K}">1</text><text x="6" y="19" font-size="8" fill="${K}">2</text><text x="6" y="26" font-size="8" fill="${K}">3</text><path d="M12 9 H27 M12 16 H27 M12 23 H27" stroke="${K}" stroke-width="2"/>`,
    growFont: `<text x="12" y="25" font-size="18" text-anchor="middle" fill="${K}">가</text><path d="M24 12 L27 7 L30 12" stroke="${K}" fill="none" stroke-width="2"/>`,
    shrinkFont: `<text x="12" y="25" font-size="14" text-anchor="middle" fill="${K}">가</text><path d="M24 7 L27 12 L30 7" stroke="${K}" fill="none" stroke-width="2"/>`,
    guides: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M16 4 V28 M4 16 H28" stroke="#c00000" stroke-dasharray="3 2"/>`,
    gridlines: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M12 4 V28 M20 4 V28 M4 12 H28 M4 20 H28" stroke="${Gy}"/>`,
    ruler: `<rect x="3" y="11" width="26" height="10" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M8 11 V15 M13 11 V17 M18 11 V15 M23 11 V17" stroke="${K}"/>`,
    master: `<rect x="3" y="6" width="22" height="16" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="7" y="10" width="22" height="16" fill="#fff" stroke="${R}" stroke-width="2"/>`,
    closeView: `<rect x="5" y="5" width="22" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M11 11 L21 21 M21 11 L11 21" stroke="#c00000" stroke-width="3"/>`,
    newfile: `<path d="M8 3 H20 L26 9 V29 H8 Z" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M17 15 V25 M12 20 H22" stroke="${A}" stroke-width="2.5"/>`,
    db: `<ellipse cx="16" cy="8" rx="10" ry="4" fill="#fff" stroke="#a4373a" stroke-width="2"/><path d="M6 8 V24 A10 4 0 0 0 26 24 V8" fill="#fff" stroke="#a4373a" stroke-width="2"/>`,
    folder: `<path d="M3 9 H13 L16 12 H29 V27 H3 Z" fill="#f3c86b" stroke="${K}" stroke-width="1.5"/>`,
    datasheet: `<rect x="4" y="5" width="24" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M4 11 H28 M4 17 H28 M4 22 H28 M12 5 V27 M20 5 V27" stroke="${K}"/>`,
    design: `<path d="M5 27 L8 19 L22 5 L27 10 L13 24 Z" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M19 8 L24 13" stroke="${K}"/>`,
    formview: `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="8" y="9" width="16" height="4" fill="${B}"/><rect x="8" y="16" width="16" height="4" fill="none" stroke="${K}"/>`,
    insertRow: `<rect x="5" y="12" width="22" height="14" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M5 19 H27" stroke="${K}"/><path d="M16 2 V10 M12 6 H20" stroke="${A}" stroke-width="2.5"/>`,
    deleteRow: `<rect x="5" y="12" width="22" height="14" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M5 19 H27" stroke="${K}"/><path d="M12 3 L20 10 M20 3 L12 10" stroke="#c00000" stroke-width="2.5"/>`,
    addTable: `<rect x="4" y="6" width="18" height="18" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="6" width="18" height="5" fill="#a4373a"/><path d="M25 17 V29 M19 23 H31" stroke="${A}" stroke-width="2.5"/>`,
    clearLayout: `<rect x="4" y="6" width="24" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M10 12 L22 22 M22 12 L10 22" stroke="#c00000" stroke-width="2"/>`,
    builder: `<path d="M5 26 L22 9" stroke="${K}" stroke-width="3"/><circle cx="24" cy="7" r="4" fill="#f0a30a"/>`,
    existing: `<rect x="5" y="3" width="22" height="26" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M9 9 H23 M9 15 H23 M9 21 H23" stroke="#a4373a" stroke-width="2"/><path d="M24 22 V30 M20 26 H28" stroke="${A}" stroke-width="2"/>`,
    tab: `<rect x="4" y="8" width="24" height="18" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="4" y="4" width="10" height="6" fill="#fff" stroke="${K}" stroke-width="2"/>`,
    combo: `<rect x="3" y="10" width="26" height="12" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M21 14 L24 18 L27 14" fill="${K}"/>`,
    button: `<rect x="3" y="9" width="26" height="14" rx="3" fill="#eee" stroke="${K}" stroke-width="2"/><text x="16" y="20" font-size="8" text-anchor="middle" fill="${K}">XXX</text>`,
    checkbox: `<rect x="7" y="7" width="18" height="18" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M10 16 L14 20 L22 11" stroke="${A}" stroke-width="3" fill="none"/>`,
    image: `<rect x="4" y="6" width="24" height="20" fill="#fff" stroke="${K}" stroke-width="2"/><circle cx="11" cy="12" r="3" fill="#f0a30a"/><path d="M6 24 L14 16 L19 21 L22 18 L27 24 Z" fill="${A}"/>`,
    chartIcon: `<rect x="4" y="16" width="5" height="12" fill="#a4373a"/><rect x="12" y="9" width="5" height="19" fill="#a4373a"/><rect x="20" y="13" width="5" height="15" fill="#a4373a"/>`,
    pageNum: `<path d="M6 3 H20 L26 9 V29 H6 Z" fill="#fff" stroke="${K}" stroke-width="2"/><text x="16" y="24" font-size="10" text-anchor="middle" fill="${K}">#</text>`,
    dateTime: `<rect x="5" y="6" width="22" height="21" fill="#fff" stroke="${K}" stroke-width="2"/><path d="M5 12 H27" stroke="#c00000" stroke-width="3"/>`,
    logo: `<rect x="5" y="5" width="22" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><circle cx="16" cy="16" r="5" fill="${Gy}"/>`,
    titleP: `<rect x="4" y="5" width="24" height="22" fill="#fff" stroke="${K}" stroke-width="2"/><rect x="7" y="9" width="18" height="4" fill="${K}"/>`,
    dot: `<circle cx="16" cy="16" r="3" fill="${K}"/>`,
  };
  function icon(name, size) {
    const s = document.createElementNS(svgNS, 'svg');
    s.setAttribute('viewBox', '0 0 32 32');
    s.setAttribute('class', 'ic');
    if (size) { s.setAttribute('width', size); s.setAttribute('height', size); }
    s.innerHTML = I[name] || `<rect x="6" y="6" width="20" height="20" rx="3" fill="#fff" stroke="${K}" stroke-width="2"/><text x="16" y="21" font-size="12" text-anchor="middle" fill="${K}">${(name || '?').slice(0, 2)}</text>`;
    return s;
  }

  // ------------------------------------------------------------------ 메뉴
  let openMenus = [];
  function closeMenus(level) {
    while (openMenus.length > (level || 0)) { const m = openMenus.pop(); m.el.remove(); }
    document.removeEventListener('mousedown', outside, true);
    if (openMenus.length) document.addEventListener('mousedown', outside, true);
  }
  function outside(e) { if (!openMenus.some(m => m.el.contains(e.target))) closeMenus(0); }
  function menu(anchor, items, opts) {
    const level = (opts && opts.level) || 0;
    closeMenus(level);
    const el = h('div', { class: 'menu' });
    const list = typeof items === 'function' ? items() : items;
    for (const it of list) {
      if (!it) continue;
      if (it.sep) { el.appendChild(h('div', { class: 'msep' })); continue; }
      if (it.header) { el.appendChild(h('div', { class: 'mhead' }, it.header)); continue; }
      if (it.html) { el.appendChild(typeof it.html === 'function' ? it.html(() => closeMenus(0)) : it.html); continue; }
      const dis = typeof it.disabled === 'function' ? it.disabled() : it.disabled;
      const mi = h('div', { class: 'mi' + (dis ? ' dis' : ''), title: it.tip || '' },
        h('span', { class: 'ck' }, (typeof it.checked === 'function' ? it.checked() : it.checked) ? '✓' : ''),
        it.icon ? icon(it.icon) : h('span', { class: 'ic' }), h('span', {}, it.label), it.sub ? h('span', { class: 'arrow' }, '▶') : null);
      if (!dis) {
        if (it.sub) {
          const open = () => { [...el.children].forEach(c => c.classList.remove('hot')); mi.classList.add('hot'); const r = mi.getBoundingClientRect(); menu({ getBoundingClientRect: () => ({ left: r.right - 2, right: r.right, top: r.top - 4, bottom: r.top }) }, it.sub, { level: level + 1, side: true }); };
          mi.addEventListener('mouseenter', open);
          mi.addEventListener('click', open);
        } else {
          mi.addEventListener('mouseenter', () => closeMenus(level + 1));
          mi.addEventListener('click', e => { e.stopPropagation(); closeMenus(0); it.onClick && it.onClick(); });
        }
      }
      if (it.onContext) mi.addEventListener('contextmenu', e => { e.preventDefault(); it.onContext(e); });
      el.appendChild(mi);
    }
    document.body.appendChild(el);
    const r = anchor.getBoundingClientRect ? anchor.getBoundingClientRect() : anchor;
    let x = opts && opts.at ? opts.at.x : (opts && opts.side ? r.left : r.left), y = opts && opts.at ? opts.at.y : (opts && opts.side ? r.top : r.bottom);
    const mw = el.offsetWidth, mh = el.offsetHeight;
    if (x + mw > innerWidth - 4) x = Math.max(4, (opts && opts.side ? r.left - mw - (r.right - r.left) : innerWidth - mw - 4));
    if (y + mh > innerHeight - 4) y = Math.max(4, innerHeight - mh - 4);
    el.style.left = x + 'px'; el.style.top = y + 'px';
    openMenus.push({ el, level });
    setTimeout(() => document.addEventListener('mousedown', outside, true), 0);
    return el;
  }

  // ------------------------------------------------------------------ 리본
  function ribbon(host, spec) {
    let active = spec.initial || spec.tabs[0].id;
    const tabsEl = h('div', { class: 'ribbon-tabs' });
    const body = h('div', { class: 'ribbon' });
    host.appendChild(tabsEl);
    host.appendChild(body);
    function visibleTabs() { return spec.tabs.filter(t => !t.when || t.when()); }
    function renderTabs() {
      tabsEl.innerHTML = '';
      tabsEl.appendChild(h('button', { class: 'file', onclick: () => spec.file() }, '파일'));
      const vis = visibleTabs();
      if (!vis.some(t => t.id === active)) active = vis.find(t => !t.ctx) ? vis.find(t => !t.ctx).id : vis[0].id;
      let lastCtx = null;
      for (const t of vis) {
        if (t.ctx && t.ctx !== lastCtx) { tabsEl.appendChild(h('span', { class: 'ctxlabel' }, t.ctx)); lastCtx = t.ctx; }
        tabsEl.appendChild(h('button', { class: (t.id === active ? 'on ' : '') + (t.ctx ? 'ctx' : ''), onclick: () => { active = t.id; render(); } }, t.label));
      }
    }
    function val(v) { return typeof v === 'function' ? v() : v; }
    function item(it) {
      if (it.type === 'stack') return h('div', { class: 'rstack' }, it.items.map(item));
      if (it.type === 'row') return h('div', { class: 'rrow' }, it.items.map(item));
      if (it.type === 'combo') {
        const inp = h('input', { value: val(it.value) == null ? '' : val(it.value), title: it.tip || it.label || '', onkeydown: e => { if (e.key === 'Enter') { it.onChange(inp.value); e.preventDefault(); } e.stopPropagation(); }, onchange: () => it.onChange(inp.value) });
        const wrap = h('span', { class: 'rcombo', style: { width: (it.width || 60) + 'px' } }, inp,
          h('button', { onclick: e => menu(wrap, (val(it.list) || []).map(v => ({ label: String(v), checked: String(v) === String(val(it.value)), onClick: () => it.onChange(v) }))) }, '▼'));
        if (val(it.disabled)) { inp.disabled = true; }
        return wrap;
      }
      if (it.type === 'check') {
        const cb = h('input', { type: 'checkbox', checked: !!val(it.checked), disabled: !!val(it.disabled), onchange: () => it.onChange(cb.checked) });
        return h('label', { class: 'rcheck', title: it.tip || '' }, cb, it.label);
      }
      if (it.type === 'html') return it.render();
      const dis = val(it.disabled) || it.out;
      const b = h('button', { class: 'rb ' + (it.type === 'big' ? 'big ' : '') + (val(it.pressed) ? 'on ' : '') + (it.menu && !it.onClick ? '' : ''), disabled: !!dis,
        title: it.out ? (it.label || '') + ' — 이 시뮬레이터에서 다루지 않는 기능' : (it.tip || it.label || '') });
      if (it.icon) b.appendChild(icon(it.icon));
      if (it.label && it.type !== 'icon') b.appendChild(h('span', {}, it.label));
      if (it.menu) b.appendChild(h('span', { class: 'dd' }, '▼'));
      b.addEventListener('click', () => {
        if (it.menu && !it.onClick) { menu(b, it.menu); return; }
        it.onClick && it.onClick(b);
      });
      if (it.menu && it.onClick) {                      // 분할 단추: 본체는 실행, ▼는 목록
        const wrap = h('span', { style: { display: 'inline-flex', flexDirection: it.type === 'big' ? 'column' : 'row' } });
        b.lastChild.remove();
        const dd = h('button', { class: 'rb', disabled: !!dis, style: it.type === 'big' ? { height: '18px', justifyContent: 'center' } : { padding: '1px' }, onclick: () => menu(wrap, it.menu), title: it.label + ' 목록' }, h('span', { class: 'dd' }, '▼'));
        if (it.type === 'big') { b.style.height = '52px'; }
        wrap.appendChild(b); wrap.appendChild(dd);
        return wrap;
      }
      return b;
    }
    function render() {
      renderTabs();
      body.innerHTML = '';
      const t = spec.tabs.find(x => x.id === active);
      if (!t) return;
      for (const g of t.groups) {
        if (g.when && !g.when()) continue;
        const grp = h('div', { class: 'rgroup' }, h('div', { class: 'ritems' }, g.items.map(item)),
          h('div', { class: 'rlabel' }, g.label, g.launcher ? h('button', { class: 'launch', title: g.label + ' 대화상자 표시', onclick: g.launcher }, '⇲') : null));
        body.appendChild(grp);
      }
      if (t.note) body.appendChild(h('div', { style: { padding: '8px 12px', color: '#777', alignSelf: 'center' } }, t.note));
    }
    render();
    return { render, select: id => { active = id; render(); }, get active() { return active; } };
  }

  // ------------------------------------------------------------------ 대화상자
  const stack = [];
  function dialog(o) {
    const back = h('div', { class: 'modal-back' });
    const title = h('div', { class: 'dlg-title' }, h('span', { class: 't' }, o.title), o.help !== false ? h('button', { class: 'help', title: '도움말' }, '?') : null, h('button', { title: '닫기', onclick: () => cancel() }, '✕'));
    const bodyEl = h('div', { class: 'dlg-body' });
    const foot = h('div', { class: 'dlg-foot' });
    const el = h('div', { class: 'dlg', style: { width: o.width ? o.width + 'px' : 'auto' } }, title, bodyEl, foot);
    const api = { el, body: bodyEl, close, setBody, hide, show, buttons: {} };
    function setBody(b) { bodyEl.innerHTML = ''; bodyEl.appendChild(typeof b === 'function' ? b(api) : b); }
    setBody(o.body || h('div'));
    const left = h('div', { class: 'left' });
    foot.appendChild(left);
    for (const bt of o.buttons || [{ label: '확인', primary: true }, { label: '취소', cancel: true }]) {
      const b = h('button', { class: 'db' + (bt.primary ? ' primary' : ''), disabled: !!bt.disabled }, bt.label);
      b.addEventListener('click', () => { if (bt.cancel) return cancel(); const r = bt.onClick ? bt.onClick(api) : undefined; if (r !== false) close(bt.label); });
      (bt.left ? left : foot).appendChild(b);
      api.buttons[bt.label] = b;
    }
    document.body.appendChild(back);
    document.body.appendChild(el);
    const place = () => { el.style.left = Math.max(10, (innerWidth - el.offsetWidth) / 2) + 'px'; el.style.top = Math.max(10, (innerHeight - el.offsetHeight) / 2.4) + 'px'; };
    place();
    // 끌기
    title.addEventListener('mousedown', e => {
      if (e.target.tagName === 'BUTTON') return;
      const sx = e.clientX - el.offsetLeft, sy = e.clientY - el.offsetTop;
      const mv = ev => { el.style.left = (ev.clientX - sx) + 'px'; el.style.top = Math.max(0, ev.clientY - sy) + 'px'; };
      const up = () => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); };
      addEventListener('mousemove', mv); addEventListener('mouseup', up);
    });
    const onKey = e => {
      if (stack[stack.length - 1] !== api || api._hidden) return;
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); cancel(); }
      if (e.key === 'Enter' && !(e.target.tagName === 'TEXTAREA') && !e.target.dataset.noenter) {
        const pb = (o.buttons || []).find(b => b.primary);
        if (pb) { e.preventDefault(); e.stopPropagation(); api.buttons[pb.label].click(); }
      }
      e.stopPropagation();
    };
    el.addEventListener('keydown', onKey);
    stack.push(api);
    let done = false;
    function close(res) {
      if (done) return; done = true;
      back.remove(); el.remove();
      const i = stack.indexOf(api); if (i >= 0) stack.splice(i, 1);
      o.onClose && o.onClose(res);
    }
    function cancel() { if (o.onCancel && o.onCancel(api) === false) return; close(null); }
    function hide() { api._hidden = true; el.style.display = 'none'; back.style.display = 'none'; }
    function show() { api._hidden = false; el.style.display = ''; back.style.display = ''; }
    setTimeout(() => { const f = el.querySelector('[autofocus]') || el.querySelector('input:not([type=checkbox]):not([type=radio]), select, textarea'); if (f) { f.focus(); if (f.select) f.select(); } }, 0);
    api.place = place;
    return api;
  }
  function anyDialog() { return stack.length > 0 && !stack[stack.length - 1]._hidden; }
  function msg(o) {
    return new Promise(res => {
      const btns = (o.buttons || ['확인']).map((label, i) => ({ label, primary: i === (o.def || 0), onClick: () => res(label) }));
      dialog({ title: o.title || 'Microsoft Office', width: o.width, help: false,
        body: h('div', { class: 'msgbox' }, h('div', { class: 'mic ' + (o.icon || 'warn') }, { warn: '!', err: '✕', info: 'i', q: '?' }[o.icon || 'warn']), h('div', {}, o.text)),
        buttons: btns, onCancel: () => { res(o.cancelValue !== undefined ? o.cancelValue : (o.buttons ? o.buttons[o.buttons.length - 1] : '확인')); } });
    });
  }
  function toast(text, ms) {
    const t = h('div', { class: 'toast' }, text);
    document.body.appendChild(t);
    setTimeout(() => t.remove(), ms || 2200);
  }
  /* 범위 선택 상자: 단추를 누르면 대화상자가 접히고 앱에서 셀을 고른 뒤 돌아온다 */
  function refEdit(o) {
    const inp = h('input', { type: 'text', value: o.value || '', style: { width: (o.width || 220) + 'px' } });
    const wrap = h('span', { class: 'refedit' }, inp, h('button', { title: '범위 선택', onclick: () => {
      if (!o.pick) return;
      const dlg = o.dialog();
      dlg.hide();
      const bar = h('div', { class: 'collapsed-dlg' }, h('span', {}, o.label || '범위'), h('input', { type: 'text', value: inp.value, style: { width: '220px' } }), h('button', { class: 'btn' }, '확인'));
      document.body.appendChild(bar);
      const live = bar.querySelector('input');
      const finish = () => { bar.remove(); stop(); dlg.show(); inp.value = live.value; o.onChange && o.onChange(inp.value); };
      bar.querySelector('button').onclick = finish;
      live.addEventListener('keydown', e => { if (e.key === 'Enter') finish(); });
      const stop = o.pick(v => { live.value = v; });
    } }, '↥'));
    wrap.input = inp;
    return wrap;
  }

  /* 탭 부품(대화상자 안) */
  function tabs(defs, initial) {
    const bar = h('div', { class: 'dtabs' });
    const panes = h('div');
    let cur = initial || defs[0].id;
    const show = id => { cur = id; [...bar.children].forEach((b, i) => b.classList.toggle('on', defs[i].id === id)); [...panes.children].forEach((p, i) => { p.hidden = defs[i].id !== id; }); };
    for (const d of defs) { bar.appendChild(h('button', { onclick: () => show(d.id) }, d.label)); panes.appendChild(d.body); }
    show(cur);
    const el = h('div', {}, bar, panes);
    el.show = show;
    el.current = () => cur;
    return el;
  }
  function labeled(text, input) { return h('label', { class: 'row' }, h('span', {}, text), input); }

  /* 색 고르기 판(테마 색·표준 색) */
  const THEME_COLORS = [['#FFFFFF', '흰색, 배경 1'], ['#000000', '검정, 텍스트 1'], ['#E7E6E6', '회색, 배경 2'], ['#44546A', '파랑 회색, 텍스트 2'], ['#4472C4', '파랑, 강조 1'],
    ['#ED7D31', '주황, 강조 2'], ['#A5A5A5', '회색, 강조 3'], ['#FFC000', '금색, 강조 4'], ['#5B9BD5', '파랑, 강조 5'], ['#70AD47', '녹색, 강조 6']];
  const SHADES = {
    '#FFFFFF': [['#F2F2F2', '5% 더 어둡게'], ['#D9D9D9', '15% 더 어둡게'], ['#BFBFBF', '25% 더 어둡게'], ['#A6A6A6', '35% 더 어둡게'], ['#808080', '50% 더 어둡게']],
    '#000000': [['#7F7F7F', '50% 더 밝게'], ['#595959', '35% 더 밝게'], ['#404040', '25% 더 밝게'], ['#262626', '15% 더 밝게'], ['#0D0D0D', '5% 더 밝게']],
  };
  function mix(hex, t) { // t>0 밝게, t<0 어둡게
    const n = parseInt(hex.slice(1), 16), r = n >> 16, g = (n >> 8) & 255, b = n & 255;
    const f = c => Math.round(t > 0 ? c + (255 - c) * t : c * (1 + t));
    return '#' + [f(r), f(g), f(b)].map(x => x.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function palette(onPick, extra) {
    return close => {
      const grid = h('div', { class: 'mgrid', style: { gridTemplateColumns: 'repeat(10, 18px)' } });
      for (const [c, nm] of THEME_COLORS) grid.appendChild(h('div', { class: 'swatch', style: { background: c }, title: nm, onclick: () => { close(); onPick(c, nm); } }));
      for (let row = 0; row < 5; row++) for (const [c, nm] of THEME_COLORS) {
        const sh = SHADES[c] ? SHADES[c][row] : [mix(c, [0.8, 0.6, 0.4, -0.25, -0.5][row]), ['80% 더 밝게', '60% 더 밝게', '40% 더 밝게', '25% 더 어둡게', '50% 더 어둡게'][row]];
        grid.appendChild(h('div', { class: 'swatch', style: { background: sh[0] }, title: nm + ', ' + sh[1], onclick: () => { close(); onPick(sh[0], nm + ', ' + sh[1]); } }));
      }
      const std = h('div', { class: 'mgrid', style: { gridTemplateColumns: 'repeat(10, 18px)' } });
      for (const [c, nm] of [['#C00000', '진한 빨강'], ['#FF0000', '빨강'], ['#FFC000', '주황'], ['#FFFF00', '노랑'], ['#92D050', '연한 녹색'], ['#00B050', '녹색'], ['#00B0F0', '연한 파랑'], ['#0070C0', '파랑'], ['#002060', '진한 파랑'], ['#7030A0', '자주']])
        std.appendChild(h('div', { class: 'swatch', style: { background: c }, title: nm, onclick: () => { close(); onPick(c, nm); } }));
      return h('div', {}, h('div', { class: 'mhead' }, '테마 색'), grid, h('div', { class: 'mhead' }, '표준 색'), std, extra ? extra(close) : null);
    };
  }
  function colorDialog(initial, onOk) {
    const n = parseInt((initial || '#999999').slice(1), 16);
    const r = h('input', { type: 'number', min: 0, max: 255, value: n >> 16, style: { width: '64px' } });
    const g = h('input', { type: 'number', min: 0, max: 255, value: (n >> 8) & 255, style: { width: '64px' } });
    const b = h('input', { type: 'number', min: 0, max: 255, value: n & 255, style: { width: '64px' } });
    const sw = h('div', { style: { width: '70px', height: '50px', border: '1px solid #888' } });
    const hex = () => '#' + [r, g, b].map(x => Math.max(0, Math.min(255, +x.value || 0)).toString(16).padStart(2, '0')).join('').toUpperCase();
    const upd = () => { sw.style.background = hex(); };
    [r, g, b].forEach(x => x.addEventListener('input', upd));
    upd();
    const stdGrid = h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(9, 22px)', gap: '3px' } });
    ['#000000', '#262626', '#404040', '#595959', '#7F7F7F', '#999999', '#A6A6A6', '#BFBFBF', '#D9D9D9', '#FFFFFF', '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0'].forEach(c =>
      stdGrid.appendChild(h('div', { style: { width: '22px', height: '22px', background: c, border: '1px solid #aaa', cursor: 'pointer' }, onclick: () => { const v = parseInt(c.slice(1), 16); r.value = v >> 16; g.value = (v >> 8) & 255; b.value = v & 255; upd(); } })));
    const t = tabs([
      { id: 'std', label: '표준', body: h('div', { style: { padding: '8px' } }, h('div', { class: 'note' }, '색:'), stdGrid) },
      { id: 'cus', label: '사용자 지정', body: h('div', { style: { padding: '8px' } }, h('div', { class: 'row' }, '색 모델:', h('select', {}, h('option', {}, 'RGB'))),
        h('div', { class: 'grid2' }, '빨강(R):', r, '녹색(G):', g, '파랑(B):', b)) },
    ], 'cus');
    dialog({ title: '색', width: 400, body: h('div', { style: { display: 'flex', gap: '12px' } }, h('div', { style: { flex: 1 } }, t), h('div', { class: 'col' }, h('span', { class: 'note' }, '새 색'), sw)),
      buttons: [{ label: '확인', primary: true, onClick: () => onOk(hex()) }, { label: '취소', cancel: true }] });
  }

  /* 끌어 놓기 도우미 */
  function drag(e, onMove, onUp) {
    e.preventDefault();
    const mv = ev => onMove(ev);
    const up = ev => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); onUp && onUp(ev); };
    addEventListener('mousemove', mv); addEventListener('mouseup', up);
  }

  OA.ui = { h, icon, menu, closeMenus, ribbon, dialog, anyDialog, msg, toast, refEdit, tabs, labeled, palette, colorDialog, drag, mix, THEME_COLORS };
})(typeof window !== 'undefined' ? window : globalThis);

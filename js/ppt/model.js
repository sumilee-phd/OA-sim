/* PowerPoint 모델(DOM 없음): 슬라이드·도형(cm 좌표)·기본 서식·연결점과 꺾인 연결선 경로·그룹·순서·맞춤·유인물 배치·도형 경로(SVG 문자열). */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util;

  const SIZES = { wide: { w: 33.867, h: 19.05, label: '와이드스크린(16:9)' }, standard: { w: 25.4, h: 19.05, label: '표준(4:3)' } };
  const PT = 1 / 28.3465;                              // 1pt = 0.03528cm
  const THEME = { accent1: '#4472C4', accent1Dark: '#2F528F' };

  /* 도형 종류: 이름(원고 분석표의 PowerPoint 이름), 범주, 연결점 수 */
  const KINDS = {
    rect: { name: '사각형', cat: '사각형', sites: 4 },
    roundRect: { name: '모서리가 둥근 사각형', cat: '사각형', sites: 4, adj: [0.16667] },
    ellipse: { name: '타원', cat: '기본 도형', sites: 8 },
    triangle: { name: '이등변 삼각형', cat: '기본 도형', sites: 6 },
    parallelogram: { name: '평행 사변형', cat: '기본 도형', sites: 4 },
    trapezoid: { name: '사다리꼴', cat: '기본 도형', sites: 4 },
    diamond: { name: '다이아몬드', cat: '기본 도형', sites: 4 },
    pentagon: { name: '정오각형', cat: '기본 도형', sites: 5 },
    hexagon: { name: '육각형', cat: '기본 도형', sites: 6 },
    can: { name: '원통형', cat: '기본 도형', sites: 4 },
    cloud: { name: '구름', cat: '기본 도형', sites: 8 },
    leftBracket: { name: '왼쪽 대괄호', cat: '기본 도형', sites: 3 },
    rightBracket: { name: '오른쪽 대괄호', cat: '기본 도형', sites: 3 },
    leftBrace: { name: '왼쪽 중괄호', cat: '기본 도형', sites: 3 },
    rightBrace: { name: '오른쪽 중괄호', cat: '기본 도형', sites: 3 },
    rightArrow: { name: '오른쪽 화살표', cat: '블록 화살표', sites: 4 },
    leftArrow: { name: '왼쪽 화살표', cat: '블록 화살표', sites: 4 },
    upArrow: { name: '위쪽 화살표', cat: '블록 화살표', sites: 4 },
    downArrow: { name: '아래쪽 화살표', cat: '블록 화살표', sites: 4 },
    flowProcess: { name: '순서도: 처리', cat: '순서도', sites: 4 },
    flowAlt: { name: '순서도: 대체 처리', cat: '순서도', sites: 4 },
    flowDecision: { name: '순서도: 결정', cat: '순서도', sites: 4 },
    flowData: { name: '순서도: 데이터', cat: '순서도', sites: 4 },
    flowDocument: { name: '순서도: 문서', cat: '순서도', sites: 4 },
    flowTerminator: { name: '순서도: 종료', cat: '순서도', sites: 4 },
    flowPreparation: { name: '순서도: 준비', cat: '순서도', sites: 4 },
    flowManualInput: { name: '순서도: 수동 입력', cat: '순서도', sites: 4 },
    flowConnector: { name: '순서도: 연결자', cat: '순서도', sites: 8 },
    flowStoredData: { name: '순서도: 저장 데이터', cat: '순서도', sites: 4 },
    flowMagneticDisk: { name: '순서도: 자기 디스크', cat: '순서도', sites: 4 },
    scrollH: { name: '가로로 말린 두루마리 모양', cat: '별 및 현수막', sites: 4 },
    scrollV: { name: '세로로 말린 두루마리 모양', cat: '별 및 현수막', sites: 4 },
    star5: { name: '별: 꼭짓점 5개', cat: '별 및 현수막', sites: 5 },
    calloutRect: { name: '사각형 설명선', cat: '설명선', sites: 4, adj: [-0.20833, 0.625] },
    calloutRoundRect: { name: '모서리가 둥근 사각형 설명선', cat: '설명선', sites: 4, adj: [-0.20833, 0.625] },
    calloutOval: { name: '타원형 설명선', cat: '설명선', sites: 8, adj: [-0.20833, 0.625] },
    textbox: { name: '텍스트 상자', cat: '기본 도형', sites: 4 },
    line: { name: '직선', cat: '선', line: true },
    connStraight: { name: '직선 연결선', cat: '선', line: true },
    connElbow: { name: '꺾인 연결선', cat: '선', line: true },
    group: { name: '그룹', cat: '그룹', sites: 0 },
    placeholder: { name: '개체 틀', cat: '개체 틀', sites: 4 },
  };
  const GALLERY = [
    ['선', [['line', '선'], ['arrow', '화살표'], ['arrow2', '양쪽 화살표'], ['connElbow', '꺾인 연결선'], ['elbowArrow', '꺾인 화살표 연결선'], ['elbowArrow2', '꺾인 양쪽 화살표 연결선']]],
    ['사각형', [['rect', '사각형'], ['roundRect', '모서리가 둥근 사각형']]],
    ['기본 도형', [['textbox', '텍스트 상자'], ['ellipse', '타원'], ['triangle', '이등변 삼각형'], ['parallelogram', '평행 사변형'], ['trapezoid', '사다리꼴'], ['diamond', '다이아몬드'],
      ['pentagon', '정오각형'], ['hexagon', '육각형'], ['can', '원통형'], ['cloud', '구름'], ['leftBracket', '왼쪽 대괄호'], ['rightBracket', '오른쪽 대괄호'], ['leftBrace', '왼쪽 중괄호'], ['rightBrace', '오른쪽 중괄호']]],
    ['블록 화살표', [['rightArrow', '오른쪽 화살표'], ['leftArrow', '왼쪽 화살표'], ['upArrow', '위쪽 화살표'], ['downArrow', '아래쪽 화살표']]],
    ['순서도', [['flowProcess', '순서도: 처리'], ['flowAlt', '순서도: 대체 처리'], ['flowDecision', '순서도: 결정'], ['flowData', '순서도: 데이터'], ['flowDocument', '순서도: 문서'],
      ['flowTerminator', '순서도: 종료'], ['flowPreparation', '순서도: 준비'], ['flowManualInput', '순서도: 수동 입력'], ['flowConnector', '순서도: 연결자'], ['flowStoredData', '순서도: 저장 데이터'], ['flowMagneticDisk', '순서도: 자기 디스크']]],
    ['별 및 현수막', [['star5', '별: 꼭짓점 5개'], ['scrollV', '세로로 말린 두루마리 모양'], ['scrollH', '가로로 말린 두루마리 모양']]],
    ['설명선', [['calloutRect', '사각형 설명선'], ['calloutRoundRect', '모서리가 둥근 사각형 설명선'], ['calloutOval', '타원형 설명선']]],
  ];
  const isLine = s => !!(KINDS[s.kind] && KINDS[s.kind].line);

  // ------------------------------------------------------------------ 프레젠테이션
  function defaults() {
    return {
      shape: { fill: THEME.accent1, line: { color: THEME.accent1Dark, weight: 1, dash: 'solid', begin: null, end: null }, shadow: null,
        text: { font: '맑은 고딕', size: 18, bold: false, color: '#FFFFFF', align: 'center', anchor: 'middle' } },
      line: { line: { color: THEME.accent1, weight: 0.75, dash: 'solid', begin: null, end: null }, shadow: null },
      textbox: { fill: null, line: null, shadow: null, text: { font: '맑은 고딕', size: 18, bold: false, color: '#000000', align: 'left', anchor: 'top' } },
    };
  }
  function newPresentation() {
    const p = { size: Object.assign({ kind: 'wide' }, SIZES.wide), slides: [], defaults: defaults(), seq: 1,
      handout: { header: true, date: true, footer: true, pageNum: true, perPage: 6, orient: 'portrait' },
      print: { target: 'all', range: '', layout: 'full', frame: false, fit: true, hq: false, orient: 'portrait', color: 'color', sides: 'single', collate: true } };
    p.slides.push({ id: U.uid('s'), layout: 'title', shapes: [] });
    applyLayout(p, p.slides[0], 'title');
    return p;
  }
  const LAYOUTS = [['title', '제목 슬라이드'], ['titleContent', '제목 및 내용'], ['section', '구역 머리글'], ['two', '콘텐츠 2개'], ['compare', '비교'],
    ['titleOnly', '제목만'], ['blank', '빈 화면'], ['captionContent', '캡션 있는 콘텐츠'], ['captionPic', '캡션 있는 그림']];
  function layoutPlaceholders(p, layout) {
    const W = p.size.w, H = p.size.h, s = W / 33.867;
    const ph = (role, x, y, w, h, prompt, size, anchor) => ({ role, x: x * s, y, w: w * s, h, prompt, size, anchor });
    switch (layout) {
      case 'title': return [ph('title', 4.23, 3.12, 25.4, 6.63, '제목을 입력하십시오', 54, 'bottom'), ph('subtitle', 4.23, 10.0, 25.4, 4.6, '부제목을 입력하십시오', 24, 'top')];
      case 'titleContent': return [ph('title', 2.33, 1.01, 29.21, 3.68, '제목을 입력하십시오', 40, 'middle'), ph('body', 2.33, 5.07, 29.21, 12.09, '텍스트를 입력하십시오', 24, 'top')];
      case 'section': return [ph('title', 2.31, 4.75, 29.21, 7.93, '제목을 입력하십시오', 54, 'bottom'), ph('body', 2.31, 12.76, 29.21, 4.17, '텍스트를 입력하십시오', 24, 'top')];
      case 'two': return [ph('title', 2.33, 1.01, 29.21, 3.68, '제목을 입력하십시오', 40, 'middle'), ph('body', 2.33, 5.07, 14.39, 12.09, '텍스트를 입력하십시오', 24, 'top'), ph('body', 17.14, 5.07, 14.39, 12.09, '텍스트를 입력하십시오', 24, 'top')];
      case 'compare': return [ph('title', 2.33, 1.01, 29.21, 3.68, '제목을 입력하십시오', 40, 'middle'), ph('body', 2.33, 4.67, 14.33, 2.29, '텍스트를 입력하십시오', 24, 'top'), ph('body', 17.15, 4.67, 14.4, 2.29, '텍스트를 입력하십시오', 24, 'top')];
      case 'titleOnly': return [ph('title', 2.33, 1.01, 29.21, 3.68, '제목을 입력하십시오', 40, 'middle')];
      case 'captionContent': case 'captionPic': return [ph('title', 2.33, 1.27, 10.92, 4.44, '제목을 입력하십시오', 32, 'bottom'), ph('body', 14.4, 2.74, 17.15, 13.53, '텍스트를 입력하십시오', 24, 'top')];
    }
    return [];
  }
  function applyLayout(p, slide, layout) {
    slide.shapes = slide.shapes.filter(s => !(s.kind === 'placeholder' && !s.text.value));
    for (const s of slide.shapes) if (s.kind === 'placeholder') s.kind = 'textbox';
    const phs = layoutPlaceholders(p, layout).map(d => ({ id: 'sh' + (p.seq++), kind: 'placeholder', role: d.role, name: (d.role === 'title' ? '제목 ' : d.role === 'subtitle' ? '부제목 ' : '내용 개체 틀 ') + (p.seq - 1),
      x: d.x, y: d.y, w: d.w, h: d.h, fill: null, line: null, shadow: null, prompt: d.prompt,
      text: { value: '', font: '맑은 고딕', size: d.size, bold: false, color: '#000000', align: d.role === 'body' ? 'left' : 'center', anchor: d.anchor, bullet: d.role === 'body' ? '•' : null } }));
    slide.shapes = phs.concat(slide.shapes);
    slide.layout = layout;
  }
  function setLayout(p, si, layout) { applyLayout(p, p.slides[si], layout); }
  function addSlide(p, after, layout) {
    const cur = p.slides[after];
    const lay = layout || (cur ? (cur.layout === 'title' ? 'titleContent' : cur.layout) : 'titleContent');
    const s = { id: U.uid('s'), layout: lay, shapes: [] };
    applyLayout(p, s, lay);
    p.slides.splice(after + 1, 0, s);
    return after + 1;
  }
  function deleteSlide(p, si) { if (p.slides.length > 1) p.slides.splice(si, 1); }
  function eachShape(list, fn, parent) { for (const s of list) { fn(s, parent); if (s.children) eachShape(s.children, fn, s); } }
  function setSlideSize(p, kind, mode) {
    const n = SIZES[kind];
    if (!n || p.size.kind === kind) return;
    const sx = n.w / p.size.w, sy = n.h / p.size.h;
    const s = mode === 'max' ? Math.max(sx, sy) : Math.min(sx, sy);
    const ox = (n.w - p.size.w * s) / 2, oy = (n.h - p.size.h * s) / 2;
    for (const sl of p.slides) eachShape(sl.shapes, sh => {
      if (sh.kind === 'group') return;
      if (isLine(sh)) { sh.x1 = sh.x1 * s + ox; sh.y1 = sh.y1 * s + oy; sh.x2 = sh.x2 * s + ox; sh.y2 = sh.y2 * s + oy; }
      else { sh.x = sh.x * s + ox; sh.y = sh.y * s + oy; sh.w *= s; sh.h *= s; if (sh.text && sh.kind !== 'placeholder') sh.text.size = Math.round(sh.text.size * s * 2) / 2; }
    });
    p.size = Object.assign({ kind }, n);
  }

  // ------------------------------------------------------------------ 도형 만들기
  function nameFor(kind, id) { return (KINDS[kind] ? KINDS[kind].name : kind) + ' ' + id; }
  function addShape(p, si, kind, geom, opts) {
    const o = opts || {};
    const id = 'sh' + (p.seq++);
    let s;
    const D = p.defaults;
    if (kind === 'line' || kind === 'connStraight' || kind === 'connElbow') {
      s = Object.assign({ id, kind, name: nameFor(kind, id), x1: geom.x1, y1: geom.y1, x2: geom.x2, y2: geom.y2, begin: null, end: null, adj: 0.5, fill: null },
        U.clone(D.line));
      if (o.arrow === 'end' || o.arrow === 'both') s.line.end = 'arrow';
      if (o.arrow === 'both') s.line.begin = 'arrow';
    } else if (kind === 'textbox') {
      s = Object.assign({ id, kind, name: nameFor(kind, id), x: geom.x, y: geom.y, w: geom.w, h: geom.h, autoFit: true, margins: [0.25, 0.13, 0.25, 0.13] }, U.clone(D.textbox));
      s.text.value = '';
    } else {
      s = Object.assign({ id, kind, name: nameFor(kind, id), x: geom.x, y: geom.y, w: geom.w, h: geom.h, margins: [0.25, 0.13, 0.25, 0.13] }, U.clone(D.shape));
      s.text.value = '';
      if (KINDS[kind] && KINDS[kind].adj) s.adj = KINDS[kind].adj.slice();
    }
    if (p.slides[si]) p.slides[si].shapes.push(s);
    return s;
  }
  function findShape(p, si, id) {
    let hit = null;
    eachShape(p.slides[si].shapes, s => { if (s.id === id) hit = s; });
    return hit;
  }
  function findByName(slide, name) { let hit = null; eachShape(slide.shapes, s => { if (s.name === name) hit = s; }); return hit; }
  function parentList(slide, id) {
    let res = null;
    const walk = (list) => { for (const s of list) { if (s.id === id) { res = list; return; } if (s.children) walk(s.children); } };
    walk(slide.shapes);
    return res;
  }
  function topLevelOf(slide, id) {
    for (const s of slide.shapes) { if (s.id === id) return s; let found = false; if (s.children) eachShape(s.children, c => { if (c.id === id) found = true; }); if (found) return s; }
    return null;
  }

  // ------------------------------------------------------------------ 기하
  function bbox(s) {
    if (s.kind === 'group') {
      let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
      for (const c of s.children) { const b = bbox(c); x1 = Math.min(x1, b.x); y1 = Math.min(y1, b.y); x2 = Math.max(x2, b.x + b.w); y2 = Math.max(y2, b.y + b.h); }
      return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
    }
    if (isLine(s)) return { x: Math.min(s.x1, s.x2), y: Math.min(s.y1, s.y2), w: Math.abs(s.x2 - s.x1), h: Math.abs(s.y2 - s.y1) };
    return { x: s.x, y: s.y, w: s.w, h: s.h };
  }
  function sitePos(s, site) {
    const k = KINDS[s.kind] || {};
    const { x, y, w, h } = s;
    const n = k.sites || 4;
    if (n === 8) {
      const a = [90, 135, 180, 225, 270, 315, 0, 45][(site - 1) % 8] * Math.PI / 180;
      const dirs = [[0, -1], [-1, 0], [-1, 0], [-1, 0], [0, 1], [1, 0], [1, 0], [1, 0]][(site - 1) % 8];
      return { x: x + w / 2 + Math.cos(a) * w / 2, y: y + h / 2 - Math.sin(a) * h / 2, dx: dirs[0], dy: dirs[1] };
    }
    const m = [[x + w / 2, y, 0, -1], [x, y + h / 2, -1, 0], [x + w / 2, y + h, 0, 1], [x + w, y + h / 2, 1, 0]][((site - 1) % 4 + 4) % 4];
    return { x: m[0], y: m[1], dx: m[2], dy: m[3] };
  }
  function sites(s) {
    const k = KINDS[s.kind] || {};
    if (!k.sites || s.kind === 'group' || isLine(s)) return [];
    const n = k.sites === 8 ? 8 : 4;
    return [...Array(n).keys()].map(i => Object.assign({ site: i + 1 }, sitePos(s, i + 1)));
  }
  /* 연결선 끝점(붙은 도형 기준) */
  function endpoints(p, slide, s) {
    const end = (att, fx, fy) => {
      if (att) { const t = findByIdIn(slide, att.id); if (t && t.kind !== 'group') { const q = sitePos(t, att.site); return { x: q.x, y: q.y, dx: q.dx, dy: q.dy, attached: true }; } }
      return { x: fx, y: fy, dx: 0, dy: 0, attached: false };
    };
    return [end(s.begin, s.x1, s.y1), end(s.end, s.x2, s.y2)];
  }
  function findByIdIn(slide, id) { let hit = null; eachShape(slide.shapes, s => { if (s.id === id) hit = s; }); return hit; }
  /* 꺾인 연결선 경로 점 목록 */
  function elbowPoints(a, b, adj) {
    const t = adj === undefined || adj === null ? 0.5 : adj;
    let d1 = { x: a.dx, y: a.dy }, d2 = { x: b.dx, y: b.dy };
    if (!a.attached && !b.attached) { d1 = { x: Math.sign(b.x - a.x) || 1, y: 0 }; d2 = { x: -d1.x, y: 0 }; }
    else if (!a.attached) d1 = Math.abs(b.x - a.x) >= Math.abs(b.y - a.y) ? { x: Math.sign(b.x - a.x) || 1, y: 0 } : { x: 0, y: Math.sign(b.y - a.y) || 1 };
    else if (!b.attached) d2 = d1.x ? { x: 0, y: -Math.sign(b.y - a.y) || -1 } : { x: -Math.sign(b.x - a.x) || -1, y: 0 };
    const H1 = d1.x !== 0, H2 = d2.x !== 0;
    const M = 0.4;
    if (H1 && H2) {
      if (Math.abs(a.y - b.y) < 1e-6 && d1.x === Math.sign(b.x - a.x)) return [a, b];
      let xm = a.x + (b.x - a.x) * t;
      if (d1.x === d2.x) xm = d1.x > 0 ? Math.max(a.x, b.x) + M : Math.min(a.x, b.x) - M;
      else if (Math.sign(b.x - a.x) !== d1.x) { const x1 = a.x + d1.x * M, x2 = b.x + d2.x * M, ym = (a.y + b.y) / 2; return [a, { x: x1, y: a.y }, { x: x1, y: ym }, { x: x2, y: ym }, { x: x2, y: b.y }, b]; }
      return [a, { x: xm, y: a.y }, { x: xm, y: b.y }, b];
    }
    if (!H1 && !H2) {
      if (Math.abs(a.x - b.x) < 1e-6 && d1.y === Math.sign(b.y - a.y)) return [a, b];
      let ym = a.y + (b.y - a.y) * t;
      if (d1.y === d2.y) ym = d1.y > 0 ? Math.max(a.y, b.y) + M : Math.min(a.y, b.y) - M;
      else if (Math.sign(b.y - a.y) !== d1.y) { const y1 = a.y + d1.y * M, y2 = b.y + d2.y * M, xm = (a.x + b.x) / 2; return [a, { x: a.x, y: y1 }, { x: xm, y: y1 }, { x: xm, y: y2 }, { x: b.x, y: y2 }, b]; }
      return [a, { x: a.x, y: ym }, { x: b.x, y: ym }, b];
    }
    if (H1 && !H2) return [a, { x: b.x, y: a.y }, b];
    return [a, { x: a.x, y: b.y }, b];
  }
  function linePoints(p, slide, s) {
    const [a, b] = endpoints(p, slide, s);
    if (s.kind === 'connElbow') return elbowPoints(a, b, s.adj);
    return [a, b];
  }

  /* 도형 윤곽 경로(SVG d, 단위 cm) */
  function f(n) { return Math.round(n * 1000) / 1000; }
  function geomPath(s) {
    const { x, y, w, h } = s;
    const P = (pts, close) => 'M' + pts.map(q => f(q[0]) + ' ' + f(q[1])).join(' L') + (close === false ? '' : ' Z');
    const ss = Math.min(w, h);
    const k = s.kind;
    switch (k) {
      case 'rect': case 'flowProcess': case 'textbox': case 'placeholder': return P([[x, y], [x + w, y], [x + w, y + h], [x, y + h]]);
      case 'roundRect': case 'flowAlt': { const r = ss * (s.adj && s.adj[0] !== undefined && k === 'roundRect' ? s.adj[0] : 0.16667); return rr(x, y, w, h, r); }
      case 'ellipse': case 'flowConnector': return ell(x, y, w, h);
      case 'diamond': case 'flowDecision': return P([[x + w / 2, y], [x + w, y + h / 2], [x + w / 2, y + h], [x, y + h / 2]]);
      case 'triangle': return P([[x + w / 2, y], [x + w, y + h], [x, y + h]]);
      case 'parallelogram': return P([[x + ss * 0.25, y], [x + w, y], [x + w - ss * 0.25, y + h], [x, y + h]]);
      case 'flowData': return P([[x + w * 0.2, y], [x + w, y], [x + w * 0.8, y + h], [x, y + h]]);
      case 'trapezoid': return P([[x + ss * 0.25, y], [x + w - ss * 0.25, y], [x + w, y + h], [x, y + h]]);
      case 'pentagon': return P([[x + w / 2, y], [x + w, y + h * 0.38], [x + w * 0.81, y + h], [x + w * 0.19, y + h], [x, y + h * 0.38]]);
      case 'hexagon': { const a = ss * 0.25; return P([[x + a, y], [x + w - a, y], [x + w, y + h / 2], [x + w - a, y + h], [x + a, y + h], [x, y + h / 2]]); }
      case 'flowPreparation': return P([[x + w * 0.2, y], [x + w * 0.8, y], [x + w, y + h / 2], [x + w * 0.8, y + h], [x + w * 0.2, y + h], [x, y + h / 2]]);
      case 'flowTerminator': { const rx = Math.min(w * 0.1609, w / 2), ry = h / 2; return `M${f(x + rx)} ${f(y)} L${f(x + w - rx)} ${f(y)} A${f(rx)} ${f(ry)} 0 0 1 ${f(x + w - rx)} ${f(y + h)} L${f(x + rx)} ${f(y + h)} A${f(rx)} ${f(ry)} 0 0 1 ${f(x + rx)} ${f(y)} Z`; }
      case 'flowDocument': return `M${f(x)} ${f(y)} L${f(x + w)} ${f(y)} L${f(x + w)} ${f(y + h * 0.83)} C${f(x + w * 0.75)} ${f(y + h * 0.66)} ${f(x + w * 0.62)} ${f(y + h * 0.8)} ${f(x + w * 0.5)} ${f(y + h * 0.88)} C${f(x + w * 0.35)} ${f(y + h * 0.99)} ${f(x + w * 0.18)} ${f(y + h * 1.03)} ${f(x)} ${f(y + h * 0.9)} Z`;
      case 'flowManualInput': return P([[x, y + h * 0.2], [x + w, y], [x + w, y + h], [x, y + h]]);
      case 'flowStoredData': { const a = w / 6; return `M${f(x + a)} ${f(y)} L${f(x + w)} ${f(y)} A${f(a)} ${f(h / 2)} 0 0 0 ${f(x + w)} ${f(y + h)} L${f(x + a)} ${f(y + h)} A${f(a)} ${f(h / 2)} 0 0 1 ${f(x + a)} ${f(y)} Z`; }
      case 'flowMagneticDisk': case 'can': { const ry = h * 0.12; return `M${f(x)} ${f(y + ry)} A${f(w / 2)} ${f(ry)} 0 0 1 ${f(x + w)} ${f(y + ry)} L${f(x + w)} ${f(y + h - ry)} A${f(w / 2)} ${f(ry)} 0 0 1 ${f(x)} ${f(y + h - ry)} Z M${f(x)} ${f(y + ry)} A${f(w / 2)} ${f(ry)} 0 0 0 ${f(x + w)} ${f(y + ry)}`; }
      case 'cloud': return cloudPath(x, y, w, h);
      case 'rightArrow': { const hl = ss * 0.5; return P([[x, y + h * 0.25], [x + w - hl, y + h * 0.25], [x + w - hl, y], [x + w, y + h / 2], [x + w - hl, y + h], [x + w - hl, y + h * 0.75], [x, y + h * 0.75]]); }
      case 'leftArrow': { const hl = ss * 0.5; return P([[x + w, y + h * 0.25], [x + hl, y + h * 0.25], [x + hl, y], [x, y + h / 2], [x + hl, y + h], [x + hl, y + h * 0.75], [x + w, y + h * 0.75]]); }
      case 'downArrow': { const hl = ss * 0.5; return P([[x + w * 0.25, y], [x + w * 0.75, y], [x + w * 0.75, y + h - hl], [x + w, y + h - hl], [x + w / 2, y + h], [x, y + h - hl], [x + w * 0.25, y + h - hl]]); }
      case 'upArrow': { const hl = ss * 0.5; return P([[x + w * 0.25, y + h], [x + w * 0.75, y + h], [x + w * 0.75, y + hl], [x + w, y + hl], [x + w / 2, y], [x, y + hl], [x + w * 0.25, y + hl]]); }
      case 'leftBrace': return `M${f(x + w)} ${f(y)} Q${f(x + w / 2)} ${f(y)} ${f(x + w / 2)} ${f(y + h * 0.08)} L${f(x + w / 2)} ${f(y + h * 0.42)} Q${f(x + w / 2)} ${f(y + h / 2)} ${f(x)} ${f(y + h / 2)} Q${f(x + w / 2)} ${f(y + h / 2)} ${f(x + w / 2)} ${f(y + h * 0.58)} L${f(x + w / 2)} ${f(y + h * 0.92)} Q${f(x + w / 2)} ${f(y + h)} ${f(x + w)} ${f(y + h)}`;
      case 'rightBrace': return `M${f(x)} ${f(y)} Q${f(x + w / 2)} ${f(y)} ${f(x + w / 2)} ${f(y + h * 0.08)} L${f(x + w / 2)} ${f(y + h * 0.42)} Q${f(x + w / 2)} ${f(y + h / 2)} ${f(x + w)} ${f(y + h / 2)} Q${f(x + w / 2)} ${f(y + h / 2)} ${f(x + w / 2)} ${f(y + h * 0.58)} L${f(x + w / 2)} ${f(y + h * 0.92)} Q${f(x + w / 2)} ${f(y + h)} ${f(x)} ${f(y + h)}`;
      case 'leftBracket': return `M${f(x + w)} ${f(y)} Q${f(x)} ${f(y)} ${f(x)} ${f(y + h * 0.08)} L${f(x)} ${f(y + h * 0.92)} Q${f(x)} ${f(y + h)} ${f(x + w)} ${f(y + h)}`;
      case 'rightBracket': return `M${f(x)} ${f(y)} Q${f(x + w)} ${f(y)} ${f(x + w)} ${f(y + h * 0.08)} L${f(x + w)} ${f(y + h * 0.92)} Q${f(x + w)} ${f(y + h)} ${f(x)} ${f(y + h)}`;
      case 'star5': { const pts = []; for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + i * Math.PI / 5, r = i % 2 ? 0.38 : 1; pts.push([x + w / 2 + Math.cos(a) * w / 2 * r, y + h / 2 + Math.sin(a) * h / 2 * r * 1.05 + h * 0.03]); } return P(pts); }
      case 'scrollH': { const e = ss * 0.125; return `M${f(x + e)} ${f(y + e)} L${f(x + w - e)} ${f(y + e)} L${f(x + w - e)} ${f(y + e / 2)} A${f(e / 2)} ${f(e / 2)} 0 0 1 ${f(x + w)} ${f(y + e / 2)} L${f(x + w)} ${f(y + h - e * 1.5)} A${f(e / 2)} ${f(e / 2)} 0 0 1 ${f(x + w - e / 2)} ${f(y + h - e)} L${f(x + e)} ${f(y + h - e)} L${f(x + e)} ${f(y + h - e / 2)} A${f(e / 2)} ${f(e / 2)} 0 0 1 ${f(x)} ${f(y + h - e / 2)} L${f(x)} ${f(y + e * 1.5)} A${f(e / 2)} ${f(e / 2)} 0 0 1 ${f(x + e / 2)} ${f(y + e)} Z M${f(x + e)} ${f(y + e * 1.5)} L${f(x + e)} ${f(y + h - e)} M${f(x + w - e)} ${f(y + e)} L${f(x + w - e / 2)} ${f(y + e)} M${f(x + e / 2)} ${f(y + e)} A${f(e / 4)} ${f(e / 4)} 0 0 1 ${f(x + e / 2)} ${f(y + e * 1.5)} L${f(x + e)} ${f(y + e * 1.5)}`; }
      case 'scrollV': { const e = ss * 0.125; return `M${f(x + e)} ${f(y + e)} L${f(x + w - e)} ${f(y + e)} L${f(x + w - e)} ${f(y + h)} L${f(x + e / 2)} ${f(y + h)} A${f(e / 2)} ${f(e / 2)} 0 0 1 ${f(x + e / 2)} ${f(y + h - e)} L${f(x + e)} ${f(y + h - e)} Z M${f(x + e)} ${f(y + e)} A${f(e / 2)} ${f(e / 2)} 0 0 1 ${f(x + w - e)} ${f(y)} L${f(x + w)} ${f(y)} A${f(e / 2)} ${f(e / 2)} 0 0 1 ${f(x + w - e / 2)} ${f(y + e)}`; }
      case 'calloutRect': case 'calloutRoundRect': return calloutBox(s, k === 'calloutRoundRect' ? ss * 0.16667 : 0);
      case 'calloutOval': return calloutOval(s);
    }
    return P([[x, y], [x + w, y], [x + w, y + h], [x, y + h]]);
  }
  function rr(x, y, w, h, r) {
    r = Math.min(r, w / 2, h / 2);
    return `M${f(x + r)} ${f(y)} L${f(x + w - r)} ${f(y)} A${f(r)} ${f(r)} 0 0 1 ${f(x + w)} ${f(y + r)} L${f(x + w)} ${f(y + h - r)} A${f(r)} ${f(r)} 0 0 1 ${f(x + w - r)} ${f(y + h)} L${f(x + r)} ${f(y + h)} A${f(r)} ${f(r)} 0 0 1 ${f(x)} ${f(y + h - r)} L${f(x)} ${f(y + r)} A${f(r)} ${f(r)} 0 0 1 ${f(x + r)} ${f(y)} Z`;
  }
  function ell(x, y, w, h) { return `M${f(x)} ${f(y + h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(x + w)} ${f(y + h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(x)} ${f(y + h / 2)} Z`; }
  function cloudPath(x, y, w, h) {
    const n = 11, cx = x + w / 2, cy = y + h / 2;
    const pts = [];
    for (let i = 0; i < n; i++) { const a = -Math.PI / 2 + i * 2 * Math.PI / n; pts.push([cx + Math.cos(a) * w * 0.43, cy + Math.sin(a) * h * 0.4]); }
    let d = `M${f(pts[0][0])} ${f(pts[0][1])}`;
    for (let i = 0; i < n; i++) {
      const q = pts[(i + 1) % n], pp = pts[i];
      const r = Math.hypot(q[0] - pp[0], q[1] - pp[1]) * 0.62;
      d += ` A${f(r)} ${f(r * 0.9)} 0 0 1 ${f(q[0])} ${f(q[1])}`;
    }
    return d + ' Z';
  }
  function tailTip(s) {
    const a = s.adj || KINDS[s.kind].adj;
    return { x: s.x + s.w / 2 + a[0] * s.w, y: s.y + s.h / 2 + a[1] * s.h };
  }
  function calloutBox(s, r) {
    const { x, y, w, h } = s;
    const t = tailTip(s);
    const rx = (t.x - (x + w / 2)) / w, ry = (t.y - (y + h / 2)) / h;
    if (Math.abs(rx) <= 0.5 && Math.abs(ry) <= 0.5) return r ? rr(x, y, w, h, r) : `M${f(x)} ${f(y)} L${f(x + w)} ${f(y)} L${f(x + w)} ${f(y + h)} L${f(x)} ${f(y + h)} Z`;
    r = Math.min(r, w / 2, h / 2);
    const vert = Math.abs(ry) >= Math.abs(rx);
    const segs = [];
    // 시계 방향 윤곽, 꼬리가 붙는 변에 노치
    const A = (cx2, cy2) => r ? ` A${f(r)} ${f(r)} 0 0 1 ${f(cx2)} ${f(cy2)}` : '';
    let d = `M${f(x + r)} ${f(y)}`;
    if (vert && ry < 0) { const b1 = x + w * (rx < 0 ? 0.25 : 0.583), b2 = b1 + w * 0.167; d += ` L${f(b1)} ${f(y)} L${f(t.x)} ${f(t.y)} L${f(b2)} ${f(y)}`; }
    d += ` L${f(x + w - r)} ${f(y)}` + A(x + w, y + r);
    if (!vert && rx > 0) { const b1 = y + h * (ry < 0 ? 0.25 : 0.583), b2 = b1 + h * 0.167; d += ` L${f(x + w)} ${f(b1)} L${f(t.x)} ${f(t.y)} L${f(x + w)} ${f(b2)}`; }
    d += ` L${f(x + w)} ${f(y + h - r)}` + A(x + w - r, y + h);
    if (vert && ry > 0) { const b2 = x + w * (rx < 0 ? 0.25 : 0.583), b1 = b2 + w * 0.167; d += ` L${f(b1)} ${f(y + h)} L${f(t.x)} ${f(t.y)} L${f(b2)} ${f(y + h)}`; }
    d += ` L${f(x + r)} ${f(y + h)}` + A(x, y + h - r);
    if (!vert && rx < 0) { const b2 = y + h * (ry < 0 ? 0.25 : 0.583), b1 = b2 + h * 0.167; d += ` L${f(x)} ${f(b1)} L${f(t.x)} ${f(t.y)} L${f(x)} ${f(b2)}`; }
    d += ` L${f(x)} ${f(y + r)}` + A(x + r, y) + ' Z';
    void segs;
    return d;
  }
  function calloutOval(s) {
    const { x, y, w, h } = s;
    const cx = x + w / 2, cy = y + h / 2;
    const t = tailTip(s);
    const ang = Math.atan2((t.y - cy) / h, (t.x - cx) / w);
    const a1 = ang - 0.22, a2 = ang + 0.22;
    const p = a => [cx + Math.cos(a) * w / 2, cy + Math.sin(a) * h / 2];
    const q1 = p(a1), q2 = p(a2);
    return `M${f(q2[0])} ${f(q2[1])} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(q1[0])} ${f(q1[1])} L${f(t.x)} ${f(t.y)} Z`;
  }

  // ------------------------------------------------------------------ 편집 명령
  function moveShapes(p, slide, list, dx, dy) {
    for (const s of list) eachShape([s], c => {
      if (c.kind === 'group') return;
      if (isLine(c)) { if (!c.begin) { c.x1 += dx; c.y1 += dy; } if (!c.end) { c.x2 += dx; c.y2 += dy; } if (c.begin && c.end && list.length === 1 && s === c) { /* 양 끝 고정 */ } }
      else { c.x += dx; c.y += dy; }
    });
  }
  function cloneShape(p, s) {
    const c = U.clone(s);
    eachShape([c], x => { x.id = 'sh' + (p.seq++); x.name = nameFor(x.kind, x.id.slice(2)); x.begin = x.begin && null; x.end = x.end && null; });
    return c;
  }
  function duplicate(p, si, list, dx, dy) {
    const out = [];
    for (const s of list) {
      const c = U.clone(s);
      const idMap = {};
      eachShape([c], x => { const nid = 'sh' + (p.seq++); idMap[x.id] = nid; x.id = nid; x.name = nameFor(x.kind, nid.slice(2)); });
      eachShape([c], x => { if (x.begin) x.begin = idMap[x.begin.id] ? { id: idMap[x.begin.id], site: x.begin.site } : null; if (x.end) x.end = idMap[x.end.id] ? { id: idMap[x.end.id], site: x.end.site } : null; });
      if (isLine(c) && !c.begin && !c.end) { const [a, b] = endpoints(p, p.slides[si], s); c.x1 = a.x; c.y1 = a.y; c.x2 = b.x; c.y2 = b.y; }
      moveShapes(p, p.slides[si], [c], dx, dy);
      p.slides[si].shapes.push(c);
      out.push(c);
    }
    return out;
  }
  function removeShapes(p, si, ids) {
    const slide = p.slides[si];
    const drop = list => list.filter(s => !ids.includes(s.id)).map(s => { if (s.children) s.children = drop(s.children); return s; });
    slide.shapes = drop(slide.shapes);
    eachShape(slide.shapes, s => {
      if (isLine(s)) {
        if (s.begin && ids.includes(s.begin.id)) { const q = endpoints(p, slide, s)[0]; s.begin = null; s.x1 = q.x; s.y1 = q.y; }
        if (s.end && ids.includes(s.end.id)) { s.end = null; }
      }
    });
  }
  function freezeEnds(p, slide, s) { const [a, b] = endpoints(p, slide, s); s.x1 = a.x; s.y1 = a.y; s.x2 = b.x; s.y2 = b.y; }
  function group(p, si, ids) {
    const slide = p.slides[si];
    const members = slide.shapes.filter(s => ids.includes(s.id));
    if (members.length < 2) return null;
    const idx = Math.max(...members.map(s => slide.shapes.indexOf(s)));
    eachShape(members, s => { if (isLine(s)) freezeEnds(p, slide, s); });
    const g = { id: 'sh' + (p.seq++), kind: 'group', children: members, fill: null, line: null };
    g.name = '그룹 ' + g.id.slice(2);
    const rest = slide.shapes.filter(s => !ids.includes(s.id));
    const before = slide.shapes.slice(0, idx + 1).filter(s => !ids.includes(s.id)).length;
    rest.splice(before, 0, g);
    slide.shapes = rest;
    // 그룹 안 도형에 붙어 있던 바깥 연결선은 떼어 낸다(그룹에는 연결점이 없음)
    const inner = new Set(); eachShape(members, s => inner.add(s.id));
    eachShape(slide.shapes, s => { if (isLine(s) && !inner.has(s.id)) { const [a, b] = endpoints(p, slide, s); if (s.begin && inner.has(s.begin.id)) { s.begin = null; s.x1 = a.x; s.y1 = a.y; } if (s.end && inner.has(s.end.id)) { s.end = null; s.x2 = b.x; s.y2 = b.y; } } });
    return g;
  }
  function ungroup(p, si, id) {
    const slide = p.slides[si];
    const i = slide.shapes.findIndex(s => s.id === id && s.kind === 'group');
    if (i < 0) return [];
    const g = slide.shapes[i];
    slide.shapes.splice(i, 1, ...g.children);
    return g.children;
  }
  function zorder(p, si, ids, cmd) {
    const slide = p.slides[si];
    const sel = slide.shapes.filter(s => ids.includes(s.id));
    const rest = slide.shapes.filter(s => !ids.includes(s.id));
    if (cmd === 'front') slide.shapes = rest.concat(sel);
    else if (cmd === 'back') slide.shapes = sel.concat(rest);
    else {
      const arr = slide.shapes.slice();
      const order = cmd === 'forward' ? [...sel].reverse() : sel;
      for (const s of order) {
        const i = arr.indexOf(s), j = cmd === 'forward' ? i + 1 : i - 1;
        if (j < 0 || j >= arr.length || ids.includes(arr[j].id)) continue;
        arr[i] = arr[j]; arr[j] = s;
      }
      slide.shapes = arr;
    }
  }
  function align(p, si, list, how, toSlide) {
    const bbs = list.map(bbox);
    const ref = toSlide || list.length === 1 ? { x: 0, y: 0, w: p.size.w, h: p.size.h } : (() => { const x1 = Math.min(...bbs.map(b => b.x)), y1 = Math.min(...bbs.map(b => b.y)), x2 = Math.max(...bbs.map(b => b.x + b.w)), y2 = Math.max(...bbs.map(b => b.y + b.h)); return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 }; })();
    list.forEach((s, i) => {
      const b = bbs[i];
      let dx = 0, dy = 0;
      if (how === 'left') dx = ref.x - b.x;
      if (how === 'center') dx = ref.x + ref.w / 2 - (b.x + b.w / 2);
      if (how === 'right') dx = ref.x + ref.w - (b.x + b.w);
      if (how === 'top') dy = ref.y - b.y;
      if (how === 'middle') dy = ref.y + ref.h / 2 - (b.y + b.h / 2);
      if (how === 'bottom') dy = ref.y + ref.h - (b.y + b.h);
      moveShapes(p, p.slides[si], [s], dx, dy);
    });
  }
  /* 가로(세로) 간격을 동일하게: 같은 열(가로 범위가 겹치는 도형)은 한 묶음으로 */
  function distribute(p, si, list, dir) {
    if (list.length < 3 && !(list.length === 2)) return;
    const H = dir === 'h';
    const items = list.map(s => ({ s, b: bbox(s) }));
    const lo = it => H ? it.b.x : it.b.y, size = it => H ? it.b.w : it.b.h;
    items.sort((a, b) => lo(a) - lo(b));
    const clusters = [];
    for (const it of items) {
      const c = clusters.find(cl => Math.abs(lo(cl[0]) - lo(it)) < 0.05 || (lo(it) < lo(cl[0]) + size(cl[0]) * 0.5 && lo(it) + size(it) * 0.5 > lo(cl[0])));
      if (c) c.push(it); else clusters.push([it]);
    }
    if (clusters.length < 3) {
      if (clusters.length === 2 && list.length === 2 && false) return;
      if (clusters.length < 3) return;
    }
    const cl = clusters.map(c => ({ items: c, lo: Math.min(...c.map(lo)), size: Math.max(...c.map(it => lo(it) + size(it))) - Math.min(...c.map(lo)) }));
    const first = cl[0], last = cl[cl.length - 1];
    const total = last.lo + last.size - first.lo;
    const gap = (total - cl.reduce((s, c) => s + c.size, 0)) / (cl.length - 1);
    let pos = first.lo;
    for (const c of cl) {
      const d = pos - c.lo;
      for (const it of c.items) moveShapes(p, p.slides[si], [it.s], H ? d : 0, H ? 0 : d);
      pos += c.size + gap;
    }
  }
  function styleOf(s) { return U.clone({ fill: s.fill, line: s.line, shadow: s.shadow, text: s.text ? Object.assign({}, s.text, { value: undefined }) : undefined }); }
  function setDefault(p, s) {
    const st = styleOf(s);
    if (isLine(s)) p.defaults.line = { line: st.line, shadow: st.shadow };
    else if (s.kind === 'textbox') { p.defaults.textbox = st; delete p.defaults.textbox.text.value; }
    else { p.defaults.shape = st; delete p.defaults.shape.text.value; }
  }

  // ------------------------------------------------------------------ 원안 불러오기(모범 풀이)
  function loadShapes(p, si, ref) {
    const slide = p.slides[si];
    const byName = {};
    const conv = r => {
      const id = 'sh' + (p.seq++);
      if (r.kind === 'group') { const g = { id, kind: 'group', name: r.name, children: r.children.map(conv), fill: null, line: null }; byName[r.name] = g; return g; }
      const s = { id, kind: r.kind, name: r.name };
      if (KINDS[r.kind] && KINDS[r.kind].line) Object.assign(s, { x1: r.x1, y1: r.y1, x2: r.x2, y2: r.y2, begin: r.begin ? { name: r.begin.shape, site: r.begin.site } : null, end: r.end ? { name: r.end.shape, site: r.end.site } : null, adj: 0.5, fill: null });
      else Object.assign(s, { x: r.x, y: r.y, w: r.w, h: r.h, margins: r.kind === 'textbox' ? [0.05, 0.05, 0.05, 0.05] : [0.1, 0.05, 0.1, 0.05] });
      s.fill = r.kind === 'line' || r.kind.startsWith('conn') ? null : r.fill;
      s.line = r.line ? { color: r.line.color, weight: r.line.weight, dash: r.line.dash, begin: r.line.begin, end: r.line.end } : null;
      s.shadow = r.shadow ? { color: '#808080', transparency: 0, blur: 0, angle: 45, dist: 5.7 } : null;
      if (r.text !== undefined) {
        let v = r.text, bullet = null;
        if (v.startsWith('❖ ')) { bullet = '❖'; v = v.slice(2); }
        s.text = { value: v.replace(/\r/g, '\n'), font: r.font || '맑은 고딕', size: r.size, bold: !!r.bold, color: '#000000', align: r.align, anchor: r.anchor, bullet };
      } else if (!(KINDS[r.kind] && KINDS[r.kind].line)) s.text = { value: '', font: '맑은 고딕', size: 18, bold: false, color: '#000000', align: 'center', anchor: 'middle' };
      if (r.adj) s.adj = r.adj.map(v => v);
      if (r.kind === 'textbox') s.autoFit = false;
      byName[r.name] = s;
      return s;
    };
    const shapes = ref.map(conv);
    eachShape(shapes, s => {
      if (s.begin && s.begin.name) { const t = byName[s.begin.name]; s.begin = t ? { id: t.id, site: s.begin.site } : null; }
      if (s.end && s.end.name) { const t = byName[s.end.name]; s.end = t ? { id: t.id, site: s.end.site } : null; }
    });
    slide.shapes = slide.shapes.filter(s => !(s.kind === 'placeholder' && !s.text.value)).concat(shapes);
  }

  // ------------------------------------------------------------------ 인쇄 배치
  const HANDOUT = {   // A4 세로, 4:3 슬라이드, 용지 맞춤 해제 기준(14.4절 실측: 12.7 × 9.52cm, 위 4.56cm·15.6cm)
    h1: { slots: [[1.6, 4.86, 17.8, 20.0]] },
    h2: { slots: [[4.15, 4.56, 12.7, 9.52], [4.15, 15.6, 12.7, 9.52]] },
    h3: { slots: [[1.57, 3.14, 7.8, 5.85], [1.57, 11.92, 7.8, 5.85], [1.57, 20.7, 7.8, 5.85]], lines: true },
    h4: { slots: [[1.66, 4.0, 8.4, 6.3], [10.94, 4.0, 8.4, 6.3], [1.66, 15.4, 8.4, 6.3], [10.94, 15.4, 8.4, 6.3]] },
    h6: { slots: [[1.66, 2.8, 8.4, 6.3], [10.94, 2.8, 8.4, 6.3], [1.66, 10.7, 8.4, 6.3], [10.94, 10.7, 8.4, 6.3], [1.66, 18.6, 8.4, 6.3], [10.94, 18.6, 8.4, 6.3]] },
    h9: { slots: [0, 1, 2].flatMap(r => [0, 1, 2].map(c => [1.4 + c * 6.2, 2.6 + r * 8.3, 5.9, 4.43])) },
  };
  const LAYOUT_NAMES = { full: '전체 페이지 슬라이드', notes: '슬라이드 노트', outline: '개요', h1: '1슬라이드', h2: '2슬라이드', h3: '3슬라이드', h4: '4슬라이드 가로', h6: '6슬라이드 가로', h9: '9슬라이드 가로' };
  function printPages(p) {
    const pr = p.print;
    let idx = p.slides.map((s, i) => i);
    if (pr.target === 'current' && p.current !== undefined) idx = [p.current];
    if (pr.target === 'range' && pr.range) { const set = new Set(); String(pr.range).split(',').forEach(part => { const m = /^(\d+)(?:-(\d+))?$/.exec(part.trim()); if (m) for (let k = +m[1]; k <= +(m[2] || m[1]); k++) set.add(k - 1); }); idx = idx.filter(i => set.has(i)); }
    const portrait = pr.orient === 'portrait';
    const paper = portrait ? { w: 21, h: 29.7 } : { w: 29.7, h: 21 };
    const pages = [];
    if (pr.layout === 'full' || pr.layout === 'notes' || pr.layout === 'outline') {
      for (const i of idx) {
        const m = 0.6;
        const availW = paper.w - 2 * m, availH = paper.h - 2 * m;
        const s = Math.min(availW / p.size.w, availH / p.size.h) * (pr.fit ? 1 : 0.94);
        const w = p.size.w * s, h = p.size.h * s;
        pages.push({ paper, frames: [{ slide: i, x: (paper.w - w) / 2, y: (paper.h - h) / 2, w, h }] });
      }
      return pages;
    }
    const lay = HANDOUT[pr.layout] || HANDOUT.h2;
    let slots = lay.slots;
    if (!portrait) slots = slots.map(([x, y, w, h]) => [y - 1.2, x + 0.2, w, h]);
    const per = slots.length;
    for (let k = 0; k < idx.length; k += per) {
      const frames = [];
      idx.slice(k, k + per).forEach((si, j) => {
        let [x, y, w, h] = slots[j];
        if (pr.fit) { const c = 1.06; x = paper.w / 2 - (paper.w / 2 - x) * c; y = paper.h / 2 - (paper.h / 2 - y) * c; w *= c; h *= c; }
        const s = Math.min(w / p.size.w, h / p.size.h);
        const fw = p.size.w * s, fh = p.size.h * s;
        frames.push({ slide: si, x: x + (w - fw) / 2, y: y + (h - fh) / 2, w: fw, h: fh });
      });
      pages.push({ paper, frames, lines: lay.lines });
    }
    return pages;
  }

  /* 채점용 도형 목록(그룹 펼침) */
  const CATEGORY = { flowProcess: 'rect', flowAlt: 'roundRect', flowDecision: 'diamond', flowConnector: 'ellipse', connStraight: 'line' };
  function inventory(p, si) {
    const slide = p.slides[si];
    const flat = [];
    eachShape(slide.shapes, s => { if (s.kind !== 'group') flat.push(s); });
    return { top: slide.shapes.length, flat, count: k => flat.filter(s => (CATEGORY[s.kind] || s.kind) === k).length };
  }
  const isGray = c => { if (!c) return false; const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(c); if (!m) return false; const r = parseInt(m[1], 16), g = parseInt(m[2], 16), b = parseInt(m[3], 16); return Math.abs(r - g) < 12 && Math.abs(g - b) < 12 && r >= 120 && r <= 175; };

  OA.ppt = { SIZES, PT, KINDS, GALLERY, LAYOUTS, LAYOUT_NAMES, HANDOUT, defaults, newPresentation, setLayout, addSlide, deleteSlide, setSlideSize, eachShape, addShape, findShape,
    findByName, parentList, topLevelOf, bbox, sitePos, sites, endpoints, elbowPoints, linePoints, geomPath, tailTip, moveShapes, duplicate, removeShapes, group, ungroup, zorder,
    align, distribute, setDefault, styleOf, loadShapes, printPages, inventory, isGray, isLine, nameFor, freezeEnds, cloneShape, CATEGORY };
})(typeof window !== 'undefined' ? window : globalThis);

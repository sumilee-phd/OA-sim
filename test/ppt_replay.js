/* PowerPoint 화면 재현 시험(확인용). ui_driver.js 를 먼저 실행한 뒤 이 파일을 실행하고
   await __T.ppt.slide(0[, 'mock1']) · await __T.ppt.slide(1) 로 14장(또는 모의 세트) 원안 슬라이드를 리본·마우스 조작만으로 다시 그린다.
   모델 함수를 직접 부르지 않는다(원안 좌표 계산에만 OA.ppt 를 쓴다). */
(function () {
  const T = window.__T;
  const P = OA.ppt;
  const X = T.ppt = {};
  const app = () => T.app();
  const wrap = () => app().querySelector('.pp-slidewrap');
  const stage = () => app().querySelector('.pp-stage');
  const LABEL = {};
  for (const [, items] of P.GALLERY) for (const [k, l] of items) LABEL[k] = l;
  const W = 25.4;   // 표준(4:3) 너비 cm
  function scr(x, y) { const r = wrap().getBoundingClientRect(), s = r.width / W; return { cx: r.left + x * s, cy: r.top + y * s }; }
  function fireDrag(x0, y0, x1, y1, mods) {
    const a = scr(x0, y0), b = scr(x1, y1), m = mods || {};
    const el = document.elementFromPoint(a.cx, a.cy) || stage();
    el.dispatchEvent(new MouseEvent('mousedown', Object.assign({ bubbles: true, cancelable: true, button: 0, clientX: a.cx, clientY: a.cy }, m)));
    window.dispatchEvent(new MouseEvent('mousemove', Object.assign({ bubbles: true, clientX: (a.cx + b.cx) / 2, clientY: (a.cy + b.cy) / 2 }, m)));
    window.dispatchEvent(new MouseEvent('mousemove', Object.assign({ bubbles: true, clientX: b.cx, clientY: b.cy }, m)));
    window.dispatchEvent(new MouseEvent('mouseup', Object.assign({ bubbles: true, clientX: b.cx, clientY: b.cy }, m)));
  }
  X.click = (x, y, mods) => { const a = scr(x, y); const el = document.elementFromPoint(a.cx, a.cy); const o = Object.assign({ bubbles: true, cancelable: true, button: 0, clientX: a.cx, clientY: a.cy }, mods || {}); el.dispatchEvent(new MouseEvent('mousedown', o)); window.dispatchEvent(new MouseEvent('mouseup', o)); return el; };
  const w = T.wait;
  const menuItem = (text) => {
    const items = [...document.querySelectorAll('.menu .mi, .menu .swatch')].filter(e => e.getClientRects().length);
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const txt = e => norm([...e.childNodes].filter(n => n.nodeName.toLowerCase() !== 'svg' && !(n.classList && n.classList.contains('ck'))).map(n => n.textContent).join(''));
    const el = items.find(e => e.title === text || txt(e) === text) || items.find(e => txt(e).startsWith(text));
    if (!el) throw new Error('메뉴 항목 없음: ' + text);
    return el;
  };
  async function menuPath(rbTitle, ...path) {
    T.rb(rbTitle); await w(40);
    for (const p of path) { const el = menuItem(p); T.clickEl(el); await w(40); }
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); await w(10);
  }
  async function tool(k) {
    T.tab('홈'); await w(20);
    if (k === 'textbox') { T.tab('삽입'); await w(20); T.rb('텍스트 상자'); await w(30); return; }
    T.rb('도형'); await w(40);
    const cell = [...document.querySelectorAll('.menu .shapecell')].find(c => c.title === LABEL[k]);
    if (!cell) throw new Error('도형 목록에 없음: ' + k);
    T.clickEl(cell); await w(30);
  }
  async function typeInto(text) {
    let ta = app().querySelector('textarea.pp-text');
    if (!ta) { T.key('가', {}, stage()); await w(30); ta = app().querySelector('textarea.pp-text'); }
    if (!ta) throw new Error('글자 편집 상자가 열리지 않음');
    ta.value = text; ta.dispatchEvent(new Event('blur')); await w(40);
  }
  async function fill(c) {
    T.tab('홈'); await w(20);
    if (c === null || c === undefined) return menuPath('도형 채우기', '채우기 없음(N)');
    if (c === '#FFFFFF') return menuPath('도형 채우기', '흰색, 배경 1');
    T.rb('도형 채우기'); await w(40); T.clickEl(menuItem('다른 채우기 색(M)...')); await w(80);
    const d = T.dlg(); const v = parseInt(c.slice(1), 16);
    const ins = [...d.querySelectorAll('input[type=number]')];
    [v >> 16, (v >> 8) & 255, v & 255].forEach((n, i) => { T.set(ins[i], String(n)); });
    T.click('확인', d); await w(60);
  }
  const WN = { 0.25: '1/4pt', 0.5: '1/2pt', 0.75: '3/4pt', 1: '1pt', 1.5: '1 1/2pt', 2.25: '2 1/4pt', 3: '3pt' };
  const DN = { solid: '실선', dash: '파선', roundDot: '둥근 점선', sysDot: '사각 점선', lgDash: '긴 파선' };
  async function line(l) {
    T.tab('홈'); await w(20);
    if (!l) return menuPath('도형 윤곽선', '윤곽선 없음(N)');
    await menuPath('도형 윤곽선', '검정, 텍스트 1');
    await menuPath('도형 윤곽선', '두께(W)', WN[l.weight]);
    await menuPath('도형 윤곽선', '대시(S)', DN[l.dash || 'solid']);
  }
  async function shadow() { T.tab('홈'); await w(20); await menuPath('도형 효과', '그림자(S)', '오프셋: 오른쪽 아래'); }
  async function font(s) {
    T.tab('홈'); await w(20);
    const combo = [...app().querySelectorAll('.ribbon .rcombo input')].filter(e => e.getClientRects().length)[1];
    T.focus(combo); combo.value = String(s.size); combo.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true, cancelable: true })); await w(40);
    const bold = [...app().querySelectorAll('.ribbon button')].find(b => b.title === '굵게');
    if (bold.classList.contains('on') !== !!s.bold) { T.clickEl(bold); await w(40); }
    T.rb({ center: '가운데 맞춤', left: '왼쪽 맞춤', right: '오른쪽 맞춤' }[s.align || 'left']); await w(40);
    await menuPath('텍스트 맞춤', { top: '위쪽(T)', middle: '중간(M)', bottom: '아래쪽(B)' }[s.anchor || 'top']);
    await menuPath('글꼴 색', '검정, 텍스트 1');
  }
  // 원안 모델(연결선 끝점 계산용)
  function refPres(setId) {
    const set = OA_DATA.sets.find(s => s.id === (setId || 'equip'));
    const ref = P.newPresentation(); P.setSlideSize(ref, 'standard', 'fit'); P.setLayout(ref, 0, 'blank'); P.addSlide(ref, 0);
    set.pt.ref.forEach((s, i) => P.loadShapes(ref, i, s));
    return { set, ref };
  }
  function lineTool(s) {
    const b = !!(s.line && s.line.begin), e = !!(s.line && s.line.end);
    if (s.kind === 'connElbow') return b && e ? 'elbowArrow2' : e ? 'elbowArrow' : 'connElbow';
    return b && e ? 'arrow2' : e ? 'arrow' : 'line';
  }
  async function drawOne(si, raw, refShape, ref) {
    if (raw.kind === 'group') {
      const kids = refShape.children;
      for (let k = 0; k < raw.children.length; k++) await drawOne(si, raw.children[k], kids[k], ref);
      const b = P.bbox(refShape);
      fireDrag(b.x - 0.15, b.y - 0.05, b.x + b.w + 0.15, b.y + b.h + 0.05); await w(60);
      T.key('g', { ctrlKey: true }, stage()); await w(60);
      return;
    }
    if (P.isLine(refShape)) {
      const [a, b] = P.endpoints(ref, ref.slides[si], refShape);
      await tool(lineTool(raw));
      fireDrag(a.x, a.y, b.x, b.y); await w(60);
      await line(raw.line);
      return;
    }
    await tool(raw.kind);
    fireDrag(raw.x, raw.y, raw.x + raw.w, raw.y + raw.h); await w(60);
    if (raw.kind === 'textbox') {
      await typeInto(raw.text || '');
    } else {
      await fill(raw.fill);
      await line(raw.line);
      if (raw.text) await typeInto(raw.text);
    }
    if (raw.text) await font(raw);
    if (raw.shadow) await shadow();
  }
  X.slide = async (si, setId) => {
    const { set, ref } = refPres(setId);
    const thumbs = [...app().querySelectorAll('.pp-thumb')];
    T.clickEl(thumbs[si]); await w(80);
    const raws = set.pt.ref[si], refs = ref.slides[si].shapes;
    const log = [];
    for (let i = 0; i < raws.length; i++) {
      try { await drawOne(si, raws[i], refs[i], ref); log.push(raws[i].name); }
      catch (e) { log.push(raws[i].name + ' 오류: ' + e.message); throw new Error(log.join(' / ')); }
    }
    return log;
  };
  return 'ppt ready';
})();

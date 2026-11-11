/* PowerPoint 화면: 시작 화면, 리본, 축소판, SVG 슬라이드 편집, 선택 창, 도형 서식 작업창, 유인물 마스터, 회색조 보기, 백스테이지 인쇄 */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const { h, icon, menu, ribbon, dialog, msg, toast, tabs } = OA.ui;
  const U = OA.util, P = OA.ppt;
  const UU = 100;                                     // SVG 단위: 1cm = 100
  const PTCM = 0.035278;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  const DASH = { solid: '', dash: '4 3', sysDot: '1 1', sysDash: '3 1', lgDash: '8 3', dashDot: '4 3 1 3', lgDashDot: '8 3 1 3', lgDashDotDot: '8 3 1 3 1 3', roundDot: '0.1 2' };
  const DASH_NAMES = [['solid', '실선'], ['roundDot', '둥근 점선'], ['sysDot', '사각 점선'], ['dash', '파선'], ['lgDash', '긴 파선'], ['dashDot', '파선-점선'], ['lgDashDot', '긴 파선-점선'], ['lgDashDotDot', '긴 파선-점선-점선']];
  const WEIGHTS = [[0.25, '1/4pt'], [0.5, '1/2pt'], [0.75, '3/4pt'], [1, '1pt'], [1.5, '1 1/2pt'], [2.25, '2 1/4pt'], [3, '3pt'], [4.5, '4 1/2pt'], [6, '6pt']];
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
  const BULLETS = [[null, '없음'], ['•', '채우기 원형 글머리 기호(작은)'], ['◆', '채우기 다이아몬드 글머리 기호'], ['■', '채우기 사각형 글머리 기호'], ['●', '채우기 원형 글머리 기호'], ['❖', '별표 글머리 기호'], ['➢', '화살표 글머리 기호'], ['✓', '체크 표시 글머리 기호']];

  // ------------------------------------------------------------------ 도형 그리기(SVG 문자열)
  function textHtml(s, scaleFont) {
    const t = s.text;
    if (!t) return '';
    const isPh = s.kind === 'placeholder' && !t.value;
    const value = isPh ? s.prompt : t.value;
    if (!value) return '';
    const m = s.margins || [0.25, 0.13, 0.25, 0.13];
    const fs = (t.size || 18) * PTCM * UU;
    const lines = String(value).split('\n');
    const bullet = t.bullet;
    const just = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[t.anchor || 'middle'];
    const body = lines.map(l => bullet && l !== '' ? `<div style="padding-left:1.1em;text-indent:-1.1em">${esc(bullet)}&#8201;${esc(l)}</div>` : `<div>${esc(l) || '&nbsp;'}</div>`).join('');
    const b = P.bbox(s);
    return `<foreignObject x="${b.x * UU}" y="${b.y * UU}" width="${Math.max(1, b.w * UU)}" height="${Math.max(1, b.h * UU)}" style="overflow:visible;pointer-events:none">` +
      `<div xmlns="http://www.w3.org/1999/xhtml" style="box-sizing:border-box;width:100%;height:100%;display:flex;flex-direction:column;justify-content:${just};padding:${m[1] * UU}px ${m[2] * UU}px ${m[3] * UU}px ${m[0] * UU}px;font-family:'${t.font || '맑은 고딕'}','Malgun Gothic',sans-serif;font-size:${fs}px;line-height:1.2;font-weight:${t.bold ? 700 : 400};font-style:${t.italic ? 'italic' : 'normal'};text-decoration:${t.underline ? 'underline' : 'none'};color:${isPh ? '#888' : (t.color || '#000')};text-align:${t.align || 'left'};word-break:keep-all;overflow-wrap:anywhere">${body}</div></foreignObject>`;
  }
  function strokeAttrs(line) {
    if (!line) return 'stroke="none"';
    const w = Math.max(0.2, (line.weight || 0.75) * PTCM * UU);
    const dash = DASH[line.dash || 'solid'];
    return `stroke="${line.color || '#000'}" stroke-width="${w}" ${dash ? `stroke-dasharray="${dash.split(' ').map(x => +x * w).join(' ')}"` : ''} stroke-linejoin="round"`;
  }
  function shadowOf(s, d) {
    const sh = s.shadow;
    if (!sh) return '';
    const dist = (sh.dist === undefined ? 4.2 : sh.dist) * PTCM * UU, ang = (sh.angle === undefined ? 45 : sh.angle) * Math.PI / 180;
    const op = 1 - (sh.transparency || 0);
    return `<path d="${d}" transform="translate(${Math.cos(ang) * dist} ${Math.sin(ang) * dist})" fill="${sh.color || '#808080'}" fill-opacity="${op}" stroke="${sh.color || '#808080'}" stroke-opacity="${op}" stroke-width="${(s.line && s.line.weight || 0.75) * PTCM * UU}" ${sh.blur ? 'filter="url(#blur)"' : ''}/>`;
  }
  function scalePath(d) { return d.replace(/-?\d+(\.\d+)?/g, n => String(Math.round(parseFloat(n) * UU * 10) / 10)); }
  function arcSafePath(dcm) {
    // A 명령의 플래그(0/1)는 비율을 곱하지 않도록 따로 처리
    return dcm.replace(/A([^A-Za-z]+)/g, (m, args) => { const a = args.trim().split(/[\s,]+/).map(Number); const out = []; for (let i = 0; i + 6 < a.length + 1; i += 7) out.push([a[i] * UU, a[i + 1] * UU, a[i + 2], a[i + 3], a[i + 4], a[i + 5] * UU, a[i + 6] * UU].map(x => Math.round(x * 10) / 10).join(' ')); return 'A' + out.join(' A') + ' '; })
      .replace(/([MLHVCQZ])([^A-Za-z]*)/g, (m, c, args) => c + args.replace(/-?\d+(\.\d+)?/g, n => String(Math.round(parseFloat(n) * UU * 10) / 10)));
  }
  function shapeSvg(pres, slide, s, opts) {
    const o = opts || {};
    if (s.hidden && !o.showHidden) return '';
    if (s.kind === 'group') return `<g data-id="${s.id}">${s.children.map(c => shapeSvg(pres, slide, c, Object.assign({}, o, { inGroup: s.id }))).join('')}</g>`;
    const hitAttr = `data-id="${o.inGroup || s.id}" data-child="${s.id}"`;
    if (P.isLine(s)) {
      const pts = P.linePoints(pres, slide, s);
      const dd = 'M' + pts.map(p => (p.x * UU).toFixed(1) + ' ' + (p.y * UU).toFixed(1)).join(' L');
      const mk = (end) => s.line && s.line[end] === 'arrow' ? `marker-${end === 'end' ? 'end' : 'start'}="url(#arr_${(s.line.color || '#000').slice(1)})"` : '';
      return `<path d="${dd}" fill="none" ${strokeAttrs(s.line)} ${mk('end')} ${mk('begin')}/><path class="hit" ${hitAttr} d="${dd}" fill="none" stroke="transparent" stroke-width="30"/>`;
    }
    if (s.kind === 'placeholder' && !s.text.value) {
      if (o.print) return '';
      return `<rect class="hit" ${hitAttr} x="${s.x * UU}" y="${s.y * UU}" width="${s.w * UU}" height="${s.h * UU}" fill="transparent" stroke="#a6a6a6" stroke-dasharray="6 4" stroke-width="2"/>` + textHtml(s);
    }
    const dcm = P.geomPath(s);
    const d = arcSafePath(dcm);
    const openPath = ['leftBrace', 'rightBrace', 'leftBracket', 'rightBracket'].includes(s.kind);
    const fill = openPath ? 'none' : (s.fill || 'none');
    return shadowOf(s, d) + `<path class="hit" ${hitAttr} d="${d}" fill="${fill === 'none' ? (o.print ? 'none' : 'transparent') : fill}" ${strokeAttrs(s.line)}/>` + textHtml(s);
  }
  function defsSvg() {
    const colors = ['000000', '4472C4', '2F528F', 'FF0000', '808080', '595959', 'FFFFFF'];
    return `<defs><filter id="blur"><feGaussianBlur stdDeviation="6"/></filter>${colors.map(c => `<marker id="arr_${c}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse" markerUnits="strokeWidth"><path d="M0 0 L10 5 L0 10 z" fill="#${c}"/></marker>`).join('')}</defs>`;
  }
  function slideSvg(pres, si, opts) {
    const slide = pres.slides[si];
    const W = pres.size.w * UU, H = pres.size.h * UU;
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="100%" height="100%" preserveAspectRatio="none" style="display:block">${defsSvg()}<rect x="0" y="0" width="${W}" height="${H}" fill="#fff"/>${slide.shapes.map(s => shapeSvg(pres, slide, s, opts)).join('')}</svg>`;
  }

  // ------------------------------------------------------------------ 앱
  function PptApp(win, ctx) {
    const app = this;
    this.el = win;
    let pres = ctx.model();
    let cur = 0, sel = [], child = null, tool = null, lock = false, pane = null, view = 'normal', backstage = null, guides = false, grid = false, fmtClip = null, shapeClip = null;
    const undo = [], redo = [];
    let fileName = (ctx.set.id === 'equip' ? 'ch14_작업' : ctx.set.id + '_작업') + '.pptx', savedOnce = false;
    this.fileName = () => (savedOnce ? fileName.replace(/\.pptx$/, '') : '프레젠테이션1') + ' - PowerPoint';
    this.model = () => pres;
    const slide = () => pres.slides[cur];
    const selShapes = () => sel.map(id => slide().shapes.find(s => s.id === id)).filter(Boolean);
    const target = () => { if (child) { const c = P.findShape(pres, cur, child); return c ? [c] : []; } return selShapes(); };
    const flatTargets = () => { const out = []; P.eachShape(target(), s => { if (s.kind !== 'group') out.push(s); }); return out; };

    function exec(fn) {
      if (!pres) return;
      const snap = JSON.stringify(pres);
      try { fn(); } catch (e) { console.error(e); msg({ title: 'Microsoft PowerPoint', text: e.message, icon: 'err' }); return; }
      undo.push(snap); if (undo.length > 80) undo.shift(); redo.length = 0;
      ctx.changed();
      render();
    }
    function doUndo() { if (!undo.length) return; redo.push(JSON.stringify(pres)); pres = JSON.parse(undo.pop()); ctx.setModel(pres); sel = sel.filter(id => slide() && slide().shapes.some(s => s.id === id)); render(); }
    function doRedo() { if (!redo.length) return; undo.push(JSON.stringify(pres)); pres = JSON.parse(redo.pop()); ctx.setModel(pres); render(); }
    function save(asNew) {
      if (!pres) return;
      const go = () => { savedOnce = true; ctx.saved(); render(); toast(fileName + ' 저장'); };
      if (savedOnce && !asNew) return go();
      const inp = h('input', { type: 'text', value: fileName.replace(/\.pptx$/, ''), style: { width: '260px' } });
      dialog({ title: '다른 이름으로 저장', width: 440, body: h('div', {}, h('div', { class: 'row' }, icon('folder'), '바탕 화면 › 비번호 폴더'), h('div', { class: 'grid2' }, '파일 이름:', inp, '파일 형식:', h('select', {}, h('option', {}, 'PowerPoint 프레젠테이션 (*.pptx)')))),
        buttons: [{ label: '저장', primary: true, onClick: () => { fileName = inp.value.trim() + '.pptx'; go(); } }, { label: '취소', cancel: true }] });
    }

    // ------------------------------------------------------------------ 틀
    win.innerHTML = '';
    win.style.setProperty('--app', '#c43e1c'); win.style.setProperty('--app2', '#a8331a');
    const titleName = h('span', { class: 'name' });
    const qU = h('button', { title: '실행 취소 (Ctrl+Z)', onclick: doUndo }, icon('undo')), qR = h('button', { title: '다시 실행 (Ctrl+Y)', onclick: doRedo }, icon('redo'));
    win.appendChild(h('div', { class: 'titlebar' }, h('span', { class: 'qat' }, h('button', { title: '저장 (Ctrl+S)', onclick: () => save() }, icon('save')), qU, qR), titleName, h('span', { style: { width: '120px' } })));
    const ribHost = h('div');
    win.appendChild(ribHost);
    const thumbs = h('div', { class: 'pp-thumbs' });
    const stage = h('div', { class: 'pp-stage', tabindex: 0 });
    const side = h('div', { style: { display: 'flex' } });
    win.appendChild(h('div', { class: 'pp-main' }, thumbs, stage, side));
    const status = h('div', { class: 'statusbar' });
    win.appendChild(status);

    // ------------------------------------------------------------------ 리본
    const out = (label, ic, type) => ({ type: type || 'small', label, icon: ic, out: true });
    const hasSel = () => !!pres && target().length > 0;
    const shapeGallery = () => [{ html: close => {
      const wrap = h('div', { style: { width: '330px', padding: '4px 0' } });
      for (const [cat, items] of P.GALLERY) {
        wrap.appendChild(h('div', { class: 'mhead' }, cat));
        const g = h('div', { class: 'mgrid', style: { gridTemplateColumns: 'repeat(10, 28px)' } });
        for (const [k, label] of items) {
          const b = h('button', { class: 'shapecell', title: label }, galleryIcon(k));
          b.addEventListener('click', () => { close(); setTool(k, false); });
          b.addEventListener('contextmenu', e => { e.preventDefault(); menu(b, [{ label: '그리기 잠금 모드(L)', onClick: () => { close(); setTool(k, true); } }], { at: { x: e.clientX, y: e.clientY }, level: 1 }); });
          g.appendChild(b);
        }
        wrap.appendChild(g);
      }
      return wrap;
    } }];
    const fillMenu = () => [{ html: OA.ui.palette(c => exec(() => flatTargets().filter(s => !P.isLine(s)).forEach(s => { s.fill = c; })), close => h('div', {},
      h('div', { class: 'mi', onclick: () => { close(); exec(() => flatTargets().forEach(s => { s.fill = null; })); } }, '채우기 없음(N)'),
      h('div', { class: 'mi', onclick: () => { close(); const t0 = flatTargets()[0]; OA.ui.colorDialog(t0 && t0.fill, c => exec(() => flatTargets().filter(s => !P.isLine(s)).forEach(s => { s.fill = c; }))); } }, '다른 채우기 색(M)...'),
      h('div', { class: 'mi dis' }, '그림(P)...'), h('div', { class: 'mi dis' }, '그라데이션(G)'), h('div', { class: 'mi dis' }, '질감(T)'))) }];
    const lineMenu = () => [{ html: OA.ui.palette(c => exec(() => flatTargets().forEach(s => { s.line = Object.assign({ weight: 0.75, dash: 'solid' }, s.line, { color: c }); })), close => h('div', {},
      h('div', { class: 'mi', onclick: () => { close(); exec(() => flatTargets().forEach(s => { s.line = null; })); } }, '윤곽선 없음(N)'),
      h('div', { class: 'mi', onclick: () => { close(); OA.ui.colorDialog('#000000', c => exec(() => flatTargets().forEach(s => { s.line = Object.assign({ weight: 0.75, dash: 'solid' }, s.line, { color: c }); }))); } }, '다른 윤곽선 색(M)...'))) },
      { sep: true },
      { label: '두께(W)', sub: WEIGHTS.map(([w, l]) => ({ label: l, onClick: () => exec(() => flatTargets().forEach(s => { s.line = Object.assign({ color: '#000000', dash: 'solid' }, s.line, { weight: w }); })) })).concat([{ sep: true }, { label: '다른 선(L)...', onClick: () => openFmtPane('fill') }]) },
      { label: '대시(S)', sub: DASH_NAMES.map(([k, l]) => ({ label: l, onClick: () => exec(() => flatTargets().forEach(s => { s.line = Object.assign({ color: '#000000', weight: 0.75 }, s.line, { dash: k }); })) })).concat([{ sep: true }, { label: '다른 선(L)...', onClick: () => openFmtPane('fill') }]) },
      { label: '화살표(R)', sub: [['none', '화살표 스타일 1(선)'], ['end', '화살표 스타일 2(끝 화살표)'], ['begin', '화살표 스타일 3(시작 화살표)'], ['both', '화살표 스타일 5(양쪽 화살표)']].map(([k, l]) => ({ label: l, disabled: !flatTargets().some(s => P.isLine(s) || s.kind === 'connElbow'), onClick: () => exec(() => flatTargets().filter(s => P.isLine(s)).forEach(s => { s.line = Object.assign({ color: '#000000', weight: 0.75, dash: 'solid' }, s.line, { begin: k === 'begin' || k === 'both' ? 'arrow' : null, end: k === 'end' || k === 'both' ? 'arrow' : null }); })) })).concat([{ sep: true }, { label: '다른 화살표(M)...', onClick: () => openFmtPane('fill') }]) }];
    const effectMenu = () => [{ label: '미리 설정(P)', disabled: true }, { label: '그림자(S)', sub: [{ header: '그림자 없음' }, { label: '그림자 없음', onClick: () => exec(() => flatTargets().forEach(s => { s.shadow = null; })) }, { header: '바깥쪽' },
      { label: '오프셋: 오른쪽 아래', onClick: () => exec(() => flatTargets().forEach(s => { s.shadow = { color: '#000000', transparency: 0.6, blur: 4, angle: 45, dist: 3 }; })) },
      { label: '오프셋: 아래쪽', onClick: () => exec(() => flatTargets().forEach(s => { s.shadow = { color: '#000000', transparency: 0.6, blur: 4, angle: 90, dist: 3 }; })) },
      { label: '오프셋: 오른쪽', onClick: () => exec(() => flatTargets().forEach(s => { s.shadow = { color: '#000000', transparency: 0.6, blur: 4, angle: 0, dist: 3 }; })) },
      { header: '안쪽' }, { label: '안쪽: 가운데', disabled: true }, { sep: true }, { label: '그림자 옵션(S)...', onClick: () => openFmtPane('effect') }] },
      { label: '반사(R)', disabled: true }, { label: '네온(G)', disabled: true }, { label: '부드러운 가장자리(E)', disabled: true }, { label: '입체 효과(B)', disabled: true }, { label: '3차원 회전(D)', disabled: true }];
    const alignMenu = () => [['left', '왼쪽 맞춤(L)'], ['center', '가운데 맞춤(C)'], ['right', '오른쪽 맞춤(R)'], ['top', '위쪽 맞춤(T)'], ['middle', '중간 맞춤(M)'], ['bottom', '아래쪽 맞춤(B)']].map(([k, l]) => ({ label: l, onClick: () => exec(() => P.align(pres, cur, selShapes(), k, app.alignToSlide)) }))
      .concat([{ sep: true }, { label: '가로 간격을 동일하게(H)', disabled: () => selShapes().length < 3, onClick: () => exec(() => P.distribute(pres, cur, selShapes(), 'h')) }, { label: '세로 간격을 동일하게(V)', disabled: () => selShapes().length < 3, onClick: () => exec(() => P.distribute(pres, cur, selShapes(), 'v')) }, { sep: true },
        { label: '슬라이드에 맞춤(A)', checked: !!app.alignToSlide, onClick: () => { app.alignToSlide = true; } }, { label: '선택한 개체 맞춤(O)', checked: !app.alignToSlide, onClick: () => { app.alignToSlide = false; } }, { sep: true },
        { label: '눈금선 보기(S)', checked: grid, onClick: () => { grid = !grid; render(); } }]);
    const groupMenu = () => [{ label: '그룹(G)', icon: 'group', disabled: () => selShapes().length < 2, onClick: groupSel }, { label: '그룹 다시 묶기(E)', disabled: true }, { label: '그룹 해제(U)', icon: 'ungroup', disabled: () => !selShapes().some(s => s.kind === 'group'), onClick: ungroupSel }];
    const orderMenu = front => (front ? [['front', '맨 앞으로 가져오기(R)'], ['forward', '앞으로 가져오기(F)']] : [['back', '맨 뒤로 보내기(K)'], ['backward', '뒤로 보내기(B)']]).map(([k, l]) => ({ label: l, icon: front ? 'front' : 'back', onClick: () => exec(() => P.zorder(pres, cur, sel, k)) }));
    const layoutMenu = (apply) => P.LAYOUTS.map(([k, l]) => ({ label: l, icon: 'slidelayout', onClick: () => apply(k) }));
    const fontSel = () => { const t = flatTargets().find(s => s.text); return t ? t.text : null; };
    const setText = patch => exec(() => flatTargets().filter(s => s.text).forEach(s => Object.assign(s.text, patch)));
    const sizeStep = dir => exec(() => flatTargets().filter(s => s.text).forEach(s => { const cs = s.text.size || 18; const i = SIZES.findIndex(x => x >= cs - 0.01); let n; if (dir > 0) n = SIZES.find(x => x > cs + 0.01) || cs + 4; else n = [...SIZES].reverse().find(x => x < cs - 0.01) || Math.max(1, cs - 1); void i; s.text.size = n; }));
    const rib = ribbon(ribHost, {
      file: () => openBackstage(pres ? 'info' : 'new'),
      initial: 'home',
      tabs: [
        { id: 'home', label: '홈', when: () => view === 'normal', groups: [
          { label: '클립보드', items: [{ type: 'big', label: '붙여넣기', icon: 'paste', onClick: () => pasteShapes() }, { type: 'stack', items: [{ type: 'small', label: '잘라내기', icon: 'cut', onClick: () => { copyShapes(); deleteSel(); } }, { type: 'small', label: '복사', icon: 'copy', onClick: () => copyShapes() }, { type: 'small', label: '서식 복사', icon: 'brush', onClick: () => copyFormat() }] }] },
          { label: '슬라이드', items: [{ type: 'big', label: '새 슬라이드', icon: 'newslide', onClick: () => newSlide(), menu: () => layoutMenu(k => newSlide(k)) }, { type: 'stack', items: [{ type: 'small', label: '레이아웃', icon: 'slidelayout', menu: () => layoutMenu(k => exec(() => P.setLayout(pres, cur, k))) }, out('원래대로', 'refresh'), out('구역', 'fieldlist')] }] },
          { label: '글꼴', items: [{ type: 'stack', items: [
            { type: 'row', items: [{ type: 'combo', width: 110, value: () => fontSel() ? fontSel().font : '맑은 고딕', list: ['맑은 고딕', '굴림', '돋움', '바탕', 'Arial'], onChange: v => setText({ font: v }) },
              { type: 'combo', width: 48, value: () => fontSel() ? fontSel().size : 18, list: SIZES, onChange: v => { const n = parseFloat(v); if (n > 0) setText({ size: n }); } },
              { type: 'icon', icon: 'growFont', label: '글꼴 크기 크게 (Ctrl+Shift+>)', onClick: () => sizeStep(1) }, { type: 'icon', icon: 'shrinkFont', label: '글꼴 크기 작게 (Ctrl+Shift+<)', onClick: () => sizeStep(-1) }] },
            { type: 'row', items: [{ type: 'icon', icon: 'bold', label: '굵게', pressed: () => fontSel() && fontSel().bold, onClick: () => setText({ bold: !(fontSel() && fontSel().bold) }) }, { type: 'icon', icon: 'italic', label: '기울임꼴', pressed: () => fontSel() && fontSel().italic, onClick: () => setText({ italic: !(fontSel() && fontSel().italic) }) }, { type: 'icon', icon: 'underline', label: '밑줄', pressed: () => fontSel() && fontSel().underline, onClick: () => setText({ underline: !(fontSel() && fontSel().underline) }) },
              { type: 'icon', icon: 'fontcolor', label: '글꼴 색', menu: () => [{ html: OA.ui.palette(c => setText({ color: c })) }] }] }] }] },
          { label: '단락', items: [{ type: 'stack', items: [
            { type: 'row', items: [{ type: 'icon', icon: 'bullets', label: '글머리 기호', onClick: () => setText({ bullet: fontSel() && fontSel().bullet ? null : '•' }), menu: () => bulletMenu() }, out('번호 매기기', 'numbering', 'icon'), out('목록 수준 줄임', 'alignL', 'icon'), out('목록 수준 늘림', 'alignR', 'icon')] },
            { type: 'row', items: [{ type: 'icon', icon: 'alignL', label: '왼쪽 맞춤', pressed: () => fontSel() && fontSel().align === 'left', onClick: () => setText({ align: 'left' }) }, { type: 'icon', icon: 'alignC', label: '가운데 맞춤', pressed: () => fontSel() && fontSel().align === 'center', onClick: () => setText({ align: 'center' }) }, { type: 'icon', icon: 'alignR', label: '오른쪽 맞춤', pressed: () => fontSel() && fontSel().align === 'right', onClick: () => setText({ align: 'right' }) },
              { type: 'icon', icon: 'alignM', label: '텍스트 맞춤', menu: () => [['top', '위쪽(T)'], ['middle', '중간(M)'], ['bottom', '아래쪽(B)']].map(([k, l]) => ({ label: l, checked: fontSel() && fontSel().anchor === k, onClick: () => setText({ anchor: k }) })) }] }] }] },
          { label: '그리기', items: [{ type: 'big', label: '도형', icon: 'shapes', menu: shapeGallery }, { type: 'big', label: '정렬', icon: 'arrange', menu: () => [{ header: '개체 순서' }, ...orderMenu(true), ...orderMenu(false), { header: '개체 그룹화' }, ...groupMenu(), { header: '개체 위치' }, { label: '맞춤(A)', sub: alignMenu }, { sep: true }, { label: '선택 창(P)...', icon: 'selpane', onClick: () => togglePane('sel') }] },
            { type: 'stack', items: [{ type: 'small', label: '도형 채우기', icon: 'shapeFill', disabled: () => !hasSel(), menu: fillMenu }, { type: 'small', label: '도형 윤곽선', icon: 'shapeLine', disabled: () => !hasSel(), menu: lineMenu }, { type: 'small', label: '도형 효과', icon: 'shapeEffect', disabled: () => !hasSel(), menu: effectMenu }] }] },
          { label: '편집', items: [{ type: 'stack', items: [out('찾기', 'find'), out('바꾸기', 'find'), { type: 'small', label: '선택', icon: 'selectArrow', menu: () => [{ label: '모두 선택(A)', onClick: selectAll }, { label: '개체 선택(O)', disabled: true }, { label: '선택 창(P)...', icon: 'selpane', onClick: () => togglePane('sel') }] }] }] }] },
        { id: 'insert', label: '삽입', when: () => view === 'normal', groups: [
          { label: '슬라이드', items: [{ type: 'big', label: '새 슬라이드', icon: 'newslide', onClick: () => newSlide(), menu: () => layoutMenu(k => newSlide(k)) }] },
          { label: '표', items: [out('표', 'table', 'big')] }, { label: '이미지', items: [out('그림', 'image', 'big')] },
          { label: '일러스트레이션', items: [{ type: 'big', label: '도형', icon: 'shapes', menu: shapeGallery }, out('아이콘', 'shapes', 'big'), out('SmartArt', 'layout', 'big'), out('차트', 'colchart', 'big')] },
          { label: '텍스트', items: [{ type: 'big', label: '텍스트 상자', icon: 'textboxP', onClick: () => setTool('textbox', false) }, out('머리글/바닥글', 'titleP', 'big'), out('WordArt', 'label', 'big')] },
          { label: '기호', items: [out('기호', 'sum', 'big')] }] },
        { id: 'draw', label: '그리기', when: () => view === 'normal', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '도구', items: [out('펜', 'design', 'big')] }] },
        { id: 'design', label: '디자인', when: () => view === 'normal', groups: [
          { label: '테마', items: [out('테마', 'colors', 'big')] }, { label: '적용', items: [out('적용', 'colors', 'big')] },
          { label: '사용자 지정', items: [{ type: 'big', label: '슬라이드 크기', icon: 'slidesize', menu: () => [{ label: '표준(4:3)', checked: pres && pres.size.kind === 'standard', onClick: () => slideSize('standard') }, { label: '와이드스크린(16:9)', checked: pres && pres.size.kind === 'wide', onClick: () => slideSize('wide') }, { sep: true }, { label: '사용자 지정 슬라이드 크기(C)...', onClick: () => customSize() }] }, out('배경 서식', 'shapeFill', 'big')] },
          { label: '디자이너', items: [out('디자인 아이디어', 'wand', 'big')] }] },
        { id: 'trans', label: '전환', when: () => view === 'normal', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '전환', items: [out('미리 보기', 'preview', 'big')] }] },
        { id: 'anim', label: '애니메이션', when: () => view === 'normal', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '애니메이션', items: [out('미리 보기', 'preview', 'big')] }] },
        { id: 'show', label: '슬라이드 쇼', when: () => view === 'normal', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '슬라이드 쇼 시작', items: [out('처음부터', 'preview', 'big')] }] },
        { id: 'review', label: '검토', when: () => view === 'normal', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '언어 교정', items: [out('맞춤법 검사', 'find', 'big')] }] },
        { id: 'view', label: '보기', when: () => view === 'normal', groups: [
          { label: '프레젠테이션 보기', items: [{ type: 'big', label: '기본', icon: 'slidelayout', pressed: true, onClick: () => {} }, out('개요 보기', 'fieldlist', 'big'), out('여러 슬라이드', 'twopage', 'big'), out('슬라이드 노트', 'paper', 'big'), out('읽기용 보기', 'preview', 'big')] },
          { label: '마스터 보기', items: [out('슬라이드 마스터', 'master', 'big'), { type: 'big', label: '유인물 마스터', icon: 'handout', onClick: () => { if (pres) { view = 'handout'; rib.select('handout'); render(); } } }, out('슬라이드 노트 마스터', 'master', 'big')] },
          { label: '표시', items: [{ type: 'stack', items: [{ type: 'check', label: '눈금자', checked: false, disabled: true, onChange: () => {} }, { type: 'check', label: '눈금선', checked: () => grid, onChange: v => { grid = v; render(); } }, { type: 'check', label: '안내선', checked: () => guides, onChange: v => { guides = v; render(); } }] }] },
          { label: '확대/축소', items: [out('확대/축소', 'zoom', 'big'), out('창에 맞춤', 'onepage', 'big')] },
          { label: '컬러/회색조', items: [{ type: 'stack', items: [{ type: 'small', label: '컬러', icon: 'colors', pressed: () => view === 'normal', onClick: () => { view = 'normal'; render(); } }, { type: 'small', label: '회색조', icon: 'grayscale', onClick: () => { if (pres) { view = 'gray'; rib.select('gray'); render(); } } }, out('흑백', 'grayscale')] }] }] },
        { id: 'help', label: '도움말', when: () => view === 'normal', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '도움말', items: [out('도움말', 'find', 'big')] }] },
        { id: 'format', label: '도형 서식', ctx: ' ', when: () => view === 'normal' && hasSel(), groups: [
          { label: '도형 삽입', items: [{ type: 'big', label: '도형', icon: 'shapes', menu: shapeGallery }, { type: 'stack', items: [out('도형 편집', 'design'), { type: 'small', label: '텍스트 상자', icon: 'textboxP', onClick: () => setTool('textbox', false) }] }] },
          { label: '도형 스타일', launcher: () => openFmtPane('fill'), items: [out('빠른 스타일', 'cellstyle', 'big'), { type: 'stack', items: [{ type: 'small', label: '도형 채우기', icon: 'shapeFill', menu: fillMenu }, { type: 'small', label: '도형 윤곽선', icon: 'shapeLine', menu: lineMenu }, { type: 'small', label: '도형 효과', icon: 'shapeEffect', menu: effectMenu }] }] },
          { label: 'WordArt 스타일', items: [out('텍스트 채우기', 'fontcolor')] },
          { label: '접근성', items: [out('대체 텍스트', 'textboxP', 'big')] },
          { label: '정렬', items: [{ type: 'stack', items: [{ type: 'small', label: '앞으로 가져오기', icon: 'front', onClick: () => exec(() => P.zorder(pres, cur, sel, 'forward')), menu: () => orderMenu(true) }, { type: 'small', label: '뒤로 보내기', icon: 'back', onClick: () => exec(() => P.zorder(pres, cur, sel, 'backward')), menu: () => orderMenu(false) }, { type: 'small', label: '선택 창', icon: 'selpane', pressed: () => pane === 'sel', onClick: () => togglePane('sel') }] },
            { type: 'stack', items: [{ type: 'small', label: '맞춤', icon: 'alignObj', menu: alignMenu }, { type: 'small', label: '그룹화', icon: 'group', menu: groupMenu }, out('회전', 'refresh')] }] },
          { label: '크기', launcher: () => openFmtPane('size'), items: [{ type: 'stack', items: [
            { type: 'row', items: [{ type: 'html', render: () => h('span', { style: { lineHeight: '22px', padding: '0 4px' } }, '높이:') }, { type: 'combo', width: 72, value: () => { const b = selShapes()[0] && P.bbox(selShapes()[0]); return b ? b.h.toFixed(2) + ' cm' : ''; }, list: [], onChange: v => resizeTo(null, parseFloat(v)) }] },
            { type: 'row', items: [{ type: 'html', render: () => h('span', { style: { lineHeight: '22px', padding: '0 4px' } }, '너비:') }, { type: 'combo', width: 72, value: () => { const b = selShapes()[0] && P.bbox(selShapes()[0]); return b ? b.w.toFixed(2) + ' cm' : ''; }, list: [], onChange: v => resizeTo(parseFloat(v), null) }] }] }] }] },
        { id: 'handout', label: '유인물 마스터', when: () => view === 'handout', groups: [
          { label: '페이지 설정', items: [{ type: 'big', label: '유인물 방향', icon: 'orient', menu: () => [{ label: '세로(P)', checked: pres.handout.orient !== 'landscape', onClick: () => exec(() => { pres.handout.orient = 'portrait'; }) }, { label: '가로(L)', checked: pres.handout.orient === 'landscape', onClick: () => exec(() => { pres.handout.orient = 'landscape'; }) }] }, out('슬라이드 크기', 'slidesize', 'big'),
            { type: 'big', label: '페이지당 슬라이드 수', icon: 'handout', menu: () => [1, 2, 3, 4, 6, 9].map(n => ({ label: n + '슬라이드', checked: pres.handout.perPage === n, onClick: () => exec(() => { pres.handout.perPage = n; }) })) }] },
          { label: '개체 틀', items: [{ type: 'stack', items: [{ type: 'check', label: '머리글', checked: () => pres.handout.header, onChange: v => exec(() => { pres.handout.header = v; }) }, { type: 'check', label: '날짜', checked: () => pres.handout.date, onChange: v => exec(() => { pres.handout.date = v; }) }] },
            { type: 'stack', items: [{ type: 'check', label: '바닥글', checked: () => pres.handout.footer, onChange: v => exec(() => { pres.handout.footer = v; }) }, { type: 'check', label: '페이지 번호', checked: () => pres.handout.pageNum, onChange: v => exec(() => { pres.handout.pageNum = v; }) }] }] },
          { label: '배경', items: [out('배경 스타일', 'shapeFill', 'big')] },
          { label: '닫기', items: [{ type: 'big', label: '마스터 보기 닫기', icon: 'closeView', onClick: () => { view = 'normal'; rib.select('view'); render(); } }] }] },
        { id: 'gray', label: '회색조', when: () => view === 'gray', groups: [
          { label: '선택한 개체 변경', items: [out('자동', 'grayscale', 'big'), out('회색조', 'grayscale', 'big')] },
          { label: '닫기', items: [{ type: 'big', label: '컬러 보기로 돌아가기', icon: 'closeView', onClick: () => { view = 'normal'; rib.select('view'); render(); } }] }] },
      ],
    });
    function galleryIcon(k) {
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('viewBox', '0 0 100 100'); svg.setAttribute('width', 20); svg.setAttribute('height', 20);
      const base = { arrow: 'connStraight', arrow2: 'connStraight', elbowArrow: 'connElbow', elbowArrow2: 'connElbow' }[k] || k;
      if (['line', 'connStraight', 'connElbow'].includes(base)) {
        const d = base === 'connElbow' ? 'M10 85 H50 V15 H90' : 'M10 90 L90 10';
        const arrowEnd = ['arrow', 'arrow2', 'elbowArrow', 'elbowArrow2'].includes(k);
        svg.innerHTML = `<defs><marker id="g${k}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0 0 L10 5 L0 10z" fill="#444"/></marker></defs><path d="${d}" stroke="#444" stroke-width="7" fill="none" ${arrowEnd ? `marker-end="url(#g${k})"` : ''} ${k.endsWith('2') ? `marker-start="url(#g${k})"` : ''}/>`;
        return svg;
      }
      if (k === 'textbox') { svg.innerHTML = '<rect x="10" y="20" width="80" height="60" fill="none" stroke="#444" stroke-width="5" stroke-dasharray="10 6"/><text x="50" y="62" font-size="36" text-anchor="middle" fill="#444">가</text>'; return svg; }
      const s = { kind: base, x: 0.1, y: 0.15, w: 0.8, h: 0.7, adj: P.KINDS[base] && P.KINDS[base].adj };
      const dd = arcSafePath(P.geomPath(s)).replace(/-?\d+(\.\d+)?/g, n => n);
      svg.innerHTML = `<path d="${dd}" fill="#fff" stroke="#444" stroke-width="5"/>`;
      return svg;
    }

    // ------------------------------------------------------------------ 명령
    function setTool(k, lk) { if (!pres) return; tool = k; lock = !!lk; sel = []; child = null; render(); stage.focus(); if (lk) toast('그리기 잠금 모드: 같은 도형을 계속 그립니다. 끝내려면 Esc'); }
    function newSlide(layout) { if (!pres) return; exec(() => { cur = P.addSlide(pres, cur, layout); }); sel = []; render(); }
    function slideSize(kind) {
      if (!pres || pres.size.kind === kind) return;
      dialog({ title: 'Microsoft PowerPoint', width: 520, help: false, body: h('div', {}, h('p', {}, '새 슬라이드 크기로 조정 중입니다. 콘텐츠의 크기를 최대화하시겠습니까, 아니면 새 슬라이드에 맞게 축소하시겠습니까?'),
        h('div', { class: 'row', style: { gap: '18px' } }, h('div', { class: 'col', style: { alignItems: 'center' } }, icon('slidesize', 60), '최대화'), h('div', { class: 'col', style: { alignItems: 'center' } }, icon('onepage', 60), '맞춤 확인'))),
        buttons: [{ label: '최대화', onClick: () => exec(() => P.setSlideSize(pres, kind, 'max')) }, { label: '맞춤 확인', primary: true, onClick: () => exec(() => P.setSlideSize(pres, kind, 'fit')) }, { label: '취소', cancel: true }] });
    }
    function customSize() {
      const sz = h('select', {}, h('option', { value: 'standard', selected: pres.size.kind === 'standard' }, '화면 슬라이드 쇼(4:3)'), h('option', { value: 'wide', selected: pres.size.kind === 'wide' }, '와이드스크린'));
      dialog({ title: '슬라이드 크기', width: 420, body: h('div', {}, h('div', { class: 'grid2' }, '슬라이드 크기:', sz, '너비:', h('input', { type: 'text', readOnly: true, value: pres.size.w + ' cm' }), '높이:', h('input', { type: 'text', readOnly: true, value: pres.size.h + ' cm' })), h('p', { class: 'note' }, '방향: 슬라이드 가로, 슬라이드 노트·유인물 및 개요 세로')),
        buttons: [{ label: '확인', primary: true, onClick: () => { setTimeout(() => slideSize(sz.value), 0); } }, { label: '취소', cancel: true }] });
    }
    function selectAll() { if (!pres) return; sel = slide().shapes.filter(s => !s.hidden).map(s => s.id); child = null; render(); }
    function deleteSel() { if (!sel.length && !child) return; const ids = child ? [child] : sel.slice(); exec(() => P.removeShapes(pres, cur, ids)); sel = []; child = null; render(); }
    function groupSel() { if (selShapes().length < 2) return; let g; exec(() => { g = P.group(pres, cur, sel); }); if (g) sel = [g.id]; render(); }
    function ungroupSel() { const gs = selShapes().filter(s => s.kind === 'group'); if (!gs.length) return; let kids = []; exec(() => { gs.forEach(g => { kids = kids.concat(P.ungroup(pres, cur, g.id)); }); }); sel = kids.map(k => k.id); render(); }
    function copyShapes() { if (!sel.length) return; shapeClip = U.clone(selShapes()); }
    function pasteShapes() { if (!shapeClip || !pres) return; let made = []; exec(() => { made = P.duplicate(pres, cur, shapeClip, 0.25, 0.25); }); sel = made.map(s => s.id); shapeClip = U.clone(made); render(); }
    function duplicateSel() { if (!sel.length) return; let made = []; exec(() => { made = P.duplicate(pres, cur, selShapes(), 0.25, 0.25); }); sel = made.map(s => s.id); render(); }
    function copyFormat() { const t = flatTargets()[0]; if (t) { fmtClip = P.styleOf(t); toast('서식을 복사했습니다(Ctrl+Shift+V로 붙여 넣기).'); } }
    function pasteFormat() { if (!fmtClip) return; exec(() => flatTargets().forEach(s => { if (!P.isLine(s)) s.fill = fmtClip.fill; s.line = U.clone(fmtClip.line); s.shadow = U.clone(fmtClip.shadow); if (s.text && fmtClip.text) { const v = s.text.value; Object.assign(s.text, U.clone(fmtClip.text)); s.text.value = v; } })); }
    function resizeTo(w, hh) { exec(() => selShapes().forEach(s => { if (s.kind === 'group' || P.isLine(s)) return; if (w > 0) s.w = w; if (hh > 0) s.h = hh; })); }
    function bulletMenu() {
      return [{ html: close => h('div', { class: 'mgrid', style: { gridTemplateColumns: 'repeat(4, 54px)' } }, BULLETS.map(([b, l]) => h('button', { class: 'btn', title: l, style: { height: '46px', fontSize: '18px' }, onclick: () => { close(); setText({ bullet: b }); } }, b || '없음'))) },
        { sep: true }, { label: '글머리 기호 및 번호 매기기(N)...', onClick: () => bulletDialog() }];
    }
    function bulletDialog() {
      const t = fontSel();
      let pick = t ? t.bullet : null;
      const cells = BULLETS.map(([b, l]) => { const c = h('div', { title: l, style: { height: '74px', border: '1px solid #ccc', display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'flex-start', padding: '4px 12px', cursor: 'pointer', fontSize: b === '•' ? '12px' : '16px', outline: pick === b ? '3px solid #f4c7a8' : '' } }, b ? [1, 2, 3].map(() => h('div', {}, b + ' ———')) : h('div', { style: { fontSize: '16px', alignSelf: 'center' } }, '없음')); c.onclick = () => { pick = b; cells.forEach((x, i) => { x.style.outline = BULLETS[i][0] === pick ? '3px solid #f4c7a8' : ''; }); }; return c; });
      dialog({ title: '글머리 기호 및 번호 매기기', width: 470, body: tabs([{ id: 'b', label: '글머리 기호(B)', body: h('div', {}, h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '6px', border: '1px solid #aaa', padding: '6px' } }, cells),
        h('div', { class: 'row' }, '텍스트 크기(S):', h('input', { type: 'number', value: 100, style: { width: '60px' }, disabled: true }), '%', h('span', { style: { flex: 1 } }), h('button', { class: 'btn', disabled: true }, '그림(P)...')),
        h('div', { class: 'row' }, '색(C):', h('span', { style: { display: 'inline-block', width: '30px', height: '18px', background: '#000' } }), h('span', { style: { flex: 1 } }), h('button', { class: 'btn', onclick: () => symbolDialog(b => { pick = b; }) }, '사용자 지정(U)...'))) },
        { id: 'n', label: '번호 매기기(N)', body: h('div', { class: 'note', style: { padding: '30px' } }, '실기 과제에 쓰지 않는 탭입니다.') }]),
        buttons: [{ label: '원래대로(E)', left: true, disabled: true }, { label: '확인', primary: true, onClick: () => setText({ bullet: pick }) }, { label: '취소', cancel: true }] });
    }
    function symbolDialog(cb) {
      const syms = ['❖', '◆', '◇', '■', '□', '●', '○', '★', '☆', '▶', '▷', '➢', '✓', '※', '◎', '◈', '▣', '♦', '·', '-'];
      let pick = null;
      const grid = h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(10, 30px)', gap: '2px' } }, syms.map(s => { const c = h('div', { style: { height: '30px', border: '1px solid #ccc', display: 'grid', placeItems: 'center', cursor: 'pointer', fontSize: '18px' } }, s); c.onclick = () => { pick = s; [...grid.children].forEach(x => { x.style.background = x === c ? '#c7e0f4' : ''; }); }; return c; }));
      dialog({ title: '기호', width: 400, body: h('div', {}, h('div', { class: 'row' }, '글꼴(F):', h('select', {}, h('option', {}, 'Wingdings'), h('option', {}, '(현재 글꼴)'))), grid), buttons: [{ label: '확인', primary: true, onClick: () => pick && cb(pick) }, { label: '취소', cancel: true }] });
    }
    function togglePane(k) { pane = pane === k ? null : k; render(); }
    function openFmtPane(section) { pane = 'fmt'; app.fmtSection = section; render(); }

    // ------------------------------------------------------------------ 백스테이지
    function openBackstage(page) {
      closeBackstage();
      backstage = h('div', { class: 'backstage' });
      const nav = h('div', { class: 'bs-nav' });
      const mainEl = h('div', { class: 'bs-main' });
      const items = [['back', '←'], ['home', '홈'], ['new', '새로 만들기'], ['open', '열기'], ['-'], ['info', '정보'], ['save', '저장'], ['saveas', '다른 이름으로 저장'], ['print', '인쇄'], ['share', '공유'], ['export', '내보내기'], ['close', '닫기']];
      const show = p => {
        [...nav.children].forEach(b => b.classList.toggle('on', b.dataset.p === p));
        mainEl.innerHTML = '';
        if (p === 'new' || p === 'home') mainEl.appendChild(h('div', {}, h('h2', {}, '새로 만들기'), h('button', { class: 'btn', style: { width: '200px', height: '140px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: '8px' }, onclick: () => newPres() }, h('div', { style: { width: '160px', height: '90px', border: '1px solid #ccc', background: '#fff' } }), '새 프레젠테이션')));
        else if (p === 'print') mainEl.appendChild(printPage());
        else if (p === 'info') mainEl.appendChild(h('div', {}, h('h2', {}, '정보'), h('p', {}, pres ? `슬라이드 ${pres.slides.length}장, 크기 ${P.SIZES[pres.size.kind].label}` : '열린 프레젠테이션이 없습니다.')));
        else mainEl.appendChild(h('div', {}, h('h2', {}, items.find(x => x[0] === p)[1]), h('p', { class: 'note' }, '실기 과제에 쓰지 않는 화면입니다.')));
      };
      for (const [p, l] of items) {
        if (p === '-') { nav.appendChild(h('div', { style: { height: '1px', background: 'rgba(255,255,255,.3)', margin: '6px 14px' } })); continue; }
        const b = h('button', { class: p === 'back' ? 'back' : '', 'data-p': p, disabled: ['open', 'share', 'export'].includes(p) || (!pres && ['save', 'saveas', 'print', 'close', 'back', 'info'].includes(p)) }, l);
        b.addEventListener('click', () => { if (p === 'back' || p === 'close') return closeBackstage(); if (p === 'save') { closeBackstage(); return save(); } if (p === 'saveas') { closeBackstage(); return save(true); } show(p); });
        nav.appendChild(b);
      }
      backstage.append(nav, mainEl);
      win.appendChild(backstage);
      show(page);
    }
    function closeBackstage() { if (backstage) { backstage.remove(); backstage = null; render(); } }
    function newPres() {
      const go = () => { pres = P.newPresentation(); cur = 0; sel = []; ctx.setModel(pres); savedOnce = false; closeBackstage(); };
      if (pres) msg({ title: 'Microsoft PowerPoint', text: '현재 프레젠테이션을 닫고 새 프레젠테이션을 만듭니다.', buttons: ['새로 만들기', '취소'], icon: 'q' }).then(r => { if (r === '새로 만들기') go(); });
      else go();
    }
    function printPage() {
      const pr = pres.print;
      let pi = 0;
      const prev = h('div', { class: 'print-prev' });
      const setsEl = h('div', { class: 'print-set' });
      const setP = patch => exec(() => Object.assign(pres.print, patch));
      const pset = (ic, main2, sub, items) => { const el = h('div', { class: 'pset' }, icon(ic), h('div', { class: 'pt' }, main2, sub ? h('small', {}, sub) : null), h('span', { class: 'pdd' }, '▼')); el.addEventListener('click', () => menu(el, items())); return el; };
      const refresh = () => {
        setsEl.innerHTML = '';
        const layoutName = P.LAYOUT_NAMES[pr.layout];
        setsEl.append(
          h('div', { style: { display: 'flex', gap: '14px', alignItems: 'center', marginBottom: '10px' } }, h('button', { class: 'big', onclick: () => { ctx.printed('5-5', { pages: P.printPages(pres).length }); closeBackstage(); } }, icon('printer', 36), '인쇄'), h('div', {}, '복사본: ', h('input', { type: 'number', value: 1, min: 1, style: { width: '50px' } }))),
          h('div', { style: { fontWeight: 700, margin: '6px 0' } }, '프린터'), h('div', { class: 'pset' }, icon('printer'), h('div', { class: 'pt' }, '시험장 지정 프린터', h('small', {}, '준비'))),
          h('div', { style: { fontWeight: 700, margin: '10px 0 4px' } }, '설정'),
          pset('slidelayout', { all: '모든 슬라이드 인쇄', current: '현재 슬라이드 인쇄', range: '범위 지정' }[pr.target], { all: '전체 프레젠테이션을 인쇄합니다.', current: '현재 슬라이드만 인쇄합니다.', range: '인쇄할 슬라이드를 입력합니다.' }[pr.target], () => [{ label: '모든 슬라이드 인쇄', onClick: () => setP({ target: 'all' }) }, { label: '선택 영역 인쇄', disabled: true }, { label: '현재 슬라이드 인쇄', onClick: () => setP({ target: 'current' }) }, { label: '범위 지정', onClick: () => setP({ target: 'range' }) }]),
          h('div', { class: 'row' }, '슬라이드:', h('input', { type: 'text', value: pr.range || '', style: { width: '180px' }, onkeydown: e => e.stopPropagation(), onchange: e => setP({ range: e.target.value, target: e.target.value ? 'range' : 'all' }) })),
          pset('handout', pr.layout === 'full' ? '전체 페이지 슬라이드' : layoutName, pr.layout === 'full' ? '슬라이드당 1페이지 인쇄' : `유인물(${layoutName}/페이지)`, () => [{ header: '인쇄 모양' }, { label: '전체 페이지 슬라이드', checked: pr.layout === 'full', onClick: () => setP({ layout: 'full' }) }, { label: '슬라이드 노트', disabled: true }, { label: '개요', disabled: true }, { header: '유인물' },
            ...['h1', 'h2', 'h3', 'h4', 'h6', 'h9'].map(k => ({ label: P.LAYOUT_NAMES[k], checked: pr.layout === k, onClick: () => setP({ layout: k }) })), { sep: true },
            { label: '슬라이드 테두리(F)', checked: pr.frame, onClick: () => setP({ frame: !pr.frame }) }, { label: '용지에 맞게 크기 조정(S)', checked: pr.fit, onClick: () => setP({ fit: !pr.fit }) }, { label: '고품질', checked: pr.hq, onClick: () => setP({ hq: !pr.hq }) }, { label: '메모 및 잉크 인쇄', disabled: true }]),
          pset('copy', '단면 인쇄', '페이지의 한 면에만 인쇄', () => [{ label: '단면 인쇄', checked: true }]),
          pset('copy', '한 부씩 인쇄', '1,2,3  1,2,3  1,2,3', () => [{ label: '한 부씩 인쇄', checked: true }]),
          pset(pr.orient === 'portrait' ? 'paper' : 'orient', pr.orient === 'portrait' ? '세로 방향' : '가로 방향', '', () => [{ label: '세로 방향', checked: pr.orient === 'portrait', onClick: () => setP({ orient: 'portrait' }) }, { label: '가로 방향', checked: pr.orient === 'landscape', onClick: () => setP({ orient: 'landscape' }) }]),
          pset(pr.color === 'color' ? 'colors' : 'grayscale', { color: '컬러', gray: '회색조', bw: '흑백' }[pr.color], '', () => [{ label: '컬러', checked: pr.color === 'color', onClick: () => setP({ color: 'color' }) }, { label: '회색조', checked: pr.color === 'gray', onClick: () => setP({ color: 'gray' }) }, { label: '흑백', checked: pr.color === 'bw', onClick: () => setP({ color: 'bw' }) }]),
          h('div', { style: { textAlign: 'right' } }, h('button', { class: 'linkbtn', disabled: true }, '머리글 및 바닥글 편집')));
        prev.innerHTML = '';
        const pages = P.printPages(pres);
        pi = Math.max(0, Math.min(pi, pages.length - 1));
        const scale = Math.min(2.4, (prev.clientHeight || 720) / 320);
        prev.appendChild(handoutPage(pres, pages[pi], pi, scale));
        const bP = h('button', { disabled: pi === 0, onclick: () => { pi--; refresh(); } }, '◀'), bN = h('button', { disabled: pi >= pages.length - 1, onclick: () => { pi++; refresh(); } }, '▶');
        prev.appendChild(h('div', { class: 'print-nav' }, bP, ` ${pi + 1} / ${pages.length} `, bN));
      };
      app._refreshPrint = refresh;
      setTimeout(refresh, 0);
      return h('div', {}, h('h2', {}, '인쇄'), h('div', { class: 'print-wrap' }, setsEl, prev));
    }

    // ------------------------------------------------------------------ 화면 그리기
    let scale = 30;
    function render() {
      titleName.textContent = app.fileName();
      qU.disabled = !undo.length; qR.disabled = !redo.length;
      rib.render();
      if (backstage && app._refreshPrint && backstage.querySelector('.print-wrap')) app._refreshPrint();
      thumbs.innerHTML = ''; stage.innerHTML = ''; side.innerHTML = '';
      if (!pres) {
        stage.appendChild(h('div', { style: { padding: '40px', textAlign: 'center' } }, h('p', {}, '열린 프레젠테이션이 없습니다.'), h('button', { class: 'btn primary', onclick: () => openBackstage('new') }, '파일 탭 > 새로 만들기 > 새 프레젠테이션')));
        status.innerHTML = '';
        return;
      }
      cur = Math.max(0, Math.min(cur, pres.slides.length - 1));
      if (view === 'handout') { renderHandoutMaster(); return; }
      pres.slides.forEach((s, i) => {
        const tb = h('div', { class: 'tb', style: { width: '150px', height: 150 * pres.size.h / pres.size.w + 'px' }, html: slideSvg(pres, i, { print: true }) });
        if (view === 'gray') tb.classList.add('pp-gray');
        const t = h('div', { class: 'pp-thumb' + (i === cur ? ' on' : '') }, h('span', { class: 'n' }, String(i + 1)), tb);
        t.addEventListener('click', () => { cur = i; sel = []; child = null; render(); });
        t.addEventListener('contextmenu', e => { e.preventDefault(); menu(t, [{ label: '새 슬라이드(N)', onClick: () => { cur = i; newSlide(); } }, { label: '슬라이드 삭제(D)', disabled: pres.slides.length < 2, onClick: () => exec(() => { P.deleteSlide(pres, i); cur = Math.min(cur, pres.slides.length - 1); }) }, { label: '레이아웃(L)', sub: layoutMenu(k => exec(() => P.setLayout(pres, i, k))) }], { at: { x: e.clientX, y: e.clientY } }); });
        thumbs.appendChild(t);
      });
      const sw = stage.clientWidth || 900, sh2 = stage.clientHeight || 600;
      scale = Math.max(10, Math.min((sw - 60) / pres.size.w, (sh2 - 50) / pres.size.h));
      const W = pres.size.w * scale, H = pres.size.h * scale;
      const wrap = h('div', { class: 'pp-slidewrap', style: { width: W + 'px', height: H + 'px', left: Math.max(20, (sw - W) / 2) + 'px', top: Math.max(16, (sh2 - H) / 2) + 'px' } });
      wrap.innerHTML = slideSvg(pres, cur, {});
      const svg = wrap.firstChild;
      svg.classList.add('pp-svg');
      if (view === 'gray') svg.style.filter = 'grayscale(1)';
      const over = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      svg.appendChild(over);
      drawOverlay(over);
      stage.appendChild(wrap);
      app._wrap = wrap; app._svg = svg; app._over = over;
      if (pane === 'sel') side.appendChild(selPane());
      if (pane === 'fmt' && hasSel()) side.appendChild(fmtPane());
      status.innerHTML = '';
      status.append(h('span', {}, `슬라이드 ${cur + 1}/${pres.slides.length}`), h('span', {}, '한국어'), h('span', { class: 'sp' }), tool ? h('span', {}, '그리기: ' + (P.KINDS[tool] ? P.KINDS[tool].name : tool) + (lock ? ' (잠금)' : '')) : null, h('span', {}, Math.round(scale / 37.8 * 100) + '%'));
      stage.style.cursor = tool ? 'crosshair' : 'default';
    }
    function drawOverlay(g) {
      const sl = slide();
      let o = '';
      const W = pres.size.w * UU, H = pres.size.h * UU;
      if (grid) for (let x = 0; x <= pres.size.w; x += 1) o += `<line class="gridl" x1="${x * UU}" y1="0" x2="${x * UU}" y2="${H}"/>`;
      if (grid) for (let y = 0; y <= pres.size.h; y += 1) o += `<line class="gridl" x1="0" y1="${y * UU}" x2="${W}" y2="${y * UU}"/>`;
      if (guides) o += `<line class="guide" x1="${W / 2}" y1="0" x2="${W / 2}" y2="${H}"/><line class="guide" x1="0" y1="${H / 2}" x2="${W}" y2="${H / 2}"/>`;
      const hs = 6 * UU / scale;
      for (const s of target()) {
        if (P.isLine(s)) {
          const pts = P.linePoints(pres, sl, s);
          const a = pts[0], b = pts[pts.length - 1];
          const ends = P.endpoints(pres, sl, s);
          o += `<circle class="${ends[0].attached ? 'green' : 'hnd'}" data-end="begin" cx="${a.x * UU}" cy="${a.y * UU}" r="${hs}"/><circle class="${ends[1].attached ? 'green' : 'hnd'}" data-end="end" cx="${b.x * UU}" cy="${b.y * UU}" r="${hs}"/>`;
          if (s.kind === 'connElbow' && pts.length >= 4) { const m1 = pts[1], m2 = pts[2]; o += `<circle class="yellow" data-adj="elbow" cx="${(m1.x + m2.x) / 2 * UU}" cy="${(m1.y + m2.y) / 2 * UU}" r="${hs}"/>`; }
          continue;
        }
        const b = P.bbox(s);
        o += `<rect class="selbox" x="${b.x * UU}" y="${b.y * UU}" width="${b.w * UU}" height="${b.h * UU}"/>`;
        if (target().length === 1 || true) {
          for (const [px, py] of [[0, 0], [0.5, 0], [1, 0], [1, 0.5], [1, 1], [0.5, 1], [0, 1], [0, 0.5]]) o += `<circle class="hnd" data-h="${px},${py}" cx="${(b.x + b.w * px) * UU}" cy="${(b.y + b.h * py) * UU}" r="${hs}"/>`;
          if (s.kind !== 'group') o += `<circle class="rot" cx="${(b.x + b.w / 2) * UU}" cy="${(b.y - 0.6) * UU}" r="${hs}" style="pointer-events:none"/>`;
        }
        if (['calloutRect', 'calloutRoundRect', 'calloutOval'].includes(s.kind)) { const t = P.tailTip(s); o += `<circle class="yellow" data-adj="tail" cx="${t.x * UU}" cy="${t.y * UU}" r="${hs}"/>`; }
      }
      if (app._sitesFor) { const t = P.findShape(pres, cur, app._sitesFor); if (t) for (const q of P.sites(t)) o += `<circle class="site" cx="${q.x * UU}" cy="${q.y * UU}" r="${hs * 0.8}"/>`; }
      if (app._marquee) { const m = app._marquee; o += `<rect class="marquee" x="${Math.min(m.x0, m.x1) * UU}" y="${Math.min(m.y0, m.y1) * UU}" width="${Math.abs(m.x1 - m.x0) * UU}" height="${Math.abs(m.y1 - m.y0) * UU}"/>`; }
      g.innerHTML = o;
    }
    function toCm(e) { const r = app._wrap.getBoundingClientRect(); return { x: (e.clientX - r.left) / scale, y: (e.clientY - r.top) / scale }; }
    function shapeAt(p) {
      const el = document.elementFromPoint(p.cx, p.cy);
      const hit = el && el.closest && el.closest('[data-id]');
      return hit ? { id: hit.getAttribute('data-id'), child: hit.getAttribute('data-child') } : null;
    }
    function nearestSite(p, exceptId) {
      let best = null, bd = 0.5;
      P.eachShape(slide().shapes, s => {
        if (s.id === exceptId || s.kind === 'group' || P.isLine(s) || s.hidden) return;
        const inGroup = !slide().shapes.includes(s);
        if (inGroup) return;
        const b = P.bbox(s);
        if (p.x < b.x - 0.6 || p.x > b.x + b.w + 0.6 || p.y < b.y - 0.6 || p.y > b.y + b.h + 0.6) return;
        for (const q of P.sites(s)) { const dd = Math.hypot(q.x - p.x, q.y - p.y); if (dd < bd) { bd = dd; best = { id: s.id, site: q.site, x: q.x, y: q.y }; } }
      });
      return best;
    }
    function hoverShape(p) { let hit = null; P.eachShape(slide().shapes, s => { if (s.kind === 'group' || P.isLine(s) || !slide().shapes.includes(s)) return; const b = P.bbox(s); if (p.x >= b.x - 0.3 && p.x <= b.x + b.w + 0.3 && p.y >= b.y - 0.3 && p.y <= b.y + b.h + 0.3) hit = s; }); return hit; }

    stage.addEventListener('mousedown', e => {
      if (!pres || view === 'handout' || e.button !== 0 || !app._wrap) return;
      if (app._editing) return;
      stage.focus({ preventScroll: true });
      const p = toCm(e);
      const t = e.target;
      // 그리기
      if (tool) { e.preventDefault(); drawShape(e, p); return; }
      // 핸들
      if (t.getAttribute && t.getAttribute('data-h')) { e.preventDefault(); return resizeShape(e, t.getAttribute('data-h').split(',').map(Number)); }
      if (t.getAttribute && t.getAttribute('data-end')) { e.preventDefault(); return moveEnd(e, t.getAttribute('data-end')); }
      if (t.getAttribute && t.getAttribute('data-adj')) { e.preventDefault(); return moveAdj(e, t.getAttribute('data-adj')); }
      const hit = t.closest && t.closest('[data-id]');
      if (!hit) {
        sel = []; child = null;
        app._marquee = { x0: p.x, y0: p.y, x1: p.x, y1: p.y };
        OA.ui.drag(e, ev => { const q = toCm(ev); app._marquee.x1 = q.x; app._marquee.y1 = q.y; drawOverlay(app._over); }, () => {
          const m = app._marquee; app._marquee = null;
          const x1 = Math.min(m.x0, m.x1), x2 = Math.max(m.x0, m.x1), y1 = Math.min(m.y0, m.y1), y2 = Math.max(m.y0, m.y1);
          if (x2 - x1 > 0.1 || y2 - y1 > 0.1) sel = slide().shapes.filter(s => { const b = P.bbox(s); return !s.hidden && b.x >= x1 && b.y >= y1 && b.x + b.w <= x2 && b.y + b.h <= y2; }).map(s => s.id);
          render();
        });
        render();
        return;
      }
      e.preventDefault();
      const id = hit.getAttribute('data-id'), cid = hit.getAttribute('data-child');
      const topS = slide().shapes.find(s => s.id === id);
      if (e.shiftKey || (e.ctrlKey && !sel.includes(id))) {
        if (e.shiftKey) { if (sel.includes(id)) sel = sel.filter(x => x !== id); else sel.push(id); child = null; render(); return; }
      }
      if (topS && topS.kind === 'group' && sel.length === 1 && sel[0] === id && cid && cid !== id) child = cid;
      else if (!sel.includes(id)) { sel = [id]; child = null; }
      render();
      dragMove(e, p);
    });
    stage.addEventListener('dblclick', e => {
      const hit = e.target.closest && e.target.closest('[data-id]');
      if (!hit || tool) return;
      const id = hit.getAttribute('data-child') || hit.getAttribute('data-id');
      const s = P.findShape(pres, cur, id);
      if (s && s.text) editText(s);
    });
    stage.addEventListener('contextmenu', e => {
      if (!pres || !app._wrap) return;
      e.preventDefault();
      const hit = e.target.closest && e.target.closest('[data-id]');
      if (hit) { const id = hit.getAttribute('data-id'); if (!sel.includes(id)) { sel = [id]; child = null; render(); } }
      const s = target()[0];
      const items = s ? [{ label: '잘라내기(T)', icon: 'cut', onClick: () => { copyShapes(); deleteSel(); } }, { label: '복사(C)', icon: 'copy', onClick: copyShapes }, { label: '붙여넣기 옵션:', icon: 'paste', disabled: !shapeClip, onClick: pasteShapes }, { sep: true },
        { label: '텍스트 편집(X)', disabled: !s.text, onClick: () => editText(s) }, { label: '점 편집(E)', disabled: true }, { label: '그룹화(G)', sub: groupMenu() }, { label: '맨 앞으로 가져오기(R)', sub: orderMenu(true) }, { label: '맨 뒤로 보내기(K)', sub: orderMenu(false) }, { sep: true },
        { label: P.isLine(s) ? '기본 선으로 설정(D)' : s.kind === 'textbox' ? '기본 텍스트 상자로 설정(D)' : '기본 도형으로 설정(D)', disabled: s.kind === 'group' || s.kind === 'placeholder', onClick: () => { exec(() => P.setDefault(pres, s)); toast('이후에 그리는 ' + (P.isLine(s) ? '선' : s.kind === 'textbox' ? '텍스트 상자' : '도형') + '은 이 서식으로 만들어집니다.'); } },
        { label: '크기 및 위치(Z)...', onClick: () => openFmtPane('size') }, { label: '도형 서식(O)...', onClick: () => openFmtPane('fill') }]
        : [{ label: '붙여넣기 옵션:', icon: 'paste', disabled: !shapeClip, onClick: pasteShapes }, { sep: true }, { label: '레이아웃(L)', sub: layoutMenu(k => exec(() => P.setLayout(pres, cur, k))) }, { label: '눈금 및 안내선(I)', sub: [{ label: '눈금선', checked: grid, onClick: () => { grid = !grid; render(); } }, { label: '안내선', checked: guides, onClick: () => { guides = !guides; render(); } }] }];
      menu(stage, items, { at: { x: e.clientX, y: e.clientY } });
    });
    stage.addEventListener('mousemove', e => {
      if (!tool || !['line', 'connStraight', 'connElbow', 'arrow', 'arrow2', 'elbowArrow', 'elbowArrow2'].includes(tool) || !app._wrap) return;
      const s = hoverShape(toCm(e));
      const id = s ? s.id : null;
      if (id !== app._sitesFor) { app._sitesFor = id; drawOverlay(app._over); }
    });
    function dragMove(e, p0) {
      const ctrl = e.ctrlKey, shift = e.shiftKey;
      const moving = child ? [P.findShape(pres, cur, child)] : selShapes();
      if (!moving.length || moving.some(s => s.kind === 'placeholder' && false)) return;
      let moved = false, snap = null, dupDone = false;
      const before = JSON.stringify(pres);
      OA.ui.drag(e, ev => {
        const q = toCm(ev);
        let dx = q.x - p0.x, dy = q.y - p0.y;
        if (!moved && Math.hypot(dx, dy) * scale < 4) return;
        if (ev.shiftKey || shift) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
        if (!moved) { moved = true; snap = before; if (ctrl || ev.ctrlKey) { const made = P.duplicate(pres, cur, moving, 0, 0); sel = made.map(s => s.id); child = null; moving.length = 0; moving.push(...made); dupDone = true; } moving.forEach(s => { s._o = JSON.stringify(P.bbox(s)); s._orig = U.clone(s); }); }
        moving.forEach(s => { const o = s._orig; const tmp = U.clone(o); P.moveShapes(pres, slide(), [tmp], dx, dy); Object.assign(s, tmp); });
        app._wrap.innerHTML = slideSvg(pres, cur, {});
        app._svg = app._wrap.firstChild; app._svg.classList.add('pp-svg');
        const g = document.createElementNS('http://www.w3.org/2000/svg', 'g'); app._svg.appendChild(g); app._over = g; drawOverlay(g);
      }, () => {
        moving.forEach(s => { delete s._o; delete s._orig; });
        if (moved) { undo.push(snap); redo.length = 0; ctx.changed(); }
        void dupDone;
        render();
      });
    }
    function resizeShape(e, [px, py]) {
      const s = selShapes()[0];
      if (!s || P.isLine(s)) return;
      const before = JSON.stringify(pres);
      const b0 = P.bbox(s);
      const snapshot = U.clone(s);
      OA.ui.drag(e, ev => {
        const q = toCm(ev);
        let x1 = b0.x, y1 = b0.y, x2 = b0.x + b0.w, y2 = b0.y + b0.h;
        if (px === 0) x1 = q.x; if (px === 1) x2 = q.x; if (py === 0) y1 = q.y; if (py === 1) y2 = q.y;
        if (ev.shiftKey && px !== 0.5 && py !== 0.5) { const r = b0.w / b0.h; const w = Math.abs(x2 - x1), hh = w / r; if (py === 0) y1 = y2 - hh; else y2 = y1 + hh; }
        const nb = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(0.05, Math.abs(x2 - x1)), h: Math.max(0.05, Math.abs(y2 - y1)) };
        const fx = nb.w / Math.max(0.0001, b0.w), fy = nb.h / Math.max(0.0001, b0.h);
        const apply = (src, dst) => { if (src.kind === 'group') { src.children.forEach((c, i) => apply(c, dst.children[i])); return; } if (P.isLine(src)) { dst.x1 = nb.x + (src.x1 - b0.x) * fx; dst.y1 = nb.y + (src.y1 - b0.y) * fy; dst.x2 = nb.x + (src.x2 - b0.x) * fx; dst.y2 = nb.y + (src.y2 - b0.y) * fy; return; } dst.x = nb.x + (src.x - b0.x) * fx; dst.y = nb.y + (src.y - b0.y) * fy; dst.w = src.w * fx; dst.h = src.h * fy; };
        apply(snapshot, s);
        app._wrap.innerHTML = slideSvg(pres, cur, {}); app._svg = app._wrap.firstChild; app._svg.classList.add('pp-svg');
        const g = document.createElementNS('http://www.w3.org/2000/svg', 'g'); app._svg.appendChild(g); app._over = g; drawOverlay(g);
      }, () => { undo.push(before); redo.length = 0; ctx.changed(); render(); });
    }
    function moveEnd(e, which) {
      const s = target()[0];
      const before = JSON.stringify(pres);
      OA.ui.drag(e, ev => {
        const q = toCm(ev);
        const site = nearestSite(q, null);
        const hover = hoverShape(q);
        app._sitesFor = hover ? hover.id : null;
        if (which === 'begin') { if (site) { s.begin = { id: site.id, site: site.site }; } else { s.begin = null; s.x1 = q.x; s.y1 = q.y; } }
        else { if (site) { s.end = { id: site.id, site: site.site }; } else { s.end = null; s.x2 = q.x; s.y2 = q.y; } }
        renderSlideOnly();
      }, () => { app._sitesFor = null; undo.push(before); redo.length = 0; ctx.changed(); render(); });
    }
    function moveAdj(e, kind) {
      const s = target()[0];
      const before = JSON.stringify(pres);
      OA.ui.drag(e, ev => {
        const q = toCm(ev);
        if (kind === 'tail') s.adj = [(q.x - (s.x + s.w / 2)) / s.w, (q.y - (s.y + s.h / 2)) / s.h];
        if (kind === 'elbow') { const [a, b] = P.endpoints(pres, slide(), s); const horiz = P.elbowPoints(a, b, 0.5)[1].y === a.y; s.adj = horiz ? (q.x - a.x) / ((b.x - a.x) || 1) : (q.y - a.y) / ((b.y - a.y) || 1); }
        renderSlideOnly();
      }, () => { undo.push(before); redo.length = 0; ctx.changed(); render(); });
    }
    function renderSlideOnly() {
      app._wrap.innerHTML = slideSvg(pres, cur, {}); app._svg = app._wrap.firstChild; app._svg.classList.add('pp-svg');
      if (view === 'gray') app._svg.style.filter = 'grayscale(1)';
      const g = document.createElementNS('http://www.w3.org/2000/svg', 'g'); app._svg.appendChild(g); app._over = g; drawOverlay(g);
    }
    function drawShape(e, p0) {
      const k0 = tool;
      const isLineTool = ['line', 'connStraight', 'connElbow', 'arrow', 'arrow2', 'elbowArrow', 'elbowArrow2'].includes(k0);
      const kind = { arrow: 'connStraight', arrow2: 'connStraight', line: 'connStraight', elbowArrow: 'connElbow', elbowArrow2: 'connElbow' }[k0] || k0;
      const arrow = { arrow: 'end', arrow2: 'both', elbowArrow: 'end', elbowArrow2: 'both' }[k0] || null;
      const before = JSON.stringify(pres);
      let s;
      const startSite = isLineTool ? nearestSite(p0, null) : null;
      if (isLineTool) {
        s = P.addShape(pres, cur, kind, { x1: p0.x, y1: p0.y, x2: p0.x, y2: p0.y }, { arrow });
        if (startSite) s.begin = { id: startSite.id, site: startSite.site };
      } else s = P.addShape(pres, cur, kind, { x: p0.x, y: p0.y, w: 0.01, h: 0.01 });
      let moved = false;
      OA.ui.drag(e, ev => {
        const q = toCm(ev);
        moved = moved || Math.hypot(q.x - p0.x, q.y - p0.y) * scale > 4;
        if (isLineTool) {
          let ex = q.x, ey = q.y;
          if (ev.shiftKey) { if (Math.abs(ex - p0.x) > Math.abs(ey - p0.y)) ey = p0.y; else ex = p0.x; }
          const site = nearestSite({ x: ex, y: ey }, null);
          const hv = hoverShape({ x: ex, y: ey });
          app._sitesFor = hv ? hv.id : null;
          if (site && !(startSite && site.id === startSite.id)) { s.end = { id: site.id, site: site.site }; } else { s.end = null; s.x2 = ex; s.y2 = ey; }
        } else {
          let w = q.x - p0.x, hh = q.y - p0.y;
          if (ev.shiftKey) { const m = Math.max(Math.abs(w), Math.abs(hh)); w = Math.sign(w || 1) * m; hh = Math.sign(hh || 1) * m; }
          s.x = Math.min(p0.x, p0.x + w); s.y = Math.min(p0.y, p0.y + hh); s.w = Math.max(0.05, Math.abs(w)); s.h = Math.max(0.05, Math.abs(hh));
        }
        renderSlideOnly();
      }, () => {
        app._sitesFor = null;
        if (!moved) {
          if (isLineTool) { s.x2 = p0.x + 2.54; s.y2 = p0.y; }
          else if (kind === 'textbox') { s.w = 4; s.h = 0.9; }
          else { s.w = 2.54; s.h = 2.54; }
        }
        if (isLineTool && s.begin && s.end && s.begin.id === s.end.id) s.end = null;
        undo.push(before); redo.length = 0; ctx.changed();
        sel = [s.id]; child = null;
        if (!lock) tool = null;
        render();
        if (kind === 'textbox') editText(s);
      });
    }
    function editText(s, initial) {
      if (app._editing) return;
      const b = P.bbox(s);
      const r = app._wrap.getBoundingClientRect(), sr = stage.getBoundingClientRect();
      const ta = h('textarea', { class: 'pp-text', spellcheck: false }, initial !== undefined ? initial : (s.text.value || ''));
      Object.assign(ta.style, { left: r.left - sr.left + stage.scrollLeft + b.x * scale + 'px', top: r.top - sr.top + stage.scrollTop + b.y * scale + 'px', width: Math.max(60, b.w * scale) + 'px', height: Math.max(30, b.h * scale) + 'px', fontSize: (s.text.size || 18) * PTCM * scale + 'px', fontWeight: s.text.bold ? 700 : 400, textAlign: s.text.align || 'left' });
      stage.appendChild(ta);
      app._editing = true;
      ta.focus();
      ta.setSelectionRange(ta.value.length, ta.value.length);
      const done = ok => {
        if (!app._editing) return;
        app._editing = false;
        const v = ta.value;
        ta.remove();
        if (ok && v !== (s.text.value || '')) exec(() => { const t = P.findShape(pres, cur, s.id); t.text.value = v; if (t.kind === 'placeholder' && v) { /* 개체 틀은 입력하면 내용이 있는 개체 틀 */ } if (t.kind === 'textbox' && t.autoFit !== false) { const lines = v.split('\n').length; t.h = Math.max(t.h, lines * (t.text.size || 18) * PTCM * 1.25 + 0.26); } });
        else render();
        stage.focus({ preventScroll: true });
      };
      ta.addEventListener('keydown', ev => { ev.stopPropagation(); if (ev.key === 'Escape') { ev.preventDefault(); done(true); } });
      ta.addEventListener('blur', () => done(true));
      ta.addEventListener('mousedown', ev => ev.stopPropagation());
    }

    // ------------------------------------------------------------------ 작업창
    function selPane() {
      const sl = slide();
      const list = h('div', { class: 'selpane' });
      const row = (s, depth) => {
        const on = sel.includes(s.id) || child === s.id;
        const nm = h('span', { class: 'nm' }, s.name);
        nm.addEventListener('dblclick', () => { const inp = h('input', { value: s.name, style: { width: '100%' } }); nm.innerHTML = ''; nm.appendChild(inp); inp.focus(); inp.select(); const fin = () => { const v = inp.value.trim(); if (v && v !== s.name) exec(() => { P.findShape(pres, cur, s.id).name = v; }); else render(); }; inp.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') fin(); if (e.key === 'Escape') render(); }); inp.addEventListener('blur', fin); });
        const eye = h('button', { class: 'eye', title: s.hidden ? '표시' : '숨기기', onclick: ev => { ev.stopPropagation(); exec(() => { const t = P.findShape(pres, cur, s.id); t.hidden = !t.hidden; }); } }, s.hidden ? '—' : '👁');
        const r = h('div', { class: 'row' + (on ? ' on' : '') + (depth ? ' child' : ''), style: { paddingLeft: 4 + depth * 14 + 'px' } }, nm, eye);
        r.addEventListener('click', ev => { if (depth) { child = s.id; } else if (ev.ctrlKey || ev.shiftKey) { if (sel.includes(s.id)) sel = sel.filter(x => x !== s.id); else sel.push(s.id); child = null; } else { sel = [s.id]; child = null; } render(); });
        list.appendChild(r);
        if (s.kind === 'group') [...s.children].reverse().forEach(c => row(c, depth + 1));
      };
      [...sl.shapes].reverse().forEach(s => row(s, 0));
      return h('div', { class: 'pane' }, h('div', { class: 'pane-h' }, h('span', { class: 't' }, '선택'), h('button', { title: '닫기', onclick: () => togglePane('sel') }, '✕')),
        h('div', { class: 'pane-b' }, h('div', { class: 'top' }, h('button', { class: 'btn', onclick: () => exec(() => P.eachShape(slide().shapes, s => { s.hidden = false; })) }, '모두 표시'), h('button', { class: 'btn', onclick: () => exec(() => P.eachShape(slide().shapes, s => { s.hidden = true; })) }, '모두 숨기기'),
          h('span', { style: { flex: 1 } }), h('button', { class: 'btn', title: '앞으로 가져오기', onclick: () => exec(() => P.zorder(pres, cur, sel, 'forward')) }, '▲'), h('button', { class: 'btn', title: '뒤로 보내기', onclick: () => exec(() => P.zorder(pres, cur, sel, 'backward')) }, '▼')),
          h('div', { class: 'note', style: { margin: '0 0 4px' } }, `이 슬라이드의 도형 ${sl.shapes.length}개(그룹은 1개). 위쪽이 앞에 있는 도형입니다.`), list));
    }
    function fmtPane() {
      const ts = flatTargets();
      const s = ts[0];
      const sec = app.fmtSection || 'fill';
      const body = h('div', { class: 'fmtpane' });
      const num = (v, step, on) => { const i = h('input', { type: 'number', step: step || 0.1, value: v, style: { width: '80px' } }); i.addEventListener('keydown', e => e.stopPropagation()); i.addEventListener('change', () => on(parseFloat(i.value))); return i; };
      const r = (label, el) => h('div', { class: 'r' }, h('span', {}, label), el);
      if (sec === 'fill') {
        const fillNone = h('input', { type: 'radio', name: 'ff', checked: !s.fill, onchange: () => exec(() => ts.forEach(x => { if (!P.isLine(x)) x.fill = null; })) });
        const fillSolid = h('input', { type: 'radio', name: 'ff', checked: !!s.fill, onchange: () => exec(() => ts.forEach(x => { if (!P.isLine(x)) x.fill = x.fill || '#FFFFFF'; })) });
        const lineNone = h('input', { type: 'radio', name: 'll', checked: !s.line, onchange: () => exec(() => ts.forEach(x => { x.line = null; })) });
        const lineSolid = h('input', { type: 'radio', name: 'll', checked: !!s.line, onchange: () => exec(() => ts.forEach(x => { x.line = x.line || { color: '#000000', weight: 0.75, dash: 'solid' }; })) });
        const colorBtn = (cur2, on) => h('button', { class: 'btn', style: { minWidth: '60px' }, onclick: ev => menu(ev.target, [{ html: OA.ui.palette(on, close => h('div', { class: 'mi', onclick: () => { close(); OA.ui.colorDialog(cur2, on); } }, '다른 색...')) }]) }, h('span', { style: { display: 'inline-block', width: '16px', height: '12px', background: cur2 || '#fff', border: '1px solid #888', verticalAlign: 'middle' } }), ' ▼');
        body.append(h('div', { class: 'sec' }, h('b', {}, '▾ 채우기'), h('label', { class: 'r' }, fillNone, '채우기 없음(N)'), h('label', { class: 'r' }, fillSolid, '단색 채우기(S)'),
          s.fill ? r('색(C)', colorBtn(s.fill, c => exec(() => ts.forEach(x => { if (!P.isLine(x)) x.fill = c; })))) : null, s.fill ? r('투명도(T)', num(0, 1, () => {})) : null),
          h('div', { class: 'sec' }, h('b', {}, '▾ 선'), h('label', { class: 'r' }, lineNone, '선 없음(N)'), h('label', { class: 'r' }, lineSolid, '실선(S)'),
            s.line ? r('색(C)', colorBtn(s.line.color, c => exec(() => ts.forEach(x => { x.line = Object.assign({}, x.line, { color: c }); })))) : null,
            s.line ? r('너비(W)', num(s.line.weight, 0.25, v => exec(() => ts.forEach(x => { if (x.line) x.line.weight = v; })))) : null,
            s.line ? r('대시 종류(D)', (() => { const sl = h('select', { onchange: () => exec(() => ts.forEach(x => { if (x.line) x.line.dash = sl.value; })) }, DASH_NAMES.map(([k, l]) => h('option', { value: k, selected: s.line.dash === k }, l))); return sl; })()) : null,
            s.line && P.isLine(s) ? r('시작 화살표 유형', (() => { const sl = h('select', { onchange: () => exec(() => ts.forEach(x => { if (x.line) x.line.begin = sl.value || null; })) }, h('option', { value: '' }, '없음'), h('option', { value: 'arrow', selected: s.line.begin === 'arrow' }, '화살표')); return sl; })()) : null,
            s.line && P.isLine(s) ? r('끝 화살표 유형', (() => { const sl = h('select', { onchange: () => exec(() => ts.forEach(x => { if (x.line) x.line.end = sl.value || null; })) }, h('option', { value: '' }, '없음'), h('option', { value: 'arrow', selected: s.line.end === 'arrow' }, '화살표')); return sl; })()) : null));
      }
      if (sec === 'effect') {
        const sh = s.shadow || {};
        const set = patch => exec(() => ts.forEach(x => { x.shadow = Object.assign({ color: '#000000', transparency: 0.6, blur: 4, angle: 45, dist: 3 }, x.shadow, patch); }));
        body.append(h('div', { class: 'sec' }, h('b', {}, '▾ 그림자'),
          r('미리 설정(P)', h('select', { onchange: e => { if (e.target.value === 'none') exec(() => ts.forEach(x => { x.shadow = null; })); else set({}); } }, h('option', { value: 'none', selected: !s.shadow }, '그림자 없음'), h('option', { value: 'outer', selected: !!s.shadow }, '바깥쪽: 오프셋 오른쪽 아래'))),
          r('색(C)', h('button', { class: 'btn', onclick: ev => menu(ev.target, [{ html: OA.ui.palette(c => set({ color: c })) }]) }, h('span', { style: { display: 'inline-block', width: '16px', height: '12px', background: sh.color || '#000', border: '1px solid #888' } }), ' ▼')),
          r('투명도(T) %', num(Math.round((sh.transparency === undefined ? 0.6 : sh.transparency) * 100), 1, v => set({ transparency: Math.max(0, Math.min(100, v)) / 100 }))),
          r('크기(S) %', num(100, 1, () => {})),
          r('흐리게(B) pt', num(sh.blur === undefined ? 4 : sh.blur, 0.5, v => set({ blur: Math.max(0, v) }))),
          r('각도(A) °', num(sh.angle === undefined ? 45 : sh.angle, 1, v => set({ angle: v }))),
          r('간격(D) pt', num(sh.dist === undefined ? 3 : sh.dist, 0.1, v => set({ dist: Math.max(0, v) })))),
          h('div', { class: 'sec' }, h('b', {}, '▸ 반사'), h('b', {}, '▸ 네온'), h('b', {}, '▸ 부드러운 가장자리'), h('b', {}, '▸ 3차원 서식'), h('b', {}, '▸ 3차원 회전')));
      }
      if (sec === 'size') {
        const t0 = target()[0];
        const b = P.bbox(t0);
        body.append(h('div', { class: 'sec' }, h('b', {}, '▾ 크기'), r('높이(E) cm', num(+b.h.toFixed(2), 0.1, v => resizeTo(null, v))), r('너비(D) cm', num(+b.w.toFixed(2), 0.1, v => resizeTo(v, null)))),
          h('div', { class: 'sec' }, h('b', {}, '▾ 위치'), r('가로 위치(H) cm', num(+b.x.toFixed(2), 0.1, v => exec(() => P.moveShapes(pres, slide(), target(), v - b.x, 0)))), r('세로 위치(V) cm', num(+b.y.toFixed(2), 0.1, v => exec(() => P.moveShapes(pres, slide(), target(), 0, v - b.y))))),
          s.text ? h('div', { class: 'sec' }, h('b', {}, '▾ 텍스트 상자'),
            r('세로 맞춤(V)', (() => { const sl = h('select', { onchange: () => setText({ anchor: sl.value }) }, [['top', '위쪽'], ['middle', '중간'], ['bottom', '아래쪽']].map(([k, l]) => h('option', { value: k, selected: s.text.anchor === k }, l))); return sl; })()),
            ...[['왼쪽 여백(L)', 0], ['오른쪽 여백(R)', 2], ['위쪽 여백(T)', 1], ['아래쪽 여백(B)', 3]].map(([l, k]) => r(l + ' cm', num(+(s.margins || [0.25, 0.13, 0.25, 0.13])[k].toFixed(2), 0.05, v => exec(() => ts.forEach(x => { x.margins = (x.margins || [0.25, 0.13, 0.25, 0.13]).slice(); x.margins[k] = Math.max(0, v); }))))),
            h('label', { class: 'r' }, h('input', { type: 'checkbox', checked: true, disabled: true }), '도형의 텍스트 배치(W)')) : null);
      }
      const tabBtn = (k, ic, label) => h('button', { class: 'btn' + (sec === k ? ' primary' : ''), title: label, onclick: () => { app.fmtSection = k; render(); } }, icon(ic), ' ', label);
      return h('div', { class: 'pane', style: { width: '320px' } }, h('div', { class: 'pane-h' }, h('span', { class: 't' }, '도형 서식'), h('button', { title: '닫기', onclick: () => { pane = null; render(); } }, '✕')),
        h('div', { class: 'pane-b' }, h('div', { class: 'row', style: { display: 'flex', gap: '6px', marginBottom: '6px' } }, h('b', {}, '도형 옵션'), h('span', { class: 'note' }, '텍스트 옵션')),
          h('div', { style: { display: 'flex', gap: '4px', marginBottom: '8px', flexWrap: 'wrap' } }, tabBtn('fill', 'shapeFill', '채우기 및 선'), tabBtn('effect', 'shapeEffect', '효과'), tabBtn('size', 'slidesize', '크기 및 속성')), body));
    }

    // ------------------------------------------------------------------ 유인물 마스터
    function renderHandoutMaster() {
      const sw = stage.clientWidth || 900, sh2 = stage.clientHeight || 600;
      const portrait = pres.handout.orient !== 'landscape';
      const PW = portrait ? 21 : 29.7, PH = portrait ? 29.7 : 21;
      const s = Math.min((sw - 60) / PW, (sh2 - 40) / PH);
      const pg = h('div', { class: 'hm-page', style: { width: PW * s + 'px', height: PH * s + 'px', left: (sw - PW * s) / 2 + 'px', top: '20px' } });
      const ph = (on, label, x, y) => on ? pg.appendChild(h('div', { class: 'hm-ph', style: { left: x * s + 'px', top: y * s + 'px', width: 9 * s + 'px', height: 1.4 * s + 'px' } }, label)) : null;
      ph(pres.handout.header, '머리글', 0.5, 0.5); ph(pres.handout.date, '2026-11-11', PW - 9.5, 0.5); ph(pres.handout.footer, '바닥글', 0.5, PH - 1.9); ph(pres.handout.pageNum, '<#>', PW - 9.5, PH - 1.9);
      const lay = P.HANDOUT['h' + pres.handout.perPage] || P.HANDOUT.h6;
      for (const [x, y, w, hh] of lay.slots) pg.appendChild(h('div', { class: 'hm-slot', style: { left: x * s + 'px', top: y * s + 'px', width: w * s + 'px', height: hh * s + 'px' } }));
      stage.appendChild(pg);
      status.innerHTML = '';
      status.append(h('span', {}, '유인물 마스터'), h('span', { class: 'sp' }));
    }

    // ------------------------------------------------------------------ 키보드(표 9-4)
    this.handleKey = e => {
      if (backstage) { if (e.key === 'Escape') closeBackstage(); return; }
      if (app._editing) return;
      const t = e.target;
      if (t && /INPUT|TEXTAREA|SELECT/.test(t.tagName)) return;
      const k = e.key, ctrl = e.ctrlKey || e.metaKey, shift = e.shiftKey, alt = e.altKey, code = e.code;
      if (ctrl && code === 'KeyS') { e.preventDefault(); save(); return; }
      if (!pres) return;
      if (ctrl && code === 'KeyP') { e.preventDefault(); openBackstage('print'); return; }
      if (ctrl && code === 'KeyZ') { e.preventDefault(); doUndo(); return; }
      if (ctrl && code === 'KeyY') { e.preventDefault(); doRedo(); return; }
      if (ctrl && code === 'KeyM') { e.preventDefault(); newSlide(); return; }
      if (alt && k === 'F10') { e.preventDefault(); togglePane('sel'); return; }
      if (alt && k === 'F9') { e.preventDefault(); guides = !guides; render(); return; }
      if (shift && k === 'F9') { e.preventDefault(); grid = !grid; render(); return; }
      if (view !== 'normal') return;
      if (ctrl && shift && code === 'KeyC') { e.preventDefault(); copyFormat(); return; }
      if (ctrl && shift && code === 'KeyV') { e.preventDefault(); pasteFormat(); return; }
      if (ctrl && shift && code === 'KeyG') { e.preventDefault(); ungroupSel(); return; }
      if (ctrl && !shift && code === 'KeyG') { e.preventDefault(); groupSel(); return; }
      if (ctrl && shift && code === 'Period') { e.preventDefault(); sizeStep(1); return; }
      if (ctrl && shift && code === 'Comma') { e.preventDefault(); sizeStep(-1); return; }
      if (ctrl && shift && code === 'BracketRight') { e.preventDefault(); exec(() => P.zorder(pres, cur, sel, 'forward')); return; }
      if (ctrl && shift && code === 'BracketLeft') { e.preventDefault(); exec(() => P.zorder(pres, cur, sel, 'backward')); return; }
      if (ctrl && code === 'KeyD') { e.preventDefault(); duplicateSel(); return; }
      if (ctrl && code === 'KeyA') { e.preventDefault(); selectAll(); return; }
      if (ctrl && code === 'KeyC') { e.preventDefault(); copyShapes(); return; }
      if (ctrl && code === 'KeyV') { e.preventDefault(); pasteShapes(); return; }
      if (k === 'Escape') { if (tool) { tool = null; lock = false; } else if (child) child = null; else sel = []; render(); return; }
      if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); deleteSel(); return; }
      const arr = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[k];
      if (arr && target().length) { e.preventDefault(); const st = ctrl ? 0.02 : 0.1; exec(() => P.moveShapes(pres, slide(), target(), arr[0] * st, arr[1] * st)); return; }
      if (!ctrl && !alt && k.length === 1 && target().length === 1 && target()[0].text) { e.preventDefault(); editText(target()[0], k); return; }
      if (k === 'F2' && target().length === 1 && target()[0].text) { e.preventDefault(); editText(target()[0]); }
    };
    this.renderSheet = (key, sc) => { if (!pres) return null; const pages = P.printPages(pres); return handoutPage(pres, pages[0], 0, sc); };
    this.onShow = () => { render(); if (!pres) openBackstage('new'); };
    addEventListener('resize', () => { if (!win.hidden) render(); });
    render();
  }

  /* 인쇄 쪽(유인물·전체 페이지): scale = px/mm */
  function handoutPage(pres, page, pi, scale) {
    const pr = pres.print;
    const portrait = page.paper.w < page.paper.h;
    const pg = OA.print.page(scale, { landscape: !portrait, gray: pr.color !== 'color' });
    if (!page) return pg;
    const s = scale * 10;
    const content = h('div', { class: 'content', style: { left: 0, top: 0, width: '100%', height: '100%' } });
    for (const f of page.frames) {
      const box = h('div', { style: { position: 'absolute', left: f.x * s + 'px', top: f.y * s + 'px', width: f.w * s + 'px', height: f.h * s + 'px', outline: pr.frame ? '1px solid #000' : 'none' } });
      box.innerHTML = slideSvg(pres, f.slide, { print: true });
      if (pr.color === 'bw') box.style.filter = 'grayscale(1) contrast(1.4)';
      content.appendChild(box);
      if (pr.layout !== 'full') content.appendChild(h('div', { style: { position: 'absolute', left: f.x * s + 'px', top: (f.y + f.h + 0.1) * s + 'px', width: f.w * s + 'px', textAlign: 'center', fontSize: Math.max(7, 2.6 * scale) + 'px', color: '#444' } }, ''));
    }
    if (pr.layout !== 'full') {
      const ho = pres.handout;
      const fs = Math.max(6, 3 * scale) + 'px';
      if (ho.date) content.appendChild(h('div', { style: { position: 'absolute', right: 6 * scale + 'px', top: 6 * scale + 'px', fontSize: fs, color: '#444' } }, new Date().toISOString().slice(0, 10)));
      if (ho.pageNum) content.appendChild(h('div', { style: { position: 'absolute', right: 6 * scale + 'px', bottom: 6 * scale + 'px', fontSize: fs, color: '#444' } }, String(pi + 1)));
    }
    pg.appendChild(content);
    return pg;
  }

  OA.pptRender = { slideSvg, handoutPage };
  OA.apps = OA.apps || {};
  OA.apps.ppt = { create: (win, ctx) => new PptApp(win, ctx) };
})(typeof window !== 'undefined' ? window : globalThis);

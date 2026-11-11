/* Excel 화면: 리본·수식 입력줄·셀 격자·시트 탭·상태 표시줄·차트·피벗 필드 창·백스테이지(인쇄). 대화상자는 ui_dialogs 부분(아래 OA.xlDialogs). */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const { h, icon, menu, ribbon, dialog, msg, toast } = OA.ui;
  const U = OA.util, XL = OA.xl, NF = OA.numfmt, FX = OA.formula, PV = OA.pivot;
  const HDR_W = 40, HDR_H = 22;
  const CUR = '₩#,##0;-₩#,##0';
  const PX_PT = 96 / 72;

  function ExcelApp(win, ctx) {
    const app = this;
    this.el = win;
    this.ctx = ctx;
    let wb = ctx.model();
    let fileName = (ctx.set.id === 'equip' ? (ctx.chapter && ctx.chapter.ch === 11 ? 'ch11_작업' : 'ch10_작업') : ctx.set.id === 'ch15' ? 'ch15_작업' : ctx.set.id + '_작업') + '.xlsx';
    let savedOnce = false;
    const sel = { ranges: [{ r1: 1, c1: 1, r2: 1, c2: 1 }], r: 1, c: 1, ar: 1, ac: 1 };
    let edit = null, clip = null, chartSel = null, paneOpen = null, pick = null, backstage = null;
    const undo = [], redo = [];
    const sh = () => wb.sheets[wb.active];
    const api = {};
    this.fileName = () => (savedOnce ? fileName : '통합 문서1').replace(/\.xlsx$/, '') + ' - Excel';
    this.model = () => wb;

    // ------------------------------------------------------------------ 명령
    function exec(fn, opts) {
      const snap = JSON.stringify(wb);
      let r;
      try { r = fn(); } catch (e) { console.error(e); msg({ title: 'Microsoft Excel', text: '이 명령을 실행할 수 없습니다.\n' + e.message, icon: 'err' }); return null; }
      if (r && r.error) { msg({ title: 'Microsoft Excel', text: r.error, icon: 'warn' }); return r; }
      undo.push(snap); if (undo.length > 80) undo.shift(); redo.length = 0;
      XL.touch(wb);
      ctx.changed();
      if (!opts || !opts.norender) render();
      return r || {};
    }
    function restore(json) { wb = XL.fromJSON(json); ctx.setModel(wb); chartSel = null; render(); }
    function doUndo() { if (!undo.length) return; redo.push(JSON.stringify(wb)); restore(undo.pop()); }
    function doRedo() { if (!redo.length) return; undo.push(JSON.stringify(wb)); restore(redo.pop()); }
    function save(asNew) {
      const go = () => { savedOnce = true; ctx.saved(); render(); toast(`${fileName} 저장`); };
      if (savedOnce && !asNew) return go();
      const name = h('input', { type: 'text', value: fileName.replace(/\.xlsx$/, ''), style: { width: '260px' } });
      dialog({ title: '다른 이름으로 저장', width: 460, body: h('div', {},
        h('div', { class: 'row' }, icon('folder'), h('span', {}, '바탕 화면 › 비번호 폴더')),
        h('div', { class: 'grid2' }, '파일 이름:', name, '파일 형식:', h('select', {}, h('option', {}, 'Excel 통합 문서 (*.xlsx)'))),
        h('p', { class: 'note' }, '시험장에서는 바탕 화면의 비번호 폴더에 시험위원이 지정한 파일명으로 저장합니다(9.1절).')),
        buttons: [{ label: '저장', primary: true, onClick: () => { fileName = name.value.trim() + '.xlsx'; go(); } }, { label: '취소', cancel: true }] });
    }
    const act = () => sel.ranges[sel.ranges.length - 1];
    const allRanges = () => sel.ranges;
    function setSel(r1, c1, r2, c2, add) {
      const g = { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
      // 병합 셀 포함 확장
      let grown = true;
      while (grown) { grown = false; for (const m of sh().merges) if (m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1 && (m.r1 < g.r1 || m.c1 < g.c1 || m.r2 > g.r2 || m.c2 > g.c2)) { g.r1 = Math.min(g.r1, m.r1); g.c1 = Math.min(g.c1, m.c1); g.r2 = Math.max(g.r2, m.r2); g.c2 = Math.max(g.c2, m.c2); grown = true; } }
      if (add) sel.ranges.push(g); else sel.ranges = [g];
    }
    function goto(r, c, extend) {
      r = Math.max(1, Math.min(XL.MAX_ROW, r)); c = Math.max(1, Math.min(XL.MAX_COL, c));
      if (extend) setSel(sel.ar, sel.ac, r, c);
      else { const m = XL.mergeAt(sh(), r, c); if (m) { r = m.r1; c = m.c1; } sel.ar = r; sel.ac = c; sel.r = r; sel.c = c; setSel(r, c, r, c); }
      if (!extend) { sel.r = r; sel.c = c; }
      chartSel = null;
      render();
      scrollInto(extend ? r : sel.r, extend ? c : sel.c);
    }

    // ------------------------------------------------------------------ 화면 틀
    win.innerHTML = '';
    win.style.setProperty('--app', '#217346'); win.style.setProperty('--app2', '#185c37');
    const qatSave = h('button', { title: '저장 (Ctrl+S)', onclick: () => save() }, icon('save'));
    const qatUndo = h('button', { title: '실행 취소 (Ctrl+Z)', onclick: doUndo }, icon('undo'));
    const qatRedo = h('button', { title: '다시 실행 (Ctrl+Y)', onclick: doRedo }, icon('redo'));
    const titleName = h('span', { class: 'name' });
    win.appendChild(h('div', { class: 'titlebar' }, h('span', { class: 'autosave' }, '자동 저장 ○ 끔'), h('span', { class: 'qat' }, qatSave, qatUndo, qatRedo), titleName, h('span', { style: { width: '120px' } })));
    const ribHost = h('div');
    win.appendChild(ribHost);
    const nameBox = h('input', { class: 'namebox', title: '이름 상자', onkeydown: e => { if (e.key === 'Enter') { const g = U.parseRange(nameBox.value); if (g) { sel.ar = g.r1; sel.ac = g.c1; sel.r = g.r1; sel.c = g.c1; setSel(g.r1, g.c1, g.r2, g.c2); render(); gridWrap.focus(); } e.preventDefault(); } e.stopPropagation(); } });
    const fInput = h('input', { class: 'finput', spellcheck: false, title: '수식 입력줄' });
    win.appendChild(h('div', { class: 'xl-fbar' }, nameBox, h('span', { class: 'fx' }, h('button', { title: '취소', onclick: () => cancelEdit() }, '✕'), h('button', { title: '입력', onclick: () => commitEdit(0, 0) }, '✓'), h('i', {}, 'fx')), fInput));
    const gridWrap = h('div', { class: 'xl-gridwrap', tabindex: 0 });
    const paneHost = h('div');
    const main = h('div', { class: 'xl-main' }, gridWrap, paneHost);
    win.appendChild(main);
    const sheetBar = h('div', { class: 'xl-sheetbar' });
    win.appendChild(sheetBar);
    const status = h('div', { class: 'statusbar' });
    win.appendChild(status);

    // ------------------------------------------------------------------ 리본
    const dlg = () => OA.xlDialogs(app, api);
    const out = (label, icon2, type) => ({ type: type || 'small', label, icon: icon2, out: true });
    const fmtRanges = patch => exec(() => XL.applyStyle(wb, sh(), allRanges(), patch));
    const curStyle = () => XL.styleOf(sh(), sel.r, sel.c);
    const toggleStyle = k => fmtRanges({ [k]: curStyle()[k] ? null : true });
    const isPivotCell = () => !!PV.hitTest(wb, sh(), sel.r, sel.c);
    const curPivot = () => { const hh = PV.hitTest(wb, sh(), sel.r, sel.c); return hh ? hh.pv : null; };
    const selChart = () => chartSel ? sh().charts.find(c => c.id === chartSel.id) : null;
    const rib = ribbon(ribHost, {
      file: () => openBackstage('print'),
      initial: 'home',
      tabs: [
        { id: 'home', label: '홈', groups: [
          { label: '클립보드', items: [
            { type: 'big', label: '붙여넣기', icon: 'paste', onClick: () => doPaste(), menu: () => [{ label: '붙여넣기', icon: 'paste', onClick: () => doPaste() }, { label: '수식', onClick: () => doPaste() }, { label: '값 붙여넣기(V)', onClick: () => doPaste({ values: true }) }] },
            { type: 'stack', items: [out('잘라내기', 'cut'), { type: 'small', label: '복사', icon: 'copy', onClick: () => doCopy() }, out('서식 복사', 'brush')] }] },
          { label: '글꼴', launcher: () => dlg().formatCells('font'), items: [{ type: 'stack', items: [
            { type: 'row', items: [{ type: 'combo', width: 110, value: () => curStyle().font || '맑은 고딕', list: ['맑은 고딕', '굴림', '돋움', '바탕', 'Arial'], onChange: v => fmtRanges({ font: v }) },
              { type: 'combo', width: 44, value: () => curStyle().sz || 11, list: [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72], onChange: v => { const n = parseFloat(v); if (n > 0) fmtRanges({ sz: n }); }, tip: '글꼴 크기' },
              { type: 'icon', icon: 'growFont', label: '글꼴 크기 크게', onClick: () => fmtRanges({ sz: Math.min(409, (curStyle().sz || 11) + (curStyle().sz >= 12 ? 2 : 1)) }) },
              { type: 'icon', icon: 'shrinkFont', label: '글꼴 크기 작게', onClick: () => fmtRanges({ sz: Math.max(1, (curStyle().sz || 11) - 1) }) }] },
            { type: 'row', items: [{ type: 'icon', icon: 'bold', label: '굵게', pressed: () => !!curStyle().b, onClick: () => toggleStyle('b') },
              { type: 'icon', icon: 'italic', label: '기울임꼴', pressed: () => !!curStyle().i, onClick: () => toggleStyle('i') },
              { type: 'icon', icon: 'underline', label: '밑줄', pressed: () => !!curStyle().u, onClick: () => toggleStyle('u') },
              { type: 'icon', icon: 'border', label: '테두리', onClick: () => borderCmd(app.lastBorder || 'bottom'), menu: borderMenu },
              { type: 'icon', icon: 'fill', label: '채우기 색', menu: () => [{ html: OA.ui.palette(c => fmtRanges({ fill: c }), close => h('div', { class: 'mi', onclick: () => { close(); fmtRanges({ fill: null }); } }, '채우기 없음')) }] },
              { type: 'icon', icon: 'fontcolor', label: '글꼴 색', menu: () => [{ html: OA.ui.palette(c => fmtRanges({ color: c })) }] }] }] }] },
          { label: '맞춤', launcher: () => dlg().formatCells('align'), items: [{ type: 'stack', items: [
            { type: 'row', items: [{ type: 'icon', icon: 'alignT', label: '위쪽 맞춤', pressed: () => curStyle().va === 'top', onClick: () => fmtRanges({ va: 'top' }) },
              { type: 'icon', icon: 'alignM', label: '가운데 맞춤', pressed: () => curStyle().va === 'middle', onClick: () => fmtRanges({ va: 'middle' }) },
              { type: 'icon', icon: 'alignB', label: '아래쪽 맞춤', pressed: () => !curStyle().va || curStyle().va === 'bottom', onClick: () => fmtRanges({ va: 'bottom' }) },
              out('방향', 'orient', 'icon'), { type: 'small', icon: 'wrap', label: '텍스트 줄 바꿈', pressed: () => !!curStyle().wrap, onClick: () => toggleStyle('wrap') }] },
            { type: 'row', items: [{ type: 'icon', icon: 'alignL', label: '왼쪽 맞춤', pressed: () => curStyle().ha === 'left', onClick: () => fmtRanges({ ha: curStyle().ha === 'left' ? null : 'left' }) },
              { type: 'icon', icon: 'alignC', label: '가운데 맞춤', pressed: () => curStyle().ha === 'center', onClick: () => fmtRanges({ ha: curStyle().ha === 'center' ? null : 'center' }) },
              { type: 'icon', icon: 'alignR', label: '오른쪽 맞춤', pressed: () => curStyle().ha === 'right', onClick: () => fmtRanges({ ha: curStyle().ha === 'right' ? null : 'right' }) },
              out('내어쓰기', 'alignL', 'icon'), out('들여쓰기', 'alignR', 'icon'),
              { type: 'small', icon: 'merge', label: '병합하고 가운데 맞춤', pressed: () => !!XL.mergeAt(sh(), sel.r, sel.c), onClick: () => mergeCmd(XL.mergeAt(sh(), sel.r, sel.c) ? 'unmerge' : 'center'),
                menu: () => [{ label: '병합하고 가운데 맞춤(C)', icon: 'merge', onClick: () => mergeCmd('center') }, { label: '전체 병합(A)', onClick: () => mergeCmd('across') }, { label: '셀 병합(M)', onClick: () => mergeCmd('merge') }, { label: '셀 분할(U)', onClick: () => mergeCmd('unmerge') }] }] }] }] },
          { label: '표시 형식', launcher: () => dlg().formatCells('number'), items: [{ type: 'stack', items: [
            { type: 'row', items: [{ type: 'combo', width: 130, value: () => catName(curStyle().nf), list: ['일반', '숫자', '통화', '회계', '간단한 날짜', '자세한 날짜', '시간', '백분율', '분수', '지수', '텍스트'], onChange: v => numCat(v) }] },
            { type: 'row', items: [{ type: 'icon', icon: 'won', label: '회계 표시 형식', onClick: () => fmtRanges({ nf: '_-₩* #,##0_-;-₩* #,##0_-;_-₩* "-"_-;_-@_-' }), menu: () => [{ label: '₩ 한국어', onClick: () => fmtRanges({ nf: '_-₩* #,##0_-;-₩* #,##0_-;_-₩* "-"_-;_-@_-' }) }, { label: '기타 회계 형식...', onClick: () => dlg().formatCells('number') }] },
              { type: 'icon', icon: 'percent', label: '백분율 스타일 (Ctrl+Shift+%)', onClick: () => fmtRanges({ nf: '0%' }) },
              { type: 'icon', icon: 'comma', label: '쉼표 스타일', onClick: () => fmtRanges({ nf: '#,##0' }) },
              { type: 'icon', icon: 'decInc', label: '자릿수 늘림', onClick: () => decimals(1) }, { type: 'icon', icon: 'decDec', label: '자릿수 줄임', onClick: () => decimals(-1) }] }] }] },
          { label: '스타일', items: [out('조건부 서식', 'condfmt', 'big'), out('표 서식', 'table', 'big'), out('셀 스타일', 'cellstyle', 'big')] },
          { label: '셀', items: [out('삽입', 'insertCells', 'big'), out('삭제', 'deleteCells', 'big'),
            { type: 'big', label: '서식', icon: 'formatCells', menu: () => [{ header: '셀 크기' }, { label: '행 높이(H)...', onClick: () => dlg().rowHeight() }, { label: '열 너비(W)...', onClick: () => dlg().colWidth() }, { label: '열 너비 자동 맞춤(I)', onClick: () => autofitCols() },
              { header: '시트 구성' }, { label: '시트 이름 바꾸기(R)', onClick: () => renameTab(wb.active) }, { sep: true }, { label: '셀 서식(E)...', icon: 'formatCells', onClick: () => dlg().formatCells() }] }] },
          { label: '편집', items: [{ type: 'stack', items: [
            { type: 'small', label: '자동 합계', icon: 'sum', onClick: () => autoSum() },
            { type: 'small', label: '채우기', icon: 'fill', menu: () => [{ label: '아래쪽(D)', onClick: () => fillDown() }, { label: '오른쪽(R)', onClick: () => fillRight() }] },
            { type: 'small', label: '지우기', icon: 'clearLayout', menu: () => [{ label: '모두 지우기(A)', onClick: () => clearSel('all') }, { label: '서식 지우기(F)', onClick: () => clearSel('format') }, { label: '내용 지우기(C)', onClick: () => clearSel('contents') }] }] },
            { type: 'big', label: '정렬 및 필터', icon: 'sortAZ', menu: () => [{ label: '텍스트 오름차순 정렬(S)', icon: 'sortAZ', onClick: () => quickSort('asc') }, { label: '텍스트 내림차순 정렬(O)', icon: 'sortZA', onClick: () => quickSort('desc') }, { label: '사용자 지정 정렬(U)...', icon: 'sort', onClick: () => dlg().sort() }, { sep: true }, { label: '필터(F)', disabled: true }] },
            out('찾기 및 선택', 'find', 'big')] }] },
        { id: 'insert', label: '삽입', groups: [
          { label: '표', items: [{ type: 'big', label: '피벗 테이블', icon: 'pivot', onClick: () => dlg().createPivot() }, out('추천 피벗 테이블', 'pivot', 'big'), out('표', 'table', 'big')] },
          { label: '일러스트레이션', items: [out('그림', 'image', 'big'), out('도형', 'shapes', 'big')] },
          { label: '차트', launcher: () => dlg().insertChart('all'), items: [{ type: 'big', label: '추천 차트', icon: 'recchart', onClick: () => dlg().insertChart('rec') },
            { type: 'stack', items: [
              { type: 'row', items: [{ type: 'icon', icon: 'colchart', label: '세로 또는 가로 막대형 차트 삽입', menu: () => [{ header: '2차원 세로 막대형' }, { label: '묶은 세로 막대형', icon: 'colchart', onClick: () => quickChart('col') }, { sep: true }, { label: '기타 세로 막대형 차트(M)...', onClick: () => dlg().insertChart('all', 'col') }] },
                { type: 'icon', icon: 'linechart', label: '꺾은선형 또는 영역형 차트 삽입', menu: () => [{ header: '2차원 꺾은선형' }, { label: '꺾은선형', icon: 'linechart', onClick: () => quickChart('line') }, { label: '표식이 있는 꺾은선형', icon: 'linechart', onClick: () => quickChart('lineMarkers') }, { sep: true }, { label: '기타 꺾은선형 차트(M)...', onClick: () => dlg().insertChart('all', 'line') }] },
                out('원형 또는 도넛형 차트 삽입', 'piechart', 'icon')] },
              { type: 'row', items: [{ type: 'icon', icon: 'combochart', label: '혼합 차트 삽입', menu: () => [{ label: '묶은 세로 막대형 - 꺾은선형', icon: 'combochart', onClick: () => quickChart('combo1') }, { label: '묶은 세로 막대형 - 꺾은선형, 보조 축', icon: 'combochart', onClick: () => quickChart('combo2') }, { sep: true }, { label: '사용자 지정 콤보 차트 만들기(C)...', onClick: () => dlg().insertChart('all', 'combo') }] }, out('피벗 차트', 'chartIcon', 'icon')] }] }] },
          { label: '스파크라인', items: [out('꺾은선형', 'linechart'), out('열', 'colchart')] },
          { label: '링크', items: [out('링크', 'selectdata', 'big')] }, { label: '텍스트', items: [out('텍스트', 'textboxP', 'big')] }, { label: '기호', items: [out('기호', 'sum', 'big')] }] },
        { id: 'layout', label: '페이지 레이아웃', groups: [
          { label: '테마', items: [out('테마', 'colors', 'big')] },
          { label: '페이지 설정', launcher: () => dlg().pageSetup('page'), items: [
            { type: 'big', label: '여백', icon: 'margins', menu: () => [
              { label: '보통  위쪽 1.91cm 아래쪽 1.91cm 왼쪽 1.78cm 오른쪽 1.78cm', onClick: () => setPage({ top: 1.9, bottom: 1.9, left: 1.8, right: 1.8, header: 0.8, footer: 0.8 }) },
              { label: '넓게  위쪽 2.54cm 아래쪽 2.54cm 왼쪽 2.54cm 오른쪽 2.54cm', onClick: () => setPage({ top: 2.5, bottom: 2.5, left: 2.5, right: 2.5, header: 1.3, footer: 1.3 }) },
              { label: '좁게  위쪽 1.91cm 아래쪽 1.91cm 왼쪽 0.64cm 오른쪽 0.64cm', onClick: () => setPage({ top: 1.9, bottom: 1.9, left: 0.6, right: 0.6, header: 0.8, footer: 0.8 }) },
              { sep: true }, { label: '사용자 지정 여백(A)...', onClick: () => dlg().pageSetup('margins') }] },
            { type: 'big', label: '용지 방향', icon: 'orient', menu: () => [{ label: '세로', checked: sh().page.orient === 'portrait', onClick: () => setPage({ orient: 'portrait' }) }, { label: '가로', checked: sh().page.orient === 'landscape', onClick: () => setPage({ orient: 'landscape' }) }] },
            { type: 'big', label: '크기', icon: 'paper', menu: () => ['A4', 'A3', 'B5', 'Letter'].map(p => ({ label: { A4: 'A4 21cm x 29.7cm', A3: 'A3 29.7cm x 42cm', B5: 'B5 18.2cm x 25.7cm', Letter: 'Letter 21.59cm x 27.94cm' }[p], checked: sh().page.paper === p, onClick: () => setPage({ paper: p }) })).concat([{ sep: true }, { label: '기타 용지 크기(M)...', onClick: () => dlg().pageSetup('page') }]) },
            { type: 'big', label: '인쇄 영역', icon: 'printarea', menu: () => [{ label: '인쇄 영역 설정(S)', onClick: () => setPrintArea() }, { label: '인쇄 영역 해제(C)', onClick: () => exec(() => { sh().printArea = null; }) }] },
            out('나누기', 'breaks', 'big'), out('배경', 'image', 'big'), { type: 'big', label: '인쇄 제목', icon: 'printtitle', onClick: () => dlg().pageSetup('sheet') }] },
          { label: '크기 조정', launcher: () => dlg().pageSetup('page'), items: [{ type: 'stack', items: [
            { type: 'row', items: [h2('너비:'), { type: 'combo', width: 80, value: () => sh().page.scaleMode === 'fit' && sh().page.fitW ? sh().page.fitW + '페이지' : '자동', list: ['자동', '1페이지', '2페이지'], onChange: v => setFit('fitW', v) }] },
            { type: 'row', items: [h2('높이:'), { type: 'combo', width: 80, value: () => sh().page.scaleMode === 'fit' && sh().page.fitH ? sh().page.fitH + '페이지' : '자동', list: ['자동', '1페이지', '2페이지'], onChange: v => setFit('fitH', v) }] },
            { type: 'row', items: [h2('배율:'), { type: 'combo', width: 80, value: () => (sh().page.scaleMode === 'fit' ? 100 : sh().page.zoom) + '%', list: ['100%', '90%', '80%', '70%'], disabled: () => sh().page.scaleMode === 'fit', onChange: v => setPage({ scaleMode: 'zoom', zoom: parseInt(v, 10) || 100 }) }] }] }] },
          { label: '시트 옵션', items: [out('눈금선', 'gridlines'), out('머리글', 'headers')] },
          { label: '정렬', items: [out('앞으로 가져오기', 'front'), out('뒤로 보내기', 'back')] }] },
        { id: 'formula', label: '수식', note: '수식은 셀이나 수식 입력줄에 = 로 시작해 직접 입력합니다(함수 마법사는 다루지 않음).', groups: [
          { label: '함수 라이브러리', items: [out('함수 삽입', 'sum', 'big'), { type: 'big', label: '자동 합계', icon: 'sum', onClick: () => autoSum() }, out('논리', 'formatCells', 'big'), out('텍스트', 'label', 'big')] }] },
        { id: 'data', label: '데이터', groups: [
          { label: '데이터 가져오기 및 변환', items: [out('데이터 가져오기', 'db', 'big')] },
          { label: '쿼리 및 연결', items: [out('모두 새로 고침', 'refresh', 'big')] },
          { label: '정렬 및 필터', items: [{ type: 'stack', items: [{ type: 'icon', icon: 'sortAZ', label: '텍스트 오름차순 정렬', onClick: () => quickSort('asc') }, { type: 'icon', icon: 'sortZA', label: '텍스트 내림차순 정렬', onClick: () => quickSort('desc') }] },
            { type: 'big', label: '정렬', icon: 'sort', onClick: () => dlg().sort() }, out('필터', 'filter', 'big')] },
          { label: '데이터 도구', items: [out('텍스트 나누기', 'merge', 'big'), out('중복된 항목 제거', 'deleteCells', 'big')] },
          { label: '개요', items: [out('그룹', 'group', 'big'), out('그룹 해제', 'ungroup', 'big'), out('부분합', 'subtotal', 'big')] }] },
        { id: 'review', label: '검토', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '언어 교정', items: [out('맞춤법 검사', 'find', 'big')] }] },
        { id: 'view', label: '보기', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '통합 문서 보기', items: [out('기본', 'datasheet', 'big'), out('페이지 나누기 미리 보기', 'breaks', 'big')] }] },
        { id: 'help', label: '도움말', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '도움말', items: [out('도움말', 'find', 'big')] }] },
        // 차트 도구
        { id: 'chartDesign', label: '차트 디자인', ctx: '차트 도구', when: () => !!selChart(), groups: [
          { label: '차트 레이아웃', items: [{ type: 'big', label: '차트 요소 추가', icon: 'plus', menu: () => chartElementMenu() }, out('빠른 레이아웃', 'quicklayout', 'big')] },
          { label: '차트 스타일', items: [out('색 변경', 'colors', 'big')] },
          { label: '데이터', items: [out('행/열 전환', 'switchrc', 'big'), { type: 'big', label: '데이터 선택', icon: 'selectdata', onClick: () => dlg().selectData(selChart()) }] },
          { label: '종류', items: [{ type: 'big', label: '차트 종류 변경', icon: 'changetype', onClick: () => dlg().insertChart('all', 'combo', selChart()) }] },
          { label: '위치', items: [out('차트 이동', 'movechart', 'big')] }] },
        { id: 'chartFormat', label: '서식', ctx: '차트 도구', when: () => !!selChart(), groups: [
          { label: '현재 선택 영역', items: [{ type: 'html', render: () => h('div', { class: 'rstack' }, h('select', { style: { width: '150px', height: '22px' }, onchange: e => { const v = e.target.value; chartSel.series = v.startsWith('s') ? +v.slice(1) : null; render(); } },
            h('option', { value: 'area' }, '차트 영역'), (selChart() ? selChart().series : []).map((s, i) => h('option', { value: 's' + i, selected: chartSel && chartSel.series === i }, `계열 "${XL.seriesName(wb, s, i)}"`)))) }] },
          { label: '도형 스타일', items: [out('도형 채우기', 'shapeFill'), out('도형 윤곽선', 'shapeLine')] }] },
        // 피벗 도구
        { id: 'pvAnalyze', label: '피벗 테이블 분석', ctx: '피벗 테이블 도구', when: () => isPivotCell(), groups: [
          { label: '피벗 테이블', items: [{ type: 'stack', items: [h2('피벗 테이블 이름:'), { type: 'html', render: () => h('input', { value: curPivot() ? curPivot().name : '', style: { width: '110px', height: '20px' }, onkeydown: e => { e.stopPropagation(); if (e.key === 'Enter') { const p = curPivot(); const v = e.target.value.trim(); if (p && v) exec(() => { p.name = v; }); } } }) },
            { type: 'small', label: '옵션', icon: 'options', onClick: () => dlg().pivotOptions(curPivot()) }] }] },
          { label: '활성 필드', items: [{ type: 'stack', items: [h2('활성 필드:'), { type: 'html', render: () => h('input', { value: activeFieldCaption(), readOnly: true, style: { width: '110px', height: '20px' } }) },
            { type: 'small', label: '필드 설정', icon: 'options', onClick: () => fieldSettings() }] },
            { type: 'big', label: '필드 확장', icon: 'expand', onClick: () => fieldCollapse(false) }, { type: 'big', label: '필드 축소', icon: 'collapse', onClick: () => fieldCollapse(true) }] },
          { label: '그룹', items: [{ type: 'stack', items: [{ type: 'small', label: '그룹 선택', icon: 'group', onClick: () => groupSel() }, { type: 'small', label: '그룹 해제', icon: 'ungroup', onClick: () => ungroupSel() }, out('그룹 필드', 'group')] }] },
          { label: '필터', items: [out('슬라이서 삽입', 'filter'), out('시간 표시 막대 삽입', 'dateTime')] },
          { label: '데이터', items: [{ type: 'big', label: '새로 고침', icon: 'refresh', onClick: () => { const p = curPivot(); if (p) exec(() => PV.refresh(wb, p)); } }, out('데이터 원본 변경', 'selectdata', 'big')] },
          { label: '동작', items: [out('지우기', 'clearLayout'), out('선택', 'selectArrow'), out('피벗 테이블 이동', 'movechart')] },
          { label: '표시', items: [{ type: 'big', label: '필드 목록', icon: 'fieldlist', pressed: () => paneOpen === 'pivot', onClick: () => { paneOpen = paneOpen === 'pivot' ? null : 'pivot'; render(); } },
            { type: 'big', label: '+/- 단추', icon: 'plusminus', pressed: () => curPivot() && curPivot().showButtons, onClick: () => { const p = curPivot(); exec(() => { p.showButtons = !p.showButtons; }); } },
            { type: 'big', label: '필드 머리글', icon: 'headers', pressed: () => curPivot() && curPivot().showHeaders, onClick: () => { const p = curPivot(); exec(() => { p.showHeaders = !p.showHeaders; }); } }] }] },
        { id: 'pvDesign', label: '디자인', ctx: '피벗 테이블 도구', when: () => isPivotCell(), groups: [
          { label: '레이아웃', items: [
            { type: 'big', label: '부분합', icon: 'subtotal', menu: () => [['none', '부분합 표시 안 함(D)'], ['bottom', '그룹 하단에 모든 부분합 표시(B)'], ['top', '그룹 상단에 모든 부분합 표시(T)']].map(([k, l]) => ({ label: l, checked: curPivot().subtotals === k, onClick: () => exec(() => { curPivot().subtotals = k; }) })) },
            { type: 'big', label: '총합계', icon: 'grand', menu: () => [['off', '행 및 열의 총합계 해제(F)'], ['on', '행 및 열의 총합계 설정(N)'], ['row', '행의 총합계만 설정(R)'], ['col', '열의 총합계만 설정(C)']].map(([k, l]) => ({ label: l, onClick: () => exec(() => { const p = curPivot(); p.grandRow = k === 'on' || k === 'col'; p.grandCol = k === 'on' || k === 'row'; }) })) },
            { type: 'big', label: '보고서 레이아웃', icon: 'layout', menu: () => [['compact', '압축 형식으로 표시(C)'], ['outline', '개요 형식으로 표시(O)'], ['tabular', '테이블 형식으로 표시(T)']].map(([k, l]) => ({ label: l, checked: curPivot().layout === k, onClick: () => exec(() => { curPivot().layout = k; }) }))
              .concat([{ sep: true }, { label: '모든 항목 레이블 반복(R)', checked: curPivot().repeat, onClick: () => exec(() => { curPivot().repeat = true; }) }, { label: '항목 레이블 반복 안 함(N)', checked: !curPivot().repeat, onClick: () => exec(() => { curPivot().repeat = false; }) }]) },
            out('빈 행', 'blankrow', 'big')] },
          { label: '피벗 테이블 스타일 옵션', items: [{ type: 'stack', items: [{ type: 'check', label: '행 머리글', checked: true, disabled: true, onChange: () => {} }, { type: 'check', label: '열 머리글', checked: true, disabled: true, onChange: () => {} }] }] },
          { label: '피벗 테이블 스타일', items: [out('스타일', 'cellstyle', 'big')] }] },
      ],
    });
    function h2(t) { return { type: 'html', render: () => h('span', { style: { color: '#444', padding: '0 4px', lineHeight: '22px' } }, t) }; }

    // ------------------------------------------------------------------ 서식 명령
    function borderMenu() {
      const it = (k, l) => ({ label: l, icon: k === 'all' ? 'borderAll' : 'border', onClick: () => borderCmd(k) });
      return [{ header: '테두리' }, it('bottom', '아래쪽 테두리(O)'), it('top', '위쪽 테두리(P)'), it('left', '왼쪽 테두리(L)'), it('right', '오른쪽 테두리(R)'), { sep: true }, it('none', '테두리 없음(N)'), it('all', '모든 테두리(A)'), it('outside', '바깥쪽 테두리(S)'), it('thick', '굵은 바깥쪽 테두리(T)'), { sep: true }, { label: '다른 테두리(M)...', onClick: () => dlg().formatCells('border') }];
    }
    function borderCmd(k) { app.lastBorder = k; exec(() => { for (const g of allRanges()) XL.setBorder(wb, sh(), g, k); }); }
    function mergeCmd(mode) {
      const g = act();
      if (mode !== 'unmerge' && g.r1 === g.r2 && g.c1 === g.c2) return;
      const lost = mode === 'unmerge' ? 0 : XL.lostByMerge(sh(), g);
      const go = () => exec(() => { for (const x of allRanges()) XL.merge(wb, sh(), x, mode); });
      if (lost) msg({ title: 'Microsoft Excel', text: '셀을 병합하면 왼쪽 위의 값만 유지되고 나머지 값은 없어집니다.', buttons: ['확인', '취소'], icon: 'warn' }).then(r => { if (r === '확인') go(); });
      else go();
    }
    function catName(nf) {
      if (!nf || NF.isGeneral(nf)) return '일반';
      if (/%/.test(nf)) return '백분율';
      if (/₩/.test(nf) && /\*/.test(nf)) return '회계';
      if (/₩/.test(nf)) return '통화';
      if (/[yd]/i.test(nf) && NF.isDateFmt(nf)) return '간단한 날짜';
      if (NF.isDateFmt(nf)) return '시간';
      if (nf === '@') return '텍스트';
      return '숫자';
    }
    function numCat(v) {
      const map = { '일반': NF.GENERAL, '숫자': '0.00', '통화': CUR, '회계': '_-₩* #,##0_-;-₩* #,##0_-;_-₩* "-"_-;_-@_-', '간단한 날짜': 'yyyy-mm-dd', '자세한 날짜': 'yyyy"년" m"월" d"일"', '시간': 'h:mm:ss', '백분율': '0.00%', '분수': '# ?/?', '지수': '0.00E+00', '텍스트': '@' };
      if (map[v] !== undefined) fmtRanges({ nf: map[v] });
    }
    function decimals(d) {
      const nf = curStyle().nf || NF.GENERAL;
      let base = NF.isGeneral(nf) ? '0' : nf;
      const secs = base.split(';').map(s => {
        if (d > 0) return /\.0+/.test(s) ? s.replace(/\.(0+)/, (m, z) => '.' + z + '0') : s.replace(/([0#])(?![0#,])/, '$1.0');
        return s.replace(/\.(0+)/, (m, z) => z.length > 1 ? '.' + z.slice(1) : '');
      });
      fmtRanges({ nf: secs.join(';') });
    }
    function setPage(patch) { exec(() => { Object.assign(sh().page, patch); }); }
    function setFit(k, v) {
      exec(() => {
        const p = sh().page;
        const n = v === '자동' ? 0 : parseInt(v, 10) || 1;
        p[k] = n;
        p.scaleMode = (p.fitW || p.fitH) ? 'fit' : 'zoom';
        if (p.scaleMode === 'zoom') p.zoom = 100;
      });
    }
    function setPrintArea() {
      const g = act();
      exec(() => { sh().printArea = sel.ranges.length > 1 ? U.rangeStr(g) : U.rangeStr(g); });
      toast('인쇄 영역: ' + U.rangeStr(act(), true));
    }
    function autofitCols() { exec(() => { const g = act(); for (let c = g.c1; c <= g.c2; c++) sh().colW[c] = fitWidth(c); }); }
    function fitWidth(c) {
      let w = 40;
      for (let r = 1; r <= XL.MAX_ROW; r++) { if (XL.mergeAt(sh(), r, c)) continue; const t = XL.display(wb, sh(), r, c).text; if (t) { const sz = (XL.styleOf(sh(), r, c).sz || 11) * PX_PT; w = Math.max(w, textW(t, sz) + 10); } }
      return Math.min(400, Math.round(w));
    }
    function textW(t, px) { let w = 0; for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2fff ? px * 1.0 : px * 0.58; return w; }
    function clearSel(what) { exec(() => { for (const g of allRanges()) XL.clearRange(wb, sh(), g, what); }); }
    function autoSum() {
      const g = act();
      // 선택 범위 바로 아래(또는 활성 셀 위쪽 숫자 범위)에 SUM
      if (g.r1 === g.r2 && g.c1 === g.c2) {
        let r = g.r1 - 1; while (r >= 1 && typeof XL.getValue(wb, sh(), r, g.c1) === 'number') r--;
        const top = r + 1;
        startEdit(top < g.r1 ? `=SUM(${U.addr(top, g.c1)}:${U.addr(g.r1 - 1, g.c1)})` : '=SUM()', true);
        return;
      }
      exec(() => { for (let c = g.c1; c <= g.c2; c++) XL.setInput(wb, sh(), g.r2 + 1, c, `=SUM(${U.addr(g.r1, c)}:${U.addr(g.r2, c)})`); });
    }
    function fillDown() {
      const g = act();
      if (g.r1 === g.r2) { if (g.r1 > 1) exec(() => XL.fill(wb, sh(), { r1: g.r1 - 1, c1: g.c1, r2: g.r1 - 1, c2: g.c2 }, { r1: g.r1 - 1, c1: g.c1, r2: g.r2, c2: g.c2 })); return; }
      exec(() => XL.fill(wb, sh(), { r1: g.r1, c1: g.c1, r2: g.r1, c2: g.c2 }, g));
    }
    function fillRight() { const g = act(); if (g.c1 === g.c2) return; exec(() => XL.fill(wb, sh(), { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c1 }, g)); }
    function quickSort(order) {
      let g = act();
      const keyCol = sel.c;
      if (g.r1 === g.r2 && g.c1 === g.c2) g = XL.currentRegion(wb, sh(), sel.r, sel.c);
      const header = typeof XL.getValue(wb, sh(), g.r1, keyCol) === 'string' && typeof XL.getValue(wb, sh(), g.r1 + 1, keyCol) !== 'string' || hasHeaderGuess(g);
      exec(() => XL.sortRange(wb, sh(), g, [{ col: keyCol, order }], header));
    }
    function hasHeaderGuess(g) { for (let c = g.c1; c <= g.c2; c++) { const a = XL.getValue(wb, sh(), g.r1, c), b = XL.getValue(wb, sh(), g.r1 + 1, c); if (typeof a === 'string' && typeof b === 'number') return true; } return false; }
    function quickChart(kind) {
      const g = sel.ranges;
      if (g.length === 1 && g[0].r1 === g[0].r2 && g[0].c1 === g[0].c2) { const cr = XL.currentRegion(wb, sh(), sel.r, sel.c); sel.ranges = [cr]; }
      const series = XL.seriesFromSelection(wb, sh(), sel.ranges);
      if (!series.length) return msg({ title: 'Microsoft Excel', text: '차트를 만들 데이터 범위를 선택하십시오.', icon: 'warn' });
      const r = exec(() => {
        if (kind === 'combo1' || kind === 'combo2') series.forEach((s, i) => { s.type = i === 0 ? 'col' : 'line'; s.secondary = kind === 'combo2' && i > 0; });
        else series.forEach(s => { s.type = kind; });
        const g0 = sel.ranges[sel.ranges.length - 1];
        const ch = XL.newChart(wb, sh(), series, kind, { x: XL.colLeft(sh(), g0.c2 + 2), y: XL.rowTop(sh(), Math.max(1, g0.r1)) });
        chartSel = { id: ch.id, series: null };
        return {};
      });
      if (r && !r.error) rib.select('chartDesign');
    }
    api.quickChart = quickChart;

    // ------------------------------------------------------------------ 복사·붙여 넣기
    function doCopy() {
      const g = act();
      clip = XL.copy(wb, sh(), g);
      clip.src = { sheet: sh().name, g };
      render();
      ctx.msg('대상을 선택한 후 Enter 키를 누르거나 붙여넣기를 선택합니다.');
    }
    function doPaste(opts) {
      if (!clip) return;
      const g = act();
      exec(() => {
        if (opts && opts.values) {
          const c2 = U.clone(clip);
          c2.pivot = null;
          c2.cells = c2.cells.map(row => row.map(cl => { if (!cl) return cl; const o = Object.assign({}, cl); return o; }));
          const src = XL.sheetOf(wb, clip.src.sheet);
          for (let i = 0; i < c2.cells.length; i++) for (let j = 0; j < c2.cells[0].length; j++) { const cl = c2.cells[i][j]; if (cl && cl.f) { cl.v = XL.getValue(wb, src, clip.rg.r1 + i, clip.rg.c1 + j); delete cl.f; } }
          return XL.paste(wb, sh(), c2, g.r1, g.c1, { dest: g });
        }
        return XL.paste(wb, sh(), clip, g.r1, g.c1, { dest: g });
      });
    }

    // ------------------------------------------------------------------ 피벗 명령
    function pivotSelection() {
      const hits = [];
      for (const g of sel.ranges) for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const x = PV.hitTest(wb, sh(), r, c); if (x && x.cell.kind === 'label') hits.push(x); }
      return hits;
    }
    function activeFieldCaption() {
      const x = PV.hitTest(wb, sh(), sel.r, sel.c);
      if (!x) return '';
      const c = PV.cacheOf(wb, x.pv);
      if (x.cell.field) return PV.caption(c, x.pv, x.cell.field);
      if (x.cell.kind === 'value' || x.cell.kind === 'valhdr') { const v = x.pv.values[x.cell.vi || 0]; return v ? v.name : ''; }
      return '';
    }
    function groupSel() {
      const hits = pivotSelection();
      if (!hits.length) return msg({ title: 'Microsoft Excel', text: '선택 항목을 그룹으로 묶을 수 없습니다.', icon: 'warn' });
      const pv = hits[0].pv, field = hits[0].cell.field;
      const keys = [...new Set(hits.filter(x => x.pv === pv && x.cell.field === field).map(x => x.cell.key))];
      const c = PV.cacheOf(wb, pv);
      if (keys.length === 1 && PV.isNumericField(c, field)) return dlg().groupRange(pv, field);
      const r = exec(() => PV.group(wb, sh(), pv, field, keys));
      if (r && !r.error) toast('그룹을 만들었습니다. 그룹 셀에 이름을 입력해 바꿀 수 있습니다.');
    }
    function ungroupSel() {
      const hits = pivotSelection();
      if (!hits.length) return;
      const pv = hits[0].pv, field = hits[0].cell.field;
      exec(() => PV.ungroup(wb, sh(), pv, field, hits.map(x => x.cell.key)));
    }
    function fieldCollapse(on) {
      const x = PV.hitTest(wb, sh(), sel.r, sel.c);
      if (!x || !x.cell.field) return msg({ title: 'Microsoft Excel', text: '활성 셀이 피벗 테이블의 행 또는 열 필드에 있어야 합니다.', icon: 'warn' });
      exec(() => PV.collapseField(wb, x.pv, x.cell.field, on));
    }
    function fieldSettings() {
      const x = PV.hitTest(wb, sh(), sel.r, sel.c);
      if (!x) return;
      if (x.cell.kind === 'value' || x.cell.kind === 'valhdr') dlg().valueField(x.pv, x.cell.vi || 0);
      else if (x.cell.field) dlg().fieldSettings(x.pv, x.cell.field);
    }
    api.exec = exec; api.wb = () => wb; api.sh = sh; api.sel = sel; api.render = () => render(); api.act = act; api.setSel = setSel; api.goto = goto;
    api.startPick = cb => { pick = cb; return () => { pick = null; render(); }; };
    api.chartSel = () => chartSel; api.setChartSel = v => { chartSel = v; render(); };
    api.paneOpen = v => { paneOpen = v; render(); };

    // ------------------------------------------------------------------ 격자 그리기
    let maxR = 60, maxC = 20;
    const layers = { cells: h('div', { class: 'xl-layer' }), borders: document.createElementNS('http://www.w3.org/2000/svg', 'svg'), over: h('div'), colhdr: h('div'), rowhdr: h('div'), corner: h('div', { class: 'xl-corner' }) };
    const grid = h('div', { class: 'xl-grid' });
    layers.borders.setAttribute('style', 'position:absolute;left:0;top:0;pointer-events:none;z-index:4;overflow:visible');
    grid.append(layers.cells, layers.borders, layers.over, layers.colhdr, layers.rowhdr, layers.corner);
    gridWrap.appendChild(grid);
    Object.assign(layers.colhdr.style, { position: 'absolute', left: 0, top: 0, zIndex: 12 });
    Object.assign(layers.rowhdr.style, { position: 'absolute', left: 0, top: 0, zIndex: 12 });
    Object.assign(layers.corner.style, { left: 0, top: 0, width: HDR_W + 'px', height: HDR_H + 'px', zIndex: 13 });
    layers.corner.addEventListener('click', () => { setSel(1, 1, XL.MAX_ROW, XL.MAX_COL); sel.r = 1; sel.c = 1; render(); });
    gridWrap.addEventListener('scroll', () => {
      layers.colhdr.style.transform = `translateY(${gridWrap.scrollTop}px)`;
      layers.rowhdr.style.transform = `translateX(${gridWrap.scrollLeft}px)`;
      layers.corner.style.transform = `translate(${gridWrap.scrollLeft}px, ${gridWrap.scrollTop}px)`;
    });
    const X = c => HDR_W + XL.colLeft(sh(), c);
    const Y = r => HDR_H + XL.rowTop(sh(), r);

    function extent() {
      const u = XL.usedRange(wb, sh());
      maxR = Math.min(XL.MAX_ROW, Math.max(50, u ? u.r2 + 20 : 0, sel.r + 10));
      maxC = Math.min(XL.MAX_COL, Math.max(18, u ? u.c2 + 6 : 0, sel.c + 4));
      for (const ch of sh().charts) { const b = XL.cellAtXY(sh(), ch.x + ch.w, ch.y + ch.h); maxR = Math.min(XL.MAX_ROW, Math.max(maxR, b.r + 8)); maxC = Math.min(XL.MAX_COL, Math.max(maxC, b.c + 2)); }
    }
    function cellStyleCss(st, disp, isNum) {
      const css = {};
      const ha = st.ha || (disp.align === 'right' ? 'right' : disp.align === 'center' ? 'center' : 'left');
      css.textAlign = ha;
      if (st.b) css.fontWeight = '700';
      if (st.i) css.fontStyle = 'italic';
      if (st.u) css.textDecoration = 'underline';
      if (st.sz) css.fontSize = (st.sz * PX_PT) + 'px';
      if (st.color) css.color = st.color;
      if (st.fill) css.background = st.fill;
      if (st.font && st.font !== '맑은 고딕') css.fontFamily = st.font;
      return css;
    }
    function renderGrid() {
      extent();
      const s = sh();
      const W = X(maxC + 1), Hh = Y(maxR + 1);
      grid.style.width = W + 'px'; grid.style.height = Hh + 'px';
      layers.cells.innerHTML = ''; layers.over.innerHTML = ''; layers.colhdr.innerHTML = ''; layers.rowhdr.innerHTML = '';
      layers.borders.setAttribute('width', W); layers.borders.setAttribute('height', Hh);
      const frag = document.createDocumentFragment();
      const covered = new Set();
      for (const m of s.merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(r + ',' + c);
      const pvCells = new Map();
      for (const pv of s.pivots) { const g = PV.grid(wb, s, pv); if (!g) continue; g.rows.forEach((row, i) => row.forEach((cell, j) => pvCells.set((g.r0 + i) + ',' + (g.c0 + j), { cell, pv }))); }
      const filled = (r, c) => { if (pvCells.has(r + ',' + c)) return true; const cl = XL.getCell(s, r, c); return !!(cl && (cl.f || (cl.v !== null && cl.v !== undefined && cl.v !== ''))) || covered.has(r + ',' + c) || !!XL.mergeAt(s, r, c); };
      for (let r = 1; r <= maxR; r++) for (let c = 1; c <= maxC; c++) {
        if (covered.has(r + ',' + c)) continue;
        const m = XL.mergeAt(s, r, c);
        const x = X(c), y = Y(r);
        let w = XL.colW(s, c), hh = XL.rowH(s, r);
        if (m) { w = X(m.c2 + 1) - x; hh = Y(m.r2 + 1) - y; }
        const pc = pvCells.get(r + ',' + c);
        const cl = XL.getCell(s, r, c);
        if (!pc && !cl && !m) continue;
        const st = pc ? {} : XL.styleOf(s, r, c);
        const disp = XL.display(wb, s, r, c);
        const div = document.createElement('div');
        div.className = 'xl-cell' + (m ? ' merged' : '');
        div.style.left = x + 'px'; div.style.top = y + 'px'; div.style.width = w + 'px'; div.style.height = hh + 'px';
        div.style.lineHeight = (m ? hh : XL.rowH(s, r)) - 1 + 'px';
        Object.assign(div.style, cellStyleCss(st, disp));
        if (st.va === 'top') div.style.lineHeight = ((st.sz || 11) * PX_PT * 1.3) + 'px';
        if (st.wrap) { div.style.whiteSpace = 'normal'; div.style.lineHeight = ((st.sz || 11) * PX_PT * 1.25) + 'px'; }
        if (pc) {
          const k = pc.cell;
          if (k.hdr) div.classList.add('pv-hdr');
          if (k.bold) div.style.fontWeight = '700';
          if (k.grand || k.kind === 'grand') div.style.borderTop = '1px solid #9bc2e6';
          if (k.indent) div.style.paddingLeft = (3 + k.indent * 14) + 'px';
          if (k.btn && pc.pv.showButtons) {
            const b = h('span', { class: 'pvbtn', title: k.btn === '-' ? '축소' : '확장' }, k.btn === '-' ? '−' : '+');
            b.dataset.pvbtn = '1'; b.dataset.r = r; b.dataset.c = c;
            div.appendChild(b);
          }
          if (k.dd && pc.pv.showHeaders) div.appendChild(h('span', { class: 'pvdd' }, '▼'));
          div.appendChild(document.createTextNode(disp.text));
          if (k.align === 'left' || k.kind === 'label') div.style.textAlign = 'left';
          if (k.placeholder) { div.style.color = '#666'; div.style.whiteSpace = 'normal'; div.style.overflow = 'visible'; div.style.width = (w * 3) + 'px'; div.style.lineHeight = '1.3'; }
        } else {
          div.textContent = disp.text;
          // 넘침: 왼쪽 맞춤 문자가 칸보다 길고 오른쪽이 비어 있으면 이어서 보이게
          if (!m && typeof XL.getValue(wb, s, r, c) === 'string' && (st.ha || 'left') === 'left' && !st.wrap) {
            const need = textW(disp.text, (st.sz || 11) * PX_PT) + 6;
            if (need > w) { let cc = c + 1, ww = w; while (ww < need && cc <= maxC && !filled(r, cc)) { ww += XL.colW(s, cc); cc++; } if (ww > w) { div.style.width = ww + 'px'; div.classList.add('ovf'); div.style.background = 'transparent'; } }
          }
          if (typeof XL.getValue(wb, s, r, c) === 'number' && disp.text.length && textW(disp.text, (st.sz || 11) * PX_PT) + 4 > w && !m) {
            // G/표준은 소수 자릿수부터 줄이고(Excel 동작), 지정 서식은 ### 로 표시
            const fpx = (st.sz || 11) * PX_PT;
            const fit = NF.isGeneral(XL.cellFormat(wb, s, r, c)) ? NF.fitGeneral(XL.getValue(wb, s, r, c), t => textW(t, fpx) + 4 <= w) : null;
            div.textContent = fit !== null ? fit : '#'.repeat(Math.max(1, Math.floor(w / 8)));
          }
        }
        frag.appendChild(div);
      }
      layers.cells.appendChild(frag);
      // 테두리
      let lines = '';
      for (let r = 1; r <= maxR; r++) for (let c = 1; c <= maxC; c++) {
        if (covered.has(r + ',' + c) && !XL.mergeAt(s, r, c)) continue;
        const b = XL.borderAt(s, r, c);
        if (!b.t && !b.b && !b.l && !b.r) continue;
        const m = XL.mergeAt(s, r, c);
        const x1 = X(c), y1 = Y(r), x2 = X(c + 1), y2 = Y(r + 1);
        const mm = m || { r1: r, c1: c, r2: r, c2: c };
        if (b.t && r === mm.r1) lines += `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y1}" stroke="#000" stroke-width="${b.t}"/>`;
        if (b.b && r === mm.r2) lines += `<line x1="${x1}" y1="${y2}" x2="${x2}" y2="${y2}" stroke="#000" stroke-width="${b.b}"/>`;
        if (b.l && c === mm.c1) lines += `<line x1="${x1}" y1="${y1}" x2="${x1}" y2="${y2}" stroke="#000" stroke-width="${b.l}"/>`;
        if (b.r && c === mm.c2) lines += `<line x1="${x2}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#000" stroke-width="${b.r}"/>`;
      }
      layers.borders.innerHTML = lines;
      // 머리글
      const g = act();
      for (let c = 1; c <= maxC; c++) {
        const d = h('div', { class: 'xl-colhdr' + (sel.ranges.some(x => c >= x.c1 && c <= x.c2) ? ' on' : ''), style: { left: X(c) + 'px', top: 0, width: XL.colW(s, c) + 'px', height: HDR_H + 'px', lineHeight: HDR_H + 'px' } }, U.colName(c));
        const rs = h('div', { class: 'rs' });
        rs.addEventListener('mousedown', e => colResize(e, c));
        rs.addEventListener('dblclick', () => exec(() => { sh().colW[c] = fitWidth(c); }));
        d.appendChild(rs);
        d.addEventListener('mousedown', e => { if (e.target === rs) return; e.preventDefault(); commitEdit(0, 0, true); setSel(1, c, XL.MAX_ROW, c, e.ctrlKey); sel.r = 1; sel.c = c; render(); });
        d.addEventListener('contextmenu', e => { e.preventDefault(); menu(d, [{ label: '열 너비(W)...', onClick: () => dlg().colWidth() }, { label: '열 너비 자동 맞춤', onClick: () => autofitCols() }, { label: '셀 서식(F)...', onClick: () => dlg().formatCells() }], { at: { x: e.clientX, y: e.clientY } }); });
        layers.colhdr.appendChild(d);
      }
      for (let r = 1; r <= maxR; r++) {
        const d = h('div', { class: 'xl-rowhdr' + (sel.ranges.some(x => r >= x.r1 && r <= x.r2) ? ' on' : ''), style: { left: 0, top: Y(r) + 'px', width: HDR_W + 'px', height: XL.rowH(s, r) + 'px', lineHeight: XL.rowH(s, r) + 'px' } }, String(r));
        d.addEventListener('mousedown', e => { e.preventDefault(); commitEdit(0, 0, true); setSel(r, 1, r, XL.MAX_COL, e.ctrlKey); sel.r = r; sel.c = 1; render(); });
        d.addEventListener('contextmenu', e => { e.preventDefault(); menu(d, [{ label: '행 높이(R)...', onClick: () => dlg().rowHeight() }, { label: '셀 서식(F)...', onClick: () => dlg().formatCells() }], { at: { x: e.clientX, y: e.clientY } }); });
        layers.rowhdr.appendChild(d);
      }
      // 선택
      for (const x of sel.ranges) {
        const box = h('div', { class: 'xl-sel' + (sel.ranges.length > 1 ? ' multi' : '') + (x.r1 === x.r2 && x.c1 === x.c2 ? ' active' : ''), style: { left: X(x.c1) - 1 + 'px', top: Y(x.r1) - 1 + 'px', width: X(x.c2 + 1) - X(x.c1) + 1 + 'px', height: Y(x.r2 + 1) - Y(x.r1) + 1 + 'px' } });
        layers.over.appendChild(box);
      }
      const am = XL.mergeAt(s, sel.r, sel.c) || { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c };
      layers.over.appendChild(h('div', { class: 'xl-sel active', style: { left: X(am.c1) - 1 + 'px', top: Y(am.r1) - 1 + 'px', width: X(am.c2 + 1) - X(am.c1) + 1 + 'px', height: Y(am.r2 + 1) - Y(am.r1) + 1 + 'px', borderWidth: sel.ranges.length > 1 || (g.r1 !== g.r2 || g.c1 !== g.c2) ? '1px' : '2px', background: 'transparent' } }));
      if (sel.ranges.length === 1 && !chartSel) {
        const fh = h('div', { class: 'xl-fill', style: { left: X(g.c2 + 1) - 4 + 'px', top: Y(g.r2 + 1) - 4 + 'px' }, title: '채우기 핸들' });
        fh.addEventListener('mousedown', fillDrag);
        layers.over.appendChild(fh);
      }
      if (clip && clip.src && clip.src.sheet === s.name) {
        const c2 = clip.src.g;
        layers.over.appendChild(h('div', { class: 'xl-ants', style: { left: X(c2.c1) - 1 + 'px', top: Y(c2.r1) - 1 + 'px', width: X(c2.c2 + 1) - X(c2.c1) + 1 + 'px', height: Y(c2.r2 + 1) - Y(c2.r1) + 1 + 'px' } }));
      }
      // 차트
      for (const ch of s.charts) layers.over.appendChild(chartEl(ch));
    }

    // ------------------------------------------------------------------ 차트 요소
    function chartEl(ch) {
      const selected = chartSel && chartSel.id === ch.id;
      const box = h('div', { class: 'xl-chart' + (selected ? ' sel' : ''), style: { left: HDR_W + ch.x + 'px', top: HDR_H + ch.y + 'px', width: ch.w + 'px', height: ch.h + 'px' } });
      box.innerHTML = OA.chart.svg(wb, ch, { w: ch.w, h: ch.h, selSeries: selected ? chartSel.series : null });
      box.addEventListener('mousedown', e => {
        if (e.button !== 0) return;
        e.stopPropagation();
        commitEdit(0, 0, true);
        const part = e.target.getAttribute && e.target.getAttribute('data-part');
        if (e.target.classList.contains('h') || e.target.classList.contains('plusbtn')) return;
        const was = chartSel && chartSel.id === ch.id;
        chartSel = { id: ch.id, series: part && part.startsWith('series-') && was ? +part.slice(7) : (part && part.startsWith('series-') ? +part.slice(7) : null) };
        if (!was) rib.select('chartDesign');
        // 끌어 옮기기
        const sx = e.clientX, sy = e.clientY, ox = ch.x, oy = ch.y;
        let moved = false;
        OA.ui.drag(e, ev => {
          const dx = ev.clientX - sx, dy = ev.clientY - sy;
          if (!moved && Math.abs(dx) + Math.abs(dy) < 4) return;
          moved = true;
          box.style.left = HDR_W + ox + dx + 'px'; box.style.top = HDR_H + oy + dy + 'px';
        }, ev => {
          if (moved) { const dx = ev.clientX - sx, dy = ev.clientY - sy; exec(() => { ch.x = Math.max(0, ox + dx); ch.y = Math.max(0, oy + dy); }); }
          else render();
        });
      });
      box.addEventListener('dblclick', e => {
        const part = e.target.getAttribute && e.target.getAttribute('data-part');
        if (part === 'title') editChartText(ch, 'title');
        else if (part === 'axisX') editChartText(ch, 'x');
        else if (part === 'axisY') editChartText(ch, 'y');
      });
      box.addEventListener('contextmenu', e => {
        e.preventDefault(); e.stopPropagation();
        chartSel = { id: ch.id, series: chartSel && chartSel.id === ch.id ? chartSel.series : null };
        menu(box, [{ label: '차트 종류 변경(Y)...', icon: 'changetype', onClick: () => dlg().insertChart('all', 'combo', ch) }, { label: '데이터 선택(E)...', icon: 'selectdata', onClick: () => dlg().selectData(ch) },
          { sep: true }, { label: '차트 제목 편집', onClick: () => editChartText(ch, 'title') }, { label: '삭제(D)', onClick: () => deleteChart(ch) }], { at: { x: e.clientX, y: e.clientY } });
      });
      if (selected) {
        for (const [px, py, cur] of [[0, 0, 'nwse'], [0.5, 0, 'ns'], [1, 0, 'nesw'], [1, 0.5, 'ew'], [1, 1, 'nwse'], [0.5, 1, 'ns'], [0, 1, 'nesw'], [0, 0.5, 'ew']]) {
          const hd = h('div', { class: 'h', style: { left: px * ch.w - 5 + 'px', top: py * ch.h - 5 + 'px', cursor: cur + '-resize' } });
          hd.addEventListener('mousedown', e => {
            e.stopPropagation();
            const sx = e.clientX, sy = e.clientY, o = { x: ch.x, y: ch.y, w: ch.w, h: ch.h };
            OA.ui.drag(e, ev => {
              const dx = ev.clientX - sx, dy = ev.clientY - sy;
              const n = Object.assign({}, o);
              if (px === 0) { n.x = o.x + dx; n.w = o.w - dx; } if (px === 1) n.w = o.w + dx;
              if (py === 0) { n.y = o.y + dy; n.h = o.h - dy; } if (py === 1) n.h = o.h + dy;
              if (n.w < 120 || n.h < 100) return;
              Object.assign(box.style, { left: HDR_W + n.x + 'px', top: HDR_H + n.y + 'px', width: n.w + 'px', height: n.h + 'px' });
              box._n = n;
            }, () => { if (box._n) { const n = box._n; exec(() => Object.assign(ch, n)); } });
          });
          box.appendChild(hd);
        }
        const plus = h('button', { class: 'plusbtn', title: '차트 요소', style: { top: '0px' } }, '+');
        plus.addEventListener('mousedown', e => e.stopPropagation());
        plus.addEventListener('click', e => { e.stopPropagation(); chartElementsPane(plus, ch); });
        box.appendChild(plus);
      }
      return box;
    }
    function deleteChart(ch) { exec(() => { sh().charts = sh().charts.filter(c => c !== ch); }); chartSel = null; render(); }
    function editChartText(ch, which) {
      const cur = which === 'title' ? XL.chartTitleText(wb, ch) : ch.axisTitle[which].text;
      const inp = h('input', { type: 'text', value: cur, style: { width: '300px' } });
      dialog({ title: which === 'title' ? '차트 제목' : (which === 'x' ? '가로 축 제목' : '세로 축 제목'), width: 380, body: h('div', {}, h('p', { class: 'note' }, '제목 상자 안의 글자를 고칩니다.'), inp),
        buttons: [{ label: '확인', primary: true, onClick: () => exec(() => { if (which === 'title') { ch.title.show = true; ch.title.text = inp.value; } else { ch.axisTitle[which].show = true; ch.axisTitle[which].text = inp.value; } }) }, { label: '취소', cancel: true }] });
    }
    const LABELS_LINE = [['center', '가운데'], ['left', '왼쪽'], ['right', '오른쪽'], ['top', '위쪽'], ['bottom', '아래쪽']];
    const LABELS_COL = [['center', '가운데'], ['inEnd', '안쪽 끝에'], ['inBase', '안쪽 기준선에'], ['outEnd', '바깥쪽 끝에']];
    function targetSeries(ch) { return chartSel && chartSel.series !== null && chartSel.series !== undefined && ch.series[chartSel.series] ? [ch.series[chartSel.series]] : ch.series; }
    function chartElementMenu() {
      const ch = selChart();
      if (!ch) return [];
      const ts = targetSeries(ch);
      const isLine = ts.every(s => s.type === 'line' || s.type === 'lineMarkers');
      return [
        { label: '축(X)', sub: [{ label: '기본 가로', checked: ch.axes, onClick: () => exec(() => { ch.axes = !ch.axes; }) }, { label: '기본 세로', checked: ch.axes, onClick: () => exec(() => { ch.axes = !ch.axes; }) }] },
        { label: '축 제목(A)', sub: [{ label: '기본 가로', checked: ch.axisTitle.x.show, onClick: () => exec(() => { ch.axisTitle.x.show = !ch.axisTitle.x.show; }) }, { label: '기본 세로', checked: ch.axisTitle.y.show, onClick: () => exec(() => { ch.axisTitle.y.show = !ch.axisTitle.y.show; }) }] },
        { label: '차트 제목(C)', sub: [{ label: '없음', checked: !ch.title.show, onClick: () => exec(() => { ch.title.show = false; }) }, { label: '차트 위', checked: ch.title.show, onClick: () => exec(() => { ch.title.show = true; }) }] },
        { label: '데이터 레이블(D)', sub: [{ label: '없음', onClick: () => exec(() => ts.forEach(s => { s.labels = null; })) }].concat((isLine ? LABELS_LINE : LABELS_COL).map(([k, l]) => ({ label: l, checked: ts.every(s => s.labels === k), onClick: () => exec(() => ts.forEach(s => { s.labels = k; })) }))) },
        { label: '데이터 표(B)', disabled: true }, { label: '오차 막대(E)', disabled: true },
        { label: '눈금선(G)', sub: [{ label: '기본 주 가로', checked: ch.gridlines, onClick: () => exec(() => { ch.gridlines = !ch.gridlines; }) }] },
        { label: '범례(L)', sub: [{ label: '없음', checked: !ch.legend.show, onClick: () => exec(() => { ch.legend.show = false; }) }].concat([['right', '오른쪽'], ['top', '위쪽'], ['left', '왼쪽'], ['bottom', '아래쪽']].map(([k, l]) => ({ label: l, checked: ch.legend.show && ch.legend.pos === k, onClick: () => exec(() => { ch.legend.show = true; ch.legend.pos = k; }) }))) },
        { label: '추세선(T)', disabled: true },
      ];
    }
    function chartElementsPane(anchor, ch) {
      const ts = targetSeries(ch);
      const isLine = ts.every(s => s.type === 'line' || s.type === 'lineMarkers');
      const row = (label, checked, toggle, sub) => {
        const cb = h('input', { type: 'checkbox', checked, onchange: () => { toggle(cb.checked); } });
        const r = h('div', { class: 'r' }, cb, h('span', {}, label));
        if (sub) { const go = h('span', { class: 'go', title: '자세히' }, '▶'); go.addEventListener('click', e => { e.stopPropagation(); menu(go, sub, { level: 1, side: true }); }); r.appendChild(go); }
        return r;
      };
      menu(anchor, [{ html: close => h('div', { class: 'chel' }, h('div', { class: 'ttl' }, '차트 요소' + (ts.length === 1 && ch.series.length > 1 ? ` — 계열 "${XL.seriesName(wb, ts[0], ch.series.indexOf(ts[0]))}"` : '')),
        row('축', ch.axes, v => exec(() => { ch.axes = v; })),
        row('축 제목', ch.axisTitle.x.show || ch.axisTitle.y.show, v => exec(() => { ch.axisTitle.x.show = v; ch.axisTitle.y.show = v; }), [{ label: '기본 가로', checked: ch.axisTitle.x.show, onClick: () => exec(() => { ch.axisTitle.x.show = !ch.axisTitle.x.show; }) }, { label: '기본 세로', checked: ch.axisTitle.y.show, onClick: () => exec(() => { ch.axisTitle.y.show = !ch.axisTitle.y.show; }) }]),
        row('차트 제목', ch.title.show, v => exec(() => { ch.title.show = v; })),
        row('데이터 레이블', ts.some(s => s.labels), v => exec(() => ts.forEach(s => { s.labels = v ? (isLine ? 'right' : 'outEnd') : null; })), (isLine ? LABELS_LINE : LABELS_COL).map(([k, l]) => ({ label: l, checked: ts.every(s => s.labels === k), onClick: () => exec(() => ts.forEach(s => { s.labels = k; })) }))),
        row('데이터 표', false, () => {}), row('오차 막대', false, () => {}),
        row('눈금선', ch.gridlines, v => exec(() => { ch.gridlines = v; })),
        row('범례', ch.legend.show, v => exec(() => { ch.legend.show = v; }), [['right', '오른쪽'], ['top', '위쪽'], ['left', '왼쪽'], ['bottom', '아래쪽']].map(([k, l]) => ({ label: l, checked: ch.legend.show && ch.legend.pos === k, onClick: () => exec(() => { ch.legend.show = true; ch.legend.pos = k; }) }))),
        row('추세선', false, () => {})) }], { at: (() => { const r = anchor.getBoundingClientRect(); return { x: r.right + 4, y: r.top }; })() });
    }

    // ------------------------------------------------------------------ 마우스
    function cellFromEvent(e) {
      const r0 = grid.getBoundingClientRect();
      const x = e.clientX - r0.left - HDR_W, y = e.clientY - r0.top - HDR_H;
      return XL.cellAtXY(sh(), Math.max(0, x), Math.max(0, y));
    }
    gridWrap.addEventListener('mousedown', e => {
      if (e.button === 2) return;
      if (e.target.closest('.xl-colhdr, .xl-rowhdr, .xl-corner, .xl-chart, .xl-edit')) return;
      if (e.target.dataset && e.target.dataset.pvbtn) {
        e.preventDefault();
        const x = PV.hitTest(wb, sh(), +e.target.dataset.r, +e.target.dataset.c);
        if (x) exec(() => PV.toggle(wb, x.pv, x.cell.field, x.cell.key));
        return;
      }
      const p = cellFromEvent(e);
      if (edit && edit.formula && refInsertable()) { e.preventDefault(); insertRef(p, e); return; }
      commitEdit(0, 0, true);
      chartSel = null;
      e.preventDefault();
      if (e.shiftKey) { setSel(sel.ar, sel.ac, p.r, p.c); render(); return; }
      const m = XL.mergeAt(sh(), p.r, p.c);
      const r = m ? m.r1 : p.r, c = m ? m.c1 : p.c;
      sel.ar = r; sel.ac = c; sel.r = r; sel.c = c;
      setSel(r, c, r, c, e.ctrlKey);
      render();
      cellInput.focus({ preventScroll: true });
      if (pick) pick(pickText());
      OA.ui.drag(e, ev => {
        const q = cellFromEvent(ev);
        const last = sel.ranges.pop();
        void last;
        setSel(sel.ar, sel.ac, q.r, q.c, true);
        render();
        if (pick) pick(pickText());
      });
    });
    function pickText() {
      const sheetName = sh().name;
      return sel.ranges.map(g => U.rangeStr(g, true, sheetName)).join(',');
    }
    gridWrap.addEventListener('dblclick', e => {
      if (e.target.closest('.xl-colhdr, .xl-rowhdr, .xl-chart')) return;
      const p = cellFromEvent(e);
      if (PV.hitTest(wb, sh(), p.r, p.c) && PV.hitTest(wb, sh(), p.r, p.c).cell.kind === 'value') return;
      startEdit(null);
    });
    gridWrap.addEventListener('contextmenu', e => {
      if (e.target.closest('.xl-colhdr, .xl-rowhdr, .xl-chart')) return;
      e.preventDefault();
      const p = cellFromEvent(e);
      if (!sel.ranges.some(g => U.inRange(g, p.r, p.c))) { sel.ar = p.r; sel.ac = p.c; sel.r = p.r; sel.c = p.c; setSel(p.r, p.c, p.r, p.c); render(); }
      const hit = PV.hitTest(wb, sh(), p.r, p.c);
      const at = { at: { x: e.clientX, y: e.clientY } };
      if (hit) {
        const k = hit.cell;
        const items = [{ label: '복사(C)', icon: 'copy', onClick: doCopy }, { label: '셀 서식(F)...', icon: 'formatCells', onClick: () => dlg().formatCells() }, { label: '새로 고침(R)', icon: 'refresh', onClick: () => exec(() => PV.refresh(wb, hit.pv)) }, { sep: true }];
        if (k.kind === 'label') {
          items.push({ label: '그룹(G)...', icon: 'group', onClick: groupSel }, { label: '그룹 해제(U)...', icon: 'ungroup', onClick: ungroupSel }, { sep: true },
            { label: '확장/축소(E)', sub: [{ label: '확장(X)', icon: 'expand', onClick: () => exec(() => PV.setCollapsed(wb, hit.pv, k.field, k.key, false)) }, { label: '축소(O)', icon: 'collapse', onClick: () => exec(() => PV.setCollapsed(wb, hit.pv, k.field, k.key, true)) }, { sep: true },
              { label: '전체 필드 확장(E)', onClick: () => exec(() => PV.collapseField(wb, hit.pv, k.field, false)) }, { label: '전체 필드 축소(C)', onClick: () => exec(() => PV.collapseField(wb, hit.pv, k.field, true)) }] },
            { label: `"${PV.caption(PV.cacheOf(wb, hit.pv), hit.pv, k.field)}" 제거(V)`, onClick: () => exec(() => PV.removeField(wb, hit.pv, k.field)) }, { sep: true });
        }
        if (k.kind === 'value' || k.kind === 'valhdr') items.push({ label: '값 필드 설정(N)...', onClick: () => dlg().valueField(hit.pv, k.vi || 0) });
        items.push({ label: '피벗 테이블 옵션(O)...', icon: 'options', onClick: () => dlg().pivotOptions(hit.pv) }, { label: paneOpen === 'pivot' ? '필드 목록 숨기기(D)' : '필드 목록 표시(D)', icon: 'fieldlist', onClick: () => { paneOpen = paneOpen === 'pivot' ? null : 'pivot'; render(); } });
        return menu(gridWrap, items, at);
      }
      menu(gridWrap, [out2('잘라내기(T)', 'cut'), { label: '복사(C)', icon: 'copy', onClick: doCopy }, { label: '붙여넣기 옵션:', icon: 'paste', disabled: !clip, onClick: () => doPaste() }, { sep: true },
        out2('삽입(I)...'), out2('삭제(D)...'), { label: '내용 지우기(N)', onClick: () => clearSel('contents') }, { sep: true }, { label: '정렬(O)', sub: [{ label: '텍스트 오름차순 정렬(S)', icon: 'sortAZ', onClick: () => quickSort('asc') }, { label: '텍스트 내림차순 정렬(O)', icon: 'sortZA', onClick: () => quickSort('desc') }, { label: '사용자 지정 정렬(U)...', icon: 'sort', onClick: () => dlg().sort() }] },
        { sep: true }, { label: '셀 서식(F)...', icon: 'formatCells', onClick: () => dlg().formatCells() }, { label: '열 너비(W)...', onClick: () => dlg().colWidth() }], at);
    });
    function out2(label, ic) { return { label, icon: ic, disabled: true, tip: '이 시뮬레이터에서 다루지 않는 기능' }; }
    function colResize(e, c) {
      e.preventDefault(); e.stopPropagation();
      const sx = e.clientX, w0 = XL.colW(sh(), c);
      OA.ui.drag(e, ev => { sh().colW[c] = Math.max(8, w0 + ev.clientX - sx); renderGrid(); }, () => { const w1 = sh().colW[c]; sh().colW[c] = w0; exec(() => { sh().colW[c] = w1; }); });
    }
    function fillDrag(e) {
      e.preventDefault(); e.stopPropagation();
      const src = U.clone(act());
      let dest = src;
      const ghost = h('div', { class: 'xl-ants' });
      layers.over.appendChild(ghost);
      OA.ui.drag(e, ev => {
        const q = cellFromEvent(ev);
        const down = q.r - src.r2, right = q.c - src.c2;
        if (Math.abs(down) >= Math.abs(right)) dest = down >= 0 ? { r1: src.r1, c1: src.c1, r2: Math.max(src.r2, q.r), c2: src.c2 } : { r1: Math.min(src.r1, q.r), c1: src.c1, r2: src.r2, c2: src.c2 };
        else dest = right >= 0 ? { r1: src.r1, c1: src.c1, r2: src.r2, c2: Math.max(src.c2, q.c) } : { r1: src.r1, c1: Math.min(src.c1, q.c), r2: src.r2, c2: src.c2 };
        Object.assign(ghost.style, { left: X(dest.c1) - 1 + 'px', top: Y(dest.r1) - 1 + 'px', width: X(dest.c2 + 1) - X(dest.c1) + 1 + 'px', height: Y(dest.r2 + 1) - Y(dest.r1) + 1 + 'px' });
      }, () => {
        ghost.remove();
        if (U.sameRange(dest, src)) return;
        exec(() => XL.fill(wb, sh(), src, dest));
        sel.ranges = [dest];
        render();
      });
    }
    function scrollInto(r, c) {
      const x = X(c), y = Y(r);
      if (x - HDR_W < gridWrap.scrollLeft) gridWrap.scrollLeft = x - HDR_W;
      if (x + XL.colW(sh(), c) > gridWrap.scrollLeft + gridWrap.clientWidth) gridWrap.scrollLeft = x + XL.colW(sh(), c) - gridWrap.clientWidth + 4;
      if (y - HDR_H < gridWrap.scrollTop) gridWrap.scrollTop = y - HDR_H;
      if (y + XL.rowH(sh(), r) > gridWrap.scrollTop + gridWrap.clientHeight) gridWrap.scrollTop = y + XL.rowH(sh(), r) - gridWrap.clientHeight + 4;
    }

    // ------------------------------------------------------------------ 셀 편집
    // 활성 셀 위에 늘 보이지 않는 입력 상자를 두고(한글 입력기·붙여 넣기 대응), 글자가 들어오면 편집 상자로 바꾼다
    const cellInput = h('input', { class: 'xl-edit', spellcheck: false, autocomplete: 'off', 'aria-label': '셀 입력' });
    grid.appendChild(cellInput);
    function editGeom() {
      const m = XL.mergeAt(sh(), sel.r, sel.c);
      const r = m ? m.r1 : sel.r, c = m ? m.c1 : sel.c;
      return { r, c, w: m ? X(m.c2 + 1) - X(c) : XL.colW(sh(), c), hh: m ? Y(m.r2 + 1) - Y(r) : XL.rowH(sh(), r) };
    }
    function hideInput() {
      const g = editGeom();
      Object.assign(cellInput.style, { left: X(g.c) + 'px', top: Y(g.r) + 'px', width: '2px', minWidth: '0', height: g.hh + 'px', opacity: 0, pointerEvents: 'none', borderWidth: '0' });
    }
    function focusCatcher() {
      const a = document.activeElement;
      if (!win.hidden && !backstage && !OA.ui.anyDialog() && (!a || a === document.body || a === gridWrap || a === cellInput || win.contains(a) && a.tagName !== 'INPUT' && a.tagName !== 'TEXTAREA' && a.tagName !== 'SELECT')) cellInput.focus({ preventScroll: true });
    }
    function beginEdit(mode, text) {
      if (chartSel) return;
      if (PV.hitTest(wb, sh(), sel.r, sel.c) && PV.hitTest(wb, sh(), sel.r, sel.c).cell.kind === 'value') { cellInput.value = ''; msg({ title: 'Microsoft Excel', text: '피벗 테이블의 이 부분은 변경할 수 없습니다.', icon: 'warn' }); return; }
      const g = editGeom();
      const st = XL.styleOf(sh(), g.r, g.c);
      if (text !== null && text !== undefined) cellInput.value = text;
      Object.assign(cellInput.style, { left: X(g.c) - 1 + 'px', top: Y(g.r) - 1 + 'px', minWidth: g.w + 2 + 'px', width: Math.max(g.w + 2, textW(cellInput.value, (st.sz || 11) * PX_PT) + 16) + 'px', height: g.hh + 2 + 'px', opacity: 1, pointerEvents: 'auto', borderWidth: '2px',
        fontSize: ((st.sz || 11) * PX_PT) + 'px', fontWeight: st.b ? '700' : '400', textAlign: 'left' });
      edit = { r: g.r, c: g.c, inp: cellInput, mode, get formula() { return (document.activeElement === fInput ? fInput.value : cellInput.value).startsWith('='); } };
      fInput.value = cellInput.value;
      if (status.firstChild) status.firstChild.textContent = mode === 'enter' ? '입력' : '편집';
    }
    function startEdit(initial) {
      beginEdit(initial !== null && initial !== undefined ? 'enter' : 'edit', initial !== null && initial !== undefined ? initial : XL.editText(wb, sh(), sel.r, sel.c));
      if (!edit) return;
      cellInput.focus({ preventScroll: true });
      const L = cellInput.value.length;
      cellInput.setSelectionRange(L, L);
    }
    cellInput.addEventListener('input', () => {
      if (!edit) { beginEdit('enter', null); return; }
      fInput.value = cellInput.value;
      const st = XL.styleOf(sh(), edit.r, edit.c);
      cellInput.style.width = Math.max(parseFloat(cellInput.style.minWidth) || 60, textW(cellInput.value, (st.sz || 11) * PX_PT) + 16) + 'px';
    });
    cellInput.addEventListener('compositionstart', () => { if (!edit) beginEdit('enter', ''); });
    cellInput.addEventListener('keydown', e => { if (edit) editKey(e); });
    cellInput.addEventListener('paste', e => { if (!edit && clip) { e.preventDefault(); doPaste(); } });
    function refInsertable() {
      if (!edit) return false;
      const inp = document.activeElement === fInput ? fInput : edit.inp;
      const pos = inp.selectionStart;
      const before = inp.value.slice(0, pos).trimEnd();
      return /[=(,+\-*/&<>:^]$/.test(before) || /[=(,+\-*/&<>^]\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$/i.test(inp.value.slice(0, pos)) && edit.lastRef;
    }
    function insertRef(p, e) {
      const inp = document.activeElement === fInput ? fInput : edit.inp;
      let pos = inp.selectionStart;
      let val = inp.value;
      if (edit.lastRef && val.slice(0, pos).endsWith(edit.lastRef)) { val = val.slice(0, pos - edit.lastRef.length) + val.slice(pos); pos -= edit.lastRef.length; }
      const prefix = sh().name !== edit.sheetName && edit.sheetName ? sh().name + '!' : '';
      let ref = prefix + U.addr(p.r, p.c);
      inp.value = val.slice(0, pos) + ref + val.slice(pos);
      edit.lastRef = ref;
      const np = pos + ref.length;
      inp.setSelectionRange(np, np);
      fInput.value = inp.value;
      // 끌어서 범위
      const start = p;
      OA.ui.drag(e, ev => {
        const q = cellFromEvent(ev);
        const rr = prefix + (q.r === start.r && q.c === start.c ? U.addr(q.r, q.c) : U.rangeStr({ r1: Math.min(start.r, q.r), c1: Math.min(start.c, q.c), r2: Math.max(start.r, q.r), c2: Math.max(start.c, q.c) }));
        inp.value = val.slice(0, pos) + rr + val.slice(pos);
        edit.lastRef = rr;
        inp.setSelectionRange(pos + rr.length, pos + rr.length);
        fInput.value = inp.value;
      }, () => inp.focus());
    }
    function editKey(e) {
      e.stopPropagation();
      const inp = e.target;
      if (e.key === 'Enter') { e.preventDefault(); commitEdit(e.shiftKey ? -1 : 1, 0); return; }
      if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); return; }
      if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); return; }
      if (e.key === 'F4') { e.preventDefault(); const r = FX.cycleRef(inp.value, inp.selectionStart); inp.value = r.text; inp.setSelectionRange(r.caret, r.caret); fInput.value = inp.value; if (edit) edit.lastRef = null; return; }
      if (e.key === 'F2') { e.preventDefault(); if (edit) edit.mode = edit.mode === 'enter' ? 'edit' : 'enter'; status.firstChild.textContent = edit.mode === 'enter' ? '입력' : '편집'; return; }
      if (edit && edit.mode === 'enter' && !edit.formula && ['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
        e.preventDefault();
        const d = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
        commitEdit(d[0], d[1]);
        return;
      }
      if (e.ctrlKey && (e.key === 's' || e.key === 'S')) { e.preventDefault(); commitEdit(0, 0); save(); return; }
    }
    function commitEdit(dr, dc, silent) {
      if (!edit) return true;
      const { r, c } = edit;
      const val = document.activeElement === fInput || edit.viaBar ? fInput.value : cellInput.value;
      edit = null;
      cellInput.value = '';
      hideInput();
      const before = XL.editText(wb, sh(), r, c);
      if (val !== before) {
        const res = exec(() => XL.setInput(wb, sh(), r, c, val), { norender: true });
        if (res && res.error) { render(); return false; }
      }
      if (dr || dc) {
        const g = act();
        const multi = g.r1 !== g.r2 || g.c1 !== g.c2;
        if (multi) { // 선택 범위 안에서 이동
          let nr = sel.r + dr, nc = sel.c + dc;
          if (nr > g.r2) { nr = g.r1; nc = nc + 1 > g.c2 ? g.c1 : nc + 1; } if (nc > g.c2) { nc = g.c1; nr = nr + 1 > g.r2 ? g.r1 : nr + 1; }
          sel.r = nr; sel.c = nc;
          render();
        } else goto(r + dr, c + dc);
      } else render();
      if (!silent || document.activeElement === fInput) cellInput.focus({ preventScroll: true });
      return true;
    }
    function cancelEdit() { if (!edit) return; edit = null; cellInput.value = ''; hideInput(); render(); cellInput.focus({ preventScroll: true }); }
    fInput.addEventListener('focus', () => {
      if (edit) { edit.viaBar = true; return; }
      if (PV.hitTest(wb, sh(), sel.r, sel.c) || chartSel) return;
      beginEdit('edit', XL.editText(wb, sh(), sel.r, sel.c));
      if (edit) edit.viaBar = true;
      fInput.value = cellInput.value;
    });
    fInput.addEventListener('input', () => { if (edit) { cellInput.value = fInput.value; } });
    fInput.addEventListener('keydown', e => {
      if (!edit) return;
      if (e.key === 'Enter') { e.preventDefault(); edit.inp.value = fInput.value; commitEdit(1, 0); }
      else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
      else if (e.key === 'F4') { e.preventDefault(); const r = FX.cycleRef(fInput.value, fInput.selectionStart); fInput.value = r.text; fInput.setSelectionRange(r.caret, r.caret); edit.inp.value = fInput.value; }
      e.stopPropagation();
    });

    // ------------------------------------------------------------------ 키보드(표 9-4 + 기본 이동)
    this.handleKey = e => {
      if (backstage) { if (e.key === 'Escape') closeBackstage(); if (e.ctrlKey && e.code === 'KeyS') { e.preventDefault(); save(); } return; }
      const t = e.target;
      const inField = t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT') && !t.classList.contains('xl-edit');
      if (edit) return;
      const k = e.key, code = e.code, ctrl = e.ctrlKey || e.metaKey, shift = e.shiftKey, alt = e.altKey;
      if (ctrl && code === 'KeyS') { e.preventDefault(); save(); return; }
      if (inField) return;
      if (ctrl && code === 'KeyZ') { e.preventDefault(); doUndo(); return; }
      if (ctrl && code === 'KeyY') { e.preventDefault(); doRedo(); return; }
      if (chartSel) {
        if (k === 'Delete') { e.preventDefault(); deleteChart(selChart()); return; }
        if (k === 'Escape') { chartSel = null; render(); return; }
        return;
      }
      if (ctrl && code === 'KeyC') { e.preventDefault(); doCopy(); return; }
      if (ctrl && code === 'KeyV') { e.preventDefault(); doPaste(); return; }
      if (ctrl && code === 'Digit1' && !shift) { e.preventDefault(); dlg().formatCells(); return; }
      if (ctrl && shift && code === 'Digit5') { e.preventDefault(); fmtRanges({ nf: '0%' }); return; }
      if (ctrl && shift && code === 'Digit4') { e.preventDefault(); fmtRanges({ nf: CUR }); return; }
      if (ctrl && code === 'KeyD') { e.preventDefault(); fillDown(); return; }
      if (ctrl && code === 'KeyA') { e.preventDefault(); const g = XL.currentRegion(wb, sh(), sel.r, sel.c); const same = U.sameRange(g, act()); if (same || (g.r1 === g.r2 && g.c1 === g.c2)) setSel(1, 1, XL.MAX_ROW, XL.MAX_COL); else setSel(g.r1, g.c1, g.r2, g.c2); render(); return; }
      if (shift && k === 'F11') { e.preventDefault(); newSheet(true); return; }
      if (ctrl && k === 'PageDown') { e.preventDefault(); switchSheet(Math.min(wb.sheets.length - 1, wb.active + 1)); return; }
      if (ctrl && k === 'PageUp') { e.preventDefault(); switchSheet(Math.max(0, wb.active - 1)); return; }
      if (alt && k === 'F1') { e.preventDefault(); quickChart('col'); return; }
      if (k === 'F2') { e.preventDefault(); startEdit(null); return; }
      if (k === 'Escape') { if (clip) { clip = null; ctx.msg(''); render(); } return; }
      if (k === 'Delete') { e.preventDefault(); clearSel('contents'); return; }
      if (k === 'Enter') { e.preventDefault(); if (clip && clip.src) { doPaste(); clip = null; render(); return; } moveIn(shift ? -1 : 1, 0); return; }
      if (k === 'Tab') { e.preventDefault(); moveIn(0, shift ? -1 : 1); return; }
      const arrows = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (arrows[k]) {
        e.preventDefault();
        const [dr, dc] = arrows[k];
        if (shift) { const g = act(); const er = (g.r1 === sel.ar ? g.r2 : g.r1) + dr, ec = (g.c1 === sel.ac ? g.c2 : g.c1) + dc; setSel(sel.ar, sel.ac, Math.max(1, er), Math.max(1, ec)); render(); scrollInto(Math.max(1, er), Math.max(1, ec)); }
        else { const m = XL.mergeAt(sh(), sel.r, sel.c); const base = m ? { r: dr > 0 ? m.r2 : m.r1, c: dc > 0 ? m.c2 : m.c1 } : { r: sel.r, c: sel.c }; goto(base.r + dr, base.c + dc); }
        return;
      }
      if (!ctrl && !alt && k.length === 1) { if (t === cellInput) return; e.preventDefault(); startEdit(k); return; }
      if ((k === 'Process' || e.isComposing) && t !== cellInput) { startEdit(''); }
    };
    function moveIn(dr, dc) {
      const g = act();
      if (g.r1 === g.r2 && g.c1 === g.c2) { goto(sel.r + dr, sel.c + dc); return; }
      let nr = sel.r + dr, nc = sel.c + dc;
      if (nr > g.r2) { nr = g.r1; nc++; } if (nr < g.r1) { nr = g.r2; nc--; } if (nc > g.c2) { nc = g.c1; nr++; } if (nc < g.c1) { nc = g.c2; nr--; }
      if (nr > g.r2) nr = g.r1; if (nc > g.c2) nc = g.c1; if (nr < g.r1) nr = g.r2; if (nc < g.c1) nc = g.c2;
      sel.r = nr; sel.c = nc; render();
    }

    // ------------------------------------------------------------------ 시트 탭
    function newSheet(before) {
      exec(() => XL.addSheet(wb, null, before ? wb.active : wb.active + 1));
      sel.ranges = [{ r1: 1, c1: 1, r2: 1, c2: 1 }]; sel.r = 1; sel.c = 1; sel.ar = 1; sel.ac = 1;
      render();
    }
    function switchSheet(i) {
      commitEdit(0, 0, true);
      if (pick) { wb.active = i; render(); return; }
      wb.active = i; chartSel = null;
      sel.ranges = [{ r1: 1, c1: 1, r2: 1, c2: 1 }]; sel.r = 1; sel.c = 1; sel.ar = 1; sel.ac = 1;
      render();
    }
    function renameTab(i) {
      const tab = sheetBar.querySelectorAll('.xl-tab')[i];
      if (!tab) return;
      const inp = h('input', { value: wb.sheets[i].name });
      tab.innerHTML = ''; tab.appendChild(inp); inp.focus(); inp.select();
      const done = ok => { if (ok && inp.value !== wb.sheets[i].name) exec(() => XL.renameSheet(wb, i, inp.value)); else render(); };
      inp.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') done(true); if (e.key === 'Escape') done(false); });
      inp.addEventListener('blur', () => done(true));
    }
    function renderSheetBar() {
      sheetBar.innerHTML = '';
      sheetBar.append(h('button', { class: 'nav', title: '이전 시트', onclick: () => switchSheet(Math.max(0, wb.active - 1)) }, '◀'), h('button', { class: 'nav', title: '다음 시트', onclick: () => switchSheet(Math.min(wb.sheets.length - 1, wb.active + 1)) }, '▶'));
      wb.sheets.forEach((s, i) => {
        const tab = h('div', { class: 'xl-tab' + (i === wb.active ? ' on' : '') }, s.name);
        tab.addEventListener('click', () => { if (i !== wb.active) switchSheet(i); });
        tab.addEventListener('dblclick', () => renameTab(i));
        tab.addEventListener('contextmenu', e => { e.preventDefault(); menu(tab, [{ label: '삽입(I)...', onClick: () => newSheet(true) }, { label: '삭제(D)', onClick: () => msg({ title: 'Microsoft Excel', text: '이 시트를 영구적으로 삭제합니다. 계속하시겠습니까?', buttons: ['삭제', '취소'] }).then(r => { if (r === '삭제') exec(() => XL.deleteSheet(wb, i)); }) }, { label: '이름 바꾸기(R)', onClick: () => renameTab(i) }], { at: { x: e.clientX, y: e.clientY - 80 } }); });
        sheetBar.appendChild(tab);
      });
      sheetBar.appendChild(h('button', { class: 'plus', title: '새 시트', onclick: () => newSheet(false) }, '⊕'));
    }

    // ------------------------------------------------------------------ 상태 표시줄·작업창
    function renderStatus() {
      const nums = [];
      let cnt = 0;
      for (const g of sel.ranges) { if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > 4000) continue; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const v = XL.getValue(wb, sh(), r, c); if (v !== null && v !== '') cnt++; if (typeof v === 'number') nums.push(v); } }
      const nf = XL.cellFormat(wb, sh(), sel.r, sel.c);
      const f = v => NF.format(v, NF.isDateFmt(nf) ? NF.GENERAL : nf).text;
      status.innerHTML = '';
      status.appendChild(h('span', {}, pick ? '범위 선택' : '준비'));
      status.appendChild(h('span', { class: 'sp' }));
      if (cnt > 1) {
        if (nums.length) status.appendChild(h('span', {}, '평균: ' + f(nums.reduce((s, x) => s + x, 0) / nums.length)));
        status.appendChild(h('span', {}, '개수: ' + cnt));
        if (nums.length) status.appendChild(h('span', {}, '합계: ' + f(nums.reduce((s, x) => s + x, 0))));
      }
      status.appendChild(h('span', {}, '100%'));
    }
    function renderPane() {
      paneHost.innerHTML = '';
      const pv = curPivot() || (paneOpen === 'pivot' && app._lastPivot);
      if (paneOpen === 'pivot' && pv) { app._lastPivot = pv; paneHost.appendChild(OA.xlPivotPane(app, api, pv, () => { paneOpen = null; render(); })); }
    }
    let lastWasPivot = false;
    function render() {
      if (edit) return;
      titleName.textContent = app.fileName();
      qatUndo.disabled = !undo.length; qatRedo.disabled = !redo.length;
      const inPv = isPivotCell();
      if (inPv && !lastWasPivot) { paneOpen = 'pivot'; }
      lastWasPivot = inPv;
      if (!inPv && paneOpen === 'pivot' && !sh().pivots.length) paneOpen = null;
      rib.render();
      renderGrid();
      renderSheetBar();
      renderStatus();
      renderPane();
      const g = act();
      nameBox.value = chartSel ? (selChart() ? selChart().name : '') : (g.r1 === g.r2 && g.c1 === g.c2) || XL.mergeAt(sh(), g.r1, g.c1) && U.sameRange(XL.mergeAt(sh(), g.r1, g.c1), g) ? U.addr(sel.r, sel.c) : U.addr(sel.r, sel.c);
      if (document.activeElement !== fInput) fInput.value = chartSel ? '' : XL.editText(wb, sh(), sel.r, sel.c);
      hideInput();
      focusCatcher();
    }
    api.renderNow = render;

    // ------------------------------------------------------------------ 백스테이지
    function openBackstage(page) {
      commitEdit(0, 0, true);
      closeBackstage();
      backstage = h('div', { class: 'backstage' });
      const nav = h('div', { class: 'bs-nav' });
      const mainEl = h('div', { class: 'bs-main' });
      const pages = [['back', '←'], ['home', '홈'], ['new', '새로 만들기'], ['open', '열기'], ['-'], ['info', '정보'], ['save', '저장'], ['saveas', '다른 이름으로 저장'], ['print', '인쇄'], ['share', '공유'], ['export', '내보내기'], ['close', '닫기']];
      const show = p => {
        [...nav.children].forEach(b => b.classList.toggle('on', b.dataset.p === p));
        mainEl.innerHTML = '';
        if (p === 'print') mainEl.appendChild(printPage());
        else if (p === 'info') mainEl.appendChild(h('div', {}, h('h2', {}, '정보'), h('p', {}, '파일: ' + (savedOnce ? fileName : '(저장 안 함)')), h('p', {}, '시트: ' + wb.sheets.map(s => s.name).join(', '))));
        else if (p === 'new') mainEl.appendChild(h('div', {}, h('h2', {}, '새로 만들기'), h('button', { class: 'print-set big', style: { width: '140px', height: '110px' }, onclick: () => msg({ title: 'Microsoft Excel', text: '현재 통합 문서를 닫고 새 통합 문서를 만듭니다. 저장하지 않은 내용은 없어집니다.', buttons: ['새 통합 문서', '취소'], icon: 'q' }).then(r => { if (r === '새 통합 문서') { exec(() => { const n = XL.newWorkbook(); Object.keys(wb).forEach(k => delete wb[k]); Object.assign(wb, JSON.parse(JSON.stringify(n))); XL.touch(wb); }); savedOnce = false; closeBackstage(); } }) }, icon('newfile', 40), '새 통합 문서')));
        else if (p === 'home') mainEl.appendChild(h('div', {}, h('h2', {}, '좋은 하루입니다'), h('p', {}, '왼쪽 목록에서 인쇄를 눌러 미리 보기를 확인합니다.')));
        else mainEl.appendChild(h('div', {}, h('h2', {}, pages.find(x => x[0] === p)[1]), h('p', { class: 'note' }, '실기 과제에 쓰지 않는 화면입니다.')));
      };
      for (const [p, l] of pages) {
        if (p === '-') { nav.appendChild(h('div', { style: { height: '1px', background: 'rgba(255,255,255,.3)', margin: '6px 14px' } })); continue; }
        const b = h('button', { class: p === 'back' ? 'back' : '', 'data-p': p, disabled: ['open', 'share', 'export'].includes(p) }, l);
        b.addEventListener('click', () => { if (p === 'back') return closeBackstage(); if (p === 'save') { closeBackstage(); return save(); } if (p === 'saveas') { closeBackstage(); return save(true); } if (p === 'close') { closeBackstage(); return; } show(p); });
        nav.appendChild(b);
      }
      backstage.append(nav, mainEl);
      win.appendChild(backstage);
      show(page || 'print');
    }
    function closeBackstage() { if (backstage) { backstage.remove(); backstage = null; render(); } }
    api.openBackstage = openBackstage; api.closeBackstage = closeBackstage;
    function printPage() {
      const s = sh();
      const pg = s.page;
      let pi = 0;
      const prev = h('div', { class: 'print-prev' });
      const navT = h('span');
      const btnP = h('button', { onclick: () => { pi--; draw(); } }, '◀'), btnN = h('button', { onclick: () => { pi++; draw(); } }, '▶');
      const draw = () => {
        const lay = XL.printLayout(wb, s);
        pi = Math.max(0, Math.min(pi, lay.pages.length - 1));
        prev.innerHTML = '';
        const scale = Math.min(2.4, (prev.clientHeight || 700) / 310, 2.4);
        prev.appendChild(renderPage(wb, s, lay, pi, scale));
        navT.textContent = ` ${pi + 1} / ${lay.pages.length} `;
        btnP.disabled = pi <= 0; btnN.disabled = pi >= lay.pages.length - 1;
        prev.appendChild(h('div', { class: 'print-nav' }, btnP, navT, btnN));
      };
      const pset = (ic, main2, sub, items) => {
        const el = h('div', { class: 'pset' }, icon(ic), h('div', { class: 'pt' }, main2, sub ? h('small', {}, sub) : null), h('span', { class: 'pdd' }, '▼'));
        el.addEventListener('click', () => menu(el, items()));
        return el;
      };
      const fitLabel = () => pg.scaleMode === 'fit' ? (pg.fitW === 1 && pg.fitH === 1 ? ['한 페이지에 시트 맞추기', '한 페이지에 맞게 출력물을 축소합니다.'] : pg.fitW === 1 && !pg.fitH ? ['한 페이지에 모든 열 맞추기', ''] : !pg.fitW && pg.fitH === 1 ? ['한 페이지에 모든 행 맞추기', ''] : ['사용자 지정 배율', '']) : pg.zoom === 100 ? ['현재 설정된 용지', '실제 크기의 용지에 시트를 인쇄합니다.'] : ['사용자 지정 배율', pg.zoom + '%'];
      const marginLabel = () => (pg.top === 1.9 && pg.left === 1.8 && pg.right === 1.8 && pg.bottom === 1.9) ? ['기본 여백', '위쪽: 1.91cm 아래쪽: 1.91cm 왼쪽: 1.78cm'] : ['사용자 지정 여백', `위쪽: ${pg.top}cm 아래쪽: ${pg.bottom}cm`];
      const setP = (patch) => { setPage(patch); refresh(); };
      const setsEl = h('div', { class: 'print-set' });
      const refresh = () => {
        setsEl.innerHTML = '';
        setsEl.append(
          h('div', { style: { display: 'flex', gap: '14px', alignItems: 'center', marginBottom: '10px' } },
            h('button', { class: 'big', onclick: () => doPrint() }, icon('printer', 36), '인쇄'),
            h('div', {}, '복사본: ', h('input', { type: 'number', value: 1, min: 1, style: { width: '50px' } }))),
          h('div', { style: { fontWeight: 700, margin: '6px 0' } }, '프린터'),
          h('div', { class: 'pset' }, icon('printer'), h('div', { class: 'pt' }, '시험장 지정 프린터', h('small', {}, '준비'))),
          h('div', { style: { fontWeight: 700, margin: '10px 0 4px' } }, '설정'),
          pset('paper', '활성 시트 인쇄', '활성 시트만 인쇄합니다.', () => [{ label: '활성 시트 인쇄', checked: true }, { label: '전체 통합 문서 인쇄', disabled: true }, { label: '선택 영역 인쇄', disabled: true }]),
          h('div', { class: 'row', style: { gap: '6px' } }, '페이지:', h('input', { type: 'text', style: { width: '40px' } }), '위치:', h('input', { type: 'text', style: { width: '40px' } })),
          pset('copy', '한 부씩 인쇄', '1,2,3  1,2,3  1,2,3', () => [{ label: '한 부씩 인쇄', checked: true }]),
          pset('orient', pg.orient === 'portrait' ? '세로 방향' : '가로 방향', '', () => [{ label: '세로 방향', checked: pg.orient === 'portrait', onClick: () => setP({ orient: 'portrait' }) }, { label: '가로 방향', checked: pg.orient === 'landscape', onClick: () => setP({ orient: 'landscape' }) }]),
          pset('paper', pg.paper, { A4: '21cm x 29.7cm', A3: '29.7cm x 42cm', B5: '18.2cm x 25.7cm', Letter: '21.59cm x 27.94cm' }[pg.paper], () => ['A4', 'A3', 'B5', 'Letter'].map(p => ({ label: p, checked: pg.paper === p, onClick: () => setP({ paper: p }) }))),
          pset('margins', ...marginLabel(), () => [{ label: '보통', onClick: () => setP({ top: 1.9, bottom: 1.9, left: 1.8, right: 1.8 }) }, { label: '넓게', onClick: () => setP({ top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 }) }, { label: '좁게', onClick: () => setP({ top: 1.9, bottom: 1.9, left: 0.6, right: 0.6 }) }, { sep: true }, { label: '사용자 지정 여백...', onClick: () => dlg().pageSetup('margins', refresh) }]),
          pset('onepage', ...fitLabel(), () => [{ label: '현재 설정된 용지', onClick: () => setP({ scaleMode: 'zoom', zoom: 100 }) }, { label: '한 페이지에 시트 맞추기', onClick: () => setP({ scaleMode: 'fit', fitW: 1, fitH: 1 }) }, { label: '한 페이지에 모든 열 맞추기', onClick: () => setP({ scaleMode: 'fit', fitW: 1, fitH: 0 }) }, { label: '한 페이지에 모든 행 맞추기', onClick: () => setP({ scaleMode: 'fit', fitW: 0, fitH: 1 }) }, { sep: true }, { label: '사용자 지정 배율 옵션...', onClick: () => dlg().pageSetup('page', refresh) }]),
          h('div', { style: { textAlign: 'right' } }, h('button', { class: 'linkbtn', onclick: () => dlg().pageSetup('page', refresh) }, '페이지 설정')));
        draw();
      };
      const wrap = h('div', {}, h('h2', {}, '인쇄'), h('div', { class: 'print-wrap' }, setsEl, prev));
      setTimeout(refresh, 0);
      return wrap;
    }
    function doPrint() {
      const s = sh();
      const key = s.pivots.length ? '5-2' : '5-1';
      const lay = XL.printLayout(wb, s);
      ctx.printed(key, { sheet: s.name, pages: lay.pages.length });
      closeBackstage();
    }
    this.renderSheet = (key, scale) => {
      const s = key === '5-2' ? wb.sheets.find(x => x.pivots.length) : (wb.sheets.find(x => x.name === '작업표') || wb.sheets.find(x => !x.pivots.length));
      if (!s) return null;
      const lay = XL.printLayout(wb, s);
      return renderPage(wb, s, lay, 0, scale, true);
    };
    this.onShow = () => { render(); setTimeout(focusCatcher, 0); };
    render();
  }

  /* A4 쪽 그리기(미리 보기·5매 축소판) */
  function renderPage(wb, s, lay, pi, scale) {
    const P = OA.print.page(scale, { landscape: s.page.orient === 'landscape' });
    const pg = lay.pages[pi];
    if (!pg || pg.empty) return P;
    const f = XL.PX_MM * scale * lay.scale;
    const pgs = s.page;
    let ox = pgs.left * 10 * scale, oy = pgs.top * 10 * scale;
    if (pgs.centerH) ox += (lay.avail.w - pg.w) / 2 * scale;
    if (pgs.centerV) oy += (lay.avail.h - pg.h) / 2 * scale;
    const content = h('div', { class: 'content', style: { left: ox + 'px', top: oy + 'px', width: pg.w * scale + 'px', height: pg.h * scale + 'px', overflow: 'hidden' } });
    const x0 = XL.colLeft(s, pg.c1), y0 = XL.rowTop(s, pg.r1);
    const covered = new Set();
    for (const m of s.merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(r + ',' + c);
    const pv = new Map();
    for (const p of s.pivots) { const g = PV.grid(wb, s, p); if (g) g.rows.forEach((row, i) => row.forEach((cell, j) => pv.set((g.r0 + i) + ',' + (g.c0 + j), { cell, pv: p }))); }
    let lines = '';
    for (let r = pg.r1; r <= pg.r2; r++) for (let c = pg.c1; c <= pg.c2; c++) {
      if (covered.has(r + ',' + c)) continue;
      const m = XL.mergeAt(s, r, c);
      const x = (XL.colLeft(s, c) - x0) * f, y = (XL.rowTop(s, r) - y0) * f;
      const w = (m ? XL.colLeft(s, m.c2 + 1) - XL.colLeft(s, c) : XL.colW(s, c)) * f, hh = (m ? XL.rowTop(s, m.r2 + 1) - XL.rowTop(s, r) : XL.rowH(s, r)) * f;
      const b = XL.borderAt(s, r, c);
      const mm = m || { r1: r, c1: c, r2: r, c2: c };
      const bw = k => Math.max(0.6, b[k] * f * 0.9);
      if (b.t) lines += `<line x1="${x}" y1="${y}" x2="${x + w}" y2="${y}" stroke="#000" stroke-width="${bw('t')}"/>`;
      if (b.l) lines += `<line x1="${x}" y1="${y}" x2="${x}" y2="${y + hh}" stroke="#000" stroke-width="${bw('l')}"/>`;
      if (b.b && (!m || r === mm.r1)) lines += `<line x1="${x}" y1="${y + hh}" x2="${x + w}" y2="${y + hh}" stroke="#000" stroke-width="${bw('b')}"/>`;
      if (b.r && (!m || c === mm.c1)) lines += `<line x1="${x + w}" y1="${y}" x2="${x + w}" y2="${y + hh}" stroke="#000" stroke-width="${bw('r')}"/>`;
      const pc = pv.get(r + ',' + c);
      const disp = XL.display(wb, s, r, c);
      if (!disp.text && !(pc && pc.cell.hdr)) continue;
      const st = pc ? {} : XL.styleOf(s, r, c);
      const d = h('div', { style: { position: 'absolute', left: x + 'px', top: y + 'px', width: w + 'px', height: hh + 'px', lineHeight: hh + 'px', fontSize: (st.sz || 11) * PX_PT * f + 'px', padding: '0 ' + 2 * f + 'px', whiteSpace: 'nowrap', overflow: 'hidden', fontFamily: 'Malgun Gothic, sans-serif',
        textAlign: st.ha || (pc && pc.cell.kind === 'label' ? 'left' : disp.align), fontWeight: st.b || (pc && pc.cell.bold) ? '700' : '400', background: pc && pc.cell.hdr ? '#ddebf7' : (st.fill || 'transparent') } });
      let txt = disp.text;
      if (pc && pc.cell.btn && pc.pv.printButtons && pc.pv.showButtons) txt = (pc.cell.btn === '-' ? '⊟ ' : '⊞ ') + txt;
      d.textContent = txt;
      const v = XL.getValue(wb, s, r, c);
      if (!m && typeof v === 'string' && !pc && (st.ha || 'left') === 'left') { d.style.overflow = 'visible'; }
      content.appendChild(d);
    }
    const svgB = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svgB.setAttribute('style', 'position:absolute;left:0;top:0;overflow:visible');
    svgB.setAttribute('width', pg.w * scale); svgB.setAttribute('height', pg.h * scale);
    svgB.innerHTML = lines;
    content.appendChild(svgB);
    for (const ch of s.charts) {
      const cx = (ch.x - x0) * f, cy = (ch.y - y0) * f;
      if (cx > pg.w * scale || cy > pg.h * scale || cx + ch.w * f < 0 || cy + ch.h * f < 0) continue;
      const box = h('div', { style: { position: 'absolute', left: cx + 'px', top: cy + 'px', width: ch.w * f + 'px', height: ch.h * f + 'px', border: '1px solid #d9d9d9' } });
      box.innerHTML = OA.chart.svg(wb, ch, { w: ch.w, h: ch.h });
      const sv = box.firstChild; sv.setAttribute('width', ch.w * f); sv.setAttribute('height', ch.h * f);
      content.appendChild(box);
    }
    P.appendChild(content);
    return P;
  }

  OA.xlRenderPage = renderPage;
  OA.apps = OA.apps || {};
  OA.apps.excel = { create: (win, ctx) => new ExcelApp(win, ctx) };
})(typeof window !== 'undefined' ? window : globalThis);

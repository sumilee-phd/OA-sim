/* Access 화면: 백스테이지(새 데이터베이스), 탐색 창, 문서 탭, 테이블 디자인·데이터시트, 관계, 쿼리 디자인·SQL·결과, 폼·보고서 디자인(OA.acDesign), 인쇄 미리 보기. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const { h, icon, menu, ribbon, dialog, msg, toast, tabs } = OA.ui;
  const U = OA.util, A = OA.acc, R = OA.report;
  const PX_CM = 37.8;

  function AccessApp(win, ctx) {
    const app = this;
    this.el = win;
    app.wizardOn = true;   // Access 기본값: 컨트롤 마법사 사용 켜짐(12.3 따라 하기 5에서 끈다)
    let db = ctx.model();
    const docs = [];
    let active = -1, navOpen = true, propOpen = false, backstage = null, navSel = null;
    const api = { db: () => db, changed: () => { ctx.changed(); }, render: () => render(), app };
    this.fileName = () => db ? db.name + ' : 데이터베이스 - Access' : 'Access';
    this.model = () => db;

    win.innerHTML = '';
    win.style.setProperty('--app', '#a4373a'); win.style.setProperty('--app2', '#8c2f32');
    const titleName = h('span', { class: 'name' });
    const qat = h('span', { class: 'qat' }, h('button', { title: '저장 (Ctrl+S)', onclick: () => saveActive() }, icon('save')), h('button', { title: '실행 취소 (Ctrl+Z)', onclick: () => undoDoc() }, icon('undo')), h('button', { title: '다시 실행 (Ctrl+Y)', disabled: true }, icon('redo')));
    win.appendChild(h('div', { class: 'titlebar' }, qat, titleName, h('span', { style: { width: '120px' } })));
    const ribHost = h('div');
    win.appendChild(ribHost);
    const navHost = h('div'), docTabs = h('div', { class: 'ac-doctabs' }), docHost = h('div', { class: 'ac-doc' }), sideHost = h('div', { style: { display: 'flex' } });
    const docsEl = h('div', { class: 'ac-docs' }, docTabs, docHost);
    const main = h('div', { class: 'ac-main' }, navHost, docsEl, sideHost);
    win.appendChild(main);
    const status = h('div', { class: 'statusbar' });
    win.appendChild(status);

    const cur = () => docs[active] || null;
    const curType = () => cur() ? cur().type : null;
    const isDesign = t => ['formDesign', 'reportDesign'].includes(curType()) && (!t || curType() === t);
    const out = (label, ic, type) => ({ type: type || 'small', label, icon: ic, out: true });

    // ------------------------------------------------------------------ 리본
    const viewMenu = () => {
      const d = cur();
      if (!d) return [];
      if (d.kind === 'table') return [{ label: '데이터시트 보기(H)', icon: 'datasheet', onClick: () => switchView(d, 'datasheet') }, { label: '디자인 보기(D)', icon: 'design', onClick: () => switchView(d, 'tableDesign') }];
      if (d.kind === 'query') return [{ label: '데이터시트 보기(H)', icon: 'datasheet', onClick: () => switchView(d, 'queryResult') }, { label: 'SQL 보기(Q)', icon: 'sql', onClick: () => switchView(d, 'querySQL') }, { label: '디자인 보기(D)', icon: 'design', onClick: () => switchView(d, 'queryDesign') }];
      if (d.kind === 'form') return [{ label: '폼 보기(F)', icon: 'formview', onClick: () => switchView(d, 'formView') }, { label: '레이아웃 보기(Y)', disabled: true }, { label: '디자인 보기(D)', icon: 'design', onClick: () => switchView(d, 'formDesign') }];
      if (d.kind === 'report') return [{ label: '보고서 보기(R)', icon: 'report', onClick: () => switchView(d, 'reportView') }, { label: '인쇄 미리 보기(V)', icon: 'preview', onClick: () => openPreview(d) }, { label: '레이아웃 보기(Y)', disabled: true }, { label: '디자인 보기(D)', icon: 'design', onClick: () => switchView(d, 'reportDesign') }];
      return [];
    };
    const viewBtn = () => ({ type: 'big', label: '보기', icon: 'view', onClick: () => { const d = cur(); if (!d) return; const m = viewMenu(); const next = m.find(x => !x.disabled && !(x.label.includes('디자인') && d.type.endsWith('Design'))); if (d.type.endsWith('Design')) m[0].onClick(); else (m.find(x => x.label.includes('디자인')) || next).onClick(); }, menu: viewMenu, disabled: () => !cur() || curType() === 'relations' || curType() === 'preview' });
    const ctlBtn = (kind, label, ic) => ({ type: 'icon', icon: ic, label, pressed: () => isDesign() && cur().tool === kind, onClick: () => { const d = cur(); if (d && d.design) d.design.setTool(kind); } });
    const controls = (report) => ({ label: '컨트롤', items: [{ type: 'stack', items: [
      { type: 'row', items: [ctlBtn(null, '선택', 'selectArrow'), ctlBtn('textbox', '텍스트 상자', 'textboxC'), ctlBtn('label', '레이블', 'label'), out('단추', 'button', 'icon'), out('탭 컨트롤', 'tab', 'icon'), out('하이퍼링크', 'selectdata', 'icon'), out('웹 브라우저 컨트롤', 'image', 'icon')] },
      { type: 'row', items: [out('탐색 컨트롤', 'fieldlist', 'icon'), out('옵션 그룹', 'group', 'icon'), out('페이지 나누기 삽입', 'breaks', 'icon'), out('콤보 상자', 'combo', 'icon'), out('차트', 'chartIcon', 'icon'), ctlBtn('line', '선', 'lineC'), out('토글 단추', 'button', 'icon')] },
      { type: 'row', items: [report ? out('목록 상자', 'listbox', 'icon') : ctlBtn('listbox', '목록 상자', 'listbox'), ctlBtn('rect', '사각형', 'rectC'), out('확인란', 'checkbox', 'icon'), out('언바운드 개체 틀', 'image', 'icon'), out('첨부 파일', 'paper', 'icon'), out('옵션 단추', 'dot', 'icon'), out('하위 폼/하위 보고서', 'form', 'icon')] }] },
      { type: 'stack', items: [{ type: 'small', label: '컨트롤 마법사 사용', icon: 'wand', pressed: () => !!app.wizardOn, onClick: () => { app.wizardOn = !app.wizardOn; rib.render(); } }, out('ActiveX 컨트롤', 'options')] }] });
    const tools = (report) => ({ label: '도구', items: [
      { type: 'big', label: '기존 필드 추가', icon: 'existing', onClick: () => { const d = cur(); if (d && d.design) d.design.toggleFieldList(); }, disabled: () => !isDesign() },
      { type: 'big', label: '속성 시트', icon: 'propsheet', pressed: () => propOpen, onClick: () => toggleProps() },
      out('탭 순서', 'fieldlist', 'big'), out(report ? '하위 보고서' : '새 창에서 하위 폼', 'form', 'big'), out('코드 보기', 'design', 'big')] });
    const rib = ribbon(ribHost, {
      file: () => openBackstage(db ? 'info' : 'new'),
      initial: 'home',
      tabs: [
        { id: 'home', label: '홈', when: () => curType() !== 'preview', groups: [
          { label: '보기', items: [viewBtn()] },
          { label: '클립보드', items: [out('붙여넣기', 'paste', 'big'), { type: 'stack', items: [out('잘라내기', 'cut'), out('복사', 'copy'), out('서식 복사', 'brush')] }] },
          { label: '정렬 및 필터', items: [out('필터', 'filter', 'big'), { type: 'stack', items: [out('오름차순', 'sortAZ'), out('내림차순', 'sortZA'), out('정렬 제거', 'clearLayout')] }] },
          { label: '레코드', items: [{ type: 'stack', items: [{ type: 'small', label: '모두 새로 고침', icon: 'refresh', onClick: () => render() }, out('새로 만들기', 'insertRow'), { type: 'small', label: '저장', icon: 'save', onClick: () => saveActive() }] },
            { type: 'stack', items: [{ type: 'small', label: '요약', icon: 'totals', disabled: () => curType() !== 'datasheet' && curType() !== 'queryResult', pressed: () => cur() && cur().totals, onClick: () => { const d = cur(); d.totals = !d.totals; render(); } }, out('맞춤법 검사', 'find'), out('삭제', 'deleteRow')] }] },
          { label: '찾기', items: [out('찾기', 'find', 'big')] }, { label: '텍스트 서식', items: [out('굵게', 'bold'), out('기울임꼴', 'italic')] }] },
        { id: 'create', label: '만들기', when: () => curType() !== 'preview', groups: [
          { label: '서식 파일', items: [out('응용 프로그램 요소', 'db', 'big')] },
          { label: '테이블', items: [{ type: 'big', label: '테이블', icon: 'tableA', onClick: () => need(() => newTableDatasheet()) }, { type: 'big', label: '테이블 디자인', icon: 'design', onClick: () => need(() => newTableDesign()) }, out('SharePoint 목록', 'tableA', 'big')] },
          { label: '쿼리', items: [out('쿼리 마법사', 'wand', 'big'), { type: 'big', label: '쿼리 디자인', icon: 'query', onClick: () => need(() => newQuery()) }] },
          { label: '폼', items: [out('폼', 'form', 'big'), { type: 'big', label: '폼 디자인', icon: 'design', onClick: () => need(() => newForm()) }, out('새 폼', 'form', 'big'), { type: 'stack', items: [out('폼 마법사', 'wand'), out('탐색', 'fieldlist'), out('기타 폼', 'form')] }] },
          { label: '보고서', items: [out('보고서', 'report', 'big'), { type: 'big', label: '보고서 디자인', icon: 'design', onClick: () => need(() => newReport()) }, out('새 보고서', 'report', 'big'), { type: 'stack', items: [{ type: 'small', label: '보고서 마법사', icon: 'wand', onClick: () => need(() => OA.acWizard(api, rpt => openDoc({ kind: 'report', name: null, obj: rpt, type: 'reportDesign', dirty: true, wizard: true }))) }, out('레이블', 'label')] }] },
          { label: '매크로 및 코드', items: [out('매크로', 'run', 'big'), out('모듈', 'design', 'big')] }] },
        { id: 'external', label: '외부 데이터', when: () => curType() !== 'preview', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '가져오기 및 연결', items: [out('새 데이터 원본', 'db', 'big')] }] },
        { id: 'dbtools', label: '데이터베이스 도구', when: () => curType() !== 'preview', groups: [
          { label: '도구', items: [out('데이터베이스 압축 및 복구', 'db', 'big')] }, { label: '매크로', items: [out('Visual Basic', 'design', 'big')] },
          { label: '관계', items: [{ type: 'big', label: '관계', icon: 'relations', onClick: () => need(() => openRelations()) }, out('개체 관계', 'relations', 'big')] },
          { label: '분석', items: [out('데이터베이스 문서화', 'report', 'big')] }] },
        { id: 'help', label: '도움말', when: () => curType() !== 'preview', note: '실기 과제에 쓰지 않는 탭입니다.', groups: [{ label: '도움말', items: [out('도움말', 'find', 'big')] }] },
        // 문맥 탭
        { id: 'tdesign', label: '테이블 디자인', ctx: '테이블 도구', when: () => curType() === 'tableDesign', groups: [
          { label: '보기', items: [viewBtn()] },
          { label: '도구', items: [{ type: 'big', label: '기본 키', icon: 'key', onClick: () => cur().td.primaryKey() }, out('작성기', 'builder', 'big'), out('유효성 검사 규칙 테스트', 'checkbox', 'big'),
            { type: 'stack', items: [{ type: 'small', label: '행 삽입', icon: 'insertRow', onClick: () => cur().td.insertRow() }, { type: 'small', label: '행 삭제', icon: 'deleteRow', onClick: () => cur().td.deleteRow() }, out('조회 수정', 'design')] }] },
          { label: '표시/숨기기', items: [out('속성 시트', 'propsheet', 'big'), out('인덱스', 'key', 'big')] },
          { label: '필드, 레코드 및 테이블 이벤트', items: [out('데이터 매크로 만들기', 'run', 'big')] },
          { label: '관계', items: [{ type: 'big', label: '관계', icon: 'relations', onClick: () => openRelations() }, out('개체 종속성', 'relations', 'big')] }] },
        { id: 'reldesign', label: '관계 디자인', ctx: '관계 도구', when: () => curType() === 'relations', groups: [
          { label: '도구', items: [{ type: 'big', label: '관계 편집', icon: 'relations', onClick: () => cur().rel.editSelected() }, { type: 'stack', items: [{ type: 'small', label: '레이아웃 지우기', icon: 'clearLayout', onClick: () => cur().rel.clearLayout() }, out('관계 보고서', 'report')] }] },
          { label: '관계', items: [{ type: 'big', label: '테이블 추가', icon: 'addTable', onClick: () => cur().rel.showAdd() }, { type: 'stack', items: [out('직접 관계 표시', 'relations'), { type: 'small', label: '모든 관계', icon: 'relations', onClick: () => cur().rel.showAll() }] }, { type: 'big', label: '닫기', icon: 'closeView', onClick: () => closeDoc(active) }] }] },
        { id: 'qdesign', label: '쿼리 디자인', ctx: '쿼리 도구', when: () => ['queryDesign', 'querySQL', 'queryResult'].includes(curType()), groups: [
          { label: '결과', items: [viewBtn(), { type: 'big', label: '실행', icon: 'run', onClick: () => runQuery(cur()) }] },
          { label: '쿼리 유형', items: [{ type: 'big', label: '선택', icon: 'query', pressed: true, onClick: () => {} }, out('테이블 만들기', 'tableA', 'big'), out('추가', 'insertRow', 'big'), out('업데이트', 'design', 'big'), out('크로스탭', 'datasheet', 'big'), out('삭제', 'deleteRow', 'big')] },
          { label: '쿼리 설정', items: [{ type: 'big', label: '테이블 추가', icon: 'addTable', disabled: () => curType() !== 'queryDesign', onClick: () => cur().qd.showAdd() }, { type: 'stack', items: [{ type: 'small', label: '행 삽입', icon: 'insertRow', disabled: true }, { type: 'small', label: '열 삽입', icon: 'insertCells', disabled: () => curType() !== 'queryDesign', onClick: () => cur().qd.insertCol() }, { type: 'small', label: '열 삭제', icon: 'deleteCells', disabled: () => curType() !== 'queryDesign', onClick: () => cur().qd.deleteCol() }] }, out('작성기', 'builder', 'big')] },
          { label: '표시/숨기기', items: [out('요약', 'totals', 'big'), out('매개 변수', 'options', 'big'), out('속성 시트', 'propsheet', 'big'), out('테이블 이름', 'tableA', 'big')] }] },
        { id: 'fdesign', label: '폼 디자인', ctx: '폼 디자인 도구', when: () => curType() === 'formDesign', groups: [{ label: '보기', items: [viewBtn()] }, { label: '테마', items: [out('테마', 'colors', 'big')] }, controls(false), { label: '머리글/바닥글', items: [out('로고', 'logo'), out('제목', 'titleP'), out('날짜 및 시간', 'dateTime')] }, tools(false)] },
        { id: 'farrange', label: '정렬', ctx: '폼 디자인 도구', when: () => curType() === 'formDesign', note: '이 시뮬레이터는 컨트롤 위치·크기를 속성 시트와 화살표 키로 정합니다.', groups: [{ label: '표', items: [out('눈금선', 'gridlines', 'big')] }] },
        { id: 'fformat', label: '서식', ctx: '폼 디자인 도구', when: () => curType() === 'formDesign', note: '글꼴·테두리는 속성 시트 형식 탭에서 정합니다.', groups: [{ label: '글꼴', items: [out('굵게', 'bold')] }] },
        { id: 'rdesign', label: '보고서 디자인', ctx: '보고서 디자인 도구', when: () => curType() === 'reportDesign', groups: [{ label: '보기', items: [viewBtn()] }, { label: '테마', items: [out('테마', 'colors', 'big')] },
          { label: '그룹화 및 요약', items: [{ type: 'big', label: '그룹화 및 정렬', icon: 'grouping', pressed: () => cur() && cur().design && cur().design.groupingOpen(), onClick: () => cur().design.toggleGrouping() }, { type: 'stack', items: [out('요약', 'totals'), out('세부 정보 숨기기', 'collapse')] }] },
          controls(true), { label: '머리글/바닥글', items: [out('페이지 번호', 'pageNum'), out('로고', 'logo'), out('제목', 'titleP'), out('날짜 및 시간', 'dateTime')] }, tools(true)] },
        { id: 'rarrange', label: '정렬', ctx: '보고서 디자인 도구', when: () => curType() === 'reportDesign', note: '컨트롤 위치·크기는 속성 시트와 화살표 키로 정합니다.', groups: [{ label: '표', items: [out('눈금선', 'gridlines', 'big')] }] },
        { id: 'rformat', label: '서식', ctx: '보고서 디자인 도구', when: () => curType() === 'reportDesign', note: '글꼴·형식은 속성 시트에서 정합니다.', groups: [{ label: '글꼴', items: [out('굵게', 'bold')] }] },
        { id: 'rpage', label: '페이지 설정', ctx: '보고서 디자인 도구', when: () => curType() === 'reportDesign', groups: [{ label: '페이지 레이아웃', items: [{ type: 'big', label: '페이지 설정', icon: 'pagesetup', onClick: () => pageSetup(cur()) }] }] },
        { id: 'preview', label: '인쇄 미리 보기', when: () => curType() === 'preview', groups: [
          { label: '인쇄', items: [{ type: 'big', label: '인쇄', icon: 'printer', onClick: () => printDoc(cur()) }] },
          { label: '페이지 크기', items: [{ type: 'big', label: '크기', icon: 'paper', menu: () => ['A4', 'A3', 'B5'].map(p => ({ label: p, checked: pv().page.paper === p, onClick: () => setPage({ paper: p }) })) },
            { type: 'big', label: '여백', icon: 'margins', menu: () => [{ label: '보통  위쪽 19.05mm 아래쪽 19.05mm 왼쪽 19.05mm 오른쪽 19.05mm', onClick: () => setPage({ top: 19.05, bottom: 19.05, left: 19.05, right: 19.05 }) }, { label: '넓게  위쪽 25.4mm 아래쪽 25.4mm 왼쪽 19.05mm 오른쪽 19.05mm', onClick: () => setPage({ top: 25.4, bottom: 25.4, left: 19.05, right: 19.05 }) }, { label: '좁게  위쪽 6.35mm 아래쪽 6.35mm 왼쪽 6.35mm 오른쪽 6.35mm', onClick: () => setPage({ top: 6.35, bottom: 6.35, left: 6.35, right: 6.35 }) }] },
            { type: 'check', label: '데이터만 인쇄', checked: () => pv() && pv().page.dataOnly, onChange: v => setPage({ dataOnly: v }) }] },
          { label: '페이지 레이아웃', items: [{ type: 'big', label: '세로', icon: 'paper', pressed: () => pv() && pv().page.orient === 'portrait', onClick: () => setPage({ orient: 'portrait' }) }, { type: 'big', label: '가로', icon: 'orient', pressed: () => pv() && pv().page.orient === 'landscape', onClick: () => setPage({ orient: 'landscape' }) }, out('열', 'twopage', 'big'), { type: 'big', label: '페이지 설정', icon: 'pagesetup', onClick: () => pageSetup(cur()) }] },
          { label: '확대/축소', items: [{ type: 'big', label: '확대/축소', icon: 'zoom', menu: () => [200, 150, 100, 75, 50].map(z => ({ label: z + '%', onClick: () => { cur().zoom = z / 100; render(); } })).concat([{ label: '페이지에 맞춤', onClick: () => { cur().zoom = null; render(); } }]) }, { type: 'big', label: '한 페이지', icon: 'onepage', onClick: () => { cur().two = false; cur().zoom = null; render(); } }, { type: 'big', label: '두 페이지', icon: 'twopage', onClick: () => { cur().two = true; render(); } }, out('기타 페이지', 'twopage', 'big')] },
          { label: '데이터', items: [out('모두 새로 고침', 'refresh', 'big'), out('PDF 또는 XPS', 'paper', 'big')] },
          { label: '미리 보기 닫기', items: [{ type: 'big', label: '인쇄 미리 보기 닫기', icon: 'closeView', onClick: () => closeDoc(active) }] }] },
      ],
    });
    function need(fn) { if (!db) return msg({ title: 'Microsoft Access', text: '먼저 파일 탭 > 새로 만들기 > 새 데이터베이스로 데이터베이스를 만드십시오.', icon: 'info' }); fn(); }
    const pv = () => cur() && cur().type === 'preview' ? cur().src.obj : null;
    function setPage(patch) { const d = cur(); if (!d || d.type !== 'preview') return; Object.assign(d.src.obj.page, patch); applyPageToSaved(d.src); render(); }
    function applyPageToSaved(doc) {
      const store = doc.kind === 'form' ? db.forms : db.reports;
      if (doc.name && store[doc.name]) { store[doc.name].page = U.clone(doc.obj.page); ctx.changed(); }
      else doc.dirty = true;
    }

    // ------------------------------------------------------------------ 문서 관리
    function openDoc(d) {
      d.id = d.id || U.uid('doc');
      docs.push(d);
      active = docs.length - 1;
      if (d.type === 'formDesign' || d.type === 'reportDesign') propOpen = propOpen || false;
      render();
      return d;
    }
    function findDoc(kind, name) { return docs.findIndex(d => d.kind === kind && d.name === name && d.type !== 'preview'); }
    function closeDoc(i, force) {
      const d = docs[i];
      if (!d) return Promise.resolve(true);
      const finish = () => { docs.splice(i, 1); if (active >= docs.length) active = docs.length - 1; render(); return true; };
      if (d.type === 'relations' && d.layoutDirty && !force) {
        return msg({ title: 'Microsoft Access', text: "'관계' 레이아웃의 변경 내용을 저장하시겠습니까?", buttons: ['예', '아니요', '취소'], icon: 'q' }).then(r => { if (r === '취소') return false; if (r === '예') { db.relLayout = U.clone(d.layout); ctx.changed(); } return finish(); });
      }
      if (d.dirty && !force && d.type !== 'preview') {
        const label = { table: '테이블', query: '쿼리', form: '폼', report: '보고서' }[d.kind];
        const nm = d.name || defaultName(d.kind);
        return msg({ title: 'Microsoft Access', text: `'${nm}' ${label}의 디자인에 대한 변경 내용을 저장하시겠습니까?`, buttons: ['예', '아니요', '취소'], icon: 'q' }).then(r => { if (r === '취소') return false; if (r === '예') return saveDoc(d).then(ok => ok ? finish() : false); return finish(); });
      }
      return Promise.resolve(finish());
    }
    function defaultName(kind) {
      const base = { table: '테이블', query: '쿼리', form: '폼', report: '보고서' }[kind];
      const store = { table: db.tables, query: db.queries, form: db.forms, report: db.reports }[kind];
      let n = 1; while (store[base + n] || docs.some(d => d.kind === kind && d.name === base + n)) n++;
      return base + n;
    }
    function askName(kind, def) {
      return new Promise(res => {
        const label = { table: '테이블 이름', query: '쿼리 이름', form: '폼 이름', report: '보고서 이름' }[kind];
        const inp = h('input', { type: 'text', value: def, style: { width: '260px' } });
        dialog({ title: '다른 이름으로 저장', width: 320, help: false, body: h('div', { class: 'col' }, label + '(N):', inp), buttons: [{ label: '확인', primary: true, onClick: () => { const v = inp.value.trim(); if (!v || /[.!`[\]]/.test(v)) { msg({ title: 'Microsoft Access', text: '개체 이름이 잘못되었습니다.', icon: 'warn' }); return false; } res(v); } }, { label: '취소', cancel: true }], onClose: r => { if (r === null) res(null); } });
      });
    }
    function saveDoc(d) {
      if (!d) return Promise.resolve(false);
      if (d.type === 'relations') { db.relLayout = U.clone(d.layout); d.layoutDirty = false; ctx.changed(); return Promise.resolve(true); }
      if (d.type === 'preview') return saveDoc(d.src);
      const go = name => {
        if (!name) return false;
        if (d.kind === 'table') { const r = d.td.commit(name); if (!r) return false; }
        else if (d.kind === 'query') { d.qd && d.qd.sync(); A.saveQuery(db, name, d.obj); }
        else if (d.kind === 'form') A.saveForm(db, name, d.obj);
        else if (d.kind === 'report') A.saveReport(db, name, d.obj);
        d.name = name; d.dirty = false; d.isNew = false;
        ctx.changed(); ctx.saved();
        render();
        return true;
      };
      if (d.kind === 'table' && d.type === 'tableDesign' && !d.obj.pk.length && !d.pkAsked) {
        return (d.name ? Promise.resolve(d.name) : askName('table', defaultName('table'))).then(name => {
          if (!name) return false;
          return msg({ title: 'Microsoft Access', text: '기본 키가 정의되어 있지 않습니다.\n\n기본 키가 꼭 필요한 것은 아니지만 사용하는 것이 좋습니다. 데이터베이스의 다른 테이블과 관계를 정의하려면 테이블에 기본 키가 있어야 합니다.\n\n지금 기본 키를 만드시겠습니까?', buttons: ['예', '아니요', '취소'], icon: 'warn' }).then(r => {
            if (r === '취소') return false;
            if (r === '예') d.td.addIdKey();
            d.pkAsked = true;
            return go(name);
          });
        });
      }
      if (d.name && !d.isNew) return Promise.resolve(go(d.name));
      return askName(d.kind, d.name || defaultName(d.kind)).then(go);
    }
    function saveActive() { if (!db) return; const d = cur(); if (!d) { ctx.saved(); return; } if (d.design) d.design.commitEdit(); saveDoc(d); }
    function undoDoc() { const d = cur(); if (d && d.undo && d.undo.length) { d.obj = JSON.parse(d.undo.pop()); d.dirty = true; if (d.design) d.design.reset(d.obj); render(); } }
    api.snapshot = d => { d.undo = d.undo || []; d.undo.push(JSON.stringify(d.obj)); if (d.undo.length > 60) d.undo.shift(); d.dirty = true; };
    function switchView(d, type) {
      if (d.design) d.design.commitEdit();
      if (d.kind === 'table') {
        if (type === 'datasheet' && (d.dirty || d.isNew)) {
          return msg({ title: 'Microsoft Access', text: '테이블을 먼저 저장해야 합니다.\n\n지금 저장하시겠습니까?', buttons: ['예', '아니요'], icon: 'q' }).then(r => { if (r === '예') saveDoc(d).then(ok => { if (ok) { d.type = 'datasheet'; d.obj = U.clone(db.tables[d.name]); render(); } }); });
        }
        if (type === 'tableDesign') { d.obj = U.clone(db.tables[d.name]); d.td = null; }
        d.type = type; render(); return;
      }
      if (d.kind === 'query') {
        if (d.qd) d.qd.sync();
        if (type === 'queryDesign' && d.type === 'querySQL') {
          const q = A.sqlToDesign(db, d.obj.sql || '');
          if (!q) return msg({ title: 'Microsoft Access', text: 'SQL 문의 구문 오류입니다. SQL 문을 확인하십시오.', icon: 'warn' });
          Object.assign(d.obj, q, { mode: q.partial ? 'sql' : 'design', sql: d.obj.sql });
          d.qd = null;
        }
        if (type === 'querySQL' && d.obj.mode !== 'sql') { try { d.obj.sql = A.designToSQL(db, d.obj); } catch (e) { d.obj.sql = 'SELECT;'; } }
        if (type === 'queryResult') return runQuery(d);
        d.type = type; render(); return;
      }
      if (d.kind === 'form' || d.kind === 'report') { d.type = type; d.design = null; render(); }
    }
    function runQuery(d, params) {
      if (!d) return;
      if (d.qd) d.qd.sync();
      let sql;
      try { sql = d.obj.mode === 'sql' ? d.obj.sql : A.designToSQL(db, d.obj); } catch (e) { return msg({ title: 'Microsoft Access', text: e.message, icon: 'warn' }); }
      let r;
      try { r = A.runSQL(db, sql, params); } catch (e) { return msg({ title: 'Microsoft Access', text: e.message, icon: 'warn' }); }
      if (r.needParam) {
        const inp = h('input', { type: 'text', style: { width: '260px' } });
        return dialog({ title: '매개 변수 값 입력', width: 320, help: false, body: h('div', { class: 'col' }, r.needParam, inp), buttons: [{ label: '확인', primary: true, onClick: () => { setTimeout(() => runQuery(d, Object.assign({}, params, { [r.needParam]: inp.value })), 0); } }, { label: '취소', cancel: true }] });
      }
      d.result = r; d.type = 'queryResult';
      render();
    }

    // ------------------------------------------------------------------ 새 개체
    function newTableDatasheet() {
      const t = { name: null, fields: [Object.assign(A.defaultField('ID', '일련 번호'), { indexed: '예(중복 불가능)', required: true })], pk: ['ID'], rows: [] };
      return openDoc({ kind: 'table', name: null, obj: t, type: 'tableDesign', isNew: true, dirty: false, fromBlank: true });
    }
    function newTableDesign() { return openDoc({ kind: 'table', name: null, obj: { name: null, fields: [], pk: [], rows: [] }, type: 'tableDesign', isNew: true, dirty: false }); }
    function newQuery() { const d = openDoc({ kind: 'query', name: null, obj: A.newQueryDesign(db, []), type: 'queryDesign', isNew: true, dirty: false }); setTimeout(() => d.qd && d.qd.showAdd(), 0); }
    function newForm() { const f = A.newForm(db); return openDoc({ kind: 'form', name: null, obj: f, type: 'formDesign', isNew: true, dirty: false }); }
    function newReport() { const r = A.newReport(db, ''); r.sections.pageHeader.height = 0.8; return openDoc({ kind: 'report', name: null, obj: r, type: 'reportDesign', isNew: true, dirty: false }); }
    function openObject(kind, name, view) {
      const i = findDoc(kind, name);
      if (i >= 0) { active = i; if (view && docs[i].type !== view) switchView(docs[i], view); else render(); return; }
      if (kind === 'table') openDoc({ kind, name, obj: U.clone(db.tables[name]), type: view || 'datasheet' });
      if (kind === 'query') { const d = openDoc({ kind, name, obj: U.clone(db.queries[name]), type: view || 'queryResult' }); if (!view || view === 'queryResult') runQuery(d); }
      if (kind === 'form') openDoc({ kind, name, obj: U.clone(db.forms[name]), type: view || 'formView' });
      if (kind === 'report') { const d = openDoc({ kind, name, obj: U.clone(db.reports[name]), type: view === 'preview' ? 'reportView' : (view || 'reportView') }); if (view === 'preview') openPreview(d); }
    }
    function openRelations() {
      const i = docs.findIndex(d => d.type === 'relations');
      if (i >= 0) { active = i; render(); return; }
      const layout = db.relLayout ? U.clone(db.relLayout) : { tables: [] };
      const d = openDoc({ kind: 'rel', name: '관계', type: 'relations', layout });
      if (!layout.tables.length) setTimeout(() => d.rel && d.rel.showAdd(), 0);
    }
    function openPreview(src) {
      if (src.design) src.design.commitEdit();
      const i = docs.findIndex(d => d.type === 'preview' && d.src === src);
      if (i >= 0) { active = i; render(); return; }
      openDoc({ kind: 'preview', name: src.name || defaultName(src.kind), type: 'preview', src, zoom: null, page: 0 });
    }
    function printDoc(d) {
      const src = d.src;
      const key = src.kind === 'form' ? '5-3' : '5-4';
      if (src.dirty || src.isNew) toast('저장하지 않은 디자인입니다. 저장한 개체와 출력물이 같아야 합니다(9.1절).');
      ctx.printed(key, { name: src.name });
    }
    api.openObject = openObject; api.openDoc = openDoc; api.openPreview = openPreview; api.cur = cur; api.rib = () => rib;

    function pageSetup(d) {
      if (!d) return;
      const src = d.type === 'preview' ? d.src : d;
      if (!(src.kind === 'form' || src.kind === 'report')) return;
      OA.acPageSetup(api, src, () => { applyPageToSaved(src); render(); });
    }
    // ------------------------------------------------------------------ 속성 시트
    function toggleProps() { propOpen = !propOpen; render(); }
    api.showProps = (on) => { propOpen = on === undefined ? true : on; render(); };

    // ------------------------------------------------------------------ 백스테이지
    function openBackstage(page) {
      closeBackstage();
      backstage = h('div', { class: 'backstage' });
      const nav = h('div', { class: 'bs-nav' });
      const mainEl = h('div', { class: 'bs-main' });
      const items = [['back', '←'], ['home', '홈'], ['new', '새로 만들기'], ['open', '열기'], ['-'], ['info', '정보'], ['save', '저장'], ['saveas', '다른 이름으로 저장'], ['print', '인쇄'], ['close', '닫기']];
      const show = p => {
        [...nav.children].forEach(b => b.classList.toggle('on', b.dataset.p === p));
        mainEl.innerHTML = '';
        if (p === 'new' || p === 'home') mainEl.appendChild(newPage());
        else if (p === 'info') mainEl.appendChild(h('div', {}, h('h2', {}, '정보'), db ? h('p', {}, `${db.folder}\\${db.name}.accdb`) : h('p', {}, '열린 데이터베이스가 없습니다.')));
        else if (p === 'print') mainEl.appendChild(h('div', {}, h('h2', {}, '인쇄'),
          h('div', { class: 'pset', onclick: () => { const d = printable(); if (d) { closeBackstage(); ctx.printed(d.kind === 'form' ? '5-3' : '5-4', { name: d.name }); } } }, icon('printer'), h('div', { class: 'pt' }, '빠른 인쇄', h('small', {}, '변경 없이 기본 프린터로 바로 보냅니다.'))),
          h('div', { class: 'pset', onclick: () => { const d = printable(); if (d) { closeBackstage(); openPreview(d); } } }, icon('printer'), h('div', { class: 'pt' }, '인쇄', h('small', {}, '인쇄하기 전에 프린터, 인쇄 매수 및 기타 인쇄 옵션을 선택합니다.'))),
          h('div', { class: 'pset', onclick: () => { const d = printable(); if (d) { closeBackstage(); openPreview(d); } } }, icon('preview'), h('div', { class: 'pt' }, '인쇄 미리 보기', h('small', {}, '인쇄하기 전에 페이지를 미리 보고 변경합니다.')))));
        else mainEl.appendChild(h('div', {}, h('h2', {}, items.find(x => x[0] === p)[1]), h('p', { class: 'note' }, '실기 과제에 쓰지 않는 화면입니다.')));
      };
      for (const [p, l] of items) {
        if (p === '-') { nav.appendChild(h('div', { style: { height: '1px', background: 'rgba(255,255,255,.3)', margin: '6px 14px' } })); continue; }
        const b = h('button', { class: p === 'back' ? 'back' : '', 'data-p': p, disabled: ['open'].includes(p) || (!db && ['save', 'saveas', 'print', 'close', 'back'].includes(p)) }, l);
        b.addEventListener('click', () => { if (p === 'back') return closeBackstage(); if (p === 'save') { closeBackstage(); return saveActive(); } if (p === 'saveas') { closeBackstage(); const d = cur(); if (d) { d.isNew = true; saveDoc(d); } return; } if (p === 'close') { closeBackstage(); return; } show(p); });
        nav.appendChild(b);
      }
      backstage.append(nav, mainEl);
      win.appendChild(backstage);
      show(page);
    }
    function printable() {
      const d = cur();
      const src = d && (d.type === 'preview' ? d.src : d);
      if (!src || !(src.kind === 'form' || src.kind === 'report')) { msg({ title: 'Microsoft Access', text: '인쇄할 폼이나 보고서를 먼저 엽니다(폼 보기 또는 보고서).', icon: 'info' }); return null; }
      return src;
    }
    function closeBackstage() { if (backstage) { backstage.remove(); backstage = null; render(); } }
    function newPage() {
      const nameIn = h('input', { type: 'text', value: (ctx.set.id === 'equip' ? 'ch12_작업' : ctx.set.id + '_작업') , style: { width: '220px' } });
      const folder = h('span', { class: 'note' }, '바탕 화면\\비번호 폴더');
      const panel = h('div', { style: { border: '1px solid #ddd', padding: '14px', width: '330px', marginTop: '14px', display: 'none' } },
        h('div', { style: { fontSize: '15px', marginBottom: '8px' } }, '새 데이터베이스'), '파일 이름', h('div', { class: 'row' }, nameIn, h('button', { class: 'btn', title: '데이터베이스를 넣을 위치 찾아보기', onclick: () => dialog({ title: '새 데이터베이스 파일', width: 420, body: h('div', {}, h('div', { class: 'row' }, icon('folder'), '바탕 화면 › 비번호 폴더'), h('p', { class: 'note' }, '시험장에서는 바탕 화면의 비번호 폴더를 고릅니다(9.1절).')), buttons: [{ label: '확인', primary: true }, { label: '취소', cancel: true }] }) }, icon('folder'))), folder,
        h('div', { style: { marginTop: '12px' } }, h('button', { class: 'btn primary', onclick: () => createDb(nameIn.value.trim() || 'Database1') }, '만들기')));
      return h('div', {}, h('h2', {}, '새로 만들기'),
        h('div', { class: 'row', style: { gap: '14px' } },
          h('button', { class: 'btn', style: { width: '150px', height: '120px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: '8px' }, onclick: () => { panel.style.display = 'block'; nameIn.focus(); nameIn.select(); } }, icon('db', 44), '새 데이터베이스'),
          h('button', { class: 'btn', disabled: true, style: { width: '150px', height: '120px' } }, '새 웹 앱')),
        panel, db ? h('p', { class: 'note' }, `현재 열린 데이터베이스: ${db.name}. 새로 만들면 이 데이터베이스를 닫습니다.`) : null);
    }
    function createDb(name) {
      const go = () => {
        db = A.newDatabase(name);
        db.folder = '바탕 화면\\비번호 폴더';
        docs.length = 0; active = -1;
        ctx.setModel(db);
        closeBackstage();
        // 자동으로 열리는 테이블1(데이터시트)
        const t = { name: null, fields: [Object.assign(A.defaultField('ID', '일련 번호'), { indexed: '예(중복 불가능)', required: true })], pk: ['ID'], rows: [] };
        openDoc({ kind: 'table', name: null, obj: t, type: 'datasheet', isNew: true, dirty: false, blankTable: true, tabName: '테이블1' });
      };
      if (db && Object.keys(db.tables).length) msg({ title: 'Microsoft Access', text: '현재 데이터베이스를 닫고 새 데이터베이스를 만듭니다. 이 시뮬레이터에서는 이전 데이터베이스가 지워집니다.', buttons: ['만들기', '취소'] }).then(r => { if (r === '만들기') go(); });
      else go();
    }

    // ------------------------------------------------------------------ 탐색 창
    function renderNav() {
      navHost.innerHTML = '';
      if (!navOpen) { navHost.appendChild(h('div', { class: 'ac-nav-collapsed', title: '탐색 창 (F11)', onclick: () => { navOpen = true; render(); } }, '탐색 창 »')); return; }
      const el = h('div', { class: 'ac-nav' }, h('div', { class: 'ac-nav-h' }, h('span', { class: 't' }, '모든 Access 개체'), h('button', { title: '탐색 창 닫기 (F11)', onclick: () => { navOpen = false; render(); } }, '«')));
      const search = h('input', { class: 'search', placeholder: '검색...', onkeydown: e => e.stopPropagation(), oninput: () => fill() });
      el.appendChild(search);
      const list = h('div', { style: { overflow: 'auto', flex: 1 } });
      el.appendChild(list);
      const fill = () => {
        list.innerHTML = '';
        if (!db) { list.appendChild(h('div', { class: 'note', style: { padding: '10px' } }, '데이터베이스가 없습니다. 파일 탭에서 새 데이터베이스를 만듭니다.')); return; }
        const groups = [['table', '테이블', db.tables, 'tableA'], ['query', '쿼리', db.queries, 'query'], ['form', '폼', db.forms, 'form'], ['report', '보고서', db.reports, 'report']];
        for (const [kind, label, store, ic] of groups) {
          const names = Object.keys(store).filter(n => !search.value || n.includes(search.value));
          if (!names.length) continue;
          list.appendChild(h('div', { class: 'grp' }, label, ' ⌃'));
          for (const n of names) {
            const o = h('div', { class: 'obj' + (navSel === kind + ':' + n ? ' on' : '') }, icon(ic), n);
            o.addEventListener('click', () => { navSel = kind + ':' + n; fill(); });
            o.addEventListener('dblclick', () => openObject(kind, n));
            o.addEventListener('contextmenu', e => {
              e.preventDefault(); navSel = kind + ':' + n; fill();
              const items = [{ label: '열기(O)', icon: 'view', onClick: () => openObject(kind, n) }, { label: '디자인 보기(D)', icon: 'design', onClick: () => openObject(kind, n, { table: 'tableDesign', query: 'queryDesign', form: 'formDesign', report: 'reportDesign' }[kind]) }];
              if (kind === 'report') items.push({ label: '인쇄 미리 보기(V)', icon: 'preview', onClick: () => openObject(kind, n, 'preview') });
              if (kind === 'form') items.push({ label: '인쇄 미리 보기(V)', icon: 'preview', onClick: () => { openObject(kind, n, 'formView'); openPreview(cur()); } });
              items.push({ sep: true }, { label: '이름 바꾸기(M)', onClick: () => renameObj(kind, n) }, { label: '삭제(L)', onClick: () => deleteObj(kind, n) });
              menu(o, items, { at: { x: e.clientX, y: e.clientY } });
            });
            list.appendChild(o);
          }
        }
      };
      fill();
      navHost.appendChild(el);
    }
    function storeOf(kind) { return { table: db.tables, query: db.queries, form: db.forms, report: db.reports }[kind]; }
    function renameObj(kind, n) {
      if (findDoc(kind, n) >= 0) return msg({ title: 'Microsoft Access', text: '열려 있는 개체는 이름을 바꿀 수 없습니다. 개체를 닫은 뒤 다시 하십시오.', icon: 'warn' });
      const inp = h('input', { type: 'text', value: n });
      dialog({ title: '이름 바꾸기', width: 300, body: inp, buttons: [{ label: '확인', primary: true, onClick: () => { const v = inp.value.trim(); if (!v || storeOf(kind)[v]) return false; const s = storeOf(kind); s[v] = s[n]; s[v].name = v; delete s[n]; if (kind === 'table') { db.relationships.forEach(r => { if (r.t1 === n) r.t1 = v; if (r.t2 === n) r.t2 = v; }); } ctx.changed(); render(); } }, { label: '취소', cancel: true }] });
    }
    function deleteObj(kind, n) {
      if (findDoc(kind, n) >= 0) return msg({ title: 'Microsoft Access', text: '열려 있는 개체는 삭제할 수 없습니다.', icon: 'warn' });
      msg({ title: 'Microsoft Access', text: `'${n}'을(를) 삭제하시겠습니까?`, buttons: ['예', '아니요'], icon: 'q' }).then(r => {
        if (r !== '예') return;
        if (kind === 'table' && db.relationships.some(x => x.t1 === n || x.t2 === n)) return msg({ title: 'Microsoft Access', text: `'${n}' 테이블은 하나 이상의 관계에 참여하고 있으므로 삭제할 수 없습니다. 먼저 관계 창에서 관계를 삭제하십시오.`, icon: 'warn' });
        delete storeOf(kind)[n]; ctx.changed(); render();
      });
    }

    // ------------------------------------------------------------------ 그리기
    function render() {
      titleName.textContent = app.fileName();
      rib.render();
      renderNav();
      docTabs.innerHTML = '';
      docs.forEach((d, i) => {
        const ic = { table: 'tableA', query: 'query', form: 'form', report: 'report', rel: 'relations', preview: 'preview' }[d.kind];
        const label = d.type === 'relations' ? '관계' : (d.name || d.tabName || defaultName(d.kind)) + (d.dirty ? ' *' : '');
        const t = h('div', { class: 'ac-doctab' + (i === active ? ' on' : '') }, icon(ic), h('span', {}, label), h('button', { title: '닫기', onclick: e => { e.stopPropagation(); closeDoc(i); } }, '✕'));
        t.addEventListener('click', () => { active = i; render(); });
        docTabs.appendChild(t);
      });
      docHost.innerHTML = '';
      sideHost.innerHTML = '';
      const d = cur();
      if (!db) docHost.appendChild(h('div', { class: 'ac-empty' }, h('p', {}, '열린 데이터베이스가 없습니다.'), h('button', { class: 'btn primary', onclick: () => openBackstage('new') }, '파일 탭 > 새로 만들기')));
      else if (!d) docHost.appendChild(h('div', { class: 'ac-empty' }, '탐색 창에서 개체를 두 번 누르거나 만들기 탭에서 새 개체를 만듭니다.'));
      else {
        try {
          if (d.type === 'tableDesign') docHost.appendChild(OA.acTable.design(api, d));
          else if (d.type === 'datasheet') docHost.appendChild(OA.acTable.datasheet(api, d));
          else if (d.type === 'relations') docHost.appendChild(OA.acRel.view(api, d, sideHost));
          else if (d.type === 'queryDesign') docHost.appendChild(OA.acQuery.design(api, d, sideHost));
          else if (d.type === 'querySQL') docHost.appendChild(OA.acQuery.sql(api, d));
          else if (d.type === 'queryResult') docHost.appendChild(OA.acQuery.result(api, d));
          else if (d.type === 'formDesign' || d.type === 'reportDesign') docHost.appendChild(OA.acDesign.view(api, d, sideHost, propOpen));
          else if (d.type === 'formView') docHost.appendChild(OA.acDesign.formView(api, d));
          else if (d.type === 'reportView') docHost.appendChild(OA.acDesign.reportView(api, d));
          else if (d.type === 'preview') docHost.appendChild(OA.acDesign.preview(api, d));
        } catch (e) { console.error(e); docHost.appendChild(h('div', { class: 'ac-empty' }, '화면 오류: ' + e.message)); }
      }
      status.innerHTML = '';
      status.append(h('span', {}, d ? ({ tableDesign: '디자인 보기. F6 = 창 전환. F1 = 도움말.', datasheet: '데이터시트 보기', relations: '준비', queryDesign: '준비', querySQL: 'SQL 보기', queryResult: '데이터시트 보기', formDesign: '디자인 보기', reportDesign: '디자인 보기', formView: '폼 보기', reportView: '보고서 보기', preview: '인쇄 미리 보기' }[d.type] || '준비') : '준비'), h('span', { class: 'sp' }), h('span', {}, db ? db.name + '.accdb' : ''));
      if (!d || (d.type !== 'preview' && rib.active === 'preview')) { /* 탭 유지 */ }
      if (d) {
        const want = { tableDesign: 'tdesign', relations: 'reldesign', queryDesign: 'qdesign', querySQL: 'qdesign', queryResult: 'qdesign', formDesign: 'fdesign', reportDesign: 'rdesign', preview: 'preview' }[d.type];
        if (want && d._lastTabFor !== d.type) { d._lastTabFor = d.type; rib.select(want); }
        if (!want && d._lastTabFor !== d.type) { d._lastTabFor = d.type; rib.select('home'); }
      }
    }
    api.renderNow = render;

    // ------------------------------------------------------------------ 키보드(표 9-4)
    this.handleKey = e => {
      if (backstage) { if (e.key === 'Escape') closeBackstage(); return; }
      const k = e.key, ctrl = e.ctrlKey || e.metaKey, shift = e.shiftKey, alt = e.altKey, code = e.code;
      if (ctrl && code === 'KeyS') { e.preventDefault(); saveActive(); return; }
      if (k === 'F11') { e.preventDefault(); navOpen = !navOpen; render(); return; }
      if (ctrl && code === 'KeyP') { e.preventDefault(); const d = printable(); if (d) openPreview(d); return; }
      if (k === 'F4' || (alt && k === 'Enter')) { if (isDesign()) { e.preventDefault(); toggleProps(); return; } }
      if (alt && k === 'F8') { if (isDesign()) { e.preventDefault(); cur().design.toggleFieldList(); return; } }
      if (ctrl && code === 'KeyZ' && !(e.target && /INPUT|TEXTAREA/.test(e.target.tagName))) { e.preventDefault(); undoDoc(); return; }
      const d = cur();
      if (d && d.keys) d.keys(e);
    };
    this.renderSheet = (key, scale) => {
      if (!db) return null;
      if (key === '5-3') {
        const f = db.forms['조회화면'] || Object.values(db.forms).find(x => x.controls.some(c => c.type === 'listbox')) || Object.values(db.forms)[0];
        return f ? OA.acDesign.renderFormPage(db, f, scale, 0) : null;
      }
      const r = Object.values(db.reports)[0];
      return r ? OA.acDesign.renderReportPage(db, r, scale, 0) : null;
    };
    this.onShow = () => { render(); if (!db) openBackstage('new'); };
    render();
  }

  // ===================================================================== 테이블
  OA.acTable = {
    design(api, d) {
      const db = api.db();
      const t = d.obj;
      if (!d.td) d.td = {};
      let curRow = d.curRow || 0;
      const rows = Math.max(t.fields.length + 12, 20);
      const wrap = h('div', { class: 'ac-tdesign' });
      const table = h('table', { style: { width: '100%' } }, h('tr', {}, h('th', { style: { width: '22px' } }), h('th', { style: { width: '30%' } }, '필드 이름'), h('th', { style: { width: '22%' } }, '데이터 형식'), h('th', {}, '설명(옵션)')));
      const touch = () => { if (!d.dirty) { d.dirty = true; api.render && api.renderNow(); } };
      const top = h('div', { class: 'top' }, table);
      const propBox = h('div', { class: 'props' });
      const selRow = i => { curRow = i; d.curRow = i; [...table.querySelectorAll('tr')].forEach((tr, k) => tr.classList.toggle('on', k - 1 === i)); drawProps(); };
      for (let i = 0; i < rows; i++) {
        const f = t.fields[i];
        const nameIn = h('input', { value: f ? f.name : '', onfocus: () => selRow(i), onkeydown: e => e.stopPropagation(),
          onchange: () => {
            const v = nameIn.value.trim();
            api.snapshot(d);
            if (!t.fields[i] && v) { while (t.fields.length < i) t.fields.push(A.defaultField('', '짧은 텍스트')); t.fields[i] = A.defaultField(v, '짧은 텍스트'); typeSel.value = '짧은 텍스트'; }
            else if (t.fields[i]) { const old = t.fields[i].name; t.fields[i].name = v; t.pk = t.pk.map(k => k === old ? v : k); t.rows.forEach(r => { if (old !== v) { r[v] = r[old]; delete r[old]; } }); }
            touch(); drawProps(); sel.textContent = t.pk.includes(v) ? '🔑' : '';
          } });
        const typeSel = h('select', { onfocus: () => selRow(i), onchange: () => {
          if (!t.fields[i]) { typeSel.value = ''; return; }
          api.snapshot(d);
          const fld = t.fields[i];
          const old = fld.type;
          fld.type = typeSel.value;
          delete fld.size; delete fld.decimals; if (fld.format === '통화') fld.format = '';
          A.applyTypeDefaults(fld);
          let lost = 0;
          t.rows.forEach(r => { const c = A.coerce(fld, r[fld.name]); if (c.error) { lost++; r[fld.name] = null; } else r[fld.name] = c.v; });
          if (lost) msg({ title: 'Microsoft Access', text: `데이터 형식을 ${old}에서 ${fld.type}(으)로 바꾸면서 ${lost}개 값이 변환되지 않았습니다.`, icon: 'warn' });
          touch(); drawProps();
        } }, h('option', { value: '' }, ''), A.TYPES.map(x => h('option', { value: x, selected: f && f.type === x, disabled: ['OLE 개체', '첨부 파일', '계산'].includes(x) }, x)));
        const descIn = h('input', { value: f ? f.desc || '' : '', onfocus: () => selRow(i), onkeydown: e => e.stopPropagation(), onchange: () => { if (t.fields[i]) { t.fields[i].desc = descIn.value; touch(); } } });
        const sel = h('td', { class: 'sel', onclick: e => { selRow(i); if (e.ctrlKey) { d.multi = d.multi || []; d.multi.push(i); } else d.multi = [i]; } }, f && t.pk.includes(f.name) ? '🔑' : '');
        table.appendChild(h('tr', { class: i === curRow ? 'on' : '' }, sel, h('td', {}, nameIn), h('td', {}, typeSel), h('td', {}, descIn)));
      }
      function drawProps() {
        propBox.innerHTML = '';
        const f = t.fields[curRow];
        const grid = h('div', { class: 'grid' });
        const help = h('div', { class: 'help' }, f ? '필드 이름은 공백을 포함하여 64자까지 입력할 수 있습니다. 필드 이름에 대한 도움말을 보려면 F1 키를 누르십시오.' : '필드 이름과 데이터 형식을 입력하면 필드 속성이 표시됩니다.');
        const prop = (label, key, kind, opts) => {
          let input;
          const val = f[key] === undefined ? '' : f[key];
          if (kind === 'select') input = h('select', { onchange: () => { api.snapshot(d); f[key] = typeof val === 'boolean' ? input.value === '예' : (key === 'size' && f.type === '짧은 텍스트' ? +input.value : input.value); if (key === 'indexed' && t.pk.includes(f.name) && input.value !== '예(중복 불가능)') { msg({ title: 'Microsoft Access', text: '기본 키 필드의 인덱스는 예(중복 불가능)여야 합니다.', icon: 'warn' }); f.indexed = '예(중복 불가능)'; } touch(); drawProps(); } }, opts.map(o => h('option', { selected: String(typeof val === 'boolean' ? (val ? '예' : '아니요') : val) === String(o) }, o)));
          else input = h('input', { value: val, onkeydown: e => e.stopPropagation(), onchange: () => {
            api.snapshot(d);
            if (key === 'size') { const n = parseInt(input.value, 10); if (!(n >= 1 && n <= 255)) { msg({ title: 'Microsoft Access', text: '설정은 1에서 255 사이의 수여야 합니다.', icon: 'warn' }); input.value = f.size; return; } if (t.rows.some(r => r[f.name] && String(r[f.name]).length > n)) msg({ title: 'Microsoft Access', text: '필드 크기를 줄이면 일부 데이터가 손실될 수 있습니다.', icon: 'warn' }); f.size = n; }
            else f[key] = input.value;
            touch();
          } });
          input.addEventListener('focus', () => { help.textContent = PROP_HELP[label] || ''; });
          grid.append(h('div', {}, label), h('div', {}, input));
        };
        if (f && f.type) {
          if (f.type === '짧은 텍스트') { prop('필드 크기', 'size', 'text'); prop('형식', 'format', 'text'); prop('입력 마스크', 'mask', 'text'); prop('캡션', 'caption', 'text'); prop('기본값', 'defaultValue', 'text'); prop('유효성 검사 규칙', 'rule', 'text'); prop('유효성 검사 텍스트', 'ruleText', 'text'); prop('필수', 'required', 'select', ['예', '아니요']); prop('빈 문자열 허용', 'allowZero', 'select', ['예', '아니요']); prop('인덱스', 'indexed', 'select', ['아니요', '예(중복 가능)', '예(중복 불가능)']); }
          else if (f.type === '숫자') { prop('필드 크기', 'size', 'select', A.NUM_SIZES); prop('형식', 'format', 'select', A.FORMATS['숫자']); prop('소수 자릿수', 'decimals', 'select', ['자동', '0', '1', '2', '3', '4']); prop('입력 마스크', 'mask', 'text'); prop('캡션', 'caption', 'text'); prop('기본값', 'defaultValue', 'text'); prop('유효성 검사 규칙', 'rule', 'text'); prop('필수', 'required', 'select', ['예', '아니요']); prop('인덱스', 'indexed', 'select', ['아니요', '예(중복 가능)', '예(중복 불가능)']); }
          else if (f.type === '통화') { prop('형식', 'format', 'select', A.FORMATS['통화']); prop('소수 자릿수', 'decimals', 'select', ['자동', '0', '1', '2']); prop('입력 마스크', 'mask', 'text'); prop('캡션', 'caption', 'text'); prop('기본값', 'defaultValue', 'text'); prop('필수', 'required', 'select', ['예', '아니요']); prop('인덱스', 'indexed', 'select', ['아니요', '예(중복 가능)', '예(중복 불가능)']); }
          else { prop('형식', 'format', 'text'); prop('캡션', 'caption', 'text'); prop('필수', 'required', 'select', ['예', '아니요']); prop('인덱스', 'indexed', 'select', ['아니요', '예(중복 가능)', '예(중복 불가능)']); }
        }
        propBox.append(h('div', { style: { flex: 1, display: 'flex', flexDirection: 'column' } }, h('div', { style: { padding: '4px 8px', fontWeight: 700 } }, '필드 속성'), h('div', { class: 'tabs' }, h('button', { class: 'btn' }, '일반'), h('button', { class: 'btn', disabled: true }, '조회')), h('div', { style: { overflow: 'auto', flex: 1 } }, grid)), help);
      }
      Object.assign(d.td, {
        primaryKey() {
          const sel = (d.multi && d.multi.length ? d.multi : [curRow]).map(i => t.fields[i]).filter(Boolean).map(f => f.name).filter(Boolean);
          if (!sel.length) return;
          api.snapshot(d);
          const same = sel.length === t.pk.length && sel.every(n => t.pk.includes(n));
          A.setPrimaryKey(t, same ? [] : sel);
          d.dirty = true; api.renderNow();
        },
        insertRow() { api.snapshot(d); t.fields.splice(curRow, 0, A.defaultField('', '짧은 텍스트')); t.fields[curRow].type = ''; d.dirty = true; api.renderNow(); },
        deleteRow() { const f = t.fields[curRow]; if (!f) return; if (db.relationships.some(r => (r.t1 === d.name && r.f1 === f.name) || (r.t2 === d.name && r.f2 === f.name))) return msg({ title: 'Microsoft Access', text: '이 필드는 관계에 사용되고 있으므로 삭제할 수 없습니다. 먼저 관계 창에서 관계를 삭제하십시오.', icon: 'warn' }); api.snapshot(d); t.fields.splice(curRow, 1); t.pk = t.pk.filter(k => k !== f.name); d.dirty = true; api.renderNow(); },
        addIdKey() { t.fields.unshift(Object.assign(A.defaultField('ID', '일련 번호'), { required: true })); A.setPrimaryKey(t, ['ID']); t.rows.forEach((r, i) => { r.ID = i + 1; }); },
        commit(name) {
          const fields = t.fields.filter(f => f && f.name);
          const names = fields.map(f => f.name);
          if (!fields.length) { msg({ title: 'Microsoft Access', text: '테이블에는 필드가 하나 이상 있어야 합니다.', icon: 'warn' }); return false; }
          if (fields.some(f => !f.type)) { msg({ title: 'Microsoft Access', text: '데이터 형식을 지정하지 않은 필드가 있습니다.', icon: 'warn' }); return false; }
          if (new Set(names).size !== names.length) { msg({ title: 'Microsoft Access', text: '필드 이름이 중복되었습니다. 필드 이름은 테이블 안에서 고유해야 합니다.', icon: 'warn' }); return false; }
          if (names.some(n => /[.!`[\]]/.test(n))) { msg({ title: 'Microsoft Access', text: '필드 이름이 잘못되었습니다.', icon: 'warn' }); return false; }
          const old = d.name && db.tables[d.name];
          if (old && d.name !== name) delete db.tables[d.name];
          db.tables[name] = { name, fields: U.clone(fields), pk: t.pk.filter(k => names.includes(k)), rows: old ? old.rows.map(r => Object.fromEntries(names.map(n => [n, r[n] === undefined ? null : r[n]]))) : [] };
          if (!db.order.includes('t:' + name)) db.order.push('t:' + name);
          d.obj = U.clone(db.tables[name]);
          return true;
        },
      });
      drawProps();
      wrap.append(top, propBox);
      d.keys = e => { if (e.key === 'F6') { e.preventDefault(); } };
      return wrap;
    },

    datasheet(api, d) {
      const db = api.db();
      const t = d.blankTable ? d.obj : db.tables[d.name];
      const wrap = h('div', { style: { display: 'flex', flexDirection: 'column', height: '100%' } });
      const box = h('div', { style: { flex: 1, overflow: 'auto' } });
      const fields = t.fields;
      d.cur = d.cur === undefined ? 0 : d.cur;
      const pending = d.pending || (d.pending = {});
      const tbl = h('table', { class: 'ds' }, h('tr', {}, h('th', {}), fields.map(f => h('th', {}, f.name)), d.blankTable ? h('th', { style: { color: '#888' } }, '추가하려면 클릭 ▼') : null));
      const commitRow = (i) => {
        if (!pending.row && pending.row !== 0) return true;
        if (d.blankTable) { pending.row = null; return true; }
        const isNew = pending.row === t.rows.length;
        const rec = Object.assign({}, isNew ? {} : t.rows[pending.row], pending.vals);
        if (isNew && Object.keys(pending.vals).every(k => pending.vals[k] === '' || pending.vals[k] === null)) { pending.row = null; pending.vals = {}; return true; }
        const r = isNew ? A.insertRow(db, d.name, rec) : A.updateRow(db, d.name, pending.row, rec);
        if (r.error) { msg({ title: 'Microsoft Access', text: r.error, icon: 'warn' }); return false; }
        pending.row = null; pending.vals = {};
        api.changed();
        return true;
      };
      d.commitRow = commitRow;
      const rowsN = t.rows.length;
      for (let i = 0; i <= rowsN; i++) {
        const r = t.rows[i];
        const isNew = i === rowsN;
        const tr = h('tr', { class: i === d.cur ? 'cur' : '' });
        const rs = h('td', { class: 'rs' }, pending.row === i ? '✎' : isNew ? '*' : (i === d.cur ? '▶' : ''));
        rs.addEventListener('click', () => { if (!commitRow()) return; d.cur = i; d.rowSel = !isNew ? i : null; api.renderNow(); });
        tr.appendChild(rs);
        fields.forEach(f => {
          const v = pending.row === i && f.name in (pending.vals || {}) ? pending.vals[f.name] : (r ? r[f.name] : null);
          const shown = (pending.row === i && f.name in (pending.vals || {})) ? (v === null ? '' : v) : (v === null || v === undefined ? (isNew && f.type === '일련 번호' ? '(새 항목)' : '') : A.formatValue(v, f));
          const inp = h('input', { value: shown, readOnly: f.type === '일련 번호' });
          inp.addEventListener('focus', () => { if (d.cur !== i) { if (!commitRow()) { setTimeout(() => api.renderNow(), 0); return; } d.cur = i; [...tbl.querySelectorAll('tr')].forEach((x, k) => x.classList.toggle('cur', k - 1 === i)); } if (r && pending.row !== i) inp.value = v === null || v === undefined ? '' : String(v); inp.select(); });
          inp.addEventListener('input', () => { pending.row = i; pending.vals = pending.vals || {}; if (Object.keys(pending.vals).length === 0 && r) Object.assign(pending.vals, {}); pending.vals[f.name] = inp.value; rs.textContent = '✎'; });
          inp.addEventListener('keydown', e => {
            e.stopPropagation();
            if (e.key === 'Enter' && e.shiftKey) { e.preventDefault(); if (commitRow()) api.renderNow(); return; }
            if (e.key === 'Enter' || e.key === 'Tab') {
              e.preventDefault();
              const tds = [...tbl.querySelectorAll('input')];
              const k = tds.indexOf(inp);
              const next = tds[k + (e.shiftKey ? -1 : 1)];
              const rowEnd = (k + 1) % fields.length === 0;
              if (rowEnd && !e.shiftKey) { if (commitRow()) { d.cur = i + 1; api.renderNow(); setTimeout(() => { const all = [...box.querySelectorAll('input')]; const n = all[(i + 1) * fields.length]; if (n) n.focus(); }, 0); } return; }
              if (next) next.focus();
            }
            if (e.key === 'Escape') { pending.row = null; pending.vals = {}; api.renderNow(); }
            if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); if (!commitRow()) return; d.cur = Math.max(0, Math.min(rowsN, i + (e.key === 'ArrowDown' ? 1 : -1))); api.renderNow(); setTimeout(() => { const all = [...box.querySelectorAll('tr')][d.cur + 1]; const n = all && all.querySelectorAll('input')[fields.indexOf(f)]; if (n) n.focus(); }, 0); }
          });
          tr.appendChild(h('td', { class: ['숫자', '통화', '일련 번호', '큰 번호'].includes(f.type) ? 'num' : '' }, inp));
        });
        if (d.blankTable) tr.appendChild(h('td'));
        tbl.appendChild(tr);
      }
      if (d.totals) {
        const tr = h('tr', { class: 'tot' }, h('td', { class: 'rs' }, 'Σ'));
        d.totFn = d.totFn || {};
        fields.forEach((f, k) => {
          const numeric = ['숫자', '통화', '큰 번호'].includes(f.type);
          const fn = d.totFn[f.name] || (k === 0 ? 'label' : '');
          const vals = t.rows.map(r => r[f.name]).filter(v => v !== null && v !== undefined && v !== '');
          const calc = { '합계': () => vals.reduce((s, x) => s + x, 0), '평균': () => vals.reduce((s, x) => s + x, 0) / vals.length, '개수': () => vals.length, '최대값': () => Math.max(...vals), '최소값': () => Math.min(...vals) };
          const opts = numeric ? ['', '합계', '평균', '개수', '최대값', '최소값'] : ['', '개수'];
          const sel = h('select', { style: { width: '100%' }, onchange: () => { d.totFn[f.name] = sel.value; api.renderNow(); } }, opts.map(o => h('option', { value: o, selected: fn === o }, o || '없음')));
          const shown = fn === 'label' ? '요약' : fn && calc[fn] ? A.formatValue(calc[fn](), fn === '개수' ? {} : f) : '';
          tr.appendChild(h('td', { title: '요약 함수 선택' }, fn === 'label' ? '요약' : h('div', { style: { display: 'flex', gap: '2px' } }, h('span', { style: { flex: 1, textAlign: 'right' } }, shown), sel)));
        });
        tbl.appendChild(tr);
      }
      box.appendChild(tbl);
      const total = t.rows.length;
      const nav = h('div', { class: 'recnav' }, '레코드: ', h('button', { onclick: () => { d.cur = 0; api.renderNow(); } }, '|◀'), h('button', { onclick: () => { d.cur = Math.max(0, d.cur - 1); api.renderNow(); } }, '◀'),
        h('input', { value: `${Math.min(d.cur + 1, total + 1)}/${total}`, readOnly: true }), h('button', { onclick: () => { d.cur = Math.min(total, d.cur + 1); api.renderNow(); } }, '▶'), h('button', { onclick: () => { d.cur = total - 1; api.renderNow(); } }, '▶|'), h('button', { onclick: () => { d.cur = total; api.renderNow(); } }, '▶*'),
        h('span', { style: { marginLeft: '12px', color: '#777' } }, '필터 없음'));
      wrap.append(box, nav);
      d.keys = e => {
        if (e.key === 'Delete' && d.rowSel !== null && d.rowSel !== undefined && !d.blankTable) {
          e.preventDefault();
          msg({ title: 'Microsoft Access', text: '1개의 레코드를 삭제합니다.\n\n[예]를 클릭하면 삭제한 레코드를 다시 실행할 수 없습니다.\n이 레코드를 삭제하시겠습니까?', buttons: ['예', '아니요'], icon: 'warn' }).then(r => { if (r === '예') { const res = A.deleteRow(db, d.name, d.rowSel); if (res.error) msg({ title: 'Microsoft Access', text: res.error, icon: 'warn' }); else api.changed(); d.rowSel = null; api.renderNow(); } });
        }
      };
      if (d.blankTable) setTimeout(() => toast('새 데이터베이스와 함께 테이블1이 열렸습니다. 과제는 테이블 디자인으로 만들므로 이 창은 닫습니다(12.1절).'), 50);
      return wrap;
    },
  };
  const PROP_HELP = { '필드 크기': '필드에 입력할 수 있는 최대 문자 수입니다. 최대 255자까지 설정할 수 있습니다.', '형식': '필드의 표시 레이아웃입니다.', '필수': '이 필드에 데이터를 반드시 입력해야 하는지 여부입니다.', '인덱스': '이 필드의 인덱스입니다. 중복 가능 또는 중복 불가능을 선택합니다.', '소수 자릿수': '소수점 오른쪽에 표시할 자릿수입니다.' };

  OA.apps = OA.apps || {};
  OA.apps.access = { create: (win, ctx) => new AccessApp(win, ctx) };
  OA.acPX = PX_CM;
})(typeof window !== 'undefined' ? window : globalThis);

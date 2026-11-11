/* 모범 풀이: 시뮬레이터 화면이 쓰는 것과 같은 모델 명령으로 과제를 처음부터 끝까지 수행한다.
   쓰임 1) 정답 JSON 대조 테스트  2) 장별 시작 상태(11장 = 10장 완성, 13장 = 12장 완성). 원본: 90_자산/그림/ch10~ch16 의 검증 스크립트 절차. */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util;
  const CUR = '₩#,##0;-₩#,##0';

  // ================================================================== Excel
  function excelWorksheet(wb, sp) {
    const XL = OA.xl, R = U.parseRange;
    const S = '작업표';
    XL.renameSheet(wb, 0, S);
    const sh = XL.sheetOf(wb, S);
    for (let c = 2; c <= 11; c++) sh.colW[c] = 82;
    XL.setInput(wb, S, 2, 2, sp.title);
    XL.merge(wb, S, R('B2:K2'), 'center');
    XL.applyStyle(wb, S, R('B2'), { sz: 20 });
    sp.head.forEach((h, j) => XL.setInput(wb, S, 4, 2 + j, h));
    sp.data.forEach((row, i) => row.forEach((v, j) => XL.setInput(wb, S, 5 + i, 2 + j, String(v))));
    XL.applyStyle(wb, S, R('G5:G16'), { nf: '0%' });
    XL.applyStyle(wb, S, R('B4:K4'), { b: true, ha: 'center' });
    XL.applyStyle(wb, S, R('B5:C16'), { ha: 'center' });
    XL.applyStyle(wb, S, R('D5:G16'), { ha: 'right' });
    XL.setBorder(wb, S, R('B4:K23'), 'all');
    // 계산 열: 5행 입력 뒤 채우기 핸들로 16행까지
    XL.setInput(wb, S, 5, 8, sp.f_h.replace(/\{r\}/g, 5));
    XL.setInput(wb, S, 5, 9, sp.f_i.replace(/\{r\}/g, 5));
    XL.setInput(wb, S, 5, 10, sp.f_j.replace(/\{r\}/g, 5));
    XL.setInput(wb, S, 5, 11, sp.f_k.replace(/\{r\}/g, 5));
    XL.fill(wb, S, R('H5:K5'), R('H5:K16'));
    XL.applyStyle(wb, S, R('H5:H16'), { ha: 'center' });
    XL.merge(wb, S, R('B17:H17'), 'center');
    XL.setInput(wb, S, 17, 2, '평균');
    ['I', 'J', 'K'].forEach((c, k) => XL.setInput(wb, S, 17, 9 + k, `=AVERAGE(${c}5:${c}16)`));
    const rowsTxt = [[18, sp.t6, sp.f6], [20, sp.t8, sp.f8], [22, sp.t10, sp.f10]];
    for (const [r, t, f] of rowsTxt) {
      XL.merge(wb, S, R(`B${r}:J${r}`), 'center');
      XL.setInput(wb, S, r, 2, t);
      XL.setInput(wb, S, r, 11, f);
    }
    XL.applyStyle(wb, S, R('B17:K23'), { ha: 'center' });
    XL.applyStyle(wb, S, [R('I5:K18'), R('K22')], { nf: CUR, ha: 'right' });
    XL.applyStyle(wb, S, R('K20'), { ha: 'right' });
    // 정렬
    XL.sortRange(wb, S, R('B4:K16'), [{ col: 8, order: sp.sort_key_desc ? 'desc' : 'asc' }, { col: 11, order: 'desc' }], true);
    // 함수식 표시(정렬 뒤 최종 수식)
    for (const [r] of rowsTxt) {
      XL.merge(wb, S, R(`B${r + 1}:K${r + 1}`), 'center');
      XL.setInput(wb, S, r + 1, 2, "'" + XL.getCell(sh, r, 11).f);
    }
    // 그래프: 정렬 뒤 해당 구분의 행
    let n = 0;
    for (let r = 5; r <= 16; r++) if (XL.getValue(wb, S, r, 8) === sp.chart_group) n++;
    const end = 4 + n;
    const series = XL.seriesFromSelection(wb, S, [R(`C4:C${end}`), R(`I4:I${end}`), R(`K4:K${end}`)]);
    const ch = XL.newChart(wb, S, series, 'combo');
    const kSer = ch.series.find(s => s.val.includes('$K$')), iSer = ch.series.find(s => s.val.includes('$I$'));
    kSer.type = 'col'; iSer.type = 'lineMarkers'; kSer.secondary = iSer.secondary = false;
    iSer.labels = 'top';
    ch.series = [kSer, iSer];
    ch.title = { show: true, text: sp.chart_title };
    ch.axisTitle = { x: { show: true, text: sp.x_title }, y: { show: true, text: sp.y_title } };
    ch.legend = { show: true, pos: 'bottom' };
    ch.x = XL.colLeft(sh, 2); ch.y = XL.rowTop(sh, 25);
    ch.w = XL.colLeft(sh, 12) - XL.colLeft(sh, 2); ch.h = 300;
    sh.printArea = 'B2:K40';
    Object.assign(sh.page, { orient: 'portrait', paper: 'A4', top: 6, centerH: true, scaleMode: 'fit', fitW: 1, fitH: 1 });
    XL.touch(wb);
    return { chart: ch, n };
  }

  function excelPivot(wb, sp) {
    const XL = OA.xl, PV = OA.pivot, R = U.parseRange;
    const pvs = sp.pivot;
    const sh2 = XL.addSheet(wb, '피벗', wb.sheets.length);
    const res = PV.create(wb, '작업표!$B$4:$K$16', { ref: '피벗!$B$4' });
    const pv = res.pv;
    PV.addField(wb, pv, pvs.row1, 'rows');
    PV.addField(wb, pv, pvs.row2, 'rows');
    PV.addField(wb, pv, pvs.vals[0], 'values');
    PV.addField(wb, pv, pvs.vals[1], 'values');
    pv.values.forEach((v, i) => PV.setValueField(wb, pv, i, { func: 'sum', nf: '#,##0' }));
    const applyNames = (p) => {
      if (pvs.kind !== 'group_items') return;
      const c = PV.cacheOf(wb, p);
      const gd = PV.groupDef(c, gfield);
      p.itemNames[gfield] = {};
      for (const [name, members] of Object.entries(pvs.groups)) {
        const keys = members.map(PV.rawKey);
        const it = gd.items.find(x => x.members.length === keys.length && keys.every(k => x.members.includes(k)));
        p.itemNames[gfield][it ? 'g:' + it.id : keys[0]] = name;
      }
      p.captions[gfield] = pvs.caption;
    };
    let gfield = null;
    if (pvs.kind === 'group_items') {
      for (const members of Object.values(pvs.groups)) {
        if (members.length < 2) continue;                         // 항목 하나짜리 그룹은 묶지 않고 이름만 바꾼다(16장 실측)
        const r = PV.group(wb, sh2, pv, pvs.row1, members.map(PV.rawKey));
        gfield = r.field;
      }
      applyNames(pv);
    }
    pv.layout = 'tabular'; pv.subtotals = 'none';
    XL.touch(wb);
    const keysOfGroups = (p, names) => Object.entries(p.itemNames[gfield] || {}).filter(([, nm]) => names.includes(nm)).map(([k]) => k);
    if (pvs.kind === 'group_items') pv.collapsed[gfield] = keysOfGroups(pv, pvs.adj1_collapse);
    else pv.collapsed[pvs.row1] = pvs.adj1_collapse.map(PV.rawKey);
    XL.touch(wb);
    const t1 = PV.tableRange(wb, sh2, pv);
    XL.setInput(wb, '피벗', 2, 2, '피벗테이블 조정1');
    XL.merge(wb, '피벗', { r1: 2, c1: 2, r2: 2, c2: t1.c2 > 6 ? t1.c2 : 6 }, 'center');
    XL.applyStyle(wb, '피벗', R('B2'), { b: true });
    // 조정2: 조정1 전체 복사 → 아래에 붙여 넣기
    const clip = XL.copy(wb, '피벗', t1);
    const dest = t1.r2 + 4;
    XL.paste(wb, '피벗', clip, dest, t1.c1);
    const pv2 = sh2.pivots[1];
    applyNames(pv2);
    const c = PV.cacheOf(wb, pv2);
    if (pvs.kind === 'group_items') {
      pv2.collapsed[gfield] = [];
      pv2.collapsed[pvs.row1] = PV.baseOrderForField(c, pv2, pvs.row1).filter(k => k !== PV.rawKey(pvs.adj2_expand_item));
    } else {
      pv2.collapsed[pvs.row1] = PV.baseOrderForField(c, pv2, pvs.row1).filter(k => k !== PV.rawKey(pvs.adj2_expand_item));
    }
    XL.touch(wb);
    const t2 = PV.tableRange(wb, sh2, pv2);
    XL.setInput(wb, '피벗', dest - 2, 2, '피벗테이블 조정2');
    XL.merge(wb, '피벗', { r1: dest - 2, c1: 2, r2: dest - 2, c2: 6 }, 'center');
    XL.applyStyle(wb, '피벗', { r1: dest - 2, c1: 2, r2: dest - 2, c2: 2 }, { b: true });
    for (let col = 2; col <= 6; col++) sh2.colW[col] = 100;
    sh2.printArea = `B2:${U.colName(Math.max(6, t2.c2))}${t2.r2}`;
    Object.assign(sh2.page, { orient: 'portrait', paper: 'A4', top: 6, centerH: true, scaleMode: 'fit', fitW: 1, fitH: 1 });
    XL.touch(wb);
    return { pv, pv2, t1, t2, gfield };
  }

  function excelSet(set, upTo) {
    const wb = OA.xl.newWorkbook();
    const a = excelWorksheet(wb, set.sp);
    if (upTo === 10) return wb;
    excelPivot(wb, set.sp);
    wb.active = 0;
    return wb;
  }

  /* 15장 과제 A~F: 원고 표의 자료·수식을 그대로 입력 */
  function excelCh15(set) {
    const XL = OA.xl, R = U.parseRange;
    const wb = XL.newWorkbook();
    XL.renameSheet(wb, 0, '작업표');
    const S = '작업표';
    const ch = set.chapters[0];
    const tbl = cap => ch.task.concat(ch.formulas).find(b => b.t === 'table' && b.caption.startsWith(cap));
    const t6 = tbl('표 15-6.');
    t6.rows[0].slice(1).forEach((h, j) => XL.setInput(wb, S, 4, 2 + j, h));
    t6.rows.slice(1).forEach((row, i) => row.slice(1).forEach((v, j) => XL.setInput(wb, S, 5 + i, 2 + j, v)));
    ['점검시간', '인건비', '할증', '청구액'].forEach((h, j) => XL.setInput(wb, S, 4, 8 + j, h));
    const t14 = tbl('표 15-14.');
    XL.setInput(wb, S, 4, 16, '설치일'); XL.setInput(wb, S, 4, 17, '점검일');
    t14.rows.forEach((row, i) => { XL.setInput(wb, S, 5 + i, 16, row[2]); XL.setInput(wb, S, 5 + i, 17, row[3]); });
    const t12 = tbl('표 15-12.');
    t12.head.forEach((h, j) => XL.setInput(wb, S, 27, 2 + j, h));
    t12.rows.forEach((row, i) => row.forEach((v, j) => XL.setInput(wb, S, 28 + i, 2 + j, v)));
    const formulas = [];
    for (const cap of ['표 15-8.', '표 15-9.', '표 15-11.', '표 15-13.', '표 15-15.']) {
      for (const row of tbl(cap).rows) {
        const a = U.parseAddr(row[0]);
        if (!a) continue;                                            // '비교' 행은 위치가 없어 아래 M26 에 둔다
        formulas.push([a, row[1]]);
      }
    }
    const cmpRow = tbl('표 15-9.').rows.find(r => r[0] === '비교');
    for (const [a, f] of formulas) {
      XL.setInput(wb, S, a.r, a.c, f);
      if (a.r === 5) XL.fill(wb, S, { r1: 5, c1: a.c, r2: 5, c2: a.c }, { r1: 5, c1: a.c, r2: 14, c2: a.c });
    }
    if (cmpRow) XL.setInput(wb, S, 26, 13, cmpRow[1]);
    // 과제 C: 차트용 사본 M16:N19(값 셀은 집계 셀 참조) → 묶은 세로 막대형(verify_ch15_excel.py 와 같은 구성)
    XL.setInput(wb, S, 16, 13, '구분'); XL.setInput(wb, S, 16, 14, '청구액 합');
    [['라우터', 'K17'], ['스위치', 'K18'], ['무선AP', 'K19']].forEach(([nm, src], i) => { XL.setInput(wb, S, 17 + i, 13, nm); XL.setInput(wb, S, 17 + i, 14, '=' + src); });
    XL.applyStyle(wb, S, R('N17:N19'), { nf: CUR });
    const chC = XL.newChart(wb, S, XL.seriesFromSelection(wb, S, [R('M16:N19')]), 'col', { x: XL.colLeft(wb.sheets[0], 13), y: XL.rowTop(wb.sheets[0], 21) });
    chC.title.text = '구분별 청구액 합'; chC.legend.show = false; chC.series[0].labels = 'outEnd';
    XL.applyStyle(wb, S, [R('G5:G14'), R('I5:K14'), R('K17:K21'), R('O5:O14'), R('U5:U14'), R('K22'), R('K24'), R('M26')], { nf: CUR });
    XL.applyStyle(wb, S, R('P5:Q14'), { nf: 'yyyy-mm-dd' });
    XL.touch(wb);
    return wb;
  }

  // ================================================================== Access(12·13장)
  function accessSet(set, upTo) {
    const A = OA.acc;
    const db = set.db;
    const d = A.newDatabase(db.dbName || '연습');
    // 테이블 디자인
    A.createTable(d, '구분', [
      { name: '구분코드', type: '짧은 텍스트', size: 1 },
      { name: db.gname, type: '짧은 텍스트', size: 20, required: true },
    ], ['구분코드']);
    A.createTable(d, db.tname, db.fields.map(([name, type]) => {
      const f = { name, type };
      if (type === '짧은 텍스트') f.size = name === '구분코드' ? 1 : 20;
      if (type === '숫자') f.size = '정수(Long)';
      if (type === '통화') { f.format = '통화'; f.decimals = 0; }
      if (name === db.who) f.required = true;
      return f;
    }), [db.key]);
    for (const [code, name] of db.gubun) A.insertRow(d, '구분', { 구분코드: code, [db.gname]: name });
    const names = db.fields.map(f => f[0]);
    for (const row of db.rows) { const rec = {}; names.forEach((n, i) => { rec[n] = row[i]; }); A.insertRow(d, db.tname, rec); }
    A.addRelationship(d, { t1: '구분', f1: '구분코드', t2: db.tname, f2: '구분코드', ri: true });
    // 조회 쿼리(디자인 눈금)
    const q = A.newQueryDesign(d, [db.tname, '구분']);
    const crit = db.q_codes.map(c => `'${c}'`).join(' Or ');
    q.cols = [
      { field: db.key, table: db.tname, sort: '오름차순', show: true, crit: [] },
      { field: db.who, table: db.tname, sort: '', show: true, crit: [] },
      { field: '구분코드', table: db.tname, sort: '', show: true, crit: [crit] },
      { field: db.fee, table: db.tname, sort: '', show: true, crit: ['<' + db.q_fee_lt] },
      { field: db.gname, table: '구분', sort: '', show: true, crit: [] },
    ];
    A.saveQuery(d, db.qname, q);
    // 조회 폼
    const f = A.newForm(d);
    Object.assign(f.props, { width: 16.5, recordSelectors: false, navButtons: false, dividingLines: false, scrollBars: '없음' });
    const lbl = A.addControl(f, 'label', { left: 0.8, top: 0.6, width: 14.9, height: 1.9, caption: db.form_title, fontSize: 16, fontWeight: '굵게', name: '제목' });
    const lst = A.addControl(f, 'listbox', { left: 0.8, top: 2.8, width: 14.9, height: 4.6, fontSize: 10, name: '목록', withLabel: false });
    Object.assign(lst, { rowSourceType: '테이블/쿼리', rowSource: db.qname, columnCount: 5, columnHeads: true, columnWidths: '2.2cm;3.6cm;2.2cm;3.2cm;2.6cm' });
    A.addControl(f, 'label', { left: 0.8, top: 7.8, width: 14.9, height: 0.8, caption: '리스트박스 조회 시 작성된 SQL문', fontSize: 12, name: 'SQL제목' });
    const sql = A.querySQL(d, db.qname);
    A.addControl(f, 'label', { left: 0.8, top: 8.7, width: 14.9, height: 4.2, caption: sql, fontSize: 10, borderStyle: '점선', borderColor: '#000000', name: 'SQL문' });
    Object.assign(f.page, { top: 60, bottom: 20, left: 20, right: 20, orient: 'portrait', paper: 'A4' });
    A.saveForm(d, '조회화면', f);
    if (upTo === 12) return d;
    // 보고서용 쿼리(계산 필드)
    const q2 = A.newQueryDesign(d, [db.tname, '구분']);
    const codes = Object.keys(db.base);
    const b = db.base;
    const calcExpr = `${db.calc}: IIf([구분코드]="${codes[0]}",${b[codes[0]]},IIf([구분코드]="${codes[1]}",${b[codes[1]]},${b[codes[2]]}))*[${db.days}]`;
    const [[l1, v1], [l2, v2], [, v3]] = db.disc_rule;
    const discExpr = `${db.disc}: IIf([${db.days}]>=${l1},${v1},IIf([${db.days}]>=${l2},${v2},${v3}))`;
    const feeTerm = db.nz ? `Nz([${db.fee}],0)` : `[${db.fee}]`;
    const totExpr = `${db.total}: ${feeTerm}+[${db.calc}]${db.disc_sign > 0 ? '+' : '-'}[${db.disc}]`;
    q2.cols = [
      { field: '구분코드', table: db.tname, show: true, crit: [] },
      { field: db.gname, table: '구분', show: true, crit: [] },
      { field: db.key, table: db.tname, show: true, crit: [] },
      { field: db.days, table: db.tname, show: true, crit: [] },
      { field: db.fee, table: db.tname, show: true, crit: [] },
      { field: calcExpr, table: '', show: true, crit: [] },
      { field: discExpr, table: '', show: true, crit: [] },
      { field: totExpr, table: '', show: true, crit: [] },
    ];
    A.saveQuery(d, db.rname_q, q2);
    // 보고서: 마법사 → 디자인 정리
    const cols = ['구분코드', db.key, db.days, db.fee, db.calc, db.disc, db.total];
    const rpt = A.reportWizard(d, { source: db.rname_q, fields: cols, groups: ['구분코드'], sorts: [[db.key, 'asc']], sums: [db.calc, db.disc, db.total], layout: '단계', orient: '세로', title: '보고서' });
    A.reportTidy(d, rpt, { title: db.report_title, cols, money: [db.fee, db.calc, db.disc, db.total], days: db.days, gname: db.gname, sums: [db.calc, db.disc, db.total] });
    Object.assign(rpt.page, { top: 60, bottom: 20, left: 20, right: 20, orient: 'portrait', paper: 'A4' });
    A.saveReport(d, '보고서', rpt);
    return d;
  }

  // ================================================================== PowerPoint
  function pptSet(set) {
    const P = OA.ppt;
    const pres = P.newPresentation();
    P.setSlideSize(pres, 'standard', 'fit');
    P.setLayout(pres, 0, 'blank');
    P.addSlide(pres, 0);
    set.pt.ref.forEach((shapes, i) => P.loadShapes(pres, i, shapes));
    Object.assign(pres.handout, { header: false, date: false, footer: false, pageNum: false });
    Object.assign(pres.print, { target: 'all', layout: 'h2', frame: true, fit: false, hq: true, orient: 'portrait', color: 'gray' });
    return pres;
  }

  OA.solutions = { excelWorksheet, excelPivot, excelSet, excelCh15, accessSet, pptSet, CUR };
})(typeof window !== 'undefined' ? window : globalThis);

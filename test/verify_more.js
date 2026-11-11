/* verify.js 가 읽는 추가 대조: Access(sql.js)·PowerPoint·자동 채점 */
(function () {
  const fs = require('fs');
  const path = require('path');
  const VROOT = path.join(__dirname, '..');
  let ready = null;
  function initSql(OA) {
    if (!ready) {
      const init = require(path.join(VROOT, 'vendor', 'sql.js', 'sql-wasm.js'));
      ready = init({ wasmBinary: fs.readFileSync(path.join(VROOT, 'vendor', 'sql.js', 'sql-wasm.wasm')) }).then(SQL => OA.sqlx.setEngine(SQL));
    }
    return ready;
  }
  const md = name => fs.readFileSync(path.join(VROOT, '..', '..', '20_원고', '2권', name), 'utf8').split('\r\n').join('\n');
  function codeBlocks(text, after) {
    const i = text.indexOf(after);
    const seg = text.slice(i);
    const out = [];
    const re = /```\n([\s\S]*?)\n```/g;
    let m;
    while ((m = re.exec(seg)) && out.length < 2) out.push(m[1]);
    return out;
  }

  globalThis.__accessChecks = async function (set, check, OA) {
    await initSql(OA);
    const A = OA.acc, R = OA.report;
    const db = set.db, e = set.expected.db;
    const d = OA.solutions.accessSet(set);
    const q = A.runSaved(d, db.qname);
    check(`${set.id} 조회 결과 행`, q.rows, e['조회']);
    const sql = A.querySQL(d, db.qname);
    check(`${set.id} SQL 에 WHERE·INNER JOIN·ORDER BY`, ['WHERE', 'INNER JOIN', 'ORDER BY'].every(k => sql.includes(k)), true);
    if (set.id === 'equip') {
      const [ansi, accessForm] = codeBlocks(md('ch12.md'), '## 12.4');
      check('equip 디자인 눈금 → SQL(원고 12.4 Access 표기)', sql.replace(/\r/g, ''), accessForm);
      const r2 = A.runSQL(d, ansi);
      check('equip 원고 12.4 표준 표기 SQL 실행 결과', r2.rows, e['조회']);
      const r3 = A.runSQL(d, accessForm);
      check('equip 원고 12.4 Access 표기 SQL 실행 결과', r3.rows, e['조회']);
      const back = A.sqlToDesign(d, ansi);
      check('equip SQL → 디자인 눈금 조건', back.cols.filter(c => c.crit[0]).map(c => [c.field, c.crit[0]]), [['구분코드', "'R' Or 'W'"], ['점검비용', '<200000']]);
      // 12장 변형 연습 정답
      const vr = s => A.runSQL(d, s).rows.map(r => r[0]);
      check('equip 12장 변형 1(S·30 이상·내림차순)', vr("SELECT 점검번호 FROM 점검 WHERE 구분코드='S' AND 점검시간>=30 ORDER BY 점검번호 DESC;"), [1002, 1001]);
      check('equip 12장 변형 2(100,000 이상·점검월·번호)', vr('SELECT 점검번호 FROM 점검 WHERE 점검비용>=100000 ORDER BY 점검월, 점검번호;'), [1001, 1002, 1005, 1010, 1006]);
      check('equip 12장 변형 3(W·50,000 미만)', vr("SELECT 점검번호 FROM 점검 WHERE 구분코드='W' AND 점검비용<50000;"), [1007, 1008, 1009]);
      // 13장 주의: 10% 구문 오류, Nz 없이 비어 있는 값
      let msg = '';
      try { A.runSQL(d, 'SELECT 점검비용*10% AS x FROM 점검;'); } catch (err) { msg = err.message; }
      check('equip 백분율 식 오류 문구', msg, "'점검비용*10%' 쿼리식의 구문 오류입니다.");
      check('equip Nz 없이 1011', A.runSQL(d, 'SELECT 점검번호, 점검비용+24000-5000 AS y FROM 점검 WHERE 점검번호=1011;').rows, [[1011, null]]);
      const pr = A.runSQL(d, 'SELECT 점검번호, [없는필드] FROM 점검;');
      check('equip 없는 필드 → 매개 변수 값 입력', pr.needParam, '없는필드');
      const sumCost = d.tables['점검'].rows.reduce((s, r) => s + (r['점검비용'] || 0), 0);
      check('equip 점검비용 합계(12.1 검산)', sumCost, 1100000);
    } else {
      const r2 = A.runSQL(d, e.SQL['조회'].replace(/\r/g, ''));
      check(`${set.id} 실측 SQL(조회) 실행 결과`, r2.rows, e['조회']);
      const r3 = A.runSQL(d, e.SQL['보고서용'].replace(/\r/g, ''));
      check(`${set.id} 실측 SQL(보고서용) 행 수`, r3.rows.length, e['보고서_행'].length);
    }
    // 보고서용 쿼리
    const rq = A.runSaved(d, db.rname_q);
    const ci = n => rq.columns.indexOf(n);
    const rows = rq.rows.map(r => ['구분코드', db.key, db.days, db.fee, db.calc, db.disc, db.total].map(n => r[ci(n)]))
      .sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]));
    check(`${set.id} 보고서용 쿼리 행(계산 필드)`, rows, e['보고서_행']);
    // 보고서 출력
    const rpt = d.reports['보고서'];
    const out = R.renderReport(d, rpt);
    const footers = {}, avg = [];
    let firstDetail = null;
    for (const pg of out.pages) for (const it of pg.items) {
      if (it.inst.section === 'groupFooter0') {
        const code = it.inst.ctx.row['구분코드'];
        footers[code] = it.controls.filter(x => /^=Sum/.test(x.c.controlSource || '')).sort((a, b) => a.c.left - b.c.left).map(x => x.value);
      }
      if (it.inst.section === 'reportFooter') it.controls.filter(x => /^=Avg/.test(x.c.controlSource || '')).sort((a, b) => a.c.left - b.c.left).forEach(x => avg.push(Math.round(x.value * 1e4) / 1e4));
      if (it.inst.section === 'detail' && !firstDetail) firstDetail = it.controls.filter(x => x.c.type === 'textbox').sort((a, b) => a.c.left - b.c.left).map(x => x.text);
    }
    check(`${set.id} 보고서 그룹 합계(그룹 바닥글)`, footers, e['보고서_그룹']);
    check(`${set.id} 보고서 총평균(보고서 바닥글)`, avg, e['보고서_평균'], 1e-3);
    check(`${set.id} 보고서 쪽 수`, out.pageCount, 1);
    const f = d.forms['조회화면'];
    check(`${set.id} 조회화면 쪽 수`, R.renderForm(d, f).pageCount, 1);
    if (set.id === 'equip') {
      check('equip 보고서 첫 행 표시', firstDetail, ['R', '1005', '50', '₩200,000', '₩250,000', '₩20,000', '₩430,000']);
      const lst = R.renderForm(d, f).items.find(x => x.c.type === 'listbox');
      check('equip 목록 상자 금액 표시', lst.list.rows.map((r, i) => A.formatValue(r[3], lst.list.info[3])), ['₩180,000', '₩40,000', '₩40,000', '₩30,000']);
      const rw = OA.util.clone(rpt); rw.props.width = 19;
      check('equip 보고서 너비 19cm → 오른쪽 빈 쪽', R.renderReport(d, rw).pageCount, 2);
      const tf = OA.util.clone(rpt);
      tf.controls.filter(c => c.section === 'reportFooter' && c.type === 'textbox').forEach(c => { c.format = ''; c.decimals = '자동'; });
      const t0 = R.renderReport(d, tf).pages[0].items.find(x => x.inst.section === 'reportFooter').controls.filter(x => x.c.type === 'textbox').sort((a, b) => a.c.left - b.c.left).map(x => x.text);
      check('equip 총평균 서식 전 표시(그림 13-8 위)', t0, ['113545.454545', '12272.727273', '201272.727273']);
    }
  };

  // ------------------------------------------------------------------ PowerPoint
  const TYPE = { 1: 'rect', 5: 'roundRect', 9: 'ellipse', 4: 'diamond', 10: 'hexagon', 69: 'flowTerminator', 64: 'flowData', 67: 'flowDocument', 71: 'flowManualInput',
    83: 'flowStoredData', 86: 'flowMagneticDisk', 33: 'rightArrow', 36: 'downArrow', 31: 'leftBrace', 32: 'rightBrace', 105: 'calloutRect', 106: 'calloutRoundRect', 107: 'calloutOval', 17: 'textbox' };
  globalThis.__pptChecks = function (set, check, OA) {
    const P = OA.ppt, e = set.expected.pt;
    const pres = OA.solutions.pptSet(set);
    check(`${set.id} 슬라이드 도형 수`, pres.slides.map(s => s.shapes.length), e['도형 수']);
    check(`${set.id} 슬라이드 크기(cm)`, [pres.size.w, pres.size.h], [25.4, 19.05]);
    if (e['회색 채우기']) {
      const gray = pres.slides[0].shapes.filter(s => s.kind !== 'textbox' && !P.isLine(s) && P.isGray(s.fill)).map(s => s.name).sort();
      check(`${set.id} 회색 40% 채우기 도형`, gray, e['회색 채우기'].slice().sort());
    }
    if (e['도형 종류']) {
      e['도형 종류'].forEach((list, i) => {
        const got = pres.slides[i].shapes.map(s => [s.name, s.kind]);
        const want = list.map(([nm, t]) => [nm, TYPE[t] || (t === -2 ? 'conn' : t === 9 ? 'line' : t === 6 ? 'group' : 'type' + t)]);
        // 실측 JSON 은 직선(Type 9 msoLine)과 타원(AutoShapeType 9)을 같은 숫자 9로 기록하므로 둘을 같은 값으로 비교한다
        const nine = k => (k === 'ellipse' || k === 'line') ? '9(타원·직선)' : k;
        const norm = got.map(([nm, k]) => [nm, nine(k.startsWith('conn') ? 'conn' : k)]);
        check(`${set.id} 제${i + 1} 슬라이드 도형 이름·종류·순서`, norm, want.map(([nm, k]) => [nm, nine(k)]));
      });
    }
    const pages = P.printPages(pres);
    check(`${set.id} 유인물 쪽 수`, pages.length, e['유인물']['쪽수']);
    check(`${set.id} 유인물 용지(cm)`, [pages[0].paper.w, pages[0].paper.h], e['유인물']['쪽 크기 cm']);
    const want = e['유인물']['슬라이드 틀(x, y, 폭, 높이) cm'];
    const fr = pages[0].frames.map(f => [f.x, f.y, f.w, f.h].map(v => Math.round(v * 100) / 100));
    check(`${set.id} 유인물 슬라이드 틀 위치·크기(±0.05cm)`, fr, [want[0], want[2]], 0.05);
    // 연결선 끝점이 붙은 도형의 연결점에 있는지(그림 14-9)
    let bad = 0;
    pres.slides.forEach(sl => P.eachShape(sl.shapes, s => { if (P.isLine(s) && s.begin) { const pts = P.linePoints(pres, sl, s); const t = P.findByName(sl, (P.eachShape(sl.shapes, () => {}), null)); void t; if (!pts[0].attached) bad++; } }));
    check(`${set.id} 연결선 끝점 붙음`, bad, 0);
  };

  // ------------------------------------------------------------------ 자동 채점
  globalThis.__gradeChecks = async function (set, check, OA, DATA) {
    await initSql(OA);
    const state = { wb: OA.solutions.excelSet(set), db: OA.solutions.accessSet(set), pres: OA.solutions.pptSet(set), status: {} };
    const res = OA.grade.grade(DATA, set, state);
    const bad = res.filter(r => r.status !== 'pass' && r.status !== 'manual').map(r => `${r.sheet} ${r.item.slice(0, 18)}: ${r.notes.join(' ')}`);
    check(`${set.id} 모범 풀이 자동 채점 통과(${res.length}항목, 자필 2항목 제외)`, bad, []);
    if (set.id !== 'equip') return;
    // 채점기 민감도: 일부러 틀린 상태
    const XL = OA.xl;
    const t1 = OA.solutions.excelSet(set);
    XL.sheetOf(t1, '작업표').charts[0].series[1].labels = null;
    XL.sortRange(t1, '작업표', OA.util.parseRange('B4:K16'), [{ col: 3, order: 'asc' }], true);
    const r1 = OA.grade.excel51(set, t1);
    check('채점기: 레이블 빠짐·정렬 틀림 → 4·6번 실패', [r1[3].ok, r1[5].ok], [false, false]);
    const t2 = OA.solutions.excelSet(set);
    XL.setInput(t2, '작업표', 18, 11, '400000');
    check('채점기: 집계 값을 직접 입력 → 3번 실패', OA.grade.excel51(set, t2)[2].ok, false);
    const t3 = OA.solutions.excelSet(set);
    XL.sheetOf(t3, '작업표').page.scaleMode = 'zoom';
    check('채점기: 자동 맞춤 해제 → 2쪽, 8번 실패', [XL.printLayout(t3, '작업표').pages.length > 1, OA.grade.excel51(set, t3)[7].ok], [true, false]);
    const t4 = OA.solutions.excelSet(set);
    XL.sheetOf(t4, '피벗').pivots[1].collapsed = {};
    check('채점기: 조정2 모두 확장 → 5-2 4번 실패', OA.grade.excel52(set, t4)[3].ok, false);
    const d5 = OA.solutions.accessSet(set);
    d5.queries['장비점검조회'].cols[2].crit = ["'R'", "'W'"];
    check('채점기: R·W 를 다른 행에 → 5행, 5-3 2번 실패', [OA.acc.runSaved(d5, '장비점검조회').rows.length, OA.grade.access53(set, d5)[1].ok], [5, false]);
    const d6 = OA.solutions.accessSet(set);
    d6.forms['조회화면'].controls = d6.forms['조회화면'].controls.filter(c => c.name !== 'SQL문');
    check('채점기: SQL 문 레이블 없음 → 5-3 4번 실패', OA.grade.access53(set, d6)[3].ok, false);
    const d7 = OA.solutions.accessSet(set);
    d7.reports['보고서'].controls.filter(c => c.section === 'reportFooter' && c.type === 'textbox').forEach(c => { c.format = ''; });
    check('채점기: 총평균 서식 없음 → 5-4 3번 실패', OA.grade.access54(set, d7)[2].ok, false);
    const p8 = OA.solutions.pptSet(set);
    p8.print.layout = 'full'; p8.handout.date = true;
    const g8 = OA.grade.ppt55(set, p8);
    check('채점기: 전체 페이지 슬라이드·날짜 켜짐 → 5-5 3·4번 실패', [g8[2].ok, g8[3].ok], [false, false]);
    const p9 = OA.solutions.pptSet(set);
    p9.slides[0].shapes.pop();
    check('채점기: 도형 하나 빠짐 → 5-5 1번 실패', OA.grade.ppt55(set, p9)[0].ok, false);
    // 14.3 "기호를 글자로 직접 입력해도 인쇄 결과는 같다": 글머리 기호 대신 ❖ 를 글자로 입력해도 통과
    const p10 = OA.solutions.pptSet(set);
    let bul = null;
    OA.ppt.eachShape(p10.slides[0].shapes, s => { if (s.text && s.text.bullet === '❖') bul = s; });
    if (bul) { bul.text.value = '❖ ' + bul.text.value; bul.text.bullet = null; }
    check('채점기: ❖ 를 글자로 입력 → 5-5 2번 통과', [!!bul, OA.grade.ppt55(set, p10)[1].ok], [true, true]);
    const p11 = OA.solutions.pptSet(set);
    OA.ppt.eachShape(p11.slides[0].shapes, s => { if (s.text && s.text.bullet === '❖') s.text.bullet = null; });
    check('채점기: 글머리 기호 빠짐 → 5-5 2번 실패', OA.grade.ppt55(set, p11)[1].ok, false);
  };
})();

/* 정답 JSON 대조 테스트(2026-11-11).
   시뮬레이터 엔진(js/)을 Node 에서 읽어 모범 풀이를 수행한 뒤, 결과를 실측 JSON(data/scenarios.json 의 expected = 90_자산/그림 의 검증 결과)과 대조한다.
   이어서 자동 채점기(grade.js)가 모범 풀이를 모든 항목 통과로 판정하는지 확인한다.
   실행: node test/verify.js   (불일치 0 이면 종료 코드 0, 하나라도 있으면 1) */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
globalThis.window = globalThis;
const FILES = ['js/core/util.js', 'js/excel/numfmt.js', 'js/excel/formula.js', 'js/excel/workbook.js', 'js/excel/pivot.js',
  'js/access/sqlx.js', 'js/access/db.js', 'js/access/report.js', 'js/ppt/model.js', 'js/core/solutions.js', 'js/core/grade.js'];
for (const f of FILES) {
  const p = path.join(ROOT, f);
  if (!fs.existsSync(p)) { console.log('  (아직 없음)', f); continue; }
  vm.runInThisContext(fs.readFileSync(p, 'utf8'), { filename: f });
}
const OA = globalThis.OA;
const DATA = JSON.parse(fs.readFileSync(path.join(ROOT, 'data', 'scenarios.json'), 'utf8'));
const only = process.argv[2] || '';

let pass = 0, fail = 0;
const fails = [];
function eq(a, b, eps) {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= (eps == null ? 1e-6 : eps);
  if (Array.isArray(a) && Array.isArray(b)) return a.length === b.length && a.every((x, i) => eq(x, b[i], eps));
  if (a && b && typeof a === 'object' && typeof b === 'object') { const ka = Object.keys(a).sort(), kb = Object.keys(b).sort(); return eq(ka, kb) && ka.every(k => eq(a[k], b[k], eps)); }
  return a === b;
}
function check(label, got, want, eps) {
  if (eq(got, want, eps)) { pass++; return true; }
  fail++; fails.push(label); console.log('  불일치', label, '\n    시뮬레이터', JSON.stringify(got), '\n    실측      ', JSON.stringify(want));
  return false;
}

// ------------------------------------------------------------------ Excel
function excelChecks(set) {
  const XL = OA.xl, PV = OA.pivot;
  const e = set.expected.sp;
  const wb = OA.solutions.excelSet(set);
  const S = '작업표';
  const v = (a) => { const g = OA.util.parseAddr(a); return XL.getValue(wb, S, g.r, g.c); };
  check(`${set.id} ⑤ 평균`, [v('I17'), v('J17'), v('K17')], e['평균'], 1e-3);
  check(`${set.id} ⑥`, v('K18'), e['⑥']);
  check(`${set.id} ⑧`, v('K20'), e['⑧']);
  check(`${set.id} ⑩`, v('K22'), e['⑩']);
  const names = []; for (let r = 5; r <= 16; r++) names.push(XL.getValue(wb, S, r, 3));
  check(`${set.id} 정렬 순서`, names, e['정렬_사람']);
  const ch = XL.sheetOf(wb, S).charts[0];
  const kSer = ch.series.find(s => s.type === 'col'), iSer = ch.series.find(s => s.type === 'lineMarkers');
  check(`${set.id} 그래프 막대 값`, XL.refValues(wb, kSer.val), e['차트_합']);
  check(`${set.id} 그래프 꺾은선 값`, XL.refValues(wb, iSer.val), e['차트_기본']);
  // 표시(원화·쉼표)
  if (set.id === 'equip') {
    check('equip 표시 I17·J17·K17', ['I17', 'J17', 'K17'].map(a => { const g = OA.util.parseAddr(a); return XL.display(wb, S, g.r, g.c).text; }), ['₩169,167', '₩12,333', '₩181,500']);
    check('equip 정렬 뒤 K5·K10·K11 표시', ['K5', 'K10', 'K11'].map(a => { const g = OA.util.parseAddr(a); return XL.display(wb, S, g.r, g.c).text; }), ['₩352,000', '₩30,000', '₩528,000']);
    check('equip 함수식 표시 B19', XL.display(wb, S, 19, 2).text, '=SUMIF(F5:F16,"<3",I5:I16)');
    check('equip F열 합·G열 합(10.1 검산)', [[5, 6], [5, 7]].map(([, c]) => { let s = 0; for (let r = 5; r <= 16; r++) s += XL.getValue(wb, S, r, c); return Math.round(s * 100) / 100; }), [36, 0.65]);
  }
  // 피벗
  const sh2 = XL.sheetOf(wb, '피벗');
  const [p1, p2] = sh2.pivots;
  const s1 = PV.summary(wb, p1), s2 = PV.summary(wb, p2);
  const groups = {}; for (const [k, it] of Object.entries(s1.items)) groups[k] = it.values;
  check(`${set.id} 피벗 그룹 합(조정1)`, groups, e['피벗_그룹']);
  const groups2 = {}; for (const [k, it] of Object.entries(s2.items)) groups2[k] = it.values;
  check(`${set.id} 피벗 그룹 합(조정2)`, groups2, e['피벗_그룹']);
  check(`${set.id} 피벗 총합계`, [s1.total, s2.total], [e['피벗_총합'], e['피벗_총합']]);
  const t1 = PV.tableRange(wb, sh2, p1), t2 = PV.tableRange(wb, sh2, p2);
  check(`${set.id} 피벗 위치(조정1 시작·끝, 조정2 시작·끝 행)`, [t1.r1, t1.r2, t2.r1, t2.r2], e['피벗_행']);
  if (set.id === 'equip') {
    const yrs = PV.summary(wb, p1, 1);
    const per = {}; for (const [k, it] of Object.entries(yrs.items)) per[k] = it.values;
    check('equip 사용연수별 합', per, e['피벗_연수']);
    check('equip 조정2 확장 행의 담당자', PV.visibleLabels(wb, sh2, p2, '담당자'), e['피벗_확장항목_사람']);
  }
  if (e['피벗_그룹항목']) {
    const c = PV.cacheOf(wb, p1);
    const gd = c.groups[0];
    check(`${set.id} 피벗 그룹 항목 순서·이름`, PV.visibleLabels(wb, sh2, p1, gd.field), e['피벗_그룹항목']);
    const p0 = Object.assign({}, p1, { itemNames: {} });
    check(`${set.id} 그룹 필드 초기 항목 이름(그룹 직후)`, Object.keys(PV.summary(wb, p0).items).sort(), e['피벗_그룹항목_원래이름'].slice().sort());
  }
  return wb;
}

function ch15Checks(set) {
  const XL = OA.xl;
  const wb = OA.solutions.excelCh15(set);
  const S = '작업표';
  const col = (c, fn) => { const out = []; for (let r = 5; r <= 14; r++) out.push(fn ? fn(r) : XL.getValue(wb, S, r, OA.util.colNum(c))); return out; };
  const v = a => { const g = OA.util.parseAddr(a); return XL.getValue(wb, S, g.r, g.c); };
  const d = a => { const g = OA.util.parseAddr(a); return XL.display(wb, S, g.r, g.c).text; };
  const p2 = set.expected.p2, abc = set.expected.abc['실측'];
  for (let r = 5; r <= 14; r++) {
    const want = abc['행' + r];
    check(`ch15 과제 A ${r}행`, { 점검시간표시: d('H' + r), 인건비: v('I' + r), 할증: v('J' + r), 청구액: v('K' + r) },
      { 점검시간표시: want['점검시간표시'], 인건비: want['인건비'], 할증: want['할증'], 청구액: want['청구액'] });
  }
  check('ch15 과제 B K17·K18·K19·K20·K21', [v('K17'), v('K18'), v('K19'), v('K20'), v('K21')], [abc.R, abc.S, abc.W, abc.RS, abc.R_wild]);
  check('ch15 과제 D 순위', col('L'), p2['순위']);
  check('ch15 과제 D 김씨 평균·건수·3위 합', [v('K22'), v('K23'), v('K24')], [p2['김씨 평균'], p2['R 10만 이상 건수'], p2['3위 이내 합']]);
  check('ch15 과제 E 구분', col('M'), p2['구분']);
  check('ch15 과제 E 설치 위치', col('N'), p2['설치 위치']);
  check('ch15 과제 E 기본료', col('O'), p2['기본료']);
  check('ch15 과제 F 사용개월', col('R'), p2['사용개월']);
  check('ch15 과제 F 사용기간', col('S'), p2['사용기간']);
  check('ch15 과제 F 다음 점검일', col('T', r => d('T' + r)), p2['다음 점검일']);
  check('ch15 과제 F 반올림', col('U'), p2['반올림']);
  check('ch15 과제 F 2024년 설치', v('K25'), p2['2024년 설치']);
  check('ch15 표시 K22·K24·S5·T5·U5', [d('K22'), d('K24'), d('S5'), d('T5'), d('U5')], [p2['표시'].K22, p2['표시'].K24, p2['표시'].S5, p2['표시'].T5, p2['표시'].U5]);
  // 15장 관찰값: 상대 참조 RANK, 근사 일치, 오류
  const obs = DATA.sets.find(s => s.id === 'ch15') && JSON.parse(fs.readFileSync(path.join(ROOT, '..', '그림', 'ch15', 'verify_ch15_p2.json'), 'utf8'))['관찰'];
  if (obs) {
    const t = XL.newWorkbook(); XL.renameSheet(t, 0, '작업표');
    // 상대 참조 RANK.EQ 채우기: 같은 작업표(17~25행 집계값 포함)의 빈 열 W 에서
    XL.setInput(wb, S, 5, 23, '=RANK.EQ(K5,K5:K14)');
    XL.fill(wb, S, { r1: 5, c1: 23, r2: 5, c2: 23 }, { r1: 5, c1: 23, r2: 14, c2: 23 });
    const rel = []; for (let r = 5; r <= 14; r++) rel.push(XL.getValue(wb, S, r, 23));
    XL.clearRange(wb, S, { r1: 5, c1: 23, r2: 14, c2: 23 });
    check('ch15 관찰 상대 참조 RANK.EQ', rel, obs['상대 참조 RANK.EQ 채우기']);
    const tb = [['W', '무선AP', 10000], ['R', '라우터', 20000], ['S', '스위치', 15000]];
    tb.forEach((row, i) => row.forEach((x, j) => XL.setInput(t, '작업표', 28 + i, 2 + j, String(x))));
    const look = k => { XL.setInput(t, '작업표', 40, 2, '=VLOOKUP("' + k + '",$B$28:$D$30,3)'); return String(XL.getValue(t, '작업표', 40, 2)); };
    check('ch15 관찰 근사 일치(정렬 안 된 표)', ['R', 'S', 'W'].map(look), obs['근사 일치(정렬 안 된 표) R·S·W']);
    const fx = f => { XL.setInput(t, '작업표', 41, 2, f); const x = XL.getValue(t, '작업표', 41, 2); return x && x.e ? x.e : x; };
    check('ch15 관찰 열 번호 초과·없는 값', [fx('=VLOOKUP("R",$B$28:$D$30,5,FALSE)'), fx('=VLOOKUP("Q",$B$28:$D$30,3,FALSE)')], [obs['열 번호 초과'], obs['없는 값']]);
    const Q5 = XL.getValue(wb, S, 5, 17), P5 = XL.getValue(wb, S, 5, 16);
    XL.setInput(t, '작업표', 1, 1, String(P5)); XL.setInput(t, '작업표', 1, 2, String(Q5));
    check('ch15 관찰 DATEDIF 역순·y·d', [fx('=DATEDIF(B1,A1,"m")'), String(fx('=DATEDIF(A1,B1,"y")')), String(fx('=DATEDIF(A1,B1,"d")'))], [obs['DATEDIF 역순'], obs['DATEDIF y·d(5행)'][0], obs['DATEDIF y·d(5행)'][1]]);
    XL.setInput(t, '작업표', 1, 3, '157500'); XL.applyStyle(t, '작업표', [{ r1: 42, c1: 2, r2: 42, c2: 3 }], { nf: '₩#,##0;-₩#,##0' });
    XL.setInput(t, '작업표', 42, 2, '=ROUNDDOWN(C1,-3)'); XL.setInput(t, '작업표', 42, 3, '=ROUNDUP(C1,-3)');
    check('ch15 관찰 ROUNDDOWN·ROUNDUP·INT', [XL.display(t, '작업표', 42, 2).text, XL.display(t, '작업표', 42, 3).text, String(fx('=INT(C1/1000)*1000'))], obs['ROUNDDOWN·ROUNDUP·INT(5행)']);
    XL.setInput(t, '작업표', 43, 2, '=B1+183'); XL.setInput(t, '작업표', 43, 3, '=DATE(2026,16,5)');
    XL.applyStyle(t, '작업표', [{ r1: 43, c1: 2, r2: 43, c2: 2 }], { nf: 'yyyy-mm-dd' });
    XL.applyStyle(t, '작업표', [{ r1: 1, c1: 2, r2: 1, c2: 2 }], { nf: 'yyyy-mm-dd' });
    check('ch15 관찰 Q5+183·DATE(2026,16,5)', [XL.display(t, '작업표', 43, 2).text, XL.display(t, '작업표', 43, 3).text], [obs['Q5+183'], obs['DATE(2026,16,5)']]);
  }
  // 15.3 버전 차이: 배열 수식이 아닌 SUM(ISNUMBER(FIND()))*범위 를 자료 범위 밖 행에서 Enter 로 입력하면 #VALUE!
  XL.setInput(wb, S, 22, 13, '=SUM(ISNUMBER(FIND("R",C5:C14))*K5:K14)');
  const sv = XL.getValue(wb, S, 22, 13);
  check('ch15 SUM 배열(구형 해석) 오류', sv && sv.e, '#VALUE!');
  check('ch15 같은 계산의 SUMPRODUCT', v('K17'), set.expected.abc['실측']['CSE 값']);
  // 15장 채점기: 모범 풀이는 6개 과제 모두 통과, 오답 사례는 해당 과제만 실패
  if (OA.gradeCh15) {
    const ok = OA.gradeCh15(set, OA.solutions.excelCh15(set));
    check('ch15 채점기: 모범 풀이 6과제 통과', ok.map(r => r.status), ok.map(() => 'pass'));
    const w1 = OA.solutions.excelCh15(set);
    XL.setInput(w1, S, 5, 9, '37500');                                   // I5 수식 대신 값
    const g1 = OA.gradeCh15(set, w1);
    check('ch15 채점기: I5 값 입력 → 과제 A만 실패', g1.map(r => r.status), ['fail', 'pass', 'pass', 'pass', 'pass', 'pass']);
    const w2 = OA.solutions.excelCh15(set);
    const shw = w2.sheets[0];
    shw.charts = [];
    XL.newChart(w2, S, XL.seriesFromSelection(w2, S, [{ r1: 17, c1: 2, r2: 19, c2: 11 }]), 'col');   // 병합된 집계 행을 그대로 범위로
    check('ch15 채점기: 병합 범위 차트 → 과제 C 실패', OA.gradeCh15(set, w2)[2].status, 'fail');
  }
  return wb;
}

// ------------------------------------------------------------------ 실행
async function main() {
  console.log('정답 JSON 대조 테스트', DATA.version);
  for (const set of DATA.sets) {
    if (only && !set.id.startsWith(only)) continue;
    console.log('■', set.id, set.title);
    if (set.ch15) { ch15Checks(set); continue; }
    excelChecks(set);
    if (OA.acc && globalThis.__accessChecks) await globalThis.__accessChecks(set, check, OA, DATA);
    if (OA.ppt && globalThis.__pptChecks) globalThis.__pptChecks(set, check, OA, DATA);
    if (OA.grade && globalThis.__gradeChecks) await globalThis.__gradeChecks(set, check, OA, DATA);
  }
  console.log(`\n대조 ${pass + fail}건, 일치 ${pass}건, 불일치 ${fail}건`);
  if (fail) { console.log('불일치 항목:', fails.join(' / ')); process.exit(1); }
}
const extra = path.join(__dirname, 'verify_more.js');
if (fs.existsSync(extra)) require(extra);
main().catch(e => { console.error(e); process.exit(1); });

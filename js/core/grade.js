/* 자동 채점: 부록 표 A-1 출력물 5매 점검표 30항목. 기대값은 data/scenarios.json 의 실측값(expected)과 모범 풀이 결과.
   결과 항목 {sheet, item, sec, status:'pass'|'fail'|'manual'|'todo', notes:[]} (DOM 없음) */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const U = OA.util;
  const refCache = {};
  const near = (a, b, e) => typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) <= (e === undefined ? 0.01 : e);
  const NT = s => U.normText(s).toUpperCase();

  function reference(set) {
    if (!refCache[set.id]) refCache[set.id] = { wb: OA.solutions.excelSet(set) };
    return refCache[set.id];
  }

  // ------------------------------------------------------------------ Excel
  function findWorksheet(wb, sp) {
    if (!wb) return null;
    return wb.sheets.find(s => s.name === '작업표') || wb.sheets.find(s => { const c = OA.xl.getCell(s, 2, 2); return c && c.v === sp.title; }) || wb.sheets.find(s => !s.pivots.length) || null;
  }
  function excel51(set, wb) {
    const XL = OA.xl, FX = OA.formula, sp = set.sp, e = set.expected.sp;
    const sh = findWorksheet(wb, sp);
    const R = [[], [], [], [], [], [], [], []];
    if (!sh) return R.map(() => ({ ok: false, notes: ['작업표 시트가 없습니다.'] }));
    const S = sh.name;
    const gv = (r, c) => XL.getValue(wb, S, r, c);
    const disp = (r, c) => XL.display(wb, S, r, c).text;
    const cell = (r, c) => XL.getCell(sh, r, c);
    const out = [];
    // 1 제목·위치
    {
      const n = [];
      if (gv(2, 2) !== sp.title) n.push(`B2 제목이 「${sp.title}」이 아닙니다.`);
      if ((XL.styleOf(sh, 2, 2).sz || 11) !== 20) n.push('B2 제목 글꼴 크기가 20pt 가 아닙니다.');
      if (!sh.merges.some(m => m.r1 === 2 && m.c1 === 2 && m.r2 === 2 && m.c2 === 11)) n.push('B2:K2 가 병합되어 있지 않습니다.');
      const head = sp.head.map((h, j) => gv(4, 2 + j));
      if (head.join('|') !== sp.head.join('|')) n.push('4행 머리글(B4:K4)이 지시와 다릅니다.');
      const want = sp.data.map(r => r.map(v => typeof v === 'number' ? +v.toFixed(6) : v).join('|')).sort();
      const got = []; for (let r = 5; r <= 16; r++) { const row = []; for (let c = 2; c <= 7; c++) { const v = gv(r, c); row.push(typeof v === 'number' ? +v.toFixed(6) : v); } got.push(row.join('|')); }
      if (got.sort().join('#') !== want.join('#')) n.push('B5:G16 원자료가 표와 다르거나 위치가 다릅니다.');
      out.push({ ok: !n.length, notes: n });
    }
    // 2 계산 열
    {
      const n = [];
      const ref = reference(set).wb;
      const refRow = {};
      for (let r = 5; r <= 16; r++) refRow[XL.getValue(ref, '작업표', r, 3)] = [8, 9, 10, 11].map(c => XL.getValue(ref, '작업표', r, c));
      let notFormula = 0, wrong = 0, noIf = 0;
      for (let r = 5; r <= 16; r++) {
        for (const c of [8, 9, 10, 11]) { const cl = cell(r, c); if (!cl || !cl.f) notFormula++; }
        const cl = cell(r, 8), ci = cell(r, 9);
        if (cl && cl.f && !FX.functionsUsed(cl.f).includes('IF')) noIf++;
        if (ci && ci.f && !FX.functionsUsed(ci.f).includes('IF')) noIf++;
        const want = refRow[gv(r, 3)];
        const got = [8, 9, 10, 11].map(c => gv(r, c));
        if (!want || !got.every((v, k) => typeof want[k] === 'number' ? near(v, want[k], 0.5) : v === want[k])) wrong++;
      }
      if (notFormula) n.push(`H5:K16 가운데 수식이 아닌 셀 ${notFormula}개(수식 미작성은 실격).`);
      if (noIf) n.push('설치위치·비용 계산 열에 IF 함수를 쓰지 않은 행이 있습니다.');
      if (wrong) n.push(`계산 결과가 기대값과 다른 행 ${wrong}개.`);
      out.push({ ok: !n.length, notes: n });
    }
    // 3 집계 행·함수식 표시
    {
      const n = [];
      ['I', 'J', 'K'].forEach((c, k) => { if (!near(gv(17, 9 + k), e['평균'][k], 0.01)) n.push(`${c}17 평균이 기대값(${Math.round(e['평균'][k])})과 다릅니다.`); });
      const chk = (r, want, fns, label) => {
        const cl = cell(r, 11);
        const v = gv(r, 11);
        if (!(typeof want === 'number' ? near(v, want, 0.5) : v === want)) n.push(`${label} K${r} 값이 기대값(${want})과 다릅니다.`);
        if (!cl || !cl.f || !FX.functionsUsed(cl.f).some(f => fns.includes(f))) n.push(`${label} K${r}에 지정 함수(${fns.join('·')})를 쓰지 않았습니다.`);
        const txt = cell(r + 1, 2);
        if (!txt || txt.f || typeof txt.v !== 'string' || !txt.v.trim().startsWith('=')) n.push(`${label} 함수식 표시 B${r + 1}이 글자로 입력되어 있지 않습니다(작은따옴표).`);
        else if (cl && cl.f && NT(txt.v) !== NT(cl.f)) n.push(`B${r + 1}의 함수식이 K${r}의 최종 수식과 다릅니다.`);
      };
      chk(18, e['⑥'], ['SUMIF', 'SUMIFS'], '⑥');
      chk(20, e['⑧'], ['COUNTIF'], '⑧');
      chk(22, e['⑩'], ['SUMIF', 'SUMIFS'], '⑩');
      out.push({ ok: !n.length, notes: n });
    }
    // 4 정렬
    {
      const n = [];
      const names = []; for (let r = 5; r <= 16; r++) names.push(gv(r, 3));
      if (names.join('|') !== e['정렬_사람'].join('|')) n.push('5~16행의 순서가 정렬 기준(1차·2차)과 다릅니다.');
      if (gv(17, 2) !== '평균') n.push('17행 평균 행이 제자리에 없습니다(정렬 범위에 집계 행이 들어갔을 수 있습니다).');
      out.push({ ok: !n.length, notes: n });
    }
    // 5 서식
    {
      const n = [];
      const money = [];
      for (let r = 5; r <= 17; r++) for (const c of [9, 10, 11]) money.push([r, c]);
      money.push([18, 11], [22, 11]);
      const badMoney = money.filter(([r, c]) => !/^-?₩[\d,]+$/.test(disp(r, c)) || /\./.test(disp(r, c)));
      if (badMoney.length) n.push(`금액 셀 ${badMoney.length}개가 ₩·천 단위 쉼표·소수 없음 형식이 아닙니다(예: ${U.addr(badMoney[0][0], badMoney[0][1])} ${disp(...badMoney[0])}).`);
      for (let r = 5; r <= 16; r++) { if (!/^\d+%$/.test(disp(r, 7))) { n.push('할인율·할증률 열(G)이 백분율로 표시되지 않습니다.'); break; } }
      for (let r = 5; r <= 16; r++) for (const c of [4, 5, 6]) if (/[,.₩]/.test(disp(r, c))) { n.push('금액 외 수치(D~F 열)에 원화 기호·쉼표·소수가 있습니다(병합 셀 때문에 선택이 넓어졌는지 확인).'); r = 99; break; }
      const ha = (r, c) => XL.styleOf(sh, r, c).ha;
      const centerTxt = []; for (let r = 5; r <= 16; r++) for (const c of [2, 3, 8]) if (ha(r, c) !== 'center') centerTxt.push(U.addr(r, c));
      if (centerTxt.length) n.push(`문자 셀 가운데 맞춤이 아닌 셀 ${centerTxt.length}개(예: ${centerTxt[0]}).`);
      const leftNum = []; for (let r = 5; r <= 16; r++) for (let c = 4; c <= 11; c++) if (c !== 8 && (ha(r, c) === 'left' || ha(r, c) === 'center')) leftNum.push(U.addr(r, c));
      if (leftNum.length) n.push(`숫자 셀 오른쪽 맞춤이 아닌 셀 ${leftNum.length}개(예: ${leftNum[0]}).`);
      out.push({ ok: !n.length, notes: n });
    }
    // 6 그래프
    const ch = sh.charts[0];
    {
      const n = [];
      if (!ch) n.push('그래프가 없습니다.');
      else {
        const rows = e['차트_합'].length, end = 4 + rows;
        const nm = s => s.name && U.parseRange(s.name);
        const bySrc = col => ch.series.find(s => { const g = nm(s); return g && g.c1 === col && g.r1 === 4; }) || ch.series.find(s => { const g = U.parseRange(s.val); return g && g.c1 === col; });
        const k = bySrc(11), i = bySrc(9);
        if (!k || !i) n.push('그래프 계열이 지시된 두 열(K·I)과 다릅니다.');
        else {
          if (k.type !== 'col') n.push(`${XL.seriesName(wb, k, 0)} 계열이 묶은 세로 막대형이 아닙니다.`);
          if (i.type !== 'lineMarkers') n.push(`${XL.seriesName(wb, i, 1)} 계열이 표식이 있는 꺾은선형이 아닙니다.`);
          if (ch.series.some(s => s.secondary)) n.push('보조 축이 켜져 있습니다(단일축이어야 합니다).');
          if (ch.series.length !== 2) n.push(`계열이 ${ch.series.length}개입니다(2개여야 합니다).`);
          for (const s of [k, i]) {
            const g = U.parseRange(s.val), c = s.cat && U.parseRange(s.cat);
            if (!g || g.r1 !== 5 || g.r2 !== end) n.push(`${XL.seriesName(wb, s, 0)} 값 범위가 5~${end}행이 아닙니다(데이터 범위 설정 오류는 실격).`);
            if (!c || c.c1 !== 3 || c.r1 !== 5 || c.r2 !== end) n.push(`가로 축 항목 범위가 C5:C${end}가 아닙니다.`);
            if (!s.name) n.push('계열 이름(범례)이 머리글과 연결되어 있지 않습니다.');
          }
          if (i.labels !== 'top') n.push(`${XL.seriesName(wb, i, 1)} 계열에 값 레이블이 위쪽으로 없습니다.`);
          if (k.labels) n.push(`${XL.seriesName(wb, k, 0)} 계열에도 데이터 레이블이 있습니다(지시된 계열에만).`);
        }
        if (XL.chartTitleText(wb, ch) !== sp.chart_title || !ch.title.show) n.push(`차트 제목이 「${sp.chart_title}」이 아닙니다.`);
        if (!ch.axisTitle.x.show || ch.axisTitle.x.text !== sp.x_title) n.push(`가로 축 제목이 「${sp.x_title}」이 아닙니다.`);
        if (!ch.axisTitle.y.show || ch.axisTitle.y.text !== sp.y_title) n.push(`세로 축 제목이 「${sp.y_title}」이 아닙니다.`);
        if (!ch.legend.show) n.push('범례가 없습니다.');
      }
      out.push({ ok: !n.length, notes: n });
    }
    // 7 배치·한 장
    const lay = XL.printLayout(wb, S);
    {
      const n = [];
      if (ch) {
        const left = XL.colLeft(sh, 2), right = XL.colLeft(sh, 12);
        if (Math.abs(ch.x - left) > 12 || Math.abs(ch.x + ch.w - right) > 16) n.push('그래프 너비가 작업표(B~K 열) 너비와 맞지 않습니다.');
        if (ch.y < XL.rowTop(sh, 24)) n.push('그래프가 작업표 아래(24행 이후)에 있지 않습니다.');
        const area = sh.printArea ? U.parseRange(sh.printArea) : null;
        if (area) { const top = XL.rowTop(sh, area.r1), bottom = XL.rowTop(sh, area.r2 + 1), l = XL.colLeft(sh, area.c1), r = XL.colLeft(sh, area.c2 + 1); if (ch.y + ch.h > bottom + 2 || ch.x < l - 2 || ch.x + ch.w > r + 2 || ch.y < top) n.push('그래프가 인쇄 영역 밖으로 나갑니다.'); }
        if (area && (area.r1 > 2 || area.c1 > 2 || area.c2 < 11 || area.r2 < 23)) n.push('인쇄 영역에 작업표 전체가 들어 있지 않습니다.');
      } else n.push('그래프가 없습니다.');
      if (lay.pages.length !== 1) n.push(`미리 보기 쪽 수가 ${lay.pages.length}입니다.`);
      out.push({ ok: !n.length, notes: n });
    }
    // 8 페이지 설정
    out.push(pageSetupCheck(sh, lay));
    return out;
  }
  function pageSetupCheck(sh, lay) {
    const n = [], pg = sh.page;
    if (pg.paper !== 'A4') n.push('용지가 A4 가 아닙니다.');
    if (pg.orient !== 'portrait') n.push('용지 방향이 세로가 아닙니다.');
    if (Math.abs(pg.top - 6) > 0.05) n.push(`위쪽 여백이 ${pg.top}cm 입니다(6cm).`);
    if (!pg.centerH) n.push('페이지 가운데 맞춤(가로)이 꺼져 있습니다.');
    if (lay.pages.length !== 1) n.push(`미리 보기 쪽 수가 ${lay.pages.length}입니다(1이어야 합니다).`);
    return { ok: !n.length, notes: n };
  }
  function excel52(set, wb) {
    const XL = OA.xl, PV = OA.pivot, pvs = set.sp.pivot;
    const sh = wb && wb.sheets.find(s => s.pivots.length >= 1);
    const fail = msg => [0, 1, 2, 3, 4].map(() => ({ ok: false, notes: [msg] }));
    if (!sh) return fail('피벗테이블이 있는 시트가 없습니다.');
    const pvList = sh.pivots.slice().sort((a, b) => a.r - b.r);
    const [p1, p2] = pvList;
    const out = [];
    // 1 원본 범위
    {
      const n = [];
      for (const p of pvList.slice(0, 2)) {
        const c = PV.cacheOf(wb, p);
        const g = U.parseRange(c.ref);
        if (!g || g.r1 !== 4 || g.r2 !== 16 || g.c1 !== 2 || g.c2 !== 11) n.push(`${p.name} 원본 범위가 ${c.sheet}!${c.ref.replace(/\$/g, '')} 입니다(B4:K16).`);
      }
      out.push({ ok: !n.length, notes: n });
    }
    const gField = p => { const c = PV.cacheOf(wb, p); const gd = c.groups.find(x => x.base === pvs.row1); return gd ? gd.field : null; };
    // 2 필드·합계 이름·쉼표, 총합계 값
    {
      const n = [];
      for (const p of pvList.slice(0, 2)) {
        const rows = p.rows.filter(f => f !== 'Σ값');
        const want = pvs.kind === 'group_items' ? [gField(p), pvs.row1, pvs.row2] : [pvs.row1, pvs.row2];
        if (rows.join('|') !== want.join('|')) n.push(`${p.name} 행 영역이 ${rows.map(f => PV.caption(PV.cacheOf(wb, p), p, f)).join(', ') || '(없음)'} 입니다.`);
        const vals = p.values.map(v => v.field + ':' + v.func);
        if (vals.join('|') !== pvs.vals.map(v => v + ':sum').join('|')) n.push(`${p.name} 값 영역이 ${pvs.vals.map(v => '합계 : ' + v).join(', ')}이 아닙니다.`);
        if (!p.values.every(v => v.nf && /#,##0/.test(v.nf))) n.push(`${p.name} 값에 천 단위 쉼표 형식이 없습니다.`);
        const sm = PV.summary(wb, p);
        if (p.values.length === 2 && !(near(sm.total[0], set.expected.sp['피벗_총합'][0]) && near(sm.total[1], set.expected.sp['피벗_총합'][1]))) n.push(`${p.name} 총합계가 기대값과 다릅니다.`);
      }
      if (pvList.length < 2) n.push('피벗테이블이 2개(조정1·조정2)가 아닙니다.');
      out.push({ ok: !n.length, notes: n });
    }
    // 3 그룹·이름·레이아웃
    {
      const n = [];
      for (const p of pvList.slice(0, 2)) {
        if (pvs.kind === 'group_items') {
          const c = PV.cacheOf(wb, p), gf = gField(p);
          if (!gf) n.push(`${p.name}에 그룹이 없습니다.`);
          else {
            const sm = PV.summary(wb, p);
            for (const [name, members] of Object.entries(pvs.groups)) {
              const it = sm.items[name];
              if (!it) { n.push(`${p.name}에 그룹 이름 「${name}」이 없습니다.`); continue; }
              const keys = new Set(it.recs.map(r => PV.itemKeyOf(c, pvs.row1, r)));
              const want = new Set(members.map(PV.rawKey));
              if (keys.size !== want.size || [...want].some(k => !keys.has(k))) n.push(`그룹 「${name}」의 항목이 지시와 다릅니다.`);
            }
            if (PV.caption(c, p, gf) !== pvs.caption) n.push(`${p.name} 그룹 필드 이름이 「${pvs.caption}」이 아닙니다.`);
          }
        }
        if (p.layout !== 'tabular') n.push(`${p.name} 보고서 레이아웃이 테이블 형식이 아닙니다.`);
        if (p.subtotals !== 'none') n.push(`${p.name} 부분합이 표시되어 있습니다.`);
        if (!p.grandRow) n.push(`${p.name} 총합계가 꺼져 있습니다.`);
      }
      out.push({ ok: !n.length, notes: n });
    }
    // 4 조정1·조정2 상태와 제목
    {
      const n = [];
      const ref = reference(set).wb;
      const rsh = ref.sheets.find(s => s.pivots.length);
      const [r1, r2] = rsh.pivots.slice().sort((a, b) => a.r - b.r);
      const rowsOf = (w, s, p) => { const g = PV.grid(w, s, p); return g.rows.length; };
      const visible = (w, s, p) => PV.visibleLabels(w, s, p, pvs.row2).join('|');
      if (p1 && (visible(wb, sh, p1) !== visible(ref, rsh, r1) || rowsOf(wb, sh, p1) !== rowsOf(ref, rsh, r1))) n.push('위쪽 피벗의 확장·축소 상태가 조정1 지시와 다릅니다.');
      if (p2 && (visible(wb, sh, p2) !== visible(ref, rsh, r2) || rowsOf(wb, sh, p2) !== rowsOf(ref, rsh, r2))) n.push('아래쪽 피벗의 확장·축소 상태가 조정2 지시와 다릅니다.');
      const title = (p, want) => {
        let hit = false;
        for (let r = Math.max(1, p.r - 3); r < p.r; r++) for (let c = Math.max(1, p.c - 1); c <= p.c + 4; c++) if (XL.getValue(wb, sh, r, c) === want) hit = true;
        if (!hit) n.push(`「${want}」 제목이 피벗 위에 없습니다.`);
      };
      if (p1) title(p1, '피벗테이블 조정1');
      if (p2) title(p2, '피벗테이블 조정2');
      out.push({ ok: !n.length, notes: n });
    }
    // 5 위·아래·한 장·단추 인쇄
    {
      const lay = XL.printLayout(wb, sh.name);
      const r = pageSetupCheck(sh, lay);
      if (pvList.length >= 2 && !(p1.r < p2.r)) r.notes.push('조정1이 위, 조정2가 아래가 아닙니다.');
      if (pvList.some(p => p.printButtons && p.showButtons)) r.notes.push('피벗 테이블 옵션에서 확장/축소 단추 인쇄가 켜져 있습니다.');
      const area = sh.printArea && U.parseRange(sh.printArea);
      if (pvList.length >= 2 && area) { const t2 = PV.tableRange(wb, sh, p2); if (area.r2 < t2.r2 || area.r1 > p1.r - 2) r.notes.push('인쇄 영역에 두 피벗과 제목이 모두 들어 있지 않습니다.'); }
      r.ok = !r.notes.length;
      out.push(r);
    }
    return out;
  }

  // ------------------------------------------------------------------ Access
  function access53(set, d) {
    const A = OA.acc, R = OA.report, db = set.db, e = set.expected.db;
    const fail = msg => [0, 1, 2, 3, 4].map(() => ({ ok: false, notes: [msg] }));
    if (!d) return fail('데이터베이스를 만들지 않았습니다.');
    const out = [];
    // 1 테이블
    {
      const n = [];
      const tg = d.tables['구분'], tt = d.tables[db.tname];
      if (!tg) n.push('구분 테이블이 없습니다.');
      if (!tt) n.push(`${db.tname} 테이블이 없습니다.`);
      if (tg) {
        const f1 = tg.fields.find(f => f.name === '구분코드'), f2 = tg.fields.find(f => f.name === db.gname);
        if (!f1 || f1.type !== '짧은 텍스트') n.push('구분코드 필드가 짧은 텍스트가 아닙니다.');
        if (!f2 || f2.type !== '짧은 텍스트') n.push(`${db.gname} 필드가 짧은 텍스트가 아닙니다.`);
        if (tg.pk.join() !== '구분코드') n.push('구분 테이블의 기본 키가 구분코드가 아닙니다.');
        if (tg.rows.length !== db.gubun.length) n.push(`구분 테이블이 ${tg.rows.length}행입니다(${db.gubun.length}행).`);
      }
      if (tt) {
        for (const [name, type] of db.fields) {
          const f = tt.fields.find(x => x.name === name);
          if (!f) { n.push(`${db.tname}.${name} 필드가 없습니다.`); continue; }
          if (f.type !== type) n.push(`${db.tname}.${name} 데이터 형식이 ${f.type} 입니다(${type}).`);
        }
        if (tt.pk.join() !== db.key) n.push(`${db.tname} 테이블의 기본 키가 ${db.key}가 아닙니다.`);
        if (tt.rows.length !== db.rows.length) n.push(`${db.tname} 테이블이 ${tt.rows.length}행입니다(${db.rows.length}행).`);
      }
      const rel = d.relationships.find(r => r.t1 === '구분' && r.f1 === '구분코드' && r.t2 === db.tname && r.f2 === '구분코드');
      if (!rel) n.push('구분코드 일대다 관계가 없습니다.');
      else if (!rel.ri) n.push('관계에 참조 무결성이 설정되어 있지 않습니다.');
      out.push({ ok: !n.length, notes: n });
    }
    const forms = Object.values(d.forms);
    const form = forms.find(f => f.controls.some(c => c.type === 'listbox')) || forms[0];
    const lst = form && form.controls.find(c => c.type === 'listbox');
    let data = null;
    if (lst) { try { data = R.listboxData(d, lst); } catch (err) { data = { error: err.message, rows: [], columns: [] }; } }
    // 2 조회 결과
    {
      const n = [];
      if (!form) n.push('조회화면 폼이 없습니다.');
      else if (!lst) n.push('폼에 목록 상자가 없습니다.');
      else if (data.error) n.push('목록 상자의 행 원본을 실행할 수 없습니다: ' + data.error);
      else {
        const keys = data.rows.map(r => r[0]);
        const want = e['조회'].map(r => r[0]);
        if (keys.length !== want.length) n.push(`조회 결과가 ${keys.length}행입니다(${want.length}행).`);
        else if (keys.join() !== want.join()) n.push('조회 결과의 행 또는 순서가 지시와 다릅니다.');
        else if (JSON.stringify(data.rows) !== JSON.stringify(e['조회'])) n.push('조회 결과의 열 값이 기대값과 다릅니다.');
      }
      out.push({ ok: !n.length, notes: n });
    }
    // 3 목록 상자 열
    {
      const n = [];
      if (lst && data && !data.error) {
        const want = [db.key, db.who, '구분코드', db.fee, db.gname];
        if (data.columns.slice(0, 5).join('|') !== want.join('|')) n.push(`목록 상자 필드가 ${data.columns.join(', ')} 입니다.`);
        if ((+lst.columnCount || 1) < 5) n.push(`열 개수가 ${lst.columnCount} 입니다(5).`);
        if (!lst.columnHeads) n.push('열 이름(머리글)이 꺼져 있습니다.');
        const w = R.parseWidths(lst.columnWidths, +lst.columnCount || 1, lst.width);
        if (w.reduce((s, x) => s + x, 0) > lst.width + 0.01) n.push('열 너비의 합이 목록 상자 너비보다 커서 뒤 열이 잘립니다.');
        const rowH = (lst.fontSize || 11) * 0.0353 * 1.45;
        if ((data.rows.length + (lst.columnHeads ? 1 : 0)) * rowH > lst.height + 0.05) n.push('목록 상자 높이가 결과 행 수보다 작아 일부 행이 인쇄되지 않습니다.');
      } else n.push('목록 상자가 없거나 행 원본이 없습니다.');
      out.push({ ok: !n.length, notes: n });
    }
    // 4 SQL 문
    const sqlLbl = form && form.controls.find(c => c.type === 'label' && /SELECT/i.test(c.caption || '') && /FROM/i.test(c.caption || ''));
    {
      const n = [];
      if (!sqlLbl) n.push('SQL 문을 적은 레이블이 없습니다(SQL 문 공란은 조회화면 0점).');
      else {
        const s = sqlLbl.caption.toUpperCase();
        for (const k of ['WHERE', 'INNER JOIN', 'ORDER BY']) if (!s.includes(k)) n.push(`SQL 문에 ${k}가 없습니다.`);
        try {
          const r = A.runSQL(d, sqlLbl.caption);
          if (r.needParam) n.push(`SQL 문에 알 수 없는 이름 「${r.needParam}」이 있습니다.`);
          else if (JSON.stringify(r.rows) !== JSON.stringify(e['조회'])) n.push('적어 둔 SQL 문의 실행 결과가 조회 결과와 다릅니다.');
        } catch (err) { n.push('적어 둔 SQL 문을 실행할 수 없습니다: ' + err.message); }
        const lines = String(sqlLbl.caption).split(/\r?\n/);
        const charsPerLine = Math.max(10, Math.floor(sqlLbl.width / ((sqlLbl.fontSize || 11) * 0.0353 * 0.62)));
        const need = lines.reduce((s, l) => s + Math.max(1, Math.ceil(l.length / charsPerLine)), 0) * (sqlLbl.fontSize || 11) * 0.0353 * 1.3;
        if (need > sqlLbl.height + 0.1) n.push('SQL 문 레이블 높이가 부족해 문장 일부가 보이지 않습니다.');
      }
      out.push({ ok: !n.length, notes: n });
    }
    // 5 서식·여백
    {
      const n = [];
      if (form) {
        const title = form.controls.filter(c => c.type === 'label' && c !== sqlLbl).sort((a, b) => a.top - b.top)[0];
        if (!title || title.caption !== db.form_title) n.push('제목 레이블의 문장이 지시와 다릅니다.');
        if (title && (title.fontSize !== 16 || title.fontWeight !== '굵게')) n.push('제목이 16pt 굵게가 아닙니다.');
        if (sqlLbl && !['점선', '드문 점선', '짧은 파선', '파선'].includes(sqlLbl.borderStyle)) n.push(`SQL 문 상자 테두리가 ${sqlLbl.borderStyle} 입니다(점선).`);
        if (data && !data.error && data.info && data.info[3] && !/통화/.test(data.info[3].format || '') && data.info[3].type !== '통화') n.push('목록 상자의 금액이 원화 형식이 아닙니다(필드 형식 통화).');
        if (Math.abs(form.page.top - 60) > 0.5) n.push(`위쪽 여백이 ${form.page.top}mm 입니다(60mm).`);
        if (form.page.orient !== 'portrait' || form.page.paper !== 'A4') n.push('A4 세로가 아닙니다.');
        if (form.props.recordSource) n.push('폼의 레코드 원본이 지정되어 있어 레코드 수만큼 반복 인쇄됩니다.');
        const pc = R.renderForm(d, form).pageCount;
        if (pc !== 1) n.push(`인쇄 쪽 수가 ${pc}입니다.`);
      } else n.push('조회화면 폼이 없습니다.');
      out.push({ ok: !n.length, notes: n });
    }
    return out;
  }
  function access54(set, d) {
    const R = OA.report, A = OA.acc, db = set.db, e = set.expected.db;
    const fail = msg => [0, 1, 2, 3, 4].map(() => ({ ok: false, notes: [msg] }));
    if (!d) return fail('데이터베이스를 만들지 않았습니다.');
    const rpt = Object.values(d.reports)[0];
    if (!rpt) return fail('보고서가 없습니다.');
    let out;
    try { out = R.renderReport(d, rpt); } catch (err) { return fail('보고서를 열 수 없습니다: ' + err.message); }
    if (out.error || out.needParam) return fail('보고서를 열 수 없습니다: ' + (out.error || ('매개 변수 ' + out.needParam)));
    const items = out.pages.flatMap(p => p.items);
    const details = items.filter(it => it.inst.section === 'detail');
    const res = [];
    const cols = ['구분코드', db.key, db.days, db.fee, db.calc, db.disc, db.total];
    // 1 그룹·정렬
    {
      const n = [];
      const g0 = (rpt.grouping || []).find(g => g.kind === 'group');
      if (!g0 || g0.field !== '구분코드' || g0.order !== 'asc') n.push('그룹화 기준이 구분코드 오름차순이 아닙니다.');
      const order = details.map(it => it.inst.ctx.row[db.key]);
      const want = e['보고서_행'].map(r => r[1]);
      if (order.join() !== want.join()) n.push('본문 행 순서가 구분코드·' + db.key + ' 오름차순과 다릅니다.');
      res.push({ ok: !n.length, notes: n });
    }
    // 2 계산 필드
    {
      const n = [];
      const rows = out.rows.map(r => cols.map(c => r[c])).sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]));
      for (const c of [db.calc, db.disc, db.total]) if (!out.rows.length || !(c in out.rows[0])) n.push(`레코드 원본에 ${c} 필드가 없습니다.`);
      if (!n.length && JSON.stringify(rows) !== JSON.stringify(e['보고서_행'])) {
        const bad = rows.findIndex((r, i) => JSON.stringify(r) !== JSON.stringify(e['보고서_행'][i]));
        n.push(`계산 필드 값이 기대값과 다릅니다(${bad + 1}번째 행 ${rows[bad] ? rows[bad][1] : ''}).`);
      }
      res.push({ ok: !n.length, notes: n });
    }
    // 3 합계·총평균
    {
      const n = [];
      const foot = items.filter(it => it.inst.section === 'groupFooter0');
      const sums = {};
      for (const it of foot) {
        const code = it.inst.ctx.row['구분코드'];
        sums[code] = it.controls.filter(x => /^=\s*Sum\s*\(/i.test(x.c.controlSource || '')).sort((a, b) => a.c.left - b.c.left).map(x => x.value);
        const name = (db.gubun.find(g => g[0] === code) || [])[1];
        if (!it.controls.some(x => x.text === name + ' 합계')) n.push(`그룹 바닥글에 「${name} 합계」 제목이 없습니다.`);
      }
      if (JSON.stringify(sums) !== JSON.stringify(e['보고서_그룹'])) n.push('그룹 바닥글의 합계 3개가 기대값과 다릅니다.');
      const rf = items.find(it => it.inst.section === 'reportFooter');
      const avg = rf ? rf.controls.filter(x => /^=\s*Avg\s*\(/i.test(x.c.controlSource || '')).sort((a, b) => a.c.left - b.c.left) : [];
      if (avg.length !== 3 || !avg.every((x, i) => near(x.value, e['보고서_평균'][i], 0.01))) n.push('보고서 바닥글의 총평균 3개가 기대값과 다릅니다.');
      else if (avg.some(x => /\./.test(x.text))) n.push('총평균이 소수 자릿수 0으로 표시되지 않습니다.');
      if (rf && !rf.controls.some(x => x.text === '총평균')) n.push('「총평균」 제목이 없습니다.');
      res.push({ ok: !n.length, notes: n });
    }
    // 4 열·중복 숨기기·작성일자
    {
      const n = [];
      const dt = rpt.controls.filter(c => c.section === 'detail' && c.type === 'textbox').map(c => c.controlSource);
      for (const c of cols) if (!dt.includes(c)) n.push(`본문에 ${c} 텍스트 상자가 없습니다.`);
      const codeShown = details.filter(it => it.controls.some(x => x.c.controlSource === '구분코드' && x.text)).length;
      const groupsN = Object.keys(e['보고서_그룹']).length;
      if (codeShown !== groupsN) n.push('구분코드가 그룹의 첫 행에만 표시되지 않습니다(중복 내용 숨기기).');
      const rh = items.filter(it => it.inst.section === 'reportHeader').flatMap(it => it.controls);
      if (!rh.some(x => /작성일자/.test(x.text) && /\d{4}-\d{2}-\d{2}/.test(x.text))) n.push('보고서 머리글에 「작성일자 : YYYY-MM-DD」가 없습니다.');
      if (!rh.some(x => x.text === db.report_title && x.c.fontSize === 16 && x.c.fontWeight === '굵게')) n.push(`제목 「${db.report_title}」 16pt 굵게가 없습니다.`);
      res.push({ ok: !n.length, notes: n });
    }
    // 5 선·통화·한 장·여백
    {
      const n = [];
      const lines = sec => rpt.controls.filter(c => c.section === sec && c.type === 'line' && c.height < 0.05).length;
      if (lines('pageHeader') < 2) n.push('열 제목 위아래 가로선이 없습니다.');
      if (lines('groupFooter0') < 2) n.push('합계 행 위아래 가로선이 없습니다.');
      if (lines('reportFooter') < 1) n.push('총평균 행 아래 가로선이 없습니다.');
      const money = details.flatMap(it => it.controls.filter(x => [db.fee, db.calc, db.disc, db.total].includes(x.c.controlSource) && x.text !== ''));
      if (money.some(x => !/^-?₩[\d,]+$/.test(x.text))) n.push('본문 금액이 원화·쉼표 형식이 아닙니다.');
      if (out.pageCount !== 1) n.push(`인쇄 쪽 수가 ${out.pageCount}입니다(보고서 너비 ${rpt.props.width}cm).`);
      if (Math.abs(rpt.page.top - 60) > 0.5) n.push(`위쪽 여백이 ${rpt.page.top}mm 입니다(60mm).`);
      if (rpt.page.orient !== 'portrait' || rpt.page.paper !== 'A4') n.push('A4 세로가 아닙니다.');
      res.push({ ok: !n.length, notes: n });
    }
    return res;
  }

  // ------------------------------------------------------------------ PowerPoint
  function ppt55(set, pres) {
    const P = OA.ppt;
    const fail = msg => [0, 1, 2, 3].map(() => ({ ok: false, notes: [msg] }));
    if (!pres) return fail('프레젠테이션을 만들지 않았습니다.');
    const ref = P.newPresentation();
    P.setSlideSize(ref, 'standard', 'fit'); P.setLayout(ref, 0, 'blank'); P.addSlide(ref, 0);
    set.pt.ref.forEach((s, i) => P.loadShapes(ref, i, s));
    const out = [];
    const n1 = [], n2 = [];
    if (pres.slides.length !== 2) n1.push(`슬라이드가 ${pres.slides.length}장입니다(2장).`);
    for (let i = 0; i < 2; i++) {
      if (!pres.slides[i]) continue;
      const a = P.inventory(pres, i), b = P.inventory(ref, i);
      const tag = `제${i + 1} 슬라이드`;
      if (a.top !== b.top) n1.push(`${tag} 도형 수 ${a.top}개(원안 ${b.top}개, 그룹은 1개로 셈).`);
      const kinds = new Set(b.flat.map(s => P.CATEGORY[s.kind] || s.kind).concat(a.flat.map(s => P.CATEGORY[s.kind] || s.kind)));
      for (const k of kinds) {
        if (k === 'line' || k === 'connElbow') continue;
        if (a.count(k) !== b.count(k)) n1.push(`${tag} ${P.KINDS[k] ? P.KINDS[k].name : k} ${a.count(k)}개(원안 ${b.count(k)}개).`);
      }
      const lines = inv => inv.flat.filter(s => P.isLine(s)).length;
      if (lines(a) !== lines(b)) n1.push(`${tag} 선·연결선 ${lines(a)}개(원안 ${lines(b)}개).`);
      if (pres.slides[i].shapes.some(s => s.kind === 'placeholder')) n1.push(`${tag}에 빈 개체 틀이 남아 있습니다(레이아웃 빈 화면).`);
      // 서식
      const cnt = (inv, fn) => inv.flat.filter(fn).length;
      const gray = s => !P.isLine(s) && P.isGray(s.fill);
      const shadow = s => !!s.shadow;
      const dash = s => s.line && s.line.dash && s.line.dash !== 'solid';
      const arrow = s => P.isLine(s) && s.line && (s.line.end || s.line.begin);
      const both = s => P.isLine(s) && s.line && s.line.end && s.line.begin;
      if (cnt(a, gray) !== cnt(b, gray)) n2.push(`${tag} 회색 40% 채우기 도형 ${cnt(a, gray)}개(원안 ${cnt(b, gray)}개).`);
      if (cnt(a, shadow) !== cnt(b, shadow)) n2.push(`${tag} 그림자 도형 ${cnt(a, shadow)}개(원안 ${cnt(b, shadow)}개).`);
      if (cnt(a, dash) !== cnt(b, dash)) n2.push(`${tag} 점선 테두리 ${cnt(a, dash)}개(원안 ${cnt(b, dash)}개).`);
      if (cnt(a, arrow) !== cnt(b, arrow) || cnt(a, both) !== cnt(b, both)) n2.push(`${tag} 화살표가 있는 선 ${cnt(a, arrow)}개(원안 ${cnt(b, arrow)}개, 양쪽 ${cnt(b, both)}개).`);
      const colorful = a.flat.filter(s => (s.fill && !/^#([0-9A-F]{2})\1\1$/i.test(s.fill)) || (s.line && s.line.color && !/^#([0-9A-F]{2})\1\1$/i.test(s.line.color)) || (s.text && s.text.value && s.text.color && !/^#([0-9A-F]{2})\1\1$/i.test(s.text.color)));
      if (colorful.length) n2.push(`${tag} 색이 있는 도형·선·글자 ${colorful.length}개(흑백으로).`);
      // 글머리 기호는 문단마다 앞에 붙여 비교한다(14.3: 기호를 글자로 직접 입력해도 인쇄 결과는 같다)
      const texts = inv => inv.flat.filter(s => s.text && s.text.value).map(s => U.normText(s.text.bullet ? s.text.value.split('\n').map(l => s.text.bullet + l).join('\n') : s.text.value));
      const at = texts(a);
      const missing = texts(b).filter(t => !at.includes(t));
      if (missing.length) n2.push(`${tag} 원안 글자 ${missing.length}건이 없거나 다릅니다(예: ${missing[0]}).`);
      const title = inv => inv.flat.filter(s => s.text && s.text.value).sort((x, y) => y.text.size - x.text.size)[0];
      const ta = title(a), tb = title(b);
      if (ta && tb && (ta.text.size !== tb.text.size || !!ta.text.bold !== !!tb.text.bold)) n2.push(`${tag} 제목 글자가 ${tb.text.size}pt${tb.text.bold ? ' 굵게' : ''}가 아닙니다.`);
    }
    if (pres.size.kind !== 'standard') n1.push('슬라이드 크기가 표준(4:3)이 아닙니다.');
    out.push({ ok: !n1.length, notes: n1 });
    out.push({ ok: !n2.length, notes: n2 });
    const pr = pres.print, n3 = [];
    if (pr.layout !== 'h2') n3.push(`인쇄 모양이 ${P.LAYOUT_NAMES[pr.layout]} 입니다(유인물 2슬라이드).`);
    if (!pr.frame) n3.push('슬라이드 테두리가 꺼져 있습니다.');
    if (pr.fit) n3.push('용지에 맞게 크기 조정이 켜져 있습니다.');
    if (!pr.hq) n3.push('고품질이 꺼져 있습니다.');
    if (pr.orient !== 'portrait') n3.push('방향이 세로가 아닙니다.');
    if (pr.color !== 'gray') n3.push('색이 회색조가 아닙니다.');
    if (pr.target !== 'all') n3.push('인쇄 대상이 모든 슬라이드 인쇄가 아닙니다.');
    out.push({ ok: !n3.length, notes: n3 });
    const h = pres.handout, n4 = [];
    for (const [k, nm] of [['header', '머리글'], ['date', '날짜'], ['footer', '바닥글'], ['pageNum', '페이지 번호']]) if (h[k]) n4.push(`유인물 마스터의 ${nm} 확인란이 켜져 있습니다.`);
    const pc = P.printPages(pres).length;
    if (pc !== 1) n4.push(`인쇄 미리 보기 쪽 수가 ${pc}입니다.`);
    out.push({ ok: !n4.length, notes: n4 });
    return out;
  }

  // ------------------------------------------------------------------ 전체
  /* state = {wb, db, pres, status:{excel:{dirty, printedDirty}, ...}}, sheets = 채점할 출력물(['5-1',...]) */
  function grade(DATA, set, state, sheets) {
    const list = DATA.checklist;
    const want = new Set(sheets || ['5-1', '5-2', '5-3', '5-4', '5-5', '공통']);
    const groups = { '5-1': () => excel51(set, state.wb), '5-2': () => excel52(set, state.wb), '5-3': () => access53(set, state.db), '5-4': () => access54(set, state.db), '5-5': () => ppt55(set, state.pres) };
    const results = {};
    for (const k of Object.keys(groups)) {
      const key = k;
      if (![...want].some(w => w.startsWith(key))) continue;
      try { results[key] = groups[key](); } catch (err) { results[key] = [0, 1, 2, 3, 4, 5, 6, 7].map(() => ({ ok: false, notes: ['채점 중 오류: ' + err.message] })); }
    }
    const out = [];
    const idx = {};
    for (const c of list) {
      const key = c.sheet.slice(0, 3) === '공통' ? '공통' : c.sheet.slice(0, 3);
      if (!want.has(key) && !(key === '공통' && want.has('공통'))) continue;
      idx[key] = (idx[key] || 0);
      let r;
      if (key === '공통') {
        const k = idx[key];
        if (k < 2) r = { status: 'manual', notes: ['출력물에 직접 확인하는 항목입니다.'] };
        else {
          const st = state.status || {};
          const bad = Object.entries(st).filter(([, s]) => s && (s.dirty || s.printedDirty)).map(([app, s]) => ({ excel: 'Excel', access: 'Access', ppt: 'PowerPoint' }[app] + (s.printedDirty ? ' 인쇄 뒤 수정' : ' 저장 안 함')));
          r = { status: bad.length ? 'fail' : 'pass', notes: bad.length ? ['마지막 저장·인쇄 이후 변경: ' + bad.join(', ')] : [] };
        }
      } else {
        const g = results[key] && results[key][idx[key]];
        r = g ? { status: g.ok ? 'pass' : 'fail', notes: g.notes } : { status: 'todo', notes: [] };
      }
      idx[key]++;
      out.push(Object.assign({ sheet: c.sheet, item: c.item, sec: c.sec }, r));
    }
    return out;
  }

  // ================================================================== 15장 유형 과제(출력물 없음)
  /* 기대값: 15.3 검산 상자와 verify_ch15_p2.json·verify_ch15_excel.py 실측값(set.expected) */
  function gradeCh15(set, wb) {
    const XL = OA.xl, FX = OA.formula;
    const S = '작업표';
    const sh = wb && (wb.sheets.find(s => s.name === S) || wb.sheets[0]);
    const row = (item, sec, notes) => ({ sheet: '15장', item, sec, status: notes.length ? 'fail' : 'pass', notes });
    if (!sh) return [row('작업표 시트', '15.3', ['작업표 시트가 없습니다.'])];
    const p2 = set.expected.p2, abc = set.expected.abc['실측'];
    const cell = a => { const g = U.parseAddr(a); return XL.getCell(sh, g.r, g.c); };
    const val = a => { const g = U.parseAddr(a); const v = XL.getValue(wb, sh.name, g.r, g.c); return v && v.e ? v.e : v; };
    const disp = a => { const g = U.parseAddr(a); return XL.display(wb, sh.name, g.r, g.c).text; };
    const eq = (a, b) => (typeof a === 'number' && typeof b === 'number') ? U.approx(a, b, 1e-6) : String(a) === String(b);
    const rows = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14];
    const colCheck = (n, colL, want, label, opts) => {
      const o = opts || {};
      const bad = [], noF = [];
      rows.forEach((r, i) => { const a = colL + r; const got = o.display ? disp(a) : val(a); if (!eq(got, want[i])) bad.push(`${a} ${got === '' || got === null || got === undefined ? '(빈 셀)' : got}(기대 ${want[i]})`); const cl = cell(a); if (!cl || !cl.f) noF.push(a); });
      if (bad.length) n.push(`${label} 값이 다른 셀 ${bad.length}개(예: ${bad[0]}).`);
      if (noF.length && noF.length < 10) n.push(`${label} 수식이 아닌 셀 ${noF.length}개(예: ${noF[0]}).`);
      else if (noF.length === 10) n.push(`${label}(${colL}5:${colL}14)에 수식이 없습니다.`);
      if (o.fns) { const cl = cell(colL + 5); if (cl && cl.f && !o.fns.every(f => FX.functionsUsed(cl.f).includes(f))) n.push(`${colL}5에 지정 함수(${o.fns.join('·')})를 쓰지 않았습니다.`); }
    };
    const one = (n, a, want, label, opts) => {
      const o = opts || {};
      const got = o.display ? disp(a) : val(a);
      if (!eq(got, want)) n.push(`${label} ${a} ${got === '' || got === null || got === undefined ? '(빈 셀)' : got}(기대 ${want}).`);
      const cl = cell(a);
      if (!cl || !cl.f) n.push(`${a} 셀이 수식이 아닙니다.`);
      else if (o.fns && !o.fns.some(f => FX.functionsUsed(cl.f).includes(f))) n.push(`${a} 셀에 지정 함수(${o.fns.join('·')})를 쓰지 않았습니다.`);
    };
    const out = [];
    { // 과제 A ①~④
      const n = [];
      colCheck(n, 'H', rows.map(r => abc['행' + r]['점검시간표시']), '점검시간(H)', { display: true });
      colCheck(n, 'I', rows.map(r => abc['행' + r]['인건비']), '인건비(I)', { fns: ['HOUR', 'MINUTE'] });
      colCheck(n, 'J', rows.map(r => abc['행' + r]['할증']), '할증(J)', { fns: ['IF', 'HOUR'] });
      colCheck(n, 'K', rows.map(r => abc['행' + r]['청구액']), '청구액(K)');
      out.push(row('과제 A 시간 계산: 점검시간 h:mm, HOUR·MINUTE 인건비, 3시간 이상 할증, 청구액', '15.3', n));
    }
    { // 과제 B ⑤⑥
      const n = [];
      one(n, 'K17', abc.R, 'R 포함 합계', { fns: ['SUMPRODUCT'] });
      one(n, 'K18', abc.S, 'S 포함 합계', { fns: ['SUMPRODUCT'] });
      one(n, 'K19', abc.W, 'W 포함 합계', { fns: ['SUMPRODUCT'] });
      one(n, 'K20', abc.RS, 'R 또는 S 포함 합계', { fns: ['SUMPRODUCT'] });
      const k21 = cell('K21');
      if (k21 && (k21.f || k21.v !== undefined) && !eq(val('K21'), abc.R_wild)) n.push(`와일드카드 검산 K21 ${val('K21')}(기대 ${abc.R_wild}).`);
      out.push(row('과제 B 코드 포함 조건 합계: SUMPRODUCT(ISNUMBER(FIND())), 두 조건은 +', '15.3', n));
    }
    { // 과제 C ⑦
      const n = [];
      const want = abc['차트 계열 값'];
      const ch = (sh.charts || [])[0];
      if (!ch) n.push('그래프가 없습니다.');
      else {
        if (ch.series.length !== 1) n.push(`계열이 ${ch.series.length}개입니다(사본 범위로 그리면 1개). 병합된 집계 행을 그대로 범위로 잡았는지 확인합니다.`);
        const s0 = ch.series[0];
        if (s0) {
          if (s0.type !== 'col') n.push('묶은 세로 막대형이 아닙니다.');
          const vals = XL.refValues(wb, s0.val);
          if (vals.length !== 3 || vals.some((v, i) => !eq(v, want[i]))) n.push(`계열 값 ${vals.join(', ')}(기대 ${want.join(', ')}).`);
          const g = U.parseRange(s0.val);
          if (g && g.r1 >= 17 && g.r2 <= 19 && g.c1 >= 2 && g.c2 <= 10) n.push('계열 범위가 병합된 집계 셀(B~J 열) 안에 있습니다. 차트용 사본 범위를 만듭니다.');
          if (!s0.cat) n.push('가로 축 항목(구분 이름) 범위가 없습니다.');
        }
      }
      out.push(row('과제 C 집계 결과 그래프: 차트용 사본 범위(예: M16:N19)로 묶은 세로 막대형', '15.3', n));
    }
    { // 과제 D ⑧~⑪
      const n = [];
      colCheck(n, 'L', p2['순위'], '순위(L)', { fns: ['RANK.EQ'] });
      one(n, 'K22', p2['김씨 평균'], '김씨 평균', { fns: ['AVERAGEIF', 'AVERAGEIFS'] });
      one(n, 'K23', p2['R 10만 이상 건수'], 'R 10만 이상 건수', { fns: ['COUNTIFS'] });
      one(n, 'K24', p2['3위 이내 합'], '3위 이내 합', { fns: ['SUMIFS', 'SUMIF'] });
      if (disp('K22') !== p2['표시'].K22) n.push(`K22 표시 ${disp('K22')}(기대 ${p2['표시'].K22}).`);
      if (disp('K24') !== p2['표시'].K24) n.push(`K24 표시 ${disp('K24')}(기대 ${p2['표시'].K24}).`);
      out.push(row('과제 D 순위·조건 평균·조건 개수·순위 조건 합(절대 참조)', '15.3', n));
    }
    { // 과제 E ⑫~⑭
      const n = [];
      colCheck(n, 'M', p2['구분'], '구분(M)', { fns: ['IF', 'LEFT'] });
      colCheck(n, 'N', p2['설치 위치'], '설치 위치(N)', { fns: ['CHOOSE', 'MID'] });
      colCheck(n, 'O', p2['기본료'], '기본료(O)', { fns: ['VLOOKUP'] });
      out.push(row('과제 E 코드 해석: LEFT·MID, CHOOSE, VLOOKUP 정확히 일치', '15.3', n));
    }
    { // 과제 F ⑮~⑱
      const n = [];
      colCheck(n, 'R', p2['사용개월'], '사용개월(R)', { fns: ['DATEDIF'] });
      colCheck(n, 'S', p2['사용기간'], '사용기간(S)');
      colCheck(n, 'T', p2['다음 점검일'], '다음 점검일(T)', { display: true, fns: ['DATE'] });
      colCheck(n, 'U', p2['반올림'], '반올림(U)', { fns: ['ROUND'] });
      one(n, 'K25', p2['2024년 설치'], '2024년 설치 건수', { fns: ['COUNTIFS'] });
      if (disp('U5') !== p2['표시'].U5) n.push(`U5 표시 ${disp('U5')}(기대 ${p2['표시'].U5}).`);
      out.push(row('과제 F 날짜 계산: DATEDIF 개월, 년·개월 표시, 6개월 뒤 날짜, 천 원 반올림, 연도 조건 개수', '15.3', n));
    }
    return out;
  }
  OA.gradeCh15 = gradeCh15;

  OA.grade = { grade, excel51, excel52, access53, access54, ppt55, reference, gradeCh15 };
})(typeof window !== 'undefined' ? window : globalThis);

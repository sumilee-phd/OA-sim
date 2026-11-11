/* Excel 대화상자와 피벗 테이블 필드 창 */
(function (G) {
  'use strict';
  const OA = G.OA = G.OA || {};
  const { h, icon, menu, dialog, msg, tabs, refEdit, toast } = OA.ui;
  const U = OA.util, XL = OA.xl, NF = OA.numfmt, PV = OA.pivot;
  const CUR = '₩#,##0;-₩#,##0';

  OA.xlDialogs = function (app, api) {
    const wb = () => api.wb(), sh = () => api.sh(), sel = api.sel;
    const D = {};
    const lst = (items, cur, onPick, height) => {
      const el = h('div', { class: 'lst', style: { height: (height || 160) + 'px' } });
      const paint = v => { [...el.children].forEach(c => c.classList.toggle('on', c.dataset.v === String(v))); };
      for (const it of items) { const [v, l] = Array.isArray(it) ? it : [it, it]; const d = h('div', { 'data-v': String(v) }, l); d.addEventListener('click', () => { paint(v); onPick(v); }); d.addEventListener('dblclick', () => { const b = el.closest('.dlg') && el.closest('.dlg').querySelector('.db.primary'); if (b) b.click(); }); el.appendChild(d); }
      paint(cur);
      el.set = v => paint(v);
      return el;
    };

    // ------------------------------------------------------------------ 셀 서식
    function numberTab(initialNf, onChange) {
      let nf = initialNf || NF.GENERAL;
      const catOf = f => {
        if (NF.isGeneral(f)) return '일반';
        if (f === '@') return '텍스트';
        if (/^0%$|^0\.0+%$/.test(f)) return '백분율';
        if (/^₩#,##0(\.0+)?;-₩#,##0(\.0+)?$/.test(f)) return '통화';
        if (/^_-₩\*/.test(f)) return '회계';
        if (/^(#,##0|0)(\.0+)?$/.test(f)) return '숫자';
        if (['yyyy-mm-dd', 'yyyy"년" m"월" d"일"', 'm/d'].includes(f)) return '날짜';
        if (['h:mm', 'h:mm:ss', 'h:mm AM/PM', '[h]:mm'].includes(f)) return '시간';
        return '사용자 지정';
      };
      let cat = catOf(nf);
      const sample = h('div', { style: { border: '1px solid #c8c6c4', padding: '4px 6px', minHeight: '24px', marginBottom: '8px' } });
      const opts = h('div');
      const v0 = XL.getValue(wb(), sh(), sel.r, sel.c);
      const upd = () => { sample.textContent = v0 === null ? '' : NF.format(v0, nf).text; onChange(nf); };
      const dec = (n) => n > 0 ? '.' + '0'.repeat(n) : '';
      const decOf = f => { const m = /\.(0+)/.exec(f); return m ? m[1].length : 0; };
      const spin = (val, on) => { const i = h('input', { type: 'number', min: 0, max: 30, value: val, style: { width: '60px' } }); i.addEventListener('input', () => on(+i.value || 0)); return i; };
      const drawOpts = () => {
        opts.innerHTML = '';
        if (cat === '일반') opts.append(h('p', { class: 'note' }, '일반 서식 셀에는 특정 서식이 지정되지 않습니다.'));
        if (cat === '숫자') {
          let d = decOf(nf), comma = /#,##0/.test(nf);
          const mk = () => { nf = (comma ? '#,##0' : '0') + dec(d); upd(); };
          const cb = h('input', { type: 'checkbox', checked: comma, onchange: () => { comma = cb.checked; mk(); } });
          opts.append(h('div', { class: 'row' }, '소수 자릿수(D):', spin(d, v => { d = v; mk(); })), h('label', { class: 'row' }, cb, '1000 단위 구분 기호(,) 사용(U)'),
            h('div', {}, '음수(N):'), lst([['-', '-1234'], ['(', '(1234)']], '-', () => {}, 60));
        }
        if (cat === '통화' || cat === '회계') {
          let d = decOf(nf);
          let sym = '₩';
          const mk = () => { const s = sym === '없음' ? '' : sym; nf = cat === '통화' ? `${s}#,##0${dec(d)};-${s}#,##0${dec(d)}` : `_-${s}* #,##0${dec(d)}_-;-${s}* #,##0${dec(d)}_-;_-${s}* "-"_-;_-@_-`; upd(); };
          const symSel = h('select', { onchange: () => { sym = symSel.value; mk(); } }, ['₩', '$', '없음'].map(x => h('option', {}, x)));
          opts.append(h('div', { class: 'row' }, '소수 자릿수(D):', spin(d, v => { d = v; mk(); })), h('div', { class: 'row' }, '기호(S):', symSel));
          if (cat === '통화') opts.append(h('div', {}, '음수(N):'), lst([['-', '-₩1,234'], ['r', '₩1,234 (빨강)']], '-', () => {}, 60));
          if (catOf(nf) !== cat) mk();
        }
        if (cat === '날짜') opts.append(h('div', {}, '형식(T):'), lst([['yyyy-mm-dd', '2012-03-14'], ['yyyy"년" m"월" d"일"', '2012년 3월 14일'], ['m/d', '3/14']], nf, v => { nf = v; upd(); }, 100));
        if (cat === '시간') opts.append(h('div', {}, '형식(T):'), lst([['h:mm', '13:30'], ['h:mm:ss', '13:30:55'], ['h:mm AM/PM', '1:30 PM'], ['[h]:mm', '37:30']], nf, v => { nf = v; upd(); }, 100));
        if (cat === '백분율') { let d = decOf(nf); opts.append(h('div', { class: 'row' }, '소수 자릿수(D):', spin(d, v => { d = v; nf = '0' + dec(d) + '%'; upd(); }))); if (catOf(nf) !== '백분율') { nf = '0%'; upd(); } }
        if (cat === '텍스트') { nf = '@'; opts.append(h('p', { class: 'note' }, '텍스트 서식 셀에서는 숫자도 텍스트로 처리됩니다.')); }
        if (cat === '분수') nf = '# ?/?';
        if (cat === '지수') nf = '0.00E+00';
        if (cat === '사용자 지정') {
          const inp = h('input', { type: 'text', value: nf, style: { width: '100%' }, 'data-noenter': '' });
          inp.dataset.noenter = '';
          inp.addEventListener('input', () => { nf = inp.value; upd(); });
          inp.addEventListener('keydown', e => { if (e.key === 'Enter') { nf = inp.value; upd(); } });
          const l = lst(NF.CUSTOM_LIST, nf, v => { nf = v; inp.value = v; upd(); }, 170);
          opts.append(h('div', {}, '형식(T):'), inp, l, h('p', { class: 'note' }, '기존 형식 중 하나를 선택한 다음 편집하여 사용자 지정 형식을 만듭니다.'));
        }
        upd();
      };
      const cats = lst(NF.CATEGORIES.map(c => c.id), cat, v => {
        cat = v;
        const s = NF.CATEGORIES.find(c => c.id === v).sample;
        if (s !== null && catOf(nf) !== v) nf = s;
        if (v === '사용자 지정' && NF.isGeneral(nf)) nf = NF.GENERAL;
        drawOpts();
      }, 250);
      drawOpts();
      return h('div', { style: { display: 'flex', gap: '12px', minHeight: '300px' } }, h('div', { style: { width: '130px' } }, h('div', {}, '범주(C):'), cats),
        h('div', { style: { flex: 1 } }, h('div', {}, '보기'), sample, opts));
    }
    D.formatCells = (tab) => {
      const st = XL.styleOf(sh(), sel.r, sel.c);
      let nf = st.nf || NF.GENERAL;
      const changes = {};
      const ha = h('select', {}, [['', '일반'], ['left', '왼쪽(들여쓰기)'], ['center', '가운데'], ['right', '오른쪽(들여쓰기)'], ['fill', '채우기'], ['justify', '양쪽 맞춤'], ['centerAcross', '선택 영역의 가운데로'], ['distributed', '균등 분할(들여쓰기)']].map(([v, l]) => h('option', { value: v, selected: (st.ha || '') === v }, l)));
      const va = h('select', {}, [['top', '위쪽'], ['middle', '가운데'], ['bottom', '아래쪽'], ['justify', '양쪽 맞춤'], ['distributed', '균등 분할']].map(([v, l]) => h('option', { value: v, selected: (st.va || 'bottom') === v }, l)));
      const wrap = h('input', { type: 'checkbox', checked: !!st.wrap });
      const shrink = h('input', { type: 'checkbox', disabled: true });
      const isMerged = !!XL.mergeAt(sh(), sel.r, sel.c);
      const mergeCb = h('input', { type: 'checkbox', checked: isMerged });
      const fontSel = lst(['맑은 고딕', '굴림', '돋움', '바탕', 'Arial'], st.font || '맑은 고딕', v => { changes.font = v; }, 110);
      const style = lst([['n', '보통'], ['i', '기울임꼴'], ['b', '굵게'], ['bi', '굵은 기울임꼴']], (st.b ? 'b' : '') + (st.i ? 'i' : '') || 'n', v => { changes.b = v.includes('b'); changes.i = v.includes('i'); }, 110);
      const sizeIn = h('input', { type: 'text', value: st.sz || 11, style: { width: '60px' } });
      const sizeSel = lst([8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36], st.sz || 11, v => { sizeIn.value = v; }, 110);
      const underline = h('select', {}, h('option', { value: '' }, '없음'), h('option', { value: 'u', selected: !!st.u }, '실선'));
      let border = null;
      const bBtn = (k, label, ic) => h('button', { class: 'btn', style: { display: 'flex', flexDirection: 'column', alignItems: 'center', width: '70px' }, onclick: () => { border = k; bInfo.textContent = '적용할 테두리: ' + label; } }, icon(ic, 28), label);
      const bInfo = h('div', { class: 'note' }, '미리 설정 단추를 눌러 테두리를 지정합니다.');
      let fill;
      const t = tabs([
        { id: 'number', label: '표시 형식', body: numberTab(nf, v => { nf = v; }) },
        { id: 'align', label: '맞춤', body: h('div', { style: { minHeight: '300px' } }, h('fieldset', {}, h('legend', {}, '텍스트 맞춤'), h('div', { class: 'grid2' }, '가로(H):', ha, '세로(V):', va)),
          h('fieldset', {}, h('legend', {}, '텍스트 조정'), h('label', { class: 'row' }, wrap, '텍스트 줄 바꿈(W)'), h('label', { class: 'row' }, shrink, '셀에 맞춤(K)'), h('label', { class: 'row' }, mergeCb, '셀 병합(M)'))) },
        { id: 'font', label: '글꼴', body: h('div', { style: { display: 'flex', gap: '10px', minHeight: '300px' } }, h('div', {}, '글꼴(F):', fontSel), h('div', {}, '글꼴 스타일(O):', style), h('div', {}, '크기(S):', sizeIn, sizeSel), h('div', {}, '밑줄(U):', underline)) },
        { id: 'border', label: '테두리', body: h('div', { style: { minHeight: '300px' } }, h('div', {}, '미리 설정'), h('div', { class: 'row' }, bBtn('none', '없음', 'border'), bBtn('outside', '윤곽선', 'rectC'), bBtn('all', '안쪽', 'borderAll')), bInfo,
          h('p', { class: 'note' }, '윤곽선은 선택 범위의 바깥 테두리, 안쪽은 안쪽 선까지(모든 테두리와 같음) 지정합니다.')) },
        { id: 'fill', label: '채우기', body: h('div', { style: { minHeight: '300px' } }, h('div', {}, '배경색(C):'), h('div', { class: 'row' }, h('button', { class: 'btn', onclick: () => { fill = null; } }, '색 없음')),
          h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(10, 20px)', gap: '3px' } }, ['#FFFFFF', '#F2F2F2', '#D9D9D9', '#BFBFBF', '#A6A6A6', '#DDEBF7', '#FCE4D6', '#E2EFDA', '#FFF2CC', '#FFFF00'].map(c => h('div', { style: { width: '20px', height: '20px', background: c, border: '1px solid #aaa', cursor: 'pointer' }, onclick: () => { fill = c; } })))) },
        { id: 'protect', label: '보호', body: h('div', { style: { minHeight: '300px' } }, h('label', { class: 'row' }, h('input', { type: 'checkbox', checked: true }), '잠금(L)'), h('label', { class: 'row' }, h('input', { type: 'checkbox' }), '숨김(I)'), h('p', { class: 'note' }, '워크시트를 보호해야 셀 잠금 또는 수식 숨기기가 적용됩니다.')) },
      ], tab || 'number');
      dialog({ title: '셀 서식', width: 560, body: t, buttons: [{ label: '확인', primary: true, onClick: () => {
        api.exec(() => {
          const patch = Object.assign({ nf: NF.isGeneral(nf) ? null : nf, ha: ha.value || null, va: va.value === 'bottom' ? null : va.value, wrap: wrap.checked || null, u: underline.value ? true : null }, changes);
          const sz = parseFloat(sizeIn.value); if (sz > 0) patch.sz = sz === 11 && !st.sz ? null : sz;
          if (patch.b === false) patch.b = null; if (patch.i === false) patch.i = null;
          if (fill !== undefined) patch.fill = fill;
          XL.applyStyle(wb(), sh(), sel.ranges, patch);
          if (border) for (const g of sel.ranges) XL.setBorder(wb(), sh(), g, border);
          if (mergeCb.checked !== isMerged) for (const g of sel.ranges) XL.merge(wb(), sh(), g, mergeCb.checked ? 'merge' : 'unmerge');
        });
      } }, { label: '취소', cancel: true }] });
    };

    // ------------------------------------------------------------------ 정렬
    D.sort = () => {
      let g = api.act();
      const proceed = (rg) => {
        api.setSel(rg.r1, rg.c1, rg.r2, rg.c2); api.render();
        const guessHeader = (() => { for (let c = rg.c1; c <= rg.c2; c++) { const a = XL.getValue(wb(), sh(), rg.r1, c), b = XL.getValue(wb(), sh(), rg.r1 + 1, c); if (typeof a === 'string' && (typeof b === 'number')) return true; } return false; })();
        const hdr = h('input', { type: 'checkbox', checked: guessHeader });
        const cols = () => { const out = []; for (let c = rg.c1; c <= rg.c2; c++) { const v = XL.getValue(wb(), sh(), rg.r1, c); out.push([c, hdr.checked && v !== null && v !== '' ? OA.formula.toStr(v) : '열 ' + U.colName(c)]); } return out; };
        let levels = [{ col: sel.c >= rg.c1 && sel.c <= rg.c2 ? sel.c : rg.c1, order: 'asc' }];
        let cur = 0;
        const body = h('div', { class: 'lst', style: { height: '170px', padding: 0 } });
        const draw = () => {
          body.innerHTML = '';
          body.appendChild(h('div', { style: { display: 'grid', gridTemplateColumns: '84px 1fr 1fr 1fr', background: '#fff', padding: '4px', borderBottom: '1px solid #ddd', cursor: 'default' } }, h('span'), h('span', {}, '세로 막대형'), h('span', {}, '정렬 기준'), h('span', {}, '정렬')));
          levels.forEach((lv, i) => {
            const colSel = h('select', { style: { width: '100%' }, onchange: () => { lv.col = +colSel.value; } }, cols().map(([c, l]) => h('option', { value: c, selected: lv.col === c }, l)));
            const ordSel = h('select', { style: { width: '100%' }, onchange: () => { lv.order = ordSel.value; } }, h('option', { value: 'asc', selected: lv.order === 'asc' }, '오름차순'), h('option', { value: 'desc', selected: lv.order === 'desc' }, '내림차순'), h('option', { value: 'custom', disabled: true }, '사용자 지정 목록...'));
            const row = h('div', { style: { display: 'grid', gridTemplateColumns: '84px 1fr 1fr 1fr', gap: '4px', padding: '3px 4px', background: i === cur ? '#e1dfdd' : '#fff', alignItems: 'center' }, onclick: () => { cur = i; [...body.children].forEach((c, k) => { if (k) c.style.background = k - 1 === cur ? '#e1dfdd' : '#fff'; }); } },
              h('span', {}, i === 0 ? '정렬 기준' : '다음 기준'), colSel, h('select', { style: { width: '100%' } }, h('option', {}, '셀 값'), h('option', { disabled: true }, '셀 색'), h('option', { disabled: true }, '글꼴 색'), h('option', { disabled: true }, '조건부 서식 아이콘')), ordSel);
            body.appendChild(row);
          });
        };
        hdr.addEventListener('change', draw);
        draw();
        const tb = h('div', { class: 'row' },
          h('button', { class: 'btn', onclick: () => { levels.push({ col: rg.c1, order: 'asc' }); cur = levels.length - 1; draw(); } }, '＋ 기준 추가(A)'),
          h('button', { class: 'btn', onclick: () => { if (levels.length > 1) { levels.splice(cur, 1); cur = Math.max(0, cur - 1); draw(); } } }, '✕ 기준 삭제(D)'),
          h('button', { class: 'btn', onclick: () => { levels.splice(cur + 1, 0, Object.assign({}, levels[cur])); cur++; draw(); } }, '기준 복사(C)'),
          h('button', { class: 'btn', title: '위로 이동', onclick: () => { if (cur > 0) { [levels[cur - 1], levels[cur]] = [levels[cur], levels[cur - 1]]; cur--; draw(); } } }, '▲'),
          h('button', { class: 'btn', title: '아래로 이동', onclick: () => { if (cur < levels.length - 1) { [levels[cur + 1], levels[cur]] = [levels[cur], levels[cur + 1]]; cur++; draw(); } } }, '▼'),
          h('button', { class: 'btn', onclick: () => dialog({ title: '정렬 옵션', width: 300, body: h('div', {}, h('label', { class: 'row' }, h('input', { type: 'checkbox' }), '대/소문자 구분(C)'), h('fieldset', {}, h('legend', {}, '방향'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'so', checked: true }), '위쪽에서 아래쪽(T)'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'so', disabled: true }), '왼쪽에서 오른쪽(L)'))) }) }, '옵션(O)...'),
          h('label', { class: 'row', style: { marginLeft: 'auto' } }, hdr, '내 데이터에 머리글 표시(H)'));
        dialog({ title: '정렬', width: 720, body: h('div', {}, tb, body), buttons: [{ label: '확인', primary: true, onClick: () => { api.exec(() => XL.sortRange(wb(), sh(), rg, levels.map(l => ({ col: l.col, order: l.order })), hdr.checked)); } }, { label: '취소', cancel: true }] });
      };
      if (g.r1 === g.r2 && g.c1 === g.c2) return proceed(XL.currentRegion(wb(), sh(), sel.r, sel.c));
      const region = XL.currentRegion(wb(), sh(), g.r1, g.c1);
      if ((g.c2 - g.c1) < (region.c2 - region.c1) && g.r1 >= region.r1 && g.r2 <= region.r2 && g.c1 === g.c2) {
        const ex = h('input', { type: 'radio', name: 'sw', checked: true }), cu = h('input', { type: 'radio', name: 'sw' });
        return dialog({ title: '정렬 경고', width: 440, body: h('div', {}, h('p', {}, '선택 영역 옆에 데이터가 있습니다. 이 데이터는 선택하지 않았으므로 정렬되지 않습니다.'), h('label', { class: 'row' }, ex, '선택 영역 확장(E)'), h('label', { class: 'row' }, cu, '현재 선택 영역으로 정렬(C)')),
          buttons: [{ label: '정렬(S)', primary: true, onClick: () => { setTimeout(() => proceed(ex.checked ? { r1: region.r1, c1: region.c1, r2: region.r2, c2: region.c2 } : g), 0); } }, { label: '취소', cancel: true }] });
      }
      proceed(g);
    };

    // ------------------------------------------------------------------ 차트 삽입·종류 변경
    D.insertChart = (tab, focus, existing) => {
      let series;
      if (existing) series = existing.series.map(s => Object.assign({}, s));
      else {
        if (sel.ranges.length === 1 && sel.ranges[0].r1 === sel.ranges[0].r2 && sel.ranges[0].c1 === sel.ranges[0].c2) { const cr = XL.currentRegion(wb(), sh(), sel.r, sel.c); sel.ranges = [cr]; }
        series = XL.seriesFromSelection(wb(), sh(), sel.ranges);
        if (!series.length) return msg({ title: 'Microsoft Excel', text: '차트를 만들 데이터를 선택하십시오.', icon: 'warn' });
      }
      const preset = (k) => series.forEach((s, i) => {
        if (k === 'col') { s.type = 'col'; s.secondary = false; }
        else if (k === 'line') { s.type = 'line'; s.secondary = false; }
        else if (k === 'lineMarkers') { s.type = 'lineMarkers'; s.secondary = false; }
        else if (k === 'combo1') { s.type = i === 0 ? 'col' : 'line'; s.secondary = false; }
        else if (k === 'combo2') { s.type = i === 0 ? 'col' : 'line'; s.secondary = i > 0; }
        else if (k === 'combo3') { s.type = i === 0 ? 'area' : 'col'; s.secondary = false; }
      });
      let kind = existing ? (existing.series.some(s => s.type !== existing.series[0].type) || existing.series.some(s => s.secondary) ? 'combo' : existing.series[0].type) : (focus || 'col');
      if (!existing) preset(kind === 'combo' ? 'combo1' : kind);
      const prev = h('div', { style: { border: '1px solid #ddd', background: '#fff', height: '250px', overflow: 'hidden' } });
      const tmp = () => ({ series, title: { show: true, text: existing ? XL.chartTitleText(wb(), existing) : null }, axisTitle: existing ? existing.axisTitle : { x: { show: false, text: '' }, y: { show: false, text: '' } }, legend: existing ? existing.legend : { show: true, pos: 'bottom' }, gridlines: true, axes: true });
      const drawPrev = () => { prev.innerHTML = OA.chart.svg(wb(), tmp(), { w: 470, h: 250 }); };
      const right = h('div', { style: { flex: 1, minWidth: 0 } });
      const CATS = [['recent', '최근에 사용한 항목'], ['tmpl', '서식 파일'], ['col', '세로 막대형'], ['line', '꺾은선형'], ['pie', '원형'], ['bar', '가로 막대형'], ['area', '영역형'], ['scatter', '분산형'], ['map', '지도'], ['stock', '주식형'], ['surface', '표면형'], ['radar', '방사형'], ['treemap', '트리맵'], ['sunburst', '선버스트'], ['hist', '히스토그램'], ['box', '상자 수염'], ['waterfall', '폭포'], ['funnel', '깔때기형'], ['combo', '혼합']];
      let cat = kind === 'col' ? 'col' : kind === 'combo' ? 'combo' : 'line';
      const left = h('div', { class: 'lst', style: { width: '170px', height: '420px' } });
      const seriesTable = () => {
        const tbl = h('div', { style: { border: '1px solid #ddd', marginTop: '6px' } }, h('div', { style: { display: 'grid', gridTemplateColumns: '1fr 190px 60px', padding: '4px', background: '#faf9f8' } }, '계열 이름', '차트 종류', '보조 축'));
        series.forEach((s, i) => {
          const tsel = h('select', { onchange: () => { s.type = tsel.value; drawPrev(); } }, Object.entries(OA.chart.TYPE_NAMES).map(([k, l]) => h('option', { value: k, selected: s.type === k, disabled: k === 'stacked' || k === 'area' }, l)));
          const cb = h('input', { type: 'checkbox', checked: !!s.secondary, onchange: () => { s.secondary = cb.checked; drawPrev(); } });
          tbl.appendChild(h('div', { style: { display: 'grid', gridTemplateColumns: '1fr 190px 60px', padding: '3px 4px', alignItems: 'center' } }, h('span', {}, h('span', { style: { display: 'inline-block', width: '10px', height: '10px', background: OA.chart.COLORS[i % 6], marginRight: '6px' } }), XL.seriesName(wb(), s, i)), tsel, h('span', { style: { textAlign: 'center' } }, cb)));
        });
        return tbl;
      };
      const drawRight = () => {
        right.innerHTML = '';
        const icons = (list) => h('div', { class: 'row' }, list.map(([k, ic, l, dis]) => h('button', { class: 'btn', title: l, disabled: !!dis, style: { width: '62px', height: '52px', outline: kind === k ? '2px solid #217346' : '' }, onclick: () => { kind = k; preset(k); drawRight(); } }, icon(ic, 34))));
        if (cat === 'col') { right.append(icons([['col', 'colchart', '묶은 세로 막대형'], ['stacked', 'colchart', '누적 세로 막대형', true], ['stacked100', 'colchart', '100% 기준 누적 세로 막대형', true]]), h('h3', { style: { margin: '6px 0', fontWeight: 400 } }, '묶은 세로 막대형'), prev); }
        else if (cat === 'line') { right.append(icons([['line', 'linechart', '꺾은선형'], ['lstack', 'linechart', '누적 꺾은선형', true], ['lineMarkers', 'linechart', '표식이 있는 꺾은선형']]), h('h3', { style: { margin: '6px 0', fontWeight: 400 } }, OA.chart.TYPE_NAMES[kind] || '꺾은선형'), prev); }
        else if (cat === 'combo') {
          if (!['combo1', 'combo2', 'combo3', 'custom', 'combo'].includes(kind)) { kind = 'combo1'; preset('combo1'); }
          right.append(icons([['combo1', 'combochart', '묶은 세로 막대형 - 꺾은선형'], ['combo2', 'combochart', '묶은 세로 막대형 - 꺾은선형, 보조 축'], ['combo3', 'combochart', '누적 영역형 - 묶은 세로 막대형', true], ['custom', 'design', '사용자 지정 조합']]),
            h('h3', { style: { margin: '6px 0', fontWeight: 400 } }, { combo1: '묶은 세로 막대형 - 꺾은선형', combo2: '묶은 세로 막대형 - 꺾은선형, 보조 축', custom: '사용자 지정 조합', combo: '사용자 지정 조합' }[kind] || '사용자 지정 조합'), prev,
            h('div', { style: { marginTop: '6px' } }, '데이터 계열에 대한 차트 종류와 축을 선택합니다:'), seriesTable());
        } else right.append(h('div', { class: 'note', style: { padding: '40px 10px' } }, '이 시뮬레이터에서는 시험 과제에 쓰는 세로 막대형·꺾은선형·혼합 차트만 만듭니다.'));
        drawPrev();
      };
      const drawLeft = () => { left.innerHTML = ''; for (const [k, l] of CATS) { const d = h('div', { class: k === cat ? 'on' : '' }, l); d.addEventListener('click', () => { cat = k; if (k === 'col') { kind = 'col'; preset('col'); } if (k === 'line') { kind = 'line'; preset('line'); } if (k === 'combo') { kind = 'combo1'; preset('combo1'); } drawLeft(); drawRight(); }); left.appendChild(d); } };
      const allBody = h('div', { style: { display: 'flex', gap: '12px' } }, left, right);
      const recList = [['col', '묶은 세로 막대형', 'colchart'], ['combo1', '묶은 세로 막대형 - 꺾은선형', 'combochart'], ['line', '꺾은선형', 'linechart'], ['lineMarkers', '표식이 있는 꺾은선형', 'linechart']];
      const recRight = h('div', { style: { flex: 1 } });
      const recBody = h('div', { style: { display: 'flex', gap: '12px' } }, h('div', { class: 'lst', style: { width: '170px', height: '420px' } }, recList.map(([k, l, ic]) => { const d = h('div', { style: { display: 'flex', gap: '6px', alignItems: 'center', padding: '6px' } }, icon(ic, 30), l); d.addEventListener('click', () => { kind = k; preset(k); [...d.parentNode.children].forEach(x => x.classList.remove('on')); d.classList.add('on'); recRight.innerHTML = ''; recRight.append(h('h3', { style: { fontWeight: 400 } }, l), prev); drawPrev(); }); return d; })), recRight);
      const t = tabs([{ id: 'rec', label: '추천 차트', body: recBody }, { id: 'all', label: '모든 차트', body: allBody }], existing ? 'all' : tab === 'rec' ? 'rec' : 'all');
      drawLeft(); drawRight();
      if (tab === 'rec' && !existing) { recRight.append(h('h3', { style: { fontWeight: 400 } }, '묶은 세로 막대형'), prev); kind = 'col'; preset('col'); drawPrev(); }
      dialog({ title: existing ? '차트 종류 변경' : '차트 삽입', width: 720, body: t, buttons: [{ label: '확인', primary: true, onClick: () => {
        if (existing) api.exec(() => { existing.series.forEach((s, i) => { s.type = series[i].type; s.secondary = series[i].secondary; }); existing.kind = kind; });
        else {
          let created = null;
          api.exec(() => { const g0 = sel.ranges[sel.ranges.length - 1]; created = XL.newChart(wb(), sh(), series, kind, { x: XL.colLeft(sh(), g0.c2 + 2), y: XL.rowTop(sh(), g0.r1) }); });
          if (created) api.setChartSel({ id: created.id, series: null });
        }
      } }, { label: '취소', cancel: true }] });
    };
    D.selectData = (ch) => {
      if (!ch) return;
      const work = ch.series.map(s => Object.assign({}, s));
      let cur = 0;
      const lstS = h('div', { class: 'lst', style: { height: '160px', width: '260px' } });
      const lstC = h('div', { class: 'lst', style: { height: '160px', width: '220px' } });
      const draw = () => {
        lstS.innerHTML = ''; lstC.innerHTML = '';
        work.forEach((s, i) => { const d = h('div', { class: i === cur ? 'on' : '' }, h('input', { type: 'checkbox', checked: true, disabled: true }), ' ', XL.seriesName(wb(), s, i)); d.onclick = () => { cur = i; draw(); }; lstS.appendChild(d); });
        if (work[0] && work[0].cat) XL.refValues(wb(), work[0].cat).forEach(v => lstC.appendChild(h('div', {}, h('input', { type: 'checkbox', checked: true, disabled: true }), ' ', OA.formula.toStr(v))));
      };
      const editSeries = (s, isNew) => {
        const nm = refEdit({ value: s.name ? '=' + s.name : '', width: 260, pick: cb => api.startPick(v => cb('=' + v)), dialog: () => inner, label: '계열 이름' });
        const vl = refEdit({ value: s.val ? '=' + s.val : '', width: 260, pick: cb => api.startPick(v => cb('=' + v)), dialog: () => inner, label: '계열 값' });
        const inner = dialog({ title: '계열 편집', width: 380, body: h('div', { class: 'col' }, '계열 이름(N):', nm, '계열 값(V):', vl), buttons: [{ label: '확인', primary: true, onClick: () => {
          const n = nm.input.value.replace(/^=/, ''), v = vl.input.value.replace(/^=/, '');
          if (!U.parseRange(v)) { msg({ title: 'Microsoft Excel', text: '입력한 참조가 올바르지 않습니다.', icon: 'warn' }); return false; }
          s.name = U.parseRange(n) ? n : (n || null); s.val = v;
          if (isNew) { s.cat = work[0] ? work[0].cat : null; s.type = 'col'; work.push(s); }
          draw();
        } }, { label: '취소', cancel: true }] });
      };
      const editCat = () => {
        const cr = refEdit({ value: work[0] && work[0].cat ? '=' + work[0].cat : '', width: 260, pick: cb => api.startPick(v => cb('=' + v)), dialog: () => inner, label: '축 레이블 범위' });
        const inner = dialog({ title: '축 레이블', width: 380, body: h('div', { class: 'col' }, '축 레이블 범위(A):', cr), buttons: [{ label: '확인', primary: true, onClick: () => { const v = cr.input.value.replace(/^=/, ''); if (!U.parseRange(v)) return false; work.forEach(s => { s.cat = v; }); draw(); } }, { label: '취소', cancel: true }] });
      };
      draw();
      const range = h('input', { type: 'text', readOnly: true, value: '=' + [...new Set(work.flatMap(s => [s.cat, s.name, s.val]).filter(Boolean))].join(','), style: { width: '420px' } });
      dialog({ title: '데이터 원본 선택', width: 600, body: h('div', {}, h('div', { class: 'row' }, '차트 데이터 범위(D):', range),
        h('div', { style: { display: 'flex', gap: '14px', marginTop: '10px' } },
          h('div', {}, h('div', {}, '범례 항목(계열)(S)'), h('div', { class: 'row' }, h('button', { class: 'btn', onclick: () => editSeries({ name: null, val: '' }, true) }, '추가(A)'), h('button', { class: 'btn', onclick: () => work[cur] && editSeries(work[cur]) }, '편집(E)'), h('button', { class: 'btn', onclick: () => { if (work.length > 1) { work.splice(cur, 1); cur = 0; draw(); } } }, '제거(R)')), lstS),
          h('div', {}, h('div', {}, '가로(항목) 축 레이블(C)'), h('div', { class: 'row' }, h('button', { class: 'btn', onclick: editCat }, '편집(T)')), lstC))),
        buttons: [{ label: '확인', primary: true, onClick: () => api.exec(() => { ch.series = work; }) }, { label: '취소', cancel: true }] });
    };

    // ------------------------------------------------------------------ 피벗
    D.createPivot = () => {
      let g = api.act();
      if (g.r1 === g.r2 && g.c1 === g.c2) g = XL.currentRegion(wb(), sh(), sel.r, sel.c);
      let dlgRef;
      const src = refEdit({ value: U.rangeStr(g, true, sh().name), width: 280, label: '표/범위', pick: cb => api.startPick(cb), dialog: () => dlgRef });
      const rNew = h('input', { type: 'radio', name: 'pvdest', checked: true }), rOld = h('input', { type: 'radio', name: 'pvdest' });
      const loc = refEdit({ value: '', width: 280, label: '위치', pick: cb => { rOld.checked = true; return api.startPick(v => cb(v.split(':')[0])); }, dialog: () => dlgRef });
      loc.input.addEventListener('focus', () => { rOld.checked = true; });
      dlgRef = dialog({ title: '피벗 테이블 만들기', width: 470, body: h('div', {},
        h('div', {}, '분석할 데이터를 선택하십시오.'),
        h('label', { class: 'row' }, h('input', { type: 'radio', name: 'pvsrc', checked: true }), '표 또는 범위 선택(S)'),
        h('div', { class: 'row', style: { paddingLeft: '24px' } }, '표/범위(T):', src),
        h('label', { class: 'row', style: { color: '#999' } }, h('input', { type: 'radio', name: 'pvsrc', disabled: true }), '외부 데이터 원본 사용(U)'),
        h('label', { class: 'row', style: { color: '#999' } }, h('input', { type: 'radio', name: 'pvsrc', disabled: true }), '이 통합 문서의 데이터 모델 사용(D)'),
        h('div', { style: { marginTop: '10px' } }, '피벗 테이블 보고서를 넣을 위치를 선택하십시오.'),
        h('label', { class: 'row' }, rNew, '새 워크시트(N)'), h('label', { class: 'row' }, rOld, '기존 워크시트(E)'),
        h('div', { class: 'row', style: { paddingLeft: '24px' } }, '위치(L):', loc),
        h('div', { style: { marginTop: '10px' } }, '여러 테이블을 분석할 것인지 선택'),
        h('label', { class: 'row', style: { color: '#999' } }, h('input', { type: 'checkbox', disabled: true }), '데이터 모델에 이 데이터 추가(M)')),
        buttons: [{ label: '확인', primary: true, onClick: () => {
          const ref = src.input.value.replace(/^=/, '');
          const r0 = U.parseRange(ref);
          if (!r0) { msg({ title: 'Microsoft Excel', text: '참조가 유효하지 않습니다.', icon: 'warn' }); return false; }
          const fullRef = r0.sheet ? ref : U.rangeStr(r0, true, sh().name);
          let dest;
          if (rNew.checked) dest = { newSheet: true };
          else {
            const l = loc.input.value.replace(/^=/, '');
            const lr = U.parseRange(l);
            if (!lr) { msg({ title: 'Microsoft Excel', text: '대상 참조가 유효하지 않습니다.', icon: 'warn' }); return false; }
            dest = { ref: lr.sheet ? l : U.rangeStr(lr, true, sh().name) };
          }
          let res;
          api.exec(() => { res = PV.create(wb(), fullRef, dest); return res; });
          if (res && res.pv) {
            wb().active = wb().sheets.indexOf(res.sheet);
            sel.r = res.pv.r; sel.c = res.pv.c; sel.ar = sel.r; sel.ac = sel.c; sel.ranges = [{ r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c }];
            api.paneOpen('pivot');
          }
          if (res && res.error) return false;
        } }, { label: '취소', cancel: true }] });
    };
    D.valueField = (pv, idx) => {
      const v = pv.values[idx];
      if (!v) return;
      let func = v.func, nf = v.nf;
      const name = h('input', { type: 'text', value: v.name, style: { width: '300px' } });
      const fl = lst(PV.FUNC_LIST.map(k => [k, { sum: '합계', count: '개수', avg: '평균', max: '최대', min: '최소', product: '곱' }[k]]), func, k => { func = k; if (/^(합계|개수|평균|최대|최소|곱) : /.test(name.value)) name.value = PV.valueName(k, v.field, null); }, 150);
      const t = tabs([{ id: 'sum', label: '값 요약 기준', body: h('div', {}, h('div', {}, '값 필드 요약 기준(S)'), h('div', { class: 'note' }, '선택한 필드의 데이터를 요약하는 데 사용할 계산 유형을 선택합니다.'), fl) },
        { id: 'show', label: '값 표시 형식', body: h('div', {}, h('div', {}, '값 표시 형식(A)'), h('select', {}, h('option', {}, '계산 없음'))) }]);
      dialog({ title: '값 필드 설정', width: 420, body: h('div', {}, h('div', { class: 'row' }, '원본 이름: ', h('b', {}, v.field)), h('div', { class: 'row' }, '사용자 지정 이름(C):', name), t),
        buttons: [{ label: '표시 형식(N)', left: true, onClick: () => { let tmp = nf; dialog({ title: '셀 서식', width: 520, body: numberTab(nf || NF.GENERAL, x => { tmp = x; }), buttons: [{ label: '확인', primary: true, onClick: () => { nf = NF.isGeneral(tmp) ? null : tmp; } }, { label: '취소', cancel: true }] }); return false; } },
          { label: '확인', primary: true, onClick: () => { const r = api.exec(() => PV.setValueField(wb(), pv, idx, { func, nf, name: name.value })); return !(r && r.error); } }, { label: '취소', cancel: true }] });
    };
    D.fieldSettings = (pv, field) => {
      const c = PV.cacheOf(wb(), pv);
      const name = h('input', { type: 'text', value: PV.caption(c, pv, field), style: { width: '280px' } });
      dialog({ title: '필드 설정', width: 420, body: h('div', {}, h('div', { class: 'row' }, '원본 이름: ', h('b', {}, field)), h('div', { class: 'row' }, '사용자 지정 이름(M):', name),
        tabs([{ id: 'a', label: '부분합 및 필터', body: h('div', {}, h('fieldset', {}, h('legend', {}, '부분합'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'st', checked: pv.subtotals !== 'none' }), '자동(A)'), h('label', { class: 'row' }, h('input', { type: 'radio', name: 'st', checked: pv.subtotals === 'none' }), '없음(E)')), h('p', { class: 'note' }, '부분합 표시는 디자인 탭 > 부분합에서도 바꿉니다.')) },
          { id: 'b', label: '레이아웃 및 인쇄', body: h('div', { class: 'note' }, '레이아웃은 디자인 탭 > 보고서 레이아웃에서 바꿉니다.') }])),
        buttons: [{ label: '확인', primary: true, onClick: () => { const v = name.value.trim(); if (v && v !== PV.caption(c, pv, field)) api.exec(() => { pv.captions[field] = v; }); } }, { label: '취소', cancel: true }] });
    };
    D.pivotOptions = (pv) => {
      if (!pv) return;
      const name = h('input', { type: 'text', value: pv.name, style: { width: '260px' } });
      const cb = (v, l, dis) => { const i = h('input', { type: 'checkbox', checked: !!v, disabled: !!dis }); return [i, h('label', { class: 'row', style: dis ? { color: '#999' } : {} }, i, l)]; };
      const [gRow, gRowL] = cb(pv.grandRow, '행 총합계 표시(S)'), [gCol, gColL] = cb(pv.grandCol, '열 총합계 표시(G)');
      const [btn, btnL] = cb(pv.showButtons, '확장/축소 단추 표시(S)'), [hdr, hdrL] = cb(pv.showHeaders, '필드 캡션 및 필터 드롭다운 표시(D)');
      const [pb, pbL] = cb(pv.printButtons, '피벗 테이블에 표시된 경우 확장/축소 단추 인쇄(X)');
      const [rep, repL] = cb(false, '각 인쇄 페이지에 행 레이블 반복(R)');
      const [ttl, ttlL] = cb(false, '인쇄 제목 설정(I)', true);
      const t = tabs([
        { id: 'layout', label: '레이아웃 및 서식', body: h('div', { style: { minHeight: '260px' } }, h('fieldset', {}, h('legend', {}, '레이아웃'), cb(false, '레이블이 있는 셀 병합 및 가운데 맞춤(M)')[1], h('div', { class: 'row' }, '압축 형식의 행 레이블 들여쓰기(C):', h('input', { type: 'number', value: 1, style: { width: '50px' } }), '자')),
          h('fieldset', {}, h('legend', {}, '서식'), cb(false, '오류 값 표시(E)')[1], cb(true, '빈 셀 표시(S)')[1], cb(true, '업데이트 시 열 자동 맞춤(A)')[1], cb(true, '업데이트 시 셀 서식 유지(P)')[1])) },
        { id: 'totals', label: '요약 및 필터', body: h('div', { style: { minHeight: '260px' } }, h('fieldset', {}, h('legend', {}, '총합계'), gRowL, gColL)) },
        { id: 'display', label: '표시', body: h('div', { style: { minHeight: '260px' } }, h('fieldset', {}, h('legend', {}, '표시'), btnL, cb(true, '상황에 맞는 도구 설명 표시(C)')[1], hdrL)) },
        { id: 'print', label: '인쇄', body: h('div', { style: { minHeight: '260px' } }, h('fieldset', {}, h('legend', {}, '인쇄'), pbL, repL, ttlL)) },
        { id: 'data', label: '데이터', body: h('div', { style: { minHeight: '260px' } }, h('fieldset', {}, h('legend', {}, '피벗 테이블 데이터'), cb(true, '파일에 원본 데이터 저장(S)')[1], cb(false, '파일을 열 때 데이터 새로 고침(R)')[1])) },
        { id: 'alt', label: '대체 텍스트', body: h('div', { style: { minHeight: '260px' } }, '제목(T)', h('input', { type: 'text', style: { width: '100%' } }), '설명(D)', h('textarea', { style: { width: '100%', height: '80px' } })) },
      ], 'layout');
      dialog({ title: '피벗 테이블 옵션', width: 520, body: h('div', {}, h('div', { class: 'row' }, '피벗 테이블 이름(N):', name), t),
        buttons: [{ label: '확인', primary: true, onClick: () => api.exec(() => { pv.name = name.value.trim() || pv.name; pv.grandRow = gRow.checked; pv.grandCol = gCol.checked; pv.showButtons = btn.checked; pv.showHeaders = hdr.checked; pv.printButtons = pb.checked; }) }, { label: '취소', cancel: true }] });
    };
    D.groupRange = (pv, field) => {
      const c = PV.cacheOf(wb(), pv);
      const i = c.fields.indexOf(field);
      const vals = c.records.map(r => r[i]).filter(v => typeof v === 'number');
      const st = h('input', { type: 'text', value: Math.min(...vals), style: { width: '90px' } }), en = h('input', { type: 'text', value: Math.max(...vals), style: { width: '90px' } }), by = h('input', { type: 'text', value: 10, style: { width: '90px' } });
      dialog({ title: '그룹화', width: 300, body: h('div', {}, h('fieldset', {}, h('legend', {}, '자동'), h('div', { class: 'grid2' }, h('label', {}, h('input', { type: 'checkbox', checked: true }), ' 시작(S):'), st, h('label', {}, h('input', { type: 'checkbox', checked: true }), ' 끝(E):'), en)), h('div', { class: 'grid2' }, '단위(B):', by),
        h('p', { class: 'note' }, '숫자 항목 하나에서 그룹을 누르면 이 대화상자가 열립니다. 항목 여러 개를 Ctrl로 골라 그룹을 누르면 선택 항목이 그룹1로 묶입니다(11.2절).')),
        buttons: [{ label: '확인', primary: true, onClick: () => api.exec(() => PV.rangeGroup(wb(), pv, field, +st.value, +en.value, +by.value)) }, { label: '취소', cancel: true }] });
    };

    // ------------------------------------------------------------------ 페이지 설정
    D.pageSetup = (tab, after) => {
      const p = Object.assign({}, sh().page);
      const num = (v, w) => h('input', { type: 'number', step: 0.1, value: v, style: { width: (w || 70) + 'px' } });
      const rP = h('input', { type: 'radio', name: 'or', checked: p.orient === 'portrait' }), rL = h('input', { type: 'radio', name: 'or', checked: p.orient === 'landscape' });
      const rZ = h('input', { type: 'radio', name: 'sc', checked: p.scaleMode !== 'fit' }), rF = h('input', { type: 'radio', name: 'sc', checked: p.scaleMode === 'fit' });
      const zoom = num(p.zoom || 100, 60);
      const fw = h('input', { type: 'number', min: 0, value: p.fitW === undefined ? 1 : p.fitW, style: { width: '50px' } }), fh = h('input', { type: 'number', min: 0, value: p.fitH === undefined ? 1 : p.fitH, style: { width: '50px' } });
      [fw, fh].forEach(x => x.addEventListener('focus', () => { rF.checked = true; }));
      zoom.addEventListener('focus', () => { rZ.checked = true; });
      const paper = h('select', {}, ['A4', 'A3', 'B5', 'Letter'].map(x => h('option', { value: x, selected: p.paper === x }, { A4: 'A4 (210 x 297mm)', A3: 'A3 (297 x 420mm)', B5: 'B5 (182 x 257mm)', Letter: 'Letter (8.5 x 11")' }[x])));
      const mT = num(p.top), mB = num(p.bottom), mL = num(p.left), mR = num(p.right), mH = num(p.header), mF = num(p.footer);
      const cH = h('input', { type: 'checkbox', checked: !!p.centerH }), cV = h('input', { type: 'checkbox', checked: !!p.centerV });
      let dref;
      const area = refEdit({ value: sh().printArea || '', width: 280, label: '인쇄 영역', pick: cb => api.startPick(v => cb(v.replace(/^[^!]+!/, '').replace(/\$/g, ''))), dialog: () => dref });
      const prevBox = h('div', { style: { width: '120px', height: '160px', border: '1px solid #444', margin: '0 auto', position: 'relative', background: '#fff' } });
      const drawPrev = () => {
        prevBox.innerHTML = '';
        const s = 160 / 29.7;
        const t = (+mT.value || 0) * s, b2 = (+mB.value || 0) * s, l = (+mL.value || 0) * s * 0.75, r = (+mR.value || 0) * s * 0.75;
        prevBox.appendChild(h('div', { style: { position: 'absolute', left: l + 'px', right: r + 'px', top: t + 'px', bottom: b2 + 'px', border: '1px dashed #999', background: 'repeating-linear-gradient(#fff 0 6px, #ddd 6px 7px)' } }));
      };
      [mT, mB, mL, mR].forEach(x => x.addEventListener('input', drawPrev));
      drawPrev();
      const t = tabs([
        { id: 'page', label: '페이지', body: h('div', { style: { minHeight: '330px' } },
          h('fieldset', {}, h('legend', {}, '용지 방향'), h('div', { class: 'row' }, h('label', { class: 'row' }, rP, icon('paper', 28), '세로(T)'), h('label', { class: 'row' }, rL, icon('orient', 28), '가로(L)'))),
          h('fieldset', {}, h('legend', {}, '배율'), h('div', { class: 'row' }, h('label', { class: 'row' }, rZ, '확대/축소 배율(A):'), zoom, '% 정상 크기'),
            h('div', { class: 'row' }, h('label', { class: 'row' }, rF, '자동 맞춤(F):'), fw, '용지 너비', fh, '용지 높이')),
          h('div', { class: 'grid2', style: { marginTop: '10px' } }, '용지 크기(Z):', paper, '인쇄 품질(Q):', h('select', {}, h('option', {}, '600 dpi')), '시작 페이지 번호(R):', h('input', { type: 'text', value: '자동', style: { width: '80px' } }))) },
        { id: 'margins', label: '여백', body: h('div', { style: { minHeight: '330px' } },
          h('div', { style: { display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '8px', alignItems: 'center', justifyItems: 'center' } },
            h('span'), h('div', { class: 'col' }, '위쪽(T):', mT), h('div', { class: 'col' }, '머리글(A):', mH),
            h('div', { class: 'col' }, '왼쪽(L):', mL), prevBox, h('div', { class: 'col' }, '오른쪽(R):', mR),
            h('span'), h('div', { class: 'col' }, '아래쪽(B):', mB), h('div', { class: 'col' }, '바닥글(F):', mF)),
          h('fieldset', {}, h('legend', {}, '페이지 가운데 맞춤'), h('div', { class: 'row' }, h('label', { class: 'row' }, cH, '가로(Z)'), h('label', { class: 'row' }, cV, '세로(V)')))) },
        { id: 'hf', label: '머리글/바닥글', body: h('div', { style: { minHeight: '330px' } }, h('div', { class: 'grid2' }, '머리글(A):', h('select', {}, h('option', {}, '(없음)')), '바닥글(F):', h('select', {}, h('option', {}, '(없음)'))), h('p', { class: 'note' }, '출력물의 인적사항과 쪽 번호는 자필로 적습니다(9.3절). 머리글·바닥글은 쓰지 않습니다.')) },
        { id: 'sheet', label: '시트', body: h('div', { style: { minHeight: '330px' } }, h('div', { class: 'row' }, '인쇄 영역(A):', area), h('fieldset', {}, h('legend', {}, '인쇄 제목'), h('div', { class: 'row', style: { color: '#999' } }, '반복할 행(R):', h('input', { type: 'text', disabled: true })), h('div', { class: 'row', style: { color: '#999' } }, '반복할 열(C):', h('input', { type: 'text', disabled: true }))),
          h('fieldset', {}, h('legend', {}, '인쇄'), h('label', { class: 'row' }, h('input', { type: 'checkbox' }), '눈금선(G)'), h('label', { class: 'row' }, h('input', { type: 'checkbox' }), '흑백으로(B)'), h('label', { class: 'row' }, h('input', { type: 'checkbox' }), '행/열 머리글(L)'))) },
      ], tab || 'page');
      const apply = () => {
        const n = (x, d) => { const v = parseFloat(x.value); return isFinite(v) && v >= 0 ? Math.round(v * 100) / 100 : d; };
        const pa = area.input.value.replace(/^=/, '').trim();
        if (pa && !U.parseRange(pa)) { msg({ title: 'Microsoft Excel', text: '인쇄 영역 참조가 올바르지 않습니다.', icon: 'warn' }); return false; }
        const fwv = parseInt(fw.value, 10) || 0, fhv = parseInt(fh.value, 10) || 0;
        api.exec(() => {
          Object.assign(sh().page, { orient: rP.checked ? 'portrait' : 'landscape', paper: paper.value, top: n(mT, p.top), bottom: n(mB, p.bottom), left: n(mL, p.left), right: n(mR, p.right), header: n(mH, p.header), footer: n(mF, p.footer),
            centerH: cH.checked, centerV: cV.checked, scaleMode: rF.checked ? 'fit' : 'zoom', zoom: Math.max(10, Math.min(400, parseInt(zoom.value, 10) || 100)), fitW: fwv, fitH: fhv });
          sh().printArea = pa ? pa.replace(/\$/g, '').replace(/^[^!]+!/, '') : null;
        });
        after && after();
      };
      dref = dialog({ title: '페이지 설정', width: 520, body: t, buttons: [
        { label: '인쇄(P)...', left: true, onClick: () => { if (apply() === false) return false; api.openBackstage('print'); } },
        { label: '인쇄 미리 보기(W)', left: true, onClick: () => { if (apply() === false) return false; api.openBackstage('print'); } },
        { label: '옵션(O)...', left: true, disabled: true },
        { label: '확인', primary: true, onClick: apply }, { label: '취소', cancel: true }] });
    };
    D.colWidth = () => {
      const c = sel.c;
      const cur = XL.colW(sh(), c);
      const inp = h('input', { type: 'text', value: ((cur - 5) / 8).toFixed(2), style: { width: '80px' } });
      dialog({ title: '열 너비', width: 260, body: h('div', { class: 'row' }, '열 너비(C):', inp), buttons: [{ label: '확인', primary: true, onClick: () => { const v = parseFloat(inp.value); if (!(v >= 0 && v <= 255)) return false; api.exec(() => { for (const g of sel.ranges) for (let k = g.c1; k <= g.c2; k++) sh().colW[k] = Math.round(v * 8 + 5); }); } }, { label: '취소', cancel: true }] });
    };
    D.rowHeight = () => {
      const cur = XL.rowH(sh(), sel.r);
      const inp = h('input', { type: 'text', value: (cur * 0.75).toFixed(2), style: { width: '80px' } });
      dialog({ title: '행 높이', width: 260, body: h('div', { class: 'row' }, '행 높이(R):', inp), buttons: [{ label: '확인', primary: true, onClick: () => { const v = parseFloat(inp.value); if (!(v >= 0 && v <= 409)) return false; api.exec(() => { for (const g of sel.ranges) for (let k = g.r1; k <= g.r2; k++) sh().rowH[k] = Math.round(v / 0.75); }); } }, { label: '취소', cancel: true }] });
    };
    return D;
  };

  // ------------------------------------------------------------------ 피벗 테이블 필드 창
  OA.xlPivotPane = function (app, api, pv, onClose) {
    const wb = api.wb();
    const c = PV.cacheOf(wb, pv);
    const fields = PV.allFields(c);
    const search = h('input', { type: 'text', placeholder: '검색', style: { width: '100%', height: '24px', border: '1px solid #c8c6c4', padding: '0 6px' } });
    const list = h('div', { class: 'pvlist', style: { border: '1px solid #c8c6c4', height: '170px', overflow: 'auto', margin: '6px 0' } });
    const drawList = () => {
      list.innerHTML = '';
      for (const f of fields) {
        if (search.value && !PV.caption(c, pv, f).includes(search.value) && !f.includes(search.value)) continue;
        const on = !!PV.placed(pv, f);
        const cb = h('input', { type: 'checkbox', checked: on, onchange: () => { api.exec(() => { if (cb.checked) PV.addField(wb, pv, f); else PV.removeField(wb, pv, f); }); } });
        const row = h('div', { class: 'f' + (on ? ' b' : '') }, cb, h('span', {}, PV.caption(c, pv, f)));
        row.addEventListener('mousedown', e => { if (e.target === cb) return; startDrag(e, { field: f, from: null }); });
        list.appendChild(row);
      }
    };
    search.addEventListener('input', drawList);
    search.addEventListener('keydown', e => e.stopPropagation());
    const areaDefs = [['filters', '필터', 'filter'], ['cols', '열', 'colchart'], ['rows', '행', 'alignL'], ['values', '값', 'sum']];
    const boxes = {};
    const areas = h('div', { class: 'pvareas' });
    for (const [k, l, ic] of areaDefs) {
      const box = h('div', { class: 'box' });
      boxes[k] = box;
      const items = k === 'values' ? pv.values.map((v, i) => ({ label: v.name, idx: i, field: v.field })) : pv[k].map((f, i) => ({ label: f === 'Σ값' ? 'Σ 값' : PV.caption(c, pv, f), idx: i, field: f }));
      for (const it of items) {
        const dd = h('button', { title: '메뉴' }, '▼');
        const el = h('div', { class: 'it' }, h('span', {}, it.label), dd);
        el.addEventListener('mousedown', e => { if (e.target === dd) return; startDrag(e, { field: it.field, from: k, idx: it.idx }); });
        dd.addEventListener('click', () => menu(dd, [
          { label: '위로 이동(U)', disabled: it.idx === 0, onClick: () => api.exec(() => PV.moveInArea(wb, pv, k, it.idx, it.idx - 1)) },
          { label: '아래로 이동(D)', disabled: it.idx === items.length - 1, onClick: () => api.exec(() => PV.moveInArea(wb, pv, k, it.idx, it.idx + 1)) },
          { label: '처음으로 이동(G)', onClick: () => api.exec(() => PV.moveInArea(wb, pv, k, it.idx, 0)) }, { label: '끝으로 이동(E)', onClick: () => api.exec(() => PV.moveInArea(wb, pv, k, it.idx, 999)) }, { sep: true },
          ...(it.field === 'Σ값' ? [] : [['filters', '보고서 필터로 이동'], ['rows', '행 레이블로 이동'], ['cols', '열 레이블로 이동'], ['values', '값으로 이동']].filter(([a]) => a !== k).map(([a, l2]) => ({ label: l2, onClick: () => api.exec(() => { if (k === 'values') PV.removeValue(wb, pv, it.idx); else PV.removeField(wb, pv, it.field, k); PV.addField(wb, pv, it.field, a); }) }))),
          { sep: true }, { label: '필드 제거(V)', disabled: it.field === 'Σ값', onClick: () => api.exec(() => { if (k === 'values') PV.removeValue(wb, pv, it.idx); else PV.removeField(wb, pv, it.field, k); }) },
          k === 'values' ? { label: '값 필드 설정(N)...', onClick: () => OA.xlDialogs(app, api).valueField(pv, it.idx) } : { label: '필드 설정(N)...', disabled: it.field === 'Σ값', onClick: () => OA.xlDialogs(app, api).fieldSettings(pv, it.field) }]));
        box.appendChild(el);
      }
      areas.appendChild(h('div', { class: 'pvarea' }, h('div', { class: 'h' }, icon(ic), l), box));
    }
    function startDrag(e, info) {
      e.preventDefault();
      let ghost = null;
      const sx = e.clientX, sy = e.clientY;
      const over = ev => { for (const [k, b] of Object.entries(boxes)) { const r = b.getBoundingClientRect(); if (ev.clientX >= r.left && ev.clientX <= r.right && ev.clientY >= r.top && ev.clientY <= r.bottom) return k; } return null; };
      OA.ui.drag(e, ev => {
        if (!ghost && Math.abs(ev.clientX - sx) + Math.abs(ev.clientY - sy) < 4) return;
        if (!ghost) { ghost = h('div', { class: 'dragghost' }, PV.caption(c, pv, info.field) || info.field); document.body.appendChild(ghost); }
        ghost.style.left = ev.clientX + 8 + 'px'; ghost.style.top = ev.clientY + 8 + 'px';
        Object.entries(boxes).forEach(([k, b]) => b.classList.toggle('drop', k === over(ev)));
      }, ev => {
        if (!ghost) return;
        ghost.remove();
        Object.values(boxes).forEach(b => b.classList.remove('drop'));
        const to = over(ev);
        const paneR = el.getBoundingClientRect();
        const outside = ev.clientX < paneR.left;
        if (!to) { if (info.from && outside) api.exec(() => { if (info.from === 'values') PV.removeValue(wb, pv, info.idx); else PV.removeField(wb, pv, info.field, info.from); }); return; }
        if (info.field === 'Σ값' && !(to === 'rows' || to === 'cols')) return;
        const b = boxes[to];
        const kids = [...b.children];
        let pos = kids.findIndex(k2 => { const r = k2.getBoundingClientRect(); return ev.clientY < r.top + r.height / 2; });
        if (pos < 0) pos = kids.length;
        api.exec(() => {
          if (info.from === to) { PV.moveInArea(wb, pv, to, info.idx, pos > info.idx ? pos - 1 : pos); return; }
          if (info.from === 'values') PV.removeValue(wb, pv, info.idx); else if (info.from) PV.removeField(wb, pv, info.field, info.from);
          if (info.field === 'Σ값') { pv[to].splice(pos, 0, 'Σ값'); return; }
          PV.addField(wb, pv, info.field, to, pos);
        });
      });
    }
    drawList();
    const el = h('div', { class: 'pane' },
      h('div', { class: 'pane-h' }, h('span', { class: 't' }, '피벗 테이블 필드'), h('button', { title: '작업창 옵션' }, '⌄'), h('button', { title: '닫기', onclick: onClose }, '✕')),
      h('div', { class: 'pane-b' },
        h('div', { class: 'row', style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' } }, h('span', {}, '보고서에 추가할 필드 선택:'), h('button', { class: 'btn', title: '도구', disabled: true }, '⚙')),
        search, list, h('div', { style: { color: '#444', margin: '6px 0' } }, '아래 영역 사이에 필드를 끌어 놓으십시오.'), areas,
        h('div', { class: 'row', style: { display: 'flex', justifyContent: 'space-between', marginTop: '10px' } }, h('label', {}, h('input', { type: 'checkbox', disabled: true }), ' 나중에 레이아웃 업데이트'), h('button', { class: 'btn', disabled: true }, '업데이트'))));
    return el;
  };
})(typeof window !== 'undefined' ? window : globalThis);

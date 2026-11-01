'use strict';
/* 공유용: 부서별 집계만 다룬다. 이 파일의 데이터에는 개인별 행이 없다. */
const D = readData();
const MIN = D.meta.minGroup;
const f = { dept: '' };

$('#chip-base').textContent = '기준일 ' + D.meta.baseDate;
renderNotices($('#notices'), D.meta.notices);

/* 부서 선택 목록 (가려진 조직은 목록에도 넣지 않는다) */
(function buildDeptOptions() {
  const sel = $('#f-dept');
  sel.innerHTML = '<option value="">전체 부서</option>';
  let curDiv = null, grp = null;
  D.orgs.forEach(r => {
    if (!r.org) {
      if (r.suppressed) { curDiv = null; return; }
      grp = document.createElement('optgroup'); grp.label = r.div;
      const o = document.createElement('option'); o.value = 'div:' + r.div; o.textContent = r.div + ' 전체'; grp.appendChild(o);
      sel.appendChild(grp); curDiv = r.div; return;
    }
    if (r.suppressed || !grp || curDiv !== r.div) return;
    const o = document.createElement('option'); o.value = 'org:' + r.div + '|' + r.org; o.textContent = r.org; grp.appendChild(o);
  });
  sel.addEventListener('change', () => { f.dept = sel.value; render(); });
})();

function deptRows() {
  if (!f.dept) return D.orgs;
  const [t, v] = [f.dept.slice(0, 3), f.dept.slice(4)];
  if (t === 'div') return D.orgs.filter(r => r.div === v);
  const [dv, og] = v.split('|');
  return D.orgs.filter(r => r.div === dv && (!r.org || r.org === og));
}

function toVals(r) {
  const v = {};
  if (r.suppressed) return v;
  const scored = r.counts[0] + r.counts[1] + r.counts[2];
  v.score = r.avgScore;
  v.hiRate = r.n ? r.counts[0] / r.n * 100 : null;
  D.signals.forEach((_, i) => { v['s' + i] = r.sig ? r.sig[i] : null; });
  if (r.src) { v.ot3 = r.src.ot3; v.rn = r.src.rn; v.raise = r.src.raise; }
  return v;
}

function columns() {
  const cols = [{ key: 'score', label: '평균 점수', dec: 1, title: '부서 구성원의 점수 평균 (확인 순서 지표, 퇴사 확률 아님)' }, { key: 'hiRate', label: '상 비율(%)', dec: 0 }];
  D.signals.forEach((s, i) => cols.push({ key: 's' + i, label: s, dec: 2, title: s + ' 신호 점수 평균 (0~1)' }));
  if (D.meta.hasSource) {
    cols.push({ key: 'ot3', label: '야근(h/월)', dec: 1, title: '최근 3개월 1인당 평균 초과근무시간' });
    cols.push({ key: 'rn', label: '평가(1–5)', dec: 2, title: '최근 확정 평가의 평균 (S=5 … D=1)' });
    cols.push({ key: 'raise', label: '인상률(%)', dec: 1, title: '최근 급여 변동률 평균' });
  }
  return cols;
}

function render() {
  const rows = deptRows();
  const selected = f.dept ? rows.filter(r => f.dept.startsWith('org:') ? !!r.org : !r.org)[0] : null;

  /* 1 요약 */
  let filtered = null;
  if (selected && !selected.suppressed) {
    filtered = { total: selected.n, counts: { '상': selected.counts[0], '중': selected.counts[1], '하': selected.counts[2], '판단불가': selected.counts[3] } };
  }
  renderTiles($('#tiles'), { baseDate: D.meta.baseDate, total: D.summary.total, counts: D.summary.counts, filtered });
  $('#lines').innerHTML = (D.summary.lines || []).map(l => `<li>${esc(l)}</li>`).join('');
  $('#f-count').textContent = f.dept ? (selected && selected.suppressed ? '선택한 부서는 표본 부족으로 값을 표시하지 않습니다' : '부서 필터 적용 중') : '';

  /* 3 조직 현황 */
  const hrows = rows.map(r => ({ label: r.org || (r.div + ' 전체'), kind: r.org ? 'org' : 'div', suppressed: r.suppressed, protect: r.protect, n: r.n, vals: toVals(r), counts: r.counts }));
  const divRows = D.orgs.filter(r => !r.org);
  if (!f.dept && D.meta.hasSource && D.company && divRows.length && !divRows.some(r => r.suppressed)) {
    const w = divRows.reduce((a, r) => a + (r.counts[0] + r.counts[1] + r.counts[2]), 0);
    const score = w ? divRows.reduce((a, r) => a + r.avgScore * (r.counts[0] + r.counts[1] + r.counts[2]), 0) / w : null;
    hrows.push({ label: '회사 전체', kind: 'div', n: D.summary.total, vals: { score, hiRate: D.summary.counts['상'] / D.summary.total * 100, ot3: D.company.ot3, rn: D.company.rating, raise: D.company.raise } });
  }
  const heatEl = $('#heat');
  renderHeatmap(heatEl, { cols: columns(), rows: hrows, minGroup: MIN });
  if (!D.meta.hasSource && D.reasons.orgMetrics) heatEl.insertAdjacentHTML('beforeend', `<div class="reason" style="margin-top:8px">야근·평가·인상률 열: ${esc(D.reasons.orgMetrics)}</div>`);
  renderDist($('#dist'), { rows: hrows.filter(r => r.counts || r.suppressed), minGroup: MIN });

  /* 4 월간 변화 */
  const mEl = $('#monthly');
  if (!D.monthly) { mEl.innerHTML = reasonBox('월간 변화: ' + D.reasons.monthly); }
  else {
    const S = D.monthly.series, keys = [];
    const dv = f.dept.startsWith('div:') ? f.dept.slice(4) : f.dept.startsWith('org:') ? f.dept.slice(4).split('|')[0] : null;
    const og = f.dept.startsWith('org:') ? f.dept.slice(4) : null;
    keys.push(['company', '회사 전체', true]);
    if (!f.dept) Object.keys(S).filter(k => k.startsWith('div:')).forEach(k => keys.push([k, k.slice(4), false]));
    else {
      keys.push(['div:' + dv, dv, false]);
      if (!og) Object.keys(S).filter(k => k.startsWith('org:' + dv + '|')).forEach(k => keys.push([k, k.split('|')[1], false]));
      else keys.push(['org:' + og, og.split('|')[1], false]);
    }
    const series = keys.filter(k => S[k[0]]).map((k, i) => ({ key: k[0], label: k[1], values: S[k[0]].values, color: k[2] ? '#c9ced6' : SERIES_COLORS[i % SERIES_COLORS.length], dash: k[2] }));
    mEl.innerHTML = '<div id="mchart"></div><div class="small muted" style="margin-top:6px">월 인원이 5명 미만인 달과 5명 미만 조직은 표시하지 않습니다.</div>';
    renderLineChart($('#mchart'), { months: D.monthly.months, series, unit: D.monthly.unit });
  }
}

renderDetail($('#detail'), D.detail);
render();

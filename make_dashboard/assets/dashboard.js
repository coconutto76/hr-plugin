'use strict';
/* 내부용 대시보드. 직원 개인의 점수·등급은 결과 파일 값 그대로 쓰고, 어떤 필터에서도 다시 계산하지 않는다. */
const D = readData();
const MIN = D.meta.minGroup;
const STATUSES = ['미확인', '면담 예정', '면담 완료', '조치 불필요'];
const KEY = 'hr-dashboard-notes:' + D.meta.baseDate + ':' + D.meta.resultFile;
const f = { dept: '', q: '', grade: '', status: '' };
let sort = { key: 'score', dir: -1 };
const open = new Set();

$('#chip-base').textContent = '기준일 ' + D.meta.baseDate;
$('#chip-src').textContent = D.meta.hasSource ? (D.meta.linked ? '원본 연결됨' : '원본 있음 · 직원 연결 안 함') : '원본 없음';
renderNotices($('#notices'), D.meta.notices);

/* ---------- 확인 기록 저장 (분석용 ID -> 담당자·상태·다음 확인일) ---------- */
let notes = {}, storageOK = true;
try { const s = localStorage.getItem(KEY); notes = s ? JSON.parse(s) : {}; } catch (e) { storageOK = false; notes = {}; }
function saveNotes() {
  try { localStorage.setItem(KEY, JSON.stringify(notes)); }
  catch (e) { storageOK = false; warnStorage(); }
}
function warnStorage() {
  $('#storage-warn').innerHTML = storageOK ? '' : '<div class="banner warn">이 브라우저에서는 확인 기록을 저장할 수 없습니다(시크릿 창이거나 저장이 막혀 있을 수 있습니다). 입력한 내용은 이 창을 닫으면 사라지니, 작업 중간에 JSON으로 내보내 두세요.</div>';
}
warnStorage();
const noteOf = id => notes[id] || { owner: '', status: '미확인', next: '' };
function setNote(id, patch) {
  const n = Object.assign({}, noteOf(id), patch);
  if (!n.owner && n.status === '미확인' && !n.next) delete notes[id]; else notes[id] = n;
  saveNotes();
}

/* ---------- 필터 ---------- */
const P = D.people;
function buildControls() {
  const dept = $('#f-dept');
  dept.innerHTML = '<option value="">전체 부서</option>';
  const divs = [...new Set(P.map(p => p.div))].sort();
  divs.forEach(dv => {
    const g = document.createElement('optgroup'); g.label = dv;
    g.innerHTML = `<option value="div:${esc(dv)}">${esc(dv)} 전체</option>` +
      [...new Set(P.filter(p => p.div === dv).map(p => p.org))].sort().map(o => `<option value="org:${esc(dv)}|${esc(o)}">${esc(o)}</option>`).join('');
    dept.appendChild(g);
  });
  $('#f-grade').innerHTML = '<option value="">전체 등급</option>' + GRADES.map(g => `<option>${g}</option>`).join('');
  $('#f-status').innerHTML = '<option value="">전체 상태</option>' + STATUSES.map(s => `<option>${s}</option>`).join('');
  dept.addEventListener('change', () => { f.dept = dept.value; render(); });
  $('#f-grade').addEventListener('change', e => { f.grade = e.target.value; render(); });
  $('#f-status').addEventListener('change', e => { f.status = e.target.value; render(); });
  let t = null;
  $('#f-q').addEventListener('input', e => { clearTimeout(t); t = setTimeout(() => { f.q = e.target.value.trim().toLowerCase(); render(); }, 120); });
  $('#f-reset').addEventListener('click', () => {
    Object.assign(f, { dept: '', q: '', grade: '', status: '' });
    $('#f-dept').value = ''; $('#f-q').value = ''; $('#f-grade').value = ''; $('#f-status').value = ''; render();
  });
}
function inDept(p) {
  if (!f.dept) return true;
  const v = f.dept.slice(4);
  return f.dept.startsWith('div:') ? p.div === v : (p.div + '|' + p.org) === v;
}
const pass = p => inDept(p) && (!f.grade || p.grade === f.grade) && (!f.q || p.id.toLowerCase().includes(f.q)) && (!f.status || noteOf(p.id).status === f.status);
const countGrades = list => { const c = {}; GRADES.forEach(g => { c[g] = list.filter(p => p.grade === g).length; }); return c; };

/* ---------- 지표(원본 + ID 연결이 있을 때) ---------- */
const metricsOn = !!(D.meta.linked && D.company);
const RN_LABEL = ['', 'D', 'C', 'B', 'A', 'S'];
function metricCell(kind, p) {
  if (!metricsOn) return '<span class="muted">—</span>';
  const m = p.m || [null, null, null, null], c = D.company;
  if (kind === 'ot') {
    const v = m[0]; if (v == null) return '<span class="muted">판단불가</span>';
    const d = c.ot3 == null ? null : v - c.ot3;
    return `<span class="metric"><span class="val">${fmt(v, 1)}h</span><span class="cmp ${cls(d)}">전사 ${fmt(c.ot3, 1)}h · ${sign(d, 1)}h</span></span>`;
  }
  if (kind === 'rating') {
    const v = m[1], pv = m[2]; if (v == null) return '<span class="muted">판단불가</span>';
    const d = c.rating == null ? null : v - c.rating;
    const chg = pv == null ? '직전 평가 없음' : (v === pv ? '직전과 같음' : `${RN_LABEL[pv]}→${RN_LABEL[v]}`);
    return `<span class="metric"><span class="val">${RN_LABEL[v]} (${v})</span><span class="cmp ${cls(-d)}">${chg} · 전사 ${fmt(c.rating, 2)}</span></span>`;
  }
  const v = m[3]; if (v == null) return '<span class="muted">인상 이력 없음</span>';
  const d = c.raise == null ? null : v - c.raise;
  return `<span class="metric"><span class="val">${fmt(v, 1)}%</span><span class="cmp ${cls(-d)}">전사 ${fmt(c.raise, 1)}% · ${sign(d, 1)}%p</span></span>`;
}
const sign = (d, n) => d == null ? '—' : (d > 0 ? '+' : '') + d.toFixed(n);
/* 위험 신호 쪽으로 벗어난 값(야근은 위, 평가·인상률은 아래)을 노란색, 반대를 초록색으로 표시한다. */
const cls = d => d == null ? 'flat' : d > 0.0001 ? 'up' : d < -0.0001 ? 'down' : 'flat';

/* ---------- 1 요약 ---------- */
function renderSummary(list) {
  const active = !!(f.dept || f.q || f.grade || f.status);
  renderTiles($('#tiles'), { baseDate: D.meta.baseDate, total: D.summary.total, counts: D.summary.counts, filtered: active ? { total: list.length, counts: countGrades(list) } : null });
  $('#lines').innerHTML = (D.summary.lines || []).map(l => `<li>${esc(l)}</li>`).join('');
  $('#f-count').textContent = active ? `필터 결과 ${list.length}명 / 전체 ${P.length}명` : `전체 ${P.length}명`;
}

/* ---------- 2 직원별 현황 ---------- */
const COLS = [
  { k: 'id', t: '분석용 ID', s: p => p.id, n: false },
  { k: 'grade', t: '등급', s: p => GRADES.indexOf(p.grade), n: true },
  { k: 'score', t: '점수', s: p => p.score, n: true },
  { k: 'org', t: '소속', s: p => p.div + p.org, n: false },
  { k: 'ot', t: '야근 (최근 3개월 평균)', s: p => p.m && p.m[0], n: true, off: true },
  { k: 'rating', t: '평가', s: p => p.m && p.m[1], n: true, off: true },
  { k: 'raise', t: '인상률', s: p => p.m && p.m[3], n: true, off: true },
  { k: 'owner', t: '담당자', s: p => noteOf(p.id).owner, n: false },
  { k: 'status', t: '확인 상태', s: p => STATUSES.indexOf(noteOf(p.id).status), n: true },
  { k: 'next', t: '다음 확인일', s: p => noteOf(p.id).next, n: false },
];
function renderPeople(list) {
  const el = $('#people');
  let msg = '';
  if (!metricsOn) msg = `<div class="reason" style="margin-bottom:10px"><b>야근·평가·인상률 열:</b> ${esc(D.reasons.metrics || '원본 정보가 없어 표시할 수 없습니다.')}</div>`;
  const col = COLS.find(c => c.k === sort.key) || COLS[2];
  const rows = list.slice().sort((a, b) => {
    const x = col.s(a), y = col.s(b);
    if (x == null && y == null) return a.id.localeCompare(b.id);
    if (x == null) return 1; if (y == null) return -1;
    const r = col.n ? (x - y) : String(x).localeCompare(String(y), 'ko');
    return (r || a.id.localeCompare(b.id)) * sort.dir;
  });
  const head = COLS.map(c => `<th class="sortable ${c.n ? 'num' : ''}" data-k="${c.k}" aria-sort="${sort.key === c.k ? (sort.dir > 0 ? 'ascending' : 'descending') : 'none'}">${esc(c.t)}${sort.key === c.k ? (sort.dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('');
  const body = rows.map(p => {
    const n = noteOf(p.id), isOpen = open.has(p.id);
    const sel = `<select data-id="${p.id}" data-f="status" aria-label="${p.id} 확인 상태">${STATUSES.map(s => `<option${s === n.status ? ' selected' : ''}>${s}</option>`).join('')}</select>`;
    const main = `<tr class="row-click" data-row="${p.id}" tabindex="0" aria-expanded="${isOpen}">
      <td><b>${esc(p.id)}</b></td><td>${gradeBadge(p.grade)}</td><td class="num">${p.score == null ? '—' : p.score.toFixed(1)}</td>
      <td>${esc(p.div)} · ${esc(p.org)}</td><td>${metricCell('ot', p)}</td><td>${metricCell('rating', p)}</td><td>${metricCell('raise', p)}</td>
      <td><input class="owner" data-id="${p.id}" data-f="owner" value="${esc(n.owner)}" maxlength="40" placeholder="담당자" aria-label="${p.id} 담당자"></td>
      <td>${sel}</td>
      <td><input class="date" type="date" data-id="${p.id}" data-f="next" value="${esc(n.next)}" aria-label="${p.id} 다음 확인일"></td></tr>`;
    return main + (isOpen ? detailRow(p) : '');
  }).join('');
  el.innerHTML = msg + (rows.length ? `<div class="tablewrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>
    <div class="small muted" style="margin-top:6px">행을 누르면 판단 근거를 볼 수 있습니다. 담당자·확인 상태·다음 확인일은 이 브라우저에 저장됩니다.</div>` : reasonBox('필터에 해당하는 직원이 없습니다.'));
}
function detailRow(p) {
  const bars = D.signals.map((s, i) => {
    const v = p.sig[i];
    return `<div class="sigbar"><span>${esc(s)}</span><span class="bar"><i style="width:${v == null ? 0 : Math.round(v * 100)}%"></i></span><span class="num">${v == null ? '판단불가' : v.toFixed(2)}</span></div>`;
  }).join('');
  return `<tr class="detail"><td colspan="${COLS.length}"><div class="detail-grid">
    <div><h4>판단 근거</h4>${esc(p.reason || '—')}</div>
    <div><h4>추가 확인 사항 (확인 제안이며 조치 결정이 아님)</h4>${esc(p.action || '—')}</div>
    <div><h4>신호별 점수 (0~1)</h4>${bars}<div class="small muted">계산에 쓴 가중치 합 ${esc(p.availW)}${p.na ? ' · 판단불가: ' + esc(p.na) : ''}</div></div>
    ${p.ref ? `<div><h4>참고</h4>${esc(p.ref)}</div>` : ''}</div></td></tr>`;
}

/* ---------- 3 조직 현황 ---------- */
function groups() {
  const out = [];
  [...new Set(P.map(p => p.div))].sort().forEach(dv => {
    out.push({ kind: 'div', label: dv + ' 전체', div: dv, org: null, all: P.filter(p => p.div === dv) });
    [...new Set(P.filter(p => p.div === dv).map(p => p.org))].sort().forEach(o => out.push({ kind: 'org', label: o, div: dv, org: o, all: P.filter(p => p.div === dv && p.org === o) }));
  });
  return out;
}
const srcFor = g => (D.srcGroups || []).find(s => s.div === g.div && s.org === g.org);
function heatCols() {
  const c = [{ key: 'score', label: '평균 점수', dec: 1, title: '필터에 걸린 구성원의 점수 평균 (확인 순서 지표)' }, { key: 'hiRate', label: '상 비율(%)', dec: 0 }];
  D.signals.forEach((s, i) => c.push({ key: 's' + i, label: s, dec: 2, title: s + ' 신호 점수 평균 (0~1)' }));
  if (D.meta.hasSource) {
    c.push({ key: 'ot3', label: '야근(h/월)', dec: 1, title: '최근 3개월 1인당 평균 초과근무시간' });
    c.push({ key: 'rn', label: '평가(1–5)', dec: 2, title: '최근 확정 평가 평균 (S=5 … D=1)' });
    c.push({ key: 'raise', label: '인상률(%)', dec: 1, title: '최근 급여 변동률 평균' });
  }
  return c;
}
const avg = a => { const v = a.filter(x => x != null); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : null; };
function renderOrgs() {
  const rows = [];
  groups().forEach(g => {
    if (f.dept) {
      const v = f.dept.slice(4);
      if (f.dept.startsWith('div:')) { if (g.div !== v) return; }
      else if (g.div !== v.split('|')[0] || (g.org && (g.div + '|' + g.org) !== v)) return;
    }
    const full = g.all.length, mem = g.all.filter(pass);
    const base = { label: g.label, kind: g.kind, n: mem.length, counts: GRADES.map(x => mem.filter(p => p.grade === x).length), vals: {} };
    if (full < MIN) { base.suppressed = true; rows.push(base); return; }
    if (!mem.length) { base.missing = true; rows.push(base); return; }
    const v = base.vals;
    v.score = avg(mem.map(p => p.score));
    v.hiRate = mem.filter(p => p.grade === '상').length / mem.length * 100;
    D.signals.forEach((_, i) => { v['s' + i] = avg(mem.map(p => p.sig[i])); });
    if (D.meta.hasSource) {
      if (metricsOn) { v.ot3 = avg(mem.map(p => p.m && p.m[0])); v.rn = avg(mem.map(p => p.m && p.m[1])); v.raise = avg(mem.map(p => p.m && p.m[3])); }
      else { const s = srcFor(g); if (s && s.n >= MIN) { v.ot3 = s.ot3; v.rn = s.rn; v.raise = s.raise; } }
    }
    rows.push(base);
  });
  const active = !!(f.dept || f.q || f.grade || f.status);
  const heat = $('#heat');
  renderHeatmap(heat, { cols: heatCols(), rows, minGroup: MIN });
  const notes_ = [];
  if (active) notes_.push('필터에 걸린 직원만으로 다시 센 값입니다. 개인의 점수·등급은 바뀌지 않습니다.');
  if (!D.meta.hasSource) notes_.push('야근·평가·인상률 열: ' + (D.reasons.orgMetrics || '원본이 없어 표시할 수 없습니다.'));
  else if (!metricsOn) notes_.push('야근·평가·인상률 열은 직원 연결 없이 부서 전체로 계산한 값이라 필터의 영향을 받지 않습니다.');
  if (notes_.length) heat.insertAdjacentHTML('beforeend', `<div class="small muted" style="margin-top:6px">${notes_.map(esc).join('<br>')}</div>`);
  renderDist($('#dist'), { rows, minGroup: MIN });
}

/* ---------- 4 월간 변화 ---------- */
function renderMonthly() {
  const el = $('#monthly');
  if (!D.monthly) { el.innerHTML = reasonBox('월간 변화: ' + (D.reasons.monthly || '원본이 없어 표시할 수 없습니다.')); return; }
  const S = D.monthly.series;
  const dv = f.dept.startsWith('div:') ? f.dept.slice(4) : f.dept.startsWith('org:') ? f.dept.slice(4).split('|')[0] : null;
  const og = f.dept.startsWith('org:') ? f.dept.slice(4) : null;
  const keys = [['company', '회사 전체', true]];
  if (!f.dept) Object.keys(S).filter(k => k.startsWith('div:')).forEach(k => keys.push([k, k.slice(4), false]));
  else {
    keys.push(['div:' + dv, dv, false]);
    if (!og) Object.keys(S).filter(k => k.startsWith('org:' + dv + '|')).forEach(k => keys.push([k, k.split('|')[1], false]));
    else keys.push(['org:' + og, og.split('|')[1], false]);
  }
  const series = keys.filter(k => S[k[0]]).map((k, i) => ({ key: k[0], label: k[1], values: S[k[0]].values, color: k[2] ? '#c9ced6' : SERIES_COLORS[i % SERIES_COLORS.length], dash: k[2] }));
  const other = (f.q || f.grade || f.status) ? '<div class="banner warn" style="margin-bottom:10px">월간 변화에는 부서 필터만 적용됩니다. 검색·등급·확인 상태 필터는 직원 연결이 필요하고, 작은 집단에서는 개인이 드러날 수 있어 적용하지 않습니다.</div>' : '';
  el.innerHTML = other + '<div id="mchart"></div><div class="small muted" style="margin-top:6px">기준일이 속한 달까지의 월별 1인당 평균 초과근무시간입니다. 월 인원이 5명 미만인 달은 비워 두었습니다.</div>';
  renderLineChart($('#mchart'), { months: D.monthly.months, series, unit: D.monthly.unit });
}

/* ---------- 이벤트 ---------- */
let listCache = [];
function render() {
  listCache = P.filter(pass);
  renderSummary(listCache); renderPeople(listCache); renderOrgs(); renderMonthly();
}
function refreshCounts() { renderSummary(listCache = P.filter(pass)); renderOrgs(); }

document.addEventListener('click', e => {
  const th = e.target.closest('th.sortable');
  if (th) { const k = th.dataset.k; sort = sort.key === k ? { key: k, dir: -sort.dir } : { key: k, dir: (COLS.find(c => c.k === k).n && k !== 'next') ? -1 : 1 }; renderPeople(listCache); return; }
  if (e.target.closest('input,select,button')) return;
  const tr = e.target.closest('tr[data-row]');
  if (tr) { const id = tr.dataset.row; open.has(id) ? open.delete(id) : open.add(id); renderPeople(listCache); }
});
document.addEventListener('keydown', e => {
  if (e.key !== 'Enter' && e.key !== ' ') return;
  const tr = e.target.closest && e.target.closest('tr[data-row]');
  if (tr && e.target === tr) { e.preventDefault(); const id = tr.dataset.row; open.has(id) ? open.delete(id) : open.add(id); renderPeople(listCache); }
});
document.addEventListener('change', e => {
  const t = e.target; if (!t.dataset || !t.dataset.f) return;
  const patch = {}; patch[t.dataset.f] = t.value;
  if (t.dataset.f === 'owner') patch.owner = t.value.trim().slice(0, 40);
  setNote(t.dataset.id, patch);
  if (t.dataset.f === 'status' && f.status) render(); else refreshCounts();
});

/* ---------- JSON 내보내기 / 불러오기 ---------- */
$('#btn-export').addEventListener('click', () => {
  const entries = {}; Object.keys(notes).forEach(id => { entries[id] = notes[id]; });
  const payload = { schema: 'hr-dashboard-notes', version: 1, baseDate: D.meta.baseDate, resultFile: D.meta.resultFile, exportedAt: new Date().toISOString(), entries };
  downloadText(`확인기록_${D.meta.ym}.json`, JSON.stringify(payload, null, 2));
  say(`확인 기록 ${Object.keys(entries).length}건을 내보냈습니다. 담당자 이름이 들어 있으니 내부 파일로 다루세요.`, false);
});
$('#btn-import').addEventListener('click', () => $('#file-import').click());
$('#file-import').addEventListener('change', e => {
  const file = e.target.files[0]; e.target.value = ''; if (!file) return;
  const rd = new FileReader();
  rd.onload = () => {
    let j; try { j = JSON.parse(rd.result); } catch (err) { say('JSON 파일을 읽지 못했습니다. 이 대시보드에서 내보낸 파일인지 확인하세요.', true); return; }
    if (!j || j.schema !== 'hr-dashboard-notes' || typeof j.entries !== 'object' || j.entries === null) { say('이 대시보드에서 내보낸 확인 기록 파일이 아닙니다.', true); return; }
    if (j.baseDate !== D.meta.baseDate) { say(`기준일이 다른 기록입니다(파일 ${j.baseDate} / 이 대시보드 ${D.meta.baseDate}). 분석용 ID가 기준일마다 달라 불러오지 않았습니다.`, true); return; }
    const ids = new Set(P.map(p => p.id)); const next = {}; let skipped = 0, ok = 0;
    Object.keys(j.entries).forEach(id => {
      const v = j.entries[id];
      if (!ids.has(id) || !v || typeof v !== 'object') { skipped++; return; }
      const status = STATUSES.includes(v.status) ? v.status : '미확인';
      const owner = typeof v.owner === 'string' ? v.owner.trim().slice(0, 40) : '';
      const nx = typeof v.next === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v.next) ? v.next : '';
      if (owner || status !== '미확인' || nx) { next[id] = { owner, status, next: nx }; ok++; }
    });
    const have = Object.keys(notes).length;
    if (have && !window.confirm(`현재 이 브라우저에 저장된 확인 기록 ${have}건을 불러온 ${ok}건으로 바꿉니다. 계속할까요?`)) { say('불러오기를 취소했습니다.', false); return; }
    notes = next; saveNotes(); render();
    say(`확인 기록 ${ok}건을 불러왔습니다.${skipped ? ` 이 대시보드에 없는 분석용 ID ${skipped}건은 건너뛰었습니다.` : ''}`, false);
  };
  rd.readAsText(file, 'utf-8');
});
function say(text, warn) { $('#import-msg').innerHTML = `<div class="banner ${warn ? 'warn' : ''}">${esc(text)}</div>`; }

buildControls();
renderDetail($('#detail'), D.detail);
render();

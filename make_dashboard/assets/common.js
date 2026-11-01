'use strict';
/* 내부용·공유용 대시보드가 같이 쓰는 도우미. 개인 단위 자료를 다루는 코드는 여기에 넣지 않는다. */
const GRADES = ['상', '중', '하', '판단불가'];
const GCLS = { '상': 'g-hi', '중': 'g-mid', '하': 'g-lo', '판단불가': 'g-na' };
const SERIES_COLORS = ['#7aa7ff', '#f2a65a', '#b78cf4', '#5fd0c1', '#e8cf6a', '#f08fb0', '#8fd16b', '#d98cf0'];
const SCORE_NOTE = '점수는 퇴사 확률이 아니라, 먼저 확인할 대상을 정하는 데 쓰는 지표입니다.';

const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = (v, d = 1) => (v == null || Number.isNaN(v)) ? '—' : Number(v).toFixed(d);
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const gradeBadge = g => `<span class="grade ${GCLS[g] || 'g-na'}">${esc(g)}</span>`;
const reasonBox = text => `<div class="reason">${esc(text)}</div>`;

function readData() {
  const node = document.getElementById('hr-data');
  try { return JSON.parse(node.textContent); }
  catch (e) { document.body.innerHTML = '<div class="wrap"><div class="reason">데이터를 읽지 못했습니다. 대시보드를 다시 생성해 주세요.</div></div>'; throw e; }
}

function renderNotices(el, notices) {
  el.innerHTML = (notices || []).map(n => `<div class="banner ${n.level === 'warn' ? 'warn' : ''}">${esc(n.text)}</div>`).join('');
}

/* 요약 타일. filtered가 있으면 "필터 N명"을 작은 글씨로 함께 보인다. */
function renderTiles(el, o) {
  const tile = (k, v, f, cls) => `<div class="tile ${cls || ''}"><div class="k">${esc(k)}</div><div class="v">${esc(v)}</div><div class="f">${esc(f || '')}</div></div>`;
  const fl = o.filtered;
  const sub = n => fl ? `필터 ${n}명` : '';
  el.innerHTML =
    tile('분석 기준일', o.baseDate, '', '') +
    tile('총인원', o.total + '명', fl ? `필터 ${fl.total}명` : '', '') +
    tile('상', o.counts['상'] + '명', sub(fl && fl.counts['상']), 'g-hi') +
    tile('중', o.counts['중'] + '명', sub(fl && fl.counts['중']), 'g-mid') +
    tile('하', o.counts['하'] + '명', sub(fl && fl.counts['하']), 'g-lo') +
    (o.counts['판단불가'] ? tile('판단불가', o.counts['판단불가'] + '명', sub(fl && fl.counts['판단불가']), 'g-na') : '');
}

/* 히트맵: 칸 색은 열마다 최소~최대를 기준으로 한 가지 색의 진하기로 나타낸다. 칸 안에 숫자를 함께 쓴다. */
function heatColor(t) {
  const a = [38, 42, 49], b = [232, 176, 74];
  const c = a.map((x, i) => Math.round(x + (b[i] - x) * t));
  return { bg: `rgb(${c[0]},${c[1]},${c[2]})`, fg: t > 0.6 ? '#14110a' : '#e7e9ec' };
}
function renderHeatmap(el, o) {
  const cols = o.cols, rows = o.rows;
  const range = {};
  cols.forEach(c => {
    const vs = rows.filter(r => !r.suppressed && !r.missing).map(r => r.vals[c.key]).filter(v => v != null && !Number.isNaN(v));
    range[c.key] = vs.length ? [Math.min(...vs), Math.max(...vs)] : null;
  });
  const head = `<tr><th class="lab">부서</th><th class="num">인원</th>${cols.map(c => `<th title="${esc(c.title || c.label)}">${esc(c.label)}</th>`).join('')}</tr>`;
  const body = rows.map(r => {
    const lab = `<td class="lab">${esc(r.label)}</td>`;
    if (r.suppressed) {
      const why = r.protect ? '보호 처리 (소규모 조직 값을 역산할 수 없도록 함께 가림)' : `표본 부족 (${o.minGroup}명 미만)`;
      return `<tr class="${r.kind} protect">${lab}<td class="num na">—</td><td class="cell na" colspan="${cols.length}">${esc(why)}</td></tr>`;
    }
    if (r.missing) return `<tr class="${r.kind}">${lab}<td class="num na">0</td><td class="cell na" colspan="${cols.length}">필터에 해당하는 직원이 없습니다</td></tr>`;
    const cells = cols.map(c => {
      const v = r.vals[c.key], rg = range[c.key];
      if (v == null || Number.isNaN(v)) return '<td class="cell na">—</td>';
      const t = rg && rg[1] > rg[0] ? (v - rg[0]) / (rg[1] - rg[0]) : 0.5;
      const col = heatColor(t);
      return `<td class="cell" style="background:${col.bg};color:${col.fg}" title="${esc(c.label)}: ${fmt(v, c.dec)}">${fmt(v, c.dec)}</td>`;
    }).join('');
    return `<tr class="${r.kind}">${lab}<td class="num">${r.n}</td>${cells}</tr>`;
  }).join('');
  el.innerHTML = `<div class="tablewrap"><table class="heat"><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
}

/* 등급 분포: 가로 누적 막대 */
function renderDist(el, o) {
  const legend = `<div class="legend">${GRADES.map((g, i) => `<span><span class="sw" style="background:var(--${['hi', 'mid', 'lo', 'na'][i]})"></span>${g}</span>`).join('')}</div>`;
  const rows = o.rows.map(r => {
    const lab = `<div class="lab ${r.kind}">${esc(r.label)}</div>`;
    if (r.suppressed) return `${lab}<div class="muted small">${r.protect ? '보호 처리' : `표본 부족 (${o.minGroup}명 미만)`}</div><div></div>`;
    const n = r.n || 0;
    if (!n) return `${lab}<div class="muted small">필터에 해당하는 직원이 없습니다</div><div class="num">0명</div>`;
    const segs = r.counts.map((c, i) => c ? `<i class="s${i}" style="width:${(c / n * 100).toFixed(2)}%" title="${GRADES[i]} ${c}명">${c / n > 0.07 ? c : ''}</i>` : '').join('');
    return `${lab}<div class="stack" role="img" aria-label="${esc(r.label)} 등급 분포 ${GRADES.map((g, i) => g + ' ' + r.counts[i] + '명').join(', ')}">${segs}</div><div class="num small muted">${n}명</div>`;
  }).join('');
  el.innerHTML = legend + `<div class="dist">${rows}</div>`;
}

/* 꺾은선 차트 (SVG 문자열). 선 켜고 끄기, 마우스를 올리면 값 표시, 표로 보기를 함께 제공한다. */
function renderLineChart(el, o) {
  const months = o.months, series = o.series;
  const visible = new Set(series.map(s => s.key));
  el.innerHTML = '<div class="serbtns"></div><div class="chart"></div><details class="fold" style="margin-top:10px"><summary>표로 보기</summary><div class="body"></div></details>';
  const btns = $('.serbtns', el), chart = $('.chart', el), tbl = $('.body', el);
  series.forEach(s => {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'ghost'; b.setAttribute('aria-pressed', 'true');
    b.innerHTML = `<span class="sw" style="background:${s.color}"></span>${esc(s.label)}`;
    b.addEventListener('click', () => {
      if (visible.has(s.key)) visible.delete(s.key); else visible.add(s.key);
      b.setAttribute('aria-pressed', visible.has(s.key) ? 'true' : 'false'); draw();
    });
    btns.appendChild(b);
  });
  tbl.innerHTML = `<div class="tablewrap"><table><thead><tr><th>월</th>${series.map(s => `<th class="num">${esc(s.label)}</th>`).join('')}</tr></thead><tbody>${months.map((m, i) => `<tr><td>${esc(m)}</td>${series.map(s => `<td class="num">${fmt(s.values[i], 1)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;

  const W = 760, H = 300, ml = 46, mr = 16, mt = 14, mb = 34, pw = W - ml - mr, ph = H - mt - mb;
  function draw() {
    const shown = series.filter(s => visible.has(s.key));
    const vals = shown.flatMap(s => s.values).filter(v => v != null);
    let max = vals.length ? Math.max(...vals) : 1;
    const step = niceStep(max / 4); max = Math.ceil(max * 1.05 / step) * step || step;
    const x = i => ml + (months.length > 1 ? pw * i / (months.length - 1) : pw / 2);
    const y = v => mt + ph - (v / max) * ph;
    let g = '', ticks = '';
    for (let v = 0; v <= max + 1e-9; v += step) { g += `<line x1="${ml}" x2="${W - mr}" y1="${y(v)}" y2="${y(v)}"/>`; ticks += `<text x="${ml - 6}" y="${y(v) + 4}" text-anchor="end">${fmt(v, step < 1 ? 1 : 0)}</text>`; }
    const every = Math.max(1, Math.ceil(months.length / 8));
    const xl = months.map((m, i) => i % every === 0 || i === months.length - 1 ? `<text x="${x(i)}" y="${H - 12}" text-anchor="middle">${esc(m.slice(2))}</text>` : '').join('');
    const lines = shown.map(s => {
      let d = '', pen = false, dots = '';
      s.values.forEach((v, i) => {
        if (v == null) { pen = false; return; }
        d += `${pen ? 'L' : 'M'}${x(i).toFixed(1)} ${y(v).toFixed(1)} `; pen = true;
        dots += `<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="2.6" fill="${s.color}"/>`;
      });
      return `<path d="${d}" fill="none" stroke="${s.color}" stroke-width="${s.dash ? 2 : 2.4}" ${s.dash ? 'stroke-dasharray="6 4"' : ''} stroke-linecap="round" stroke-linejoin="round"/>${dots}`;
    }).join('');
    chart.innerHTML = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="월간 연장근로 시간 추이"><g class="grid">${g}</g>${ticks}${xl}${lines}<line class="guide" y1="${mt}" y2="${mt + ph}" stroke="#5b626d" stroke-width="1" style="display:none"/></svg><div class="tip" style="display:none"></div><div class="small muted" style="margin-top:4px">${esc(o.unit || '')}${shown.length ? '' : ' · 표시할 선을 위 버튼에서 켜 주세요'}</div>`;
    const svg = $('svg', chart), tip = $('.tip', chart), guide = $('.guide', svg);
    svg.addEventListener('mousemove', ev => {
      const r = svg.getBoundingClientRect(), vx = (ev.clientX - r.left) / r.width * W;
      let i = Math.round((vx - ml) / pw * (months.length - 1)); i = Math.max(0, Math.min(months.length - 1, i));
      guide.setAttribute('x1', x(i)); guide.setAttribute('x2', x(i)); guide.style.display = '';
      tip.innerHTML = `<b>${esc(months[i])}</b>` + shown.map(s => `<div class="r"><span style="color:${s.color}">${esc(s.label)}</span><span>${fmt(s.values[i], 1)}</span></div>`).join('');
      tip.style.display = '';
      const cr = chart.getBoundingClientRect(); let left = ev.clientX - cr.left + 14;
      if (left + tip.offsetWidth > cr.width) left = ev.clientX - cr.left - tip.offsetWidth - 14;
      tip.style.left = Math.max(0, left) + 'px'; tip.style.top = Math.max(0, ev.clientY - cr.top - 20) + 'px';
    });
    svg.addEventListener('mouseleave', () => { tip.style.display = 'none'; guide.style.display = 'none'; });
  }
  draw();
}
function niceStep(raw) {
  if (!(raw > 0)) return 1;
  const p = Math.pow(10, Math.floor(Math.log10(raw))), f = raw / p;
  return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
}

/* 상세 설명: 기본은 접힌 상태 */
function tableHtml(head, rows) {
  const th = (head || []).map(c => `<th>${esc(c)}</th>`).join('');
  const tr = (rows || []).map(r => `<tr>${(head || r).map((_, i) => `<td style="white-space:normal">${esc(r[i] == null ? '' : r[i])}</td>`).join('')}</tr>`).join('');
  return `<div class="tablewrap"><table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table></div>`;
}
function renderDetail(el, d) {
  const hyp = (d.hypotheses || []).length ? `<h3>가설</h3><p class="muted small">아래는 검증하기 전의 가정입니다. 검증 결과는 이 가정이 실제 퇴사자 쪽에서 더 많이 나타났는지 사후에 확인한 값입니다.</p><ul>${d.hypotheses.map(h => `<li><b>${esc(h[0])}</b> — ${esc(h[1])}</li>`).join('')}</ul>` : '';
  const crit = d.criteria ? `<h3>점수 산정 기준</h3>${tableHtml(d.criteria.head, d.criteria.rows)}` : '';
  const val = (d.validation || []).map(b => `<h3>${esc(b.title)}</h3>${(b.lines || []).length ? `<ul>${b.lines.map(l => `<li>${esc(l)}</li>`).join('')}</ul>` : ''}${(b.tables || []).map(t => tableHtml(t.head, t.rows)).join('')}`).join('');
  const checks = d.dataChecks ? `<h3>데이터 점검</h3>${tableHtml(d.dataChecks.head, d.dataChecks.rows)}` : '';
  const cond = (d.conditions || []).length ? `<h3>분석 조건</h3><ul>${d.conditions.map(l => `<li>${esc(l)}</li>`).join('')}</ul>` : '';
  el.innerHTML = `<details class="fold"><summary>가설 · 검증 결과 · 분석 한계 (펼쳐서 보기)</summary><div class="body">
    <div class="banner warn"><b>분석의 한계</b><ul>${(d.limits || []).map(l => `<li>${esc(l)}</li>`).join('')}</ul></div>
    ${hyp}${val}${crit}${cond}${checks}</div></details>`;
}

function downloadText(name, text, mime) {
  const blob = new Blob([text], { type: mime || 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

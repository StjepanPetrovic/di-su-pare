import * as XLSX from 'xlsx';
import './style.css';

const STORE = 'mirno-finance-v1';
const app = document.querySelector('#app');
const money = new Intl.NumberFormat('hr-HR', { style: 'currency', currency: 'EUR' });
const monthFmt = new Intl.DateTimeFormat('hr-HR', { month: 'long', year: 'numeric' });
let state = loadState();
let activePeriod = { type: 'all' };
let query = '';
let notice = '';

function loadState() {
  try { return JSON.parse(localStorage.getItem(STORE)) || { transactions: [], rules: {} }; }
  catch { return { transactions: [], rules: {} }; }
}
function save() { localStorage.setItem(STORE, JSON.stringify(state)); }
function esc(s = '') { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function groupKey(desc) { return desc.toLocaleLowerCase('hr').replace(/[0-9]+/g, ' ').replace(/\b(hr|ee|de|ie|com|eu)\b/g, ' ').replace(/\s+/g, ' ').trim(); }
function fallbackGroup(t) {
  const d = t.description.toLocaleLowerCase('hr');
  if (/prijenos na vlastiti račun|isplata plaće|uplata plaće|povrat|refund/.test(d)) return 'Prihodi i prijenosi';
  if (/kupovina|kartica|nalog|plaćanje|placanje/.test(d)) return 'Ostalo';
  return 'Nepoznato';
}
function groupOf(t) { return state.rules[groupKey(t.description)] || t.group || fallbackGroup(t); }
function parseDate(v, wb) {
  if (typeof v === 'number') { const d = XLSX.SSF.parse_date_code(v, { date1904: !!wb?.Workbook?.WBProps?.date1904 }); if (d) return new Date(Date.UTC(d.y, d.m - 1, d.d)).toISOString().slice(0, 10); }
  const d = new Date(v); return Number.isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
}
function parseWorkbook(file) {
  return file.arrayBuffer().then(buf => {
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const headerAt = rows.findIndex(r => r.some(x => String(x).trim().toLocaleLowerCase('hr') === 'datum'));
    if (headerAt < 0) throw new Error('Nisam pronašao red zaglavlja s datumom. Provjerite je li ovo ZABA izvoz.');
    const headers = rows[headerAt].map(x => String(x).trim().toLocaleLowerCase('hr'));
    const col = name => headers.findIndex(h => h === name);
    const dateI = col('datum'), descI = col('opis'), inI = col('uplata'), outI = col('isplata'), currI = col('valuta');
    if ([dateI, descI, inI, outI].some(i => i < 0)) throw new Error('Nedostaju očekivani stupci Datum, Opis, Uplata ili Isplata.');
    return rows.slice(headerAt + 1).filter(r => r[descI] && r[dateI] !== '').map((r, i) => ({
      id: `${file.name}-${i}-${r[1] || ''}`, date: parseDate(r[dateI], wb), description: String(r[descI]).trim(),
      income: Number(r[inI]) || 0, expense: Number(r[outI]) || 0, balance: Number(r[col('saldo')]) || 0, currency: String(r[currI] || 'EUR'), group: ''
    })).filter(t => t.date);
  });
}
function periodRows() {
  return state.transactions.filter(t => {
    const [year, month, day] = t.date.split('-');
    if (activePeriod.year && year !== activePeriod.year) return false;
    if (activePeriod.month && month !== activePeriod.month) return false;
    if (activePeriod.dayFrom && Number(day) < Number(activePeriod.dayFrom)) return false;
    if (activePeriod.dayTo && Number(day) > Number(activePeriod.dayTo)) return false;
    return true;
  });
}
function totals(rows) { return rows.reduce((a, t) => ({ income: a.income + t.income, expense: a.expense + t.expense }), { income: 0, expense: 0 }); }
function balanceChart(rows) {
  const points = [...rows].sort((a,b) => a.date.localeCompare(b.date) || a.id.localeCompare(b.id));
  if (!points.length) return '<div class="empty-small">Nema transakcija za prikaz stanja u ovom periodu.</div>';
  const values = points.map(t => t.balance).filter(Number.isFinite);
  if (!values.length || values.every(v => v === 0)) return '<div class="empty-small">Izvod ne sadrži saldo za izradu grafa.</div>';
  const width = 760, height = 230, left = 66, right = 14, top = 18, bottom = 34;
  let min = Math.min(...values), max = Math.max(...values);
  if (min === max) { min -= Math.abs(min)*.05 || 1; max += Math.abs(max)*.05 || 1; }
  const pad = (max-min)*.12; min -= pad; max += pad;
  const x = i => left + (points.length === 1 ? (width-left-right)/2 : i*(width-left-right)/(points.length-1));
  const y = value => top + (max-value)*(height-top-bottom)/(max-min);
  const d = points.map((t,i)=>`${i?'L':'M'}${x(i).toFixed(1)},${y(t.balance).toFixed(1)}`).join(' ');
  const yTicks = Array.from({length:4},(_,i)=>max-i*(max-min)/3);
  const maxLabels = 5, stride = Math.max(1, Math.ceil(points.length/maxLabels));
  const xIndices = [...new Set([0,...Array.from({length:Math.ceil(points.length/stride)},(_,i)=>i*stride),points.length-1])].filter(i=>i<points.length);
  return `<div class="chart-scroll"><svg class="balance-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="Kretanje stanja računa u odabranom periodu">${yTicks.map(v=>`<line class="chart-grid" x1="${left}" x2="${width-right}" y1="${y(v)}" y2="${y(v)}"/><text class="chart-label" x="${left-9}" y="${y(v)+4}" text-anchor="end">${esc(new Intl.NumberFormat('hr-HR',{notation:'compact',maximumFractionDigits:1}).format(v))} €</text>`).join('')}<path class="chart-area" d="${d} L${x(points.length-1)},${height-bottom} L${x(0)},${height-bottom} Z"/><path class="chart-line" d="${d}"/>${xIndices.map(i=>`<text class="chart-label" x="${x(i)}" y="${height-9}" text-anchor="${i===0?'start':i===points.length-1?'end':'middle'}">${new Date(`${points[i].date}T12:00:00`).toLocaleDateString('hr-HR',{day:'numeric',month:'short'})}</text>`).join('')}${points.map((t,i)=>`<circle class="chart-point" cx="${x(i)}" cy="${y(t.balance)}" r="8"><title>${new Date(`${t.date}T12:00:00`).toLocaleDateString('hr-HR')} · ${money.format(t.balance)}</title></circle>`).join('')}</svg></div><p class="chart-hint">Dodirnite ili prijeđite pokazivačem preko grafa za datum i saldo.</p>`;
}
function render() {
  const years = [...new Set(state.transactions.map(t => t.date.slice(0, 4)))].sort();
  if (activePeriod.year && !years.includes(activePeriod.year)) activePeriod.year = '';
  const rows = periodRows(), sum = totals(rows);
  const periodLabel = activePeriod.month && activePeriod.year
    ? `${new Intl.DateTimeFormat('hr-HR', { month: 'long' }).format(new Date(2020, Number(activePeriod.month) - 1, 1))} ${activePeriod.year}.`
    : activePeriod.year ? `${activePeriod.year}. godina`
      : activePeriod.month ? `${new Intl.DateTimeFormat('hr-HR', { month: 'long' }).format(new Date(2020, Number(activePeriod.month) - 1, 1))} — sve godine`
        : activePeriod.dayFrom || activePeriod.dayTo ? 'Dani kroz cijeli period' : 'Cijeli period';
  const maxDay = activePeriod.year && activePeriod.month ? new Date(Number(activePeriod.year), Number(activePeriod.month), 0).getDate() : 31;
  const selectedDayFrom = Math.min(Number(activePeriod.dayFrom) || 1, maxDay), selectedDayTo = Math.min(Number(activePeriod.dayTo) || maxDay, maxDay);
  const filtered = rows.filter(t => t.description.toLocaleLowerCase('hr').includes(query.toLocaleLowerCase('hr'))).slice(0, 80);
  app.innerHTML = `<header class="top"><a class="brand" href="#" aria-label="Di su pare?!"><svg class="brand-logo" viewBox="0 0 196 42" role="img" aria-labelledby="logo-title"><title id="logo-title">?! DI SU PARE</title><rect x="1" y="1" width="179" height="40" rx="12" fill="#203b31"/><text x="14" y="28" fill="#f5c95d" font-family="Manrope, sans-serif" font-size="20" font-weight="800">?!</text><text x="38" y="28" fill="#fff" font-family="Manrope, sans-serif" font-size="18" font-weight="800" letter-spacing=".3">DI SU PARE</text></svg></a><span class="privacy"><i></i> Podaci ostaju na ovom uređaju</span></header>
  <main><section class="intro"><div class="eyebrow">FINANCIRANJA, BEZ KOMPLICIRANJA</div><h1>DI<br><em>idu tvoje pare?</em></h1><p>Uvezi izvod transakcija i brzo provjeri koliko trošiš, koliko ti sjeda i kako se mijenja stanje računa. Sve ostaje na ovom uređaju.</p></section>
  ${notice ? `<div class="notice">${esc(notice)} <button data-action="dismiss" aria-label="Zatvori">×</button></div>` : ''}
  ${state.transactions.length ? `<section class="dashboard"><div class="section-head"><div><div class="eyebrow">VAŠ PREGLED</div><h2>${esc(periodLabel)}</h2></div></div><div class="period-panel"><div class="period-heading">Odaberi period</div><div class="period-bar"><label class="period-field"><span>GODINA</span><select id="year-select"><option value="">Sve godine</option>${years.map(y=>`<option value="${y}" ${activePeriod.year === y ? 'selected' : ''}>${y}.</option>`).join('')}</select></label><label class="period-field"><span>MJESEC</span><select id="month-select"><option value="">Svi mjeseci</option>${Array.from({length:12},(_,i)=>String(i+1).padStart(2,'0')).map(m=>`<option value="${m}" ${activePeriod.month === m ? 'selected' : ''}>${new Intl.DateTimeFormat('hr-HR',{month:'long'}).format(new Date(2020,Number(m)-1,1))}</option>`).join('')}</select></label><label class="period-field day-field"><span>DAN OD</span><select id="day-from" ${!activePeriod.month || !activePeriod.year ? 'disabled' : ''}><option value="">1.</option>${Array.from({length:maxDay},(_,i)=>i+1).map(d=>`<option value="${d}" ${Number(activePeriod.dayFrom) === d ? 'selected' : ''}>${d}.</option>`).join('')}</select></label><label class="period-field day-field"><span>DAN DO</span><select id="day-to" ${!activePeriod.month || !activePeriod.year ? 'disabled' : ''}><option value="">${maxDay}.</option>${Array.from({length:maxDay},(_,i)=>i+1).map(d=>`<option value="${d}" ${Number(activePeriod.dayTo) === d ? 'selected' : ''}>${d}.</option>`).join('')}</select></label><span class="period-hint">${activePeriod.month && activePeriod.year ? `Dani se odnose na ${new Intl.DateTimeFormat('hr-HR',{month:'long'}).format(new Date(2020,Number(activePeriod.month)-1,1))} ${activePeriod.year}.` : 'Za filtriranje po danima prvo odaberite godinu i mjesec.'}</span><button class="reset-period" data-action="reset-period" ${!activePeriod.year && !activePeriod.month && !activePeriod.dayFrom && !activePeriod.dayTo ? 'disabled' : ''}>Očisti</button></div></div>
    <div class="cards"><article class="metric"><span>Otišlo</span><strong>${money.format(sum.expense)}</strong><small>u odabranom periodu</small></article><article class="metric"><span>Sjelo</span><strong>${money.format(sum.income)}</strong><small>u odabranom periodu</small></article><article class="metric balance"><span>Ostalo nakon troškova</span><strong>${money.format(sum.income - sum.expense)}</strong><small>prihodi umanjeni za potrošnju</small></article></div>
    <article class="panel balance-panel"><div class="panel-title"><div><h3>Graf kretanja</h3><p>Stanje računa nakon transakcija u odabranom periodu.</p></div></div>${balanceChart(rows)}</article>
    <div class="content-grid single-column"><article class="panel"><div class="panel-title"><div><h3>Di su točno odlazile i dolazile</h3><p>${rows.length} transakcija u odabranom periodu</p></div></div><input class="search" id="search" placeholder="Pronađi transakciju…" value="${esc(query)}"><div class="transactions">${filtered.map(t => `<div class="transaction"><div class="tx-main"><span class="tx-desc">${esc(t.description)}</span><span class="tx-meta">${new Date(`${t.date}T12:00:00`).toLocaleDateString('hr-HR')}</span></div><strong class="${t.income ? 'positive' : ''}">${t.income ? '+' : '−'}${money.format(t.income || t.expense)}</strong></div>`).join('') || '<div class="empty-small">Nema transakcija za ovaj period.</div>'}</div></article></div>
    <div class="actions"><button class="button secondary" data-action="clear">Ukloni podatke s uređaja</button><label class="button primary">Uvezi još transakcija<input id="file" type="file" accept=".xls,.xlsx,.csv" hidden></label></div>
  </section>` : `<section class="import-card"><div class="file-icon">↥</div><h2>Ajmo vidjeti di su pare</h2><p>Uvezi Excel ili CSV izvod transakcija iz ZABA internet bankarstva i odmah pogledaj kamo novac odlazi. Datoteka se obrađuje samo u pregledniku.</p><label class="button primary">Uvezi izvod<input id="file" type="file" accept=".xls,.xlsx,.csv" hidden></label><span class="formats">Podržani formati: .xls, .xlsx i .csv</span></section>`}
  <section class="trust"><span class="shield">◇</span><div><strong>Pare su tvoje. Podaci isto.</strong><p>Izvod se obrađuje u pregledniku, a podaci ostaju na ovom uređaju.</p></div></section>
  <footer>?! DI SU PARE <span>Uvezi. Pogledaj. Saznaj di idu pare.</span></footer></main>`;
  bind();
}
function bind() {
  document.querySelector('#file')?.addEventListener('change', async e => {
    const file = e.target.files?.[0]; if (!file) return;
    notice = 'Obrađujem datoteku na ovom uređaju…'; render();
    try {
      const incoming = await parseWorkbook(file), seen = new Set(state.transactions.map(t => `${t.date}|${t.description}|${t.income}|${t.expense}`));
      const fresh = incoming.filter(t => !seen.has(`${t.date}|${t.description}|${t.income}|${t.expense}`));
      state.transactions.push(...fresh); save(); notice = `Uvezeno ${fresh.length} transakcija${incoming.length !== fresh.length ? ` (preskočeno ${incoming.length - fresh.length} duplikata)` : ''}.`; activeMonth = fresh.at(-1)?.date.slice(0, 7) || activeMonth;
    } catch (err) { notice = err.message || 'Datoteku nije bilo moguće pročitati.'; }
    render();
  });
  const updatePeriod = () => { const year = document.querySelector('#year-select').value, month = document.querySelector('#month-select').value; let dayFrom = document.querySelector('#day-from').value, dayTo = document.querySelector('#day-to').value; const max = year && month ? new Date(Number(year), Number(month), 0).getDate() : 31; if (!year || !month) dayFrom = dayTo = ''; if (Number(dayFrom) > max) dayFrom = String(max); if (Number(dayTo) > max) dayTo = String(max); activePeriod = { year, month, dayFrom, dayTo }; render(); };
  ['#month-select', '#year-select', '#day-from', '#day-to'].forEach(selector => document.querySelector(selector)?.addEventListener('change', updatePeriod));
  document.querySelector('[data-action="reset-period"]')?.addEventListener('click', () => { activePeriod = { type: 'all', year: '', month: '', dayFrom: '', dayTo: '' }; render(); });
  document.querySelector('#search')?.addEventListener('input', e => { const p = e.target.selectionStart; query = e.target.value; render(); const s = document.querySelector('#search'); s.focus(); s.setSelectionRange(p, p); });
  document.querySelector('[data-action="dismiss"]')?.addEventListener('click', () => { notice = ''; render(); });
  document.querySelector('[data-action="clear"]')?.addEventListener('click', () => { if (confirm('Trajno ukloniti sve uvezene transakcije i pravila s ovog uređaja?')) { state = { transactions: [], rules: {} }; save(); notice = 'Podaci su uklonjeni.'; render(); } });
}
render();

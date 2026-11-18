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
      income: Number(r[inI]) || 0, expense: Number(r[outI]) || 0, currency: String(r[currI] || 'EUR'), group: ''
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
function render() {
  const years = [...new Set(state.transactions.map(t => t.date.slice(0, 4)))].sort();
  if (activePeriod.year && !years.includes(activePeriod.year)) activePeriod.year = '';
  const rows = periodRows(), sum = totals(rows);
  const periodLabel = activePeriod.month && activePeriod.year
    ? `${new Intl.DateTimeFormat('hr-HR', { month: 'long' }).format(new Date(2020, Number(activePeriod.month) - 1, 1))} ${activePeriod.year}.`
    : activePeriod.year ? `${activePeriod.year}. godina`
      : activePeriod.month ? `${new Intl.DateTimeFormat('hr-HR', { month: 'long' }).format(new Date(2020, Number(activePeriod.month) - 1, 1))} — sve godine`
        : activePeriod.dayFrom || activePeriod.dayTo ? 'Dani kroz cijelo razdoblje' : 'Cijelo razdoblje';
  const maxDay = activePeriod.year && activePeriod.month ? new Date(Number(activePeriod.year), Number(activePeriod.month), 0).getDate() : 31;
  const selectedDayFrom = Math.min(Number(activePeriod.dayFrom) || 1, maxDay), selectedDayTo = Math.min(Number(activePeriod.dayTo) || maxDay, maxDay);
  const filtered = rows.filter(t => t.description.toLocaleLowerCase('hr').includes(query.toLocaleLowerCase('hr'))).slice(0, 80);
  app.innerHTML = `<header class="top"><a class="brand" href="#"><span class="brand-mark">m</span> mirno</a><span class="privacy"><i></i> Podaci ostaju na ovom uređaju</span></header>
  <main><section class="intro"><div class="eyebrow">OSOBNE FINANCIJE, BEZ KOMPLICIRANJA</div><h1>Znajte kamo<br><em>novac odlazi.</em></h1><p>Uvezite izvod i dobijte jasan pregled troškova. Vaši podaci ne napuštaju ovaj uređaj.</p></section>
  ${notice ? `<div class="notice">${esc(notice)} <button data-action="dismiss" aria-label="Zatvori">×</button></div>` : ''}
  ${state.transactions.length ? `<section class="dashboard"><div class="section-head"><div><div class="eyebrow">VAŠ PREGLED</div><h2>${esc(periodLabel)}</h2></div></div><div class="period-panel"><div class="period-heading">Period</div><div class="period-bar"><label class="period-field"><span>GODINA</span><select id="year-select"><option value="">Sve godine</option>${years.map(y=>`<option value="${y}" ${activePeriod.year === y ? 'selected' : ''}>${y}.</option>`).join('')}</select></label><label class="period-field"><span>MJESEC</span><select id="month-select"><option value="">Svi mjeseci</option>${Array.from({length:12},(_,i)=>String(i+1).padStart(2,'0')).map(m=>`<option value="${m}" ${activePeriod.month === m ? 'selected' : ''}>${new Intl.DateTimeFormat('hr-HR',{month:'long'}).format(new Date(2020,Number(m)-1,1))}</option>`).join('')}</select></label><label class="period-field day-field"><span>DAN OD</span><select id="day-from" ${!activePeriod.month || !activePeriod.year ? 'disabled' : ''}><option value="">1.</option>${Array.from({length:maxDay},(_,i)=>i+1).map(d=>`<option value="${d}" ${Number(activePeriod.dayFrom) === d ? 'selected' : ''}>${d}.</option>`).join('')}</select></label><label class="period-field day-field"><span>DAN DO</span><select id="day-to" ${!activePeriod.month || !activePeriod.year ? 'disabled' : ''}><option value="">${maxDay}.</option>${Array.from({length:maxDay},(_,i)=>i+1).map(d=>`<option value="${d}" ${Number(activePeriod.dayTo) === d ? 'selected' : ''}>${d}.</option>`).join('')}</select></label><span class="period-hint">${activePeriod.month && activePeriod.year ? `Dani se odnose na ${new Intl.DateTimeFormat('hr-HR',{month:'long'}).format(new Date(2020,Number(activePeriod.month)-1,1))} ${activePeriod.year}.` : 'Za filtriranje po danima prvo odaberite godinu i mjesec.'}</span><button class="reset-period" data-action="reset-period" ${!activePeriod.year && !activePeriod.month && !activePeriod.dayFrom && !activePeriod.dayTo ? 'disabled' : ''}>Očisti</button></div></div>
    <div class="cards"><article class="metric"><span>Potrošeno</span><strong>${money.format(sum.expense)}</strong><small>ukupni izdaci u mjesecu</small></article><article class="metric"><span>Prihodi</span><strong>${money.format(sum.income)}</strong><small>ukupne uplate u mjesecu</small></article><article class="metric balance"><span>Razlika</span><strong>${money.format(sum.income - sum.expense)}</strong><small>prihodi umanjeni za izdatke</small></article></div>
    <div class="content-grid single-column"><article class="panel"><div class="panel-title"><div><h3>Transakcije</h3><p>${rows.length} zapisa u odabranom periodu</p></div></div><input class="search" id="search" placeholder="Pretraži transakcije…" value="${esc(query)}"><div class="transactions">${filtered.map(t => `<div class="transaction"><div class="tx-main"><span class="tx-desc">${esc(t.description)}</span><span class="tx-meta">${new Date(`${t.date}T12:00:00`).toLocaleDateString('hr-HR')}</span></div><strong class="${t.income ? 'positive' : ''}">${t.income ? '+' : '−'}${money.format(t.income || t.expense)}</strong></div>`).join('') || '<div class="empty-small">Nema transakcija.</div>'}</div></article></div>
    <div class="actions"><button class="button secondary" data-action="export">Izvezi sigurnosnu kopiju</button><button class="button secondary" data-action="clear">Ukloni podatke s uređaja</button><label class="button primary">Uvezi još transakcija<input id="file" type="file" accept=".xls,.xlsx,.csv" hidden></label></div>
  </section>` : `<section class="import-card"><div class="file-icon">↥</div><h2>Počnite s izvodom</h2><p>Odaberite Excel ili CSV datoteku preuzetu iz ZABA internet bankarstva. Analiza se odvija lokalno u vašem pregledniku.</p><label class="button primary">Odaberite datoteku<input id="file" type="file" accept=".xls,.xlsx,.csv" hidden></label><span class="formats">Podržani formati: .xls, .xlsx, .csv</span></section>`}
  <section class="trust"><span class="shield">◇</span><div><strong>Vaša privatnost je na prvom mjestu</strong><p>Datoteka se obrađuje u pregledniku. Podaci se spremaju samo u memoriju ovog uređaja. Izvezite kopiju ako želite sačuvati pravila i transakcije.</p></div></section>
  <footer>mirno <span>Jednostavniji odnos s novcem.</span></footer></main>`;
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
  document.querySelector('[data-action="export"]')?.addEventListener('click', () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' }), url = URL.createObjectURL(blob), a = document.createElement('a'); a.href = url; a.download = 'mirno-sigurnosna-kopija.json'; a.click(); URL.revokeObjectURL(url);
  });
  document.querySelector('[data-action="clear"]')?.addEventListener('click', () => { if (confirm('Trajno ukloniti sve uvezene transakcije i pravila s ovog uređaja?')) { state = { transactions: [], rules: {} }; save(); notice = 'Podaci su uklonjeni.'; render(); } });
}
render();

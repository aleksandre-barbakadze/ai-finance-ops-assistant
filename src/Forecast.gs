// Forecast.gs – 13-კვირიანი ფულადი ნაკადის პროგნოზი. ვარაუდები სინთეტიკურია!
function buildForecast() {
  const ss = SpreadsheetApp.getActive(), tz = Session.getScriptTimeZone();
  const r2 = x => Math.round(x * 100) / 100;
  const iso = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').trim();
  const add = (d, n) => Utilities.formatDate(new Date(new Date(d + 'T12:00:00').getTime() + n * 864e5), tz, 'yyyy-MM-dd');
  const wday = d => new Date(d + 'T12:00:00').getDay();
  const read = name => { const v = ss.getSheetByName(name).getDataRange().getValues(), h = v[0].map(String);
    return v.slice(1).filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; }); };
  const S = {}; ss.getSheetByName('Settings').getDataRange().getValues().forEach(r => S[r[0]] = r[1]);
  const start = iso(S['Forecast start']), weeks = Number(S['Forecast weeks']), opening = Number(S['Opening cash (GEL)']);
  const buffer = Number(S['Minimum cash buffer (GEL)']), cut = Number(S['Pessimistic sales cut (%)']) / 100;
  const fee = Number(S['PayKart fee (%)']) / 100;
  if (!start || !weeks || !opening) throw new Error('Settings: Forecast start / weeks / Opening cash ცარიელია');
  const end = add(start, weeks * 7 - 1);
  const week = d => { const i = Math.floor((new Date(d + 'T12:00:00') - new Date(start + 'T12:00:00')) / (7 * 864e5)); return i >= 0 && i < weeks ? i : -1; };

  // კურსი: ყველაზე ახალი ინვოისის კურსი თითოეულ ვალუტაზე
  const inv = read('Invoices');
  const fx = { GEL: 1 }, fxDate = {};
  inv.forEach(o => { const c = o['Currency'], d = iso(o['Invoice date']), r = Number(o['FX rate']);
    if (c !== 'GEL' && r && (!fxDate[c] || d > fxDate[c])) { fx[c] = r; fxDate[c] = d; } });
  const gel = (a, c) => { if (!fx[c]) throw new Error('კურსი არ არის: ' + c); return r2(a * fx[c]); };

  const ev = [];   // dB/dP – თარიღი საბაზისოში/პესიმისტურში; aB/aP – თანხა (+ შემოსავალი, − გასავალი)
  const push = (dB, dP, item, type, aB, aP) => ev.push({ dB, dP, item, type, aB, aP });

  // 1. გაყიდვები: ბარათი – ორშაბათს (საკომისიოს გამოკლებით), ნაღდი – პარასკევს
  const sales = {}; read('Forecast sales').forEach(o => sales[String(o['Month']).trim()] = o);
  const count = (ym, wd) => { let n = 0; for (let d = ym + '-01'; d.slice(0, 7) === ym; d = add(d, 1)) if (wday(d) === wd) n++; return n; };
  for (let d = start; d <= end; d = add(d, 1)) {
    const ym = d.slice(0, 7), s = sales[ym]; if (!s) throw new Error('Forecast sales: თვე არ არის ' + ym);
    if (wday(d) === 1) { const a = r2(Number(s['Card sales (GEL)']) / count(ym, 1) * (1 - fee)); push(d, d, 'PayKart payout', 'Sales', a, r2(a * (1 - cut))); }
    if (wday(d) === 5) { const a = r2(Number(s['COD sales (GEL)']) / count(ym, 5)); push(d, d, 'COD remittance', 'Sales', a, r2(a * (1 - cut))); }
  }

  // 2. დებიტორები (reconcileReceivables-ის შედეგიდან)
  let arSkipped = 0;
  read('Sales invoices').forEach(o => {
    const out = Number(o['Outstanding (GEL)']), late = Number(o['Days late / overdue']), due = iso(o['Due date']);
    if (!(out > 0.009)) return;
    let dB, dP;
    if (!late) { dB = due < start ? start : due; dP = add(due, 14); }
    else if (late <= 60) { dB = add(start, 7); dP = add(start, 28); }
    else { arSkipped = r2(arSkipped + out); return; }          // 61+ – კონსერვატიულად არ ვითვლით
    push(dB, dP, o['Invoice No'] + ' (customer)', 'Receivables', out, out);
  });

  // 3. დადასტურებული გადასახდელები (Approved, არა HOLD/BLOCKED)
  inv.forEach(o => {
    const dec = String(o['Payment decision'] || '').toUpperCase();
    const open = r2(Number(o['Total amount']) - (Number(o['Paid amount']) || 0));
    if (o['Status'] !== 'Approved' || dec === 'HOLD' || dec === 'BLOCKED' || open <= 0.009) return;
    let d = iso(o['Due date']) || start; if (d < start) d = start;
    const a = -gel(open, o['Currency']);
    push(d, d, o['Invoice ID'] + ' ' + o['Supplier name'], 'Approved payables', a, a);
  });

  // 4. ხარჯები: ყოველთვიური (Day/From/To) და ერთჯერადი (One-off date)
  read('Forecast costs').forEach(o => {
    const a = -gel(Number(o['Amount']), String(o['Currency']).trim());
    const pessOnly = String(o['Scenario']).trim().toUpperCase() === 'PESSIMISTIC';
    const type = o['One-off date'] ? 'One-off (held/blocked)' : (/inventory/i.test(o['Item']) ? 'Inventory' : 'Recurring');
    const dates = [];
    if (o['One-off date']) dates.push(iso(o['One-off date']));
    else for (let ym = String(o['From']).trim(); ym <= String(o['To']).trim(); ym = add(ym + '-28', 7).slice(0, 7)) {
      const last = new Date(Number(ym.slice(0, 4)), Number(ym.slice(5, 7)), 0).getDate();
      dates.push(ym + '-' + ('0' + Math.min(Number(o['Day']), last)).slice(-2));
    }
    dates.forEach(d => push(d, d, o['Item'], type, pessOnly ? 0 : a, a));
  });

  // 5. კვირებად დაჯამება
  const W = Array.from({ length: weeks }, (_, i) => ({ from: add(start, 7 * i), to: add(start, 7 * i + 6), inB: 0, outB: 0, inP: 0, outP: 0 }));
  ev.forEach(e => {
    const iB = week(e.dB), iP = week(e.dP);
    if (iB >= 0) { if (e.aB > 0) W[iB].inB += e.aB; else W[iB].outB -= e.aB; }
    if (iP >= 0) { if (e.aP > 0) W[iP].inP += e.aP; else W[iP].outP -= e.aP; }
  });
  let cB = opening, cP = opening, minB = [Infinity, 0], minP = [Infinity, 0];
  const rows = W.map((w, i) => {
    cB = r2(cB + w.inB - w.outB); cP = r2(cP + w.inP - w.outP);
    if (cB < minB[0]) minB = [cB, i + 1]; if (cP < minP[0]) minP = [cP, i + 1];
    const flag = [cB < buffer ? 'BASE below buffer' : '', cP < buffer ? 'PESSIMISTIC below buffer' : ''].filter(Boolean).join(', ');
    Logger.log('W' + (i + 1) + ' ' + w.from + ': base ' + cB + ' | pess ' + cP + (flag ? '  ⚠️ ' + flag : ''));
    return [i + 1, w.from, w.to, r2(w.inB), r2(w.outB), cB, r2(w.inP), r2(w.outP), cP, flag];
  });

  // 6. ჩაწერა: Cash forecast (კვირები) + Forecast detail (ყველა ჩანაწერი)
  const sh = ss.getSheetByName('Cash forecast') || ss.insertSheet('Cash forecast'); sh.clear();
  sh.getRange(1, 1).setValue('13-WEEK CASH FORECAST (synthetic assumptions) | opening ' + opening + ' GEL | buffer ' + buffer + ' | USD ' + fx.USD + ' (' + fxDate.USD + ')').setFontWeight('bold');
  const head = ['Week', 'From', 'To', 'In (base)', 'Out (base)', 'Closing (base)', 'In (pess.)', 'Out (pess.)', 'Closing (pess.)', 'Warning'];
  sh.getRange(3, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.getRange(4, 1, rows.length, head.length).setValues(rows);
  const dt = ss.getSheetByName('Forecast detail') || ss.insertSheet('Forecast detail'); dt.clear();
  const det = ev.sort((a, b) => a.dB < b.dB ? -1 : 1).map(e => [e.dB, week(e.dB) + 1, e.item, e.type, e.aB, e.dP, week(e.dP) + 1, e.aP]);
  dt.getRange(1, 1, 1, 8).setValues([['Date (base)', 'Week', 'Item', 'Type', 'GEL (base)', 'Date (pess.)', 'Week (pess.)', 'GEL (pess.)']]).setFontWeight('bold');
  dt.getRange(2, 1, det.length, 8).setValues(det);

  // 7. თვითშემოწმება
  const sum = t => r2(ev.filter(e => e.type === t && week(e.dB) >= 0).reduce((s, e) => s + e.aB, 0));
  Logger.log('Check: approved payables ' + sum('Approved payables') + ' | receivables (base) ' + sum('Receivables') + ' | skipped 61+: ' + arSkipped + ' | USD rate ' + fx.USD + ' (' + fxDate.USD + ')');
  Logger.log('Lowest: base ' + minB[0] + ' (W' + minB[1] + ') | pessimistic ' + minP[0] + ' (W' + minP[1] + ') | buffer ' + buffer);
}
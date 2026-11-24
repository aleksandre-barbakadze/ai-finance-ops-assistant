// PaymentRun.gs – კვირის გადასახდელი სია მკაცრი ფილტრით. ფულს არ რიცხავს!
function buildPaymentRun() {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const r2 = x => Math.round(x * 100) / 100;
  const days = (a, b) => Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 864e5);
  const settings = {}; ss.getSheetByName('Settings').getDataRange().getValues().forEach(r => settings[r[0]] = r[1]);
  const runDate = toISO(settings['Payment run date'], tz);
  const horizon = Number(settings['Payment run horizon (days)']);
  if (!runDate || !horizon) throw new Error('Settings: "Payment run date" ან "Payment run horizon (days)" ცარიელია');
  const lastDue = Utilities.formatDate(new Date(new Date(runDate + 'T12:00:00').getTime() + horizon * 864e5), tz, 'yyyy-MM-dd');

  // 1. ცნობარის IBAN-ები
  const sup = ss.getSheetByName('Suppliers').getDataRange().getValues();
  const sh0 = sup[0].map(String), masterIban = {};
  sup.slice(1).forEach(r => { const id = String(r[sh0.indexOf('Supplier ID')]).trim(); if (id) masterIban[id] = String(r[sh0.indexOf('IBAN')]).trim(); });

  // 2. ინვოისები
  const v = ss.getSheetByName('Invoices').getDataRange().getValues();
  const h = v[0].map(String), g = (r, n) => r[h.indexOf(n)];
  const RISK = ['IBAN_CHANGED', 'UNKNOWN_SUPPLIER', 'DUPLICATE'];
  const pay = [], skip = [], result = {};
  v.slice(1).forEach(r => {
    const id = String(g(r, 'Invoice ID')).trim(); if (!id) return;
    if (g(r, 'Payment status') === 'FEE_DEDUCTED') return;
    const open = r2(Number(g(r, 'Total amount')) - (Number(g(r, 'Paid amount')) || 0));
    const decision = String(g(r, 'Payment decision') || '').trim().toUpperCase();
    if (open <= 0.009 && !decision) return;               // გადახდილია და გადაწყვეტილება არ აქვს – არ გვაინტერესებს

    const due = toISO(g(r, 'Due date'), tz);
    const flags = String(g(r, 'Check flags') || '').split(',').map(s => s.trim());
    const supId = String(g(r, 'Supplier ID')).trim();
    let reason = '';
    if (decision === 'BLOCKED') reason = 'BLOCKED';
    else if (decision === 'HOLD') reason = 'HOLD';
    else if (g(r, 'Status') !== 'Approved') reason = 'NOT_APPROVED';
    else if (flags.some(f => RISK.includes(f))) reason = 'RISK_FLAG: ' + flags.filter(f => RISK.includes(f)).join(', ');
    else if (open <= 0.009) reason = 'NOTHING_OPEN';
    else if (!masterIban[supId]) reason = 'NO_MASTER_IBAN';
    else if (!due) reason = 'NO_DUE_DATE';
    else if (due > lastDue) reason = 'NOT_DUE_THIS_RUN';

    const base = { id, sup: g(r, 'Supplier name'), open, cur: g(r, 'Currency'), due, note: g(r, 'Decision note') || '' };
    if (reason) { skip.push(Object.assign(base, { reason })); result[id] = { inRun: 'NO', amt: 0, reason }; }
    else {
      pay.push(Object.assign(base, { iban: masterIban[supId], late: Math.max(0, days(due, runDate)), ref: g(r, 'Invoice number'), by: g(r, 'Decided by') || '' }));
      result[id] = { inRun: 'YES', amt: open, reason: '' };
    }
  });
  pay.sort((a, b) => b.late - a.late || (a.due < b.due ? -1 : 1));

  // 3. ჩაწერა
  const sh = ss.getSheetByName('Payment run') || ss.insertSheet('Payment run');
  sh.clear();
  const totals = {}; pay.forEach(p => totals[p.cur] = r2((totals[p.cur] || 0) + p.open));
  const out = [];
  out.push(['PAYMENT RUN ' + runDate + ' (due ≤ ' + lastDue + ' + overdue)', 'Proposed: ' + pay.length + ' | ' + JSON.stringify(totals), '', '', '', '', '', '', '', '', '', '']);
  out.push(['Priority', 'Invoice ID', 'Supplier', 'Pay to IBAN (master data)', 'Amount', 'Currency', 'Due date', 'Days overdue', 'Reference', 'Approval 1 (accountant)', 'Approval 2 (owner)', 'Status']);
  pay.forEach((p, i) => out.push([i + 1, p.id, p.sup, p.iban, p.open, p.cur, p.due, p.late, p.ref, p.by, '', 'Proposed – awaiting 2nd approval']));
  out.push(['', '', '', '', '', '', '', '', '', '', '', '']);
  out.push(['NOT IN THIS RUN', '', '', '', '', '', '', '', '', '', '', '']);
  out.push(['', 'Invoice ID', 'Supplier', 'Reason', 'Open amount', 'Currency', 'Due date', 'Decision note', '', '', '', '']);
  skip.forEach(s => out.push(['', s.id, s.sup, s.reason, s.open, s.cur, s.due, s.note, '', '', '', '']));
  sh.getRange(1, 1, out.length, 12).setNumberFormat('@').setValues(out);
  sh.getRange(2, 1, 1, 12).setFontWeight('bold'); sh.getRange(pay.length + 4, 1, 3, 12).setFontWeight('bold');
  sh.getRange(1, 1).setFontWeight('bold');

  // 4. შედარება პასუხთან
  const key = ss.getSheetByName('Payment run key').getDataRange().getValues().slice(1).filter(r => r[0]);
  let ok = 0;
  key.forEach(k => {
    const id = String(k[0]).trim(), got = result[id] || { inRun: 'MISSING', amt: 0, reason: '' };
    const exp = String(k[1]).trim() + ' | ' + r2(Number(k[2])) + ' | ' + String(k[3]).trim();
    const res = got.inRun + ' | ' + r2(got.amt) + ' | ' + got.reason;
    if (exp === res) ok++; else Logger.log('❌ ' + id + ': expected ' + exp + '  |  got ' + res);
  });
  Object.keys(result).filter(id => !key.some(k => String(k[0]).trim() === id)).forEach(id => Logger.log('❓ not in key: ' + id + ' → ' + result[id].inRun + ' ' + result[id].reason));
  Logger.log('Payment run ' + runDate + ': ' + pay.length + ' payments | ' + JSON.stringify(totals));
  Logger.log('Payment run correct: ' + ok + ' / ' + key.length);
}
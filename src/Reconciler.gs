function reconcileBank() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const settings = getSettings();
  const end = toISO(settings['Data period end'], tz);
  const tol = Number(settings['Amount tolerance (GEL)']) || 1;
  const cardTol = (Number(settings['Card FX tolerance (%)']) || 3) / 100;
  const norm = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');   // "PM-1297" → "PM1297"
  const NON_INVOICE = ['Treasury of Georgia', 'Kartuli Bank'];

  const read = name => {
    const sh = ss.getSheetByName(name), v = sh.getDataRange().getValues();
    return { sh, h: v[0], rows: v.slice(1).map((r, i) => { const o = { _row: i + 2 }; v[0].forEach((k, j) => o[k] = r[j]); return o; }) };
  };

  // ცნობარი და კურსები
  const sup = {}, supByName = {};
  read('Suppliers').rows.filter(o => o['Supplier ID']).forEach(o => { sup[o['Supplier ID']] = o; supByName[o['Name']] = o; });
  const fx = {}; read('FX rates').rows.forEach(o => { if (o['Date']) fx[toISO(o['Date'], tz)] = o; });
  const fxDates = Object.keys(fx).sort();
  const rate = (cur, d) => { const k = fx[d] ? d : fxDates.filter(x => x <= d).pop(); return k ? Number(fx[k][cur]) : null; };

  // ინვოისები → ჯგუფები (მომწოდებელი + ნომერი). დუბლიკატები ერთ ჯგუფში ხვდება
  const inv = read('Invoices');
  const groups = {};
  inv.rows.filter(o => o['Invoice ID']).forEach(o => {
    const k = o['Supplier ID'] + '|' + norm(o['Invoice number']);
    if (!groups[k]) groups[k] = { sid: o['Supplier ID'], num: norm(o['Invoice number']), total: Number(o['Total amount']) || 0,
      cur: o['Currency'], date: toISO(o['Invoice date'], tz), due: toISO(o['Due date'], tz), rows: [], pays: [] };
    groups[k].rows.push(o);
  });
  const G = Object.values(groups);

  // ბანკის ტრანზაქციები
  const bank = read('Bank');
  const txns = bank.rows.filter(t => t['Txn ID']);
  txns.forEach(t => { t._date = toISO(t['Date'], tz); t._debit = Number(t['Debit']) || 0; t._credit = Number(t['Credit']) || 0; });
  const inGroupCur = (t, g) => {                                     // გადახდა ინვოისის ვალუტაში
    if (t['Currency'] === g.cur) return t._debit;
    if (t['Currency'] === 'GEL' && (g.cur === 'USD' || g.cur === 'EUR')) { const r = rate(g.cur, t._date); return r ? t._debit / r : null; }
    return null;
  };
  const daysApart = (a, b) => Math.abs((new Date(a) - new Date(b)) / 86400000);

  // 1. კლასიფიკაცია + დამთხვევა ინვოისის ნომრით
  txns.forEach(t => {
    if (t._credit > 0) { t.status = 'INCOMING'; return; }
    if (NON_INVOICE.includes(t['Counterparty']) || /^Payroll/.test(t['Counterparty'])) { t.status = 'NON_INVOICE'; return; }
    const s = supByName[t['Counterparty']];
    if (!s) { t.status = 'NO_INVOICE'; t.note = 'Counterparty not in supplier list'; return; }
    t.sid = s['Supplier ID'];
    const d = norm(t['Description']);
    const hit = G.filter(g => g.sid === t.sid && g.num && d.includes(g.num)).sort((a, b) => b.num.length - a.num.length)[0];
    if (hit) { hit.pays.push(t); t.group = hit; t.note = 'Matched by invoice number'; }
  });

  // 2. დამთხვევა მომწოდებლით + თანხით (როცა ნომერი არ წერია); უახლოესი ვადით
  txns.filter(t => t.sid && !t.group).forEach(t => {
    const cands = G.filter(g => g.sid === t.sid && !g.pays.length).filter(g => {
      const a = inGroupCur(t, g); if (a === null) return false;
      return Math.abs(a - g.total) <= (t['Currency'] === g.cur ? tol : g.total * cardTol);
    }).sort((a, b) => daysApart(a.due || a.date, t._date) - daysApart(b.due || b.date, t._date));
    if (cands.length) { cands[0].pays.push(t); t.group = cands[0]; t.note = 'Matched by supplier + amount (no invoice number in reference)'; }
    else { t.status = 'NO_INVOICE'; t.note = 'Supplier payment without matching invoice'; }
  });

  // 3. თითოეული ინვოისის სტატუსი
  G.forEach(g => {
    const s = sup[g.sid], terms = String(s ? s['Payment terms'] : '');
    const paid = g.pays.reduce((x, t) => x + (inGroupCur(t, g) || 0), 0);
    const lim = g.pays.some(t => t['Currency'] !== g.cur) ? g.total * cardTol : tol;
    let st, note = '';
    if (/deducted/i.test(terms)) st = 'FEE_DEDUCTED';
    else if (!g.pays.length) st = g.due ? (g.due <= end ? 'OVERDUE' : 'NOT_DUE') : 'UNPAID';
    else if (paid >= g.total * 1.9) st = 'DOUBLE_PAID';
    else if (paid > g.total + lim) st = 'OVERPAID';
    else if (Math.abs(paid - g.total) <= lim) st = 'PAID';
    else st = /30%/.test(terms) ? 'PARTIAL' : 'UNDERPAID';

    // ფული ცნობარისგან განსხვავებულ ანგარიშზე წავიდა?
    const master = normIban(s ? s['IBAN'] : '');
    const bad = g.pays.find(t => /^[A-Z]{2}\d/.test(master) && /^[A-Z]{2}\d/.test(normIban(t['Counterparty IBAN'])) && normIban(t['Counterparty IBAN']) !== master);
    if (bad) { st = 'PAID_TO_UNVERIFIED_IBAN'; note = 'Paid to ' + bad['Counterparty IBAN'] + ' – supplier master has ' + s['IBAN']; }
    if (!note && g.pays.length) note = g.pays.map(t => t['Txn ID'] + ' ' + t._date).join('; ');
        // გადაიხადეს, თუმცა ინვოისი ჯერ შესამოწმებელი იყო (Approved არ არის)
    if (st === 'PAID' && g.rows.every(o => o['Status'] === 'Needs review')) {
      const fl = [g.rows[0]['AI flags'], g.rows[0]['Check flags']].filter(Boolean).join(', ');
      st = 'PAID_BEFORE_REVIEW'; note = 'Paid while flagged for review (' + fl + ') – ' + note;
    }
    g.status = st; g.paid = Math.round(paid * 100) / 100; g.note = note;
        g.pays.forEach(t => { t.status = ['DOUBLE_PAID', 'PAID_TO_UNVERIFIED_IBAN', 'PAID_BEFORE_REVIEW'].includes(st) ? st : 'MATCHED'; });
  });

  // 4. ჩაწერა — Bank ფურცელი
  const bc = n => bank.h.indexOf(n) + 1;
  txns.forEach(t => bank.sh.getRange(t._row, bc('Match status'), 1, 3)
    .setValues([[t.status || '', t.group ? t.group.rows.map(r => r['Invoice ID']).join(', ') : '', t.note || '']]));

  // 5. ჩაწერა — Invoices ფურცელი (სვეტები თავისით ემატება)
  ['Payment status', 'Paid amount', 'Payment note'].forEach(n => {
    if (inv.h.indexOf(n) < 0) { inv.sh.getRange(1, inv.h.length + 1).setValue(n).setFontWeight('bold'); inv.h.push(n); }
  });
  const ic = n => inv.h.indexOf(n) + 1;
  G.forEach(g => g.rows.forEach(o => inv.sh.getRange(o._row, ic('Payment status'), 1, 3).setValues([[g.status, g.paid, g.note]])));

  // 6. შედარება Answer key-სთან
  const exp = {};
  read('Answer key').rows.forEach(o => { if (o['Invoice ID']) exp[o['Invoice ID']] = o['Expected payment status']; });
  let ok = 0, n = 0; const miss = [];
  G.forEach(g => g.rows.forEach(o => { const e = exp[o['Invoice ID']]; if (!e) return; n++; if (e === g.status) ok++; else miss.push(o['Invoice ID'] + ': expected ' + e + ', got ' + g.status); }));
  Logger.log('Payment status correct: ' + ok + ' / ' + n);
  miss.forEach(m => Logger.log('❌ ' + m));
  Logger.log('NO_INVOICE: ' + txns.filter(t => t.status === 'NO_INVOICE').map(t => t['Txn ID'] + ' ' + t['Counterparty']).join('; '));
}
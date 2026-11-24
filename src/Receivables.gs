// Receivables.gs – დებიტორები: კლიენტების ჩარიცხვები ↔ ჩვენი ინვოისები
const AR_OVERDUE_FLAG_DAYS = 30;   // ძველი ვალი ამაზე მეტით თუ არის ვადაგადაცილებული, ახალი ნისია ფლაგდება

function reconcileReceivables() {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const r2 = x => Math.round(x * 100) / 100;
  const days = (a, b) => Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 864e5);
  const table = name => {
    const v = ss.getSheetByName(name).getDataRange().getValues();
    const h = v[0].map(String);
    return v.slice(1).filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
  };

  // 1. პარამეტრები
  const settings = {};
  ss.getSheetByName('Settings').getDataRange().getValues().forEach(r => settings[r[0]] = r[1]);
  const asOf = toISO(settings['AR as-of date'], tz);
  if (!asOf) throw new Error('Settings: "AR as-of date" ცარიელია');

  // 2. კლიენტები (IBAN-ით ვცნობთ)
  const cust = {}, byIban = {};
  table('Customers').forEach(c => { cust[c['Customer ID']] = c; byIban[normIban(c['IBAN'])] = c['Customer ID']; });

  // 3. ჩვენი ინვოისები
  const invSheet = ss.getSheetByName('Sales invoices');
  const inv = table('Sales invoices').map(o => ({
    no: String(o['Invoice No']).trim(), cust: o['Customer ID'],
    issue: toISO(o['Issue date'], tz), due: toISO(o['Due date'], tz),
    total: Number(o['Total (GEL)']), paid: 0, pays: []
  }));
  const open = i => r2(i.total - i.paid);

  // 4. ბანკის შემომავალი ჩანაწერები კლიენტებისგან (A=TX, B=თარიღი, D=დანიშნულება, F=IBAN, H=კრედიტი)
  const bankSheet = ss.getSheetByName('Bank');
  const bank = bankSheet.getDataRange().getValues();
  const arCol = new Array(bank.length - 1).fill('');
  const lines = [];
  bank.slice(1).forEach((r, k) => {
    const credit = Number(r[7]) || 0, date = toISO(r[1], tz), cid = byIban[normIban(r[5])];
    if (credit > 0 && cid && date <= asOf) lines.push({ k, date, cid, amt: credit, desc: String(r[3]), tx: r[0] });
  });
  lines.sort((a, b) => a.date < b.date ? -1 : 1);

  // 5. თანხის განაწილება
  const apply = (line, list) => {
    let left = line.amt; const used = [];
    list.forEach(i => {
      const x = Math.min(left, open(i));
      if (x <= 0.009) return;
      i.paid = r2(i.paid + x); i.pays.push({ date: line.date, amt: x }); left = r2(left - x); used.push(i.no);
    });
    return { used, left };
  };
  lines.forEach(line => {
    const mine = inv.filter(i => i.cust === line.cid && i.issue <= line.date && open(i) > 0.009)
                    .sort((a, b) => a.due < b.due ? -1 : 1);
    const refs = (line.desc.match(/WW-S-\d+/gi) || []).map(s => s.toUpperCase());
    const refInv = mine.filter(i => refs.includes(i.no));
    let res, how;
    if (refInv.length) { res = apply(line, refInv); how = 'by reference'; }
    else {
      const exact = mine.find(i => Math.abs(open(i) - line.amt) < 0.01);
      if (exact) { res = apply(line, [exact]); how = 'no reference – matched by amount'; }
      else { res = apply(line, mine); how = 'no reference – oldest first'; }
    }
    arCol[line.k] = (res.used.join(', ') || 'UNAPPLIED') + ' (' + how + ')' + (res.left > 0.009 ? ' | unapplied ' + res.left : '');
    if (res.left > 0.009) Logger.log('⚠️ ' + line.tx + ': unapplied ' + res.left + ' GEL');
  });

  // 6. სტატუსი, დღეები, aging
  inv.forEach(i => {
    i.out = open(i);
    i.last = i.pays.length ? i.pays[i.pays.length - 1].date : '';
    if (i.out <= 0.009) { i.out = 0; i.days = Math.max(0, days(i.due, i.last)); i.status = i.days > 0 ? 'PAID_LATE' : 'PAID'; }
    else { i.days = Math.max(0, days(i.due, asOf)); i.status = i.paid > 0 ? 'PARTIAL' : (i.days > 0 ? 'OVERDUE' : 'NOT_DUE'); }
    i.bucket = i.out === 0 ? '' : i.days === 0 ? 'Not due' : i.days <= 30 ? '1-30' : i.days <= 60 ? '31-60' : '61+';
  });

  // 7. კონტროლი: როგორი იყო მდგომარეობა გაყიდვის დღეს
  inv.forEach(x => {
    const c = cust[x.cust], flags = [];
    let exposure = x.total, worst = 0;
    inv.filter(i => i.cust === x.cust && i !== x && i.issue < x.issue).forEach(i => {
      const openThen = r2(i.total - i.pays.filter(p => p.date <= x.issue).reduce((s, p) => s + p.amt, 0));
      if (openThen > 0.009) { exposure += openThen; worst = Math.max(worst, days(i.due, x.issue)); }
    });
    if (worst > AR_OVERDUE_FLAG_DAYS) flags.push('SOLD_WHILE_OVERDUE');
    if (c && exposure > Number(c['Credit limit (GEL)'])) flags.push('OVER_CREDIT_LIMIT');
    x.flags = flags.sort().join(', ');
    const n = [];
    if (x.pays.length > 1) n.push(x.pays.length + ' payments');
    if (x.status === 'PARTIAL') n.push('paid ' + x.paid + ' of ' + x.total);
    if (flags.includes('OVER_CREDIT_LIMIT')) n.push('exposure ' + r2(exposure) + ' > limit ' + c['Credit limit (GEL)']);
    if (flags.includes('SOLD_WHILE_OVERDUE')) n.push('older invoice ' + worst + ' days overdue at sale');
    x.note = n.join('; ');
  });

  // 8. ჩაწერა
  const head = ['Paid (GEL)', 'Outstanding (GEL)', 'Last payment', 'AR status', 'Days late / overdue', 'Aging bucket', 'Control flags', 'AR note'];
  invSheet.getRange(1, 9, 1, head.length).setValues([head]).setFontWeight('bold');
  invSheet.getRange(2, 14, inv.length, 1).setNumberFormat('@');   // Aging bucket – ტექსტად, რომ 1-30 თარიღად არ იქცეს
  invSheet.getRange(2, 9, inv.length, head.length).setValues(inv.map(i => [i.paid, i.out, i.last, i.status, i.days, i.bucket, i.flags, i.note]));
  bankSheet.getRange(1, 13).setValue('AR match').setFontWeight('bold');
  if (arCol.length) bankSheet.getRange(2, 13, arCol.length, 1).setValues(arCol.map(v => [v]));

  // 9. შედარება პასუხებთან + ჯამები
  const key = {}; table('AR answer key').forEach(k => key[String(k['Invoice No']).trim()] = k);
  const normF = s => String(s || '').split(',').map(t => t.trim()).filter(Boolean).sort().join(', ');
  let ok = 0;
  inv.forEach(i => {
    const k = key[i.no]; if (!k) { Logger.log('❓ ' + i.no + ' not in AR answer key'); return; }
    const exp = [k['Expected status'], Number(k['Days late / overdue']), r2(Number(k['Expected outstanding (GEL)'])),
                 String(k['Expected aging bucket']).trim(), normF(k['Expected control flags'])].join(' | ');
    const got = [i.status, i.days, i.out, i.bucket, i.flags].join(' | ');
    if (exp === got) ok++; else Logger.log('❌ ' + i.no + ': expected ' + exp + '  |  got ' + got);
  });
  Logger.log('AR correct: ' + ok + ' / ' + inv.length);
  const sumB = {}; inv.forEach(i => { if (i.out) sumB[i.bucket] = r2((sumB[i.bucket] || 0) + i.out); });
  Logger.log('Outstanding ' + asOf + ': ' + r2(inv.reduce((s, i) => s + i.out, 0)) + ' GEL | ' + JSON.stringify(sumB));
}
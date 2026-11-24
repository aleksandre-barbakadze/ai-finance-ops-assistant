function generateBankStatement() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const settings = getSettings();
  const start = toISO(settings['Data period start'], tz);
  const end = toISO(settings['Data period end'], tz);

  // ცხრილებს ობიექტებად ვკითხულობთ
  const objs = name => {
    const v = ss.getSheetByName(name).getDataRange().getValues();
    return v.slice(1).filter(r => r[0]).map(r => { const o = {}; v[0].forEach((h, i) => o[h] = r[i]); return o; });
  };
  const key = objs('Answer key');                                   // „სიმართლე“ — რეალურად მოსული ინვოისები
  const src = {}; objs('Invoice source').forEach(o => src[o['Invoice ID']] = o);
  const sup = {}; objs('Suppliers').forEach(o => sup[o['Supplier ID']] = o);
  const fx = {};  objs('FX rates').forEach(o => fx[toISO(o['Date'], tz)] = o);

  const addDays = (iso, n) => { const d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() + n); return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
  const round2 = x => Math.round(x * 100) / 100;
  const rows = [];
  const pay = (date, cur, desc, cp, iban, debit, credit) => {
    if (date >= start && date <= end) rows.push([date, cur, desc, cp, iban, debit || '', credit || '']);
  };

  // ===== 1. ინვოისების გადახდები =====
  const HOLD = ['INV-012', 'INV-014', 'INV-017', 'INV-027', 'INV-031'];   // არ გადაიხადეს
  key.forEach(k => {
    const id = k['Invoice ID'], s = sup[k['True supplier ID']], x = src[id] || {};
    if (!s || HOLD.includes(id) || s['Supplier ID'] === 'S06') return;  // PayKart — საკომისიო უკვე დაკავებულია
    const cur = k['True currency'], total = Number(k['True total']), num = String(k['True invoice number']);
    const date = toISO(k['True invoice date'], tz), due = toISO(k['True due date'], tz);
    const name = s['Name'], iban = x['IBAN'] || s['IBAN'];              // IBAN — ინვოისიდან (როგორც ბუღალტერი იხდის)
    let desc = 'Payment for invoice ' + num;

    if (s['Supplier ID'] === 'S07') {                                    // რეკლამა — ბარათით, ლარში
      const r = fx[date] ? Number(fx[date]['USD']) : 2.6;
      pay(date, 'GEL', 'Card purchase ADREACH LTD DUBLIN', name, 'Visa ••••0001', round2(total * r * 1.015));
      return;
    }
    if (['S01', 'S02'].includes(s['Supplier ID'])) {                    // ჩინეთი — 30% / 70%
      pay(addDays(date, 2), cur, desc + ' – 30% advance', name, iban, round2(total * 0.3));
      pay(addDays(date, 2), cur, 'SWIFT transfer fee', 'Kartuli Bank', '', 25);
      if (id !== 'INV-006') {                                            // ხაფანგი #3: INV-006 — მხოლოდ 30%
        pay(addDays(date, 20), cur, desc + ' – 70% balance', name, iban, round2(total * 0.7));
        pay(addDays(date, 20), cur, 'SWIFT transfer fee', 'Kartuli Bank', '', 25);
      }
      return;
    }
    if (!due) return;
    let payDate = due, amount = total;
    if (id === 'INV-016') amount = total - 1000;                         // ხაფანგი #4: არასწორი თანხა
    if (id === 'INV-022') payDate = addDays(due, 3);                     // ხაფანგი #1: დუბლიკატიც გადაიხადეს
    if (id === 'INV-025' || id === 'INV-026') desc = 'internet';         // ხაფანგი #7: ნომრის გარეშე
    pay(payDate, cur, desc, name, iban, amount);
    if (cur !== 'GEL') pay(payDate, cur, 'SWIFT transfer fee', 'Kartuli Bank', '', 20);
  });

  // ===== 2. ხარჯები ინვოისის გარეშე (ნორმალური) =====
  ['2026-08', '2026-09', '2026-10'].forEach(m => {
    pay(m + '-15', 'GEL', 'Personal income tax (20%) – previous month payroll', 'Treasury of Georgia', '', 4875);
    pay(m + '-15', 'GEL', 'VAT payment – previous month', 'Treasury of Georgia', '', 6500);
    pay(m + '-25', 'GEL', 'Salaries ' + m, 'Payroll – 9 employees', '', 19500);
    pay(m + '-28', 'GEL', 'Monthly account maintenance fee', 'Kartuli Bank', '', 15);
  });
  pay('2026-09-18', 'GEL', 'Service payment', 'შპს „ალფა სერვისი“', 'GE00XX0000000000000123', 1200);  // ხაფანგი #6

    // ===== 3. შემომავალი ჩარიცხვები =====
  const card = { '2026-08': 78000, '2026-09': 90000, '2026-10': 102000 };   // ბარათით გაყიდვები (PayKart-ის ამონაწერის ბრუნვა)
  const cod  = { '2026-08': 52000, '2026-09': 58000, '2026-10': 64000 };    // ნაღდი კურიერთან
  // რამდენი ორშაბათი / პარასკევია თითოეულ თვეში (აგვისტოს 5 ორშაბათი აქვს!)
  const countDays = wdTarget => {
    const c = {};
    for (let d = start; d <= end; d = addDays(d, 1))
      if (new Date(d + 'T12:00:00').getDay() === wdTarget) c[d.slice(0, 7)] = (c[d.slice(0, 7)] || 0) + 1;
    return c;
  };
  const mondays = countDays(1), fridays = countDays(5);
  const sums = {};
  for (let d = start; d <= end; d = addDays(d, 1)) {
    const wd = new Date(d + 'T12:00:00').getDay(), m = d.slice(0, 7);
    if (wd === 1) {
      const amt = round2(card[m] / mondays[m] * 0.98);
      pay(d, 'GEL', 'PayKart settlement (card sales, net of 2% fee)', 'შპს „პეიკარტი“', 'GE00XX0000000000000006', '', amt);
      sums[m] = (sums[m] || 0) + amt;
    }
    if (wd === 5) pay(d, 'GEL', 'SwiftGo cash-on-delivery remittance', 'შპს „სვიფტგო“', 'GE00XX0000000000000005', '', round2(cod[m] / fridays[m]));
  }
  Object.keys(sums).forEach(m => Logger.log('PayKart ' + m + ': payouts ' + round2(sums[m]) + ' | expected ' + round2(card[m] * 0.98) + ' (' + mondays[m] + ' Mondays)'));

    // ===== 3ბ. კლიენტების ჩარიცხვები (Receipts source) =====
  {
    const tzR = Session.getScriptTimeZone();
    const book = SpreadsheetApp.getActive();
    const custName = {};
    book.getSheetByName('Customers').getDataRange().getValues().slice(1)
      .forEach(r => { if (r[0]) custName[r[0]] = r[1]; });
    const rec = book.getSheetByName('Receipts source').getDataRange().getValues();
    const c = name => rec[0].indexOf(name);
    let n = 0;
    rec.slice(1).forEach(r => {
      if (!r[c('Receipt ID')]) return;
      pay(toISO(r[c('Date')], tzR), 'GEL', String(r[c('Bank reference')]),
          custName[r[c('Customer ID')]] || '', String(r[c('Payer IBAN')]),
          '', Number(r[c('Amount (GEL)')]));
      n++;
    });
    Logger.log('Customer receipts added: ' + n);
  }

  // ===== 4. ჩაწერა თარიღის მიხედვით =====
  rows.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);
  const out = rows.map((r, i) => ['TX-' + String(i + 1).padStart(4, '0'), r[0], r[1] + ' account', r[2], r[3], r[4], r[5], r[6], r[1]]);
  const sh = ss.getSheetByName('Bank');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  sh.getRange(2, 1, out.length, 9).setValues(out);
  Logger.log('Bank statement: ' + out.length + ' transactions');
}
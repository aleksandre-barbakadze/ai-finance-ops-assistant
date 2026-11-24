// PayKartCheck.gs – სამმხრივი შემოწმება: მაღაზია ↔ ბანკი ↔ PayKart-ის ამონაწერი
function reconcilePayKart() {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const r2 = x => Math.round(x * 100) / 100;
  const TOL = 1;   // ₾ – დამრგვალების დაშვება
  const settings = {}; ss.getSheetByName('Settings').getDataRange().getValues().forEach(r => settings[r[0]] = r[1]);
  const feePct = Number(settings['PayKart fee (%)']);
  if (!feePct) throw new Error('Settings: "PayKart fee (%)" ცარიელია');

  // 1. PayKart-ის IBAN Suppliers-იდან (S06)
  const sup = ss.getSheetByName('Suppliers').getDataRange().getValues();
  const sh0 = sup[0].map(String);
  const s06 = sup.find(r => String(r[sh0.indexOf('Supplier ID')]).trim() === 'S06');
  if (!s06) throw new Error('Suppliers: S06 ვერ მოიძებნა');
  const pkIban = normIban(s06[sh0.indexOf('IBAN')]);

  // 2. ჩარიცხვები ბანკიდან, თვეების მიხედვით (B=თარიღი, F=IBAN, H=კრედიტი)
  const payouts = {}, count = {};
  ss.getSheetByName('Bank').getDataRange().getValues().slice(1).forEach(r => {
    const credit = Number(r[7]) || 0;
    if (credit > 0 && normIban(r[5]) === pkIban) {
      const m = toISO(r[1], tz).slice(0, 7);
      payouts[m] = r2((payouts[m] || 0) + credit); count[m] = (count[m] || 0) + 1;
    }
  });

  // 3. საკომისიო PayKart-ის ამონაწერებიდან (Invoices, S06)
  const inv = ss.getSheetByName('Invoices').getDataRange().getValues();
  const ih = inv[0].map(String), col = n => ih.indexOf(n);
  const fees = {}, feeDocs = {};
  inv.slice(1).forEach(r => {
    if (String(r[col('Supplier ID')]).trim() !== 'S06') return;
    const m = toISO(r[col('Invoice date')], tz).slice(0, 7);
    fees[m] = r2((fees[m] || 0) + Number(r[col('Total amount')]));
    feeDocs[m] = (feeDocs[m] ? feeDocs[m] + ', ' : '') + r[col('Invoice ID')];
  });

  // 4. შედარება მაღაზიის მონაცემებთან
  const sh = ss.getSheetByName('Card sales (shop)');
  const data = sh.getDataRange().getValues();
  const h = data[0].map(String), c = n => h.indexOf(n);
  const out = [], logs = [];
  let ok = 0, n = 0;
  data.slice(1).forEach(r => {
    const m = String(r[c('Month')]).trim(); if (!m) { out.push(['', '', '', '', '', '', '']); return; }
    const shop = Number(r[c('Card sales (GEL)')]);
    const pay = payouts[m] || 0, fee = fees[m] || 0, settled = r2(pay + fee);
    const diff = r2(shop - settled);
    const rate = settled ? r2(fee / settled * 100) : 0;
    const st = [], note = [];
    if (Math.abs(diff) > TOL) {
      st.push('MISMATCH');
      note.push(diff > 0 ? 'Shop sales exceed PayKart settlement by ' + diff + ' – check refunds / chargebacks / missing payout'
                         : 'PayKart settled ' + (-diff) + ' more than shop sales – check shop export');
    }
    if (!fee) { st.push('NO_FEE_STATEMENT'); note.push('No PayKart statement for this month'); }
    else if (Math.abs(rate - feePct) > 0.01) { st.push('FEE_RATE_MISMATCH'); note.push('Fee ' + rate + '% vs contract ' + feePct + '%'); }
    const status = st.length ? st.join(', ') : 'OK';
    note.unshift((count[m] || 0) + ' payouts; statement ' + (feeDocs[m] || '–'));
    out.push([pay, fee, settled, diff, rate, status, note.join('; ')]);

    // შედარება პასუხთან
    n++;
    const exp = String(r[c('Expected status')]).trim() + ' | ' + r2(Number(r[c('Expected difference (GEL)')]));
    const got = status + ' | ' + diff;
    if (exp === got) ok++; else logs.push('❌ ' + m + ': expected ' + exp + '  |  got ' + got);
    Logger.log('PayKart ' + m + ': shop ' + shop + ' | payouts ' + pay + ' + fee ' + fee + ' = ' + settled + ' | diff ' + diff + ' | fee ' + rate + '% → ' + status);
  });

  // 5. ჩაწერა F–L სვეტებში
  const head = ['Payouts (bank)', 'Fee (statement)', 'Settled total', 'Difference', 'Fee rate %', 'Status', 'Note'];
  sh.getRange(1, 6, 1, head.length).setValues([head]).setFontWeight('bold');
  if (out.length) sh.getRange(2, 6, out.length, head.length).setValues(out);
  logs.forEach(l => Logger.log(l));
  Logger.log('PayKart check correct: ' + ok + ' / ' + n);
}
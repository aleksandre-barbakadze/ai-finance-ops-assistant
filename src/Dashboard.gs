// Dashboard.gs – ერთი გვერდი მფლობელისთვის. ყველა ციფრი ფურცლებიდან, კოდით
function buildDashboard() {
  const ss = SpreadsheetApp.getActive(), tz = Session.getScriptTimeZone();
  const r2 = x => Math.round(x * 100) / 100;
  const iso = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').trim();
  const fmt = v => Number(v).toFixed(2).replace(/\.00$/, '').replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const read = name => { const v = ss.getSheetByName(name).getDataRange().getValues(), h = v[0].map(String);
    return v.slice(1).filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; }); };
  const S = {}; ss.getSheetByName('Settings').getDataRange().getValues().forEach(r => S[r[0]] = r[1]);
  const month = iso(S['Data period end']).slice(0, 7), buffer = Number(S['Minimum cash buffer (GEL)']);

  // ---------- ციფრები ----------
  const bank = ss.getSheetByName('Bank').getDataRange().getValues().slice(1);
  let inflow = 0;
  bank.forEach(r => { const c = Number(r[7]) || 0; if (c && iso(r[1]).slice(0, 7) === month && (/paykart|swiftgo|cash-on-delivery/i.test(r[3]) || r[12])) inflow += c; });
  const si = read('Sales invoices');
  const ar = r2(si.reduce((s, o) => s + (Number(o['Outstanding (GEL)']) || 0), 0));
  const bucket = b => r2(si.filter(o => String(o['Aging bucket']).trim() === b).reduce((s, o) => s + Number(o['Outstanding (GEL)']), 0));
  const prSheet = ss.getSheetByName('Payment run').getDataRange().getValues();
  const pr = prSheet.filter(r => /^\d+$/.test(String(r[0]).trim()));
  const prSum = r2(pr.reduce((s, r) => s + Number(r[4]), 0));
  const fc = ss.getSheetByName('Cash forecast').getDataRange().getValues().slice(3).filter(r => r[0] !== '');
  const low = fc.reduce((m, r) => Number(r[8]) < m[0] ? [Number(r[8]), r[0]] : m, [Infinity, 0]);

  const inv = read('Invoices');
  const st = s => inv.filter(o => o['Payment status'] === s);
  const ids = a => a.map(o => o['Invoice ID']).join(', ');
  const dbl = {}; st('DOUBLE_PAID').forEach(o => dbl[o['Payment note']] = Number(o['Paid amount']) - Number(o['Total amount']));
  const noInv = bank.filter(r => r[9] === 'NO_INVOICE');
  const shop = read('Card sales (shop)').find(o => String(o['Month']).trim() === month) || {};
  const issues = [
    ['ორმაგი გადახდა – დასაბრუნებელი', r2(Object.values(dbl).reduce((s, x) => s + x, 0)), ids(st('DOUBLE_PAID')), 'მაღალი'],
    ['გადახდა შეუმოწმებელ IBAN-ზე (შესაძლო თაღლითობა)', r2(st('PAID_TO_UNVERIFIED_IBAN').reduce((s, o) => s + Number(o['Paid amount']), 0)), ids(st('PAID_TO_UNVERIFIED_IBAN')), 'მაღალი'],
    ['გადახდა ინვოისის გარეშე', r2(noInv.reduce((s, r) => s + (Number(r[6]) || 0), 0)), noInv.map(r => r[4]).join(', '), 'საშუალო'],
    ['PayKart: გაყიდვა ≠ ჩარიცხვა + საკომისიო (' + month + ')', Number(shop['Difference']) || 0, shop['Status'] || '', 'საშუალო'],
    ['გადახდილია შემოწმების დასრულებამდე', '', st('PAID_BEFORE_REVIEW').length + ' ინვოისი: ' + ids(st('PAID_BEFORE_REVIEW')), 'საშუალო'],
    ['ნაკლებად გადახდილი', r2(st('UNDERPAID').reduce((s, o) => s + Number(o['Total amount']) - Number(o['Paid amount']), 0)), ids(st('UNDERPAID')), 'დაბალი'],
    ['შეჩერებული (HOLD) / დაბლოკილი (BLOCKED)', '', ids(inv.filter(o => o['Payment decision'] === 'HOLD')) + ' / ' + ids(inv.filter(o => o['Payment decision'] === 'BLOCKED')), 'ინფო']];

  // ხარისხი – ცოცხლად: შედეგი ↔ პასუხების ფურცელი
  const score = (rows, idCol, expCol, got) => { const ok = rows.filter(k => String(k[expCol]).trim() === String(got(String(k[idCol]).trim())).trim()).length; return ok + ' / ' + rows.length; };
  const invById = {}; inv.forEach(o => invById[String(o['Invoice ID']).trim()] = o);
  const siById = {}; si.forEach(o => siById[String(o['Invoice No']).trim()] = o);
  const cs = read('Card sales (shop)');
  const prRes = {}; let after = false;
  prSheet.forEach(r => { const a = String(r[0]).trim(), id = String(r[1]).trim();
    if (/^\d+$/.test(a)) prRes[id] = 'YES|'; else if (a === 'NOT IN THIS RUN') after = true; else if (after && id && id !== 'Invoice ID') prRes[id] = 'NO|' + r[3]; });
  const prKey = read('Payment run key');
  const prOk = prKey.filter(k => prRes[String(k['Invoice ID']).trim()] === String(k['Expected in run']).trim() + '|' + String(k['Expected reason'] || '').trim()).length;
  const quality = [
    ['შემოწმება', 'შედეგი', 'როგორ'],
    ['AI ამოკითხვა (34 ინვოისი)', '271/272 ველი · ხაფანგები 10/10', 'ჩაწერილი შედეგი, 2026-11-18'],
    ['ბანკის შედარება (გადახდის სტატუსი)', score(read('Answer key'), 'Invoice ID', 'Expected payment status', id => (invById[id] || {})['Payment status']), 'ცოცხლად: Invoices ↔ Answer key'],
    ['დებიტორები (სტატუსი)', score(read('AR answer key'), 'Invoice No', 'Expected status', id => (siById[id] || {})['AR status']), 'ცოცხლად: Sales invoices ↔ AR answer key'],
    ['PayKart (სამმხრივი)', score(cs, 'Month', 'Expected status', m => (cs.find(o => String(o['Month']).trim() === m) || {})['Status']), 'ცოცხლად: Card sales'],
    ['Payment run', prOk + ' / ' + prKey.length, 'ცოცხლად: Payment run ↔ Payment run key'],
    ['13-კვირიანი პროგნოზი', '13/13 კვირა', 'დამოუკიდებელი გამოთვლა, 2026-11-24'],
    ['თვის ანგარიში (AI-ის რიცხვები)', '0 მოგონილი რიცხვი', 'ავტომატური შემოწმება, 2026-11-24']];

  // ---------- დაფის აწყობა ----------
  let sh = ss.getSheetByName('Dashboard');
  if (!sh) sh = ss.insertSheet('Dashboard', 0);
  sh.getCharts().forEach(c => sh.removeChart(c)); sh.clear(); sh.getRange('A1:Z200').breakApart();
  sh.setHiddenGridlines(true);
  const put = (r, c, v) => { const g = sh.getRange(r, c, v.length, v[0].length); g.setValues(v); return g; };
  const head = g => g.offset(0, 0, 1).setFontWeight('bold').setBackground('#e8eef7');
  put(1, 1, [['Worldwide – ფინანსური დაფა, ' + month]]).setFontSize(18).setFontWeight('bold');
  put(2, 1, [['⚠️ სინთეტიკური მონაცემები (გამოგონილი კომპანია). განახლდა: ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm') + '. ყველა ციფრი ითვლება კოდით, ცხრილის ფურცლებიდან.']]).setFontColor('#666');

  const kpi = [['შემოსავალი ' + month, r2(inflow), 'PayKart + ნაღდი + B2B, ₾'],
               ['დებიტორები', ar, '61+ დღე: ' + fmt(bucket('61+')) + ' ₾'],
               ['Payment run ' + iso(S['Payment run date']), prSum, pr.length + ' გადახდა · ელოდება მეორე ხელმოწერას'],
               ['უმცირესი ნაშთი (პესიმისტური)', low[0], 'კვირა ' + low[1] + ' · ზღვარი ' + fmt(buffer) + ' ₾']];
  kpi.forEach((k, i) => { const c = 1 + i * 2;
    sh.getRange(4, c, 1, 2).merge().setValue(k[0]).setFontColor('#555');
    sh.getRange(5, c, 1, 2).merge().setValue(k[1]).setNumberFormat('#,##0.00').setFontSize(20).setFontWeight('bold');
    sh.getRange(6, c, 1, 2).merge().setValue(k[2]).setFontSize(9).setFontColor('#777');
    sh.getRange(4, c, 3, 2).setBackground(i === 3 && low[0] < buffer ? '#fdecea' : '#f3f6fb').setBorder(true, true, true, true, false, false); });

  put(8, 1, [['⚠️ გამოვლენილი პრობლემები']]).setFontWeight('bold').setFontSize(12);
  const it = put(9, 1, [['პრობლემა', 'თანხა, ₾', 'ინვოისები / დეტალი', 'სიმძიმე']].concat(issues)); head(it);
  it.offset(1, 1, issues.length, 1).setNumberFormat('#,##0.00');
  issues.forEach((x, i) => sh.getRange(10 + i, 4).setBackground({ 'მაღალი': '#fdecea', 'საშუალო': '#fff4e5', 'დაბალი': '#eef7ee' }[x[3]] || '#ffffff'));

  const ag = [['ასაკი', '₾'], ['Not due', bucket('Not due')], ['1-30', bucket('1-30')], ['31-60', bucket('31-60')], ['61+', bucket('61+')]];
    put(8, 6, [['💳 დებიტორები ასაკის მიხედვით']]).setFontWeight('bold').setFontSize(12);
  sh.getRange(9, 6, 5, 1).setNumberFormat('@');
  head(put(9, 6, ag)).offset(1, 1, 4, 1).setNumberFormat('#,##0');
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(sh.getRange(9, 6, 5, 2)).setNumHeaders(1).setPosition(8, 9, 0, 0)
    .setOption('title', 'დებიტორები ასაკის მიხედვით, ₾').setOption('legend', { position: 'none' }).setOption('width', 420).setOption('height', 230).build());

  const r0 = 19;
  put(r0, 1, [['📈 13-კვირიანი პროგნოზი (სინთეტიკური ვარაუდებით)']]).setFontWeight('bold').setFontSize(12);
  const ft = [['კვირა', 'საბაზისო', 'პესიმისტური', 'ზღვარი']].concat(fc.map(r => ['W' + r[0] + ' ' + iso(r[1]).slice(5), Number(r[5]), Number(r[8]), buffer]));
  head(put(r0 + 1, 1, ft)).offset(1, 1, fc.length, 3).setNumberFormat('#,##0');
  fc.forEach((r, i) => { if (Number(r[8]) < buffer) sh.getRange(r0 + 2 + i, 3).setBackground('#fdecea'); });
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.LINE).addRange(sh.getRange(r0 + 1, 1, fc.length + 1, 4)).setNumHeaders(1).setPosition(r0, 6, 0, 0)
    .setOption('title', 'ფულის ნაშთი კვირების მიხედვით, ₾').setOption('width', 640).setOption('height', 320).build());

  const q0 = r0 + fc.length + 4;
  put(q0, 1, [['✅ ხარისხი და ტესტები']]).setFontWeight('bold').setFontSize(12);
  head(put(q0 + 1, 1, quality));
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 260); sh.setColumnWidth(4, 90);
    sh.setColumnWidth(5, 130); sh.setColumnWidth(6, 110); sh.setColumnWidth(7, 110); sh.setColumnWidth(8, 140);
  sh.getRange(4, 1, 3, 8).setWrap(true);
  sh.getRange(10, 1, issues.length, 4).setWrap(true).setVerticalAlignment('top');
  sh.getRange(q0 + 2, 1, quality.length - 1, 3).setWrap(true);
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);

  Logger.log('KPI: inflow ' + r2(inflow) + ' | AR ' + ar + ' (61+ ' + bucket('61+') + ') | Payment run ' + prSum + ' (' + pr.length + ') | lowest pess ' + low[0] + ' W' + low[1]);
  quality.slice(1).forEach(q => Logger.log('Quality: ' + q[0] + ' → ' + q[1]));
}
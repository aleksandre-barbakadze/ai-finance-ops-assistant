// MonthEnd.gs – თვის დახურვის ანგარიში: ციფრები – კოდი, ახსნა – AI, რიცხვების შემოწმება – კოდი
function buildMonthEndReport() {
  const ss = SpreadsheetApp.getActive(), tz = Session.getScriptTimeZone();
  const r2 = x => Math.round(x * 100) / 100;
  const iso = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').trim();
  const fmtN = v => typeof v === 'number' ? v.toFixed(2).replace(/\.00$/, '').replace(/\B(?=(\d{3})+(?!\d))/g, ',') : v;
  const read = name => { const v = ss.getSheetByName(name).getDataRange().getValues(), h = v[0].map(String);
    return v.slice(1).filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; }); };
  const S = {}; ss.getSheetByName('Settings').getDataRange().getValues().forEach(r => S[r[0]] = r[1]);
  const month = iso(S['Data period end']).slice(0, 7), asOf = iso(S['AR as-of date']);
  const F = [], add = (sec, label, value) => F.push([sec, label, fmtN(value)]);

  // 1. შემოსავალი თვეში (Bank: B=თარიღი, D=დანიშნულება, H=კრედიტი, M=AR match)
  const bank = ss.getSheetByName('Bank').getDataRange().getValues().slice(1);
  let pk = 0, cod = 0, cust = 0;
  bank.forEach(r => { const c = Number(r[7]) || 0; if (!c || iso(r[1]).slice(0, 7) !== month) return;
    if (/paykart/i.test(r[3])) pk += c; else if (/swiftgo|cash-on-delivery/i.test(r[3])) cod += c; else if (r[12]) cust += c; });
  const IN = 'შემოსავალი ' + month;
  add(IN, 'PayKart-ის ჩარიცხვები (ბარათით გაყიდვა, საკომისიოს შემდეგ), ₾', r2(pk));
  add(IN, 'კურიერის ნაღდი ფული, ₾', r2(cod));
  add(IN, 'B2B კლიენტების გადახდები, ₾', r2(cust));

  // 2. გამოვლენილი პრობლემები
  const P = 'პრობლემები', inv = read('Invoices');
  const st = s => inv.filter(o => o['Payment status'] === s);
  const ids = arr => arr.map(o => o['Invoice ID']).join(', ');
  const dbl = {}; st('DOUBLE_PAID').forEach(o => dbl[o['Payment note']] = r2(Number(o['Paid amount']) - Number(o['Total amount'])));
  add(P, 'ორმაგად გადახდილი, დასაბრუნებელი, ₾ (' + ids(st('DOUBLE_PAID')) + ')', r2(Object.values(dbl).reduce((s, x) => s + x, 0)));
  add(P, 'გადახდილია შეუმოწმებელ IBAN-ზე (შესაძლო თაღლითობა), ₾ (' + ids(st('PAID_TO_UNVERIFIED_IBAN')) + ')',
      r2(st('PAID_TO_UNVERIFIED_IBAN').reduce((s, o) => s + Number(o['Paid amount']), 0)));
  add(P, 'გადახდილია შემოწმების დასრულებამდე, ინვოისი (' + ids(st('PAID_BEFORE_REVIEW')) + ')', st('PAID_BEFORE_REVIEW').length);
  add(P, 'ნაკლებად გადახდილი, ნაშთი, ₾ (' + ids(st('UNDERPAID')) + ')',
      r2(st('UNDERPAID').reduce((s, o) => s + Number(o['Total amount']) - Number(o['Paid amount']), 0)));
  const noInv = bank.filter(r => r[9] === 'NO_INVOICE');
  add(P, 'გადახდა ინვოისის გარეშე, ₾ (' + noInv.map(r => r[4]).join(', ') + ')', r2(noInv.reduce((s, r) => s + (Number(r[6]) || 0), 0)));
  const card = inv.filter(o => o['Payment status'] === 'PAID' && Number(o['Paid amount']) > Number(o['Total amount']) + 0.01);
  add(P, 'ბარათით გადახდა ინვოისზე მეტი (ფარული კურსის სხვაობა), ' + [...new Set(card.map(o => o['Currency']))].join('/') + ' (' + ids(card) + ')',
      r2(card.reduce((s, o) => s + Number(o['Paid amount']) - Number(o['Total amount']), 0)));
  const shop = read('Card sales (shop)').find(o => String(o['Month']).trim() === month) || {};
  add(P, 'PayKart: მაღაზიის გაყიდვები − (ჩარიცხვა + საკომისიო), ₾ (' + (shop['Status'] || '–') + ')', Number(shop['Difference']) || 0);

  // 3. დებიტორები
  const A = 'დებიტორები ' + asOf, si = read('Sales invoices');
  const outOf = arr => r2(arr.reduce((s, o) => s + (Number(o['Outstanding (GEL)']) || 0), 0));
  add(A, 'გადაუხდელი ნაშთი, ₾', outOf(si));
  ['61+', '31-60', '1-30', 'Not due'].forEach(b => add(A, 'მათ შორის ' + b + ', ₾', outOf(si.filter(o => String(o['Aging bucket']).trim() === b))));
  add(A, 'ვადაგადაცილებული კლიენტი', new Set(si.filter(o => Number(o['Outstanding (GEL)']) > 0 && Number(o['Days late / overdue']) > 0).map(o => o['Customer ID'])).size);
  add(A, 'შეხსენების მონახაზი (არ გაგზავნილა)', read('AR reminders').length);
  add(A, 'საკრედიტო ფლაგი', si.filter(o => o['Control flags']).map(o => o['Invoice No'] + ' (' + o['Customer ID'] + '): ' + o['Control flags']).join('; ') || 'არა');

  // 4. გადასახდელები
  const G = 'გადასახდელები';
  const pr = ss.getSheetByName('Payment run').getDataRange().getValues().filter(r => /^\d+$/.test(String(r[0]).trim()));
  add(G, 'Payment run ' + iso(S['Payment run date']) + ': გადახდა', pr.length);
  add(G, 'Payment run: ჯამი, ₾', r2(pr.reduce((s, r) => s + Number(r[4]), 0)));
  ['HOLD', 'BLOCKED'].forEach(d => add(G, d + ' (ბუღალტრის გადაწყვეტილება)', ids(inv.filter(o => o['Payment decision'] === d)) || 'არა'));

  // 5. პროგნოზი + ყველაზე სახიფათო კვირის მიზეზი
  const R = 'პროგნოზი (13 კვირა)';
  const fc = ss.getSheetByName('Cash forecast').getDataRange().getValues().slice(3).filter(r => r[0] !== '');
  const minOf = c => fc.reduce((m, r) => Number(r[c]) < m[0] ? [Number(r[c]), r[0], iso(r[1])] : m, [Infinity]);
  const mb = minOf(5), mp = minOf(8);
  add(R, 'უმცირესი ნაშთი, საბაზისო, ₾ (კვირა ' + mb[1] + ', ' + mb[2] + ')', mb[0]);
  add(R, 'უმცირესი ნაშთი, პესიმისტური, ₾ (კვირა ' + mp[1] + ', ' + mp[2] + ')', mp[0]);
  add(R, 'კვირები ზღვარს (' + fmtN(Number(S['Minimum cash buffer (GEL)'])) + ' ₾) ქვემოთ', fc.filter(r => r[9]).map(r => 'W' + r[0] + ': ' + r[9]).join('; ') || 'არა');
  const top = ss.getSheetByName('Forecast detail').getDataRange().getValues().slice(1)
    .filter(r => Number(r[6]) === Number(mp[1]) && Number(r[7]) < 0).sort((a, b) => a[7] - b[7]).slice(0, 3);
  add(R, 'უდიდესი გასავალი კვირა ' + mp[1] + '-ში, ₾', top.map(r => r[2] + ' ' + fmtN(-Number(r[7]))).join('; '));
  F.forEach(f => Logger.log('[' + f[0] + '] ' + f[1] + ': ' + f[2]));

  // 6. AI – ახსნა მხოლოდ ფაქტებიდან
  const facts = F.map(f => '- [' + f[0] + '] ' + f[1] + ': ' + f[2]).join('\n');
  const prompt = 'შენ ხარ გამოცდილი ფინანსური მენეჯერი. დაწერე თვის დახურვის მოკლე ანგარიში მცირე კომპანიის მფლობელისთვის, ქართულად, ბუნებრივი და მარტივი ენით, როგორც ადამიანი წერს ადამიანს.\n' +
    'წესები:\n1. გამოიყენე მხოლოდ ქვემოთ მოცემული ფაქტები. ახალი ფაქტი არ გამოიგონო.\n' +
    '2. ახალი რიცხვი არ გამოთვალო (ჯამი, პროცენტი, სხვაობა). რიცხვი გადმოიტანე ზუსტად ფაქტებიდან, იმავე ფორმატით (მაგ. 7,074.10).\n' +
    '3. ფაქტების სათაურები სიტყვასიტყვით ნუ გადაიწერე: აზრი თავისი სიტყვებით, რიცხვი ზუსტად.\n' +
    '4. „მთავარი“: 3 წინადადება. რა არის ამ თვის ყველაზე მნიშვნელოვანი (რისკის ქვეშ მყოფი ფული, ფულის ნაკლებობის საფრთხე და მისი მიზეზი). არა მხოლოდ შემოსავლის ჩამონათვალი.\n' +
    '5. „რისკები“: მაქსიმუმ 5 პუნქტი, ყველაზე სერიოზული პირველი. თითოეულში მოკლედ: რატომ არის მნიშვნელოვანი.\n' +
    '6. „რეკომენდებული ნაბიჯები“: 4–6 პუნქტი, ფორმატით „ვინ – რა – როდემდე (ამ კვირაში / თვის ბოლომდე)“. მხოლოდ კონკრეტული მოქმედება (დარეკვა, მოთხოვნა, დადასტურება, გადადება, შეჩერება), არა ზოგადი რჩევა.\n' +
    '7. თაღლითობის და იურიდიული შედეგების შესახებ დასკვნა არ გააკეთო, დაწერე, რომ ბანკთან და იურისტთან დასაზუსტებელია. პრაქტიკული პირველი ნაბიჯი (ბანკთან დაკავშირება) შეგიძლია ურჩიო.\n' +
    '8. ერთ წინადადებაში აღნიშნე, რომ მონაცემები სინთეტიკურია.\n' +
    '9. მარტივი ტექსტი: სათაურები ცალკე ხაზზე, პუნქტები „- “-ით, markdown-ის გარეშე.\n\nფაქტები:\n' + facts;
  const props = PropertiesService.getScriptProperties().getProperties();
  const key = props.GEMINI_API_KEY || props[Object.keys(props).find(k => /gemini|api/i.test(k))];
  if (!key) throw new Error('Gemini API key ვერ მოიძებნა Script Properties-ში');
  const models = [S['Report model'], S['Gemini model']].map(m => String(m || '').trim()).filter(Boolean);
  let ai = '', used = '';
  for (const m of models) {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2 } }) });
    if (res.getResponseCode() === 200) { ai = JSON.parse(res.getContentText()).candidates[0].content.parts.map(p => p.text || '').join('').trim(); used = m; break; }
    Logger.log('⚠️ ' + m + ': HTTP ' + res.getResponseCode() + ' – trying next model');
  }
  if (!ai) throw new Error('Gemini: არც ერთმა მოდელმა არ უპასუხა');

  // 7. შემოწმება: AI-ის ყველა რიცხვი ფაქტებშია?
  const num = s => { const n = Number(s.replace(/[\s,]/g, '')); return isNaN(n) ? s : String(n); };
  const known = new Set((facts.match(/\d[\d,.]*/g) || []).map(s => num(s.replace(/[.,]$/, ''))));
  const found = [...new Set((ai.match(/\d[\d,.]*/g) || []).map(s => num(s.replace(/[.,]$/, ''))))];
  const unknown = found.filter(n => !known.has(n) && !(Number(n) <= 31) && !/^202[67]$/.test(n));

  // 8. Google Doc – მონახაზი
  const H1 = DocumentApp.ParagraphHeading.HEADING1;
  const doc = DocumentApp.create('Worldwide – Month-end report ' + month + ' (synthetic)'), b = doc.getBody();
  b.appendParagraph('Worldwide – თვის დახურვის ანგარიში, ' + month).setHeading(DocumentApp.ParagraphHeading.TITLE);
  b.appendParagraph('⚠️ სინთეტიკური მონაცემები (გამოგონილი კომპანია). სტატუსი: მონახაზი – გაგზავნამდე გადაამოწმეთ. შექმნილია: ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'));
  b.appendParagraph('1. ციფრები (ითვლის კოდი)').setHeading(H1);
  b.appendTable([['განყოფილება', 'მაჩვენებელი', 'მნიშვნელობა']].concat(F.map(f => f.map(String))));
  b.appendParagraph('2. ახსნა და ნაბიჯები (წერს AI, ' + used + ')').setHeading(H1);
  b.appendParagraph(unknown.length ? '⚠️ AI-ის ტექსტში არის რიცხვი, რომელიც ფაქტებში არ არის: ' + unknown.join(', ') + ' – გადაამოწმეთ!'
                                   : '✅ AI-ის ტექსტის ყველა რიცხვი ფაქტებს ემთხვევა.');
  ai.split('\n').forEach(line => b.appendParagraph(line));
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFileById(ss.getId()).getParents().next());
  Logger.log('Model: ' + used + ' | Numbers check: ' + (unknown.length ? '⚠️ NOT IN FACTS: ' + unknown.join(', ') : 'OK – all AI numbers are in facts'));
  Logger.log('Report: ' + doc.getUrl());
}
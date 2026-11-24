// Reminders.gs – შეხსენებების მონახაზები. არაფერს აგზავნის!
function draftReminders() {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const read = name => {
    const v = ss.getSheetByName(name).getDataRange().getValues(); const h = v[0].map(String);
    return v.slice(1).filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
  };
  const fmt = n => { const p = Number(n).toFixed(2).split('.'); return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1]; };
  const settings = {}; ss.getSheetByName('Settings').getDataRange().getValues().forEach(r => settings[r[0]] = r[1]);
  const asOf = toISO(settings['AR as-of date'], tz);
  const ourIban = String(settings['Company IBAN'] || '').trim();
  if (!ourIban) throw new Error('Settings: "Company IBAN" ცარიელია');
  const cust = {}; read('Customers').forEach(c => cust[c['Customer ID']] = c);

  // 1. ვადაგადაცილებული ნაშთები კლიენტების მიხედვით (reconcileReceivables-ის შედეგიდან)
  const due = {}, flags = {};
  read('Sales invoices').forEach(i => {
    const cid = i['Customer ID'], out = Number(i['Outstanding (GEL)']), d = Number(i['Days late / overdue']);
    if (i['Control flags']) flags[cid] = (flags[cid] ? flags[cid] + '; ' : '') + i['Invoice No'] + ': ' + i['Control flags'];
    if (out > 0.009 && d > 0) (due[cid] = due[cid] || []).push({ no: i['Invoice No'], date: toISO(i['Due date'], tz), out, d, total: Number(i['Total (GEL)']) });
  });

  // 2. ტექსტები
  const subj = { 1: 'შეხსენება: ვადაგადაცილებული ინვოისი', 2: 'განმეორებითი შეხსენება: ვადაგადაცილებული დავალიანება', 3: 'საბოლოო შეხსენება: ვადაგადაცილებული დავალიანება' };
  const ask = {
    1: 'გთხოვთ, დაფაროთ დავალიანება უახლოეს დღეებში.',
    2: 'ვადაგადაცილება 30 დღეს აღემატება. გთხოვთ, დაფაროთ 7 დღის განმავლობაში ან დაგვიკავშირდეთ გადახდის გრაფიკის შესათანხმებლად.',
    3: 'ვადაგადაცილება 60 დღეს აღემატება. დავალიანების დაფარვის გარეშე ახალი მიწოდებების ნისიად შეჩერება მოგვიწევს. გთხოვთ, დაგვიკავშირდეთ 5 სამუშაო დღის განმავლობაში.'
  };
  const rows = [], log = [];
  Object.keys(due).sort().forEach(cid => {
    const c = cust[cid], list = due[cid];
    const sum = Math.round(list.reduce((s, i) => s + i.out, 0) * 100) / 100;
    const worst = Math.max(...list.map(i => i.d));
    const level = worst > 60 ? 3 : worst > 30 ? 2 : 1;
    const lines = list.map(i => '• ' + i.no + ' – ვადა ' + i.date + ', დარჩენილი ' + fmt(i.out) + ' ₾' +
      (i.out < i.total ? ' (ნაწილობრივ გადახდილი)' : '') + ', ვადაგადაცილება ' + i.d + ' დღე').join('\n');
    const thanks = list.filter(i => i.out < i.total)
      .map(i => 'მადლობა ' + i.no + '-ზე უკვე ჩარიცხული ' + fmt(i.total - i.out) + ' ₾-ისთვის.').join(' ');
    const body = 'გამარჯობა,\n\n' + (thanks ? thanks + '\n\n' : '') +
      asOf + '-ის მდგომარეობით ჩვენს აღრიცხვაში ვადაგადაცილებულად ფიქსირდება:\n\n' + lines +
      '\n\nსულ: ' + fmt(sum) + ' ₾\n\n' + ask[level] +
      '\n\nგადახდა შეგიძლიათ ანგარიშზე: ' + ourIban + ' (შპს „Worldwide“). გადარიცხვისას დანიშნულებაში მიუთითეთ ინვოისის ნომერი.' +
      '\n\nთუ გადახდა უკვე განხორციელდა, გთხოვთ, გამოგვიგზავნოთ საგადახდო დავალება – ამ შემთხვევაში წერილი უგულებელყავით.' +
      '\n\nპატივისცემით,\nშპს „Worldwide“ – ფინანსური განყოფილება';
    const note = [level === 3 ? 'წერილი აფრთხილებს მიწოდების შეჩერებაზე – გაგზავნამდე დაადასტურეთ, ნამდვილად ვაპირებთ თუ არა' : '',
                  flags[cid] ? 'Control flags: ' + flags[cid] : ''].filter(Boolean).join(' | ');
    rows.push([cid, c['Name'], c['Contact email'], level, sum, worst, subj[level] + ' – ' + c['Name'], body, note, 'Draft – not sent', '']);
    log.push(cid + ' L' + level + ' ' + sum + ' (' + worst + 'd)');
  });

  // 3. ჩაწერა ცალკე ფურცელზე
  const sh = ss.getSheetByName('AR reminders') || ss.insertSheet('AR reminders');
  sh.clear();
  const head = ['Customer ID', 'Customer', 'To', 'Level', 'Overdue total (GEL)', 'Worst days', 'Subject', 'Body', 'Internal note', 'Status', 'Approved by'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.getRange(1, 1, sh.getMaxRows(), head.length).setVerticalAlignment('top');
  sh.getRange('G:I').setWrap(true);
  sh.setColumnWidth(7, 220); sh.setColumnWidth(8, 520); sh.setColumnWidth(9, 260);
  Logger.log('Drafts: ' + rows.length + ' | ' + log.join(' | '));
}
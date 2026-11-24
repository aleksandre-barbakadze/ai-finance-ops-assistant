function scoreResults() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const toObjs = name => {
    const v = ss.getSheetByName(name).getDataRange().getValues();
    return v.slice(1).filter(r => r[0]).map(r => { const o = {}; v[0].forEach((h, i) => o[h] = r[i]); return o; });
  };
  const inv = {};
    toObjs('Invoices').forEach(o => inv[o['Invoice ID']] = o);
  const key = toObjs('Answer key');

  // რომელ ველს რომელს ვადარებთ
  const fields = [
    ['Supplier ID', 'True supplier ID', 'text'], ['Invoice number', 'True invoice number', 'text'],
    ['Invoice date', 'True invoice date', 'date'], ['Due date', 'True due date', 'date'],
    ['Currency', 'True currency', 'text'], ['Net amount', 'True net', 'num'],
    ['VAT amount', 'True VAT', 'num'], ['Total amount', 'True total', 'num']
  ];
  const same = (a, b, type) => {
    if (type === 'num') return Math.abs((Number(a) || 0) - (Number(b) || 0)) <= 0.01;
    if (type === 'date') return toISO(a, tz) === toISO(b, tz);
    return String(a || '').trim() === String(b || '').trim();
  };

  const correct = {}; fields.forEach(f => correct[f[0]] = 0);
  const errors = [];
  let missing = 0, trapsTotal = 0, trapsCaught = 0, cleanTotal = 0, cleanOk = 0, falseAlarms = 0;

  key.forEach(k => {
        const o = inv[k['Invoice ID']];
    if (!o) { missing++; errors.push([k['Invoice ID'], '(not extracted)', '', '']); return; }

    // 1. ველების სიზუსტე
    fields.forEach(([f, t, type]) => {
      if (same(o[f], k[t], type)) correct[f]++;
      else errors.push([k['Invoice ID'], f, toISO(k[t], tz), toISO(o[f], tz)]);
    });

    // 2. ხაფანგები და ცრუ განგაში
    const predicted = [o['AI flags'], o['Check flags']].join(',').split(',').map(s => s.trim()).filter(Boolean);
    const expected = String(k['Expected flag'] || '').trim();
    if (expected) {
      trapsTotal++;
      if (predicted.includes(expected)) trapsCaught++;
      else errors.push([k['Invoice ID'], 'MISSED FLAG', expected, predicted.join(', ')]);
    } else {
      cleanTotal++;
      if (!predicted.length) cleanOk++;
    }
    predicted.filter(p => p !== expected).forEach(p => { falseAlarms++; errors.push([k['Invoice ID'], 'FALSE ALARM', '', p]); });
  });

  // 3. შედეგების ფურცელი
  const n = key.length - missing;
  const out = ss.getSheetByName('Test results') || ss.insertSheet('Test results');
  out.clear();
  const rows = [
    ['AI Finance Ops – test results (synthetic data)', ''],
    ['Run at', new Date()],
    ['Model(s)', [...new Set(Object.values(inv).map(o => o['Model']))].join(', ')],
    ['Documents in answer key', key.length],
    ['Extracted', n],
    ['', ''],
    ['FIELD ACCURACY', '']
  ];
  fields.forEach(([f]) => rows.push([f, correct[f] + ' / ' + n]));
  rows.push(['', ''], ['FLAGS', ''],
    ['Planted issues caught', trapsCaught + ' / ' + trapsTotal],
    ['Clean documents left clean', cleanOk + ' / ' + cleanTotal],
    ['False alarms (extra flags)', falseAlarms]);
  out.getRange(1, 1, rows.length, 2).setValues(rows);
  out.getRange(rows.length + 2, 1, 1, 4).setValues([['Invoice', 'Field / issue', 'Expected', 'Got']]).setFontWeight('bold');
  if (errors.length) out.getRange(rows.length + 3, 1, errors.length, 4).setValues(errors);
  out.getRange(1, 1).setFontWeight('bold');
  Logger.log('Planted issues ' + trapsCaught + '/' + trapsTotal + ', false alarms ' + falseAlarms + ', error lines ' + errors.length);
}

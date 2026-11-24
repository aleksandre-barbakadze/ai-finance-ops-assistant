function toISO(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v || '').trim();
}
function normIban(s) { return String(s || '').replace(/\s/g, '').toUpperCase(); }

function validateInvoices() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const settings = getSettings();
  const tol = Number(settings['Amount tolerance (GEL)']) || 1;
  const windowDays = Number(settings['Duplicate check window (days)']) || 60;

  // 1. მომწოდებლების ცნობარი
  const sup = {};
  const sRows = ss.getSheetByName('Suppliers').getDataRange().getValues();
  sRows.slice(1).forEach(r => { const o = {}; sRows[0].forEach((h, i) => o[h] = r[i]); if (o['Supplier ID']) sup[o['Supplier ID']] = o; });

  // 2. ეროვნული ბანკის კურსები
  const fx = {};
  ss.getSheetByName('FX rates').getDataRange().getValues().slice(1)
    .forEach(r => { if (r[0]) fx[toISO(r[0], tz)] = { USD: Number(r[1]), EUR: Number(r[2]) }; });
  const fxDates = Object.keys(fx).sort();
  const rateFor = (cur, date) => {
    if (cur === 'GEL') return 1;
    let d = date;
    if (!fx[d]) { const before = fxDates.filter(x => x <= date); d = before.length ? before[before.length - 1] : fxDates[0]; }
    return fx[d] ? fx[d][cur] : '';
  };

  // 3. ამოკითხული ინვოისები
  const sheet = ss.getSheetByName('Invoices');
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const c = name => h.indexOf(name) + 1;
  const rows = data.slice(1).map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; })
    .filter(o => o['Invoice ID']);
  rows.forEach(o => { o._date = toISO(o['Invoice date'], tz); o._due = toISO(o['Due date'], tz); });

  // 4. თითოეული ინვოისის შემოწმება
  rows.forEach((o, idx) => {
    const flags = [];
    const s = sup[o['Supplier ID']];
    const net = Number(o['Net amount']) || 0, vat = Number(o['VAT amount']) || 0, total = Number(o['Total amount']) || 0;

    if (!s) flags.push('UNKNOWN_SUPPLIER');
    else {
      const masterIban = normIban(s['IBAN']);
      if (/^[A-Z]{2}\d/.test(masterIban) && o['Supplier IBAN'] && normIban(o['Supplier IBAN']) !== masterIban) flags.push('IBAN_CHANGED');
      if (String(s['VAT payer (GE)']).toLowerCase() === 'no' && vat > 0) flags.push('VAT_NOT_EXPECTED');
      if (o['Currency'] && o['Currency'] !== s['Currency']) flags.push('CURRENCY_MISMATCH');
      if (!o._due && o['Document type'] === 'invoice' && s['Country'] === 'Georgia' && !/deducted/i.test(s['Payment terms'])) flags.push('MISSING_DUE_DATE');
    }
    if (Math.abs(net + vat - total) > tol) flags.push('TOTAL_MISMATCH');

    // დუბლიკატები — ვადარებთ იმავე მომწოდებლის ადრინდელ ინვოისებს
    const recurring = s && String(s['Recurring']).toLowerCase() === 'yes';
    rows.forEach((p, j) => {
      if (j === idx || p['Supplier ID'] !== o['Supplier ID']) return;
      const isEarlier = p._date < o._date || (p._date === o._date && j < idx);
      if (!isEarlier) return;
      if (String(p['Invoice number']) === String(o['Invoice number'])) {
        if (!flags.includes('DUPLICATE')) flags.push('DUPLICATE');
      } else if (!recurring && Math.abs((Number(p['Total amount']) || 0) - total) <= 0.01) {
        const days = (new Date(o._date) - new Date(p._date)) / 86400000;
        if (days <= windowDays && !flags.includes('POSSIBLE_DUPLICATE')) flags.push('POSSIBLE_DUPLICATE');
      }
    });

    // ლარში გადაყვანა — კოდით, არა AI-ით
    const rate = rateFor(o['Currency'], o._date);
    const totalGel = rate ? Math.round(total * rate * 100) / 100 : '';
    const allFlags = [o['AI flags'], flags.join(', ')].filter(Boolean).join(', ');

    sheet.getRange(o._row, c('FX rate')).setValue(rate);
    sheet.getRange(o._row, c('Total GEL')).setValue(totalGel);
    sheet.getRange(o._row, c('Category')).setValue(s ? s['Default category'] : 'Uncategorized');
    sheet.getRange(o._row, c('Check flags')).setValue(flags.join(', '));
    if (o['Status'] !== 'Approved') sheet.getRange(o._row, c('Status')).setValue(allFlags ? 'Needs review' : 'Extracted');
  });
  Logger.log('Validated ' + rows.length + ' invoices');
}
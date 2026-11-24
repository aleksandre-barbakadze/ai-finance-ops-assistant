function loadFxRates() {
  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName('Settings');
  const fxSheet = ss.getSheetByName('FX rates');
  const tz = ss.getSpreadsheetTimeZone();

  // 1. პერიოდს Settings-დან ვკითხულობთ
  const start = new Date(settings.getRange('B5').getValue());
  const end = new Date(settings.getRange('B6').getValue());

  // 2. თითოეულ დღეზე ეროვნულ ბანკს ვეკითხებით
  const rows = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    const dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    const url = 'https://nbg.gov.ge/gw/api/ct/monetarypolicy/currencies/en/json/?date=' + dateStr;
    const data = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    const list = data[0].currencies;
    const usd = list.find(c => c.code === 'USD');
    const eur = list.find(c => c.code === 'EUR');
    rows.push([dateStr, usd.rate / usd.quantity, eur.rate / eur.quantity, 'NBG official']);
    Utilities.sleep(200); // სერვერს არ ვტვირთავთ
  }

  // 3. ძველს ვშლით და ახალს ერთბაშად ვწერთ
  fxSheet.getRange(2, 1, fxSheet.getMaxRows() - 1, 4).clearContent();
  fxSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  Logger.log('Loaded ' + rows.length + ' days');
}

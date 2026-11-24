// Settings-დან პარამეტრებს სახელით ვკითხულობთ
function getSettings() {
  const rows = SpreadsheetApp.getActive().getSheetByName('Settings').getDataRange().getValues();
  const map = {};
  rows.forEach(r => map[r[0]] = r[1]);
  return map;
}

// 5000 → "5,000.00"
let euFormat = false;   // true = ევროპული ციფრები (16.420,00)

function fmt(n) {
  const s = Number(String(n).replace(/,/g, '')).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return euFormat ? s.replace(/,/g, '#').replace('.', ',').replace(/#/g, '.') : s;
}

function generateInvoices() {
  const settings = getSettings();
  const folder = DriveApp.getFolderById(settings['Invoices folder ID']);

  // შაბლონებს Settings-დან თავისით ვპოულობთ: "GE template ID" → GE, "CN template ID" → CN ...
  const templates = {};
  Object.keys(settings).forEach(k => {
    if (k.endsWith(' template ID')) templates[k.replace(' template ID', '')] = settings[k];
  });

  const data = SpreadsheetApp.getActive().getSheetByName('Invoice source').getDataRange().getDisplayValues();
  const header = data[0];
  let created = 0;

  data.slice(1).forEach(row => {
    const inv = {};
    header.forEach((h, i) => inv[h] = row[i]);
        if (!inv['Invoice ID']) return;                                          // ცარიელი სტრიქონი
    if (!templates[inv['Template']]) { Logger.log('⚠️ ' + inv['Invoice ID'] + ': შაბლონი "' + inv['Template'] + '" Settings-ში ვერ მოიძებნა'); return; }                                                                  // ცარიელი ან უშაბლონო
    
    if (folder.getFilesByName(inv['File name']).hasNext()) return;           // უკვე შექმნილია
        euFormat = (inv['Template'] === 'TR');                                   // თურქულზე — ევროპული ციფრები

    const copy = DriveApp.getFileById(templates[inv['Template']]).makeCopy('tmp ' + inv['Invoice ID'], folder);
    try {
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();

      // 1. დამალული ტექსტი (ხაფანგი #10)
      const found = body.findText('\\{\\{hidden_note\\}\\}');
      if (found) {
        const text = found.getElement().asText();
        text.replaceText('\\{\\{hidden_note\\}\\}', inv['Hidden note']);
        if (inv['Hidden note']) text.setForegroundColor('#FFFFFF').setFontSize(1);
      }

      // 2. პოზიციების ცხრილი
      const table = body.getTables()[0];
      const templateRow = table.getRow(1);
      const items = inv['Items'].split(';').map(s => s.trim()).filter(s => s);
      items.forEach((item, i) => {
        const [name, qty, price] = item.split('|');
        const r = table.insertTableRow(1 + i, templateRow.copy());
        r.replaceText('\\{\\{item\\}\\}', name);
        r.replaceText('\\{\\{qty\\}\\}', qty);
        r.replaceText('\\{\\{price\\}\\}', fmt(price));
        r.replaceText('\\{\\{line_total\\}\\}', fmt(qty * price));
      });
      table.removeRow(1 + items.length);

      // 3. დანარჩენი ადგილის დამჭერები
      const values = {
        invoice_number: inv['Invoice number'], invoice_date: inv['Invoice date'], due_date: inv['Due date'],
        supplier_name: inv['Supplier name'], tax_id: inv['Tax ID'], supplier_address: inv['Address'],
        iban: inv['IBAN'], currency: inv['Currency'], terms: inv['Terms'] || '',
        net: fmt(inv['Net']), vat: fmt(inv['VAT']), total: fmt(inv['Total'])
      };
      Object.keys(values).forEach(k => body.replaceText('\\{\\{' + k + '\\}\\}', values[k]));
      doc.saveAndClose();

      // 4. PDF
      folder.createFile(copy.getAs('application/pdf')).setName(inv['File name']);
      created++;
    } catch (e) {
      Logger.log('❌ ' + inv['Invoice ID'] + ': ' + e.message);                // შეცდომა — სხვებს ვაგრძელებთ
    } finally {
      copy.setTrashed(true);                                                   // დროებითი ასლი ყოველთვის იშლება
    }
  });

  Logger.log('Created ' + created + ' PDF(s)');
}
// Runs in a SEPARATE inbox Google account (least privilege: it can only see the shared Inbox folder).
// Inbox folder ID – copy it from the main spreadsheet's Settings sheet (row "Inbox folder ID").
const INBOX_FOLDER_ID = 'YOUR_INBOX_FOLDER_ID';

function collectInvoiceEmails() {
  const folder = DriveApp.getFolderById(INBOX_FOLDER_ID);
  const label = GmailApp.getUserLabelByName('processed') || GmailApp.createLabel('processed');
  const tz = Session.getScriptTimeZone();

  // ახალი წერილები დანართით, რომლებიც ჯერ არ დამუშავებულა
  const threads = GmailApp.search('has:attachment -label:processed', 0, 20);
  let saved = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      msg.getAttachments({ includeInlineImages: false }).forEach(att => {
        const type = att.getContentType();
        const isPdf = /pdf/i.test(type);
        const isPhoto = /image\/(jpeg|jpg|png)/i.test(type) && att.getSize() > 30 * 1024;
        if (!isPdf && !isPhoto) return;                      // ლოგოები და სხვა ფაილები — გამოტოვება

        let name = att.getName();
        if (folder.getFilesByName(name).hasNext()) {         // იგივე სახელი უკვე არის — დრო ვუმატებთ
          const stamp = Utilities.formatDate(msg.getDate(), tz, 'yyyyMMdd-HHmmss');
          name = name.replace(/(\.[^.]+)$/, '_' + stamp + '$1');
        }
        folder.createFile(att.copyBlob()).setName(name)
          .setDescription('From: ' + msg.getFrom() + ' | Received: ' + msg.getDate());
        saved++;
      });
    });
    thread.addLabel(label);                                  // აღარ ავიღებთ მეორედ
  });

  Logger.log('Saved ' + saved + ' file(s) from ' + threads.length + ' email(s)');
}

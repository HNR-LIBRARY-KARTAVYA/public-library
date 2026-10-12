// ================================================================
// MEMBER MAIL (Koha-style e-mails) - Honavar
// Used by Code.gs for: Book Issued, Book Returned, Book Renewed,
// and the footer of every reminder e-mail.
// Every e-mail shows: Library Card No, member name, Accession No,
// title, author, issue date, due date + our website link.
// To change wording, edit only this file.
// ================================================================
var MM_ = {
  library: 'District Central Library, Karwar',
  branch: 'Branch: Honavar',
  shortName: 'Public Library Honavar',
  website: 'https://hnr-library-kartavya.github.io/public-library/honavar/',
  email: 'dclkarwarbranchhonnavar@gmail.com',
  timings: '8:00-11:30 AM and 4:00-8:00 PM (closed Mondays, 2nd Tuesday, 4th Saturday and government holidays)',
  colour: '#0d3b8c'
};

// Accession No -> { title, author } from Books Catalogue (read once per run).
var MM_BOOK_CACHE_ = null;
function mmBook_(bookId) {
  if (!MM_BOOK_CACHE_) {
    MM_BOOK_CACHE_ = {};
    try {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Books Catalogue');
      var last = sh.getLastRow();
      if (last >= BOOKS_START_ROW) {
        sh.getRange(BOOKS_START_ROW, 1, last - BOOKS_START_ROW + 1, 3).getValues().forEach(function (r) {
          var id = normId_(r[0]); if (id) MM_BOOK_CACHE_[id] = { title: String(r[1] || ''), author: String(r[2] || '') };
        });
      }
    } catch (e) {}
  }
  return MM_BOOK_CACHE_[normId_(bookId)] || { title: '', author: '' };
}

function mmEsc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function mmDate_(v) { return v instanceof Date ? fmtDate_(v) : String(v || ''); }

// Short label used inside reminder lists: "Title (Acc No: G-123)"
function mmItemLabel_(title, bookId) { return String(title || '') + (bookId ? ' (Acc No: ' + bookId + ')' : ''); }

// Footer added to every member e-mail (plain text).
function mmFooter_() {
  return 'Search books, check availability and see today\'s newspapers on our website:\n' + MM_.website + '\n\n' +
    'Library timings: ' + MM_.timings + '\n' +
    'Questions? Just reply to this e-mail or write to ' + MM_.email + '\n\n' +
    'This is an automatically generated e-mail.\n\n' +
    'Thank you,\n' + MM_.library + '\n' + MM_.branch;
}

function mmFooterHtml_() {
  return '<p style="margin:18px 0 6px">Search books, check availability and see today\'s newspapers on our website:<br>' +
    '<a href="' + MM_.website + '" style="color:' + MM_.colour + ';font-weight:bold">' + MM_.website + '</a></p>' +
    '<p style="margin:6px 0;color:#555">Library timings: ' + mmEsc_(MM_.timings) + '<br>' +
    'Questions? Just reply to this e-mail or write to <a href="mailto:' + MM_.email + '">' + MM_.email + '</a></p>' +
    '<p style="margin:14px 0 4px;color:#888;font-size:12px">This is an automatically generated e-mail.</p>' +
    '<p style="margin:10px 0 0">Thank you,<br><b>' + MM_.library + '</b><br>' + MM_.branch + '</p>';
}

// Builds one e-mail. kind: 'issue' | 'return' | 'renew'
// items: [{ bookId, title, issueDate, dueDate, returnDate, overdueDays, fine }]
function mmBuild_(kind, memberId, memberName, items) {
  var heads = {
    issue: 'The following item(s) have been issued:',
    'return': 'The following item(s) have been returned. Thank you!',
    renew: 'The following item(s) have been renewed:'
  };
  var subjWord = { issue: 'Issued', 'return': 'Returned', renew: 'Renewed' }[kind];
  var subject = (items.length === 1
      ? ('Book ' + subjWord + ': ' + (items[0].title || items[0].bookId))
      : (items.length + ' Books ' + subjWord)) + ' - ' + MM_.shortName;

  var txt = 'Dear ' + memberName + ',\n\n' + heads[kind] + '\n\n' +
    'Library Card No: ' + memberId + '\nName of the Member: ' + memberName + '\n\n' +
    (items.length > 1 ? 'Book details as follows' : 'Book details as follows') + '\n';
  var rowsHtml = '';
  var totalFine = 0;
  items.forEach(function (it, i) {
    var b = mmBook_(it.bookId);
    var title = it.title || b.title, author = b.author;
    var lines = [['Acc No / Barcode', it.bookId], ['Title', title], ['Author', author]];
    if (kind === 'issue') { lines.push(['Issue Date', mmDate_(it.issueDate)]); lines.push(['Due Date', mmDate_(it.dueDate)]); }
    if (kind === 'renew') { lines.push(['Renewed On', mmDate_(it.issueDate)]); lines.push(['New Due Date', mmDate_(it.dueDate)]); }
    if (kind === 'return') {
      lines.push(['Returned On', mmDate_(it.returnDate)]);
      if (it.fine > 0) { lines.push(['Late fee', 'Rs.' + it.fine + ' (' + it.overdueDays + ' day(s) late)']); totalFine += it.fine; }
    }
    txt += '-\n' + lines.filter(function (l) { return l[1]; }).map(function (l) { return l[0] + ': ' + l[1]; }).join('\n') + '\n';
    rowsHtml += '<table style="border-collapse:collapse;width:100%;margin:0 0 10px;background:#fff;border:1px solid #e4d9bd;border-radius:8px">' +
      lines.filter(function (l) { return l[1]; }).map(function (l) {
        return '<tr><td style="padding:6px 10px;color:#666;width:38%;border-bottom:1px solid #f1ead8">' + l[0] +
          '</td><td style="padding:6px 10px;font-weight:bold;border-bottom:1px solid #f1ead8">' + mmEsc_(l[1]) + '</td></tr>';
      }).join('') + '</table>';
  });
  var note = kind === 'return'
    ? (totalFine > 0 ? 'Total late fee: Rs.' + totalFine + ' - payable at the library.' : '')
    : 'Please return or renew the book(s) on or before the due date.';
  if (note) txt += '\n' + note + '\n';
  txt += '\n' + mmFooter_();

  var html = '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;color:#222;font-size:14px;line-height:1.5">' +
    '<div style="background:' + MM_.colour + ';color:#fff;padding:14px 16px;border-radius:10px 10px 0 0">' +
    '<div style="font-size:16px;font-weight:bold">' + MM_.library + '</div><div style="opacity:.9">' + MM_.branch + '</div></div>' +
    '<div style="background:#fbf6ea;padding:16px;border-radius:0 0 10px 10px">' +
    '<p style="margin:0 0 10px">Dear <b>' + mmEsc_(memberName) + '</b>,</p><p style="margin:0 0 12px">' + heads[kind] + '</p>' +
    '<p style="margin:0 0 12px">Library Card No: <b>' + mmEsc_(memberId) + '</b><br>Name of the Member: <b>' + mmEsc_(memberName) + '</b></p>' +
    rowsHtml + (note ? '<p style="margin:8px 0;font-weight:bold">' + note + '</p>' : '') + mmFooterHtml_() + '</div></div>';
  return { subject: subject, body: txt, htmlBody: html };
}

// Sends one Koha-style e-mail, logs it, returns 'sent' | 'no-email' | 'failed' | 'off'.
function mmSend_(kind, memberId, items) {
  if (!items || !items.length) return 'off';
  if (!LIB_EMAIL_NOTIFICATIONS_LIVE) return 'off';
  var type = { issue: 'Book Issued', 'return': 'Book Returned', renew: 'Book Renewed' }[kind] + (items.length > 1 ? ' (' + items.length + ')' : '');
  var detail = items.map(function (it) { return mmItemLabel_(it.title, it.bookId); }).join('; ');
  var contact = getMemberContact_(memberId);
  if (!contact || !contact.email) {
    logEmail_('Other', type, memberId, contact ? contact.name : '', '', detail, 'Skipped-NoEmail');
    return 'no-email';
  }
  try {
    var m = mmBuild_(kind, normId_(memberId), contact.name, items);
    MailApp.sendEmail({ to: contact.email, subject: m.subject, body: m.body, htmlBody: m.htmlBody, name: MM_.shortName, replyTo: MM_.email });
    logEmail_('Other', type, memberId, contact.name, contact.email, detail, 'Sent');
    return 'sent';
  } catch (e) {
    logEmail_('Other', type, memberId, contact.name, contact.email, detail, 'Failed');
    return 'failed';
  }
}

// Renewal e-mail (called from quickRenewBook in Code.gs).
function sendBookRenewEmail_(memberId, bookId, bookTitle, newDue) {
  return mmSend_('renew', memberId, [{ bookId: bookId, title: bookTitle, issueDate: new Date(), dueDate: newDue }]);
}

// Run once from the editor to see a sample in your own inbox.
function testMemberMail() {
  var to = Session.getActiveUser().getEmail() || MM_.email;
  var m = mmBuild_('issue', 'M-SAMPLE', 'Sample Member', [{ bookId: 'SAMPLE-1', title: 'Sample Book Title', issueDate: new Date(), dueDate: new Date(Date.now() + 15 * 86400000) }]);
  MailApp.sendEmail({ to: to, subject: '[SAMPLE] ' + m.subject, body: m.body, htmlBody: m.htmlBody, name: MM_.shortName, replyTo: MM_.email });
  console.log('Sample sent to ' + to);
}

// ================================================================
// QUICK REPLY - answer a member's e-mail question in one click.
// Admin panel tab "Quick Reply". Same file for Kumta and Honavar
// (branch name, website and e-mail come from MM_ in MemberMail.gs).
// Every reply is logged in the Email_Log sheet (Category: Reply).
// ================================================================

var QR_TYPES_ = [
  ['available',  'Book is available',            true],
  ['issued',     'Book is issued to someone',    true],
  ['notinlib',   'Book not in our library',      false],
  ['keptaside',  'Book kept aside for you',      true],
  ['renewed',    'Renewal done',                 true],
  ['renewno',    'Renewal not possible',         true],
  ['membership', 'How to become a member',       false],
  ['cardready',  'Membership card ready',        false],
  ['lost',       'Lost or damaged book',         true],
  ['timings',    'Timings and holidays',         false],
  ['donation',   'Thanks for donating books',    false],
  ['feedback',   'Feedback / complaint received',false],
  ['search',     'How to search on website',     false],
  ['overdue',    'Gentle overdue reminder',      true],
  ['received',   'Message received (general)',   false],
  ['custom',     'My own message',               false]
];

/** List for the drop-down: [{id, label, needsBook}] */
function qrTypes() {
  return QR_TYPES_.map(function (t) { return { id: t[0], label: t[1], needsBook: t[2] }; });
}

/** Find a member by Library Card No or e-mail. Returns {found, id, name, email}. */
function qrFindMember(q) {
  q = String(q || '').trim();
  if (!q) return { found: false };
  if (q.indexOf('@') > 0) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
    var last = sh.getLastRow();
    if (last >= MEMBERS_START_ROW) {
      var rows = sh.getRange(MEMBERS_START_ROW, 1, last - MEMBERS_START_ROW + 1, MEMBERS_LAST_COL).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][MEMBERS_COL.email - 1] || '').trim().toLowerCase() === q.toLowerCase()) {
          return { found: true, id: normId_(rows[i][MEMBERS_COL.id - 1]), name: rows[i][MEMBERS_COL.name - 1], email: q };
        }
      }
    }
    return { found: false, email: q };            // not a member - still allowed to reply
  }
  var c = getMemberContact_(q);
  if (!c) return { found: false };
  return { found: true, id: normId_(q), name: c.name, email: c.email };
}

/** Book details + loan status from your sheets. */
function qrFindBook(acc) {
  acc = normId_(acc);
  if (!acc) return { found: false };
  var st = lookupBookLoanStatus(acc);
  if (!st.bookFound) return { found: false, id: acc };
  var b = mmBook_(acc);
  var out = { found: true, id: acc, title: st.bookTitle || b.title, author: b.author, issued: !!st.issued, dueDate: st.dueDate || '' };
  if (st.issued) {   // issue date + renew count for this loan
    try {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BI_SHEET);
      var last = biLastRealRow_(sh), n = Math.max(last - BI_START_ROW + 1, 0);
      var v = n ? sh.getRange(BI_START_ROW, 1, n, BI_LAST_COL).getValues() : [];
      for (var i = v.length - 1; i >= 0; i--) {
        if (normId_(v[i][BI_COL.bookId - 1]) === acc && String(v[i][BI_COL.status - 1]).trim().toUpperCase() === 'ISSUED') {
          out.issueDate = fmtDate_(v[i][BI_COL.issueDate - 1]); out.holderId = normId_(v[i][BI_COL.memberId - 1]); break;
        }
      }
    } catch (e) {}
  }
  return out;
}

function qrAddDays_(n) {
  var d = new Date(); var added = 0;
  while (added < n) { d.setDate(d.getDate() + 1); if (d.getDay() !== 1) added++; }   // skip Mondays (weekly holiday)
  return fmtDate_(d);
}

/** Builds subject + message for staff to check (and edit) before sending. */
function qrDraft(p) {
  p = p || {};
  var name = String(p.name || '').trim() || 'Reader';
  var bk = p.acc ? qrFindBook(p.acc) : { found: false };
  var T = bk.found ? bk.title : (p.title || '[book title]');
  var A = bk.found && bk.author ? ' by ' + bk.author : '';
  var ACC = bk.found ? ' (Acc No: ' + bk.id + ')' : '';
  var due = bk.found && bk.dueDate ? bk.dueDate : (p.date || '[date]');
  var site = MM_.website;
  var card = p.cardNo || '[card number]';
  var note = String(p.note || '').trim();
  var s = '', b = '';
  switch (p.type) {
    case 'available':
      s = 'Re: ' + T + ' - available';
      b = 'Thank you for your enquiry. The book "' + T + '"' + A + ACC + ' is available in our library now.\n\n' +
          'You can borrow it with your library card during library hours. If you wish, we can keep it aside for you for 2 working days - just reply "Please keep it".';
      if (bk.found && bk.issued) b = 'Note to staff: this book is currently ISSUED (due ' + bk.dueDate + '). Choose "Book is issued to someone" instead.\n\n' + b;
      break;
    case 'issued':
      s = 'Re: ' + T + ' - currently issued';
      b = 'The book "' + T + '"' + A + ACC + ' is currently issued to another reader and is expected back by ' + due + '.\n\n' +
          'We will keep it aside for you when it is returned and inform you by e-mail. You can also check its status anytime on our website (search the title - it shows Available or Issued):\n' + site;
      break;
    case 'notinlib':
      s = 'Re: ' + (p.title || 'your book request');
      b = 'Thank you for your request. Sorry, "' + (p.title || T) + '" is not available in our library at present.\n\n' +
          'We have noted it in our list of suggested books for future purchase.' + (note ? '\n\nMeanwhile, you may like these books available with us: ' + note : '') +
          '\n\nYou can search all our books on our website:\n' + site;
      note = '';
      break;
    case 'keptaside':
      s = T + ' - kept aside for you';
      b = 'As requested, the book "' + T + '"' + ACC + ' has been kept aside for you at the counter. Please collect it within 2 working days (by ' + qrAddDays_(2) + '), after which it will be returned to the shelf.';
      break;
    case 'renewed':
      s = 'Re: Renewal - ' + T;
      b = 'Your book "' + T + '"' + ACC + ' has been renewed. The new due date is ' + due + '.';
      break;
    case 'renewno':
      s = 'Re: Renewal - ' + T;
      b = 'Sorry, we are unable to renew "' + T + '"' + ACC + ' because ' + (note || 'it has reached the renewal limit') + '. Kindly return it on or before ' + due + '. You may borrow it again later if it is free.';
      note = '';
      break;
    case 'membership':
      s = 'Re: Library membership';
      b = 'Thank you for your interest in joining our library. To become a member:\n\n' +
          '1. Fill the online Membership Form on our website: ' + site + ' (Membership Form button). Keep your photo, Aadhaar and signature ready to upload.\n' +
          '2. Print the form we send you by e-mail and get it signed by a Gazetted Officer.\n' +
          '3. Bring the signed form, your documents and the membership fee to the library.\n\n' +
          'If you cannot fill the form online, you can collect a blank form at the library counter.';
      break;
    case 'cardready':
      s = 'Your library card is ready';
      b = 'Your membership has been approved. Your Library Card No is ' + card + '. Please collect your card from the library counter during library hours. You can start borrowing books right away.';
      break;
    case 'lost':
      s = 'Re: ' + T + ' - lost / damaged';
      b = 'Thank you for informing us about "' + T + '"' + ACC + '. As per library rules, a lost or badly damaged book must be replaced with a new copy of the same book, or its cost paid at the library. Please visit the library with your card so we can close the record.';
      break;
    case 'timings':
      s = 'Re: Library timings';
      b = 'Our library is open ' + MM_.timings + '.\n\nToday\'s open or closed status is shown at the top of our website:\n' + site;
      break;
    case 'donation':
      s = 'Thank you for your book donation';
      b = 'On behalf of the library, thank you for donating ' + (p.count || 'your') + ' books. They will be checked, entered in our register and placed on the shelves for all readers. Your generosity is greatly appreciated.';
      break;
    case 'feedback':
      s = 'Re: Your feedback';
      b = 'Thank you for writing to us. We have received your message and it has been noted.' + (note ? ' ' + note : '') + ' We value your feedback in improving our library.';
      note = '';
      break;
    case 'search':
      s = 'Re: Searching for books';
      b = 'You can search all our books from your phone, without signing in:\n\n1. Open ' + site + '\n2. Type the book title, author or subject in the search box.\n3. The result shows whether the book is Available or Issued.\n\nThe website also shows today\'s newspapers, magazines and new arrivals.';
      break;
    case 'overdue':
      s = 'Reminder - ' + T;
      b = 'This is a gentle reminder that the book "' + T + '"' + ACC + (bk.issueDate ? ', issued on ' + bk.issueDate + ',' : '') + ' was due on ' + due + '. Kindly return or renew it at your earliest convenience so other readers can also use it.\n\nIf you have already returned it, please ignore this message.';
      break;
    case 'received':
      s = 'Re: Your message';
      b = 'Thank you for your message. We have received it and will reply within 2 working days. For anything urgent, please visit the library during working hours.';
      break;
    default:
      s = p.subject || 'Message from ' + MM_.shortName;
      b = note || '';
      note = '';
  }
  if (note) b += '\n\n' + note;
  var body = 'Dear ' + name + ',\n\n' + b + '\n\n' + mmFooter_();
  return { subject: s, body: body, book: bk };
}

/** Sends the (possibly edited) reply. p: {to, name, memberId, subject, body, type} */
function qrSend(p) {
  p = p || {};
  var to = String(p.to || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) throw new Error('Enter a valid e-mail address.');
  if (!String(p.subject || '').trim() || !String(p.body || '').trim()) throw new Error('Subject and message cannot be empty.');
  MailApp.sendEmail({ to: to, subject: p.subject, body: p.body, name: MM_.shortName, replyTo: MM_.email });
  var label = (QR_TYPES_.filter(function (t) { return t[0] === p.type; })[0] || [0, 'Reply'])[1];
  logEmail_('Reply', label, p.memberId || '', p.name || '', to, p.subject, 'Sent');
  return { ok: true, left: MailApp.getRemainingDailyQuota() };
}

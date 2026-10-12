// ================================================================
// LIBRARY MODULES (Koha-style) - same file for BCL Kumta and BCL Honavar
//   1. Reservations (holds queue)
//   2. My Library - member login on the website (e-mail code),
//      my loans, renew online, history, reserve, suggest a book
//   3. Book Suggestions from members (with e-mails to the reader)
//   4. Reports
//   5. Book Lists (reading lists / book displays)
//
// Needs: Code.gs (your existing project) and MemberMail.gs (MM_ settings).
// First time: run setupLibraryModules once from the editor.
// Your existing sheets are only READ. Book Issue is changed only by the
// normal renew function when a member renews online.
// ================================================================

var LM_ = {
  RES_MAX_PER_MEMBER: 3,      // active reservations one member may have
  RES_HOLD_DAYS: 3,           // working days a ready book is kept aside
  OTP_MINUTES: 10,            // login code valid for
  SESSION_HOURS: 6            // member stays logged in for
};

var LM_SHEETS_ = {
  RES:   ['RESERVATIONS', ['Res ID', 'Placed On', 'Member ID', 'Member Name', 'Acc No', 'Title', 'Status', 'Ready On', 'Collect By', 'Source', 'Note', 'Closed On']],
  SUG:   ['SUGGESTIONS', ['Sug ID', 'Date', 'Member ID', 'Name', 'E-mail', 'Title', 'Author', 'Publisher', 'Reason', 'Status', 'Staff Note', 'Acc No']],
  LIST:  ['LISTS', ['List ID', 'Name', 'Description', 'Public', 'Created', 'Order']],
  ITEM:  ['LIST_ITEMS', ['List ID', 'Acc No', 'Note', 'Added']]
};

// ---------------------------------------------------------------- setup
function setupLibraryModules() {
  var made = [];
  Object.keys(LM_SHEETS_).forEach(function (k) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(LM_SHEETS_[k][0])) { lmSheet_(k); made.push(LM_SHEETS_[k][0]); }
  });
  var msg = made.length ? 'Created sheets: ' + made.join(', ') : 'All module sheets already exist. Nothing changed.';
  console.log(msg);
  return msg;
}

function lmSheet_(key) {
  var def = LM_SHEETS_[key];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(def[0]);
  if (!sh) {
    sh = ss.insertSheet(def[0]);
    sh.getRange(1, 1, 1, def[1].length).setValues([def[1]]).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// All data rows of a module sheet as arrays, each with ._row (sheet row number)
function lmRows_(key) {
  var sh = lmSheet_(key), last = sh.getLastRow(), n = LM_SHEETS_[key][1].length;
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, n).getValues().map(function (r, i) { r._row = i + 2; return r; })
    .filter(function (r) { return String(r[0]).trim() !== ''; });
}

function lmNextId_(key, prefix) {
  var max = 0;
  lmRows_(key).forEach(function (r) { var m = String(r[0]).match(/(\d+)$/); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  return prefix + ('0000' + (max + 1)).slice(-4);
}

function lmToday_() { var d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function lmAddWorkDays_(n) {                     // skips Mondays (weekly holiday)
  var d = lmToday_(), added = 0;
  while (added < n) { d.setDate(d.getDate() + 1); if (d.getDay() !== 1) added++; }
  return d;
}
function lmStr_(v) { return String(v == null ? '' : v).trim(); }
function lmClip_(v, n) { var s = lmStr_(v).replace(/[\r\n]+/g, ' ').slice(0, n || 200); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function lmDate_(v) { return v instanceof Date ? fmtDate_(v) : lmStr_(v); }
function lmParseDate_(s) {                       // 'yyyy-mm-dd' from <input type=date>
  var m = lmStr_(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}

// ---------------------------------------------------------------- shared reads (cached per run)
var LM_CACHE_ = {};
function lmCatalogue_() {
  if (LM_CACHE_.cat) return LM_CACHE_.cat;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Books Catalogue');
  var last = sh.getLastRow(), list = [], byId = {};
  if (last >= BOOKS_START_ROW) {
    sh.getRange(BOOKS_START_ROW, 1, last - BOOKS_START_ROW + 1, BOOKS_LAST_COL).getValues().forEach(function (r, i) {
      var id = normId_(r[BOOKS_COL.id - 1]); if (!id) return;
      var b = { id: id, title: lmStr_(r[BOOKS_COL.title - 1]), author: lmStr_(r[BOOKS_COL.author - 1]),
        genre: lmStr_(r[BOOKS_COL.genre - 1]), shelf: lmStr_(r[BOOKS_COL.shelfLocation - 1]),
        publisher: lmStr_(r[BOOKS_COL.publisher - 1]), year: lmStr_(r[BOOKS_COL.year - 1]),
        status: lmStr_(r[BOOKS_COL.status - 1]), row: BOOKS_START_ROW + i };
      list.push(b); byId[id] = b;
    });
  }
  LM_CACHE_.cat = { list: list, byId: byId };
  return LM_CACHE_.cat;
}

// Book Issue rows as objects (all history)
function lmIssues_() {
  if (LM_CACHE_.iss) return LM_CACHE_.iss;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BI_SHEET);
  var last = biLastRealRow_(sh), n = Math.max(last - BI_START_ROW + 1, 0);
  var rows = n ? sh.getRange(BI_START_ROW, 1, n, BI_LAST_COL).getValues() : [];
  var out = [], current = {};
  rows.forEach(function (r, i) {
    var o = { row: BI_START_ROW + i, issueId: r[BI_COL.issueId - 1], memberId: normId_(r[BI_COL.memberId - 1]),
      bookId: normId_(r[BI_COL.bookId - 1]), memberName: lmStr_(r[BI_COL.memberName - 1]), title: lmStr_(r[BI_COL.bookTitle - 1]),
      issueDate: r[BI_COL.issueDate - 1], dueDate: r[BI_COL.dueDate - 1], returnDate: r[BI_COL.returnDate - 1],
      status: lmStr_(r[BI_COL.status - 1]).toUpperCase(), fine: Number(r[BI_COL.fine - 1]) || 0, renewCount: Number(r[BI_COL.renewCount - 1]) || 0 };
    if (!o.bookId) return;
    out.push(o);
    if (o.status === 'ISSUED') current[o.bookId] = o;
  });
  LM_CACHE_.iss = { list: out, current: current };
  return LM_CACHE_.iss;
}

function lmMembers_() {
  if (LM_CACHE_.mem) return LM_CACHE_.mem;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var last = sh.getLastRow(), list = [], byId = {};
  if (last >= MEMBERS_START_ROW) {
    sh.getRange(MEMBERS_START_ROW, 1, last - MEMBERS_START_ROW + 1, MEMBERS_LAST_COL).getValues().forEach(function (r) {
      var id = normId_(r[MEMBERS_COL.id - 1]); if (!id) return;
      var m = { id: id, name: lmStr_(r[MEMBERS_COL.name - 1]), phone: lmStr_(r[MEMBERS_COL.phone - 1]), email: lmStr_(r[MEMBERS_COL.email - 1]),
        start: r[MEMBERS_COL.membStart - 1], end: r[MEMBERS_COL.membEnd - 1], status: lmStr_(r[MEMBERS_COL.status - 1]) };
      list.push(m); byId[id] = m;
    });
  }
  LM_CACHE_.mem = { list: list, byId: byId };
  return LM_CACHE_.mem;
}

function lmSettings_() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var v = s ? s.getRange('B4:B6').getValues() : [[0], [15], ['']];
  return { finePerDay: Number(v[0][0]) || 0, loanDays: Number(v[1][0]) || 15, maxRenew: parseInt(v[2][0], 10) };
}

function lmMail_(to, subject, text) {
  if (!to) return 'no-email';
  try {
    MailApp.sendEmail({ to: to, subject: subject + ' - ' + MM_.shortName, body: text + '\n\n' + mmFooter_(), name: MM_.shortName, replyTo: MM_.email });
    return 'sent';
  } catch (e) { return 'failed'; }
}

// ================================================================
// 1. RESERVATIONS
// ================================================================
var RES_ACTIVE_ = { 'WAITING': 1, 'READY': 1 };
var RC_ = { id: 0, placed: 1, mid: 2, mname: 3, acc: 4, title: 5, status: 6, ready: 7, collectBy: 8, source: 9, note: 10, closed: 11 };

function resActive_() {
  return lmRows_('RES').filter(function (r) { return RES_ACTIVE_[lmStr_(r[RC_.status]).toUpperCase()]; });
}
// active reservations for one book: READY first, then waiting in date order
function resQueue_(bookId) {
  bookId = normId_(bookId);
  return resActive_().filter(function (r) { return normId_(r[RC_.acc]) === bookId; }).sort(function (a, b) {
    var ra = lmStr_(a[RC_.status]).toUpperCase() === 'READY' ? 0 : 1, rb = lmStr_(b[RC_.status]).toUpperCase() === 'READY' ? 0 : 1;
    if (ra !== rb) return ra - rb;
    return new Date(a[RC_.placed]) - new Date(b[RC_.placed]);
  });
}
function resSet_(row, col, value) { lmSheet_('RES').getRange(row, col + 1).setValue(value); }

/** Place a reservation. source: 'Staff' | 'Website' */
function resPlace(memberId, bookId, source, note) {
  memberId = normId_(memberId); bookId = normId_(bookId);
  if (!memberId || !bookId) throw new Error('Card No and Acc No are both needed.');
  var mem = lmMembers_().byId[memberId];
  if (!mem) throw new Error('Card No ' + memberId + ' not found.');
  if (mem.end instanceof Date && mem.end < lmToday_()) throw new Error('Membership has expired on ' + fmtDate_(mem.end) + '. Please renew membership at the library.');
  var book = lmCatalogue_().byId[bookId];
  if (!book) throw new Error('Acc No ' + bookId + ' not found.');
  if (/discard|withdrawn|lost/i.test(book.status)) throw new Error('This book is not available for reservation (' + book.status + ').');
  var cur = lmIssues_().current[bookId];
  if (cur && cur.memberId === memberId) throw new Error('This book is already issued to you.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var act = resActive_();
    if (act.some(function (r) { return normId_(r[RC_.mid]) === memberId && normId_(r[RC_.acc]) === bookId; })) throw new Error('You have already reserved this book.');
    var mine = act.filter(function (r) { return normId_(r[RC_.mid]) === memberId; }).length;
    if (mine >= LM_.RES_MAX_PER_MEMBER) throw new Error('Only ' + LM_.RES_MAX_PER_MEMBER + ' reservations are allowed at a time.');
    var queueBefore = resQueue_(bookId).length;
    var id = lmNextId_('RES', 'R');
    var onShelf = !cur;
    lmSheet_('RES').appendRow([id, new Date(), memberId, mem.name, bookId, book.title, 'Waiting', '', '', source || 'Staff',
      lmClip_(note, 200) || (onShelf ? 'On shelf - staff to keep aside' : ''), '']);
    var pos = queueBefore + 1;
    lmMail_(mem.email, 'Reservation placed: ' + book.title,
      'Dear ' + mem.name + ',\n\nYour reservation has been placed.\n\nLibrary Card No: ' + memberId + '\nAcc No: ' + bookId + '\nTitle: ' + book.title +
      (book.author ? '\nAuthor: ' + book.author : '') + '\nYour position in the queue: ' + pos + '\n\n' +
      (onShelf && pos === 1 ? 'The book is on the shelf. Staff will keep it aside for you and inform you by e-mail.' :
        'We will e-mail you as soon as the book is returned and kept aside for you.'));
    return { resId: id, position: pos, onShelf: onShelf, title: book.title };
  } finally { lock.releaseLock(); }
}

function resCancel(resId, byMemberId) {
  var r = lmRows_('RES').filter(function (x) { return lmStr_(x[RC_.id]) === lmStr_(resId); })[0];
  if (!r) throw new Error('Reservation not found.');
  if (byMemberId && normId_(r[RC_.mid]) !== normId_(byMemberId)) throw new Error('This reservation is not yours.');
  if (!RES_ACTIVE_[lmStr_(r[RC_.status]).toUpperCase()]) throw new Error('This reservation is already closed.');
  var wasReady = lmStr_(r[RC_.status]).toUpperCase() === 'READY';
  resSet_(r._row, RC_.status, 'Cancelled'); resSet_(r._row, RC_.closed, new Date());
  if (wasReady) resPromote_(r[RC_.acc]);           // next reader gets it
  return { ok: true };
}

/** Staff pulled an on-shelf book and kept it aside -> Ready + e-mail. */
function resMarkReady(resId) {
  var r = lmRows_('RES').filter(function (x) { return lmStr_(x[RC_.id]) === lmStr_(resId); })[0];
  if (!r) throw new Error('Reservation not found.');
  if (lmIssues_().current[normId_(r[RC_.acc])]) throw new Error('This book is still issued to someone. It becomes Ready automatically when it is returned.');
  resMakeReady_(r);
  return { ok: true };
}

function resMakeReady_(r) {
  var by = lmAddWorkDays_(LM_.RES_HOLD_DAYS);
  resSet_(r._row, RC_.status, 'Ready'); resSet_(r._row, RC_.ready, new Date()); resSet_(r._row, RC_.collectBy, by);
  var mem = lmMembers_().byId[normId_(r[RC_.mid])] || {};
  lmMail_(mem.email, 'Your reserved book is ready: ' + r[RC_.title],
    'Dear ' + (mem.name || r[RC_.mname]) + ',\n\nGood news! The book you reserved is now kept aside for you at the counter.\n\n' +
    'Library Card No: ' + r[RC_.mid] + '\nAcc No: ' + r[RC_.acc] + '\nTitle: ' + r[RC_.title] + '\nPlease collect it by: ' + fmtDate_(by) +
    '\n\nAfter this date it will be given to the next reader.');
  return by;
}

// next waiting reader (if the book is on the shelf) becomes Ready
function resPromote_(bookId) {
  if (lmIssues_().current[normId_(bookId)]) return null;
  var q = resQueue_(bookId);
  if (!q.length || lmStr_(q[0][RC_.status]).toUpperCase() === 'READY') return null;
  var by = resMakeReady_(q[0]);
  return { memberId: normId_(q[0][RC_.mid]), name: q[0][RC_.mname], collectBy: fmtDate_(by) };
}

// ---- hooks called from Code.gs (no locks here: the caller already holds one)
/** After a book is returned. Returns '' or a staff notice text. */
function resOnReturn_(bookId) {
  try {
    LM_CACHE_ = {};                                  // Book Issue just changed
    var q = resQueue_(bookId);
    if (!q.length) return '';
    var top = q[0];
    if (lmStr_(top[RC_.status]).toUpperCase() !== 'READY') resMakeReady_(top);
    return '  \u26a0 RESERVED - keep aside for ' + top[RC_.mname] + ' (Card ' + top[RC_.mid] + ')';
  } catch (e) { return ''; }
}
/** Before issuing. Returns '' if fine, or the reason it must not be issued. */
function resBlocker_(bookId, memberId) {
  try {
    var q = resQueue_(bookId);
    if (!q.length) return '';
    var top = q[0];
    if (normId_(top[RC_.mid]) === normId_(memberId)) return '';
    return 'Reserved for ' + top[RC_.mname] + ' (Card ' + top[RC_.mid] + ')' +
      (lmStr_(top[RC_.status]).toUpperCase() === 'READY' ? ' until ' + lmDate_(top[RC_.collectBy]) : ' - waiting in queue');
  } catch (e) { return ''; }
}
/** After a successful issue: close that member's reservation for this book. */
function resOnIssue_(bookId, memberId) {
  try {
    resActive_().forEach(function (r) {
      if (normId_(r[RC_.acc]) === normId_(bookId) && normId_(r[RC_.mid]) === normId_(memberId)) {
        resSet_(r._row, RC_.status, 'Collected'); resSet_(r._row, RC_.closed, new Date());
      }
    });
  } catch (e) {}
}
/** Before renewing: someone else waiting? */
function resRenewBlocker_(bookId, memberId) {
  try {
    var other = resQueue_(bookId).filter(function (r) { return normId_(r[RC_.mid]) !== normId_(memberId); });
    return other.length ? 'Another reader has reserved this book - it cannot be renewed. Please return it by the due date.' : '';
  } catch (e) { return ''; }
}
/** Daily: ready books not collected in time -> Expired, next reader gets it. */
function resDailyCheck_() {
  var today = lmToday_(), n = 0;
  resActive_().forEach(function (r) {
    if (lmStr_(r[RC_.status]).toUpperCase() === 'READY' && r[RC_.collectBy] instanceof Date && r[RC_.collectBy] < today) {
      resSet_(r._row, RC_.status, 'Expired'); resSet_(r._row, RC_.closed, new Date()); n++;
      LM_CACHE_ = {};
      resPromote_(r[RC_.acc]);
    }
  });
  return n;
}

/** Admin tab list. filter: 'active' | 'all' */
function resList(filter) {
  var iss = lmIssues_().current;
  var rows = lmRows_('RES');
  if (filter !== 'all') rows = rows.filter(function (r) { return RES_ACTIVE_[lmStr_(r[RC_.status]).toUpperCase()]; });
  var posMap = {};
  return rows.reverse().map(function (r) {
    var acc = normId_(r[RC_.acc]), cur = iss[acc], st = lmStr_(r[RC_.status]);
    var pos = '';
    if (RES_ACTIVE_[st.toUpperCase()]) { var q = resQueue_(acc); for (var i = 0; i < q.length; i++) if (lmStr_(q[i][RC_.id]) === lmStr_(r[RC_.id])) pos = i + 1; }
    return { id: lmStr_(r[RC_.id]), placed: lmDate_(r[RC_.placed]), memberId: lmStr_(r[RC_.mid]), name: lmStr_(r[RC_.mname]), acc: acc,
      title: lmStr_(r[RC_.title]), status: st, collectBy: lmDate_(r[RC_.collectBy]), source: lmStr_(r[RC_.source]), note: lmStr_(r[RC_.note]),
      position: pos, bookNow: cur ? 'Issued to ' + cur.memberId + ', due ' + lmDate_(cur.dueDate) : 'On shelf' };
  });
}

// ================================================================
// 2. MY LIBRARY - member login + account (called from the website)
// ================================================================
function mpMask_(email) { var p = email.split('@'); return p[0].slice(0, 2) + '***@' + (p[1] || ''); }

function mpRequestCode(cardNo) {
  cardNo = normId_(cardNo);
  var mem = lmMembers_().byId[cardNo];
  if (!mem) throw new Error('Library Card No not found. Please check the number.');
  if (!/@/.test(mem.email)) throw new Error('No e-mail is registered for this card. Please ask the library staff to add your e-mail.');
  var cache = CacheService.getScriptCache();
  var cnt = Number(cache.get('mp_req_' + cardNo) || 0);
  if (cnt >= 3) throw new Error('Too many codes requested. Please try again after 30 minutes.');
  cache.put('mp_req_' + cardNo, String(cnt + 1), 1800);
  var code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('mp_code_' + cardNo, code + '|0', LM_.OTP_MINUTES * 60);
  MailApp.sendEmail({ to: mem.email, subject: 'Your login code: ' + code + ' - ' + MM_.shortName, name: MM_.shortName, replyTo: MM_.email,
    body: 'Dear ' + mem.name + ',\n\nYour code to log in to My Library is:\n\n    ' + code + '\n\nIt is valid for ' + LM_.OTP_MINUTES +
      ' minutes. If you did not ask for this code, you can ignore this e-mail.\n\n' + mmFooter_() });
  return { sentTo: mpMask_(mem.email) };
}

function mpVerify(cardNo, code) {
  cardNo = normId_(cardNo); code = lmStr_(code);
  var cache = CacheService.getScriptCache();
  var v = cache.get('mp_code_' + cardNo);
  if (!v) throw new Error('The code has expired. Please ask for a new code.');
  var parts = v.split('|'), tries = Number(parts[1]) || 0;
  if (tries >= 5) { cache.remove('mp_code_' + cardNo); throw new Error('Too many wrong tries. Please ask for a new code.'); }
  if (parts[0] !== code) { cache.put('mp_code_' + cardNo, parts[0] + '|' + (tries + 1), LM_.OTP_MINUTES * 60); throw new Error('Wrong code. Please check the e-mail and try again.'); }
  cache.remove('mp_code_' + cardNo);
  var token = Utilities.getUuid();
  cache.put('mp_tok_' + token, cardNo, LM_.SESSION_HOURS * 3600);
  return { token: token, name: lmMembers_().byId[cardNo].name };
}

function mpCard_(token) {
  var c = token ? CacheService.getScriptCache().get('mp_tok_' + token) : null;
  if (!c) throw new Error('LOGIN_EXPIRED');
  return c;
}

function mpLogout(token) { try { CacheService.getScriptCache().remove('mp_tok_' + token); } catch (e) {} return { ok: true }; }

function mpMyAccount(token) {
  var card = mpCard_(token);
  var mem = lmMembers_().byId[card] || { name: '', email: '' };
  var set = lmSettings_(), today = lmToday_();
  var cat = lmCatalogue_().byId, iss = lmIssues_().list;
  var loans = [], history = [];
  iss.forEach(function (o) {
    if (o.memberId !== card) return;
    var b = cat[o.bookId] || {};
    if (o.status === 'ISSUED') {
      var overdue = o.dueDate instanceof Date && o.dueDate < today;
      var reason = '';
      if (overdue) reason = 'Overdue - please return it at the library';
      else if (!isNaN(set.maxRenew) && o.renewCount >= set.maxRenew) reason = 'Renewal limit reached';
      else reason = resRenewBlocker_(o.bookId, card) ? 'Reserved by another reader' : '';
      loans.push({ acc: o.bookId, title: o.title || b.title, author: b.author || '', issueDate: lmDate_(o.issueDate), dueDate: lmDate_(o.dueDate),
        overdue: overdue, renewals: o.renewCount, canRenew: !reason, reason: reason });
    } else if (o.status === 'RETURNED') {
      history.push({ acc: o.bookId, title: o.title || b.title, author: b.author || '', issueDate: lmDate_(o.issueDate), returnDate: lmDate_(o.returnDate) });
    }
  });
  history.reverse();
  var res = lmRows_('RES').filter(function (r) { return normId_(r[RC_.mid]) === card; }).reverse().slice(0, 20).map(function (r) {
    var st = lmStr_(r[RC_.status]), pos = '';
    if (RES_ACTIVE_[st.toUpperCase()]) { var q = resQueue_(r[RC_.acc]); for (var i = 0; i < q.length; i++) if (lmStr_(q[i][RC_.id]) === lmStr_(r[RC_.id])) pos = i + 1; }
    return { id: lmStr_(r[RC_.id]), acc: lmStr_(r[RC_.acc]), title: lmStr_(r[RC_.title]), placed: lmDate_(r[RC_.placed]), status: st,
      collectBy: lmDate_(r[RC_.collectBy]), position: pos, active: !!RES_ACTIVE_[st.toUpperCase()] };
  });
  var sug = lmRows_('SUG').filter(function (r) { return normId_(r[2]) === card; }).reverse().slice(0, 20).map(function (r) {
    return { title: lmStr_(r[5]), author: lmStr_(r[6]), date: lmDate_(r[1]), status: lmStr_(r[9]) };
  });
  return { card: card, name: mem.name, email: mem.email ? mpMask_(mem.email) : '', membEnd: lmDate_(mem.end),
    expired: mem.end instanceof Date && mem.end < today, loans: loans, history: history.slice(0, 100), historyCount: history.length,
    reservations: res, suggestions: sug, maxRes: LM_.RES_MAX_PER_MEMBER };
}

function mpRenew(token, acc) {
  var card = mpCard_(token); acc = normId_(acc);
  var o = lmIssues_().current[acc];
  if (!o || o.memberId !== card) throw new Error('This book is not issued to you.');
  if (o.dueDate instanceof Date && o.dueDate < lmToday_()) throw new Error('This book is overdue. Please return it at the library.');
  var block = resRenewBlocker_(acc, card); if (block) throw new Error(block);
  var r = quickRenewBook(card, acc);
  return { ok: true, newDueDate: r.newDueDate };
}

function mpReserve(token, acc) { var card = mpCard_(token); return resPlace(card, acc, 'Website', ''); }
function mpCancelReservation(token, resId) { var card = mpCard_(token); return resCancel(resId, card); }

function mpSuggest(token, p) {
  var card = mpCard_(token), mem = lmMembers_().byId[card] || {};
  p = p || {}; p.memberId = card; p.name = mem.name; p.email = mem.email;
  return sugAdd(p);
}

/** Public book search for the website (no login). */
function mpSearch(q) {
  q = lmStr_(q).toLowerCase();
  if (q.length < 2) throw new Error('Type at least 2 letters.');
  var words = q.split(/\s+/).filter(Boolean);
  var cat = lmCatalogue_().list, cur = lmIssues_().current, resCount = {};
  resActive_().forEach(function (r) { var a = normId_(r[RC_.acc]); resCount[a] = (resCount[a] || 0) + 1; });
  var out = [], total = 0;
  for (var i = 0; i < cat.length; i++) {
    var b = cat[i];
    if (/discard|withdrawn/i.test(b.status)) continue;
    var hay = (b.title + ' ' + b.author + ' ' + b.id + ' ' + b.publisher).toLowerCase();
    if (!words.every(function (w) { return hay.indexOf(w) >= 0; })) continue;
    total++;
    if (out.length < 40) {
      var c = cur[b.id];
      out.push({ acc: b.id, title: b.title, author: b.author, shelf: b.shelf, issued: !!c, dueDate: c ? lmDate_(c.dueDate) : '', reserved: resCount[b.id] || 0 });
    }
  }
  return { results: out, total: total };
}

// ================================================================
// 3. BOOK SUGGESTIONS (from members, with e-mails)
//   New -> Accepted (reader e-mailed: we will try to get it)
//       -> Arrived  (reader e-mailed: it is here, reserve it)
//       -> Rejected (reader e-mailed with your reason)
// ================================================================
var SC_ = { id: 0, date: 1, mid: 2, name: 3, email: 4, title: 5, author: 6, pub: 7, reason: 8, status: 9, note: 10, acc: 11 };

function sugAdd(p) {
  p = p || {};
  var title = lmClip_(p.title, 200);
  if (title.length < 2) throw new Error('Please type the book title.');
  var mid = normId_(p.memberId), name = lmClip_(p.name, 80), email = lmClip_(p.email, 120);
  if (mid && (!name || !email)) { var m = lmMembers_().byId[mid]; if (m) { name = name || m.name; email = email || m.email; } }
  var id = lmNextId_('SUG', 'S');
  lmSheet_('SUG').appendRow([id, new Date(), mid, name, email, title, lmClip_(p.author, 120),
    lmClip_(p.publisher, 120), lmClip_(p.reason, 300), 'New', '', '']);
  if (email) lmMail_(email, 'We received your book suggestion',
    'Dear ' + (name || 'Reader') + ',\n\nThank you for suggesting "' + title + '"' + (p.author ? ' by ' + lmClip_(p.author, 120) : '') +
    '. Our staff will look at it and inform you by e-mail.');
  return { id: id };
}

/** status: 'Accepted' | 'Arrived' | 'Rejected'. acc = Acc No (for Arrived, optional). */
function sugSetStatus(id, status, note, acc) {
  var r = lmRows_('SUG').filter(function (x) { return lmStr_(x[SC_.id]) === lmStr_(id); })[0];
  if (!r) throw new Error('Suggestion not found.');
  if (['Accepted', 'Arrived', 'Rejected'].indexOf(status) < 0) throw new Error('Unknown status.');
  var sh = lmSheet_('SUG');
  sh.getRange(r._row, SC_.status + 1).setValue(status);
  if (note) sh.getRange(r._row, SC_.note + 1).setValue(lmClip_(note, 300));
  acc = normId_(acc);
  if (acc) sh.getRange(r._row, SC_.acc + 1).setValue(acc);
  var email = lmStr_(r[SC_.email]), name = lmStr_(r[SC_.name]) || 'Reader', title = lmStr_(r[SC_.title]);
  var mail = 'no-email';
  if (status === 'Accepted') mail = lmMail_(email, 'Your book suggestion is accepted',
    'Dear ' + name + ',\n\nGood news! We have accepted your suggestion "' + title + '". We will try to get it for the library and e-mail you when it arrives.' +
    (note ? '\n\nNote from the library: ' + note : ''));
  if (status === 'Arrived') mail = lmMail_(email, 'The book you suggested has arrived: ' + title,
    'Dear ' + name + ',\n\nThe book you suggested, "' + title + '", is now in our library' + (acc ? ' (Acc No: ' + acc + ')' : '') + '.' +
    '\n\nYou can reserve it in My Library on our website, or borrow it at the counter.' + (note ? '\n\nNote from the library: ' + note : ''));
  if (status === 'Rejected') mail = lmMail_(email, 'About your book suggestion',
    'Dear ' + name + ',\n\nThank you for suggesting "' + title + '". We are unable to get it at present' + (note ? ' (' + note + ')' : '') +
    '. We will keep it in mind for the future.');
  return { ok: true, mail: mail };
}

function sugList(filter) {
  var rows = lmRows_('SUG').reverse().map(function (r) {
    return { id: lmStr_(r[SC_.id]), date: lmDate_(r[SC_.date]), memberId: lmStr_(r[SC_.mid]), name: lmStr_(r[SC_.name]), email: lmStr_(r[SC_.email]),
      title: lmStr_(r[SC_.title]), author: lmStr_(r[SC_.author]), publisher: lmStr_(r[SC_.pub]), reason: lmStr_(r[SC_.reason]),
      status: lmStr_(r[SC_.status]), note: lmStr_(r[SC_.note]), acc: lmStr_(r[SC_.acc]) };
  });
  if (filter === 'open') rows = rows.filter(function (s) { return s.status === 'New' || s.status === 'Accepted'; });
  return rows;
}

// ================================================================
// 4. REPORTS
// ================================================================
var RPT_LIST_ = [
  ['top_books', 'Most issued books (copies)', true],
  ['top_titles', 'Most issued titles (all copies together)', true],
  ['top_members', 'Most active members', true],
  ['monthly', 'Monthly statistics (issues, returns, readers, new members)', true],
  ['daily', 'Day-wise issues and returns', true],
  ['genre', 'Subject / genre-wise issues', true],
  ['overdue', 'Overdue books right now', false],
  ['inactive', 'Inactive members (no book taken since the From date)', true],
  ['new_members', 'New members joined', true],
  ['expiring', 'Memberships ending between the dates', true],
  ['never_issued', 'Books never issued since the From date', true],
  ['reservations', 'Reservations placed', true],
  ['suggestions', 'Book suggestions received', true]
];
function rptTypes() { return RPT_LIST_.map(function (r) { return { id: r[0], label: r[1], usesDates: r[2] }; }); }

function rptRun(type, fromS, toS, topN) {
  var from = lmParseDate_(fromS) || new Date(2000, 0, 1), to = lmParseDate_(toS) || lmToday_();
  var toEnd = new Date(to.getFullYear(), to.getMonth(), to.getDate(), 23, 59, 59);
  var inR = function (d) { return d instanceof Date && d >= from && d <= toEnd; };
  var N = Math.max(1, Math.min(1000, parseInt(topN, 10) || 50));
  var label = (RPT_LIST_.filter(function (r) { return r[0] === type; })[0] || ['', 'Report'])[1];
  var cat = lmCatalogue_().byId, iss = lmIssues_().list, mem = lmMembers_();
  var cols = [], rows = [], note = '';
  var count = function (keyFn) { var m = {}; iss.forEach(function (o) { if (inR(o.issueDate)) { var k = keyFn(o); m[k] = (m[k] || 0) + 1; } }); return m; };
  var sortTop = function (m) { return Object.keys(m).sort(function (a, b) { return m[b] - m[a]; }).slice(0, N); };
  switch (type) {
    case 'top_books':
      cols = ['#', 'Acc No', 'Title', 'Author', 'Times issued'];
      var m1 = count(function (o) { return o.bookId; });
      sortTop(m1).forEach(function (k, i) { var b = cat[k] || {}; rows.push([i + 1, k, b.title || '', b.author || '', m1[k]]); });
      break;
    case 'top_titles':
      cols = ['#', 'Title', 'Author', 'Times issued'];
      var m2 = count(function (o) { var b = cat[o.bookId] || {}; return ((b.title || o.title) + '|' + (b.author || '')).toLowerCase(); });
      var show = {}; iss.forEach(function (o) { var b = cat[o.bookId] || {}; show[((b.title || o.title) + '|' + (b.author || '')).toLowerCase()] = [b.title || o.title, b.author || '']; });
      sortTop(m2).forEach(function (k, i) { rows.push([i + 1, show[k][0], show[k][1], m2[k]]); });
      break;
    case 'top_members':
      cols = ['#', 'Card No', 'Name', 'Phone', 'Books taken'];
      var m3 = count(function (o) { return o.memberId; });
      sortTop(m3).forEach(function (k, i) { var p = mem.byId[k] || {}; rows.push([i + 1, k, p.name || '', p.phone || '', m3[k]]); });
      break;
    case 'monthly':
    case 'daily':
      cols = type === 'monthly' ? ['Month', 'Issues', 'Returns', 'Different readers', 'New members'] : ['Date', 'Issues', 'Returns', 'Different readers'];
      var fmt = type === 'monthly' ? 'yyyy-MM' : 'yyyy-MM-dd', tz = Session.getScriptTimeZone(), agg = {};
      var slot = function (d) { var k = Utilities.formatDate(d, tz, fmt); agg[k] = agg[k] || { i: 0, r: 0, who: {}, n: 0 }; return agg[k]; };
      iss.forEach(function (o) {
        if (inR(o.issueDate)) { var s = slot(o.issueDate); s.i++; s.who[o.memberId] = 1; }
        if (inR(o.returnDate)) slot(o.returnDate).r++;
      });
      if (type === 'monthly') mem.list.forEach(function (p) { if (inR(p.start)) slot(p.start).n++; });
      Object.keys(agg).sort().forEach(function (k) {
        var s = agg[k], row = [type === 'monthly' ? Utilities.formatDate(lmParseDate_(k + '-01'), tz, 'MMM yyyy') : fmtDate_(lmParseDate_(k)), s.i, s.r, Object.keys(s.who).length];
        if (type === 'monthly') row.push(s.n);
        rows.push(row);
      });
      var tot = rows.reduce(function (t, r) { return [t[0] + r[1], t[1] + r[2]]; }, [0, 0]);
      note = 'Total issues: ' + tot[0] + ', total returns: ' + tot[1];
      break;
    case 'genre':
      cols = ['Subject / Genre', 'Times issued'];
      var m4 = count(function (o) { var b = cat[o.bookId] || {}; return b.genre || '(no genre)'; });
      Object.keys(m4).sort(function (a, b) { return m4[b] - m4[a]; }).forEach(function (k) { rows.push([k, m4[k]]); });
      break;
    case 'overdue':
      cols = ['Acc No', 'Title', 'Card No', 'Name', 'Phone', 'Due date', 'Days overdue'];
      var today = lmToday_();
      iss.filter(function (o) { return o.status === 'ISSUED' && o.dueDate instanceof Date && o.dueDate < today; })
        .sort(function (a, b) { return a.dueDate - b.dueDate; })
        .forEach(function (o) { var p = mem.byId[o.memberId] || {}; rows.push([o.bookId, o.title, o.memberId, p.name || o.memberName, p.phone || '', fmtDate_(o.dueDate), Math.round((today - o.dueDate) / 86400000)]); });
      break;
    case 'inactive':
      cols = ['Card No', 'Name', 'Phone', 'Joined', 'Last book taken'];
      var lastIss = {};
      iss.forEach(function (o) { if (o.issueDate instanceof Date && (!lastIss[o.memberId] || o.issueDate > lastIss[o.memberId])) lastIss[o.memberId] = o.issueDate; });
      mem.list.forEach(function (p) {
        var li = lastIss[p.id];
        if ((!li || li < from) && !(p.start instanceof Date && p.start > from)) rows.push([p.id, p.name, p.phone, lmDate_(p.start), li ? fmtDate_(li) : 'Never']);
      });
      break;
    case 'new_members':
      cols = ['Card No', 'Name', 'Phone', 'E-mail', 'Joined', 'Valid till'];
      mem.list.filter(function (p) { return inR(p.start); }).sort(function (a, b) { return a.start - b.start; })
        .forEach(function (p) { rows.push([p.id, p.name, p.phone, p.email, fmtDate_(p.start), lmDate_(p.end)]); });
      break;
    case 'expiring':
      cols = ['Card No', 'Name', 'Phone', 'E-mail', 'Valid till'];
      mem.list.filter(function (p) { return inR(p.end); }).sort(function (a, b) { return a.end - b.end; })
        .forEach(function (p) { rows.push([p.id, p.name, p.phone, p.email, fmtDate_(p.end)]); });
      break;
    case 'never_issued':
      cols = ['Acc No', 'Title', 'Author', 'Shelf'];
      var used = {}; iss.forEach(function (o) { if (!(o.issueDate instanceof Date) || o.issueDate >= from) used[o.bookId] = 1; });
      var all = lmCatalogue_().list.filter(function (b) { return !used[b.id] && !/discard|withdrawn/i.test(b.status); });
      all.slice(0, N).forEach(function (b) { rows.push([b.id, b.title, b.author, b.shelf]); });
      note = all.length + ' book(s) not issued since ' + fmtDate_(from) + (all.length > N ? ' - showing the first ' + N + ' (change "How many")' : '');
      break;
    case 'reservations':
      cols = ['Res ID', 'Placed on', 'Card No', 'Name', 'Acc No', 'Title', 'Status', 'Source'];
      lmRows_('RES').filter(function (r) { return inR(r[RC_.placed]); })
        .forEach(function (r) { rows.push([r[RC_.id], lmDate_(r[RC_.placed]), r[RC_.mid], r[RC_.mname], r[RC_.acc], r[RC_.title], r[RC_.status], r[RC_.source]]); });
      break;
    case 'suggestions':
      cols = ['Sug ID', 'Date', 'Card No', 'Name', 'Title', 'Author', 'Status'];
      lmRows_('SUG').filter(function (r) { return inR(r[SC_.date]); })
        .forEach(function (r) { rows.push([r[SC_.id], lmDate_(r[SC_.date]), r[SC_.mid], r[SC_.name], r[SC_.title], r[SC_.author], r[SC_.status]]); });
      break;
    default: throw new Error('Unknown report.');
  }
  if (!note) note = rows.length + ' row(s)';
  return { title: label, period: fmtDate_(from) + ' to ' + fmtDate_(to), columns: cols, rows: rows, note: note };
}

/** Saves a report as a new sheet tab in this spreadsheet. Returns the link. */
function rptToSheet(type, fromS, toS, topN) {
  var r = rptRun(type, fromS, toS, topN);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = ('RPT ' + r.title).slice(0, 60) + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM HHmm');
  var sh = ss.insertSheet(name);
  sh.getRange(1, 1).setValue(r.title + ' (' + r.period + ')').setFontWeight('bold');
  sh.getRange(2, 1).setValue(r.note);
  sh.getRange(4, 1, 1, r.columns.length).setValues([r.columns]).setFontWeight('bold').setBackground('#e8eef8');
  if (r.rows.length) sh.getRange(5, 1, r.rows.length, r.columns.length).setValues(r.rows);
  sh.autoResizeColumns(1, r.columns.length);
  return { url: ss.getUrl() + '#gid=' + sh.getSheetId(), name: name };
}

// ================================================================
// 5. BOOK LISTS (reading lists / book displays)
// ================================================================
function listAll() {
  var counts = {};
  lmRows_('ITEM').forEach(function (r) { var k = lmStr_(r[0]); counts[k] = (counts[k] || 0) + 1; });
  return lmRows_('LIST').map(function (r) {
    return { id: lmStr_(r[0]), name: lmStr_(r[1]), desc: lmStr_(r[2]), isPublic: /^y/i.test(lmStr_(r[3])), created: lmDate_(r[4]), order: Number(r[5]) || 0, count: counts[lmStr_(r[0])] || 0 };
  }).sort(function (a, b) { return a.order - b.order || (a.name > b.name ? 1 : -1); });
}

function listSave(p) {
  p = p || {};
  var name = lmClip_(p.name, 100); if (!name) throw new Error('List name is needed.');
  var sh = lmSheet_('LIST');
  if (p.id) {
    var r = lmRows_('LIST').filter(function (x) { return lmStr_(x[0]) === lmStr_(p.id); })[0];
    if (!r) throw new Error('List not found.');
    sh.getRange(r._row, 2, 1, 3).setValues([[name, lmClip_(p.desc, 300), p.isPublic ? 'Yes' : 'No']]);
    if (p.order != null) sh.getRange(r._row, 6).setValue(Number(p.order) || 0);
    return { id: p.id };
  }
  var id = lmNextId_('LIST', 'L');
  sh.appendRow([id, name, lmClip_(p.desc, 300), p.isPublic ? 'Yes' : 'No', new Date(), Number(p.order) || 0]);
  return { id: id };
}

function listDelete(id) {
  var sh = lmSheet_('LIST'), it = lmSheet_('ITEM');
  lmRows_('ITEM').filter(function (r) { return lmStr_(r[0]) === lmStr_(id); }).reverse().forEach(function (r) { it.deleteRow(r._row); });
  var r = lmRows_('LIST').filter(function (x) { return lmStr_(x[0]) === lmStr_(id); })[0];
  if (r) sh.deleteRow(r._row);
  return { ok: true };
}

function listItems(id) {
  var cat = lmCatalogue_().byId, cur = lmIssues_().current;
  return lmRows_('ITEM').filter(function (r) { return lmStr_(r[0]) === lmStr_(id); }).map(function (r) {
    var acc = normId_(r[1]), b = cat[acc] || {}, c = cur[acc];
    return { acc: acc, title: b.title || '(not in catalogue)', author: b.author || '', shelf: b.shelf || '', note: lmStr_(r[2]),
      issued: !!c, dueDate: c ? lmDate_(c.dueDate) : '' };
  });
}

function listAddItems(id, accText, note) {
  if (!lmRows_('LIST').some(function (x) { return lmStr_(x[0]) === lmStr_(id); })) throw new Error('List not found.');
  var have = {}; lmRows_('ITEM').forEach(function (r) { if (lmStr_(r[0]) === lmStr_(id)) have[normId_(r[1])] = 1; });
  var cat = lmCatalogue_().byId, added = [], skipped = [];
  String(accText || '').split(/[\s,]+/).map(normId_).filter(Boolean).forEach(function (a) {
    if (have[a]) { skipped.push(a + ' (already in list)'); return; }
    if (!cat[a]) { skipped.push(a + ' (not found)'); return; }
    have[a] = 1; added.push([id, a, lmClip_(note, 200), new Date()]);
  });
  if (added.length) { var sh = lmSheet_('ITEM'); sh.getRange(sh.getLastRow() + 1, 1, added.length, 4).setValues(added); }
  return { added: added.length, skipped: skipped };
}

function listRemoveItem(id, acc) {
  var sh = lmSheet_('ITEM');
  lmRows_('ITEM').filter(function (r) { return lmStr_(r[0]) === lmStr_(id) && normId_(r[1]) === normId_(acc); }).reverse()
    .forEach(function (r) { sh.deleteRow(r._row); });
  return { ok: true };
}

/** Public lists with their books, for the website (no login). */
function mpLists() {
  return listAll().filter(function (l) { return l.isPublic && l.count; }).map(function (l) {
    l.items = listItems(l.id).slice(0, 60); return l;
  });
}

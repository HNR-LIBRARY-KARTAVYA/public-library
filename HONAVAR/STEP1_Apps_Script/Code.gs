// ================================================================
// BCL_HNR_EXPANDED - Stack Verification, Barcode & Card System
// Branch Central Library, Honavar
// ================================================================
// Built ON TOP OF your existing live workbook - does NOT touch
// Book Issue, Dashboard, Awards & Statistics, or any existing formula.
// Points directly at the REAL 'Books Catalogue' and 'Members' sheets.
//
// ASSUMPTIONS - please confirm/correct:
//  1. 'Books Catalogue' data starts at ROW 4 (rows 1-3 are title/note/header).
//  2. 'Members' data starts at ROW 4 (rows 1-3 are title/note/header).
//  3. Barcode value = Book ID (Acc. No.) - Books Catalogue col A.
//  4. Member Card "category" line shows STATUS (Active/Closed) - there is
//     no Student/Public-style field in your real Members sheet.
//  5. Member Card "Valid upto" uses the REAL Membership End date (col F)
//     per member - not a fixed settings value like the earlier sample.
//  6. Photos: Drive folder (one file per Member ID / Book ID), same
//     convention as your other systems - folder IDs go in new SETTINGS
//     rows (see MANUAL).
//  7. New columns you already added to Books Catalogue (H-L: CALL NUMBER,
//     KARWAR NUMBER, PUBLICATION YEAR, NUMBER OF PAGES, PRICE) are used
//     directly on the Book Card.
//  8. No Gate Entry Card in this build - no non-member/visitor sheet
//     exists in this workbook yet.
// ================================================================

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function norm_(s) {
  return String(s || '').trim().toUpperCase().replace(/\s/g, '');
}

function fmtDate_(v) {
  if (!(v instanceof Date)) return String(v || '');
  return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
}

// Same idea as fmtDate_ but as yyyy-MM-dd, the format an HTML <input
// type="date"> box needs to show an existing value pre-filled.
function isoDate_(v) {
  if (!(v instanceof Date) || isNaN(v.getTime())) return '';
  return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ---- Fixed column layouts (1-indexed), matching your REAL sheets ------

var BOOKS_COL = {
  id: 1, title: 2, author: 3, genre: 4, totalCopies: 5, shelfLocation: 6,
  remarks: 7, callno: 8, karwarno: 9, year: 10, pages: 11, price: 12,
  publisher: 13, place: 14, barcode: 15, status: 16
};
var BOOKS_START_ROW = 4;
var BOOKS_LAST_COL = 16;

// ================================================================
// LIVE STATUS SYNC - keeps the Books Catalogue's own Status column
// (col 16) matching reality the instant a book is issued, returned,
// or discarded - not just when Public Book Search derives it live.
// This is what the Accession Register / Authority Report screens
// read directly, so without this they'd show stale text forever.
//
// Rule: a book already marked "Discarded" is never silently flipped
// back to Available/Issued by a stray issue/return call - discarding
// is the one status that must be undone on purpose, not overwritten
// as a side effect.
// ================================================================

// Single-book lookup: scans once, stops at first match. Fine for the
// single-book actions (Quick Issue, Quick Return) that call this.
function findBooksCatalogueRow_(booksSheet, bookIdNorm) {
  var last = booksSheet.getLastRow();
  var numRows = Math.max(last - BOOKS_START_ROW + 1, 0);
  if (!numRows) return -1;
  var ids = booksSheet.getRange(BOOKS_START_ROW, BOOKS_COL.id, numRows, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (norm_(ids[i][0]) === bookIdNorm) return BOOKS_START_ROW + i;
  }
  return -1;
}

// Bulk lookup: ONE scan of the whole column, returns a {id: row} map
// for every id in wantedIdsSet (or every id at all, if that's null).
// Used by the bulk issue/return functions so a 20-book bulk action
// still only reads the Books Catalogue once, not 20 times.
function buildBookRowMap_(booksSheet, wantedIdsSet) {
  var last = booksSheet.getLastRow();
  var numRows = Math.max(last - BOOKS_START_ROW + 1, 0);
  var map = {};
  if (!numRows) return map;
  var ids = booksSheet.getRange(BOOKS_START_ROW, BOOKS_COL.id, numRows, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    var id = norm_(ids[i][0]);
    if (id && (!wantedIdsSet || wantedIdsSet[id])) map[id] = BOOKS_START_ROW + i;
  }
  return map;
}

function syncBookCatalogueStatusByRow_(booksSheet, row, newStatus) {
  if (!row || row < BOOKS_START_ROW) return;
  if (newStatus !== 'Discarded') {
    var current = String(booksSheet.getRange(row, BOOKS_COL.status).getValue() || '').trim();
    if (current === 'Discarded') return; // don't un-discard as a side effect
  }
  booksSheet.getRange(row, BOOKS_COL.status).setValue(newStatus);
}

function syncBookCatalogueStatus_(booksSheet, bookId, newStatus) {
  var row = findBooksCatalogueRow_(booksSheet, norm_(bookId));
  syncBookCatalogueStatusByRow_(booksSheet, row, newStatus);
}

var MEMBERS_COL = {
  id: 1, name: 2, phone: 3, email: 4, membStart: 5, membEnd: 6,
  status: 7, notes: 8, age: 9, gender: 10, profession: 11, qualification: 12,
  address: 13, group: 14,
  // NEW: head-office fee/receipt reporting fields (columns O-R). Blank for
  // every existing member until filled in via the Member Report tab's
  // "Update Fee/Receipt Details" panel - see addFeeReceiptFieldsToMembers.gs
  // for the one-time header setup.
  amount: 15, receiptNo: 16, receiptDate: 17, challanDetail: 18,
  // NEW: a photo link column right in Members itself (column S), so it's
  // visible and editable alongside every other field without opening
  // MEMBER_GATE_PROFILE. This is now the primary place a photo link is
  // saved; MEMBER_GATE_PROFILE.PhotoPath is kept in sync alongside it
  // (that's what the Gate Entry kiosk screen reads), and old photos saved
  // only there before this column existed are still picked up as a
  // fallback - see getMemberFullDetails() and buildGateProfilePhotoMap_().
  photoLink: 19
};
var MEMBERS_START_ROW = 4;
var MEMBERS_LAST_COL = 19;

// NON_MEMBERS and GATE_ENTRY - added by you directly, headers match the
// original sample exactly. NON_MEMBERS data starts at row 2 (single header
// row) - different convention from Books Catalogue/Members (row 4), so
// don't copy that constant by mistake.
var NONMEM_COL = {
  id: 1, name: 2, address: 3, phone: 4, purpose: 5, vip: 6,
  specialMsg: 7, photoPath: 8, registerDate: 9, occupation: 10,
  // NEW: email, for the "Visitor Registration Approved" email below. Blank
  // for any visitor whose form response has no "Email ID" answer - that's
  // expected (the form's Email ID question should be optional), and every
  // caller treats a blank email as "skip sending, no error".
  email: 11
};
var NONMEM_START_ROW = 2;
var NONMEM_LAST_COL = 11;

var VLOG_TOTAL_COLS = 5; // Timestamp, Rack, Book ID, Match Status, Staff Name

// Gets the VERIFIED LOG sheet, creating it (with headers) if it's missing
// or was accidentally renamed/deleted - every stack-verification function
// depends on this sheet existing, so this stops the whole feature from
// silently breaking if the sheet ever goes missing.
function getVlog_() {
  var ss = ss_();
  var sh = ss.getSheetByName('VERIFIED LOG');
  if (!sh) {
    sh = ss.insertSheet('VERIFIED LOG');
    sh.getRange(1, 1, 1, VLOG_TOTAL_COLS)
      .setValues([['Timestamp', 'Rack', 'Book ID', 'Match Status', 'Staff Name']])
      .setFontWeight('bold');
  }
  return sh;
}

// ---- Settings / init ---------------------------------------------------
// NOTE: rows 1-11 of your existing 'Settings' sheet are untouched. These
// are NEW rows appended below - see MANUAL for exact values to add.

function getInitData() {
  var sh = ss_();
  var st = sh.getSheetByName('SETTINGS_CARDS'); // new sheet, see MANUAL - kept
                                                 // separate from your existing
                                                 // 'Settings' sheet so nothing
                                                 // there is at risk of being
                                                 // overwritten by mistake.
  var racksSh = sh.getSheetByName('RACKS');
  var staffSh = sh.getSheetByName('STAFF');

  var settings = {
    institution: st.getRange('B2').getValue(),
    library: st.getRange('B3').getValue(),
    address: st.getRange('B4').getValue(),
    labelFooter: st.getRange('B5').getValue(),
    memberPhotoFolderId: st.getRange('B6').getValue(),
    booksAllowedText: st.getRange('B7').getValue(),
    rulesText: st.getRange('B8').getValue(),
    labelHeader: st.getRange('B9').getValue()
  };

  var racks = [];
  var rLast = racksSh.getLastRow();
  if (rLast >= 2) {
    racksSh.getRange(2, 1, rLast - 1, 1).getValues().forEach(function(r) {
      if (r[0]) racks.push(String(r[0]).trim());
    });
  }

  var staff = [];
  var sLast = staffSh.getLastRow();
  if (sLast >= 2) {
    staffSh.getRange(2, 1, sLast - 1, 1).getValues().forEach(function(r) {
      if (r[0]) staff.push(String(r[0]).trim());
    });
  }

  return { settings: settings, racks: racks, staff: staff };
}

// ---- ID lists ----------------------------------------------------------

function getBookIdList() {
  var sh = ss_().getSheetByName('Books Catalogue');
  var lastRow = sh.getLastRow();
  var numRows = lastRow - BOOKS_START_ROW + 1;
  if (numRows <= 0) return [];
  var col = sh.getRange(BOOKS_START_ROW, BOOKS_COL.id, numRows, 1).getValues();
  var out = [], seen = {};
  col.forEach(function(r) {
    var id = norm_(r[0]);
    if (id && !seen[id]) { seen[id] = true; out.push(id); }
  });
  return out;
}

function getMemberIdList() {
  var sh = ss_().getSheetByName('Members');
  var lastRow = sh.getLastRow();
  var numRows = lastRow - MEMBERS_START_ROW + 1;
  if (numRows <= 0) return [];
  var col = sh.getRange(MEMBERS_START_ROW, MEMBERS_COL.id, numRows, 1).getValues();
  var out = [], seen = {};
  col.forEach(function(r) {
    var id = norm_(r[0]);
    if (id && !seen[id]) { seen[id] = true; out.push(id); }
  });
  return out;
}

// ---- Stack verification (Books Catalogue) -------------------------------

function getBooksMeta() {
  var sh = ss_().getSheetByName('Books Catalogue');
  var lastRow = sh.getLastRow();
  var MAX_SAFE_ROWS = 200000;
  if (lastRow - BOOKS_START_ROW > MAX_SAFE_ROWS) {
    throw new Error("Books Catalogue reports far more rows than expected - check for phantom rows before running this.");
  }
  return { totalBooks: Math.max(lastRow - BOOKS_START_ROW + 1, 0) };
}

function getVerifiedIds_() {
  var sh = getVlog_();
  var lastRow = sh.getLastRow();
  var out = [];
  if (lastRow < 2) return out;
  var data = sh.getRange(2, 1, lastRow - 1, VLOG_TOTAL_COLS).getValues();
  data.forEach(function(r) { out.push(norm_(r[2])); });
  return out;
}

function appendLogRows(rows) {
  if (!rows || !rows.length) return 0;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = getVlog_();
    var startRow = sh.getLastRow() + 1;
    var now = new Date();
    var out = rows.map(function(r) { return [now].concat(r); });
    sh.getRange(startRow, 1, out.length, VLOG_TOTAL_COLS).setValues(out);
    return out.length;
  } finally {
    lock.releaseLock();
  }
}

function getProgressData(dateStr) {
  var sh = getVlog_();
  var last = sh.getLastRow();
  var tz = Session.getScriptTimeZone();
  var result = { total: 0, found: 0, notFound: 0, duplicates: 0 };
  if (last >= 2) {
    var data = sh.getRange(2, 1, last - 1, VLOG_TOTAL_COLS).getValues();
    var targetKey = dateStr ? Utilities.formatDate(new Date(dateStr), tz, 'yyyy-MM-dd') : null;
    data.forEach(function(r) {
      var ts = r[0];
      if (!(ts instanceof Date)) return;
      var key = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
      if (targetKey && key !== targetKey) return;
      var status = String(r[VLOG_TOTAL_COLS - 2] || '');
      result.total++;
      if (status.indexOf('FOUND') === 0) result.found++;
      else if (status.indexOf('NOT FOUND') === 0) result.notFound++;
      if (status.indexOf('DUPLICATE') !== -1) result.duplicates++;
    });
  }
  return result;
}

function getStaffSummary() {
  var sh = getVlog_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2, 1, last - 1, VLOG_TOTAL_COLS).getValues();
  var tz = Session.getScriptTimeZone();
  var map = {};
  data.forEach(function(r) {
    var name = String(r[VLOG_TOTAL_COLS - 1] || '').trim();
    if (!name) return;
    var ts = r[0];
    if (!(ts instanceof Date)) return;
    if (!map[name]) map[name] = { name: name, start: ts, end: ts, count: 0 };
    var e = map[name];
    if (ts < e.start) e.start = ts;
    if (ts > e.end) e.end = ts;
    e.count++;
  });
  var out = [];
  Object.keys(map).forEach(function(name) {
    var e = map[name];
    out.push({
      name: e.name,
      start: Utilities.formatDate(e.start, tz, 'dd-MMM-yyyy'),
      end: Utilities.formatDate(e.end, tz, 'dd-MMM-yyyy'),
      count: e.count
    });
  });
  out.sort(function(a, b) { return b.count - a.count; });
  return out;
}

// ---- Photo lookup helper -----------------------------------------------

function findPhotoByFolder_(folderNameOrId, id) {
  if (!folderNameOrId) return null;
  var name = extractFolderId_(String(folderNameOrId).trim());
  var folder = null;

  // Try by NAME first - e.g. "Library_Photos_HNR", the same folder your
  // Gate Register / PhotoPath entries already point to.
  try {
    var it = DriveApp.getFoldersByName(name);
    if (it.hasNext()) folder = it.next();
  } catch (e) { /* ignore, try ID below */ }

  // Fallback: treat the value as a Drive folder ID instead.
  if (!folder) {
    try { folder = DriveApp.getFolderById(name); } catch (e) { return null; }
  }
  if (!folder) return null;

  var exts = ['jpg', 'jpeg', 'png'];
  for (var i = 0; i < exts.length; i++) {
    var fit = folder.getFilesByName(id + '.' + exts[i]);
    if (fit.hasNext()) {
      var blob = fit.next().getBlob();
      return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    }
  }
  return null;
}

// Accepts a plain folder ID, a full Drive folder URL, or a folder name -
// if it looks like a Drive URL, extracts just the ID portion automatically.
function extractFolderId_(value) {
  var m = value.match(/\/folders\/([-\w]+)/);
  return m ? m[1] : value;
}

// ---- Card data: BOOK CARD (6cm x 9cm portrait) -------------------------

function getBookCardData(ids) {
  var sh = ss_().getSheetByName('Books Catalogue');
  var wanted = {};
  ids.forEach(function(id) { wanted[norm_(id)] = true; });

  var lastRow = sh.getLastRow();
  var numRows = lastRow - BOOKS_START_ROW + 1;
  if (numRows <= 0) return [];
  var values = sh.getRange(BOOKS_START_ROW, 1, numRows, BOOKS_LAST_COL).getValues();

  var init = getInitData().settings;

  var out = [];
  values.forEach(function(row) {
    var id = norm_(row[BOOKS_COL.id - 1]);
    if (!id || !wanted[id]) return;
    out.push({
      id: id,
      callno: row[BOOKS_COL.callno - 1] || '',
      bookno: row[BOOKS_COL.karwarno - 1] || '',
      location: row[BOOKS_COL.shelfLocation - 1] || '',
      title: row[BOOKS_COL.title - 1] || '',
      author: row[BOOKS_COL.author - 1] || '',
      year: row[BOOKS_COL.year - 1] || '',
      pages: row[BOOKS_COL.pages - 1] || '',
      price: row[BOOKS_COL.price - 1] || '',
      institution: init.institution || '',
      library: init.library || '',
      footer: init.labelFooter || ''
    });
  });
  return out;
}

// ---- Card data: MEMBER CARD (6x8cm front + 6cm foldable back) ---------

// Wraps the same folder-photo lookup Gate Entry Cards uses, but never lets
// a Drive error break the whole card. Folder lookup first (reliable, same
// as Gate Entry Cards), then the pasted Drive link (Members!Photo_Link /
// MEMBER_GATE_PROFILE) as a fallback for anyone without a folder photo.
function getMemberCardPhoto_(folderSetting, id) {
  var photo = null;
  try {
    photo = findPhotoByFolder_(folderSetting, id);
  } catch (e) {}
  if (!photo) {
    try {
      photo = resolvePhotoPath_(gateLookupProfile_(id).photoPath);
    } catch (e) {}
  }
  return photo;
}

function getMemberCardData(ids) {
  var sh = ss_().getSheetByName('Members');
  var wanted = {};
  ids.forEach(function(id) { wanted[norm_(id)] = true; });

  var lastRow = sh.getLastRow();
  var rows = lastRow >= MEMBERS_START_ROW
    ? sh.getRange(MEMBERS_START_ROW, 1, lastRow - MEMBERS_START_ROW + 1, MEMBERS_LAST_COL).getValues()
    : [];

  var init = getInitData().settings;
  var out = [];
  rows.forEach(function(r) {
    var id = norm_(r[MEMBERS_COL.id - 1]);
    if (!id || !wanted[id]) return;
    var status = String(r[MEMBERS_COL.status - 1] || '');
    var membEnd = r[MEMBERS_COL.membEnd - 1];
    out.push({
      id: id,
      name: String(r[MEMBERS_COL.name - 1] || ''),
      category: status, // no Student/Public field in this sheet - shows Status instead
      phone: String(r[MEMBERS_COL.phone - 1] || ''),
      email: String(r[MEMBERS_COL.email - 1] || ''),
      memberAddress: String(r[MEMBERS_COL.address - 1] || ''),
      regDate: fmtDate_(r[MEMBERS_COL.membStart - 1]),
      isClosed: status.toUpperCase() === 'CLOSED' || status.toUpperCase() === 'INACTIVE',
      booksAllowed: init.booksAllowedText || '',
      validUpto: fmtDate_(membEnd),
      institution: init.institution || '',
      library: init.library || '',
      address2: init.address || '',
      rulesText: init.rulesText || '1. One book can be borrowed on this card.\n2. Books will not be borrowed unless cards are presented.\n3. Beware of this card. You will be responsible for each borrowed book from this card.\n4. Inform the change of address immediately.',
      // Same photo-by-folder lookup Gate Entry Cards already uses reliably,
      // tried FIRST - then the pasted Drive link (Members!Photo_Link /
      // MEMBER_GATE_PROFILE) as a fallback for anyone without a folder photo.
      // Wrapped in try/catch + logged: if the folder lookup ever throws
      // (e.g. a Drive permission/quota hiccup), it must NOT take the whole
      // card down with it - every other field on the card still needs to
      // render even if the photo genuinely can't be found this time.
      photo: getMemberCardPhoto_(init.memberPhotoFolderId, id)
    });
  });
  return out;
}

function saveLabelHeader(text) {
  var st = ss_().getSheetByName('SETTINGS_CARDS');
  st.getRange('B9').setValue(text || '');
  return true;
}

function getNonMemberIdList() {
  var sh = ss_().getSheetByName('NON_MEMBERS');
  var lastRow = sh.getLastRow();
  var numRows = lastRow - NONMEM_START_ROW + 1;
  if (numRows <= 0) return [];
  var col = sh.getRange(NONMEM_START_ROW, NONMEM_COL.id, numRows, 1).getValues();
  var out = [], seen = {};
  col.forEach(function(r) {
    var id = norm_(r[0]);
    if (id && !seen[id]) { seen[id] = true; out.push(id); }
  });
  return out;
}

// Resolves a Drive file ID or URL found in a PhotoPath cell. Harmlessly
// returns null for local Windows paths (e.g. C:\Library_Photos\5001.jpg) -
// falls back to the folder-by-ID convention in that case.
// Accepts a full Drive share link (any common format), a thumbnail link, or a
// bare Drive file ID - extracts the file ID and returns a direct, embeddable
// image URL. Pure string handling, no DriveApp calls needed here at all - the
// browser loads the image directly from Google, sidestepping the DriveApp
// permission issues entirely. The linked file must be shared as "Anyone with
// the link can view" for this to actually render.
function resolvePhotoPath_(pathValue) {
  var v = String(pathValue || '').trim();
  if (!v) return null;
  var m = v.match(/\/d\/([-\w]{15,})/) || v.match(/[?&]id=([-\w]{15,})/) || v.match(/^([-\w]{15,})$/);
  var fileId = m ? m[1] : null;
  if (!fileId) return null;
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';
}

// Same idea as resolvePhotoPath_, but specifically for files uploaded through
// our own Google Form (Photo/Signature/Address Proofs). These come in private
// to the form owner by default, so this also fixes the sharing to "Anyone
// with the link can view" - safe to do automatically here since we own the
// form and know these files need to be viewable by staff. Not used for
// manually-pasted links elsewhere, since those shouldn't be silently
// re-shared without the person's own intent.
function resolveFormUploadedFile_(pathValue) {
  var v = String(pathValue || '').trim();
  if (!v) return null;
  var m = v.match(/\/d\/([-\w]{15,})/) || v.match(/[?&]id=([-\w]{15,})/) || v.match(/^([-\w]{15,})$/);
  var fileId = m ? m[1] : null;
  if (!fileId) return null;
  try {
    var file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* if sharing can't be changed, still return the URL - it may already be shared correctly */ }
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';
}

// ---- Card data: GATE ENTRY CARD (9x5.5cm landscape, member OR non-member)

function getGateCardData(type, ids) {
  type = String(type || '').toUpperCase();
  var init = getInitData().settings;
  var out = [];
  var wanted = {};
  ids.forEach(function(id) { wanted[norm_(id)] = true; });

  if (type === 'MEMBER') {
    var msh = ss_().getSheetByName('Members');
    var lastRow = msh.getLastRow();
    if (lastRow >= MEMBERS_START_ROW) {
      msh.getRange(MEMBERS_START_ROW, 1, lastRow - MEMBERS_START_ROW + 1, MEMBERS_LAST_COL).getValues().forEach(function(r) {
        var id = norm_(r[MEMBERS_COL.id - 1]);
        if (!id || !wanted[id]) return;
        out.push({
          type: 'MEMBER', id: id,
          name: String(r[MEMBERS_COL.name - 1] || ''),
          categoryOrPurpose: String(r[MEMBERS_COL.status - 1] || ''),
          phone: String(r[MEMBERS_COL.phone - 1] || ''),
          vip: '',
          address: String(r[MEMBERS_COL.address - 1] || ''),
          gender: String(r[MEMBERS_COL.gender - 1] || ''),
          qualification: String(r[MEMBERS_COL.qualification - 1] || ''),
          email: String(r[MEMBERS_COL.email - 1] || ''),
          regDate: fmtDate_(r[MEMBERS_COL.membStart - 1]),
          booksAllowed: init.booksAllowedText || '',
          institution: init.institution || '', library: init.library || '',
          photo: findPhotoByFolder_(init.memberPhotoFolderId, id)
        });
      });
    }
  } else if (type === 'NONMEMBER') {
    var nsh = ss_().getSheetByName('NON_MEMBERS');
    var lastRow2 = nsh.getLastRow();
    if (lastRow2 >= NONMEM_START_ROW) {
      nsh.getRange(NONMEM_START_ROW, 1, lastRow2 - NONMEM_START_ROW + 1, NONMEM_LAST_COL).getValues().forEach(function(r) {
        var id = norm_(r[NONMEM_COL.id - 1]);
        if (!id || !wanted[id]) return;
        var photoPath = String(r[NONMEM_COL.photoPath - 1] || '');
        out.push({
          type: 'NON-MEMBER', id: id,
          name: String(r[NONMEM_COL.name - 1] || ''),
          categoryOrPurpose: String(r[NONMEM_COL.purpose - 1] || ''),
          phone: String(r[NONMEM_COL.phone - 1] || ''),
          vip: String(r[NONMEM_COL.vip - 1] || ''),
          address: String(r[NONMEM_COL.address - 1] || ''),
          qualification: String(r[NONMEM_COL.occupation - 1] || ''),
          regDate: fmtDate_(r[NONMEM_COL.registerDate - 1]),
          specialMsg: String(r[NONMEM_COL.specialMsg - 1] || ''),
          institution: init.institution || '', library: init.library || '',
          photo: resolvePhotoPath_(photoPath) || findPhotoByFolder_(init.memberPhotoFolderId, id)
        });
      });
    }
  }
  return out;
}

// ---- Stack Verification Summary sheet (overall + per-staff stats) -----
// Computed on demand (button in the Staff tab), not live formulas - keeps
// this fast even as VERIFIED LOG grows into the thousands of rows.

function refreshStackVerificationSummary() {
  var ss = ss_();
  var vlog = getVlog_();
  var summary = ss.getSheetByName('STACK VERIFICATION SUMMARY');
  if (!summary) summary = ss.insertSheet('STACK VERIFICATION SUMMARY');

  var lastRow = vlog.getLastRow();
  var data = lastRow >= 2 ? vlog.getRange(2, 1, lastRow - 1, VLOG_TOTAL_COLS).getValues() : [];

  var overallStart = null, overallEnd = null;
  var verifiedIds = {}, notFoundIds = {};
  var staffMap = {}; // name -> {first, last, scans, ids:{}}

  data.forEach(function(r) {
    var ts = r[0], id = norm_(r[2]), status = String(r[3] || ''), staff = String(r[4] || '').trim();
    if (!(ts instanceof Date)) return;
    if (!overallStart || ts < overallStart) overallStart = ts;
    if (!overallEnd || ts > overallEnd) overallEnd = ts;
    if (status.indexOf('FOUND') === 0) verifiedIds[id] = true;
    else if (status.indexOf('NOT FOUND') === 0) notFoundIds[id] = true;

    if (staff) {
      if (!staffMap[staff]) staffMap[staff] = { first: ts, last: ts, scans: 0, ids: {} };
      var s = staffMap[staff];
      if (ts < s.first) s.first = ts;
      if (ts > s.last) s.last = ts;
      s.scans++;
      if (status.indexOf('FOUND') === 0) s.ids[id] = true;
    }
  });

  var totalBooks = getBooksMeta().totalBooks;
  var uniqueVerified = Object.keys(verifiedIds).length;
  var uniqueNotFound = Object.keys(notFoundIds).length;
  var pct = totalBooks > 0 ? Math.round((uniqueVerified / totalBooks) * 1000) / 10 : 0;

  summary.clear();
  summary.getRange(1, 1).setValue('STACK VERIFICATION - SUMMARY').setFontWeight('bold').setFontSize(13);

  var overview = [
    ['Overall Started', overallStart ? fmtDate_(overallStart) : '-'],
    ['Last Activity', overallEnd ? fmtDate_(overallEnd) : '-'],
    ['Total Books in Catalogue', totalBooks],
    ['Unique Books Verified', uniqueVerified],
    ['Unique Not-Found Entries', uniqueNotFound],
    ['% Complete', pct + '%']
  ];
  summary.getRange(3, 1, overview.length, 2).setValues(overview);
  summary.getRange(3, 1, overview.length, 1).setFontWeight('bold');

  var headerRow = 3 + overview.length + 2;
  summary.getRange(headerRow, 1, 1, 5)
    .setValues([['Staff Name', 'First Scan', 'Last Scan', 'Total Scans', 'Unique Books Verified']])
    .setFontWeight('bold');

  var staffRows = Object.keys(staffMap).map(function(name) {
    var s = staffMap[name];
    return [name, fmtDate_(s.first), fmtDate_(s.last), s.scans, Object.keys(s.ids).length];
  });
  staffRows.sort(function(a, b) { return b[4] - a[4]; });
  if (staffRows.length) {
    summary.getRange(headerRow + 1, 1, staffRows.length, 5).setValues(staffRows);
  }

  summary.autoResizeColumns(1, 5);
  return { totalBooks: totalBooks, uniqueVerified: uniqueVerified, pct: pct, staffCount: staffRows.length };
}

// ---- Stack Verification Detail sheet (per-book, with Title/Author) ----
// One row per scan, joined against Books Catalogue for Title/Author -
// this is the "which 100 books did I actually scan" view. Kept as a
// separate on-demand sheet (not live formulas) for the same performance
// reason as the Summary sheet.

function refreshStackVerificationDetail() {
  var ss = ss_();
  var vlog = getVlog_();
  var books = ss.getSheetByName('Books Catalogue');
  var detail = ss.getSheetByName('STACK VERIFICATION DETAIL');
  if (!detail) detail = ss.insertSheet('STACK VERIFICATION DETAIL');

  // Build a Book ID -> full accession-register row lookup, ALL columns
  // exactly as they appear in Books Catalogue.
  var bLast = books.getLastRow();
  var lookup = {};
  if (bLast >= BOOKS_START_ROW) {
    books.getRange(BOOKS_START_ROW, 1, bLast - BOOKS_START_ROW + 1, BOOKS_LAST_COL).getValues().forEach(function(r) {
      var id = norm_(r[BOOKS_COL.id - 1]);
      if (id) lookup[id] = r;
    });
  }

  var vLast = vlog.getLastRow();
  var rows = vLast >= 2 ? vlog.getRange(2, 1, vLast - 1, VLOG_TOTAL_COLS).getValues() : [];

  var bookHeaders = ['Book ID', 'Title', 'Author', 'Genre', 'Total Copies', 'Shelf Location',
    'Remarks', 'CALL NUMBER', 'KARWAR NUMBER', 'PUBLICATION YEAR', 'NUMBER OF PAGES', 'PRICE',
    'Publisher', 'Place', 'Barcode', 'Status'];
  var scanHeaders = ['Timestamp', 'Rack', 'Match Status', 'Staff Name'];

  detail.clear();
  detail.getRange(1, 1, 1, scanHeaders.length).setValues([scanHeaders]).setFontWeight('bold');
  detail.getRange(1, scanHeaders.length + 1, 1, bookHeaders.length).setValues([bookHeaders]).setFontWeight('bold');

  if (rows.length) {
    var out = rows.map(function(r) {
      var id = norm_(r[2]);
      var full = lookup[id] || new Array(BOOKS_LAST_COL).fill('');
      // scanHeaders order: Timestamp, Rack, Match Status, Staff Name
      return [r[0], r[1], r[3], r[4]].concat(full);
    });
    // Newest scans first, easiest to review a just-finished session.
    out.sort(function(a, b) { return b[0] - a[0]; });
    detail.getRange(2, 1, out.length, scanHeaders.length + bookHeaders.length).setValues(out);
    detail.getRange(2, 1, out.length, 1).setNumberFormat('dd-mmm-yyyy hh:mm');
  }
  detail.autoResizeColumns(1, scanHeaders.length + bookHeaders.length);

  // ---- Per-rack breakdown (start/finish per "stack") -------------------
  var rackMap = {};
  rows.forEach(function(r) {
    var ts = r[0], rack = String(r[1] || '').trim() || '(no rack)';
    var status = String(r[3] || '');
    if (!(ts instanceof Date)) return;
    if (!rackMap[rack]) rackMap[rack] = { first: ts, last: ts, scans: 0, ids: {} };
    var rk = rackMap[rack];
    if (ts < rk.first) rk.first = ts;
    if (ts > rk.last) rk.last = ts;
    rk.scans++;
    if (status.indexOf('FOUND') === 0) rk.ids[norm_(r[2])] = true;
  });
  var rackRows = Object.keys(rackMap).map(function(name) {
    var rk = rackMap[name];
    return [name, fmtDate_(rk.first), fmtDate_(rk.last), rk.scans, Object.keys(rk.ids).length];
  });
  rackRows.sort(function(a, b) { return a[0].localeCompare(b[0]); });

  var rackStartCol = scanHeaders.length + bookHeaders.length + 2; // one blank column gap
  detail.getRange(1, rackStartCol, 1, 5)
    .setValues([['Rack', 'Started', 'Finished', 'Total Scans', 'Unique Books Verified']])
    .setFontWeight('bold');
  if (rackRows.length) {
    detail.getRange(2, rackStartCol, rackRows.length, 5).setValues(rackRows);
  }
  detail.autoResizeColumns(rackStartCol, 5);

  return { rows: rows.length, racks: rackRows.length };
}

// ================================================================
// QUICK ISSUE & RETURN - actions on top of your working 'Book Issue'
// sheet. Does NOT touch the 'Quick Issue' sheet or its formulas -
// this reads 'Book Issue' independently and adds the one thing Quick
// Issue can't do: actually process a return/new issue with one click.
// ================================================================

var BI_SHEET = 'Book Issue';
var BI_START_ROW = 5; // row 4 is an instruction row, not data
var BI_COL = { issueId:1, memberId:2, bookId:3, memberName:4, bookTitle:5,
               issueDate:6, dueDate:7, returnDate:8, status:9, fine:10, renewCount:11 };
var BI_LAST_COL = 11;

function normId_(v){ return String(v||'').trim().toUpperCase(); }

// ================================================================
// BOOK ISSUE/RETURN EMAIL CONFIRMATIONS - OFF by default. These two
// send a confirmation the instant a book is issued or returned, to
// whatever email address is on file for that member (Members sheet,
// column D). Review the exact wording in sendBookIssueEmail_ and
// sendBookReturnEmail_ below - once you're happy with it, flip this
// one flag to true and every future issue/return starts emailing
// automatically. Leave it false and nothing changes from today.
// ================================================================
var LIB_EMAIL_NOTIFICATIONS_LIVE = true;

// Looks up one member's name + email from the Members sheet. Returns
// null if the member isn't found, or { name, email } if they are -
// email may still be '' if that member has no email on file, which
// every caller below treats as "skip silently", never an error.
function getMemberContact_(memberId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var last = sh.getLastRow();
  if (last < MEMBERS_START_ROW) return null;
  var rows = sh.getRange(MEMBERS_START_ROW, 1, last - MEMBERS_START_ROW + 1, MEMBERS_LAST_COL).getValues();
  var target = normId_(memberId);
  for (var i = 0; i < rows.length; i++) {
    if (normId_(rows[i][MEMBERS_COL.id - 1]) === target) {
      return { name: rows[i][MEMBERS_COL.name - 1], email: String(rows[i][MEMBERS_COL.email - 1] || '').trim() };
    }
  }
  return null;
}

// ================================================================
// EMAIL LOG - every real email this project sends (book issue/return,
// due/overdue reminders, catch-up, membership/visitor approval) gets one
// row here: who, what kind, about which book, and whether it actually
// sent. Two categories, kept separate per the user's request:
//   'Reminder' - due-tomorrow, 5/10/15-day overdue, 30-day long-overdue,
//                and the one-time catch-up
//   'Other'    - book issued, book returned, membership approved,
//                visitor approved
// This is what printMemberEmailRecordPdf_ below reads to build a
// member's printable record - nothing here, nothing to print.
// ================================================================
var EMAIL_LOG_SHEET_ = 'Email_Log';
var EMAIL_LOG_HEADERS_ = ['Timestamp', 'Category', 'Type', 'Member ID', 'Member Name', 'Email', 'Detail', 'Status'];

function emailLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(EMAIL_LOG_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(EMAIL_LOG_SHEET_);
    sh.getRange(1, 1, 1, EMAIL_LOG_HEADERS_.length).setValues([EMAIL_LOG_HEADERS_]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// category: 'Reminder' | 'Other'. status: 'Sent' | 'Failed' | 'Skipped-NoEmail'.
// Logging itself must never be able to break a real send, so every call
// site wraps this in a try/catch is unnecessary - this function catches
// its own errors internally instead.
function logEmail_(category, type, memberId, memberName, email, detail, status) {
  try {
    emailLogSheet_().appendRow([new Date(), category, type, normId_(memberId), memberName || '', email || '', detail || '', status]);
  } catch (e) { /* a logging failure must never affect the real email or the calling function */ }
}

// Every caller gets back one of these status strings, so the admin panel
// can show staff exactly what happened instead of emailing silently in
// the background with no visible sign it ran at all:
//   'off'      - LIB_EMAIL_NOTIFICATIONS_LIVE is still false
//   'no-email' - that member has no email address on file
//   'sent'     - the email went out
//   'failed'   - MailApp threw (quota, bad address, etc.)
// Koha-style e-mails: the wording and layout now live in MemberMail.gs
// (mmSend_). These four keep their old names so nothing else changes.
function sendBookIssueEmail_(memberId, bookTitle, dueDateFmt, bookId) {
  return mmSend_('issue', memberId, [{ bookId: bookId || '', title: bookTitle, issueDate: new Date(), dueDate: dueDateFmt }]);
}

function sendBookReturnEmail_(memberId, bookTitle, overdueDays, fine, bookId) {
  return mmSend_('return', memberId, [{ bookId: bookId || '', title: bookTitle, returnDate: new Date(), overdueDays: overdueDays, fine: fine }]);
}

// items: [{ title, dueDateFmt, bookId }]
function sendBookIssueBatchEmail_(memberId, items) {
  return mmSend_('issue', memberId, (items || []).map(function (it) {
    return { bookId: it.bookId || '', title: it.title, issueDate: new Date(), dueDate: it.dueDateFmt };
  }));
}

// items: [{ title, overdueDays, fine, bookId }]
function sendBookReturnBatchEmail_(memberId, items) {
  return mmSend_('return', memberId, (items || []).map(function (it) {
    return { bookId: it.bookId || '', title: it.title, returnDate: new Date(), overdueDays: it.overdueDays, fine: it.fine };
  }));
}

// ================================================================
// DUE-DATE REMINDERS - scans every currently-issued book once a day
// and emails anyone whose book is due back TOMORROW. This is opt-in
// by nature: nothing runs until you install the daily trigger from
// the "\ud83d\udce7 Library Email Notifications" menu (one click, one time).
// ================================================================
// Overdue follow-ups fire at exactly these many days overdue - 5, 10, 15 -
// each one once only (the daily check runs once a day, so a book overdue
// by exactly 5 days today gets the "1st reminder" email today, and won't
// match 5 again tomorrow).
var LIB_OVERDUE_REMINDER_DAYS_ = [5, 10, 15];

// After day 15, instead of stopping completely, send one more reminder
// every 30 days for as long as the book stays unreturned (day 45, 75,
// 105, ...) - this is what also catches your EXISTING historic backlog
// going forward: a book that's already 200 days overdue won't line up
// with 5/10/15 again, but it will eventually land on a 30-day mark from
// day 15 (every (day - 15) that's a multiple of 30). For the backlog
// that's already sitting there right now, use the one-time
// emailAllOverdueBorrowersNow_ catch-up below instead of waiting for
// that to line up on its own.
var LIB_LONG_OVERDUE_INTERVAL_DAYS_ = 30;

// One pass over Book Issue, used by both the real daily run and the
// on-demand test. Returns a result object with real counts (never just
// "it ran") so sendTestReminderNow_ can show exactly what happened, and
// sends nothing twice: a "due tomorrow" book and a book crossing one of
// the overdue thresholds are different rows, so no member gets duplicate
// mail for the same book on the same day.
function runDailyLibraryReminders_() {
  try { if (typeof resDailyCheck_ === 'function') resDailyCheck_(); } catch (e) {}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var last = biLastRealRow_(sh);
  var numRows = Math.max(last - BI_START_ROW + 1, 0);
  var today = stripTime_(new Date());
  var tomorrow = stripTime_(new Date(today.getTime() + 86400000));
  var result = {
    today: Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd-MMM-yyyy'),
    booksDueTomorrow: 0, overdueBooksAtThreshold: 0, longOverdueBooksAtInterval: 0,
    membersWithNoEmail: 0, emailsSent: 0, emailsFailed: 0
  };
  if (!numRows) return result;
  var rows = sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues();

  // memberId -> { name, email, dueTomorrow: [title,...], overdue: { 5: [...], 10: [...], 15: [...] }, longOverdue: [{title, overdueDays}, ...] }
  var byMember = {};
  var noEmailMembers = {};

  function memberBucket_(memberId) {
    if (byMember[memberId] || noEmailMembers[memberId]) return byMember[memberId] || null;
    var contact = getMemberContact_(memberId);
    if (!contact || !contact.email) { noEmailMembers[memberId] = true; return null; }
    byMember[memberId] = { name: contact.name, email: contact.email, dueTomorrow: [], overdue: {}, longOverdue: [] };
    return byMember[memberId];
  }

  rows.forEach(function(r) {
    if (String(r[BI_COL.status - 1]).trim().toUpperCase() !== 'ISSUED') return;
    var due = r[BI_COL.dueDate - 1];
    if (!(due instanceof Date)) return;
    due = stripTime_(due);
    var memberId = normId_(r[BI_COL.memberId - 1]);
    var bookTitle = r[BI_COL.bookTitle - 1];

    if (due.getTime() === tomorrow.getTime()) {
      result.booksDueTomorrow++;
      var mb = memberBucket_(memberId);
      if (mb) mb.dueTomorrow.push(mmItemLabel_(bookTitle, r[BI_COL.bookId - 1]));
      return;
    }

    var overdueDays = Math.round((today.getTime() - due.getTime()) / 86400000);
    if (LIB_OVERDUE_REMINDER_DAYS_.indexOf(overdueDays) !== -1) {
      result.overdueBooksAtThreshold++;
      var mb2 = memberBucket_(memberId);
      if (mb2) {
        if (!mb2.overdue[overdueDays]) mb2.overdue[overdueDays] = [];
        mb2.overdue[overdueDays].push(mmItemLabel_(bookTitle, r[BI_COL.bookId - 1]));
      }
    } else if (overdueDays > 15 && (overdueDays - 15) % LIB_LONG_OVERDUE_INTERVAL_DAYS_ === 0) {
      result.longOverdueBooksAtInterval++;
      var mb3 = memberBucket_(memberId);
      if (mb3) mb3.longOverdue.push({ title: mmItemLabel_(bookTitle, r[BI_COL.bookId - 1]), overdueDays: overdueDays });
    }
  });

  result.membersWithNoEmail = Object.keys(noEmailMembers).length;

  Object.keys(byMember).forEach(function(memberId) {
    var m = byMember[memberId];

    if (m.dueTomorrow.length) {
      try {
        MailApp.sendEmail({
          to: m.email, name: MM_.shortName, replyTo: MM_.email,
          subject: 'Reminder: Book(s) due tomorrow - BCL Honavar',
          body: 'Dear ' + m.name + ',\n\n' +
            'This is a reminder that the following book(s) you borrowed from Branch Central Library, Honavar ' +
            'are due back TOMORROW:\n\n' +
            m.dueTomorrow.map(function(t) { return '  - ' + t; }).join('\n') + '\n\n' +
            'Please return or renew them by the due date to avoid a late fine.\n\n' +
            mmFooter_()
        });
        result.emailsSent++;
        logEmail_('Reminder', 'Due Tomorrow', memberId, m.name, m.email, m.dueTomorrow.join(', '), 'Sent');
      } catch (e) {
        result.emailsFailed++;
        logEmail_('Reminder', 'Due Tomorrow', memberId, m.name, m.email, m.dueTomorrow.join(', '), 'Failed');
      }
    }

    LIB_OVERDUE_REMINDER_DAYS_.forEach(function(threshold) {
      var titles = m.overdue[threshold];
      if (!titles || !titles.length) return;
      var reminderNum = LIB_OVERDUE_REMINDER_DAYS_.indexOf(threshold) + 1; // 1st/2nd/3rd
      try {
        MailApp.sendEmail({
          to: m.email, name: MM_.shortName, replyTo: MM_.email,
          subject: 'Overdue Reminder #' + reminderNum + ' - ' + threshold + ' days - BCL Honavar',
          body: 'Dear ' + m.name + ',\n\n' +
            'Our records show the following book(s) borrowed from Branch Central Library, Honavar are now ' +
            threshold + ' days overdue:\n\n' +
            titles.map(function(t) { return '  - ' + t; }).join('\n') + '\n\n' +
            'Please return them at the earliest to avoid further fine accumulation.\n\n' +
            mmFooter_()
        });
        result.emailsSent++;
        logEmail_('Reminder', 'Overdue #' + reminderNum + ' (' + threshold + 'd)', memberId, m.name, m.email, titles.join(', '), 'Sent');
      } catch (e) {
        result.emailsFailed++;
        logEmail_('Reminder', 'Overdue #' + reminderNum + ' (' + threshold + 'd)', memberId, m.name, m.email, titles.join(', '), 'Failed');
      }
    });

    if (m.longOverdue.length) {
      try {
        MailApp.sendEmail({
          to: m.email, name: MM_.shortName, replyTo: MM_.email,
          subject: 'Still Overdue - Please Return - BCL Honavar',
          body: 'Dear ' + m.name + ',\n\n' +
            'Our records show the following book(s) borrowed from Branch Central Library, Honavar remain overdue:\n\n' +
            m.longOverdue.map(function(it) { return '  - ' + it.title + ' (' + it.overdueDays + ' days overdue)'; }).join('\n') + '\n\n' +
            'Please return them at the earliest to avoid further fine accumulation. If you are unable to return them, ' +
            'please contact the library to discuss.\n\n' +
            mmFooter_()
        });
        result.emailsSent++;
        logEmail_('Reminder', 'Still Overdue (30d+)', memberId, m.name, m.email,
          m.longOverdue.map(function(it) { return it.title + ' (' + it.overdueDays + 'd)'; }).join(', '), 'Sent');
      } catch (e) {
        result.emailsFailed++;
        logEmail_('Reminder', 'Still Overdue (30d+)', memberId, m.name, m.email,
          m.longOverdue.map(function(it) { return it.title + ' (' + it.overdueDays + 'd)'; }).join(', '), 'Failed');
      }
    }
  });

  return result;
}

// ONE-TIME CATCH-UP - run manually, whenever you want (not on the daily
// trigger). Emails EVERY member who has ANY currently overdue book right
// now, however overdue it is, in one message per member. This is what
// clears your historic backlog - loans that have been overdue for weeks
// or months already, which the 5/10/15-day and 30-day-interval checks
// above would otherwise only catch once they happen to land on one of
// those exact day counts going forward.
function emailAllOverdueBorrowersNow_() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'Email ALL Overdue Borrowers Now',
    'This sends one catch-up email to every member who currently has any overdue book, listing all their overdue ' +
    'titles - however overdue, old or new. Use this to clear your historic backlog once; day-to-day coverage ' +
    'going forward is handled by the daily reminder. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var last = biLastRealRow_(sh);
  var numRows = Math.max(last - BI_START_ROW + 1, 0);
  var today = stripTime_(new Date());
  var result = { overdueBooks: 0, membersWithNoEmail: 0, emailsSent: 0, emailsFailed: 0 };
  if (!numRows) { ui.alert('No book issue records found.'); return; }
  var rows = sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues();

  var byMember = {}; // memberId -> { name, email, items: [{title, overdueDays}] }
  var noEmailMembers = {};

  rows.forEach(function(r) {
    if (String(r[BI_COL.status - 1]).trim().toUpperCase() !== 'ISSUED') return;
    var due = r[BI_COL.dueDate - 1];
    if (!(due instanceof Date)) return;
    due = stripTime_(due);
    var overdueDays = Math.round((today.getTime() - due.getTime()) / 86400000);
    if (overdueDays <= 0) return;
    result.overdueBooks++;
    var memberId = normId_(r[BI_COL.memberId - 1]);
    if (!byMember[memberId] && !noEmailMembers[memberId]) {
      var contact = getMemberContact_(memberId);
      if (!contact || !contact.email) { noEmailMembers[memberId] = true; return; }
      byMember[memberId] = { name: contact.name, email: contact.email, items: [] };
    }
    if (byMember[memberId]) byMember[memberId].items.push({ title: mmItemLabel_(r[BI_COL.bookTitle - 1], r[BI_COL.bookId - 1]), overdueDays: overdueDays });
  });

  result.membersWithNoEmail = Object.keys(noEmailMembers).length;

  Object.keys(byMember).forEach(function(memberId) {
    var m = byMember[memberId];
    try {
      MailApp.sendEmail({
        to: m.email, name: MM_.shortName, replyTo: MM_.email,
        subject: 'Overdue Books - Please Return - BCL Honavar',
        body: 'Dear ' + m.name + ',\n\n' +
          'Our records show the following book(s) borrowed from Branch Central Library, Honavar are currently overdue:\n\n' +
          m.items.map(function(it) { return '  - ' + it.title + ' (' + it.overdueDays + ' days overdue)'; }).join('\n') + '\n\n' +
          'Please return them at the earliest to avoid further fine accumulation.\n\n' +
          mmFooter_()
      });
      result.emailsSent++;
      logEmail_('Reminder', 'Catch-up (backlog)', memberId, m.name, m.email,
        m.items.map(function(it) { return it.title + ' (' + it.overdueDays + 'd)'; }).join(', '), 'Sent');
    } catch (e) {
      result.emailsFailed++;
      logEmail_('Reminder', 'Catch-up (backlog)', memberId, m.name, m.email,
        m.items.map(function(it) { return it.title + ' (' + it.overdueDays + 'd)'; }).join(', '), 'Failed');
    }
  });

  ui.alert(
    'Catch-up complete.\n\n' +
    result.overdueBooks + ' overdue book(s) found, across ' + (Object.keys(byMember).length + result.membersWithNoEmail) + ' member(s).\n' +
    '\u2709 Emails sent: ' + result.emailsSent + '\n' +
    (result.membersWithNoEmail > 0 ? '\u26a0 Skipped (no email on file): ' + result.membersWithNoEmail + '\n' : '') +
    (result.emailsFailed > 0 ? '\u2717 Failed to send: ' + result.emailsFailed + '\n' : '')
  );
}

// Run once from the menu: installs a trigger that fires runDailyLibraryReminders_
// automatically every day around 8 AM. Safe to run again later (e.g. after
// changing the hour) - it removes any previous copy of this trigger first
// (both the current name and the old pre-overdue-reminders name, in case
// you installed before this update), so you never end up with duplicates.
function installDueDateReminderTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'runDailyLibraryReminders_' || fn === 'sendDueDateReminders_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyLibraryReminders_')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  SpreadsheetApp.getUi().alert(
    'Daily reminders installed. Every day around 8 AM this will automatically email: anyone whose book is due back ' +
    'the next day; anyone whose book is overdue by exactly 5, 10, or 15 days (3 reminders total); and after that, ' +
    'anyone still holding it gets one more email every 30 days for as long as it stays unreturned. ' +
    'No further setup needed.\n\n' +
    'This covers everything GOING FORWARD. For books that are ALREADY overdue right now (your existing backlog), ' +
    'use "\ud83d\udce7 Email All Currently Overdue Borrowers Now (one-time catch-up)" once to cover those too.'
  );
}

// Runs runDailyLibraryReminders_ right now, on demand, so you can check the
// wording/behaviour before (or after) installing the daily trigger above.
function sendTestReminderNow_() {
  var r = runDailyLibraryReminders_();
  var lines = ['Checked as of ' + r.today + ':', ''];
  if (r.booksDueTomorrow === 0 && r.overdueBooksAtThreshold === 0 && r.longOverdueBooksAtInterval === 0) {
    lines.push(
      'Nothing to send right now - no book is due back tomorrow, none crossed the 5/10/15-day overdue mark today, ' +
      'and none hit a 30-day long-overdue interval today. This is expected on a normal day, not an error.\n\n' +
      'To see a real email without waiting, use "\ud83e\uddea Send Me A Sample Reminder Email" instead - it sends an ' +
      'example straight to an address you choose. To cover books that are ALREADY overdue right now (not just ' +
      'ones crossing a mark today), use "\ud83d\udce7 Email All Currently Overdue Borrowers Now" instead.'
    );
  } else {
    lines.push(r.booksDueTomorrow + ' due tomorrow, ' + r.overdueBooksAtThreshold + ' crossing a 5/10/15-day mark, ' + r.longOverdueBooksAtInterval + ' crossing a 30-day long-overdue mark.');
    lines.push('\u2709 Emails sent: ' + r.emailsSent);
    if (r.membersWithNoEmail > 0) lines.push('\u26a0 Skipped (no email on file): ' + r.membersWithNoEmail);
    if (r.emailsFailed > 0) lines.push('\u2717 Failed to send: ' + r.emailsFailed);
  }
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

// On-demand proof that sending itself works, without needing a real book
// that's actually due tomorrow or actually overdue by 5/10/15 days - asks
// for an email address and sends one example of each wording straight to
// it, clearly marked SAMPLE so it's never mistaken for a real notice.
function sendSampleReminderEmail_() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Send Sample Reminder Email', 'Send a sample to which email address?', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var email = String(resp.getResponseText() || '').trim();
  if (!email) { ui.alert('No email address entered.'); return; }

  try {
    MailApp.sendEmail({
      to: email,
      subject: '[SAMPLE] Reminder: Book(s) due tomorrow - BCL Honavar',
      body: 'This is a SAMPLE - no real book is due. This is exactly what a real reminder looks like:\n\n' +
        '----------\n\n' +
        'Dear [Member Name],\n\n' +
        'This is a reminder that the following book(s) you borrowed from Branch Central Library, Honavar ' +
        'are due back TOMORROW:\n\n' +
        '  - [Book Title]\n\n' +
        'Please return or renew them by the due date to avoid a late fine.\n\n' +
        mmFooter_()
    });
    MailApp.sendEmail({
      to: email,
      subject: '[SAMPLE] Overdue Reminder #2 - 10 days - BCL Honavar',
      body: 'This is a SAMPLE - no real book is overdue. This is exactly what a real overdue reminder looks like:\n\n' +
        '----------\n\n' +
        'Dear [Member Name],\n\n' +
        'Our records show the following book(s) borrowed from Branch Central Library, Honavar are now ' +
        '10 days overdue:\n\n' +
        '  - [Book Title]\n\n' +
        'Please return them at the earliest to avoid further fine accumulation.\n\n' +
        mmFooter_()
    });
    ui.alert('Sent 2 sample emails to ' + email + ' - check that inbox now.');
  } catch (e) {
    ui.alert('Could not send: ' + e.message);
  }
}

function addLibraryEmailMenu_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('\ud83d\udce7 Library Email Notifications')
    .addItem('\ud83d\udce7 Install Daily Due-Date + Overdue Reminders (run once)', 'installDueDateReminderTrigger_')
    .addItem('\ud83d\udce7 Email All Currently Overdue Borrowers Now (one-time catch-up)', 'emailAllOverdueBorrowersNow_')
    .addSeparator()
    .addItem('\ud83e\uddea Run Reminder Check Now', 'sendTestReminderNow_')
    .addItem('\ud83e\uddea Send Me A Sample Reminder Email', 'sendSampleReminderEmail_')
    .addSeparator()
    .addItem('\ud83d\udda8\ufe0f Print Member Email Record (PDF)', 'printMemberEmailRecordPdf_')
    .addItem('\ud83d\udcee Send Bulk Email Record to Head Office (PDF)', 'sendBulkEmailRecordToHeadOffice_')
    .addToUi();
}

// ================================================================
// PRINTABLE PER-MEMBER EMAIL RECORD - one-page A4 PDF, saved to Drive.
// Every email this system sends gets logged (see logEmail_ above) in two
// separate categories, exactly as asked for:
//   'Reminder' - due-tomorrow / overdue (5,10,15-day) / 30-day long-overdue /
//                historic catch-up reminders
//   'Other'    - book issued/returned confirmations, membership and visitor
//                approval emails, application-received emails
// This tool lets staff pick one Member ID, pick which category (or both),
// pick a time window (week/month/year/all), and get back a single
// printable PDF - handy as proof a particular member really was reminded.
// ================================================================

function getEmailRecordsFolder_() {
  var folderName = 'Library_Email_Records';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function printMemberEmailRecordPdf_() {
  var ui = SpreadsheetApp.getUi();

  var idResp = ui.prompt('Print Member Email Record', 'Enter the Member ID (or Visitor ID):', ui.ButtonSet.OK_CANCEL);
  if (idResp.getSelectedButton() !== ui.Button.OK) return;
  var memberId = normId_(idResp.getResponseText());
  if (!memberId) { ui.alert('No Member ID entered.'); return; }

  var catResp = ui.prompt(
    'Which emails?',
    'Type one of: Reminder, Other, or Both\n\n' +
    '(Reminder = due-date/overdue reminders. Other = issue/return, membership & visitor approvals.)',
    ui.ButtonSet.OK_CANCEL
  );
  if (catResp.getSelectedButton() !== ui.Button.OK) return;
  var catChoice = String(catResp.getResponseText() || 'Both').trim().toLowerCase();
  if (['reminder', 'other', 'both'].indexOf(catChoice) === -1) catChoice = 'both';

  var periodResp = ui.prompt(
    'Which time period?',
    'Type one of: Week, Month, Year, or All',
    ui.ButtonSet.OK_CANCEL
  );
  if (periodResp.getSelectedButton() !== ui.Button.OK) return;
  var periodChoice = String(periodResp.getResponseText() || 'All').trim().toLowerCase();
  if (['week', 'month', 'year', 'all'].indexOf(periodChoice) === -1) periodChoice = 'all';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSh = ss.getSheetByName(EMAIL_LOG_SHEET_);
  var lastRow = logSh ? logSh.getLastRow() : 0;
  if (!logSh || lastRow < 2) { ui.alert('No email records have been logged yet.'); return; }

  var now = new Date();
  var cutoff = null;
  if (periodChoice === 'week') cutoff = new Date(now.getTime() - 7 * 86400000);
  else if (periodChoice === 'month') cutoff = new Date(now.getTime() - 30 * 86400000);
  else if (periodChoice === 'year') cutoff = new Date(now.getTime() - 365 * 86400000);

  var data = logSh.getRange(2, 1, lastRow - 1, EMAIL_LOG_HEADERS_.length).getValues();
  var matches = data.filter(function(r) {
    if (normId_(r[3]) !== memberId) return false; // Member ID col
    var category = String(r[1] || '').trim().toLowerCase();
    if (catChoice !== 'both' && category !== catChoice) return false;
    if (cutoff && !(r[0] instanceof Date && r[0].getTime() >= cutoff.getTime())) return false;
    return true;
  });

  if (!matches.length) {
    ui.alert('No matching email records found for Member ID ' + memberId + ' in that category/period.');
    return;
  }

  // Newest first reads better on a printed record.
  matches.sort(function(a, b) { return b[0].getTime() - a[0].getTime(); });

  var memberName = matches[0][4] || '';
  var memberEmail = matches[0][5] || '';
  var tz = Session.getScriptTimeZone();

  // Build (or reuse) a dedicated sheet for this one printout.
  var PRINT_SHEET_NAME = 'Print_Email_Record';
  var psh = ss.getSheetByName(PRINT_SHEET_NAME);
  if (psh) { psh.clear(); } else { psh = ss.insertSheet(PRINT_SHEET_NAME); }

  var title = 'Branch Central Library, Honavar - Email Record';
  var catLabel = catChoice === 'both' ? 'Reminder + Other' : (catChoice.charAt(0).toUpperCase() + catChoice.slice(1));
  var periodLabel = periodChoice.charAt(0).toUpperCase() + periodChoice.slice(1);

  psh.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(14);
  psh.getRange(2, 1).setValue('Member ID: ' + memberId + '    Name: ' + memberName + '    Email: ' + memberEmail);
  psh.getRange(3, 1).setValue('Category: ' + catLabel + '    Period: ' + periodLabel +
    '    Generated: ' + Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm'));
  psh.getRange(2, 1, 2, 1).setFontSize(10);

  var tableHeaders = ['Date', 'Category', 'Type', 'Detail', 'Status'];
  psh.getRange(5, 1, 1, tableHeaders.length).setValues([tableHeaders]).setFontWeight('bold').setBackground('#f1f3f4');

  var rowsOut = matches.map(function(r) {
    return [
      Utilities.formatDate(r[0], tz, 'dd-MMM-yyyy HH:mm'),
      r[1], r[2],
      String(r[6] || '').substring(0, 70),
      r[7]
    ];
  });
  psh.getRange(6, 1, rowsOut.length, tableHeaders.length).setValues(rowsOut);

  psh.setColumnWidth(1, 120);
  psh.setColumnWidth(2, 80);
  psh.setColumnWidth(3, 140);
  psh.setColumnWidth(4, 260);
  psh.setColumnWidth(5, 90);
  psh.getRange(1, 1, 5 + rowsOut.length, tableHeaders.length).setFontSize(9);
  psh.getRange(5, 1, 1, tableHeaders.length).setFontSize(9);
  psh.setFrozenRows(5);

  SpreadsheetApp.flush();

  // Export just this one sheet as a single-page A4 PDF.
  var ssUrl = ss.getUrl();
  var exportUrl = ssUrl.replace(/\/edit.*$/, '') +
    '/export?format=pdf&gid=' + psh.getSheetId() +
    '&size=A4&portrait=true&fitw=true&scale=4&top_margin=0.3&bottom_margin=0.3' +
    '&left_margin=0.3&right_margin=0.3&gridlines=false&printtitle=false&sheetnames=false&pagenum=false&horizontal_alignment=CENTER';

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token } });
  var pdfBlob = response.getBlob().setName(
    'EmailRecord_' + memberId + '_' + catChoice + '_' + periodChoice + '_' +
    Utilities.formatDate(now, tz, 'yyyyMMdd_HHmm') + '.pdf'
  );

  var folder = getEmailRecordsFolder_();
  var file = folder.createFile(pdfBlob);

  ui.alert(
    'Printable record ready.\n\n' +
    matches.length + ' email(s) found for Member ID ' + memberId + ' (' + memberName + ').\n\n' +
    'Saved to Drive as: ' + file.getName() + '\n' +
    file.getUrl()
  );
}

// ================================================================
// BULK EMAIL RECORD FOR HEAD OFFICE - same Email_Log, but instead of one
// member, this covers EVERY member/visitor for a chosen period/category
// in one summary PDF (one row per member: how many Reminder emails, how
// many Other emails, when they last got one) - then both saves it to
// Drive AND emails it straight to Head Office as an attachment, since
// that's who this report is for.
// ================================================================

function sendBulkEmailRecordToHeadOffice_() {
  var ui = SpreadsheetApp.getUi();

  var catResp = ui.prompt(
    'Which emails?',
    'Type one of: Reminder, Other, or Both\n\n' +
    '(Reminder = due-date/overdue reminders. Other = issue/return, membership & visitor approvals.)',
    ui.ButtonSet.OK_CANCEL
  );
  if (catResp.getSelectedButton() !== ui.Button.OK) return;
  var catChoice = String(catResp.getResponseText() || 'Both').trim().toLowerCase();
  if (['reminder', 'other', 'both'].indexOf(catChoice) === -1) catChoice = 'both';

  var periodResp = ui.prompt(
    'Which time period?',
    'Type one of: Week, Month, Year, or All',
    ui.ButtonSet.OK_CANCEL
  );
  if (periodResp.getSelectedButton() !== ui.Button.OK) return;
  var periodChoice = String(periodResp.getResponseText() || 'All').trim().toLowerCase();
  if (['week', 'month', 'year', 'all'].indexOf(periodChoice) === -1) periodChoice = 'all';

  var props = PropertiesService.getScriptProperties();
  var savedEmail = props.getProperty('HEAD_OFFICE_EMAIL') || '';
  var emailResp = ui.prompt(
    'Head Office Email',
    'Enter the email address to send this bulk record to' +
    (savedEmail ? ' (leave blank to reuse ' + savedEmail + ')' : '') + ':',
    ui.ButtonSet.OK_CANCEL
  );
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  var headOfficeEmail = String(emailResp.getResponseText() || '').trim() || savedEmail;
  if (!headOfficeEmail) { ui.alert('No Head Office email entered.'); return; }
  props.setProperty('HEAD_OFFICE_EMAIL', headOfficeEmail);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSh = ss.getSheetByName(EMAIL_LOG_SHEET_);
  var lastRow = logSh ? logSh.getLastRow() : 0;
  if (!logSh || lastRow < 2) { ui.alert('No email records have been logged yet.'); return; }

  var now = new Date();
  var cutoff = null;
  if (periodChoice === 'week') cutoff = new Date(now.getTime() - 7 * 86400000);
  else if (periodChoice === 'month') cutoff = new Date(now.getTime() - 30 * 86400000);
  else if (periodChoice === 'year') cutoff = new Date(now.getTime() - 365 * 86400000);

  var data = logSh.getRange(2, 1, lastRow - 1, EMAIL_LOG_HEADERS_.length).getValues();
  var matches = data.filter(function(r) {
    var category = String(r[1] || '').trim().toLowerCase();
    if (catChoice !== 'both' && category !== catChoice) return false;
    if (cutoff && !(r[0] instanceof Date && r[0].getTime() >= cutoff.getTime())) return false;
    return true;
  });

  if (!matches.length) {
    ui.alert('No matching email records found for that category/period.');
    return;
  }

  // Aggregate into one row per member - count of each category + last sent date.
  var byMember = {};
  var order = [];
  matches.forEach(function(r) {
    var mid = normId_(r[3]);
    if (!byMember[mid]) {
      byMember[mid] = { id: mid, name: r[4] || '', email: r[5] || '', reminder: 0, other: 0, total: 0, lastDate: r[0] };
      order.push(mid);
    }
    var b = byMember[mid];
    var category = String(r[1] || '').trim().toLowerCase();
    if (category === 'reminder') b.reminder++; else b.other++;
    b.total++;
    if (r[0] instanceof Date && r[0].getTime() > b.lastDate.getTime()) b.lastDate = r[0];
  });
  order.sort(function(a, b) {
    var na = parseInt(a, 10), nb = parseInt(b, 10);
    if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
    return a.localeCompare(b);
  });

  var tz = Session.getScriptTimeZone();
  var catLabel = catChoice === 'both' ? 'Reminder + Other' : (catChoice.charAt(0).toUpperCase() + catChoice.slice(1));
  var periodLabel = periodChoice.charAt(0).toUpperCase() + periodChoice.slice(1);

  var PRINT_SHEET_NAME = 'Print_Bulk_Email_Record';
  var psh = ss.getSheetByName(PRINT_SHEET_NAME);
  if (psh) { psh.clear(); } else { psh = ss.insertSheet(PRINT_SHEET_NAME); }

  psh.getRange(1, 1).setValue('Branch Central Library, Honavar - Email Notification Record (All Members)').setFontWeight('bold').setFontSize(14);
  psh.getRange(2, 1).setValue('Category: ' + catLabel + '    Period: ' + periodLabel +
    '    Generated: ' + Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm'));
  psh.getRange(3, 1).setValue('Total members: ' + order.length + '    Total emails: ' + matches.length);
  psh.getRange(2, 1, 2, 1).setFontSize(10);

  var tableHeaders = ['Member ID', 'Name', 'Email', 'Reminder Emails', 'Other Emails', 'Total', 'Last Sent'];
  psh.getRange(5, 1, 1, tableHeaders.length).setValues([tableHeaders]).setFontWeight('bold').setBackground('#f1f3f4');

  var rowsOut = order.map(function(mid) {
    var b = byMember[mid];
    return [b.id, b.name, b.email, b.reminder, b.other, b.total, Utilities.formatDate(b.lastDate, tz, 'dd-MMM-yyyy')];
  });
  psh.getRange(6, 1, rowsOut.length, tableHeaders.length).setValues(rowsOut);

  psh.setColumnWidth(1, 90);
  psh.setColumnWidth(2, 160);
  psh.setColumnWidth(3, 200);
  psh.setColumnWidth(4, 110);
  psh.setColumnWidth(5, 95);
  psh.setColumnWidth(6, 70);
  psh.setColumnWidth(7, 100);
  psh.getRange(1, 1, 5 + rowsOut.length, tableHeaders.length).setFontSize(9);
  psh.setFrozenRows(5);

  SpreadsheetApp.flush();

  var ssUrl = ss.getUrl();
  var exportUrl = ssUrl.replace(/\/edit.*$/, '') +
    '/export?format=pdf&gid=' + psh.getSheetId() +
    '&size=A4&portrait=true&fitw=true&top_margin=0.3&bottom_margin=0.3' +
    '&left_margin=0.3&right_margin=0.3&gridlines=false&printtitle=false&sheetnames=false&pagenum=true';

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token } });
  var fileLabel = catChoice + '_' + periodChoice + '_' + Utilities.formatDate(now, tz, 'yyyyMMdd_HHmm');
  var pdfBlob = response.getBlob().setName('BulkEmailRecord_' + fileLabel + '.pdf');

  var folder = getEmailRecordsFolder_();
  var file = folder.createFile(pdfBlob);

  var sendStatus = 'Sent';
  try {
    MailApp.sendEmail({
      to: headOfficeEmail,
      subject: 'Library Email Notification Record (' + catLabel + ', ' + periodLabel + ') - BCL Honavar',
      body: 'Dear Sir/Madam,\n\n' +
        'Please find attached the email notification record for Branch Central Library, Honavar.\n\n' +
        'Category: ' + catLabel + '\n' +
        'Period: ' + periodLabel + '\n' +
        'Total members covered: ' + order.length + '\n' +
        'Total emails sent: ' + matches.length + '\n\n' +
        'Thank you,\nBranch Central Library, Honavar',
      attachments: [pdfBlob]
    });
  } catch (mailErr) {
    sendStatus = 'Failed: ' + mailErr.message;
  }
  logEmail_('Other', 'Bulk Record Sent to Head Office', '', '', headOfficeEmail,
    catLabel + ' / ' + periodLabel + ' (' + order.length + ' members)', sendStatus);

  ui.alert(
    (sendStatus === 'Sent' ? 'Bulk record sent to Head Office.' : 'Bulk record saved, but the email to Head Office failed: ' + sendStatus) +
    '\n\n' + order.length + ' member(s), ' + matches.length + ' email(s) total.\n\n' +
    'Saved to Drive as: ' + file.getName() + '\n' +
    file.getUrl() + '\n\n' +
    'Sent to: ' + headOfficeEmail
  );
}

// True last row with a real Issue ID, ignoring trailing formula-filled-
// but-blank rows (Book Issue is formatted out to row 40004 even though
// real data currently ends around row 1129 - see earlier bloat findings).
function biLastRealRow_(sh) {
  var last = sh.getLastRow();
  var numRows = Math.max(last - BI_START_ROW + 1, 0);
  if (!numRows) return BI_START_ROW - 1;
  var col = sh.getRange(BI_START_ROW, BI_COL.issueId, numRows, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]).trim() !== '') return BI_START_ROW + i;
  }
  return BI_START_ROW - 1;
}

function getMemberTransactions(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var members = ss.getSheetByName('Members');
  memberId = normId_(memberId);
  if (!memberId) throw new Error('Enter a Member ID.');

  var memberName = '', phone = '';
  var mLast = members.getLastRow();
  if (mLast >= 4) {
    members.getRange(4, 1, mLast - 3, 3).getValues().some(function(r) {
      if (normId_(r[0]) === memberId) { memberName = r[1]; phone = r[2]; return true; }
      return false;
    });
  }

  var last = biLastRealRow_(sh);
  var numRows = Math.max(last - BI_START_ROW + 1, 0);
  var rows = numRows ? sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];
  var out = [];
  var currentlyHeld = 0, returned = 0, fineDue = 0, overdue = 0;
  var today = new Date();

  rows.forEach(function(r, i) {
    if (normId_(r[BI_COL.memberId - 1]) !== memberId) return;
    var status = String(r[BI_COL.status - 1] || '').trim();
    var isIssued = status.toUpperCase() === 'ISSUED';
    var dueDate = r[BI_COL.dueDate - 1];
    var isOverdue = isIssued && dueDate instanceof Date && today > dueDate;
    if (isIssued) currentlyHeld++; else returned++;
    if (isOverdue) overdue++;
    fineDue += Number(r[BI_COL.fine - 1]) || 0;
    out.push({
      row: BI_START_ROW + i,
      issueId: r[BI_COL.issueId - 1],
      bookId: r[BI_COL.bookId - 1],
      bookTitle: r[BI_COL.bookTitle - 1],
      issueDate: fmtDate_(r[BI_COL.issueDate - 1]),
      issueDateTs: (r[BI_COL.issueDate - 1] instanceof Date) ? r[BI_COL.issueDate - 1].getTime() : 0,
      dueDate: fmtDate_(dueDate),
      returnDate: fmtDate_(r[BI_COL.returnDate - 1]),
      status: status,
      fine: r[BI_COL.fine - 1] || 0,
      renewCount: Number(r[BI_COL.renewCount - 1]) || 0,
      overdue: isOverdue
    });
  });

  out.sort(function(a, b) { return b.row - a.row; }); // most recent first

  return {
    memberId: memberId, memberName: memberName, phone: phone,
    totalBorrowed: out.length, currentlyHeld: currentlyHeld, returned: returned,
    fineDue: fineDue, overdue: overdue, rows: out
  };
}

function quickReturnBook(memberId, bookId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var books = ss.getSheetByName('Books Catalogue');
  memberId = normId_(memberId); bookId = normId_(bookId);
  if (!memberId || !bookId) throw new Error('Member ID and Book ID are both required.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var last = biLastRealRow_(sh);
    var numRows = Math.max(last - BI_START_ROW + 1, 0);
    var values = numRows ? sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];
    var foundRow = -1;
    for (var i = values.length - 1; i >= 0; i--) { // most recent matching issue first
      var r = values[i];
      if (normId_(r[BI_COL.memberId - 1]) === memberId &&
          normId_(r[BI_COL.bookId - 1]) === bookId &&
          String(r[BI_COL.status - 1]).trim().toUpperCase() === 'ISSUED') {
        foundRow = BI_START_ROW + i;
        break;
      }
    }
    if (foundRow === -1) {
      throw new Error('No currently-issued record found for Member ' + memberId + ' / Book ' + bookId + '. Check the IDs - it may already be returned.');
    }

    var today = stripTime_(new Date());
    var dueDate = sh.getRange(foundRow, BI_COL.dueDate).getValue();
    var finePerDay = Number(ss.getSheetByName('Settings').getRange('B4').getValue()) || 0;
    var overdueDays = (dueDate instanceof Date && today > dueDate)
      ? Math.floor((today - dueDate) / 86400000) : 0;
    var fine = overdueDays * finePerDay;

    sh.getRange(foundRow, BI_COL.returnDate).setValue(today);
    sh.getRange(foundRow, BI_COL.status).setValue('Returned');
    sh.getRange(foundRow, BI_COL.fine).setValue(fine);
    syncBookCatalogueStatus_(books, bookId, 'Available');

    var returnedTitle = sh.getRange(foundRow, BI_COL.bookTitle).getValue();
    var emailStatus = sendBookReturnEmail_(memberId, returnedTitle, overdueDays, fine, bookId);
    var holdNote = (typeof resOnReturn_ === 'function' ? resOnReturn_(bookId) : '');

    return {
      issueId: sh.getRange(foundRow, BI_COL.issueId).getValue(),
      bookTitle: returnedTitle + holdNote, holdNote: holdNote,
      overdueDays: overdueDays, fine: fine, row: foundRow, emailStatus: emailStatus
    };
  } finally {
    lock.releaseLock();
  }
}

function quickRenewBook(memberId, bookId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  memberId = normId_(memberId); bookId = normId_(bookId);
  if (!memberId || !bookId) throw new Error('Member ID and Book ID are both required.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var last = biLastRealRow_(sh);
    var numRows = Math.max(last - BI_START_ROW + 1, 0);
    var values = numRows ? sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];
    var foundRow = -1;
    for (var i = values.length - 1; i >= 0; i--) {
      var r = values[i];
      if (normId_(r[BI_COL.memberId - 1]) === memberId &&
          normId_(r[BI_COL.bookId - 1]) === bookId &&
          String(r[BI_COL.status - 1]).trim().toUpperCase() === 'ISSUED') {
        foundRow = BI_START_ROW + i;
        break;
      }
    }
    if (foundRow === -1) {
      throw new Error('No currently-issued record found for Member ' + memberId + ' / Book ' + bookId + '.');
    }

    var settings = ss.getSheetByName('Settings');
    var maxRenewalRaw = settings.getRange('B6').getValue();
    var currentCount = Number(sh.getRange(foundRow, BI_COL.renewCount).getValue()) || 0;
    var maxRenewalNum = parseInt(maxRenewalRaw, 10);
    if (!isNaN(maxRenewalNum) && currentCount >= maxRenewalNum) {
      throw new Error('Renewal limit reached (' + maxRenewalNum + ') for this book - it must be returned instead.');
    }
    var resRenew = (typeof resRenewBlocker_ === 'function' ? resRenewBlocker_(bookId, memberId) : '');
    if (resRenew) throw new Error(resRenew);

    var loanDays = Number(settings.getRange('B5').getValue()) || 15;
    var today = stripTime_(new Date());
    var newDue = new Date(today.getTime() + loanDays * 86400000);

    sh.getRange(foundRow, BI_COL.dueDate).setValue(newDue);
    sh.getRange(foundRow, BI_COL.renewCount).setValue(currentCount + 1);
    var renewedTitle = sh.getRange(foundRow, BI_COL.bookTitle).getValue();
    var renewEmailStatus = sendBookRenewEmail_(memberId, bookId, renewedTitle, newDue);

    return {
      issueId: sh.getRange(foundRow, BI_COL.issueId).getValue(),
      bookTitle: renewedTitle,
      newDueDate: fmtDate_(newDue),
      renewCount: currentCount + 1,
      row: foundRow,
      emailStatus: renewEmailStatus
    };
  } finally {
    lock.releaseLock();
  }
}

// -------- Lookup a single book's current loan status, before acting on it --------
function lookupBookLoanStatus(bookId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var books = ss.getSheetByName('Books Catalogue');
  bookId = normId_(bookId);
  if (!bookId) throw new Error('Enter an Accession Number.');

  var bookTitle = '';
  var bLast = books.getLastRow();
  if (bLast >= BOOKS_START_ROW) {
    books.getRange(BOOKS_START_ROW, 1, bLast - BOOKS_START_ROW + 1, 2).getValues().some(function(r) {
      if (normId_(r[0]) === bookId) { bookTitle = r[1]; return true; }
      return false;
    });
  }
  if (!bookTitle) return { bookFound: false };

  var last = biLastRealRow_(sh);
  var numRows = Math.max(last - BI_START_ROW + 1, 0);
  var values = numRows ? sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];
  for (var i = values.length - 1; i >= 0; i--) {
    var r = values[i];
    if (normId_(r[BI_COL.bookId - 1]) === bookId && String(r[BI_COL.status - 1]).trim().toUpperCase() === 'ISSUED') {
      var dueDate = r[BI_COL.dueDate - 1];
      var overdue = dueDate instanceof Date && new Date() > dueDate;
      return {
        bookFound: true, bookTitle: bookTitle, issued: true,
        memberId: r[BI_COL.memberId - 1], memberName: r[BI_COL.memberName - 1],
        dueDate: fmtDate_(dueDate), overdue: overdue
      };
    }
  }
  return { bookFound: true, bookTitle: bookTitle, issued: false };
}

// -------- Bulk return: just a list of Accession Numbers, no Member ID needed --------
// Looks up each book's own active loan automatically.
function bulkReturnByAccessionNumbers(bookIdsText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var books = ss.getSheetByName('Books Catalogue');
  var ids = String(bookIdsText || '').split(/[\n,]+/).map(normId_).filter(Boolean);
  if (!ids.length) throw new Error('Enter at least one Accession Number.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var last = biLastRealRow_(sh);
    var numRows = Math.max(last - BI_START_ROW + 1, 0);
    var values = numRows ? sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];
    var finePerDay = Number(ss.getSheetByName('Settings').getRange('B4').getValue()) || 0;
    var today = new Date();

    var wantedIds = {};
    ids.forEach(function(id) { wantedIds[id] = true; });
    var bookRows = buildBookRowMap_(books, wantedIds);

    var success = [], failed = [];
    var byMember = {}; // memberId -> { items: [{title, overdueDays, fine}], successIdx: [index into success[]] }
    ids.forEach(function(bookId) {
      var foundRow = -1;
      for (var i = values.length - 1; i >= 0; i--) {
        if (normId_(values[i][BI_COL.bookId - 1]) === bookId && String(values[i][BI_COL.status - 1]).trim().toUpperCase() === 'ISSUED') {
          foundRow = BI_START_ROW + i;
          break;
        }
      }
      if (foundRow === -1) { failed.push({ bookId: bookId, reason: 'Not currently issued' }); return; }

      var dueDate = sh.getRange(foundRow, BI_COL.dueDate).getValue();
      var overdueDays = (dueDate instanceof Date && today > dueDate) ? Math.floor((today - dueDate) / 86400000) : 0;
      var fine = overdueDays * finePerDay;
      var bookTitle = sh.getRange(foundRow, BI_COL.bookTitle).getValue();

      sh.getRange(foundRow, BI_COL.returnDate).setValue(today);
      sh.getRange(foundRow, BI_COL.status).setValue('Returned');
      sh.getRange(foundRow, BI_COL.fine).setValue(fine);
      syncBookCatalogueStatusByRow_(books, bookRows[bookId], 'Available');
      // reflect locally so a second occurrence of the same ID in this batch doesn't double-return
      values[foundRow - BI_START_ROW][BI_COL.status - 1] = 'Returned';

      var returnedMemberId = normId_(values[foundRow - BI_START_ROW][BI_COL.memberId - 1]);
      var holdNote = (typeof resOnReturn_ === 'function' ? resOnReturn_(bookId) : '');
      success.push({ bookId: bookId, bookTitle: bookTitle + holdNote, holdNote: holdNote, fine: fine, emailStatus: null });
      if (!byMember[returnedMemberId]) byMember[returnedMemberId] = { items: [], successIdx: [] };
      byMember[returnedMemberId].items.push({ title: bookTitle, overdueDays: overdueDays, fine: fine, bookId: bookId });
      byMember[returnedMemberId].successIdx.push(success.length - 1);
    });

    // One combined email per member, however many of their books were
    // returned in this one batch - not one email per book.
    Object.keys(byMember).forEach(function(mid) {
      var m = byMember[mid];
      var status = sendBookReturnBatchEmail_(mid, m.items);
      m.successIdx.forEach(function(idx) { success[idx].emailStatus = status; });
    });

    return { success: success, failed: failed };
  } finally {
    lock.releaseLock();
  }
}

// -------- Bulk issue/return for ONE member, MULTIPLE accession numbers --------
function bulkIssueForMember(memberId, bookIdsText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var members = ss.getSheetByName('Members');
  var books = ss.getSheetByName('Books Catalogue');
  memberId = normId_(memberId);
  if (!memberId) throw new Error('Enter a Member ID.');
  var ids = String(bookIdsText || '').split(/[\n,]+/).map(normId_).filter(Boolean);
  if (!ids.length) throw new Error('Enter at least one Accession Number.');

  var memberName = '';
  var mLast = members.getLastRow();
  if (mLast >= 4) {
    members.getRange(4, 1, mLast - 3, 2).getValues().some(function(r) {
      if (normId_(r[0]) === memberId) { memberName = r[1]; return true; }
      return false;
    });
  }
  if (!memberName) throw new Error('Member ID ' + memberId + ' not found in Members.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var last = biLastRealRow_(sh);
    var numRows = Math.max(last - BI_START_ROW + 1, 0);
    var values = numRows ? sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];
    var lastIssueId = 0;
    values.forEach(function(r) {
      var n = parseInt(r[BI_COL.issueId - 1], 10);
      if (!isNaN(n) && n > lastIssueId) lastIssueId = n;
    });

    var bLast = books.getLastRow();
    var bookTitles = {}, bookRows = {};
    if (bLast >= BOOKS_START_ROW) {
      books.getRange(BOOKS_START_ROW, 1, bLast - BOOKS_START_ROW + 1, 2).getValues().forEach(function(r, idx) {
        var id = normId_(r[0]);
        if (id) { bookTitles[id] = r[1]; bookRows[id] = BOOKS_START_ROW + idx; }
      });
    }

    var loanDays = Number(ss.getSheetByName('Settings').getRange('B5').getValue()) || 15;
    var today = stripTime_(new Date());
    var due = new Date(today.getTime() + loanDays * 86400000);
    var nextRow = last + 1;

    var success = [], failed = [];
    ids.forEach(function(bookId) {
      if (!bookTitles[bookId]) { failed.push({ bookId: bookId, reason: 'Book not found' }); return; }
      var alreadyIssued = values.some(function(r) {
        return normId_(r[BI_COL.bookId - 1]) === bookId && String(r[BI_COL.status - 1]).trim().toUpperCase() === 'ISSUED';
      });
      if (alreadyIssued) { failed.push({ bookId: bookId, reason: 'Already issued' }); return; }
      var resBlock = (typeof resBlocker_ === 'function' ? resBlocker_(bookId, memberId) : '');
      if (resBlock) { failed.push({ bookId: bookId, reason: resBlock }); return; }

      lastIssueId++;
      sh.getRange(nextRow, 1, 1, BI_LAST_COL).setValues([[
        lastIssueId, memberId, bookId, memberName, bookTitles[bookId], today, due, '', 'Issued', 0, 0
      ]]);
      values.push([lastIssueId, memberId, bookId, memberName, bookTitles[bookId], today, due, '', 'Issued', 0, 0]);
      syncBookCatalogueStatusByRow_(books, bookRows[bookId], 'Issued');
      if (typeof resOnIssue_ === 'function') resOnIssue_(bookId, memberId);
      success.push({ bookId: bookId, bookTitle: bookTitles[bookId], emailStatus: null });
      nextRow++;
    });

    // One combined email for this member listing every book just issued,
    // however many (one or several) - not one email per book.
    if (success.length) {
      var issueItems = success.map(function(s) { return { title: s.bookTitle, dueDateFmt: fmtDate_(due), bookId: s.bookId }; });
      var issueEmailStatus = sendBookIssueBatchEmail_(memberId, issueItems);
      success.forEach(function(s) { s.emailStatus = issueEmailStatus; });
    }

    return { success: success, failed: failed, dueDate: fmtDate_(due) };
  } finally {
    lock.releaseLock();
  }
}

function bulkReturnForMember(memberId, bookIdsText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var books = ss.getSheetByName('Books Catalogue');
  memberId = normId_(memberId);
  if (!memberId) throw new Error('Enter a Member ID.');
  var ids = String(bookIdsText || '').split(/[\n,]+/).map(normId_).filter(Boolean);
  if (!ids.length) throw new Error('Enter at least one Accession Number.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var last = biLastRealRow_(sh);
    var numRows = Math.max(last - BI_START_ROW + 1, 0);
    var values = numRows ? sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];
    var finePerDay = Number(ss.getSheetByName('Settings').getRange('B4').getValue()) || 0;
    var today = new Date();

    var wantedIds = {};
    ids.forEach(function(id) { wantedIds[id] = true; });
    var bookRows = buildBookRowMap_(books, wantedIds);

    var success = [], failed = [];
    ids.forEach(function(bookId) {
      var foundRow = -1;
      for (var i = values.length - 1; i >= 0; i--) {
        if (normId_(values[i][BI_COL.memberId - 1]) === memberId &&
            normId_(values[i][BI_COL.bookId - 1]) === bookId &&
            String(values[i][BI_COL.status - 1]).trim().toUpperCase() === 'ISSUED') {
          foundRow = BI_START_ROW + i;
          break;
        }
      }
      if (foundRow === -1) { failed.push({ bookId: bookId, reason: 'No active loan for this member' }); return; }

      var dueDate = sh.getRange(foundRow, BI_COL.dueDate).getValue();
      var overdueDays = (dueDate instanceof Date && today > dueDate) ? Math.floor((today - dueDate) / 86400000) : 0;
      var fine = overdueDays * finePerDay;
      var bookTitle = sh.getRange(foundRow, BI_COL.bookTitle).getValue();

      sh.getRange(foundRow, BI_COL.returnDate).setValue(today);
      sh.getRange(foundRow, BI_COL.status).setValue('Returned');
      sh.getRange(foundRow, BI_COL.fine).setValue(fine);
      syncBookCatalogueStatusByRow_(books, bookRows[bookId], 'Available');
      values[foundRow - BI_START_ROW][BI_COL.status - 1] = 'Returned';

      var holdNote = (typeof resOnReturn_ === 'function' ? resOnReturn_(bookId) : '');
      success.push({ bookId: bookId, bookTitle: bookTitle + holdNote, plainTitle: bookTitle, holdNote: holdNote, fine: fine, emailStatus: null, overdueDays: overdueDays });
    });

    // One combined email for this member listing every book just returned,
    // however many (one or several) - not one email per book.
    if (success.length) {
      var returnItems = success.map(function(s) { return { title: s.plainTitle || s.bookTitle, overdueDays: s.overdueDays, fine: s.fine, bookId: s.bookId }; });
      var returnEmailStatus = sendBookReturnBatchEmail_(memberId, returnItems);
      success.forEach(function(s) { s.emailStatus = returnEmailStatus; });
    }

    return { success: success, failed: failed };
  } finally {
    lock.releaseLock();
  }
}

function quickIssueBook(memberId, bookId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BI_SHEET);
  var members = ss.getSheetByName('Members');
  var books = ss.getSheetByName('Books Catalogue');
  memberId = normId_(memberId); bookId = normId_(bookId);
  if (!memberId || !bookId) throw new Error('Member ID and Book ID are both required.');

  var memberName = '';
  var mLast = members.getLastRow();
  if (mLast >= 4) {
    members.getRange(4, 1, mLast - 3, 2).getValues().some(function(r) {
      if (normId_(r[0]) === memberId) { memberName = r[1]; return true; }
      return false;
    });
  }
  if (!memberName) throw new Error('Member ID ' + memberId + ' not found in Members.');

  var bookTitle = '';
  var bLast = books.getLastRow();
  if (bLast >= BOOKS_START_ROW) {
    books.getRange(BOOKS_START_ROW, 1, bLast - BOOKS_START_ROW + 1, 2).getValues().some(function(r) {
      if (normId_(r[0]) === bookId) { bookTitle = r[1]; return true; }
      return false;
    });
  }
  if (!bookTitle) throw new Error('Book ID ' + bookId + ' not found in Books Catalogue.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var last = biLastRealRow_(sh);
    var numRows = Math.max(last - BI_START_ROW + 1, 0);
    var checkVals = numRows ? sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];
    var lastIssueId = 0;
    for (var k = 0; k < checkVals.length; k++) {
      if (normId_(checkVals[k][BI_COL.bookId - 1]) === bookId &&
          String(checkVals[k][BI_COL.status - 1]).trim().toUpperCase() === 'ISSUED') {
        throw new Error('Book ID ' + bookId + ' is already issued (Issue ID ' + checkVals[k][BI_COL.issueId - 1] + ') - return it first.');
      }
      var idNum = parseInt(checkVals[k][BI_COL.issueId - 1], 10);
      if (!isNaN(idNum) && idNum > lastIssueId) lastIssueId = idNum;
    }
    var newIssueId = lastIssueId + 1;
    var resBlock = (typeof resBlocker_ === 'function' ? resBlocker_(bookId, memberId) : '');
    if (resBlock) throw new Error('Cannot issue ' + bookId + ': ' + resBlock + '.');

    var loanDays = Number(ss.getSheetByName('Settings').getRange('B5').getValue()) || 15;
    var today = stripTime_(new Date());
    var due = new Date(today.getTime() + loanDays * 86400000);
    var newRow = last + 1;

    sh.getRange(newRow, 1, 1, BI_LAST_COL).setValues([[
      newIssueId, memberId, bookId, memberName, bookTitle, today, due, '', 'Issued', 0, 0
    ]]);
    syncBookCatalogueStatus_(books, bookId, 'Issued');
    if (typeof resOnIssue_ === 'function') resOnIssue_(bookId, memberId);
    var emailStatus = sendBookIssueEmail_(memberId, bookTitle, fmtDate_(due), bookId);

    return { issueId: newIssueId, bookTitle: bookTitle, dueDate: fmtDate_(due), row: newRow, emailStatus: emailStatus };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// PUBLIC BOOK SEARCH & DASHBOARD - no login, read-only, book-level
// data only. Deliberately excludes due dates and member names -
// only Available / Issued / Discarded status is shown publicly.
// ================================================================

function stripTime_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function getPublicDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var books = ss.getSheetByName('Books Catalogue');
  var decard = ss.getSheetByName('Decard Books');
  var bi = ss.getSheetByName('Book Issue');

  var totalBooks = getBooksMeta().totalBooks;

  var dLast = decard.getLastRow();
  var discardedBooks = 0;
  if (dLast >= 6) {
    decard.getRange(6, 1, dLast - 5, 1).getValues().forEach(function(r) {
      if (String(r[0]).trim() !== '') discardedBooks++;
    });
  }

  var biLast = biLastRealRow_(bi);
  var numRows = Math.max(biLast - BI_START_ROW + 1, 0);
  var rows = numRows ? bi.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues() : [];

  var today = stripTime_(new Date());
  var dow = today.getDay(); // 0=Sun..6=Sat
  var mondayOffset = dow === 0 ? 6 : dow - 1;
  var weekStart = new Date(today.getFullYear(), today.getMonth(), today.getDate() - mondayOffset);
  var monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  var yearStart = new Date(today.getFullYear(), 0, 1);

  var issuedCount = 0, overdueCount = 0;
  var todayIssued = 0, weekIssued = 0, monthIssued = 0, yearIssued = 0;

  function countIssue(d) {
    if (!(d instanceof Date)) return;
    var dd = stripTime_(d);
    if (dd.getTime() === today.getTime()) todayIssued++;
    if (dd >= weekStart) weekIssued++;
    if (dd >= monthStart) monthIssued++;
    if (dd >= yearStart) yearIssued++;
  }

  rows.forEach(function(r) {
    var status = String(r[BI_COL.status - 1] || '').trim().toUpperCase();
    var dueDate = r[BI_COL.dueDate - 1];
    if (status === 'ISSUED') {
      issuedCount++;
      if (dueDate instanceof Date && today > stripTime_(dueDate)) overdueCount++;
    }
    countIssue(r[BI_COL.issueDate - 1]); // borrowed-books count ONLY - returns/renewals excluded on purpose
  });

  return {
    totalBooks: totalBooks,
    discardedBooks: discardedBooks,
    netBooks: totalBooks - discardedBooks,
    issuedBooks: issuedCount,
    overdueBooks: overdueCount,
    todayTx: todayIssued, weekTx: weekIssued, monthTx: monthIssued, yearTx: yearIssued
  };
}

function getPublicGenres() {
  return getPublicDistinctColumn_(BOOKS_COL.genre);
}

function getPublicPublishers() {
  return getPublicDistinctColumn_(BOOKS_COL.publisher);
}

function getPublicLocations() {
  return getPublicDistinctColumn_(BOOKS_COL.shelfLocation);
}

function getPublicDistinctColumn_(colIndex) {
  var books = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Books Catalogue');
  var last = books.getLastRow();
  var numRows = Math.max(last - BOOKS_START_ROW + 1, 0);
  var seen = {};
  if (numRows) {
    books.getRange(BOOKS_START_ROW, colIndex, numRows, 1).getValues().forEach(function(r) {
      var v = String(r[0] || '').trim();
      if (v) seen[v] = true;
    });
  }
  return Object.keys(seen).sort();
}

function publicSearchBooks(query, field) {
  query = String(query || '').trim().toLowerCase();
  field = String(field || 'title').trim();
  if (!query) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var books = ss.getSheetByName('Books Catalogue');
  var decard = ss.getSheetByName('Decard Books');
  var bi = ss.getSheetByName('Book Issue');

  var discardedSet = {};
  var dLast = decard.getLastRow();
  if (dLast >= 6) {
    decard.getRange(6, 1, dLast - 5, 1).getValues().forEach(function(r) {
      var id = norm_(r[0]);
      if (id) discardedSet[id] = true;
    });
  }

  var issuedSet = {};
  var biLast = biLastRealRow_(bi);
  var biRows = Math.max(biLast - BI_START_ROW + 1, 0);
  if (biRows) {
    bi.getRange(BI_START_ROW, 1, biRows, BI_LAST_COL).getValues().forEach(function(r) {
      if (String(r[BI_COL.status - 1] || '').trim().toUpperCase() === 'ISSUED') {
        issuedSet[norm_(r[BI_COL.bookId - 1])] = true;
      }
    });
  }

  var bLast = books.getLastRow();
  var numRows = Math.max(bLast - BOOKS_START_ROW + 1, 0);
  var results = [];
  var MAX_RESULTS = 50;
  if (numRows) {
    var data = books.getRange(BOOKS_START_ROW, 1, numRows, BOOKS_LAST_COL).getValues();
    for (var i = 0; i < data.length && results.length < MAX_RESULTS; i++) {
      var r = data[i];
      var id = String(r[BOOKS_COL.id - 1] || '');
      var title = String(r[BOOKS_COL.title - 1] || '');
      var author = String(r[BOOKS_COL.author - 1] || '');
      var publisherVal = String(r[BOOKS_COL.publisher - 1] || '');
      var place = String(r[BOOKS_COL.place - 1] || '');
      var bookGenre = String(r[BOOKS_COL.genre - 1] || '');
      var shelfLoc = String(r[BOOKS_COL.shelfLocation - 1] || '');

      var fieldMap = { title: title, author: author, id: id, publisher: publisherVal, genre: bookGenre, location: shelfLoc };
      var target = String(fieldMap.hasOwnProperty(field) ? fieldMap[field] : title).toLowerCase();
      if (target.indexOf(query) === -1) continue;

      var normId = norm_(id);
      var status = discardedSet[normId] ? 'Discarded' : (issuedSet[normId] ? 'Issued' : 'Available');
      results.push({
        id: id, title: title, author: author,
        genre: bookGenre,
        location: shelfLoc,
        publisher: publisherVal, place: place,
        status: status
      });
    }
  }
  return results;
}

// ================================================================
// DISCARD BOOKS - single + bulk, writes to Decard Books, looks up
// Title/Author/Genre/Location automatically from Books Catalogue.
// Discarded status is derived (presence in Decard Books), not a
// separate flag - so Public dashboard stats update automatically,
// no extra step needed.
// ================================================================

var DECARD_START_ROW = 6;
// Expanded so a discarded book carries its FULL bibliographic record -
// needed for the committee/chairperson approval & write-off process,
// not just enough to identify it. All of these auto-fill from Books
// Catalogue the moment a book is discarded (see writeDiscardRow_) -
// nothing to type by hand except Reason and Remarks.
var DECARD_COL = {
  id: 1, title: 2, author: 3, publisher: 4, place: 5, year: 6,
  callno: 7, karwarno: 8, pages: 9, price: 10, genre: 11, location: 12,
  barcode: 13, decardDate: 14, staffName: 15, reason: 16, remarks: 17
};
var DECARD_LAST_COL = 17;

function decardLastRealRow_(sh) {
  var last = sh.getLastRow();
  var numRows = Math.max(last - DECARD_START_ROW + 1, 0);
  if (!numRows) return DECARD_START_ROW - 1;
  var col = sh.getRange(DECARD_START_ROW, DECARD_COL.id, numRows, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]).trim() !== '') return DECARD_START_ROW + i;
  }
  return DECARD_START_ROW - 1;
}

function lookupBookForDiscard_(booksSheet, bookId) {
  var last = booksSheet.getLastRow();
  var numRows = Math.max(last - BOOKS_START_ROW + 1, 0);
  if (!numRows) return null;
  var data = booksSheet.getRange(BOOKS_START_ROW, 1, numRows, BOOKS_LAST_COL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (norm_(data[i][BOOKS_COL.id - 1]) === bookId) {
      return {
        title: data[i][BOOKS_COL.title - 1] || '',
        author: data[i][BOOKS_COL.author - 1] || '',
        genre: data[i][BOOKS_COL.genre - 1] || '',
        location: data[i][BOOKS_COL.shelfLocation - 1] || '',
        publisher: data[i][BOOKS_COL.publisher - 1] || '',
        place: data[i][BOOKS_COL.place - 1] || '',
        year: data[i][BOOKS_COL.year - 1] || '',
        pages: data[i][BOOKS_COL.pages - 1] || '',
        price: data[i][BOOKS_COL.price - 1] || '',
        callno: data[i][BOOKS_COL.callno - 1] || '',
        karwarno: data[i][BOOKS_COL.karwarno - 1] || '',
        barcode: data[i][BOOKS_COL.barcode - 1] || '',
        row: BOOKS_START_ROW + i
      };
    }
  }
  return null;
}

function buildDiscardedSet_(decardSheet) {
  var last = decardSheet.getLastRow();
  var numRows = Math.max(last - DECARD_START_ROW + 1, 0);
  var set = {};
  if (numRows) {
    decardSheet.getRange(DECARD_START_ROW, DECARD_COL.id, numRows, 1).getValues().forEach(function(r) {
      var id = norm_(r[0]);
      if (id) set[id] = true;
    });
  }
  return set;
}

function getDiscardReasons() {
  return ['Damaged', 'Torn / Worn Out', 'Lost', 'Water Damaged', 'Missing Pages', 'Outdated', 'Duplicate Copy', 'Withdrawn', 'Other'];
}

// Writes one discard row. Preserves any existing auto-formula in
// Title/Author (matching your sheet's "(auto)" columns) - only writes
// plain values there if that cell has no formula already.
function writeDiscardRow_(decardSheet, row, bookId, book, reason, remarks, staffName) {
  decardSheet.getRange(row, DECARD_COL.id).setValue(bookId);

  var titleCell = decardSheet.getRange(row, DECARD_COL.title);
  if (!titleCell.getFormula()) titleCell.setValue(book.title);
  var authorCell = decardSheet.getRange(row, DECARD_COL.author);
  if (!authorCell.getFormula()) authorCell.setValue(book.author);

  decardSheet.getRange(row, DECARD_COL.genre).setValue(book.genre);
  decardSheet.getRange(row, DECARD_COL.location).setValue(book.location);
  decardSheet.getRange(row, DECARD_COL.publisher).setValue(book.publisher);
  decardSheet.getRange(row, DECARD_COL.place).setValue(book.place);
  decardSheet.getRange(row, DECARD_COL.year).setValue(book.year);
  decardSheet.getRange(row, DECARD_COL.pages).setValue(book.pages);
  decardSheet.getRange(row, DECARD_COL.price).setValue(book.price);
  decardSheet.getRange(row, DECARD_COL.callno).setValue(book.callno);
  decardSheet.getRange(row, DECARD_COL.karwarno).setValue(book.karwarno);
  decardSheet.getRange(row, DECARD_COL.barcode).setValue(book.barcode);
  decardSheet.getRange(row, DECARD_COL.decardDate).setValue(new Date());
  decardSheet.getRange(row, DECARD_COL.staffName).setValue(staffName || '');
  decardSheet.getRange(row, DECARD_COL.reason).setValue(reason);
  decardSheet.getRange(row, DECARD_COL.remarks).setValue(remarks || '');
}

// ================================================================
// ONE-TIME MIGRATION - run this ONCE by hand from the Apps Script
// editor (select it in the function dropdown, click Run), NOT from
// any button in the app. It rebuilds your existing ~5,600 Decard
// Books rows into the new 16-column full-detail layout, looking up
// each book's Publisher/Place/Year/Call No./Karwar No./Pages/Price/
// Barcode from Books Catalogue by its Book ID - the same fields new
// discards now get automatically. Your existing Decard Date, Reason
// and Remarks for each row are kept exactly as they were.
//
// Safe to run more than once by accident: it checks the header row
// first and refuses to run again once the sheet is already in the
// new layout, so it can't double-migrate or corrupt data.
// ================================================================
function migrateDecardBooksToFullDetails_ONE_TIME() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var decard = ss.getSheetByName('Decard Books');
  var books = ss.getSheetByName('Books Catalogue');

  var headerRow = 5;
  var col4Header = String(decard.getRange(headerRow, 4).getValue() || '').trim();
  if (col4Header !== 'Genre') {
    Logger.log('Column D header is "' + col4Header + '", not "Genre" - this sheet does not look like ' +
      'the OLD 8-column layout any more (maybe already migrated?). Stopping without changing anything.');
    return;
  }

  var last = decard.getLastRow();
  var numRows = Math.max(last - DECARD_START_ROW + 1, 0);
  if (!numRows) { Logger.log('No discard rows found - nothing to migrate.'); return; }

  // Old 8-column layout, read before we overwrite anything.
  var OLD = { id: 1, title: 2, author: 3, genre: 4, location: 5, decardDate: 6, reason: 7, remarks: 8 };
  var oldData = decard.getRange(DECARD_START_ROW, 1, numRows, 8).getValues();

  var newHeaders = ['Book ID (Acc.No.)', 'Title (auto)', 'Author (auto)', 'Publisher (auto)', 'Place (auto)',
    'Year (auto)', 'Call No. (auto)', 'Karwar No. (auto)', 'Pages (auto)', 'Price (auto)',
    'Genre (auto)', 'Shelf Location (auto)', 'Barcode (auto)', 'Decard Date', 'Reason', 'Remarks'];

  // Read Books Catalogue ONCE into an in-memory {id: details} map, instead
  // of re-scanning all ~40,000+ rows separately for each of the ~5,600
  // discarded books - that one-lookup-per-book approach would re-read the
  // whole catalogue thousands of times and almost certainly blow past
  // Apps Script's 6-minute execution limit partway through.
  var booksLast = books.getLastRow();
  var booksNumRows = Math.max(booksLast - BOOKS_START_ROW + 1, 0);
  var booksMap = {};
  if (booksNumRows) {
    var booksData = books.getRange(BOOKS_START_ROW, 1, booksNumRows, BOOKS_LAST_COL).getValues();
    for (var bi = 0; bi < booksData.length; bi++) {
      var bid = norm_(booksData[bi][BOOKS_COL.id - 1]);
      if (!bid) continue;
      booksMap[bid] = {
        title: booksData[bi][BOOKS_COL.title - 1] || '',
        author: booksData[bi][BOOKS_COL.author - 1] || '',
        genre: booksData[bi][BOOKS_COL.genre - 1] || '',
        location: booksData[bi][BOOKS_COL.shelfLocation - 1] || '',
        publisher: booksData[bi][BOOKS_COL.publisher - 1] || '',
        place: booksData[bi][BOOKS_COL.place - 1] || '',
        year: booksData[bi][BOOKS_COL.year - 1] || '',
        pages: booksData[bi][BOOKS_COL.pages - 1] || '',
        price: booksData[bi][BOOKS_COL.price - 1] || '',
        callno: booksData[bi][BOOKS_COL.callno - 1] || '',
        karwarno: booksData[bi][BOOKS_COL.karwarno - 1] || '',
        barcode: booksData[bi][BOOKS_COL.barcode - 1] || ''
      };
    }
  }

  var newRows = [];
  var notFoundIds = [];
  oldData.forEach(function(r) {
    var id = norm_(r[OLD.id - 1]);
    if (!id) { newRows.push(new Array(DECARD_LAST_COL).fill('')); return; }

    var book = booksMap[id];
    if (!book) notFoundIds.push(id);

    var row = new Array(DECARD_LAST_COL).fill('');
    row[DECARD_COL.id - 1] = r[OLD.id - 1];
    // Prefer a fresh Books Catalogue lookup; fall back to whatever was
    // already in this row if the book can no longer be found there.
    row[DECARD_COL.title - 1] = book ? book.title : r[OLD.title - 1];
    row[DECARD_COL.author - 1] = book ? book.author : r[OLD.author - 1];
    row[DECARD_COL.genre - 1] = book ? book.genre : r[OLD.genre - 1];
    row[DECARD_COL.location - 1] = book ? book.location : r[OLD.location - 1];
    row[DECARD_COL.publisher - 1] = book ? book.publisher : '';
    row[DECARD_COL.place - 1] = book ? book.place : '';
    row[DECARD_COL.year - 1] = book ? book.year : '';
    row[DECARD_COL.pages - 1] = book ? book.pages : '';
    row[DECARD_COL.price - 1] = book ? book.price : '';
    row[DECARD_COL.callno - 1] = book ? book.callno : '';
    row[DECARD_COL.karwarno - 1] = book ? book.karwarno : '';
    row[DECARD_COL.barcode - 1] = book ? book.barcode : '';
    row[DECARD_COL.decardDate - 1] = r[OLD.decardDate - 1];
    row[DECARD_COL.reason - 1] = r[OLD.reason - 1];
    row[DECARD_COL.remarks - 1] = r[OLD.remarks - 1];
    newRows.push(row);
  });

  // Widen the sheet if it doesn't already have 16 columns, then write
  // the new header row and every rebuilt data row in one batch call.
  if (decard.getMaxColumns() < DECARD_LAST_COL) {
    decard.insertColumnsAfter(decard.getMaxColumns(), DECARD_LAST_COL - decard.getMaxColumns());
  }
  decard.getRange(headerRow, 1, 1, DECARD_LAST_COL).setValues([newHeaders]);
  decard.getRange(DECARD_START_ROW, 1, newRows.length, DECARD_LAST_COL).setValues(newRows);

  Logger.log('Migrated ' + newRows.length + ' rows to the full-detail layout.');
  if (notFoundIds.length) {
    Logger.log(notFoundIds.length + ' book ID(s) were NOT found in Books Catalogue (kept their old ' +
      'Title/Author/Genre/Location text as-is, new fields left blank): ' + notFoundIds.join(', '));
  }
}

// ================================================================
// ONE-TIME - run this ONCE from the Apps Script editor, AFTER you've
// already run migrateDecardBooksToFullDetails_ONE_TIME above (this one
// only makes sense on a sheet that's already in the 16-column full-
// detail layout). It inserts a new "Staff Name" column just before
// Reason - Reason and Remarks (and every existing row's values in
// them) shift one column to the right automatically; nothing is lost.
// Existing discard rows will simply have this new column blank, since
// there's no way to know after the fact who did an old discard - only
// discards made from now on (via the app) will fill it in.
//
// Safe to run more than once by accident - it checks first and
// refuses to insert a second column if this has already been done.
// ================================================================
function addStaffNameColumnToDecardBooks_ONE_TIME() {
  var decard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Decard Books');
  var headerRow = 5;
  var col15Header = String(decard.getRange(headerRow, 15).getValue() || '').trim();

  if (col15Header === 'Staff Name') {
    Logger.log('Column O is already "Staff Name" - this has already been done. Stopping without changing anything.');
    return;
  }
  if (col15Header !== 'Reason') {
    Logger.log('Column O header is "' + col15Header + '", not "Reason" as expected. This sheet may not yet be ' +
      'in the full-detail layout - run migrateDecardBooksToFullDetails_ONE_TIME first if you haven\'t already. Stopping without changing anything.');
    return;
  }

  decard.insertColumnBefore(15);
  decard.getRange(headerRow, 15).setValue('Staff Name');
  Logger.log('Done - inserted the "Staff Name" column. Reason and Remarks (and all their existing values) shifted one column right automatically.');
}

function discardBookSingle(bookId, reason, remarks, staffName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var books = ss.getSheetByName('Books Catalogue');
  var decard = ss.getSheetByName('Decard Books');
  bookId = norm_(bookId);
  if (!bookId) throw new Error('Enter an Accession Number.');
  if (!reason) throw new Error('Select a reason.');
  staffName = String(staffName || '').trim();
  if (!staffName) throw new Error('Enter your Staff Name - needed for the discard record.');

  var book = lookupBookForDiscard_(books, bookId);
  if (!book) throw new Error('Book ID ' + bookId + ' not found in Books Catalogue.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var discardedSet = buildDiscardedSet_(decard);
    if (discardedSet[bookId]) throw new Error('Book ID ' + bookId + ' is already in Decard Books.');

    var row = decardLastRealRow_(decard) + 1;
    writeDiscardRow_(decard, row, bookId, book, reason, remarks, staffName);
    books.getRange(book.row, BOOKS_COL.status).setValue('Discarded');

    return { bookId: bookId, title: book.title, row: row };
  } finally {
    lock.releaseLock();
  }
}

function discardBooksBulk(bookIdsText, reason, remarks, staffName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var books = ss.getSheetByName('Books Catalogue');
  var decard = ss.getSheetByName('Decard Books');
  if (!reason) throw new Error('Select a reason.');
  staffName = String(staffName || '').trim();
  if (!staffName) throw new Error('Enter your Staff Name - needed for the discard record.');

  var ids = String(bookIdsText || '').split('\n').map(function(s) { return norm_(s); }).filter(Boolean);
  if (!ids.length) throw new Error('Enter at least one Accession Number.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var discardedSet = buildDiscardedSet_(decard);
    var nextRow = decardLastRealRow_(decard) + 1;
    var success = [], alreadyDiscarded = [], notFound = [];

    ids.forEach(function(id) {
      if (discardedSet[id]) { alreadyDiscarded.push(id); return; }
      var book = lookupBookForDiscard_(books, id);
      if (!book) { notFound.push(id); return; }
      writeDiscardRow_(decard, nextRow, id, book, reason, remarks, staffName);
      books.getRange(book.row, BOOKS_COL.status).setValue('Discarded');
      discardedSet[id] = true; // guard against duplicate IDs within the same paste
      success.push({ bookId: id, title: book.title });
      nextRow++;
    });

    return { success: success, alreadyDiscarded: alreadyDiscarded, notFound: notFound };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// GATE ENTRY KIOSK - auto-detects Member vs Non-Member, auto-decides
// IN vs OUT from today's last entry, computes Reading Time on OUT.
// Writes to your existing GATE_ENTRY sheet - no new log sheet.
// Member VIP/welcome-message/photo come from the separate
// MEMBER_GATE_PROFILE sheet (kept apart from core Members data).
// ================================================================

var GATE_COL = { date:1, time:2, id:3, name:4, type:5, vip:6, message:7,
                  inOut:8, photoPath:9, season:10, remarks:11, readingTime:12 };
var GATE_START_ROW = 2;
var GATE_LAST_COL = 12;

// Your two real library sessions.
var SESSION1_START_MIN = 8*60;      // 08:00
var SESSION1_END_MIN = 11*60+30;    // 11:30
var SESSION2_START_MIN = 16*60;     // 16:00
var SESSION2_END_MIN = 20*60;       // 20:00

// Auto-close waits this long PAST the session's real end before treating a
// forgotten check-out as closeable - gives late readers a grace window
// instead of cutting them off the instant the session officially ends.
// Morning: 11:30 + 1:00 = 12:30. Evening: 20:00 + 0:30 = 20:30.
var SESSION1_CLOSE_GRACE_MIN = 60;
var SESSION2_CLOSE_GRACE_MIN = 30;

function gateCurrentSeason_(now) {
  var mins = now.getHours()*60 + now.getMinutes();
  return mins < SESSION2_START_MIN ? 'Morning' : 'Evening';
}

function parseReadingTimeToSeconds_(str) {
  if (!str) return 0;
  var parts = String(str).split(':').map(Number);
  if (parts.some(isNaN)) return 0;
  if (parts.length === 3) return parts[0]*3600 + parts[1]*60 + parts[2];
  if (parts.length === 2) return parts[0]*60 + parts[1];
  return parts[0] || 0;
}

function formatSecondsToHMS_(total) {
  var h = Math.floor(total/3600), m = Math.floor((total%3600)/60), s = Math.floor(total%60);
  return (h>0?h+':':'') + (h>0&&m<10?'0':'') + m + ':' + (s<10?'0':'') + s;
}

function gateLastRealRow_(sh) {
  var last = sh.getLastRow();
  var numRows = Math.max(last - GATE_START_ROW + 1, 0);
  if (!numRows) return GATE_START_ROW - 1;
  var col = sh.getRange(GATE_START_ROW, GATE_COL.id, numRows, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]).trim() !== '') return GATE_START_ROW + i;
  }
  return GATE_START_ROW - 1;
}

function gateLookupProfile_(id) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MEMBER_GATE_PROFILE');
  if (!sh) return { vip: '', message: '', photoPath: '' };
  var last = sh.getLastRow();
  if (last < 2) return { vip: '', message: '', photoPath: '' };
  var vals = sh.getRange(2, 1, last - 1, 6).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (norm_(vals[i][0]) === id) {
      return { vip: vals[i][1] || '', message: vals[i][2] || '', photoPath: vals[i][3] || '' };
    }
  }
  return { vip: '', message: '', photoPath: '' };
}

function gateLookupPerson_(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var members = ss.getSheetByName('Members');
  var mLast = members.getLastRow();
  if (mLast >= MEMBERS_START_ROW) {
    var mVals = members.getRange(MEMBERS_START_ROW, 1, mLast - MEMBERS_START_ROW + 1, MEMBERS_LAST_COL).getValues();
    for (var i = 0; i < mVals.length; i++) {
      if (norm_(mVals[i][MEMBERS_COL.id - 1]) === id) {
        var profile = gateLookupProfile_(id);
        return { id: id, name: mVals[i][MEMBERS_COL.name - 1], type: 'Member',
                 vip: profile.vip, message: profile.message, photoPath: profile.photoPath };
      }
    }
  }

  var nonmem = ss.getSheetByName('NON_MEMBERS');
  var nLast = nonmem.getLastRow();
  if (nLast >= NONMEM_START_ROW) {
    var nVals = nonmem.getRange(NONMEM_START_ROW, 1, nLast - NONMEM_START_ROW + 1, NONMEM_LAST_COL).getValues();
    for (var j = 0; j < nVals.length; j++) {
      if (norm_(nVals[j][NONMEM_COL.id - 1]) === id) {
        return { id: id, name: nVals[j][NONMEM_COL.name - 1], type: 'NonMember',
                 vip: nVals[j][NONMEM_COL.vip - 1] || '', message: nVals[j][NONMEM_COL.specialMsg - 1] || '',
                 photoPath: nVals[j][NONMEM_COL.photoPath - 1] || '' };
      }
    }
  }
  return null;
}

function gateGetTodaysLastEntryInSession_(sh, id, season) {
  var lastRow = gateLastRealRow_(sh);
  var numRows = Math.max(lastRow - GATE_START_ROW + 1, 0);
  if (!numRows) return null;
  var vals = sh.getRange(GATE_START_ROW, 1, numRows, GATE_LAST_COL).getValues();
  var today = stripTime_(new Date());
  var lastMatch = null;
  vals.forEach(function(r) {
    var d = r[GATE_COL.date - 1];
    if (!(d instanceof Date) || stripTime_(d).getTime() !== today.getTime()) return;
    if (norm_(r[GATE_COL.id - 1]) !== id) return;
    if (r[GATE_COL.season - 1] !== season) return;
    lastMatch = r; // rows are in chronological append order - last match wins
  });
  return lastMatch;
}

// Haversine distance in meters between two lat/lng points.
function distanceMeters_(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLng = (lng2 - lng1) * Math.PI / 180;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function gateCheckInOut(rawId, lat, lng) {
  var id = norm_(rawId);
  if (!id) throw new Error('Enter or scan an ID.');
  var person = gateLookupPerson_(id);
  if (!person) throw new Error('ID ' + id + ' not found in Members or Non-Members.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('GATE_ENTRY');
  var now = new Date();
  var season = gateCurrentSeason_(now); // Morning = Session 1, Evening = Session 2

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Optional location enforcement - off by default (SETTINGS_CARDS!B12).
    // Set to "Yes" only once you're ready to go public and want to actually
    // reject far-away check-ins, not just record distance.
    var stCheck = ss.getSheetByName('SETTINGS_CARDS');
    var enforceLocation = stCheck && String(stCheck.getRange('B12').getValue() || '').trim().toLowerCase() === 'yes';
    if (enforceLocation) {
      var offLat = stCheck.getRange('B9').getValue();
      var offLng = stCheck.getRange('B10').getValue();
      var maxDist = Number(stCheck.getRange('B11').getValue()) || 100;
      if (!lat || !lng) {
        throw new Error('Location access is required to check in here - please allow location access and try again.');
      }
      if (offLat && offLng) {
        var dist = distanceMeters_(Number(lat), Number(lng), Number(offLat), Number(offLng));
        if (dist > maxDist) {
          throw new Error('You appear to be ' + Math.round(dist) + 'm from the library - check-in is only allowed within ' + maxDist + 'm.');
        }
      }
    }

    var lastInSession = gateGetTodaysLastEntryInSession_(sh, id, season);

    var isIn = !lastInSession || String(lastInSession[GATE_COL.inOut - 1]).toUpperCase() === 'OUT';
    var readingTime = '';

    if (!isIn && lastInSession) {
      var inTimeCell = lastInSession[GATE_COL.time - 1];
      if (inTimeCell instanceof Date) {
        var inMoment = new Date(now.getFullYear(), now.getMonth(), now.getDate(),
                                 inTimeCell.getHours(), inTimeCell.getMinutes(), inTimeCell.getSeconds());
        var diffMs = Math.max(now.getTime() - inMoment.getTime(), 0);
        readingTime = formatSecondsToHMS_(Math.floor(diffMs / 1000));
      }
    }

    var row = gateLastRealRow_(sh) + 1;
    var remarks = '';
    var distanceMeters = null;
    if (lat && lng) {
      remarks = 'GPS: ' + Number(lat).toFixed(5) + ',' + Number(lng).toFixed(5);
      var st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SETTINGS_CARDS');
      var officeLat = st ? st.getRange('B9').getValue() : '';
      var officeLng = st ? st.getRange('B10').getValue() : '';
      if (officeLat && officeLng) {
        distanceMeters = Math.round(distanceMeters_(Number(lat), Number(lng), Number(officeLat), Number(officeLng)));
        remarks += ' (' + distanceMeters + 'm from office)';
      }
    }
    sh.getRange(row, 1, 1, GATE_LAST_COL).setValues([[
      now, now, id, person.name, person.type, person.vip, person.message,
      isIn ? 'IN' : 'OUT', person.photoPath, season, remarks, readingTime
    ]]);

    var settings = getInitData().settings;
    var photo = resolvePhotoPath_(person.photoPath) || findPhotoByFolder_(settings.memberPhotoFolderId, id);

    return {
      id: id, name: person.name, type: person.type, vip: person.vip, message: person.message,
      photo: photo, inOut: isIn ? 'IN' : 'OUT', readingTime: readingTime, season: season,
      distanceMeters: distanceMeters
    };
  } finally {
    lock.releaseLock();
  }
}

// ---- Auto-close forgotten checkouts -------------------------------------
// Closes anyone still IN past their session's end time, tagged clearly as
// auto-closed. Call manually (button) or via an installed time trigger.

// UPDATED: now walks the WHOLE Gate_Entry history sequentially, person by
// person, day by day - needed because gate data doesn't only arrive live
// through the kiosk any more. Most of it now gets pasted in bulk from an
// offline VBA/Excel tool, sometimes a full day (or several days) at a time,
// well after the fact, and that pasted data has NO Season/session label
// filled in (columns D/E/J are blank) - only Date, Time, ID and IN/OUT.
//
// Two real-world patterns this has to handle correctly (both confirmed):
//  1. Someone can visit twice in the SAME session (e.g. in at 8:00, out at
//     9:00, back in at 10:00) - a second IN for the same person on the same
//     day, before any OUT closed the first one, means the first visit was
//     really over the moment the second IN happened; it's auto-closed AT
//     that second IN's own timestamp (not at session-end), tagged
//     "checked in again before checking out". The second IN then becomes
//     the new open visit.
//  2. Someone who is IN and simply never has a matching OUT anywhere after
//     it (kiosk forgot, or the offline tool never logged the exit) gets
//     auto-closed at that session's grace cutoff instead: Morning session
//     closes at 11:30 + 1:00 = 12:30 PM, Evening session closes at
//     20:00 + 0:30 = 8:30 PM (SESSION1_CLOSE_GRACE_MIN / _2_, set above -
//     change those two numbers to retune this without touching the logic
//     below). Which session applies is worked out from the IN's own
//     time-of-day (before/after 4:00 PM), never from the Season column,
//     since bulk-pasted rows don't have one.
//
// For TODAY's dangling IN (case 2), it's only closed once that grace period
// has actually passed - so someone genuinely still reading right now is
// never prematurely checked out. Any PAST day's dangling IN is always
// closeable (that day is already over), and gets stamped with THAT day's
// own date - never today's - which matters once this runs against
// back-dated bulk-pasted rows.
function gateAutoCloseSessions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('GATE_ENTRY');
  var lastRow = gateLastRealRow_(sh);
  var numRows = Math.max(lastRow - GATE_START_ROW + 1, 0);
  if (!numRows) return { closed: 0 };

  var vals = sh.getRange(GATE_START_ROW, 1, numRows, GATE_LAST_COL).getValues();
  var today = stripTime_(new Date());
  var now = new Date();
  var nowMinutes = now.getHours() * 60 + now.getMinutes();
  var closeRows = [];

  // Combines a date-only Date with a separate time-of-day cell into one
  // real Date/time moment.
  function combine_(dateOnly, timeCell) {
    if (timeCell instanceof Date) {
      return new Date(dateOnly.getFullYear(), dateOnly.getMonth(), dateOnly.getDate(),
                       timeCell.getHours(), timeCell.getMinutes(), timeCell.getSeconds());
    }
    return new Date(dateOnly.getFullYear(), dateOnly.getMonth(), dateOnly.getDate());
  }

  // Works out Morning vs Evening from the IN's own time-of-day - the
  // Season column is trusted only when it's already a valid label,
  // otherwise (blank, as with all bulk-pasted rows) it's derived instead.
  function seasonFor_(storedSeason, inTimeCell) {
    if (storedSeason === 'Morning' || storedSeason === 'Evening') return storedSeason;
    if (inTimeCell instanceof Date) {
      var mins = inTimeCell.getHours() * 60 + inTimeCell.getMinutes();
      return mins < SESSION2_START_MIN ? 'Morning' : 'Evening';
    }
    return 'Evening';
  }

  function buildOutRow_(entry, closeMoment, season, remarks) {
    var inTimeCell = entry[GATE_COL.time - 1];
    var readingTime = '';
    if (inTimeCell instanceof Date) {
      var entryDate = stripTime_(entry[GATE_COL.date - 1]);
      var inMoment = combine_(entryDate, inTimeCell);
      readingTime = formatSecondsToHMS_(Math.max(Math.floor((closeMoment - inMoment) / 1000), 0));
    }
    return [closeMoment, closeMoment, entry[GATE_COL.id - 1], entry[GATE_COL.name - 1],
      entry[GATE_COL.type - 1], entry[GATE_COL.vip - 1], entry[GATE_COL.message - 1], 'OUT',
      entry[GATE_COL.photoPath - 1], season, remarks, readingTime];
  }

  // Sequential per-person-per-day pairing: walk every row in sheet order
  // (chronological within a person+day), tracking whichever IN is
  // currently "open" for that person on that day.
  var openByKey = {}; // "dayKey|personId" -> the still-open IN row
  vals.forEach(function(r) {
    var d = r[GATE_COL.date - 1];
    if (!(d instanceof Date) || isNaN(d.getTime())) return;
    var dayKey = stripTime_(d).getTime();
    var id = norm_(r[GATE_COL.id - 1]);
    if (!id) return;
    var key = dayKey + '|' + id;
    var inOut = String(r[GATE_COL.inOut - 1] || '').toUpperCase();

    if (inOut === 'IN') {
      var openEntry = openByKey[key];
      if (openEntry) {
        // Second IN before any OUT closed the first - that first visit was
        // really over the moment this new IN happened.
        var entryDate = stripTime_(openEntry[GATE_COL.date - 1]);
        var closeMoment = combine_(entryDate, r[GATE_COL.time - 1]);
        var season = seasonFor_(openEntry[GATE_COL.season - 1], openEntry[GATE_COL.time - 1]);
        closeRows.push(buildOutRow_(openEntry, closeMoment, season, 'Auto-closed - checked in again before checking out'));
      }
      openByKey[key] = r;
    } else if (inOut === 'OUT') {
      if (openByKey[key]) delete openByKey[key]; // real, explicit close
    }
  });

  // Anything still open after walking every row is a genuinely dangling
  // IN - close it at its session's grace cutoff, per the rules above.
  Object.keys(openByKey).forEach(function(key) {
    var entry = openByKey[key];
    var entryDate = stripTime_(entry[GATE_COL.date - 1]);
    if (entryDate.getTime() > today.getTime()) return; // data error - never touch a future-dated row

    var isToday = entryDate.getTime() === today.getTime();
    var season = seasonFor_(entry[GATE_COL.season - 1], entry[GATE_COL.time - 1]);
    var sessionEndMin = season === 'Morning'
      ? (SESSION1_END_MIN + SESSION1_CLOSE_GRACE_MIN)
      : (SESSION2_END_MIN + SESSION2_CLOSE_GRACE_MIN);

    // Only skip when it's TODAY and the grace period hasn't passed yet -
    // any earlier day is always closeable, since that day is already over.
    if (isToday && nowMinutes < sessionEndMin) return;

    var endTime = new Date(entryDate.getFullYear(), entryDate.getMonth(), entryDate.getDate(),
                            Math.floor(sessionEndMin / 60), sessionEndMin % 60, 0);
    closeRows.push(buildOutRow_(entry, endTime, season, 'Auto-closed - checkout not recorded'));
  });

  if (closeRows.length) {
    var startRow = gateLastRealRow_(sh) + 1;
    sh.getRange(startRow, 1, closeRows.length, GATE_LAST_COL).setValues(closeRows);
  }
  return { closed: closeRows.length };
}

// ---- Footfall & Time-Spent Awards ---------------------------------------

function refreshFootfallAwards() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('GATE_ENTRY');
  var awards = ss.getSheetByName('FOOTFALL AWARDS');
  if (!awards) awards = ss.insertSheet('FOOTFALL AWARDS');

  var lastRow = gateLastRealRow_(sh);
  var numRows = Math.max(lastRow - GATE_START_ROW + 1, 0);
  var rows = numRows ? sh.getRange(GATE_START_ROW, 1, numRows, GATE_LAST_COL).getValues() : [];

  var today = stripTime_(new Date());
  var yesterday = new Date(today); yesterday.setDate(yesterday.getDate() - 1);
  var dow = today.getDay();
  var weekStart = new Date(today.getFullYear(), today.getMonth(), today.getDate() - (dow === 0 ? 6 : dow - 1));
  var monthStart = new Date(today.getFullYear(), today.getMonth(), 1);

  var footfallToday = 0, footfallWeek = 0, footfallMonth = 0, footfallAllTime = 0;
  var visitWeek = {}, visitMonth = {}, visitAllTime = {};
  var timeYesterday = {}, timeWeek = {}, timeMonth = {}, timeAllTime = {};
  var nameById = {};

  rows.forEach(function(r) {
    var d = r[GATE_COL.date - 1];
    if (!(d instanceof Date)) return;
    var dd = stripTime_(d);
    var id = norm_(r[GATE_COL.id - 1]);
    if (id) nameById[id] = r[GATE_COL.name - 1];
    var inOut = String(r[GATE_COL.inOut - 1] || '').toUpperCase();

    if (inOut === 'IN') {
      footfallAllTime++;
      if (dd.getTime() === today.getTime()) footfallToday++;
      if (dd >= weekStart) { footfallWeek++; visitWeek[id] = (visitWeek[id]||0) + 1; }
      if (dd >= monthStart) { footfallMonth++; visitMonth[id] = (visitMonth[id]||0) + 1; }
      visitAllTime[id] = (visitAllTime[id]||0) + 1;
    }

    if (inOut === 'OUT' && id) {
      var secs = parseReadingTimeToSeconds_(r[GATE_COL.readingTime - 1]);
      if (secs > 0) {
        if (dd.getTime() === yesterday.getTime()) timeYesterday[id] = (timeYesterday[id]||0) + secs;
        if (dd >= weekStart) timeWeek[id] = (timeWeek[id]||0) + secs;
        if (dd >= monthStart) timeMonth[id] = (timeMonth[id]||0) + secs;
        timeAllTime[id] = (timeAllTime[id]||0) + secs;
      }
    }
  });

  function topN(map, n) {
    return Object.keys(map).map(function(id) { return { id: id, name: nameById[id] || '', value: map[id] }; })
      .sort(function(a, b) { return b.value - a.value; }).slice(0, n);
  }

  awards.clear();
  awards.getRange(1, 1).setValue('FOOTFALL & TIME-SPENT AWARDS').setFontWeight('bold').setFontSize(13);

  var overview = [
    ['Footfall Today', footfallToday],
    ['Footfall This Week', footfallWeek],
    ['Footfall This Month', footfallMonth],
    ['Footfall All-Time', footfallAllTime]
  ];
  awards.getRange(3, 1, overview.length, 2).setValues(overview);
  awards.getRange(3, 1, overview.length, 1).setFontWeight('bold');

  var row = 3 + overview.length + 2;
  awards.getRange(row, 1).setValue('MOST VISITED (Footfall Count)').setFontWeight('bold');
  row++;
  awards.getRange(row, 1, 1, 4).setValues([['Period', 'Rank', 'ID - Name', 'Visits']]).setFontWeight('bold');
  row++;
  [['This Week', visitWeek], ['This Month', visitMonth], ['All-Time', visitAllTime]].forEach(function(pair) {
    var top = topN(pair[1], 5);
    if (!top.length) { awards.getRange(row, 1, 1, 4).setValues([[pair[0], '-', '(no data)', '']]); row++; return; }
    top.forEach(function(t, i) {
      awards.getRange(row, 1, 1, 4).setValues([[pair[0], i + 1, t.id + ' - ' + t.name, t.value]]);
      row++;
    });
  });

  row += 2;
  awards.getRange(row, 1).setValue('MOST TIME SPENT (Longest Cumulative Visits)').setFontWeight('bold');
  row++;
  awards.getRange(row, 1, 1, 4).setValues([['Period', 'Rank', 'ID - Name', 'Time Spent']]).setFontWeight('bold');
  row++;
  [['Yesterday', timeYesterday], ['This Week', timeWeek], ['This Month', timeMonth], ['All-Time', timeAllTime]].forEach(function(pair) {
    var top = topN(pair[1], 5);
    if (!top.length) { awards.getRange(row, 1, 1, 4).setValues([[pair[0], '-', '(no data)', '']]); row++; return; }
    top.forEach(function(t, i) {
      awards.getRange(row, 1, 1, 4).setValues([[pair[0], i + 1, t.id + ' - ' + t.name, formatSecondsToHMS_(t.value)]]);
      row++;
    });
  });

  awards.autoResizeColumns(1, 4);
  return { footfallToday: footfallToday, footfallWeek: footfallWeek, footfallMonth: footfallMonth, footfallAllTime: footfallAllTime };
}

// ================================================================
// ADD NEW MEMBER / NON-MEMBER - webcam photo capture, auto-incrementing
// IDs. For Non-Members, reuses an existing pre-allocated placeholder
// row (ID + PhotoPath already filled, Name still blank) if one exists,
// instead of always appending a brand new row.
// ================================================================

function membersLastRealRow_(sh) {
  var last = sh.getLastRow();
  var numRows = Math.max(last - MEMBERS_START_ROW + 1, 0);
  if (!numRows) return MEMBERS_START_ROW - 1;
  var col = sh.getRange(MEMBERS_START_ROW, MEMBERS_COL.id, numRows, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]).trim() !== '') return MEMBERS_START_ROW + i;
  }
  return MEMBERS_START_ROW - 1;
}

function getNextMemberId() {
  var sh = ss_().getSheetByName('Members');
  var last = membersLastRealRow_(sh);
  var numRows = Math.max(last - MEMBERS_START_ROW + 1, 0);
  var maxId = 0;
  if (numRows) {
    sh.getRange(MEMBERS_START_ROW, MEMBERS_COL.id, numRows, 1).getValues().forEach(function(r) {
      var n = parseInt(r[0], 10);
      if (!isNaN(n) && n > maxId) maxId = n;
    });
  }
  return maxId + 1;
}

function savePhotoToLibraryFolder_(id, photoBase64) {
  if (!photoBase64) return '';
  var init = getInitData().settings;
  var folderNameOrId = init.memberPhotoFolderId;
  var folder = null;
  try {
    var it = DriveApp.getFoldersByName(String(folderNameOrId).trim());
    if (it.hasNext()) folder = it.next();
  } catch (e) { /* try ID below */ }
  if (!folder) {
    try { folder = DriveApp.getFolderById(String(folderNameOrId).trim()); } catch (e) { return ''; }
  }
  if (!folder) return '';

  var base64Data = String(photoBase64).split(',').pop(); // strip data:image/jpeg;base64, prefix
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', id + '.jpg');

  // Replace any existing photo with the same name (re-registration case).
  var existing = folder.getFilesByName(id + '.jpg');
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var file = folder.createFile(blob);
  return file.getUrl();
}

function addNewMember(data) {
  if (!data || !String(data.name || '').trim()) throw new Error('Enter a name.');
  var sh = ss_().getSheetByName('Members');
  var id = getNextMemberId();
  var row = membersLastRealRow_(sh) + 1;

  // FIXED: this used to write a fixed 8-value literal into a
  // MEMBERS_LAST_COL-wide range - fine back when Members only had 8
  // columns, but silently out of sync ever since Age/Gender/Profession/
  // Qualification/Address/Group (and now Amount/Receipt_No/Receipt_Date/
  // Challan_Detail) were added, which would throw a range-width error the
  // moment "Add New Member" was used. Building a full-width array and
  // filling in just the known fields keeps this correct no matter how many
  // columns Members ends up with.
  var newRow = new Array(MEMBERS_LAST_COL).fill('');
  newRow[MEMBERS_COL.id - 1] = id;
  newRow[MEMBERS_COL.name - 1] = data.name || '';
  newRow[MEMBERS_COL.phone - 1] = data.phone || '';
  newRow[MEMBERS_COL.email - 1] = data.email || '';
  newRow[MEMBERS_COL.membStart - 1] = data.membStart ? new Date(data.membStart) : new Date();
  newRow[MEMBERS_COL.membEnd - 1] = data.membEnd ? new Date(data.membEnd) : '';
  newRow[MEMBERS_COL.status - 1] = data.status || 'Active';
  newRow[MEMBERS_COL.notes - 1] = data.notes || '';
  sh.getRange(row, 1, 1, MEMBERS_LAST_COL).setValues([newRow]);

  var photoUrl = savePhotoToLibraryFolder_(id, data.photoBase64);
  return { id: id, name: data.name, photoSaved: !!photoUrl };
}

function nonMemLastRealRow_(sh) {
  var last = sh.getLastRow();
  var numRows = Math.max(last - NONMEM_START_ROW + 1, 0);
  if (!numRows) return NONMEM_START_ROW - 1;
  var idCol = sh.getRange(NONMEM_START_ROW, NONMEM_COL.id, numRows, 1).getValues();
  var nameCol = sh.getRange(NONMEM_START_ROW, NONMEM_COL.name, numRows, 1).getValues();
  for (var i = idCol.length - 1; i >= 0; i--) {
    if (String(idCol[i][0]).trim() !== '' || String(nameCol[i][0]).trim() !== '') return NONMEM_START_ROW + i;
  }
  return NONMEM_START_ROW - 1;
}

// Finds the lowest-ID row that has an ID but no Name yet (a pre-allocated
// placeholder), or otherwise the next brand-new row/ID after the highest
// existing one. Read-only - doesn't write anything itself.
function getNextNonMemberSlot_() {
  var sh = ss_().getSheetByName('NON_MEMBERS');
  var last = nonMemLastRealRow_(sh);
  var numRows = Math.max(last - NONMEM_START_ROW + 1, 0);
  var maxId = 0, placeholderRow = null, placeholderId = null;

  if (numRows) {
    var idVals = sh.getRange(NONMEM_START_ROW, NONMEM_COL.id, numRows, 1).getValues();
    var nameVals = sh.getRange(NONMEM_START_ROW, NONMEM_COL.name, numRows, 1).getValues();
    for (var i = 0; i < idVals.length; i++) {
      var idNum = parseInt(idVals[i][0], 10);
      var hasName = String(nameVals[i][0]).trim() !== '';
      if (!isNaN(idNum)) {
        if (idNum > maxId) maxId = idNum;
        if (!hasName && placeholderId === null) {
          placeholderId = idNum;
          placeholderRow = NONMEM_START_ROW + i;
        }
      }
    }
  }
  if (placeholderRow) return { row: placeholderRow, id: placeholderId, isNewRow: false };
  return { row: last + 1, id: maxId + 1, isNewRow: true };
}

function getNextNonMemberPreview() {
  var slot = getNextNonMemberSlot_();
  return { id: slot.id, reusedPlaceholder: !slot.isNewRow };
}

function addNewNonMember(data) {
  if (!data || !String(data.name || '').trim()) throw new Error('Enter a name.');
  var sh = ss_().getSheetByName('NON_MEMBERS');
  var slot = getNextNonMemberSlot_();
  var photoUrl = savePhotoToLibraryFolder_(slot.id, data.photoBase64);

  sh.getRange(slot.row, 1, 1, NONMEM_LAST_COL).setValues([[
    slot.id, data.name || '', data.address || '', data.phone || '', data.purpose || '',
    data.vip || 'No', data.message || '',
    photoUrl || sh.getRange(slot.row, NONMEM_COL.photoPath).getValue() || '',
    new Date(), data.occupation || ''
  ]]);

  return { id: slot.id, name: data.name, photoSaved: !!photoUrl, reusedPlaceholder: !slot.isNewRow };
}

function getPublicAwards() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bi = ss.getSheetByName('Book Issue');
  var books = ss.getSheetByName('Books Catalogue');
  var members = ss.getSheetByName('Members');
  var gate = ss.getSheetByName('GATE_ENTRY');

  var today = stripTime_(new Date());
  var dow = today.getDay();
  var weekStart = new Date(today.getFullYear(), today.getMonth(), today.getDate() - (dow === 0 ? 6 : dow - 1));
  var monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  var yearStart = new Date(today.getFullYear(), 0, 1);

  // ---- Most borrowed books + top readers, from Book Issue ----
  var biLast = biLastRealRow_(bi);
  var biRows = Math.max(biLast - BI_START_ROW + 1, 0);
  var bookCountsWeek = {}, bookCountsMonth = {}, bookCountsYear = {};
  var readerCountsWeek = {}, readerCountsMonth = {}, readerCountsYear = {};

  if (biRows) {
    bi.getRange(BI_START_ROW, 1, biRows, BI_LAST_COL).getValues().forEach(function(r) {
      var d = r[BI_COL.issueDate - 1];
      if (!(d instanceof Date)) return;
      var dd = stripTime_(d);
      var bookId = norm_(r[BI_COL.bookId - 1]);
      var memberId = norm_(r[BI_COL.memberId - 1]);
      if (!bookId || !memberId) return;
      if (dd >= weekStart) { bookCountsWeek[bookId] = (bookCountsWeek[bookId]||0)+1; readerCountsWeek[memberId] = (readerCountsWeek[memberId]||0)+1; }
      if (dd >= monthStart) { bookCountsMonth[bookId] = (bookCountsMonth[bookId]||0)+1; readerCountsMonth[memberId] = (readerCountsMonth[memberId]||0)+1; }
      if (dd >= yearStart) { bookCountsYear[bookId] = (bookCountsYear[bookId]||0)+1; readerCountsYear[memberId] = (readerCountsYear[memberId]||0)+1; }
    });
  }

  function topIds_(map, n) {
    return Object.keys(map).sort(function(a, b) { return map[b] - map[a]; }).slice(0, n);
  }

  var neededBookIds = {};
  [bookCountsWeek, bookCountsMonth, bookCountsYear].forEach(function(m) {
    topIds_(m, 5).forEach(function(id) { neededBookIds[id] = true; });
  });
  var bookLookup = {};
  var bLast = books.getLastRow();
  var bNumRows = Math.max(bLast - BOOKS_START_ROW + 1, 0);
  if (bNumRows && Object.keys(neededBookIds).length) {
    books.getRange(BOOKS_START_ROW, 1, bNumRows, 3).getValues().forEach(function(r) {
      var id = norm_(r[0]);
      if (neededBookIds[id]) bookLookup[id] = { title: r[1] || '', author: r[2] || '' };
    });
  }
  function topBooks_(map, n) {
    return topIds_(map, n).map(function(id) {
      var t = bookLookup[id] || { title: '', author: '' };
      return { id: id, title: t.title, author: t.author, count: map[id] };
    });
  }

  var neededMemberIds = {};
  [readerCountsWeek, readerCountsMonth, readerCountsYear].forEach(function(m) {
    topIds_(m, 5).forEach(function(id) { neededMemberIds[id] = true; });
  });
  var memberNameLookup = {};
  var mLast = members.getLastRow();
  var mNumRows = Math.max(mLast - MEMBERS_START_ROW + 1, 0);
  if (mNumRows && Object.keys(neededMemberIds).length) {
    members.getRange(MEMBERS_START_ROW, 1, mNumRows, 2).getValues().forEach(function(r) {
      var id = norm_(r[0]);
      if (neededMemberIds[id]) memberNameLookup[id] = r[1] || '';
    });
  }
  function topReaders_(map, n) {
    return topIds_(map, n).map(function(id) { return { id: id, name: memberNameLookup[id] || '', count: map[id] }; });
  }

  // ---- Footfall + most visited + longest visit, from GATE_ENTRY ----
  var gLast = gateLastRealRow_(gate);
  var gRows = Math.max(gLast - GATE_START_ROW + 1, 0);
  var footfallToday = 0, footfallWeek = 0, footfallMonth = 0, footfallYear = 0;
  var visitWeek = {}, visitMonth = {}, visitYear = {};
  var timeWeek = {}, timeMonth = {}, timeYear = {};
  var gateNameById = {};

  if (gRows) {
    gate.getRange(GATE_START_ROW, 1, gRows, GATE_LAST_COL).getValues().forEach(function(r) {
      var d = r[GATE_COL.date - 1];
      if (!(d instanceof Date)) return;
      var dd = stripTime_(d);
      var id = norm_(r[GATE_COL.id - 1]);
      if (id) gateNameById[id] = r[GATE_COL.name - 1];
      var inOut = String(r[GATE_COL.inOut - 1] || '').toUpperCase();
      if (inOut === 'IN') {
        if (dd.getTime() === today.getTime()) footfallToday++;
        if (dd >= weekStart) { footfallWeek++; visitWeek[id] = (visitWeek[id]||0)+1; }
        if (dd >= monthStart) { footfallMonth++; visitMonth[id] = (visitMonth[id]||0)+1; }
        if (dd >= yearStart) { footfallYear++; visitYear[id] = (visitYear[id]||0)+1; }
      }
      if (inOut === 'OUT' && id) {
        var secs = parseReadingTimeToSeconds_(r[GATE_COL.readingTime - 1]);
        if (secs > 0) {
          if (dd >= weekStart) timeWeek[id] = (timeWeek[id]||0) + secs;
          if (dd >= monthStart) timeMonth[id] = (timeMonth[id]||0) + secs;
          if (dd >= yearStart) timeYear[id] = (timeYear[id]||0) + secs;
        }
      }
    });
  }

  function topVisits_(map, n) {
    return topIds_(map, n).map(function(id) { return { id: id, name: gateNameById[id] || '', value: map[id] }; });
  }
  function topTimes_(map, n) {
    return topIds_(map, n).map(function(id) { return { id: id, name: gateNameById[id] || '', value: formatSecondsToHMS_(map[id]) }; });
  }

  return {
    books: { week: topBooks_(bookCountsWeek, 5), month: topBooks_(bookCountsMonth, 5), year: topBooks_(bookCountsYear, 5) },
    readers: { week: topReaders_(readerCountsWeek, 5), month: topReaders_(readerCountsMonth, 5), year: topReaders_(readerCountsYear, 5) },
    footfall: { today: footfallToday, week: footfallWeek, month: footfallMonth, year: footfallYear },
    mostVisited: { week: topVisits_(visitWeek, 5), month: topVisits_(visitMonth, 5), year: topVisits_(visitYear, 5) },
    longestVisit: { week: topTimes_(timeWeek, 5), month: topTimes_(timeMonth, 5), year: topTimes_(timeYear, 5) }
  };
}

// ================================================================
// PUBLIC NEWSPAPER/MAGAZINE ATTENDANCE - reads the copied
// Daily_This_Month / Daily_Last_Month sheets. Only 2 months of daily
// history are kept in this system - dates outside that range return
// "not available" rather than guessing.
// ================================================================

var NP_MONTHS_ = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY',
                   'AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
var NP_NAME_ROW = 7;   // header row
var NP_START_ROW = 8;  // first data row
var NP_NAME_COL = 2;   // column B
var NP_FIRST_DAY_COL = 3; // column C = day 1

function npLastRealRow_(sh) {
  var last = sh.getLastRow();
  var numRows = Math.max(last - NP_START_ROW + 1, 0);
  if (!numRows) return NP_START_ROW - 1;
  var col = sh.getRange(NP_START_ROW, NP_NAME_COL, numRows, 1).getValues();
  // Walk DOWN from the first data row and stop at the first blank cell -
  // the real daily newspaper/magazine list is one contiguous block
  // starting at NP_START_ROW, one name per row, with no blank rows inside
  // it. The old version scanned from the BOTTOM of the whole column for
  // the last non-blank cell, which reaches straight past that block's own
  // blank divider and picks up whatever unrelated text happens to sit
  // further down column B on the sheet (e.g. a branch name like "Honavar"
  // or a stray date) - those then got listed as if they were newspaper
  // names in the daily attendance checklist ("Save Today's Attendance").
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() === '') return NP_START_ROW + i - 1;
  }
  return NP_START_ROW + col.length - 1;
}

function getPublicNewspaperStatus(dateStr, nameFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target = dateStr ? new Date(dateStr) : new Date();
  var targetMonthIdx = target.getMonth();
  var targetYear = target.getFullYear();
  var targetDay = target.getDate();

  var today = stripTime_(new Date());
  var curMonthIdx = today.getMonth(), curYear = today.getFullYear();
  var lastMonthDate = new Date(curYear, curMonthIdx - 1, 1);
  var lastMonthIdx = lastMonthDate.getMonth(), lastMonthYear = lastMonthDate.getFullYear();

  var matchedSheet = null, sheetLabel = '';
  if (targetMonthIdx === curMonthIdx && targetYear === curYear) {
    matchedSheet = ss.getSheetByName('Daily_This_Month');
    sheetLabel = 'Daily_This_Month';
  } else if (targetMonthIdx === lastMonthIdx && targetYear === lastMonthYear) {
    matchedSheet = ss.getSheetByName('Daily_Last_Month');
    sheetLabel = 'Daily_Last_Month';
  } else {
    // Older than last month - look for the archived sheet, up to 12 months back.
    var monthsBack = (curYear - targetYear) * 12 + (curMonthIdx - targetMonthIdx);
    if (monthsBack < 0 || monthsBack > 12) {
      return { found: false, message: 'That date is outside the 12-month tracked range.' };
    }
    var archiveName = 'Daily_Archive_' + targetYear + '_' + ('0' + (targetMonthIdx + 1)).slice(-2);
    matchedSheet = ss.getSheetByName(archiveName);
    sheetLabel = archiveName;
    if (!matchedSheet) {
      return { found: false, message: 'No data filed yet for that month (expected sheet "' + archiveName + '" doesn\'t exist yet).' };
    }
  }

  if (!matchedSheet) {
    return { found: false, message: 'No attendance data available for that date.' };
  }

  var dayCol = NP_FIRST_DAY_COL + targetDay - 1;
  var lastRow = npLastRealRow_(matchedSheet);
  var numRows = Math.max(lastRow - NP_START_ROW + 1, 0);
  var items = [];
  if (numRows) {
    var names = matchedSheet.getRange(NP_START_ROW, NP_NAME_COL, numRows, 1).getValues();
    var marks = matchedSheet.getRange(NP_START_ROW, dayCol, numRows, 1).getValues();
    for (var r = 0; r < numRows; r++) {
      var name = String(names[r][0] || '').trim();
      if (!name) continue;
      if (nameFilter && name.toLowerCase().indexOf(String(nameFilter).toLowerCase()) === -1) continue;
      items.push({ name: name, present: String(marks[r][0]).trim() === '\u2713' });
    }
  }

  return {
    found: true,
    date: Utilities.formatDate(target, Session.getScriptTimeZone(), 'dd-MMM-yyyy'),
    sheet: sheetLabel,
    items: items
  };
}

// ---- Staff daily newspaper attendance entry -----------------------------
// Finds the same sheet getPublicNewspaperStatus would (This/Last Month or
// a Daily_Archive_YYYY_MM), for both reading current state and writing.

function npSheetForDate_(target) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetMonthIdx = target.getMonth(), targetYear = target.getFullYear();
  var today = stripTime_(new Date());
  var curMonthIdx = today.getMonth(), curYear = today.getFullYear();
  var lastMonthDate = new Date(curYear, curMonthIdx - 1, 1);
  var lastMonthIdx = lastMonthDate.getMonth(), lastMonthYear = lastMonthDate.getFullYear();

  if (targetMonthIdx === curMonthIdx && targetYear === curYear) return ss.getSheetByName('Daily_This_Month');
  if (targetMonthIdx === lastMonthIdx && targetYear === lastMonthYear) return ss.getSheetByName('Daily_Last_Month');
  var archiveName = 'Daily_Archive_' + targetYear + '_' + ('0' + (targetMonthIdx + 1)).slice(-2);
  return ss.getSheetByName(archiveName);
}

// Daily_This_Month/Daily_Last_Month are reused, fixed-name sheets - the
// ONLY thing that says which real calendar month a sheet's day columns
// (C=day 1 ... AG=day 31) currently represent is its own month-name label
// cell (row 5, col O = 15), which "\ud83d\udd01 Roll Over To Next Month" updates.
// Reading it back here, and comparing it to the date you're trying to
// save, is what catches "it's the 1st of a new month but nobody rolled
// over yet" BEFORE today's checkmarks get written into cells that are
// still really last month's.
function npSheetMonthLabel_(sh) {
  return String(sh.getRange(5, 15).getValue() || '').trim().toUpperCase();
}

function npCheckMonthMatch_(sh, target) {
  var label = npSheetMonthLabel_(sh);
  var expected = NP_MONTHS_[target.getMonth()];
  return { label: label, expected: expected, matches: !label || label === expected };
}

function getNewspaperNamesForDate(dateStr) {
  var target = dateStr ? new Date(dateStr) : new Date();
  var sh = npSheetForDate_(target);
  if (!sh) return { found: false, items: [] };

  var monthCheck = npCheckMonthMatch_(sh, target);
  var dayCol = NP_FIRST_DAY_COL + target.getDate() - 1;
  var lastRow = npLastRealRow_(sh);
  var numRows = Math.max(lastRow - NP_START_ROW + 1, 0);
  var items = [];
  if (numRows) {
    var names = sh.getRange(NP_START_ROW, NP_NAME_COL, numRows, 1).getValues();
    var marks = sh.getRange(NP_START_ROW, dayCol, numRows, 1).getValues();
    for (var i = 0; i < numRows; i++) {
      var name = String(names[i][0] || '').trim();
      if (!name) continue;
      items.push({ name: name, present: String(marks[i][0]).trim() === '\u2713' });
    }
  }
  return {
    found: true,
    items: items,
    sheet: sh.getName(),
    monthLabel: monthCheck.label,
    monthMismatch: !monthCheck.matches,
    monthMismatchMessage: monthCheck.matches ? '' :
      'Warning: ' + sh.getName() + ' is still labeled "' + monthCheck.label + '", not "' +
      monthCheck.expected + '". Saving now would write ' + monthCheck.expected +
      ' attendance into cells that are still really ' + monthCheck.label +
      '\'s. Run "\ud83d\udd01 Roll Over To Next Month" from the Newspaper Billing menu first.'
  };
}

function saveNewspaperAttendance(dateStr, marks) {
  var target = dateStr ? new Date(dateStr) : new Date();
  var sh = npSheetForDate_(target);
  if (!sh) throw new Error('No sheet available for that date - create it first (see Daily_Archive_YYYY_MM naming).');

  var monthCheck = npCheckMonthMatch_(sh, target);
  if (!monthCheck.matches) {
    throw new Error(
      sh.getName() + ' is still labeled "' + monthCheck.label + '", not "' + monthCheck.expected +
      '". Saving would overwrite ' + monthCheck.label + '\'s attendance with ' + monthCheck.expected +
      ' data in the same cells. Run "\ud83d\udd01 Roll Over To Next Month" from the Newspaper Billing menu first, then save again.'
    );
  }

  var dayCol = NP_FIRST_DAY_COL + target.getDate() - 1;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var lastRow = npLastRealRow_(sh);
    var numRows = Math.max(lastRow - NP_START_ROW + 1, 0);
    if (!numRows) throw new Error('No newspapers listed in that sheet.');

    var names = sh.getRange(NP_START_ROW, NP_NAME_COL, numRows, 1).getValues();
    var existing = sh.getRange(NP_START_ROW, dayCol, numRows, 1).getValues();
    var markMap = {};
    (marks || []).forEach(function(m) { markMap[m.name] = m.present; });

    var updates = [];
    for (var i = 0; i < numRows; i++) {
      var name = String(names[i][0] || '').trim();
      if (name && markMap.hasOwnProperty(name)) {
        updates.push([markMap[name] ? '\u2713' : 0]);
      } else {
        updates.push([existing[i][0]]); // leave untouched if not included
      }
    }
    sh.getRange(NP_START_ROW, dayCol, numRows, 1).setValues(updates);

    return {
      saved: true,
      date: Utilities.formatDate(target, Session.getScriptTimeZone(), 'dd-MMM-yyyy'),
      count: (marks || []).length,
      sheet: sh.getName()
    };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// WEEKLY / QUARTERLY / MONTHLY MAGAZINE SEARCH - different structure
// from Daily: repeating sections ("N. MagazineName" header + issue
// rows + total row), not a day-grid. Quatarly_Monthly additionally
// switches to a flat list after the "Monthly Magazines" marker row.
// ================================================================

function npFmtCell_(v) {
  return (v instanceof Date) ? fmtDate_(v) : String(v || '');
}

function parseSectionedMagazines_(sh, startRow, nameFilter, endRow) {
  // NOTE: this used to also cap the scan at "startRow + 400" on top of the
  // sheet's real last row. Weekly/Quatarly_Monthly are single sheets that
  // keep growing every month and are never reset, so once one grew past
  // ~row 406 that cap started silently dropping everything added after it
  // from every search result, admin total, and print sheet - a magazine's
  // later issues (or a whole magazine added further down) would just
  // disappear with no error. sh.getLastRow() (or an explicit endRow, used
  // only to stop before the "Monthly Magazines" marker row) is already
  // the real bound, so the extra cap was pure data loss - removed.
  var lastRow = endRow || sh.getLastRow();
  var numRows = Math.max(lastRow - startRow + 1, 0);
  if (!numRows) return [];
  var data = sh.getRange(startRow, 1, numRows, 8).getValues();

  var sections = [];
  var current = null;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var a = String(row[0] || '').trim();
    var b = row[1], c = row[2], d = row[3], e = row[4];

    var sectionMatch = a.match(/^\d+\.\s*(.+)/);
    if (sectionMatch && !c && !d) {
      current = { name: sectionMatch[1].trim(), issues: [] };
      sections.push(current);
      continue;
    }
    if (!current) continue;
    if (a.toLowerCase().indexOf('add issue') !== -1) continue; // instruction row
    var serial = parseInt(b, 10);
    if (isNaN(serial)) continue; // blank/total row
    // Skip a pre-added blank template row (a serial number with no
    // Volume-Issue/dates filled in yet - per the sheet's own "Insert Row
    // Below -> copy format" instructions). Otherwise it counts as a real
    // issue with no data, inflating counts by 1 and - since it's last in
    // the array - getting shown as "the latest issue" with blank dates.
    if (!c && !d && !e) continue;
    current.issues.push({
      volIssue: c || '', pubDate: npFmtCell_(d), recvDate: npFmtCell_(e), pubDateRaw: d, recvDateRaw: e
    });
  }

  if (nameFilter) {
    sections = sections.filter(function(s) { return s.name.toLowerCase().indexOf(nameFilter.toLowerCase()) !== -1; });
  }
  return sections;
}

function findMonthlyMarkerRow_(sh) {
  // Same fix as parseSectionedMagazines_ above - the quarterly section
  // above this marker row can itself grow past 200 rows over time, which
  // would silently stop this from ever finding the marker (making every
  // monthly magazine below it disappear too). sh.getLastRow() is the real
  // bound.
  var lastRow = sh.getLastRow();
  var data = sh.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    var v = String(data[i][0] || '');
    if (v.indexOf('\u0cae\u0cbe\u0cb8\u0cbf\u0c95 \u0caa\u0ca4\u0ccd\u0cb0\u0cbf\u0c95\u0cc6\u0c97\u0cb3\u0cc1') !== -1 || v.toLowerCase().indexOf('monthly magazines') !== -1) return i + 1;
  }
  return null;
}

function parseMonthlyFlatList_(sh, startRow, nameFilter) {
  // Same fix as parseSectionedMagazines_ above - this had a
  // "startRow + 100" cap on top of the sheet's real last row, which would
  // silently drop any monthly magazine rows added further down as this
  // ever-growing sheet passed ~100 rows past the marker.
  var lastRow = sh.getLastRow();
  var numRows = Math.max(lastRow - startRow + 1, 0);
  if (!numRows) return [];
  var data = sh.getRange(startRow, 1, numRows, 8).getValues();
  var items = [];
  data.forEach(function(row) {
    var rawName = String(row[0] || '').trim(); // column A, e.g. "5.\u0c9a\u0c82\u0caa\u0c95"
    var name = rawName.replace(/^\d+\.\s*/, ''); // strip the leading "5." style prefix
    if (!name) return;
    if (name.indexOf('\u0c92\u0c9f\u0ccd\u0c9f\u0cc1') !== -1) return; // skip total/summary rows ("\u0c92\u0c9f\u0ccd\u0c9f\u0cc1" = Total)
    var volIssue = row[2] || '', pubDate = row[3], recvDate = row[4];
    if (!volIssue && !pubDate && !recvDate) return; // total rows have a label but no real data columns - skip
    if (nameFilter && name.toLowerCase().indexOf(nameFilter.toLowerCase()) === -1) return;
    items.push({
      name: name, volIssue: volIssue, pubDate: npFmtCell_(pubDate), recvDate: npFmtCell_(recvDate), pubDateRaw: pubDate, recvDateRaw: recvDate
    });
  });
  return items;
}

function getPublicMagazineList() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MAGAZINE_LIST');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2, 1, last - 1, 3).getValues();
  var out = [];
  data.forEach(function(r) {
    var name = String(r[0] || '').trim();
    if (!name) return;
    if (String(r[2] || '').trim().toLowerCase() === 'no') return; // inactive
    out.push({ name: name, category: r[1] || '' });
  });
  out.sort(function(a, b) { return a.name.localeCompare(b.name); });
  return out;
}

// Handles both real Sheets date values AND plain text like "23/03/2026" -
// Indian DD/MM/YYYY dates typed into a cell don't always get auto-recognized
// as real dates by Sheets, so filtering must work either way.
//
// IMPORTANT LIMITATION this function cannot fix: if a Weekly/Quatarly_Monthly
// date cell is NOT plain text - i.e. Sheets already auto-converted what was
// typed into a real Date value - Sheets did that conversion using the
// SPREADSHEET's own locale, which for an ambiguous pair like "05/09/2026"
// (day 5 or month 5? both are valid) can silently pick MM/DD instead of the
// DD/MM the library actually uses, turning "5 September" into "9 May" with
// no error shown anywhere. Once that's baked into a real Date value there is
// no way for this function (or any code) to recover which order was
// intended - the ambiguous information is already gone. A day above 12
// (e.g. "24/09/2026") is immune to this because no locale can read "24" as
// a month, which is exactly why some rows in a magazine's issue list can
// look fine while a sibling row from the same batch silently isn't.
// The permanent fix is forcing those two columns to Plain Text format so
// Sheets never attempts this auto-conversion again (menu: "\ud83d\udcf0 Newspaper
// Billing" -> "\ud83d\udd24 Force Weekly/Quatarly_Monthly Dates To Plain Text (run
// once)") - after that, whatever is typed is kept exactly as typed and the
// day-first parsing below is what actually reads it. Any date already
// silently flipped before that needs to be retyped once, same as the
// Letter Date fix earlier.
function parseFlexibleWqmDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (!v) return null;
  var s = String(v).trim();
  if (!s) return null;

  // DD/MM/YYYY or DD-MM-YYYY, day first always (Indian convention) -
  // this only runs on values Sheets left as literal text.
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    var day = parseInt(m[1], 10), month = parseInt(m[2], 10), year = parseInt(m[3], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) return new Date(year, month - 1, day);
  }

  // "15Sep2026" / "15-Sep-2026" / "15 Sep 2026" - day + 3-letter (or full)
  // month name + year, with or without separators. The Quarterly section's
  // Published Date column on Quatarly_Monthly uses exactly this style with
  // no separators at all, which JavaScript's own Date() cannot parse and
  // silently returns "Invalid Date" for - previously making that column
  // useless for month-matching (e.g. \u0ca8\u0cbe\u0c97\u0cb0\u0cbf\u0c95's Published Date never counted
  // for anything, only its Received Date did, which is the one column
  // vulnerable to the MM/DD-vs-DD/MM mixup above).
  var monNames_ = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  var m2 = s.match(/^(\d{1,2})[\s\-]*([A-Za-z]{3})[A-Za-z]*[\s\-]*(\d{4})$/);
  if (m2) {
    var mon = monNames_[m2[2].toLowerCase()];
    if (mon !== undefined) {
      var day2 = parseInt(m2[1], 10), year2 = parseInt(m2[3], 10);
      if (day2 >= 1 && day2 <= 31) return new Date(year2, mon, day2);
    }
  }

  // "September 2026" - a month name + year with no specific day (the flat
  // Monthly list's Publishing Date column uses this) - treat as the 1st.
  var m3 = s.match(/^([A-Za-z]+)\s+(\d{4})$/);
  if (m3) {
    var mon2 = monNames_[m3[1].toLowerCase().slice(0, 3)];
    if (mon2 !== undefined) return new Date(parseInt(m3[2], 10), mon2, 1);
  }

  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// One-time utility: forces the Published Date and Received Date columns on
// both Weekly (D:E) and Quatarly_Monthly (D:E, both the quarterly section
// AND the flat Monthly list below it) to Plain Text number format, for every
// row (current data plus a generous margin for future rows). This is the
// permanent fix for the MM/DD-vs-DD/MM silent date corruption described in
// the big comment on parseFlexibleWqmDate_ above - once a column is Plain
// Text, Sheets never reinterprets what you type using its own locale, so
// whatever you type (always day-first, DD/MM/YYYY) is exactly what gets
// read back. Run this ONCE from the menu; it only changes formatting, never
// touches any existing values, so nothing already typed is erased - but a
// date that was ALREADY silently flipped before this runs is still wrong
// and needs to be retyped once after running this (same as the Letter Date
// fix earlier in this project).
function npbForceWqmDateColumnsToPlainText() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var margin = 500; // room for future rows without re-running this
  ['Weekly', 'Quatarly_Monthly'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var lastRow = Math.max(sh.getLastRow() + margin, margin);
    sh.getRange(1, 4, lastRow, 2).setNumberFormat('@'); // columns D:E
  });
}

function npbForceWqmDateColumnsToPlainTextMenuAction_() {
  var ui = SpreadsheetApp.getUi();
  try {
    npbForceWqmDateColumnsToPlainText();
    ui.alert('Done. Published Date and Received Date on Weekly and Quatarly_Monthly are now Plain Text, so new dates you type will never get silently reinterpreted. Any date already flipped before now needs to be retyped once for the fix to apply to it.');
  } catch (e) {
    ui.alert('Failed: ' + e.message);
  }
}

function issueMatchesMonth_(rawDate, year, month) {
  if (!year) return true; // no month filter requested - show everything
  var d = parseFlexibleWqmDate_(rawDate);
  if (!d) return false;
  if (d.getFullYear() !== Number(year)) return false;
  if (month && (d.getMonth() + 1) !== Number(month)) return false;
  return true;
}

// An issue counts for a month if EITHER its Published Date OR its Received
// Date falls in that month/year - not Received Date alone.
//
// Why: a weekly issue "published" in the first days of a month is very
// often physically received a few days *before* that, still in the tail
// end of the previous calendar month (e.g. the issue dated 06/09/2026 can
// arrive on 28/08/2026). Filtering by Received Date only was silently
// dropping those issues from the current month's bill and count, even
// though the issue itself unambiguously belongs to this month (its
// Volume-Issue and Published Date say so). Matching on either date fixes
// that real-world early-delivery case, and as a side effect also tolerates
// a mistyped year in one of the two date cells (e.g. Received Date typed
// as 09/09/2025 instead of 2026) as long as the other date is correct.
function issueMatchesMonthEither_(pubRawDate, recvRawDate, year, month) {
  if (!year) return true;
  return issueMatchesMonth_(pubRawDate, year, month) || issueMatchesMonth_(recvRawDate, year, month);
}

function filterSectionsByMonth_(sections, year, month) {
  return sections.map(function(s) {
    var filtered = s.issues.filter(function(it) { return issueMatchesMonthEither_(it.pubDateRaw, it.recvDateRaw, year, month); })
      .map(function(it) { return { volIssue: it.volIssue, pubDate: it.pubDate, recvDate: it.recvDate }; });
    return { name: s.name, issues: filtered };
  }).filter(function(s) { return s.issues.length > 0; });
}

// NOTE: Weekly and Quatarly_Monthly are ONE continuous sheet each - each
// magazine's issue list just keeps growing across many months, unlike
// Daily which gets a fresh sheet per month. So month filtering here works
// by checking each issue's real Received Date, not by looking for a
// different sheet per month (there isn't one for these two).
function getWeeklyQuarterlyMonthlyStatus(magazineName, monthStr) {
  var results = { weekly: [], quarterly: [], monthly: [] };
  if (!magazineName) return results;

  var year = '', month = '';
  if (monthStr) {
    var parts = monthStr.split('-');
    year = parts[0]; month = parts[1];
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // startRow=1 for both (was hardcoded 6/7) - the parser already ignores
  // any row that isn't a real "N. Name" section header or a numbered issue
  // row, so scanning from row 1 is always safe, while a hardcoded start
  // that's even one row too LATE silently loses that sheet's first
  // magazine entirely if a row shifted (this is what made \u0ca8\u0cbe\u0c97\u0cb0\u0cbf\u0c95 vanish).
  var weekly = ss.getSheetByName('Weekly');
  if (weekly) {
    var wSections = parseSectionedMagazines_(weekly, 1, magazineName)
      .filter(function(s) { return s.name.toLowerCase() === magazineName.toLowerCase(); });
    results.weekly = filterSectionsByMonth_(wSections, year, month);
  }

  var qm = ss.getSheetByName('Quatarly_Monthly');
  if (qm) {
    var marker = findMonthlyMarkerRow_(qm);
    var qSections = parseSectionedMagazines_(qm, 1, magazineName, marker ? marker - 1 : null)
      .filter(function(s) { return s.name.toLowerCase() === magazineName.toLowerCase(); });
    results.quarterly = filterSectionsByMonth_(qSections, year, month);

    if (marker) {
      var mItems = parseMonthlyFlatList_(qm, marker + 1, magazineName)
        .filter(function(it) { return it.name.toLowerCase() === magazineName.toLowerCase(); });
      results.monthly = mItems.filter(function(it) { return issueMatchesMonthEither_(it.pubDateRaw, it.recvDateRaw, year, month); })
        .map(function(it) { return { name: it.name, volIssue: it.volIssue, pubDate: it.pubDate, recvDate: it.recvDate }; });
    }
  }

  var anyFound = results.weekly.length || results.quarterly.length || results.monthly.length;
  if (!anyFound && !weekly && !qm) {
    results.notAvailable = 'Weekly/Quatarly_Monthly sheets not found.';
  }
  return results;
}

function addMagazineIssue(category, magazineName, volIssue, pubDate, recvDate, price) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = category === 'Weekly' ? 'Weekly' : 'Quatarly_Monthly';
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(sheetName + ' sheet not found.');

  var startRow = category === 'Weekly' ? 6 : 7;
  var endRow = null;
  if (category === 'Quarterly') {
    var marker = findMonthlyMarkerRow_(sh);
    endRow = marker ? marker - 1 : null;
  }

  var lastRow = endRow || sh.getLastRow();
  var maxScan = Math.min(lastRow, startRow + 400);
  var data = sh.getRange(startRow, 1, maxScan - startRow + 1, 8).getValues();

  var sectionHeaderRow = -1, lastIssueRow = -1, totalRow = -1, maxSerial = 0, inSection = false;
  for (var i = 0; i < data.length; i++) {
    var absRow = startRow + i;
    var a = String(data[i][0] || '').trim();
    var b = data[i][1], c = data[i][2], d = data[i][3];
    var sectionMatch = a.match(/^\d+\.\s*(.+)/);

    if (sectionMatch && !c && !d) {
      if (inSection) break; // moved into the next section - stop
      if (sectionMatch[1].trim().toLowerCase() === magazineName.toLowerCase()) {
        inSection = true;
        sectionHeaderRow = absRow;
      }
      continue;
    }
    if (!inSection) continue;
    if (a.toLowerCase().indexOf('add issue') !== -1) break; // instruction row = end of section
    var serial = parseInt(b, 10);
    if (!isNaN(serial)) {
      lastIssueRow = absRow;
      if (serial > maxSerial) maxSerial = serial;
    } else if (!b && !c && !d) {
      totalRow = absRow;
      break;
    }
  }

  if (sectionHeaderRow === -1) throw new Error('Magazine "' + magazineName + '" not found in ' + sheetName + '.');
  if (lastIssueRow === -1) throw new Error('No existing issue row found for "' + magazineName + '" to copy formatting from.');

  var insertAt = totalRow !== -1 ? totalRow : lastIssueRow + 1;
  sh.insertRowBefore(insertAt);
  sh.getRange(lastIssueRow, 1, 1, 8).copyTo(sh.getRange(insertAt, 1, 1, 8), { formatOnly: true });

  sh.getRange(insertAt, 2).setValue(maxSerial + 1);
  sh.getRange(insertAt, 3).setValue(volIssue || '');
  sh.getRange(insertAt, 4).setValue(pubDate ? new Date(pubDate) : '');
  sh.getRange(insertAt, 5).setValue(recvDate ? new Date(recvDate) : '');
  if (price) sh.getRange(insertAt, 6).setValue(Number(price));

  return { saved: true, magazine: magazineName, serial: maxSerial + 1, row: insertAt };
}

function addMonthlyMagazineIssue(magazineName, volIssue, pubDate, recvDate, price) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Quatarly_Monthly');
  var marker = findMonthlyMarkerRow_(sh);
  if (!marker) throw new Error('"Monthly Magazines" marker row not found in Quatarly_Monthly.');

  var lastRow = Math.min(sh.getLastRow(), marker + 150);
  var numRows = Math.max(lastRow - (marker + 1) + 1, 0);
  var lastRealRow = marker, maxNo = 0;
  if (numRows) {
    var data = sh.getRange(marker + 1, 1, numRows, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      var name = String(data[i][2] || '').trim();
      if (name) {
        lastRealRow = marker + 1 + i;
        var no = parseInt(data[i][1], 10);
        if (!isNaN(no) && no > maxNo) maxNo = no;
      }
    }
  }

  var insertAt = lastRealRow + 1;
  if (lastRealRow > marker) {
    sh.insertRowAfter(lastRealRow);
    sh.getRange(lastRealRow, 1, 1, 8).copyTo(sh.getRange(insertAt, 1, 1, 8), { formatOnly: true });
  }

  sh.getRange(insertAt, 2).setValue(maxNo + 1);
  sh.getRange(insertAt, 3).setValue(magazineName);
  sh.getRange(insertAt, 4).setValue(volIssue || '');
  sh.getRange(insertAt, 5).setValue(pubDate || ''); // often a text month label, not a strict date
  sh.getRange(insertAt, 6).setValue(recvDate ? new Date(recvDate) : '');
  if (price) sh.getRange(insertAt, 7).setValue(Number(price));

  return { saved: true, magazine: magazineName, no: maxNo + 1, row: insertAt };
}

// ================================================================
// IMPORT NEW ACCESSION DATA - staff pastes rows copied from Excel
// (tab-separated) in the 17-column accession-register layout:
// SL_NO, ACCESSION NUMBER, BARCODE, KARWAR NUMBER, CALL NUMBER,
// AUTHOR, TITTLE, (blank), (blank), Publisher, Place,
// PUBLICATION YEAR, NUMBER OF PAGES, PRICE, location, Status, Remark
// - and this maps each field into Books Catalogue's actual column
// order automatically. SL_NO and the two blank columns are skipped
// (row position already serves as the sequence in Books Catalogue).
// ================================================================

function importAccessionData(pastedText, hasHeaderRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bc = ss.getSheetByName('Books Catalogue');
  var lines = String(pastedText || '').split('\n').map(function(l) { return l.replace(/\r$/, ''); }).filter(function(l) { return l.trim() !== ''; });
  if (hasHeaderRow) lines.shift();
  if (!lines.length) throw new Error('No data rows found in the pasted text.');

  // Build a set of existing Accession Numbers to skip duplicates.
  var lastRow = bc.getLastRow();
  var numRows = Math.max(lastRow - BOOKS_START_ROW + 1, 0);
  var existingIds = {};
  if (numRows) {
    bc.getRange(BOOKS_START_ROW, BOOKS_COL.id, numRows, 1).getValues().forEach(function(r) {
      var id = norm_(r[0]);
      if (id) existingIds[id] = true;
    });
  }

  var imported = [], skippedDuplicate = [], skippedNoId = [];
  var nextRow = lastRow + 1; // Books Catalogue is packed with no phantom bloat (confirmed earlier)

  var out = [];
  lines.forEach(function(line) {
    var cols = line.split('\t');
    var accNum = String(cols[1] || '').trim();
    if (!accNum) { skippedNoId.push(line); return; }
    if (existingIds[norm_(accNum)]) { skippedDuplicate.push(accNum); return; }

    var row = new Array(BOOKS_LAST_COL).fill('');
    row[BOOKS_COL.id - 1] = accNum;
    row[BOOKS_COL.title - 1] = cols[6] || '';
    row[BOOKS_COL.author - 1] = cols[5] || '';
    // genre, totalCopies: not in the accession format - left blank
    row[BOOKS_COL.shelfLocation - 1] = cols[14] || '';
    row[BOOKS_COL.remarks - 1] = cols[16] || '';
    row[BOOKS_COL.callno - 1] = cols[4] || '';
    row[BOOKS_COL.karwarno - 1] = cols[3] || '';
    row[BOOKS_COL.year - 1] = cols[11] || '';
    row[BOOKS_COL.pages - 1] = cols[12] || '';
    row[BOOKS_COL.price - 1] = cols[13] || '';
    row[BOOKS_COL.publisher - 1] = cols[9] || '';
    row[BOOKS_COL.place - 1] = cols[10] || '';
    row[BOOKS_COL.barcode - 1] = cols[2] || '';
    row[BOOKS_COL.status - 1] = cols[15] || '';

    out.push(row);
    existingIds[norm_(accNum)] = true; // guard against duplicates within the same paste
    imported.push(accNum);
  });

  if (out.length) {
    bc.getRange(nextRow, 1, out.length, BOOKS_LAST_COL).setValues(out);
  }

  return {
    imported: imported.length,
    skippedDuplicate: skippedDuplicate,
    skippedNoId: skippedNoId.length
  };
}

function getPublicGallery() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GALLERY');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2, 1, last - 1, 3).getValues();
  var out = [];
  data.forEach(function(r) {
    var url = String(r[2] || '').trim();
    if (!url) return;
    out.push({ category: r[0] || '', caption: r[1] || '', imageUrl: url });
  });
  return out;
}

function getPublicEvents() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EVENTS');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2, 1, last - 1, 3).getValues();
  var out = [];
  data.forEach(function(r) {
    var title = String(r[1] || '').trim();
    if (!title) return;
    out.push({ date: npFmtCell_(r[0]), title: title, photoUrl: r[2] || '' });
  });
  out.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
  return out;
}

// ================================================================
// PRINTABLE REPORTS - Member Report, Discard Book Report, Accession
// Register. All READ-ONLY, no writes to any sheet. Print handled by the
// browser (same pattern as Cards/Labels) - no server-side PDF needed.
// ================================================================

// mode: 'new' (by Year/Month, as before), 'all' (full register), or
// 'fy' (financial year - Indian convention, April of `year` through March
// of `year`+1, e.g. year=2026 means April 2026 - March 2027). `month` is
// ignored in 'fy' mode - a financial year report always covers all 12
// months of that year.
function getMemberReportData(mode, year, month) {
  var sh = ss_().getSheetByName('Members');
  var last = sh.getLastRow();
  var numRows = Math.max(last - MEMBERS_START_ROW + 1, 0);
  if (!numRows) return [];
  var data = sh.getRange(MEMBERS_START_ROW, 1, numRows, MEMBERS_LAST_COL).getValues();
  var photoMap = buildGateProfilePhotoMap_();

  var fyStart, fyEnd;
  if (mode === 'fy') {
    fyStart = new Date(Number(year), 3, 1); // April 1
    fyEnd = new Date(Number(year) + 1, 2, 31, 23, 59, 59); // March 31 next year, end of day
  }

  var out = [];
  data.forEach(function(r) {
    var id = String(r[MEMBERS_COL.id - 1] || '').trim();
    if (!id) return;
    var membStart = r[MEMBERS_COL.membStart - 1];

    if (mode === 'new') {
      if (!(membStart instanceof Date)) return;
      if (membStart.getFullYear() !== Number(year)) return;
      if (month && (membStart.getMonth() + 1) !== Number(month)) return;
    } else if (mode === 'fy') {
      if (!(membStart instanceof Date)) return;
      if (membStart < fyStart || membStart > fyEnd) return;
    }

    out.push({
      id: id, name: r[MEMBERS_COL.name - 1] || '', phone: r[MEMBERS_COL.phone - 1] || '',
      email: r[MEMBERS_COL.email - 1] || '', membStart: fmtDate_(membStart),
      membEnd: fmtDate_(r[MEMBERS_COL.membEnd - 1]), status: r[MEMBERS_COL.status - 1] || '',
      age: r[MEMBERS_COL.age - 1] || '', gender: r[MEMBERS_COL.gender - 1] || '',
      profession: r[MEMBERS_COL.profession - 1] || '', qualification: r[MEMBERS_COL.qualification - 1] || '',
      address: r[MEMBERS_COL.address - 1] || '', group: r[MEMBERS_COL.group - 1] || '',
      amount: r[MEMBERS_COL.amount - 1] || '', receiptNo: r[MEMBERS_COL.receiptNo - 1] || '',
      receiptDate: fmtDate_(r[MEMBERS_COL.receiptDate - 1]),
      challanDetail: r[MEMBERS_COL.challanDetail - 1] || '',
      // Resolved to a directly-embeddable thumbnail URL (same conversion
      // used for Member Cards). Members!Photo_Link (column S) is now the
      // primary source; MEMBER_GATE_PROFILE.PhotoPath is the fallback for
      // photos saved before that column existed.
      photo: resolvePhotoPath_(r[MEMBERS_COL.photoLink - 1] || photoMap[norm_(id)]) || ''
    });
  });
  return out;
}

// Reads MEMBER_GATE_PROFILE ONCE and returns { id: photoPath } for every
// member that has one set. Used by getMemberReportData() so a report of
// hundreds of members does one sheet read instead of one lookup per row
// (gateLookupProfile_ scans the whole sheet per call - fine for a single
// member on the Edit Member tab, too slow repeated per report row).
function buildGateProfilePhotoMap_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MEMBER_GATE_PROFILE');
  var map = {};
  if (!sh) return map;
  var last = sh.getLastRow();
  if (last < 2) return map;
  var vals = sh.getRange(2, 1, last - 1, 4).getValues(); // ID, VIP, Message, PhotoPath
  vals.forEach(function(r) {
    var id = norm_(r[0]);
    if (id) map[id] = r[3] || '';
  });
  return map;
}

// A photo link that resolves to a valid thumbnail URL can still fail to
// load for everyone but the file's owner if the file itself is set to
// "Restricted" sharing in Drive - a plain "Copy link" from Drive does NOT
// always mean "Anyone with the link can view". Since Edit Member is a
// deliberate staff action to set THIS member's photo (unlike an arbitrary
// applicant-submitted link elsewhere), it's safe and expected to fix the
// file's sharing here automatically, the same way resolveFormUploadedFile_
// already does for registration-form uploads. Silently does nothing if the
// link isn't a Drive link, or if this script account doesn't have
// permission to change sharing on that file (e.g. it lives in someone
// else's personal Drive, not the library's) - the link is still saved
// either way, just may not display until sharing is fixed by hand.
// Returns 'fixed' (sharing set OK), 'not-drive-link' (nothing to do - e.g.
// a local file path), or 'failed' (looked like a Drive link but this script
// account couldn't change its sharing - most likely the file lives in
// someone else's personal Drive, not one the library account can manage).
function tryFixDrivePhotoSharing_(photoPath) {
  var v = String(photoPath || '').trim();
  if (!v) return 'not-drive-link';
  var m = v.match(/\/d\/([-\w]{15,})/) || v.match(/[?&]id=([-\w]{15,})/) || v.match(/^([-\w]{15,})$/);
  if (!m) return 'not-drive-link';
  try {
    DriveApp.getFileById(m[1]).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'fixed';
  } catch (e) {
    return 'failed';
  }
}

// Finds (or creates) a member's row in MEMBER_GATE_PROFILE and sets just
// their PhotoPath (column D) - the same sheet/column the gate kiosk
// already reads from (gateLookupProfile_), so Edit Member's Photo Path
// field is the SAME photo the gate check-in screen shows, not a second,
// separately-drifting copy.
function setMemberGatePhotoPath_(id, photoPath) {
  var sharingResult = tryFixDrivePhotoSharing_(photoPath);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MEMBER_GATE_PROFILE');
  if (!sh) return sharingResult; // sheet not set up - nothing to persist to, safe no-op
  var last = sh.getLastRow();
  if (last >= 2) {
    var idCol = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < idCol.length; i++) {
      if (norm_(idCol[i][0]) === id) {
        sh.getRange(2 + i, 4).setValue(photoPath || '');
        return sharingResult;
      }
    }
  }
  // No existing row for this member yet - create a minimal one.
  var numCols = Math.max(sh.getLastColumn(), 4);
  var newRow = new Array(numCols).fill('');
  newRow[0] = id;
  newRow[3] = photoPath || '';
  sh.getRange(last + 1, 1, 1, numCols).setValues([newRow]);
  return sharingResult;
}

// Full member record for the Edit Member tab - every editable Members
// field plus Photo Path (sourced from MEMBER_GATE_PROFILE, see above).
// Dates come back as yyyy-MM-dd so they drop straight into <input
// type="date"> boxes already filled in.
function getMemberFullDetails(memberId) {
  var id = norm_(memberId);
  if (!id) throw new Error('Enter a Member ID.');
  var sh = ss_().getSheetByName('Members');
  var last = membersLastRealRow_(sh);
  var numRows = Math.max(last - MEMBERS_START_ROW + 1, 0);
  if (!numRows) throw new Error('Member ID ' + id + ' not found.');
  var data = sh.getRange(MEMBERS_START_ROW, 1, numRows, MEMBERS_LAST_COL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (norm_(data[i][MEMBERS_COL.id - 1]) === id) {
      var r = data[i];
      // Photo: Members!Photo_Link (column S) is now the primary source;
      // if that's blank, fall back to MEMBER_GATE_PROFILE (older photos
      // saved before this column existed).
      var photoLink = r[MEMBERS_COL.photoLink - 1] || '';
      if (!photoLink) {
        var profile = gateLookupProfile_(id);
        photoLink = profile.photoPath || '';
      }
      return {
        id: id, name: r[MEMBERS_COL.name - 1] || '', phone: r[MEMBERS_COL.phone - 1] || '',
        email: r[MEMBERS_COL.email - 1] || '',
        membStart: isoDate_(r[MEMBERS_COL.membStart - 1]), membEnd: isoDate_(r[MEMBERS_COL.membEnd - 1]),
        status: r[MEMBERS_COL.status - 1] || '', notes: r[MEMBERS_COL.notes - 1] || '',
        age: r[MEMBERS_COL.age - 1] || '', gender: r[MEMBERS_COL.gender - 1] || '',
        profession: r[MEMBERS_COL.profession - 1] || '', qualification: r[MEMBERS_COL.qualification - 1] || '',
        address: r[MEMBERS_COL.address - 1] || '', group: r[MEMBERS_COL.group - 1] || '',
        amount: r[MEMBERS_COL.amount - 1] || '', receiptNo: r[MEMBERS_COL.receiptNo - 1] || '',
        receiptDate: isoDate_(r[MEMBERS_COL.receiptDate - 1]), challanDetail: r[MEMBERS_COL.challanDetail - 1] || '',
        photoPath: photoLink
      };
    }
  }
  throw new Error('Member ID ' + id + ' not found.');
}

// Saves every editable Members field for one member (Member ID itself is
// never changed here - that's their permanent identity, used everywhere
// else in the app). membStart is only overwritten when a value is actually
// given - clearing it by accident would silently break the "New
// Registrations by Year" report, which depends on it. Photo Path is saved
// separately to MEMBER_GATE_PROFILE (see setMemberGatePhotoPath_ above).
function updateMember(memberId, data) {
  var id = norm_(memberId);
  if (!id) throw new Error('Enter a Member ID.');
  if (!data || !String(data.name || '').trim()) throw new Error('Name cannot be blank.');
  var sh = ss_().getSheetByName('Members');
  var last = membersLastRealRow_(sh);
  var numRows = Math.max(last - MEMBERS_START_ROW + 1, 0);
  if (!numRows) throw new Error('Member ID ' + id + ' not found.');
  var idCol = sh.getRange(MEMBERS_START_ROW, MEMBERS_COL.id, numRows, 1).getValues();
  var rowOffset = -1;
  for (var i = 0; i < idCol.length; i++) {
    if (norm_(idCol[i][0]) === id) { rowOffset = i; break; }
  }
  if (rowOffset === -1) throw new Error('Member ID ' + id + ' not found.');
  var row = MEMBERS_START_ROW + rowOffset;

  sh.getRange(row, MEMBERS_COL.name).setValue(data.name || '');
  sh.getRange(row, MEMBERS_COL.phone).setValue(data.phone || '');
  sh.getRange(row, MEMBERS_COL.email).setValue(data.email || '');
  if (data.membStart) sh.getRange(row, MEMBERS_COL.membStart).setValue(new Date(data.membStart));
  sh.getRange(row, MEMBERS_COL.membEnd).setValue(data.membEnd ? new Date(data.membEnd) : '');
  sh.getRange(row, MEMBERS_COL.status).setValue(data.status || '');
  sh.getRange(row, MEMBERS_COL.notes).setValue(data.notes || '');
  sh.getRange(row, MEMBERS_COL.age).setValue(data.age || '');
  sh.getRange(row, MEMBERS_COL.gender).setValue(data.gender || '');
  sh.getRange(row, MEMBERS_COL.profession).setValue(data.profession || '');
  sh.getRange(row, MEMBERS_COL.qualification).setValue(data.qualification || '');
  sh.getRange(row, MEMBERS_COL.address).setValue(data.address || '');
  sh.getRange(row, MEMBERS_COL.group).setValue(data.group || '');
  sh.getRange(row, MEMBERS_COL.amount).setValue(data.amount === '' || data.amount == null ? '' : (Number(data.amount) || data.amount));
  sh.getRange(row, MEMBERS_COL.receiptNo).setValue(data.receiptNo || '');
  sh.getRange(row, MEMBERS_COL.receiptDate).setValue(data.receiptDate ? new Date(data.receiptDate) : '');
  sh.getRange(row, MEMBERS_COL.challanDetail).setValue(data.challanDetail || '');

  // Photo: saved directly into Members!Photo_Link (column S) - the primary,
  // directly-visible copy - AND kept in sync in MEMBER_GATE_PROFILE, which
  // is what the Gate Entry kiosk check-in screen reads from.
  sh.getRange(row, MEMBERS_COL.photoLink).setValue(data.photoPath || '');
  var photoSharing = setMemberGatePhotoPath_(id, data.photoPath || '');

  return { id: id, saved: true, photoSharing: photoSharing };
}

// filters (all optional - pass {} or omit for the full register):
//   fromAcc, toAcc  - Book ID / Accession No. range (serial-number-wise)
//   year            - Decard Date's year, e.g. 2026 (year-wise)
//   reason          - exact match against getDiscardReasons() (purpose-wise)
//   staffName       - who did it; matched as "contains", case-insensitive
function getDiscardReportData(filters) {
  filters = filters || {};
  var fromNum = filters.fromAcc ? parseInt(filters.fromAcc, 10) : null;
  var toNum = filters.toAcc ? parseInt(filters.toAcc, 10) : null;
  var yearNum = filters.year ? parseInt(filters.year, 10) : null;
  var reasonWant = filters.reason ? String(filters.reason).trim() : '';
  var staffWant = filters.staffName ? String(filters.staffName).trim().toLowerCase() : '';

  var sh = ss_().getSheetByName('Decard Books');
  var last = decardLastRealRow_(sh);
  var numRows = Math.max(last - DECARD_START_ROW + 1, 0);
  if (!numRows) return [];
  var data = sh.getRange(DECARD_START_ROW, 1, numRows, DECARD_LAST_COL).getValues();
  var out = [];
  data.forEach(function(r) {
    var id = String(r[DECARD_COL.id - 1] || '').trim();
    if (!id) return;

    var idNum = parseInt(id, 10);
    if (fromNum !== null && !isNaN(idNum) && idNum < fromNum) return;
    if (toNum !== null && !isNaN(idNum) && idNum > toNum) return;

    var decardDateVal = r[DECARD_COL.decardDate - 1];
    if (yearNum !== null) {
      if (!(decardDateVal instanceof Date)) return;
      if (decardDateVal.getFullYear() !== yearNum) return;
    }

    var reasonVal = String(r[DECARD_COL.reason - 1] || '');
    if (reasonWant && reasonVal !== reasonWant) return;

    var staffVal = String(r[DECARD_COL.staffName - 1] || '');
    if (staffWant && staffVal.toLowerCase().indexOf(staffWant) === -1) return;

    out.push({
      id: id, title: r[DECARD_COL.title - 1] || '', author: r[DECARD_COL.author - 1] || '',
      publisher: r[DECARD_COL.publisher - 1] || '', place: r[DECARD_COL.place - 1] || '',
      year: r[DECARD_COL.year - 1] || '', callno: r[DECARD_COL.callno - 1] || '',
      karwarno: r[DECARD_COL.karwarno - 1] || '', pages: r[DECARD_COL.pages - 1] || '',
      price: r[DECARD_COL.price - 1] || '',
      genre: r[DECARD_COL.genre - 1] || '', location: r[DECARD_COL.location - 1] || '',
      barcode: r[DECARD_COL.barcode - 1] || '',
      date: fmtDate_(decardDateVal), staffName: staffVal, reason: reasonVal,
      remarks: r[DECARD_COL.remarks - 1] || ''
    });
  });
  return out;
}

function getAccessionRegisterData(fromAcc, toAcc) {
  var sh = ss_().getSheetByName('Books Catalogue');
  var last = sh.getLastRow();
  var numRows = Math.max(last - BOOKS_START_ROW + 1, 0);
  if (!numRows) return [];
  var data = sh.getRange(BOOKS_START_ROW, 1, numRows, BOOKS_LAST_COL).getValues();

  var fromNum = fromAcc ? parseInt(fromAcc, 10) : null;
  var toNum = toAcc ? parseInt(toAcc, 10) : null;

  var out = [];
  data.forEach(function(r) {
    var id = String(r[BOOKS_COL.id - 1] || '').trim();
    if (!id) return;
    var idNum = parseInt(id, 10);
    if (fromNum !== null && !isNaN(idNum) && idNum < fromNum) return;
    if (toNum !== null && !isNaN(idNum) && idNum > toNum) return;

    out.push({
      id: id, title: r[BOOKS_COL.title - 1] || '', author: r[BOOKS_COL.author - 1] || '',
      genre: r[BOOKS_COL.genre - 1] || '', totalCopies: r[BOOKS_COL.totalCopies - 1] || '',
      location: r[BOOKS_COL.shelfLocation - 1] || '', remarks: r[BOOKS_COL.remarks - 1] || '',
      callno: r[BOOKS_COL.callno - 1] || '', karwarno: r[BOOKS_COL.karwarno - 1] || '',
      year: r[BOOKS_COL.year - 1] || '', pages: r[BOOKS_COL.pages - 1] || '',
      price: r[BOOKS_COL.price - 1] || '', publisher: r[BOOKS_COL.publisher - 1] || '',
      place: r[BOOKS_COL.place - 1] || '', barcode: r[BOOKS_COL.barcode - 1] || '',
      status: r[BOOKS_COL.status - 1] || ''
    });
  });
  return out;
}

function getAccessionRegisterCount() {
  var sh = ss_().getSheetByName('Books Catalogue');
  var last = sh.getLastRow();
  return Math.max(last - BOOKS_START_ROW + 1, 0);
}

function markSingleNewspaper(dateStr, paperName, present) {
  var target = dateStr ? new Date(dateStr) : new Date();
  var sh = npSheetForDate_(target);
  if (!sh) throw new Error('No sheet available for that date.');
  var dayCol = NP_FIRST_DAY_COL + target.getDate() - 1;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var lastRow = npLastRealRow_(sh);
    var numRows = Math.max(lastRow - NP_START_ROW + 1, 0);
    if (!numRows) throw new Error('No newspapers listed in that sheet.');
    var names = sh.getRange(NP_START_ROW, NP_NAME_COL, numRows, 1).getValues();
    for (var i = 0; i < numRows; i++) {
      if (String(names[i][0] || '').trim() === paperName) {
        sh.getRange(NP_START_ROW + i, dayCol).setValue(present ? '\u2713' : 0);
        return { saved: true, paper: paperName, present: present };
      }
    }
    throw new Error('Paper "' + paperName + '" not found in the list.');
  } finally {
    lock.releaseLock();
  }
}

function searchPeopleByName(query) {
  query = String(query || '').trim().toLowerCase();
  if (query.length < 2) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  var members = ss.getSheetByName('Members');
  var mLast = members.getLastRow();
  if (mLast >= MEMBERS_START_ROW) {
    members.getRange(MEMBERS_START_ROW, 1, mLast - MEMBERS_START_ROW + 1, 2).getValues().forEach(function(r) {
      var id = String(r[0] || '').trim();
      var name = String(r[1] || '').trim();
      if (id && name && name.toLowerCase().indexOf(query) !== -1) {
        results.push({ id: id, name: name, type: 'Member' });
      }
    });
  }

  var nonmem = ss.getSheetByName('NON_MEMBERS');
  var nLast = nonmem.getLastRow();
  if (nLast >= NONMEM_START_ROW) {
    nonmem.getRange(NONMEM_START_ROW, 1, nLast - NONMEM_START_ROW + 1, 2).getValues().forEach(function(r) {
      var id = String(r[0] || '').trim();
      var name = String(r[1] || '').trim();
      if (id && name && name.toLowerCase().indexOf(query) !== -1) {
        results.push({ id: id, name: name, type: 'Non-Member' });
      }
    });
  }

  return results.slice(0, 20);
}

// ================================================================
// FACE RECOGNITION ATTENDANCE (PILOT) - browser-based, no local
// software needed (learning from Face Gate's earlier deployment
// friction). Stores a numeric "face descriptor" per enrolled person,
// not a raw photo comparison - matching happens in the browser using
// face-api.js. Start with a small group, test real accuracy at your
// actual kiosk location/lighting, tune the match threshold before
// wider rollout.
// ================================================================

function setupFaceEnrollment() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName('FACE_ENROLLMENT');
  if (existing) {
    SpreadsheetApp.getUi().alert('FACE_ENROLLMENT already exists - not touching it.');
    return;
  }
  var sh = ss.insertSheet('FACE_ENROLLMENT');
  sh.getRange(1, 1, 1, 5).setValues([['ID', 'Name', 'Type', 'Descriptor', 'Enrolled Date']]).setFontWeight('bold');
  SpreadsheetApp.getUi().alert('\u2705 Created FACE_ENROLLMENT sheet. Ready for pilot enrollment.');
}

function saveFaceDescriptor(id, name, type, descriptorArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('FACE_ENROLLMENT');
  if (!sh) throw new Error('FACE_ENROLLMENT sheet not found - run setupFaceEnrollment first.');
  id = norm_(id);
  if (!id) throw new Error('Select a person first.');
  if (!descriptorArray || !descriptorArray.length) throw new Error('No face data captured.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var last = sh.getLastRow();
    var numRows = Math.max(last - 1, 0);
    var existingIds = numRows ? sh.getRange(2, 1, numRows, 1).getValues() : [];
    var foundRow = -1;
    for (var i = 0; i < existingIds.length; i++) {
      if (norm_(existingIds[i][0]) === id) { foundRow = i + 2; break; }
    }
    var row = [id, name || '', type || '', JSON.stringify(descriptorArray), new Date()];
    if (foundRow > -1) {
      sh.getRange(foundRow, 1, 1, 5).setValues([row]); // re-enrollment overwrites the old face data
    } else {
      sh.getRange(last + 1, 1, 1, 5).setValues([row]);
    }
    return { saved: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function getAllFaceDescriptors() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FACE_ENROLLMENT');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2, 1, last - 1, 5).getValues();
  var out = [];
  data.forEach(function(r) {
    var id = String(r[0] || '').trim();
    if (!id) return;
    try {
      var desc = JSON.parse(r[3]);
      out.push({ id: id, name: r[1] || '', type: r[2] || '', descriptor: desc });
    } catch (e) { /* skip a malformed row rather than fail the whole list */ }
  });
  return out;
}

function deleteFaceEnrollment(id) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FACE_ENROLLMENT');
  if (!sh) return { deleted: false };
  id = norm_(id);
  var last = sh.getLastRow();
  if (last < 2) return { deleted: false };
  var data = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (norm_(data[i][0]) === id) { sh.deleteRow(i + 2); return { deleted: true }; }
  }
  return { deleted: false };
}

// ================================================================
// MEMBERSHIP APPLICATION PRINTABLE COPY - recreates the official
// government form, pre-filled from a Pending Registrations row, so the
// applicant gets a typed copy instead of handwriting the whole thing.
// The actual gazetted-officer signature step stays fully physical -
// this only removes the handwriting burden.
// ================================================================

// ================================================================
// APPROVAL QUEUE - three tracked stages before someone becomes a real
// member: Documents Received, Fee Paid, then Approve (which is the only
// step that actually creates the membership). Keeps people who finish
// the paperwork but forget to pay visible, not lost.
// ================================================================

var PENDREG_TRACK_COLS = ['Documents Received', 'Fee Paid', 'Status', 'Approved Member ID'];

function ensurePendingRegTrackingCols_(sh) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var missing = PENDREG_TRACK_COLS.filter(function(c) { return headers.indexOf(c) === -1; });
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
}

function pendRegColIndex_(headers, label) {
  for (var i = 0; i < headers.length; i++) { if (headers[i].indexOf(label) === 0) return i + 1; } // 1-indexed
  return -1;
}

function getPendingRegistrationsList() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pending Registrations');
  if (!sh) return [];
  var headers = ensurePendingRegTrackingCols_(sh);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var data = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var nameCol = pendRegColIndex_(headers, 'Full Name') - 1;
  var phoneCol = pendRegColIndex_(headers, 'Phone Number') - 1;
  var docsCol = pendRegColIndex_(headers, 'Documents Received') - 1;
  var feeCol = pendRegColIndex_(headers, 'Fee Paid') - 1;
  var statusCol = pendRegColIndex_(headers, 'Status') - 1;
  var approvedIdCol = pendRegColIndex_(headers, 'Approved Member ID') - 1;

  var out = [];
  data.forEach(function(row, i) {
    var name = String(row[nameCol] || '').trim();
    if (!name) return;
    var status = String(row[statusCol] || '').trim();
    out.push({
      row: i + 2,
      name: name,
      phone: row[phoneCol] || '',
      docsReceived: String(row[docsCol] || '').toUpperCase() === 'YES',
      feePaid: String(row[feeCol] || '').toUpperCase() === 'YES',
      status: status || 'Awaiting Documents',
      approvedId: row[approvedIdCol] || ''
    });
  });
  out.reverse(); // newest first
  return out;
}

function markPendingRegField_(rowNum, label, value) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pending Registrations');
  var headers = ensurePendingRegTrackingCols_(sh);
  var col = pendRegColIndex_(headers, label);
  sh.getRange(rowNum, col).setValue(value ? 'YES' : 'NO');
  updatePendingRegStatus_(sh, headers, rowNum);
}

function markDocumentsReceived(rowNum, received) { markPendingRegField_(rowNum, 'Documents Received', received); return { success: true }; }
function markFeePaid(rowNum, paid) { markPendingRegField_(rowNum, 'Fee Paid', paid); return { success: true }; }

function updatePendingRegStatus_(sh, headers, rowNum) {
  var docsCol = pendRegColIndex_(headers, 'Documents Received');
  var feeCol = pendRegColIndex_(headers, 'Fee Paid');
  var statusCol = pendRegColIndex_(headers, 'Status');
  var docs = String(sh.getRange(rowNum, docsCol).getValue() || '').toUpperCase() === 'YES';
  var fee = String(sh.getRange(rowNum, feeCol).getValue() || '').toUpperCase() === 'YES';
  var status = !docs ? 'Awaiting Documents' : (!fee ? 'Documents Received - Fee Pending' : 'Ready to Approve');
  sh.getRange(rowNum, statusCol).setValue(status);
}

// Moves an applicant's uploaded file into one shared folder and renames it
// clearly by Member ID, instead of leaving hundreds of loose files scattered
// through Drive. Safe to call even if the file can't be found/moved - a
// failure here should never block someone's actual approval.
function organizeApplicantFile_(rawValue, folder, newName) {
  var v = String(rawValue || '').trim();
  if (!v) return;
  var m = v.match(/\/d\/([-\w]{15,})/) || v.match(/[?&]id=([-\w]{15,})/) || v.match(/^([-\w]{15,})$/);
  var fileId = m ? m[1] : null;
  if (!fileId) return;
  try {
    var file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    file.setName(newName);
    folder.addFile(file);
    var parents = file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      if (parent.getId() !== folder.getId()) parent.removeFile(file);
    }
  } catch (e) { /* skip this one file rather than block the whole approval */ }
}

function getRegistrationDocsFolder_() {
  var folderName = 'BCL Honavar - Registration Documents';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// Saves a PDF the browser already generated (via its own reliable print-to-PDF,
// which renders Kannada correctly) into the shared records folder, alongside
// that applicant's other documents. Kept as a permanent record for staff to
// print again anytime, without needing the applicant to return.
function saveApplicationPdfRecord(rowNum, base64Data, applicantName) {
  var folder = getRegistrationDocsFolder_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pending Registrations');
  var headers = ensurePendingRegTrackingCols_(sh);
  var approvedIdCol = pendRegColIndex_(headers, 'Approved Member ID');
  var approvedId = approvedIdCol > 0 ? sh.getRange(rowNum, approvedIdCol).getValue() : '';

  var safeName = String(applicantName || 'Applicant').replace(/[^a-zA-Z0-9]+/g, '_');
  var fileName = approvedId
    ? (approvedId + '_Application.pdf')
    : ('PENDING_' + safeName + '_row' + rowNum + '_Application.pdf');

  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, 'application/pdf', fileName);
  var file = folder.createFile(blob);
  return { success: true, fileName: fileName, url: file.getUrl() };
}

function approveRegistration(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Pending Registrations');
  var headers = ensurePendingRegTrackingCols_(sh);
  var docsCol = pendRegColIndex_(headers, 'Documents Received');
  var feeCol = pendRegColIndex_(headers, 'Fee Paid');
  var statusCol = pendRegColIndex_(headers, 'Status');
  var approvedIdCol = pendRegColIndex_(headers, 'Approved Member ID');

  var docs = String(sh.getRange(rowNum, docsCol).getValue() || '').toUpperCase() === 'YES';
  var fee = String(sh.getRange(rowNum, feeCol).getValue() || '').toUpperCase() === 'YES';
  if (!docs || !fee) throw new Error('Both Documents Received and Fee Paid must be ticked before approving.');

  var existingApprovedId = sh.getRange(rowNum, approvedIdCol).getValue();
  if (existingApprovedId) throw new Error('Already approved as Member ID ' + existingApprovedId + '.');

  var reg = getPendingRegistrationForPrint(rowNum);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // -- Assign new Member ID and write into the real Members sheet --
    var membersSh = ss.getSheetByName('Members');
    var mLast = membersSh.getLastRow();
    var maxId = 0;
    if (mLast >= MEMBERS_START_ROW) {
      membersSh.getRange(MEMBERS_START_ROW, MEMBERS_COL.id, mLast - MEMBERS_START_ROW + 1, 1).getValues().forEach(function(r) {
        var n = parseInt(r[0], 10);
        if (!isNaN(n) && n > maxId) maxId = n;
      });
    }
    var newId = maxId + 1;
    var addressLine = [reg.houseNo, reg.street, reg.village, reg.district, reg.state, reg.pincode].filter(Boolean).join(', ');

    var newRow = new Array(MEMBERS_LAST_COL).fill('');
    newRow[MEMBERS_COL.id - 1] = newId;
    newRow[MEMBERS_COL.name - 1] = reg.name;
    newRow[MEMBERS_COL.phone - 1] = reg.phone;
    newRow[MEMBERS_COL.email - 1] = reg.email;
    newRow[MEMBERS_COL.membStart - 1] = new Date();
    newRow[MEMBERS_COL.status - 1] = 'Active';
    newRow[MEMBERS_COL.age - 1] = reg.age;
    newRow[MEMBERS_COL.gender - 1] = reg.gender;
    newRow[MEMBERS_COL.profession - 1] = reg.occupation;
    newRow[MEMBERS_COL.qualification - 1] = reg.education;
    newRow[MEMBERS_COL.address - 1] = addressLine;
    var writeRow = Math.max(mLast + 1, MEMBERS_START_ROW);
    membersSh.getRange(writeRow, 1, 1, MEMBERS_LAST_COL).setValues([newRow]);

    // -- Photo/signature go into MEMBER_GATE_PROFILE, same as everywhere else --
    var gateProfileSh = ss.getSheetByName('MEMBER_GATE_PROFILE');
    if (gateProfileSh) {
      var gpLast = gateProfileSh.getLastRow();
      var gpHeaders = gateProfileSh.getRange(1, 1, 1, gateProfileSh.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
      if (gpHeaders.indexOf('SignaturePath') === -1) {
        gateProfileSh.getRange(1, gpHeaders.length + 1).setValue('SignaturePath').setFontWeight('bold');
      }
      var sh_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pending Registrations');
      var rawRow = sh_data.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      var sigCol = pendRegColIndex_(headers, 'Signature');
      var signatureUrl = sigCol > 0 ? resolveFormUploadedFile_(rawRow[sigCol - 1]) : '';

      gateProfileSh.appendRow(['' + newId, 'No', '', reg.photoUrl || '', addressLine, reg.gender, signatureUrl || '']);

      // Move all 4 uploaded files into one organized folder, renamed by this Member ID.
      var docsFolder = getRegistrationDocsFolder_();
      var photoCol = pendRegColIndex_(headers, 'Photo');
      var proof1Col = pendRegColIndex_(headers, 'Address Proof #1 - Upload');
      var proof2Col = pendRegColIndex_(headers, 'Address Proof #2 - Upload');
      if (photoCol > 0) organizeApplicantFile_(rawRow[photoCol - 1], docsFolder, newId + '_Photo');
      if (sigCol > 0) organizeApplicantFile_(rawRow[sigCol - 1], docsFolder, newId + '_Signature');
      if (proof1Col > 0) organizeApplicantFile_(rawRow[proof1Col - 1], docsFolder, newId + '_AddressProof1');
      if (proof2Col > 0) organizeApplicantFile_(rawRow[proof2Col - 1], docsFolder, newId + '_AddressProof2');
    }

    sh.getRange(rowNum, approvedIdCol).setValue(newId);
    sh.getRange(rowNum, statusCol).setValue('GRANTED - Membership Active');

    if (reg.email) {
      try {
        MailApp.sendEmail({
          to: reg.email,
          subject: 'Membership Approved - Branch Central Library, Honavar',
          body: 'Dear ' + reg.name + ',\n\n' +
            'Congratulations! Your library membership has been approved.\n\n' +
            'Your Member ID is: ' + newId + '\n\n' +
            'You can now borrow books at Branch Central Library, Honavar. Please visit us to collect your membership card.\n\n' +
            mmFooter_()
        });
        logEmail_('Other', 'Membership Approved', newId, reg.name, reg.email, '', 'Sent');
      } catch (mailErr) {
        logEmail_('Other', 'Membership Approved', newId, reg.name, reg.email, '', 'Failed');
      }
    }

    return { success: true, memberId: newId, name: reg.name };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// VISITOR (NON-MEMBER) ONLINE REGISTRATION - same pattern as the
// Membership queue above, but simpler on purpose: no Fee Paid step and
// no gazetted-officer print step, since visitors don't need either.
// One click on Approve creates the Non-Member record immediately.
//
// Expects a Google Form whose responses feed a sheet named EXACTLY
// "Pending Visitor Registrations", with questions titled starting with:
// "Full Name", "Phone Number", "Address", "Purpose", "Occupation", "Photo"
// (prefix-matched, so exact wording after that doesn't matter).
// ================================================================

var PENDVISREG_TRACK_COLS = ['Status', 'Approved Visitor ID'];

function ensurePendingVisitorRegTrackingCols_(sh) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var missing = PENDVISREG_TRACK_COLS.filter(function(c) { return headers.indexOf(c) === -1; });
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
}

function pendVisRegColIndex_(headers, label) {
  for (var i = 0; i < headers.length; i++) { if (headers[i].indexOf(label) === 0) return i + 1; } // 1-indexed
  return -1;
}

// Builds one readable address line out of the form's separate
// House Number / Street / Town-Village / District / State / Pincode
// questions (this form reuses the Member form's address layout instead
// of a single "Address" question).
function buildVisitorAddress_(headers, rawRow) {
  var parts = ['House Number', 'Street', 'Town / Village', 'District', 'State'];
  var bits = parts.map(function(label) {
    var col = pendVisRegColIndex_(headers, label);
    return col > 0 ? String(rawRow[col - 1] || '').trim() : '';
  }).filter(function(v) { return v; });
  var line = bits.join(', ');
  var pinCol = pendVisRegColIndex_(headers, 'Pincode');
  var pin = pinCol > 0 ? String(rawRow[pinCol - 1] || '').trim() : '';
  if (pin) line += (line ? ' - ' : '') + pin;
  return line;
}

function getPendingVisitorRegistrationsList() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pending Visitor Registrations');
  if (!sh) return [];
  var headers = ensurePendingVisitorRegTrackingCols_(sh);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var data = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var nameCol = pendVisRegColIndex_(headers, 'Full Name') - 1;
  var phoneCol = pendVisRegColIndex_(headers, 'Phone Number') - 1;
  var occupationCol = pendVisRegColIndex_(headers, 'Occupation') - 1;
  var approvedIdCol = pendVisRegColIndex_(headers, 'Approved Visitor ID') - 1;

  var out = [];
  data.forEach(function(row, i) {
    var name = nameCol > -1 ? String(row[nameCol] || '').trim() : '';
    if (!name) return;
    out.push({
      row: i + 2,
      name: name,
      phone: phoneCol > -1 ? (row[phoneCol] || '') : '',
      purpose: (occupationCol > -1 ? (row[occupationCol] || '') : '') + (buildVisitorAddress_(headers, row) ? ' \u2014 ' + buildVisitorAddress_(headers, row) : ''),
      approvedId: approvedIdCol > -1 ? (row[approvedIdCol] || '') : ''
    });
  });
  out.reverse(); // newest first
  return out;
}

function approveVisitorRegistration(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Pending Visitor Registrations');
  if (!sh) throw new Error('"Pending Visitor Registrations" sheet not found - link your Visitor Registration Google Form\'s responses to a sheet with exactly this name.');
  var headers = ensurePendingVisitorRegTrackingCols_(sh);
  var statusCol = pendVisRegColIndex_(headers, 'Status');
  var approvedIdCol = pendVisRegColIndex_(headers, 'Approved Visitor ID');

  var existingApprovedId = sh.getRange(rowNum, approvedIdCol).getValue();
  if (existingApprovedId) throw new Error('Already approved as Visitor ID ' + existingApprovedId + '.');

  var nameCol = pendVisRegColIndex_(headers, 'Full Name');
  var phoneCol = pendVisRegColIndex_(headers, 'Phone Number');
  var occupationCol = pendVisRegColIndex_(headers, 'Occupation');
  var photoCol = pendVisRegColIndex_(headers, 'Photo');
  // Optional - only present once an "Email ID" question is added to the
  // Visitor Registration form. Until then this is -1 and every visitor's
  // email stays blank, same as before this field existed.
  var emailCol = pendVisRegColIndex_(headers, 'Email ID');

  var rawRow = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var name = nameCol > 0 ? String(rawRow[nameCol - 1] || '').trim() : '';
  if (!name) throw new Error('This row has no Full Name - cannot approve.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var nsh = ss.getSheetByName('NON_MEMBERS');
    if (!nsh) throw new Error('NON_MEMBERS sheet not found.');
    var nLast = nsh.getLastRow();
    var maxId = 5000; // non-member IDs start at 5001, members stay below 5000
    if (nLast >= NONMEM_START_ROW) {
      nsh.getRange(NONMEM_START_ROW, NONMEM_COL.id, nLast - NONMEM_START_ROW + 1, 1).getValues().forEach(function(r) {
        var n = parseInt(r[0], 10);
        if (!isNaN(n) && n > maxId) maxId = n;
      });
    }
    var newId = maxId + 1;

    // This form's "Photo" question is a plain text link (staff paste one in),
    // not a File-upload question - Google Forms forces sign-in for uploads,
    // which would break the no-login requirement for visitors.
    var photoRaw = photoCol > 0 ? String(rawRow[photoCol - 1] || '').trim() : '';
    var photoUrl = photoRaw ? (resolveFormUploadedFile_(photoRaw) || photoRaw) : '';
    if (photoRaw && resolveFormUploadedFile_(photoRaw)) {
      // it's a Drive link/ID we can organize into the shared folder
      var docsFolder = getRegistrationDocsFolder_();
      organizeApplicantFile_(photoRaw, docsFolder, newId + '_Photo');
    }
    // otherwise (any other link, e.g. WhatsApp/photo host) it's stored as-is

    var newRow = new Array(NONMEM_LAST_COL).fill('');
    newRow[NONMEM_COL.id - 1] = newId;
    newRow[NONMEM_COL.name - 1] = name;
    newRow[NONMEM_COL.address - 1] = buildVisitorAddress_(headers, rawRow);
    newRow[NONMEM_COL.phone - 1] = phoneCol > 0 ? rawRow[phoneCol - 1] : '';
    newRow[NONMEM_COL.purpose - 1] = ''; // this form has no Purpose question
    newRow[NONMEM_COL.vip - 1] = '';
    newRow[NONMEM_COL.specialMsg - 1] = '';
    newRow[NONMEM_COL.photoPath - 1] = photoUrl || '';
    newRow[NONMEM_COL.registerDate - 1] = new Date();
    newRow[NONMEM_COL.occupation - 1] = occupationCol > 0 ? rawRow[occupationCol - 1] : '';
    var visitorEmail = emailCol > 0 ? String(rawRow[emailCol - 1] || '').trim() : '';
    newRow[NONMEM_COL.email - 1] = visitorEmail;

    var writeRow = Math.max(nLast + 1, NONMEM_START_ROW);
    nsh.getRange(writeRow, 1, 1, NONMEM_LAST_COL).setValues([newRow]);

    sh.getRange(rowNum, approvedIdCol).setValue(newId);
    sh.getRange(rowNum, statusCol).setValue('GRANTED - Visitor ID ' + newId);

    if (visitorEmail) {
      try {
        MailApp.sendEmail({
          to: visitorEmail,
          subject: 'Visitor Registration Approved - Branch Central Library, Honavar',
          body: 'Dear ' + name + ',\n\n' +
            'Your visitor registration at Branch Central Library, Honavar has been approved.\n\n' +
            'Your Visitor ID is: ' + newId + '\n\n' +
            mmFooter_()
        });
        logEmail_('Other', 'Visitor Registration Approved', newId, name, visitorEmail, '', 'Sent');
      } catch (mailErr) {
        logEmail_('Other', 'Visitor Registration Approved', newId, name, visitorEmail, '', 'Failed');
      }
    }

    return { success: true, visitorId: newId, name: name };
  } finally {
    lock.releaseLock();
  }
}

// Fires automatically when someone submits the membership form - emails them
// a direct link to their own pre-filled printable copy, so they don't need
// to come back to the library or ask staff for their row number.
function onMembershipFormSubmit(e) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pending Registrations');
    if (!sh) return;
    var headers = ensurePendingRegTrackingCols_(sh);
    var lastRow = sh.getLastRow();
    var emailCol = pendRegColIndex_(headers, 'Email ID');
    var nameCol = pendRegColIndex_(headers, 'Full Name');
    if (emailCol === -1 || nameCol === -1) return;

    var email = sh.getRange(lastRow, emailCol).getValue();
    var name = sh.getRange(lastRow, nameCol).getValue();
    if (!email) return; // nothing to send to - applicant can still get a printout from staff in person

    var printLink = ScriptApp.getService().getUrl() + '?page=regprint&row=' + lastRow;

    MailApp.sendEmail({
      to: email,
      subject: 'Your Library Membership Application - BCL Honavar',
      body: 'Dear ' + name + ',\n\n' +
        'Thank you for applying for membership at Branch Central Library, Honavar.\n\n' +
        'Click the link below to view and print your pre-filled application form:\n' + printLink + '\n\n' +
        'Next steps:\n' +
        '1. Print this form.\n' +
        '2. Get it signed by a Gazetted Officer.\n' +
        '3. Bring the signed form, your documents, and the nominal fee to the library.\n\n' +
        'You will receive another email once your membership is approved.\n\n' +
        mmFooter_()
    });
    logEmail_('Other', 'Application Received', '', name, email, '', 'Sent');
  } catch (err) {
    try { logEmail_('Other', 'Application Received', '', name, email, '', 'Failed'); } catch (e2) {}
  }
}

function getPendingRegistrationForPrint(rowNum) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pending Registrations');
  if (!sh) throw new Error('Pending Registrations sheet not found - run setupMemberRegistrationForm first.');
  rowNum = parseInt(rowNum, 10);
  if (!rowNum || rowNum < 2) throw new Error('Invalid row.');

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  var data = {};
  headers.forEach(function(h, i) { data[String(h).trim()] = row[i]; });

  function get(label) {
    for (var key in data) { if (key.indexOf(label) === 0) return data[key]; }
    return '';
  }

  var dob = get('Date of Birth');
  var age = '';
  if (dob instanceof Date) {
    var diffMs = Date.now() - dob.getTime();
    age = Math.floor(diffMs / (365.25 * 86400000));
  }

  return {
    row: rowNum,
    name: get('Full Name'),
    dob: fmtDate_(dob),
    age: age,
    gender: get('Gender'),
    state: get('State'),
    district: get('District'),
    village: get('Town / Village'),
    street: get('Street'),
    houseNo: get('House Number'),
    pincode: get('Pincode'),
    occupation: get('Occupation'),
    officeAddress: get('Office Address'),
    education: get('Education'),
    phone: get('Phone Number'),
    email: get('Email ID'),
    photoUrl: resolveFormUploadedFile_(get('Photo')),
    submittedDate: fmtDate_(row[0]), // Timestamp is always column A on a Form response sheet
    submittedMonth: (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'MMMM yyyy') : ''
  };
}

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || '';
  if (page === 'deadstock') return dsWebPage_();   // Dead Stock Register page (DeadStockWeb.gs)
  if (page === 'public') {
    return HtmlService.createHtmlOutputFromFile('Public')
      .setTitle('BCL HNR - Public Book Search')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'gate') {
    return HtmlService.createHtmlOutputFromFile('Gate')
      .setTitle('BCL HNR - Gate Entry')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'entry') {
    return HtmlService.createHtmlOutputFromFile('PublicGate')
      .setTitle('BCL HNR - Welcome')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'kiosk') {
    return HtmlService.createHtmlOutputFromFile('Kiosk')
      .setTitle('BCL HNR - Field Kiosk')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'links') {
    return HtmlService.createHtmlOutputFromFile('Links')
      .setTitle('BCL HNR - All Links')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'regprint') {
    var tmpl = HtmlService.createTemplateFromFile('RegistrationPrint');
    tmpl.rowParam = (e && e.parameter && e.parameter.row) || '';
    return tmpl.evaluate()
      .setTitle('BCL HNR - Membership Application')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'irkiosk') {
    return HtmlService.createHtmlOutputFromFile('IssueReturnKiosk')
      .setTitle('BCL HNR - Issue/Return Kiosk')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('BCL_HNR_EXPANDED - Stack Verification & Cards')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ================================================================
// AUTHORITY REPORT - instant printable answers for the statistics
// and document questions higher authorities/visitors ask for.
//
// Paste this ENTIRE section at the END of your Code.gs (after the
// last line of the file). Nothing here touches or redefines anything
// above it - it only adds new sheets/columns it needs (via
// setupAuthorityReportSystem.gs, run once first) and new functions.
//
// Design: every answer is computed live from your real sheets each
// time you generate a report - nothing is pre-calculated or cached,
// so it's always current. A handful of facts genuinely can't be
// computed from existing data (founding year, staff list, monthly
// bills) - those come from the small LIBRARY_PROFILE / STAFF_LIST /
// MONTHLY_BILLS / NEWSPAPER_MASTER sheets you fill in once (see the
// setup script and the install guide).
// ================================================================

var AR_BOOKS_ACC_DATE_COL = 17; // Books Catalogue column Q (Accession Date)
var AR_BOOKS_LANG_COL = 18;     // Books Catalogue column R (Language)
var AR_DECARD_STAFF_COL = 9;    // Decard Books column I (Staff Name)

// ---- period + date helpers ---------------------------------------------

// period = { year: 2026, month: 9 }  -> that specific month
// period = { year: 2026, month: 0 }  -> that whole year
// period = { year: 0,    month: 0 }  -> all-time, no date filter
function arPeriodRange_(period) {
  var year = Number(period && period.year) || 0;
  var month = Number(period && period.month) || 0;
  if (!year) return { start: null, end: null, allTime: true };
  if (month) {
    var start = new Date(year, month - 1, 1);
    var end = new Date(year, month, 1); // exclusive
    return { start: start, end: end, allTime: false };
  }
  return { start: new Date(year, 0, 1), end: new Date(year + 1, 0, 1), allTime: false };
}

function arInRange_(d, range) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return false;
  if (range.allTime) return true;
  return d >= range.start && d < range.end;
}

function arDaysInRange_(range, todayIfOpen) {
  if (range.allTime) return null;
  var end = range.end;
  if (todayIfOpen && todayIfOpen < end) end = todayIfOpen;
  var ms = end.getTime() - range.start.getTime();
  return Math.max(1, Math.round(ms / 86400000));
}

function arPeriodLabel_(period) {
  var year = Number(period && period.year) || 0;
  var month = Number(period && period.month) || 0;
  if (!year) return 'All time';
  if (month) return Utilities.formatDate(new Date(year, month - 1, 1), Session.getScriptTimeZone(), 'MMMM yyyy');
  return String(year);
}

// ---- small config sheets -------------------------------------------------

function arReadProfile_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LIBRARY_PROFILE');
  var map = {};
  if (!sh) return map;
  var last = sh.getLastRow();
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function(r) {
    var key = String(r[0] || '').trim();
    if (key) map[key] = r[1];
  });
  return map;
}

function arReadStaffList_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STAFF_LIST');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues()
    .map(function(r) { return { name: r[0], designation: r[1], joined: r[2], status: r[3], notes: r[4] }; })
    .filter(function(r) { return String(r.name || '').trim() && String(r.name).indexOf('(example)') === -1; });
}

function arReadMonthlyBills_(monthKey) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MONTHLY_BILLS');
  if (!sh) return null;
  var last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, 6).getValues();
  for (var i = 0; i < rows.length; i++) {
    var key = String(rows[i][0] || '').trim();
    if (key === monthKey) {
      return { electric: rows[i][1], newspaper: rows[i][2], other: rows[i][3], oldPaperSale: rows[i][4], notes: rows[i][5] };
    }
  }
  return null;
}

function arReadNewspaperMaster_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NEWSPAPER_MASTER');
  var map = {};
  if (!sh) return map;
  var last = sh.getLastRow();
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 4).getValues().forEach(function(r) {
    var name = String(r[0] || '').trim();
    if (!name || name.indexOf('(example)') === 0) return;
    map[name.toLowerCase()] = { type: String(r[1] || '').trim(), language: String(r[2] || '').trim(), postal: String(r[3] || '').trim() };
  });
  return map;
}

// ---- INSTITUTION & STAFF --------------------------------------------------

function ar_institutionFacts_() {
  var p = arReadProfile_();
  return {
    establishedYear: p['Library Established Year'] || '',
    buildingStartYear: p['Current Building - Construction Started Year'] || '',
    buildingOpenedYear: p['Current Building - Opened/Inaugurated Year'] || '',
    workingHours: p['Working Hours'] || '',
    weeklyHoliday: p['Weekly Holiday(s)'] || '',
    otherHolidays: p['Other Holidays'] || '',
    seatingCapacity: p['Sitting Capacity'] || '',
    floorArea: p['Total Floor Area (sq ft)'] || '',
    note: p['Any Other Note'] || ''
  };
}

function ar_totalStaffActive_() {
  var staff = arReadStaffList_();
  return staff.filter(function(s) { return String(s.status || '').trim().toLowerCase() === 'active'; }).length;
}

function ar_staffListText_() {
  var staff = arReadStaffList_();
  if (!staff.length) return '(STAFF_LIST sheet is empty - add staff there)';
  return staff.map(function(s) {
    return s.name + (s.designation ? ' - ' + s.designation : '') + (String(s.status||'').toLowerCase() !== 'active' ? ' [' + s.status + ']' : '');
  }).join('\n');
}

// ---- MEMBERSHIP ------------------------------------------------------------

function arAllMembersRows_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var last = sh.getLastRow();
  var numRows = Math.max(last - MEMBERS_START_ROW + 1, 0);
  if (!numRows) return [];
  return sh.getRange(MEMBERS_START_ROW, 1, numRows, MEMBERS_LAST_COL).getValues();
}

function ar_totalActiveMembers_() {
  var rows = arAllMembersRows_();
  var n = 0;
  rows.forEach(function(r) {
    var id = norm_(r[MEMBERS_COL.id - 1]);
    if (!id) return;
    var status = String(r[MEMBERS_COL.status - 1] || '').toUpperCase();
    if (status !== 'CLOSED' && status !== 'INACTIVE') n++;
  });
  return n;
}

function ar_membersMaleFemale_() {
  var rows = arAllMembersRows_();
  var male = 0, female = 0, other = 0;
  rows.forEach(function(r) {
    var id = norm_(r[MEMBERS_COL.id - 1]);
    if (!id) return;
    var g = String(r[MEMBERS_COL.gender - 1] || '').trim().toUpperCase();
    if (g.indexOf('M') === 0) male++;
    else if (g.indexOf('F') === 0) female++;
    else other++;
  });
  return { male: male, female: female, other: other };
}

function ar_newMembersInPeriod_(period) {
  var range = arPeriodRange_(period);
  var rows = arAllMembersRows_();
  var n = 0;
  rows.forEach(function(r) {
    var id = norm_(r[MEMBERS_COL.id - 1]);
    if (!id) return;
    var d = r[MEMBERS_COL.membStart - 1];
    if (arInRange_(d, range)) n++;
  });
  return n;
}

function ar_yearwiseMembershipTable_(numYears) {
  var rows = arAllMembersRows_();
  var thisYear = new Date().getFullYear();
  var counts = {};
  for (var y = thisYear - numYears + 1; y <= thisYear; y++) counts[y] = 0;
  rows.forEach(function(r) {
    var id = norm_(r[MEMBERS_COL.id - 1]);
    if (!id) return;
    var d = r[MEMBERS_COL.membStart - 1];
    if (d instanceof Date) {
      var y = d.getFullYear();
      if (counts.hasOwnProperty(y)) counts[y]++;
    }
  });
  var out = [];
  Object.keys(counts).sort().forEach(function(y) { out.push([y, counts[y]]); });
  return { columns: ['Year', 'New Members'], rows: out };
}

function ar_feeCollectedInPeriod_(period) {
  var range = arPeriodRange_(period);
  var rows = arAllMembersRows_();
  var total = 0, count = 0;
  rows.forEach(function(r) {
    var id = norm_(r[MEMBERS_COL.id - 1]);
    if (!id) return;
    var d = r[MEMBERS_COL.receiptDate - 1];
    if (arInRange_(d, range)) { total += Number(r[MEMBERS_COL.amount - 1]) || 0; count++; }
  });
  return { total: total, count: count };
}

// ---- BOOKS / ACCESSION REGISTER --------------------------------------------

function arAllBooksRows_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Books Catalogue');
  var last = sh.getLastRow();
  var numRows = Math.max(last - BOOKS_START_ROW + 1, 0);
  if (!numRows) return [];
  return sh.getRange(BOOKS_START_ROW, 1, numRows, AR_BOOKS_LANG_COL).getValues();
}

function ar_totalBooksInCatalogue_() {
  var rows = arAllBooksRows_();
  var n = 0;
  rows.forEach(function(r) { if (norm_(r[BOOKS_COL.id - 1])) n++; });
  return n;
}

function ar_booksAddedInPeriod_(period) {
  var range = arPeriodRange_(period);
  var rows = arAllBooksRows_();
  var n = 0, untracked = 0;
  rows.forEach(function(r) {
    if (!norm_(r[BOOKS_COL.id - 1])) return;
    var d = r[AR_BOOKS_ACC_DATE_COL - 1];
    if (!(d instanceof Date)) { untracked++; return; }
    if (arInRange_(d, range)) n++;
  });
  return { count: n, untracked: untracked };
}

function ar_yearwiseAccessionTable_(numYears) {
  var rows = arAllBooksRows_();
  var thisYear = new Date().getFullYear();
  var counts = {}, untracked = 0;
  for (var y = thisYear - numYears + 1; y <= thisYear; y++) counts[y] = 0;
  rows.forEach(function(r) {
    if (!norm_(r[BOOKS_COL.id - 1])) return;
    var d = r[AR_BOOKS_ACC_DATE_COL - 1];
    if (!(d instanceof Date)) { untracked++; return; }
    var y = d.getFullYear();
    if (counts.hasOwnProperty(y)) counts[y]++;
  });
  var out = [];
  Object.keys(counts).sort().forEach(function(y) { out.push([y, counts[y]]); });
  return { columns: ['Year', 'Books Accessioned'], rows: out, note: untracked ? (untracked + ' book(s) have no Accession Date filled in yet, so are not counted in any year.') : '' };
}

// ---- DISCARDED BOOKS --------------------------------------------------------

function arAllDiscardRows_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Decard Books');
  var last = sh.getLastRow();
  var numRows = Math.max(last - DECARD_START_ROW + 1, 0);
  if (!numRows) return [];
  return sh.getRange(DECARD_START_ROW, 1, numRows, AR_DECARD_STAFF_COL).getValues();
}

function ar_totalDiscarded_() {
  var rows = arAllDiscardRows_();
  var n = 0;
  rows.forEach(function(r) { if (norm_(r[DECARD_COL.id - 1])) n++; });
  return n;
}

function ar_discardedInPeriod_(period) {
  var range = arPeriodRange_(period);
  var rows = arAllDiscardRows_();
  var n = 0;
  rows.forEach(function(r) {
    if (!norm_(r[DECARD_COL.id - 1])) return;
    if (arInRange_(r[DECARD_COL.decardDate - 1], range)) n++;
  });
  return n;
}

function ar_yearwiseDiscardTable_(numYears) {
  var rows = arAllDiscardRows_();
  var thisYear = new Date().getFullYear();
  var counts = {};
  for (var y = thisYear - numYears + 1; y <= thisYear; y++) counts[y] = { count: 0, staff: {} };
  rows.forEach(function(r) {
    if (!norm_(r[DECARD_COL.id - 1])) return;
    var d = r[DECARD_COL.decardDate - 1];
    if (!(d instanceof Date)) return;
    var y = d.getFullYear();
    if (!counts.hasOwnProperty(y)) return;
    counts[y].count++;
    var staffName = String(r[AR_DECARD_STAFF_COL - 1] || '').trim() || 'Unrecorded';
    counts[y].staff[staffName] = (counts[y].staff[staffName] || 0) + 1;
  });
  var out = [];
  Object.keys(counts).sort().forEach(function(y) {
    var c = counts[y];
    var staffSummary = Object.keys(c.staff).map(function(n) { return n + ' (' + c.staff[n] + ')'; }).join(', ');
    out.push([y, c.count, staffSummary || '-']);
  });
  return { columns: ['Year', 'Discarded', 'By Staff'], rows: out };
}

// ---- CIRCULATION (Book Issue) -----------------------------------------------

function arAllIssueRows_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BI_SHEET);
  var last = sh.getLastRow();
  var numRows = Math.max(last - BI_START_ROW + 1, 0);
  if (!numRows) return [];
  return sh.getRange(BI_START_ROW, 1, numRows, BI_LAST_COL).getValues();
}

function ar_issuedInPeriod_(period) {
  var range = arPeriodRange_(period);
  var rows = arAllIssueRows_();
  var n = 0;
  rows.forEach(function(r) {
    if (!norm_(r[BI_COL.issueId - 1]) && !norm_(r[BI_COL.bookId - 1])) return;
    if (arInRange_(r[BI_COL.issueDate - 1], range)) n++;
  });
  return n;
}

function ar_avgIssuedPerDay_(period) {
  var total = ar_issuedInPeriod_(period);
  var range = arPeriodRange_(period);
  var days = arDaysInRange_(range, stripTime_(new Date()));
  if (!days) return { total: total, avgPerDay: null };
  return { total: total, avgPerDay: Math.round((total / days) * 100) / 100 };
}

function ar_returnedInPeriod_(period) {
  var range = arPeriodRange_(period);
  var rows = arAllIssueRows_();
  var n = 0;
  rows.forEach(function(r) {
    if (arInRange_(r[BI_COL.returnDate - 1], range)) n++;
  });
  return n;
}

function ar_pendingDueTotal_() {
  var rows = arAllIssueRows_();
  var n = 0;
  var today = stripTime_(new Date());
  rows.forEach(function(r) {
    if (!norm_(r[BI_COL.bookId - 1])) return;
    var returned = r[BI_COL.returnDate - 1];
    if (returned instanceof Date) return; // already returned
    var status = String(r[BI_COL.status - 1] || '').toUpperCase();
    if (status === 'RETURNED') return;
    n++;
  });
  return n;
}

function ar_pendingDueInPeriod_(period) {
  // Books issued within this period that are STILL not returned (as of today).
  var range = arPeriodRange_(period);
  var rows = arAllIssueRows_();
  var n = 0;
  rows.forEach(function(r) {
    if (!norm_(r[BI_COL.bookId - 1])) return;
    if (!arInRange_(r[BI_COL.issueDate - 1], range)) return;
    var returned = r[BI_COL.returnDate - 1];
    if (returned instanceof Date) return;
    var status = String(r[BI_COL.status - 1] || '').toUpperCase();
    if (status === 'RETURNED') return;
    n++;
  });
  return n;
}

function ar_fineCollectedInPeriod_(period) {
  var range = arPeriodRange_(period);
  var rows = arAllIssueRows_();
  var total = 0;
  rows.forEach(function(r) {
    if (arInRange_(r[BI_COL.returnDate - 1], range)) total += Number(r[BI_COL.fine - 1]) || 0;
  });
  return total;
}

// ---- VISITORS / GATE ENTRY ---------------------------------------------------

function arAllGateRows_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GATE_ENTRY');
  var last = sh.getLastRow();
  var numRows = Math.max(last - GATE_START_ROW + 1, 0);
  if (!numRows) return [];
  return sh.getRange(GATE_START_ROW, 1, numRows, GATE_LAST_COL).getValues();
}

function ar_visitorsInPeriod_(period) {
  var range = arPeriodRange_(period);
  var rows = arAllGateRows_();
  var n = 0;
  rows.forEach(function(r) {
    if (String(r[GATE_COL.inOut - 1] || '').toUpperCase() !== 'IN') return;
    if (arInRange_(r[GATE_COL.date - 1], range)) n++;
  });
  return n;
}

function ar_avgVisitorsPerDay_(period) {
  var total = ar_visitorsInPeriod_(period);
  var range = arPeriodRange_(period);
  var days = arDaysInRange_(range, stripTime_(new Date()));
  if (!days) return { total: total, avgPerDay: null };
  return { total: total, avgPerDay: Math.round((total / days) * 100) / 100 };
}

// Best-effort: non-member visitors whose stated purpose mentions reading/
// reference. This is an estimate, not an exact log - flagged as such in
// the printed report, since "reference reader" isn't its own tracked field.
function ar_referenceReadersInPeriod_(period) {
  var range = arPeriodRange_(period);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var purposeById = {};
  var nm = ss.getSheetByName('NON_MEMBERS');
  if (nm) {
    var last = nm.getLastRow();
    var numRows = Math.max(last - NONMEM_START_ROW + 1, 0);
    if (numRows) {
      nm.getRange(NONMEM_START_ROW, 1, numRows, NONMEM_LAST_COL).getValues().forEach(function(r) {
        var id = norm_(r[NONMEM_COL.id - 1]);
        if (id) purposeById[id] = String(r[NONMEM_COL.purpose - 1] || '').toLowerCase();
      });
    }
  }
  var rows = arAllGateRows_();
  var n = 0;
  rows.forEach(function(r) {
    if (String(r[GATE_COL.inOut - 1] || '').toUpperCase() !== 'IN') return;
    if (!arInRange_(r[GATE_COL.date - 1], range)) return;
    var id = norm_(r[GATE_COL.id - 1]);
    var purpose = purposeById[id] || '';
    if (purpose.indexOf('read') !== -1 || purpose.indexOf('reference') !== -1) n++;
  });
  return n;
}

// ---- NEWSPAPERS / MAGAZINES ----------------------------------------------

function arDailyNewspaperCountsInPeriod_(period) {
  // Sums checkmarks across Daily_This_Month / Daily_Last_Month / Daily_Archive_YYYY_MM
  // sheets, per newspaper name, for every day inside the period.
  var range = arPeriodRange_(period);
  if (range.allTime) return {}; // daily attendance isn't practical to sum with no bound - require a year
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var counts = {}; // name -> count
  var cur = new Date(range.start);
  var monthsSeen = {};
  while (cur < range.end) {
    var key = cur.getFullYear() + '-' + ('0' + (cur.getMonth() + 1)).slice(-2);
    monthsSeen[key] = new Date(cur.getFullYear(), cur.getMonth(), 1);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Object.keys(monthsSeen).forEach(function(key) {
    var monthStart = monthsSeen[key];
    var sh = npSheetForDate_(monthStart);
    if (!sh) return;
    var lastRow = npLastRealRow_(sh);
    var numRows = Math.max(lastRow - NP_START_ROW + 1, 0);
    if (!numRows) return;
    var daysInMonth = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0).getDate();
    var names = sh.getRange(NP_START_ROW, NP_NAME_COL, numRows, 1).getValues();
    var marks = sh.getRange(NP_START_ROW, NP_FIRST_DAY_COL, numRows, daysInMonth).getValues();
    for (var i = 0; i < numRows; i++) {
      var name = String(names[i][0] || '').trim();
      if (!name) continue;
      for (var d = 0; d < daysInMonth; d++) {
        var dayDate = new Date(monthStart.getFullYear(), monthStart.getMonth(), d + 1);
        if (!arInRange_(dayDate, range)) continue;
        if (String(marks[i][d]).trim() === '\u2713') counts[name] = (counts[name] || 0) + 1;
      }
    }
  });
  return counts;
}

function arWeeklyQuarterlyMonthlyCountsInPeriod_(period) {
  // Sums issues received within the period, per magazine name, from the
  // Weekly and Quatarly_Monthly sheets (both continuous, filtered by
  // real Received Date - see getWeeklyQuarterlyMonthlyStatus for the
  // same parsing approach).
  var year = Number(period && period.year) || 0;
  var month = Number(period && period.month) || 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var counts = { weekly: {}, quarterly: {}, monthly: {} };

  // startRow=1 for both (was hardcoded 6/7) - see the comment in
  // getWeeklyQuarterlyMonthlyStatus above for why starting earlier is
  // always safe here, while starting even one row too late silently drops
  // the sheet's first magazine entirely.
  var weekly = ss.getSheetByName('Weekly');
  if (weekly) {
    parseSectionedMagazines_(weekly, 1, null, null).forEach(function(s) {
      s.issues.forEach(function(it) {
        if (issueMatchesMonthEither_(it.pubDateRaw, it.recvDateRaw, year, month)) counts.weekly[s.name] = (counts.weekly[s.name] || 0) + 1;
      });
    });
  }
  var qm = ss.getSheetByName('Quatarly_Monthly');
  if (qm) {
    var marker = findMonthlyMarkerRow_(qm);
    parseSectionedMagazines_(qm, 1, null, marker ? marker - 1 : null).forEach(function(s) {
      s.issues.forEach(function(it) {
        if (issueMatchesMonthEither_(it.pubDateRaw, it.recvDateRaw, year, month)) counts.quarterly[s.name] = (counts.quarterly[s.name] || 0) + 1;
      });
    });
    if (marker) {
      parseMonthlyFlatList_(qm, marker + 1, null).forEach(function(it) {
        if (issueMatchesMonthEither_(it.pubDateRaw, it.recvDateRaw, year, month)) counts.monthly[it.name] = (counts.monthly[it.name] || 0) + 1;
      });
    }
  }
  return counts;
}

function ar_newspaperStats_(period) {
  var master = arReadNewspaperMaster_();
  var daily = arDailyNewspaperCountsInPeriod_(period);
  var wqm = arWeeklyQuarterlyMonthlyCountsInPeriod_(period);

  var totals = { daily: 0, weekly: 0, quarterly: 0, monthly: 0, kannada: 0, english: 0, other: 0, postal: 0, local: 0, unclassified: 0 };

  function tally(name, count, typeBucket) {
    totals[typeBucket] += count;
    var meta = master[String(name).trim().toLowerCase()];
    if (!meta) { totals.unclassified += count; return; }
    var lang = (meta.language || '').toLowerCase();
    if (lang.indexOf('kan') === 0) totals.kannada += count;
    else if (lang.indexOf('eng') === 0) totals.english += count;
    else totals.other += count;
    if ((meta.postal || '').toLowerCase().indexOf('y') === 0) totals.postal += count;
    else totals.local += count;
  }

  Object.keys(daily).forEach(function(name) { tally(name, daily[name], 'daily'); });
  Object.keys(wqm.weekly).forEach(function(name) { tally(name, wqm.weekly[name], 'weekly'); });
  Object.keys(wqm.quarterly).forEach(function(name) { tally(name, wqm.quarterly[name], 'quarterly'); });
  Object.keys(wqm.monthly).forEach(function(name) { tally(name, wqm.monthly[name], 'monthly'); });

  return totals;
}

// ---- FINANCE / BILLS --------------------------------------------------------

function ar_billsForMonth_(period) {
  var year = Number(period && period.year) || 0;
  var month = Number(period && period.month) || 0;
  if (!year || !month) return null;
  var key = year + '-' + ('0' + month).slice(-2);
  var row = arReadMonthlyBills_(key);
  if (!row) return { found: false, monthKey: key };
  return {
    found: true, monthKey: key,
    electric: Number(row.electric) || 0,
    newspaper: Number(row.newspaper) || 0,
    other: Number(row.other) || 0,
    oldPaperSale: Number(row.oldPaperSale) || 0,
    notes: row.notes || ''
  };
}

// ================================================================
// AUTHORITY REPORT - DOCUMENT LIBRARY. Lets you upload and keep actual
// documents (building RTC / land records, staff documents, awards,
// important photos, etc.) - stored in Google Drive, linked from a new
// AUTHORITY_DOCUMENTS sheet (added by setupAuthorityReportSystem), and
// pulled straight into the printable report: photos are embedded inline,
// everything else shows as a clickable link.
// ================================================================

var AR_DOC_CATEGORIES = [
  'Building / Land Records (RTC, Sale Deed, etc.)',
  'Staff Documents',
  'Awards & Recognition',
  'Important Photos',
  'Other Records'
];

var AR_DOC_COL = {
  timestamp: 1, category: 2, title: 3, description: 4, fileName: 5,
  mimeType: 6, url: 7, fileId: 8, isPhoto: 9, uploadedBy: 10
};
var AR_DOC_START_ROW = 2;
var AR_DOC_LAST_COL = 10;

function getAuthorityDocCategories() {
  return AR_DOC_CATEGORIES;
}

function getAuthorityDocsFolder_() {
  var folderName = 'BCL Honavar - Authority Report Documents';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function arDocsSheet_() {
  var sh = ss_().getSheetByName('AUTHORITY_DOCUMENTS');
  if (!sh) throw new Error('AUTHORITY_DOCUMENTS sheet not found - run setupAuthorityReportSystem once first (it now also creates this sheet).');
  return sh;
}

function arDocsLastRow_(sh) {
  var last = sh.getLastRow();
  var numRows = Math.max(last - AR_DOC_START_ROW + 1, 0);
  if (!numRows) return AR_DOC_START_ROW - 1;
  var col = sh.getRange(AR_DOC_START_ROW, AR_DOC_COL.timestamp, numRows, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]).trim() !== '') return AR_DOC_START_ROW + i;
  }
  return AR_DOC_START_ROW - 1;
}

// dataUrl: "data:<mime>;base64,....." straight from a <input type=file>'s
// FileReader on the browser side - see uploadAuthorityDoc() in Index.html.
function uploadAuthorityDocument(category, title, description, dataUrl, fileName, uploadedBy) {
  category = String(category || '').trim();
  title = String(title || '').trim();
  if (AR_DOC_CATEGORIES.indexOf(category) === -1) throw new Error('Choose a valid category.');
  if (!title) throw new Error('Enter a title for this document.');
  if (!dataUrl) throw new Error('Choose a file to upload.');

  var m = String(dataUrl).match(/^data:([^;]+);base64,(.*)$/);
  if (!m) throw new Error('Could not read the file - please try again.');
  var mimeType = m[1];
  var base64Data = m[2];
  var isPhoto = mimeType.indexOf('image/') === 0;
  var safeName = String(fileName || (title + (isPhoto ? '.jpg' : ''))).replace(/[\\/:*?"<>|]/g, '_');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var folder = getAuthorityDocsFolder_();
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, safeName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var sh = arDocsSheet_();
    var row = arDocsLastRow_(sh) + 1;
    sh.getRange(row, 1, 1, AR_DOC_LAST_COL).setValues([[
      new Date(), category, title, description || '', safeName, mimeType,
      file.getUrl(), file.getId(), isPhoto, uploadedBy || ''
    ]]);
    return { success: true, url: file.getUrl(), row: row };
  } finally {
    lock.releaseLock();
  }
}

function getAuthorityDocuments() {
  var sh = ss_().getSheetByName('AUTHORITY_DOCUMENTS');
  if (!sh) return [];
  var last = arDocsLastRow_(sh);
  var numRows = Math.max(last - AR_DOC_START_ROW + 1, 0);
  if (!numRows) return [];
  var values = sh.getRange(AR_DOC_START_ROW, 1, numRows, AR_DOC_LAST_COL).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!String(r[AR_DOC_COL.title - 1] || '').trim()) continue;
    out.push({
      row: AR_DOC_START_ROW + i,
      timestamp: fmtDate_(r[AR_DOC_COL.timestamp - 1]),
      category: r[AR_DOC_COL.category - 1],
      title: r[AR_DOC_COL.title - 1],
      description: r[AR_DOC_COL.description - 1],
      fileName: r[AR_DOC_COL.fileName - 1],
      mimeType: r[AR_DOC_COL.mimeType - 1],
      url: r[AR_DOC_COL.url - 1],
      fileId: r[AR_DOC_COL.fileId - 1],
      isPhoto: r[AR_DOC_COL.isPhoto - 1] === true || String(r[AR_DOC_COL.isPhoto - 1]).toUpperCase() === 'TRUE',
      uploadedBy: r[AR_DOC_COL.uploadedBy - 1]
    });
  }
  out.sort(function(a, b) { return b.row - a.row; }); // most recent first
  return out;
}

function deleteAuthorityDocument(row) {
  var sh = arDocsSheet_();
  row = parseInt(row, 10);
  if (!row || row < AR_DOC_START_ROW) throw new Error('Invalid row.');
  var fileId = sh.getRange(row, AR_DOC_COL.fileId).getValue();
  sh.deleteRow(row);
  if (fileId) {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* file already gone - ignore */ }
  }
  return { success: true };
}

// For the printable report: photos are embedded inline as base64, so the
// PDF is self-contained (works even if the person printing it doesn't have
// Drive access to your account). Anything too large to embed reasonably,
// or that isn't a photo (a PDF land record, say), is listed with its
// clickable Drive link instead - embedding a whole PDF inline on a printed
// page isn't practical.
var AR_DOC_EMBED_MAX_BYTES = 4 * 1024 * 1024;

function ar_documentsForCategory_(category) {
  var all = getAuthorityDocuments().filter(function(d) { return d.category === category; });
  var rows = all.map(function(d) {
    var out = { title: d.title, description: d.description, url: d.url, isPhoto: d.isPhoto, dataUrl: '' };
    if (d.isPhoto && d.fileId) {
      try {
        var blob = DriveApp.getFileById(d.fileId).getBlob();
        if (blob.getBytes().length <= AR_DOC_EMBED_MAX_BYTES) {
          out.dataUrl = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
        }
      } catch (e) { /* couldn't fetch the file for embedding - it'll still show as a link */ }
    }
    return out;
  });
  return {
    kind: 'documents', category: category, rows: rows,
    note: rows.length ? '' : 'No documents uploaded yet in this category.'
  };
}

// ================================================================
// QUESTION REGISTRY - one entry per printable answer. `compute`
// takes the selected period and returns either:
//   { kind:'value', value: <string/number>, note: <string?> }
//   { kind:'table', columns: [...], rows: [[...],...], note: <string?> }
//   { kind:'documents', category: <string>, rows: [...], note: <string?> }
// ================================================================

function AR_REGISTRY_() {
  return [
    // -- Institution & Staff --
    { id: 'lib_established_year', cat: 'Institution & Staff', label: 'Library established (year)',
      compute: function() { return { kind: 'value', value: ar_institutionFacts_().establishedYear || 'Not filled in yet (LIBRARY_PROFILE)' }; } },
    { id: 'building_start_year', cat: 'Institution & Staff', label: 'Current building - construction started (year)',
      compute: function() { return { kind: 'value', value: ar_institutionFacts_().buildingStartYear || 'Not filled in yet (LIBRARY_PROFILE)' }; } },
    { id: 'building_opened_year', cat: 'Institution & Staff', label: 'Current building - opened/inaugurated (year)',
      compute: function() { return { kind: 'value', value: ar_institutionFacts_().buildingOpenedYear || 'Not filled in yet (LIBRARY_PROFILE)' }; } },
    { id: 'working_hours', cat: 'Institution & Staff', label: 'Working hours',
      compute: function() { return { kind: 'value', value: ar_institutionFacts_().workingHours || 'Not filled in yet (LIBRARY_PROFILE)' }; } },
    { id: 'holidays', cat: 'Institution & Staff', label: 'Weekly / other holidays',
      compute: function() { var f = ar_institutionFacts_(); return { kind: 'value', value: [f.weeklyHoliday, f.otherHolidays].filter(Boolean).join(' | ') || 'Not filled in yet (LIBRARY_PROFILE)' }; } },
    { id: 'staff_total', cat: 'Institution & Staff', label: 'Total staff currently working',
      compute: function() { return { kind: 'value', value: ar_totalStaffActive_() }; } },
    { id: 'staff_list', cat: 'Institution & Staff', label: 'Staff list (name - designation)',
      compute: function() { return { kind: 'value', value: ar_staffListText_() }; } },

    // -- Membership --
    { id: 'members_active_total', cat: 'Membership', label: 'Total active members (as of today)',
      compute: function() { return { kind: 'value', value: ar_totalActiveMembers_() }; } },
    { id: 'members_male_female', cat: 'Membership', label: 'Total members - male / female',
      compute: function() { var g = ar_membersMaleFemale_(); return { kind: 'value', value: 'Male: ' + g.male + '  |  Female: ' + g.female + (g.other ? '  |  Not specified: ' + g.other : '') }; } },
    { id: 'members_new_in_period', cat: 'Membership', label: 'New members in selected period',
      compute: function(period) { return { kind: 'value', value: ar_newMembersInPeriod_(period) }; } },
    { id: 'members_yearwise', cat: 'Membership', label: 'Year-wise new membership (last 10 years)',
      compute: function() { return Object.assign({ kind: 'table' }, ar_yearwiseMembershipTable_(10)); } },
    { id: 'fee_collected_period', cat: 'Membership', label: 'Membership fee collected in selected period',
      compute: function(period) { var f = ar_feeCollectedInPeriod_(period); return { kind: 'value', value: '\u20b9' + f.total + ' (' + f.count + ' receipt(s))' }; } },

    // -- Books / Accession Register --
    { id: 'books_total', cat: 'Books / Accession Register', label: 'Total books in accession register (as of today)',
      compute: function() { return { kind: 'value', value: ar_totalBooksInCatalogue_() }; } },
    { id: 'books_added_period', cat: 'Books / Accession Register', label: 'Books accessioned in selected period',
      compute: function(period) { var r = ar_booksAddedInPeriod_(period); return { kind: 'value', value: r.count, note: r.untracked ? (r.untracked + ' book(s) have no Accession Date on file, not included.') : '' }; } },
    { id: 'books_yearwise', cat: 'Books / Accession Register', label: 'Year-wise accession (last 10 years)',
      compute: function() { return Object.assign({ kind: 'table' }, ar_yearwiseAccessionTable_(10)); } },

    // -- Discarded Books --
    { id: 'discard_total', cat: 'Discarded Books', label: 'Total discarded books (all time)',
      compute: function() { return { kind: 'value', value: ar_totalDiscarded_() }; } },
    { id: 'discard_period', cat: 'Discarded Books', label: 'Discarded in selected period',
      compute: function(period) { return { kind: 'value', value: ar_discardedInPeriod_(period) }; } },
    { id: 'discard_yearwise', cat: 'Discarded Books', label: 'Year-wise discards, by staff (last 10 years)',
      compute: function() { return Object.assign({ kind: 'table' }, ar_yearwiseDiscardTable_(10)); } },

    // -- Circulation --
    { id: 'issued_period', cat: 'Circulation', label: 'Books issued in selected period',
      compute: function(period) { return { kind: 'value', value: ar_issuedInPeriod_(period) }; } },
    { id: 'issued_avg_per_day', cat: 'Circulation', label: 'Average books issued per day (selected period)',
      compute: function(period) { var r = ar_avgIssuedPerDay_(period); return { kind: 'value', value: r.avgPerDay === null ? 'Select a year or month to compute an average' : r.avgPerDay + ' per day (total ' + r.total + ')' }; } },
    { id: 'returned_period', cat: 'Circulation', label: 'Books returned in selected period',
      compute: function(period) { return { kind: 'value', value: ar_returnedInPeriod_(period) }; } },
    { id: 'pending_total', cat: 'Circulation', label: 'Total pending/overdue books (as of today)',
      compute: function() { return { kind: 'value', value: ar_pendingDueTotal_() }; } },
    { id: 'pending_period', cat: 'Circulation', label: 'Issued in selected period, still pending (as of today)',
      compute: function(period) { return { kind: 'value', value: ar_pendingDueInPeriod_(period) }; } },
    { id: 'fine_collected_period', cat: 'Circulation', label: 'Fine collected in selected period',
      compute: function(period) { return { kind: 'value', value: '\u20b9' + ar_fineCollectedInPeriod_(period) }; } },

    // -- Visitors / Gate --
    { id: 'visitors_period', cat: 'Visitors', label: 'Total visitors in selected period',
      compute: function(period) { return { kind: 'value', value: ar_visitorsInPeriod_(period) }; } },
    { id: 'visitors_avg_per_day', cat: 'Visitors', label: 'Average visitors per day (selected period)',
      compute: function(period) { var r = ar_avgVisitorsPerDay_(period); return { kind: 'value', value: r.avgPerDay === null ? 'Select a year or month to compute an average' : r.avgPerDay + ' per day (total ' + r.total + ')' }; } },
    { id: 'reference_readers_period', cat: 'Visitors', label: 'Reference readers in selected period (estimate)',
      compute: function(period) { return { kind: 'value', value: ar_referenceReadersInPeriod_(period), note: 'Estimated from visitor-stated purpose containing "read"/"reference" - not a separately tracked field.' }; } },

    // -- Newspapers & Magazines --
    { id: 'np_daily_total', cat: 'Newspapers & Magazines', label: 'Total daily newspaper copies received (selected period)',
      compute: function(period) { return { kind: 'value', value: ar_newspaperStats_(period).daily }; } },
    { id: 'np_weekly_total', cat: 'Newspapers & Magazines', label: 'Total weekly issues received (selected period)',
      compute: function(period) { return { kind: 'value', value: ar_newspaperStats_(period).weekly }; } },
    { id: 'np_qm_total', cat: 'Newspapers & Magazines', label: 'Total quarterly/monthly issues received (selected period)',
      compute: function(period) { var s = ar_newspaperStats_(period); return { kind: 'value', value: s.quarterly + s.monthly }; } },
    { id: 'np_kannada', cat: 'Newspapers & Magazines', label: 'Kannada-language total (selected period)',
      compute: function(period) { return { kind: 'value', value: ar_newspaperStats_(period).kannada }; } },
    { id: 'np_english', cat: 'Newspapers & Magazines', label: 'English-language total (selected period)',
      compute: function(period) { var s = ar_newspaperStats_(period); return { kind: 'value', value: s.english, note: s.unclassified ? (s.unclassified + ' copies are of papers not yet classified in NEWSPAPER_MASTER.') : '' }; } },
    { id: 'np_postal', cat: 'Newspapers & Magazines', label: 'Postal subscriptions total (selected period)',
      compute: function(period) { return { kind: 'value', value: ar_newspaperStats_(period).postal }; } },
    { id: 'np_local', cat: 'Newspapers & Magazines', label: 'Local-delivery total (selected period)',
      compute: function(period) { return { kind: 'value', value: ar_newspaperStats_(period).local }; } },

    // -- Bills & Finance --
    { id: 'bill_electric', cat: 'Bills & Finance', label: 'Electric bill (selected month)',
      compute: function(period) { var b = ar_billsForMonth_(period); return { kind: 'value', value: b && b.found ? ('\u20b9' + b.electric) : 'Select a specific month; not filled in yet for that month (MONTHLY_BILLS)' }; } },
    { id: 'bill_newspaper', cat: 'Bills & Finance', label: 'Newspaper bill (selected month)',
      compute: function(period) { var b = ar_billsForMonth_(period); return { kind: 'value', value: b && b.found ? ('\u20b9' + b.newspaper) : 'Select a specific month; not filled in yet for that month (MONTHLY_BILLS)' }; } },
    { id: 'bill_other', cat: 'Bills & Finance', label: 'Other bills (selected month)',
      compute: function(period) { var b = ar_billsForMonth_(period); return { kind: 'value', value: b && b.found ? ('\u20b9' + b.other) : 'Select a specific month; not filled in yet for that month (MONTHLY_BILLS)' }; } },
    { id: 'bill_old_paper_sale', cat: 'Bills & Finance', label: 'Old newspaper sale amount (selected month)',
      compute: function(period) { var b = ar_billsForMonth_(period); return { kind: 'value', value: b && b.found ? ('\u20b9' + b.oldPaperSale) : 'Select a specific month; not filled in yet for that month (MONTHLY_BILLS)' }; } }
  ].concat(AR_DOC_CATEGORIES.map(function(cat, i) {
    // -- Documents & Records -- one checklist entry per category, pulling
    // whatever's been uploaded to that category in the Document Library
    // below (photos embedded inline, everything else linked).
    return {
      id: 'doc_cat_' + i, cat: 'Documents & Records', label: cat,
      compute: function() { return ar_documentsForCategory_(cat); }
    };
  }));
}

function getAuthorityReportQuestions() {
  return AR_REGISTRY_().map(function(q) { return { id: q.id, cat: q.cat, label: q.label }; });
}

function getAuthorityReportAnswers(selectedIds, period) {
  var registry = AR_REGISTRY_();
  var byId = {};
  registry.forEach(function(q) { byId[q.id] = q; });

  var out = [];
  (selectedIds || []).forEach(function(id) {
    var q = byId[id];
    if (!q) return;
    var result;
    try {
      result = q.compute(period);
    } catch (e) {
      result = { kind: 'value', value: 'Error: ' + e.message };
    }
    out.push(Object.assign({ id: q.id, cat: q.cat, label: q.label }, result));
  });

  var init = getInitData().settings;
  return {
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy hh:mm a'),
    periodLabel: arPeriodLabel_(period),
    institution: init.institution || '',
    library: init.library || '',
    address: init.address || '',
    answers: out
  };
}
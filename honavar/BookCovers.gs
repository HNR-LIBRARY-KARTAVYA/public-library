// ================================================================
//  BookCovers.gs  -  NEW FILE (adds only; changes nothing old)
//  Gives the website pages the cover photo of books.
//  bclCovers(listOfAccessionNos) -> { accNo: driveFileId }  (only books that have a cover)
//  Reads only the column "Cover Image" of Books Catalogue. Writes nothing.
// ================================================================
function bclCovers(accs) {
  if (!Array.isArray(accs)) accs = [accs];
  accs = accs.slice(0, 60).map(function (a) { return String(a == null ? '' : a); });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Books Catalogue');
  if (!sh) return {};
  var head = sh.getRange(BOOKS_START_ROW - 1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = -1;
  for (var i = 0; i < head.length; i++) { if (String(head[i]).trim().toLowerCase() === 'cover image') { col = i + 1; break; } }
  var last = sh.getLastRow();
  if (col < 0 || last < BOOKS_START_ROW) return {};
  var vals = sh.getRange(BOOKS_START_ROW, col, last - BOOKS_START_ROW + 1, 1).getValues();
  var byId = lmCatalogue_().byId, out = {};
  accs.forEach(function (a) {
    var b = byId[normId_(a)];
    if (!b) return;
    var v = String((vals[b.row - BOOKS_START_ROW] || [''])[0] || '');
    var m = v.match(/\/d\/([-\w]{15,})/) || v.match(/[?&]id=([-\w]{15,})/);
    if (m) out[a] = m[1];
  });
  return out;
}

/** Run once from the editor: should show how many of the tried books have a cover. */
function testBclCovers() {
  var r = bclCovers(['1', '2', '3', '1818', '40342']);
  console.log('Covers found: ' + JSON.stringify(r));
}

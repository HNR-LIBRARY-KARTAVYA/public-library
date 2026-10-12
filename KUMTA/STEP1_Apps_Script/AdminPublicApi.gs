// ================================================================
// ADMIN PUBLIC API - lets pages hosted on GitHub Pages (Gate Check-In
// and My Library) reach this admin project without Google sign-in.
// Same file for BCL Kumta and BCL Honavar admin projects.
// ONLY the functions in this list can be called from outside.
// Member-account functions check the member's login token themselves.
// ================================================================
var ADMIN_PUBLIC_API_ALLOWED_ = [
  // Gate Check-In page
  'searchPeopleByName', 'gateCheckInOut',
  // My Library page (LibraryModules.gs)
  'mpRequestCode', 'mpVerify', 'mpLogout', 'mpMyAccount', 'mpRenew',
  'mpReserve', 'mpCancelReservation', 'mpSuggest', 'mpSearch', 'mpLists'
];

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = String(req.fn || '');
    if (ADMIN_PUBLIC_API_ALLOWED_.indexOf(fn) < 0) throw new Error('Not allowed: ' + fn);
    var f = globalThis[fn];
    if (typeof f !== 'function') throw new Error('Function missing: ' + fn);
    var args = Array.isArray(req.args) ? req.args.slice(0, 4) : [];
    out = { ok: true, result: f.apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor: should log {"ok":true,...} */
function testAdminPublicApi() {
  var r = doPost({ postData: { contents: JSON.stringify({ fn: 'mpSearch', args: ['a'] }) } });
  console.log(r.getContent().slice(0, 300));
}

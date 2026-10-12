// ================================================================
// AUTO THANK-YOU REPLY - every new e-mail that arrives in this
// library Gmail gets one short automatic reply with the website link.
// Same file for Kumta and Honavar (uses MM_ from MemberMail.gs).
//
// IMPORTANT: it replies from the Google account that owns this
// script, so the admin project must belong to the library Gmail
// (dclkarwarbranch...@gmail.com).
//
// Switch ON : run installAutoReply  (checks every 10 minutes)
// Switch OFF: run removeAutoReply
// ================================================================

var AR_LABEL_ = 'Library-AutoReplied';
var AR_SKIP_FROM_ = /(no-?reply|mailer-daemon|postmaster|notification|notify|alerts?@|bounce|newsletter|calendar-|google\.com|accounts\.google|facebook|linkedin|amazon|flipkart|bank|otp)/i;

function autoReplyText_(name) {
  return 'Dear ' + (name || 'Reader') + ',\n\n' +
    'Thank you for writing to ' + MM_.library + ', ' + MM_.branch + '. We have received your message, and our staff will reply within 2 working days.\n\n' +
    'Meanwhile, you can search our books and check whether a book is Available or Issued on our website (no sign-in needed):\n' + MM_.website + '\n\n' +
    'Library timings: ' + MM_.timings + '\n\n' +
    'This is an automatic reply - you do not need to answer it.\n\n' +
    'Thank you,\n' + MM_.library + '\n' + MM_.branch;
}

/** Runs every 10 minutes (after installAutoReply). */
function autoReplyRun() {
  var label = GmailApp.getUserLabelByName(AR_LABEL_) || GmailApp.createLabel(AR_LABEL_);
  var me = Session.getEffectiveUser().getEmail().toLowerCase();
  var cache = CacheService.getScriptCache();
  var threads = GmailApp.search('in:inbox -label:' + AR_LABEL_ + ' newer_than:1d -from:me -category:promotions -category:social -category:updates -category:forums', 0, 20);
  var sent = 0;
  threads.forEach(function (th) {
    try {
      var msgs = th.getMessages();
      var first = msgs[0], fromRaw = first.getFrom();
      var m = fromRaw.match(/<([^>]+)>/); var from = (m ? m[1] : fromRaw).trim().toLowerCase();
      var name = (fromRaw.split('<')[0] || '').replace(/"/g, '').trim();
      var weStarted = from === me;
      var skip = weStarted || AR_SKIP_FROM_.test(from) || msgs.length > 1 ||
                 /auto-?reply|out of office|undeliverable|delivery status/i.test(first.getSubject()) ||
                 cache.get('ar_' + from);                       // max one auto reply per sender per 6 hours
      try { if (!skip && /auto-(generated|replied)/i.test(first.getHeader('Auto-Submitted') || '')) skip = true; } catch (e) {}
      if (!skip) {
        th.reply(autoReplyText_(name), { name: MM_.shortName, replyTo: MM_.email });
        cache.put('ar_' + from, '1', 21600);
        logEmail_('Reply', 'Automatic thank-you', '', name, from, first.getSubject(), 'Sent');
        sent++;
      }
      th.addLabel(label);        // never looked at twice; the thread stays UNREAD for staff
    } catch (e) { console.log('autoReplyRun: ' + e); }
  });
  return sent;
}

function installAutoReply() {
  removeAutoReply();
  ScriptApp.newTrigger('autoReplyRun').timeBased().everyMinutes(10).create();
  // mark what is already in the inbox so old e-mails do NOT get a reply now
  var label = GmailApp.getUserLabelByName(AR_LABEL_) || GmailApp.createLabel(AR_LABEL_);
  GmailApp.search('in:inbox -label:' + AR_LABEL_, 0, 100).forEach(function (t) { t.addLabel(label); });
  console.log('Automatic thank-you reply is ON for ' + Session.getEffectiveUser().getEmail() + ' (checks every 10 minutes).');
}

function removeAutoReply() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'autoReplyRun') ScriptApp.deleteTrigger(t); });
  console.log('Automatic thank-you reply is OFF.');
}

/** Sends you a sample of the automatic reply. */
function testAutoReply() {
  var me = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to: me, subject: '[SAMPLE] Automatic reply - ' + MM_.shortName, body: autoReplyText_('Sample Reader'), name: MM_.shortName });
  console.log('Sample sent to ' + me);
}

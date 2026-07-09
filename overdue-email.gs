/***** KelMar Tool Tracker — Daily Overdue-Tool Email *****/
/* Paste this whole file into a new Google Apps Script project (script.google.com),
   then run setupDailyTrigger() once to schedule it. See setup steps from Claude. */

// ===== SETTINGS — tweak these if you like =====
var OVERDUE_AFTER_DAYS = 2;                     // list tools out longer than this many days (set 0 to list everything that's out)
var EMAIL_TO          = 'mitch@kelmarprojects.com.au';   // who gets the email
var SEND_HOUR         = 6;                       // hour of the morning to send (24h, Melbourne time)
var SEND_WHEN_NONE    = false;                   // true = send an "all good" email even when nothing's overdue
// ==============================================

// (leave the rest as-is)
var API_KEY = 'AIzaSyAjnuthiLHGzDov4v-sMwVgp0S0N568TeU';
var PROJECT = 'kelmar-tool-tracker';
var BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents';

function sendOverdueEmail() {
  var token  = anonToken_();
  var people = fetchCollection_('people', token);
  var tools  = fetchCollection_('tools', token);
  var nameById = {};
  people.forEach(function (p) { nameById[p.id] = p.name; });

  var now = new Date();
  var overdue = tools.filter(function (t) {
    if (t.status !== 'out') return false;
    if (t.dueDate) return now.getTime() > new Date(t.dueDate).getTime() + 86400000;   // past its return-by date
    return t.since && daysBetween_(new Date(t.since), now) >= OVERDUE_AFTER_DAYS;      // or out longer than the fallback
  }).sort(function (a, b) { return new Date(a.since) - new Date(b.since); });

  if (overdue.length === 0 && !SEND_WHEN_NONE) return;

  var rows = overdue.map(function (t) {
    var who = nameById[t.holder] || 'someone';
    var days = t.since ? daysBetween_(new Date(t.since), now) : 0;
    return '<tr>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;font-weight:600;">' + esc_(t.name) + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;">' + esc_(t.code || '') + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;">' + esc_(who) + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;">' + esc_(t.job || '—') + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;color:#c62828;font-weight:700;">' + days + ' day' + (days === 1 ? '' : 's') + '</td>' +
    '</tr>';
  }).join('');

  var body = '<div style="font-family:Arial,Helvetica,sans-serif;color:#1A1A1A;max-width:620px;">' +
    '<div style="background:#1A1A1A;padding:16px;border-bottom:4px solid #F5A800;">' +
      '<span style="color:#fff;font-size:20px;font-weight:800;">KelMar</span> ' +
      '<span style="color:#fff;font-size:20px;font-weight:300;">Projects</span>' +
      '<div style="color:#F5A800;font-size:11px;letter-spacing:2px;margin-top:4px;">TOOL TRACKER &middot; OVERDUE TOOLS</div>' +
    '</div>' +
    '<div style="padding:18px 16px;">' +
      (overdue.length
        ? '<p style="margin:0 0 12px;">Morning Mitch — these tools have been out more than ' + OVERDUE_AFTER_DAYS + ' day' + (OVERDUE_AFTER_DAYS === 1 ? '' : 's') + ':</p>' +
          '<table style="border-collapse:collapse;width:100%;font-size:14px;">' +
            '<tr style="background:#555;color:#fff;">' +
              '<th style="text-align:left;padding:8px 10px;">Tool</th>' +
              '<th style="text-align:left;padding:8px 10px;">ID</th>' +
              '<th style="text-align:left;padding:8px 10px;">Who has it</th>' +
              '<th style="text-align:left;padding:8px 10px;">Job</th>' +
              '<th style="text-align:left;padding:8px 10px;">Out for</th>' +
            '</tr>' + rows +
          '</table>'
        : '<p style="margin:0;">Morning Mitch — nothing overdue today. All tools accounted for. &#9989;</p>'
      ) +
      '<p style="margin:16px 0 0;font-size:13px;">Open the tracker (admin): ' +
        '<a href="https://mitch492.github.io/kelmar-tool-tracker/#admin" style="color:#F5A800;font-weight:700;">KelMar Tool Tracker</a></p>' +
    '</div>' +
    '<div style="padding:12px 16px;color:#999;font-size:11px;border-top:1px solid #eee;">KelMar Projects &middot; Tool Tracker &middot; sent automatically each morning</div>' +
  '</div>';

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: 'KelMar Tools — ' + (overdue.length ? overdue.length + ' overdue' : 'all accounted for'),
    htmlBody: body
  });
}

// ===== Run this ONCE to schedule the daily email =====
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendOverdueEmail') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendOverdueEmail')
    .timeBased().everyDays(1).atHour(SEND_HOUR).inTimezone('Australia/Melbourne')
    .create();
}

// ===== helpers (leave alone) =====
function anonToken_() {
  var res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:signUp?key=' + API_KEY, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify({ returnSecureToken: true })
  });
  return JSON.parse(res.getContentText()).idToken;
}
function fetchCollection_(name, token) {
  var out = [], pageToken = '';
  do {
    var url = BASE + '/' + name + '?pageSize=300' + (pageToken ? '&pageToken=' + pageToken : '');
    var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var data = JSON.parse(res.getContentText() || '{}');
    (data.documents || []).forEach(function (d) {
      var o = { id: d.name.split('/').pop() }, f = d.fields || {};
      for (var k in f) o[k] = valueOf_(f[k]);
      out.push(o);
    });
    pageToken = data.nextPageToken || '';
  } while (pageToken);
  return out;
}
function valueOf_(v) {
  if (v.stringValue  !== undefined) return v.stringValue;
  if (v.integerValue !== undefined) return Number(v.integerValue);
  if (v.doubleValue  !== undefined) return Number(v.doubleValue);
  if (v.booleanValue !== undefined) return v.booleanValue;
  return null;
}
function daysBetween_(a, b) { return Math.floor((b - a) / 86400000); }
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

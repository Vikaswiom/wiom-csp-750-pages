/**
 * Wiom — Project Dominance ₹750
 * Callback request log: "मुझे कॉल करके समझाएं" → one row per CSP, for the team to call.
 *
 * Reads exactly what the live pages already send, so switching to this script is a
 * one-line URL change in flow1/2/3, qa-flow1/2/3 and faq.html — no page edits:
 *     ?flow=P750CALL&screen=<page>_<lang>&uid=<cspId>&t=<ms>
 *
 * SETUP (2 minutes)
 *  1. Create a Google Sheet, name it e.g. "CSP 750 — Callback Requests".
 *  2. Extensions → Apps Script. Delete the stub, paste this file, Save.
 *  3. Deploy → New deployment → type "Web app".
 *       Execute as:      Me
 *       Who has access:  Anyone            <-- must be "Anyone", not "Anyone with Google account"
 *  4. Authorise (it will warn "Google hasn't verified this app" → Advanced → Go to project → Allow).
 *  5. Copy the /exec URL and send it to Claude, or paste it into P750_CALL_LOG in the 7 files.
 *
 * NOTE: a Google Workspace account (…@wiom.in) usually blocks "Anyone" access, which is why
 * the earlier tracking script was deployed from a personal Gmail. Deploy this from an account
 * where "Anyone" is selectable, or the beacon will be silently rejected from the CSP's phone.
 */

var SHEET_NAME  = 'Callbacks';
var DEDUP_HOURS = 6;     /* a CSP tapping twice in this window updates the row, not adds one */
var TZ          = 'Asia/Kolkata';

function doGet(e) {
  var p = (e && e.parameter) || {};

  if (p.action === 'json') return json_(summary_(), p.callback);
  if (p.action === 'ping') return json_({ ok: true, sheet: SHEET_NAME }, p.callback);

  var csp  = String(p.uid || p.csp || p.csp_id || '').trim().substring(0, 80);
  var scr  = String(p.screen || p.page || '').trim().substring(0, 40);   /* e.g. flow1_hi */
  var page_ = scr.split('_')[0] || '';
  var lang_ = scr.split('_')[1] || '';
  var flow = String(p.flow || '').trim().substring(0, 20);
  var src  = String(p.src || '').trim().substring(0, 60);   /* which CTA / question */

  if (csp) logCallback_(csp, page_, lang_, flow, src);
  return gif_();          /* the page calls this via new Image(), so answer with a pixel */
}

function logCallback_(csp, page, lang, flow, src) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (err) { /* proceed anyway */ }
  try {
    var sh   = sheet_();
    var now  = new Date();
    var row = findRecent_(sh, csp);
    if (row) {
      sh.getRange(row, 8).setValue(Number(sh.getRange(row, 8).getValue() || 1) + 1);
      sh.getRange(row, 9).setValue(fmt_(now, 'yyyy-MM-dd HH:mm'));
      /* keep every place he asked from, so the caller has the context */
      var seen = String(sh.getRange(row, 11).getValue() || '');
      if (src && seen.indexOf(src) < 0) {
        sh.getRange(row, 11).setValue(seen ? seen + ', ' + src : src);
      }
      return;
    }
    sh.appendRow([
      fmt_(now, 'yyyy-MM-dd'),        /* A date        */
      fmt_(now, 'HH:mm'),             /* B time (IST)  */
      csp,                            /* C csp_id      */
      page,                           /* D page        */
      lang,                           /* E language    */
      flow,                           /* F source flow */
      'Pending',                      /* G status      */
      1,                              /* H requests    */
      fmt_(now, 'yyyy-MM-dd HH:mm'),  /* I last request*/
      '',                             /* J notes       */
      src                             /* K asked from  */
    ]);
    var last = sh.getLastRow();
    sh.getRange(last, 7).setDataValidation(statusRule_());
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

/* Same CSP inside the dedup window → return its row number, else 0. */
function findRecent_(sh, csp) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var n    = Math.min(last - 1, 500);               /* only scan the recent tail */
  var from = last - n + 1;
  var vals = sh.getRange(from, 1, n, 9).getValues();
  var cut  = new Date().getTime() - DEDUP_HOURS * 3600 * 1000;
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][2]) !== csp) continue;
    var when = new Date(String(vals[i][0]) + ' ' + String(vals[i][1]) + ':00');
    if (isNaN(when.getTime()) || when.getTime() >= cut) return from + i;
    return 0;
  }
  return 0;
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['date','time (IST)','csp_id','page','lang','flow','status','requests','last request','notes','asked from']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 11).setFontWeight('bold');
    sh.setColumnWidth(3, 120);
    sh.setColumnWidth(10, 280);
  }
  return sh;
}

function statusRule_() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Called', 'Done', 'Unreachable'], true)
    .setAllowInvalid(true).build();
}

/* Counts for a dashboard or a quick check: ?action=json */
function summary_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var out = { total: 0, pending: 0, byDate: {}, byPage: {}, updated: new Date().toISOString() };
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < v.length; i++) {
    out.total++;
    if (String(v[i][6]) === 'Pending') out.pending++;
    var d = String(v[i][0]); out.byDate[d] = (out.byDate[d] || 0) + 1;
    var pg = String(v[i][3]); out.byPage[pg] = (out.byPage[pg] || 0) + 1;
  }
  return out;
}

function fmt_(d, pattern) { return Utilities.formatDate(d, TZ, pattern); }

function json_(obj, cb) {
  var s = JSON.stringify(obj);
  if (cb) return ContentService.createTextOutput(cb + '(' + s + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

/* 1x1 transparent GIF — keeps the <img> beacon from logging an error on the CSP's phone. */
function gif_() {
  return ContentService
    .createTextOutput(Utilities.newBlob(Utilities.base64Decode(
      'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'
    )).getDataAsString())
    .setMimeType(ContentService.MimeType.TEXT);
}

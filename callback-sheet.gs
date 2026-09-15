/**
 * Wiom — CSP ₹750 · callback request log
 *
 * Every "मुझे कॉल करके समझाएं" tap (offer page AND FAQ) sends the SAME event here:
 *     ?flow=P750CALL&screen=<page>_<lang>&uid=<cspId>
 * One row per CSP per day — repeat taps bump the count instead of adding rows.
 *
 * PASTE ALL OF THIS into Code.gs (select everything first, then paste), Save, then
 * Deploy → Manage deployments → ✏️ → Version: New version → Deploy.
 * Access must be "Anyone". The last line of this file is the closing brace of doGet.
 */
function doGet(e) {
  var p   = (e && e.parameter) || {};
  var csp = String(p.uid || p.csp || p.csp_id || '').trim().substring(0, 80);
  if (!csp) return ContentService.createTextOutput('no-csp');

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName('Callbacks');
  if (!sh) {
    sh = ss.insertSheet('Callbacks');
    sh.appendRow(['date', 'time (IST)', 'csp_id', 'page', 'lang', 'status', 'requests', 'notes']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  var now  = new Date();
  var date = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
  var time = Utilities.formatDate(now, 'Asia/Kolkata', 'HH:mm');
  var scr  = String(p.screen || '').split('_');

  /* same CSP already asked today -> bump the counter, keep one row */
  var last = sh.getLastRow();
  if (last > 1) {
    var n    = Math.min(last - 1, 300);
    var from = last - n + 1;
    var vals = sh.getRange(from, 1, n, 3).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      /* Sheets turns the date cell into a Date object, so normalise before comparing */
      var d0 = vals[i][0];
      var dStr = (d0 instanceof Date) ? Utilities.formatDate(d0, 'Asia/Kolkata', 'yyyy-MM-dd') : String(d0);
      if (String(vals[i][2]) === csp && dStr === date) {
        var row = from + i;
        sh.getRange(row, 7).setValue(Number(sh.getRange(row, 7).getValue() || 1) + 1);
        sh.getRange(row, 2).setValue(time);
        return ContentService.createTextOutput('ok-dup');
      }
    }
  }

  sh.appendRow([date, time, csp, scr[0] || '', scr[1] || '', 'Pending', 1, '']);
  return ContentService.createTextOutput('ok');
}

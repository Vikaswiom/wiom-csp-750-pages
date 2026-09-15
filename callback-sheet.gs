/**
 * Wiom — CSP ₹750 · callback + visit log
 *
 * Two kinds of event, two tabs, one script:
 *   flow=P750CALL  -> "Callbacks" tab  (a CSP asked to be called)
 *   flow=P750VIEW  -> "Visits"    tab  (a CSP opened flow1 / faq)
 *
 * Both arrive as:  ?flow=<...>&screen=<page>_<lang>&uid=<cspId>
 *
 * Callbacks dedup per CSP per day (one call row however many CTAs he taps).
 * Visits dedup per CSP per page per day (reloads bump "opens").
 *
 * PASTE ALL OF THIS into Code.gs (select everything first), Save, then
 * Deploy -> Manage deployments -> pencil -> Version: New version -> Deploy.
 * Access must be "Anyone". The last line of this file is the closing brace of doGet.
 */
function doGet(e) {
  var p      = (e && e.parameter) || {};
  var isView = String(p.flow || '').trim().toUpperCase() === 'P750VIEW';
  /* visits arrive as vid= so that an older deployment of this script (which only
     knows uid=) ignores them instead of dumping page views into Callbacks */
  var csp    = String((isView ? p.vid : p.uid) || p.uid || p.csp || p.csp_id || '').trim().substring(0, 80);
  if (!csp) return ContentService.createTextOutput('no-csp');
  var name   = isView ? 'Visits' : 'Callbacks';
  var header = isView
    ? ['date', 'time (IST)', 'csp_id', 'page', 'lang', 'opens']
    : ['date', 'time (IST)', 'csp_id', 'page', 'lang', 'status', 'requests', 'notes'];
  var countCol = isView ? 6 : 7;

  header = header.concat(['key']);          /* last column: dedup key, plain text */
  var keyCol = header.length;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }

  var now  = new Date();
  var date = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
  var time = Utilities.formatDate(now, 'Asia/Kolkata', 'HH:mm');
  var scr  = String(p.screen || '').split('_');
  var page = scr[0] || '';
  var lang = scr[1] || '';

  /* Dedup on a key the script writes itself. Never compare the date CELL: Sheets
     turns it into a Date in the spreadsheet's own timezone, and re-formatting that
     in IST can land on the previous day — which is why dedup silently never matched. */
  var key  = 'k' + date + '|' + csp + (isView ? '|' + page : '');
  var last = sh.getLastRow();
  if (last > 1) {
    var n    = Math.min(last - 1, 400);
    var from = last - n + 1;
    var keys = sh.getRange(from, keyCol, n, 1).getValues();
    for (var i = keys.length - 1; i >= 0; i--) {
      if (String(keys[i][0]) === key) {
        var row = from + i;
        sh.getRange(row, countCol).setValue(Number(sh.getRange(row, countCol).getValue() || 1) + 1);
        sh.getRange(row, 2).setValue(time);
        return ContentService.createTextOutput('ok-dup');
      }
    }
  }

  sh.appendRow(isView
    ? [date, time, csp, page, lang, 1, key]
    : [date, time, csp, page, lang, 'Pending', 1, '', key]);
  return ContentService.createTextOutput('ok');
}

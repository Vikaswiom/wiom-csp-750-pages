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

  /* ── Dashboard feed: aggregates only, never a csp_id ──────────────────
     ?action=stats            -> JSON
     ?action=stats&callback=f -> JSONP (the dashboard uses this; Apps Script
                                 redirects break a plain cross-origin fetch) */
  if (String(p.action || '') === 'stats') {
    var ssx   = SpreadsheetApp.getActiveSpreadsheet();
    var today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
    var dstr  = function (v) {
      return (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd') : String(v);
    };
    var hstr  = function (v) {
      if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Kolkata', 'HH');
      var s = String(v); return s.length >= 2 ? s.substring(0, 2) : '';
    };
    var days = {}, hours = {}, uniq = { flow: {}, faq: {}, call: {} };
    var day  = function (d) {
      if (!days[d]) days[d] = { d: d, flow: {}, faq: {}, calls: {} };
      return days[d];
    };
    var hour = function (h) {
      if (!hours[h]) hours[h] = { h: h, flow: {}, faq: {}, calls: {} };
      return hours[h];
    };

    var vsh = ssx.getSheetByName('Visits');
    if (vsh && vsh.getLastRow() > 1) {
      var vv = vsh.getRange(2, 1, vsh.getLastRow() - 1, 4).getValues();
      for (var a = 0; a < vv.length; a++) {
        var vd = dstr(vv[a][0]), vc = String(vv[a][2]), vp = String(vv[a][3]);
        if (!vc) continue;
        var key = (vp === 'faq') ? 'faq' : 'flow';
        uniq[key][vc] = 1;
        day(vd)[key === 'faq' ? 'faq' : 'flow'][vc] = 1;
        if (vd === today) hour(hstr(vv[a][1]))[key === 'faq' ? 'faq' : 'flow'][vc] = 1;
      }
    }
    var csh = ssx.getSheetByName('Callbacks');
    var pending = 0;
    if (csh && csh.getLastRow() > 1) {
      var cv = csh.getRange(2, 1, csh.getLastRow() - 1, 6).getValues();
      for (var b = 0; b < cv.length; b++) {
        var cd = dstr(cv[b][0]), cc = String(cv[b][2]);
        if (!cc) continue;
        uniq.call[cc] = 1;
        day(cd).calls[cc] = 1;
        if (cd === today) hour(hstr(cv[b][1])).calls[cc] = 1;
        if (String(cv[b][5]) === 'Pending') pending++;
      }
    }

    var n = function (o) { var c = 0, k; for (k in o) if (o.hasOwnProperty(k)) c++; return c; };
    var dayList = [], hourList = [], kk;
    for (kk in days) if (days.hasOwnProperty(kk)) {
      dayList.push({ d: days[kk].d, flow: n(days[kk].flow), faq: n(days[kk].faq), calls: n(days[kk].calls) });
    }
    dayList.sort(function (x, y) { return x.d < y.d ? -1 : 1; });
    for (kk in hours) if (hours.hasOwnProperty(kk)) {
      hourList.push({ h: hours[kk].h, flow: n(hours[kk].flow), faq: n(hours[kk].faq), calls: n(hours[kk].calls) });
    }
    hourList.sort(function (x, y) { return x.h < y.h ? -1 : 1; });

    var stats = {
      updated: new Date().toISOString(),
      today: today,
      totals: { flow: n(uniq.flow), faq: n(uniq.faq), calls: n(uniq.call), pending: pending },
      days: dayList,
      hours: hourList
    };
    var body = JSON.stringify(stats);
    if (p.callback) {
      return ContentService.createTextOutput(p.callback + '(' + body + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
  }

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

  header = header.concat(['key', 'last_t']);   /* dedup key + the beacon's own timestamp */
  var keyCol = header.length - 1;
  var tCol   = header.length;
  var stamp  = String(p.t || '');

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
    var keys = sh.getRange(from, keyCol, n, 2).getValues();   /* key + last_t */
    for (var i = keys.length - 1; i >= 0; i--) {
      if (String(keys[i][0]) === key) {
        var row = from + i;
        /* The page fires two channels (no-cors fetch + Image) with the SAME t, so the
           second arrival is the same tap, not a second one. Same t = duplicate: skip it. */
        if (stamp && String(keys[i][1]) === stamp) {
          return ContentService.createTextOutput('ok-same-t');
        }
        sh.getRange(row, countCol).setValue(Number(sh.getRange(row, countCol).getValue() || 1) + 1);
        sh.getRange(row, 2).setValue(time);
        if (stamp) sh.getRange(row, tCol).setValue(stamp);
        return ContentService.createTextOutput('ok-dup');
      }
    }
  }

  sh.appendRow(isView
    ? [date, time, csp, page, lang, 1, key, stamp]
    : [date, time, csp, page, lang, 'Pending', 1, '', key, stamp]);
  return ContentService.createTextOutput('ok');
}

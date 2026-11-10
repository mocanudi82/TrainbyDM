// Train log - receives the app's data and keeps these tabs up to date:
// Days (main tab: one line per day, colour-coded), Weekly (summary), Log (every exercise), Backup.
// The app always sends the whole log, so sending twice never duplicates anything.

var DAYS_HEAD = ['Date', 'Day', 'Training type', 'Week / Day', 'Cardio type', 'Cardio mode', 'Cardio min', 'Mobility', 'Tango', 'Trip', 'Readiness', 'Minutes', 'Comment'];
var NCOL = 13;
var DEFAULT_LISTS = {
  'Training type': ['Ski 5 Week', 'Ski Training 8 Weeks', 'Strength', 'Athletic 3+3', 'Conditioning', 'Other sport', 'Rest'],
  'Cardio type': ['Zone 2', 'VO\u2082', 'Threshold'],
  'Cardio mode': ['Swim', 'Run', 'Bike indoor', 'Bike outdoor', 'Row', 'Hike', 'Other'],
  'Mobility': ['Morning', 'Evening', 'Morning + Evening', 'Knee & feet'],
  'Readiness': ['Low', 'OK', 'Great']
};
var COLORS = {
  'Ski Training': '#cfe8f5', 'Strength': '#f8d0c6', 'Athletic': '#c9e9e8', 'Conditioning': '#fbe3b8',
  'Cardio': '#d3ecc8', 'Mobility': '#e1d8f0', 'Tango': '#f6cfe0', 'Trip': '#f3e6a8'
};

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'backup') {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Backup');
    if (!sh || sh.getLastRow() < 3) return out({ ok: false, error: 'The sheet has no backup yet' });
    var at = String(sh.getRange(1, 2).getValue());
    var parts = sh.getRange(3, 1, sh.getLastRow() - 2, 1).getValues().map(function (r) { return r[0]; });
    return out({ ok: true, at: at, json: parts.join('') });
  }
  return out({ ok: true, hello: 'Train log is ready' });
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d || d.v < 4) return out({ ok: false, error: 'Old app version: update the app first' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureLists(ss);
    var days = mergeDays(ss, d.days);                            // app data fills cells; what you typed in empty cells is kept
    put(ss, 'Days', DAYS_HEAD, days);
    var wk = weekly(d, days);
    put(ss, 'Weekly', ['Week starting (Mon)', 'Sessions', 'Sets done', 'Minutes', 'Cardio min', 'Tango days'], wk);
    put(ss, 'Log', ['Date', 'Program', 'Week', 'Session', 'Exercise', 'Planned', 'Sets done', 'Sets planned', 'kg', 'Minutes', 'Pain', 'Notes', 'Readiness'], d.log);
    formatDays(ss.getSheetByName('Days'), days.length);
    styleTable(ss.getSheetByName('Days'), NCOL, days.length, [1, 2, 4, 7, 9, 11, 12], 'dd mmm yyyy', '#ec4624');
    styleTable(ss.getSheetByName('Weekly'), 6, wk.length, [1, 2, 3, 4, 5, 6], 'dd mmm yyyy', '#4fbbbc');
    styleTable(ss.getSheetByName('Log'), 13, d.log.length, [1, 3, 7, 8, 9, 10, 11, 13], 'dd mmm yyyy', '#999999');
    mondayLines(ss.getSheetByName('Days'), days);
    addDropdowns(ss, ss.getSheetByName('Days'));
    if (d.backup) saveBackup(ss, d.backup);
    var old = ss.getSheetByName('Trips');
    if (old && ss.getSheets().length > 1) ss.deleteSheet(old);   // trips now live in the Days tab
    var junk = ss.getSheetByName('Sheet1');
    if (junk && ss.getSheets().length > 1 && junk.getLastRow() === 0) ss.deleteSheet(junk);
    ss.setActiveSheet(ss.getSheetByName('Days'));
    ss.moveActiveSheet(1);                                       // Days is always the first tab
    return out({ ok: true, rows: d.log.length });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function formatDays(sh, n) {
  var last = Math.max(n + 1, 400), rules = [];     // rules reach row 400 so rows you add by hand are coloured too
  function rule(range, formula, color) {
    return SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(range)])
      .whenFormulaSatisfied(formula).setBackground(color).build();
  }
  rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange('A2:M' + last)])
    .whenFormulaSatisfied('=AND($A2<>"",$C2="",$E2="",$F2="",$H2="",$I2="",$J2="")').setFontColor('#999999').build());   // rest day
  ['Ski', 'Strength', 'Athletic', 'Conditioning'].forEach(function (k) {
    rules.push(rule('C2:D' + last, '=ISNUMBER(SEARCH("' + k + '",$C2))', COLORS[k === 'Ski' ? 'Ski Training' : k]));
  });
  rules.push(rule('E2:G' + last, '=OR($E2<>"",$F2<>"",$G2<>"")', COLORS['Cardio']));
  rules.push(rule('H2:H' + last, '=$H2<>""', COLORS['Mobility']));
  rules.push(rule('I2:I' + last, '=$I2="Yes"', COLORS['Tango']));
  rules.push(rule('J2:J' + last, '=$J2<>""', COLORS['Trip']));
  rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange('K2:K' + last)])
    .whenTextEqualTo('Low').setFontColor('#c0392b').build());
  sh.setConditionalFormatRules(rules);
  var w = [100, 50, 130, 90, 110, 110, 80, 150, 60, 190, 80, 70, 240];
  w.forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  sh.getRange('G2:G' + last).setNumberFormat('0');
  sh.getRange('O1').setValue('Colour code').setFontWeight('bold');
  var i = 2;
  Object.keys(COLORS).forEach(function (k) { sh.getRange(i++, 15).setValue(k).setBackground(COLORS[k]); });
  sh.getRange(i, 15).setValue('Rest day (grey text)').setFontColor('#999999');
  sh.getRange(i + 2, 15).setValue('Dropdowns: Training type, Cardio type, Cardio mode, Mobility, Readiness. Edit the choices in the Lists tab.').setFontStyle('italic').setFontColor('#666666');
  sh.setColumnWidth(15, 150);
}

// The Lists tab holds the dropdown choices. It is created once and never overwritten, so you can edit it.
function ensureLists(ss) {
  var sh = ss.getSheetByName('Lists');
  if (sh && sh.getRange(1, 3).getValue() === 'Cardio mode') return sh;   // current layout: leave it alone, it may hold your edits
  if (sh) sh.clear(); else sh = ss.insertSheet('Lists');                  // first run or old layout: write the defaults
  var keys = Object.keys(DEFAULT_LISTS);
  sh.getRange(1, 1, 1, keys.length).setValues([keys]).setFontWeight('bold').setBackground('#222a2a').setFontColor('#e2dac2');
  keys.forEach(function (k, c) {
    var v = DEFAULT_LISTS[k].map(function (x) { return [x]; });
    sh.getRange(2, c + 1, v.length, 1).setValues(v);
    sh.setColumnWidth(c + 1, 140);
  });
  sh.setTabColor('#e2dac2');
  return sh;
}

function addDropdowns(ss, sh) {
  var lists = ss.getSheetByName('Lists'), last = 400;
  [['C', 'A'], ['E', 'B'], ['F', 'C'], ['H', 'D'], ['K', 'E']].forEach(function (p) {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(lists.getRange(p[1] + '2:' + p[1] + '30'), true)
      .setAllowInvalid(true)                                    // combined values such as "Strength + Conditioning" are allowed
      .build();
    sh.getRange(p[0] + '2:' + p[0] + last).setDataValidation(rule);
  });
  sh.getRange('G2:G' + last).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 600).setAllowInvalid(false).setHelpText('Minutes, 0-600').build());
  sh.getRange('I2:I' + last).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes'], true).setAllowInvalid(true).build());
}

// Keep what you typed. The app fills a cell only when it has something for it; empty app cells never erase your entries.
function mergeDays(ss, appRows) {
  var tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName('Days'), old = {};
  function key(v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || ''); }
  if (sh && sh.getLastRow() > 1) {
    var oldCols = sh.getLastColumn() >= NCOL ? NCOL : 12;     // 12 = the earlier layout without Cardio mode
    sh.getRange(2, 1, sh.getLastRow() - 1, oldCols).getValues().forEach(function (r) {
      var k = key(r[0]);
      if (!k) return;
      if (oldCols === 12) r = r.slice(0, 5).concat(['']).concat(r.slice(5));
      old[k] = r;
    });
  }
  var out = {}, k;
  Object.keys(old).forEach(function (k) { out[k] = old[k].map(function (v, i) { return i === 0 ? k : v; }); });
  appRows.forEach(function (a) {
    var m = out[a[0]] ? out[a[0]].slice() : [];
    for (var c = 0; c < NCOL; c++) {
      var has = a[c] !== '' && a[c] != null;
      m[c] = (c < 2 || has) ? a[c] : (m[c] == null ? '' : m[c]);
    }
    out[a[0]] = m;
  });
  return Object.keys(out).sort().map(function (k) { return out[k]; });
}

// Same look on every tab: dark header with an orange rule, banded rows, thin lines, filter, frozen header.
function styleTable(sh, ncols, nrows, centerCols, dateFmt, tabColor) {
  var n = Math.max(nrows, 1), all = sh.getRange(1, 1, n + 1, ncols);
  sh.getBandings().forEach(function (b) { b.remove(); });
  all.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  var band = all.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#222a2a').setFirstRowColor('#ffffff').setSecondRowColor('#f7f5ee');
  sh.getRange(1, 1, 1, ncols).setFontColor('#e2dac2').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(1, 1, 1, ncols).setBorder(null, null, true, null, null, null, '#ec4624', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  all.setBorder(true, true, true, true, false, false, '#222a2a', SpreadsheetApp.BorderStyle.SOLID);
  if (nrows > 0) {
    sh.getRange(2, 1, nrows, ncols).setBorder(null, null, null, null, false, true, '#d9d5c7', SpreadsheetApp.BorderStyle.SOLID);
    centerCols.forEach(function (c) { sh.getRange(2, c, nrows, 1).setHorizontalAlignment('center'); });
    sh.getRange(2, 1, nrows, 1).setNumberFormat(dateFmt);
  }
  sh.setRowHeight(1, 28);
  if (nrows > 0) sh.setRowHeights(2, nrows, 19);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, n + 1, ncols).createFilter();
  sh.setTabColor(tabColor);
  sh.setFrozenRows(1);
}

// A heavier line above every Monday in Days, so the weeks read as blocks.
function mondayLines(sh, days) {
  var cells = [];
  days.forEach(function (r, i) { if (r[1] === 'Mon' && i > 0) cells.push('A' + (i + 2) + ':M' + (i + 2)); });
  for (var k = 0; k < cells.length; k += 100) {
    sh.getRangeList(cells.slice(k, k + 100)).setBorder(true, null, null, null, null, null, '#222a2a', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

// The full copy the app restores from. Cells are forced to plain text so nothing is turned into a formula.
function saveBackup(ss, json) {
  var sh = ss.getSheetByName('Backup') || ss.insertSheet('Backup');
  sh.clear();
  var rows = [['Saved at', new Date().toISOString()], ['Do not edit: the app restores from this tab', '']];
  for (var i = 0; i < json.length; i += 40000) rows.push([json.substr(i, 40000), '']);
  var r = sh.getRange(1, 1, rows.length, 2);
  r.setNumberFormat('@');
  r.setValues(rows);
}

function put(ss, name, head, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#222a2a').setFontColor('#e2dac2');
  if (rows && rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.setFrozenRows(1);
}

function monday(s) {
  var p = s.split('-'), x = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  x.setUTCDate(x.getUTCDate() - ((x.getUTCDay() + 6) % 7));
  return x.toISOString().slice(0, 10);
}

function weekly(d, days) {
  var w = {}, seen = {};
  function get(k) { return w[k] || (w[k] = { s: 0, sets: 0, min: 0, c: 0, t: 0 }); }
  d.log.forEach(function (r) {
    var wk = monday(String(r[0])), k = get(wk), id = r[0] + '|' + r[1] + '|' + r[3];
    k.sets += Number(r[6]) || 0;
    if (!seen[id]) { seen[id] = 1; k.s += 1; k.min += Number(r[9]) || 0; }
  });
  days.forEach(function (r) {
    var k = get(monday(String(r[0])));
    if (r[8] === 'Yes') k.t += 1;
    k.c += Number(r[6]) || 0;
  });
  return Object.keys(w).sort().map(function (k) { return [k, w[k].s, w[k].sets, w[k].min, w[k].c, w[k].t]; });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

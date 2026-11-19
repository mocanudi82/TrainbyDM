/* Training Suite by DM - Google Sheet script (v27)
   Paste into Extensions > Apps Script of the sheet, then Deploy > Web app (Execute as: Me, Access: Anyone).
   - Receives the log from the app and writes it to tabs of this sheet.
   - Saves a backup in Drive after every finished session.
   - Input: tabs "In · ..." hold weights, exercises and the year plan; the app reads them.
   - Output: a workbook in Drive, updated after every finished workout when the app switches it on. */

var DRIVE_FOLDER = 'Training Suite backups';
var DRIVE_KEEP = 30;
var PFX = 'In · ';
var TEAL = '#2b8f91', TEAL_L = '#d9eeee', GREEN_L = '#d8f0d2', GREY = '#8a8f94';

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var a = (e && e.parameter && e.parameter.action) || '';
  try {
    if (a === 'backup') return json_(readBackup_());
    if (a === 'input') return json_(readInput_());
    if (a === 'status') {
      var o = PropertiesService.getScriptProperties().getProperty('outInfo');
      return json_({ ok: true, out: o ? JSON.parse(o) : null, input: inputTabs_().length > 0 });
    }
    return json_({ ok: true, service: 'training-suite' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var p = JSON.parse(e.postData.contents);
    if (p.action === 'makeInput') return json_(makeInput_(p));
    var res = { ok: true };
    res.rows = writeSync_(p);
    try { res.drive = driveBackup_(p); } catch (err) { res.drive = 'error: ' + err; }
    try { res.out = buildOutput_(p); } catch (err) { res.out = 'error: ' + err; }
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- plain tabs ---------- */

function tab_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function writeTab_(ss, name, header, rows) {
  var sh = tab_(ss, name);
  sh.clearContents();
  var all = [header].concat(rows || []);
  var w = 0;
  all.forEach(function (r) { if (r.length > w) w = r.length; });
  all = all.map(function (r) { var x = r.slice(); while (x.length < w) x.push(''); return x; });
  sh.getRange(1, 1, all.length, w).setValues(all);
  sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground(TEAL).setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function writeSync_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  writeTab_(ss, 'Sets log', ['Date', 'Program', 'Week', 'Session', 'Exercise', 'Prescribed', 'Sets done', 'Sets planned', 'kg', 'Minutes', 'Pain', 'Notes', 'Readiness'], p.log || []);
  writeTab_(ss, 'Days', ['Date', 'Day', 'Training', 'Week / day', 'Cardio', 'Mode', 'Cardio min', 'Mobility', 'Tango', 'Trip', 'Readiness', 'Minutes', 'Comment'], p.days || []);
  writeTab_(ss, 'Trips', ['Name', 'From', 'To', 'Kind', 'Maybe'], p.trips || []);
  writeTab_(ss, 'Profile', ['Item', 'Value'], p.profile || []);
  writeTab_(ss, 'Cardio', ['Date', 'Session', 'Mode', 'Minutes', 'Avg HR', 'km', 'Pace', 'Zone'], p.cardio || []);
  if (p.backup) writeBackup_(ss, p.backup);
  return (p.log || []).length;
}

function writeBackup_(ss, text) {
  var sh = tab_(ss, 'Backup');
  sh.clearContents();
  var parts = [], n = 40000;
  for (var i = 0; i < text.length; i += n) parts.push([text.substring(i, i + n)]);
  sh.getRange(1, 1, 1, 2).setValues([['saved', new Date().toISOString()]]);
  if (parts.length) sh.getRange(2, 1, parts.length, 1).setValues(parts);
}

function readBackup_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Backup');
  if (!sh || sh.getLastRow() < 2) return { ok: false, error: 'The sheet has no backup yet' };
  var at = sh.getRange(1, 2).getValue();
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  return { ok: true, at: String(at), json: vals.map(function (r) { return r[0]; }).join('') };
}

/* ---------- Drive backup after every finished session ---------- */

function folder_() {
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(DRIVE_FOLDER);
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function driveBackup_(p) {
  if (!p.backup || p.sessFin == null) return 'skipped';
  var props = PropertiesService.getScriptProperties();
  var prev = props.getProperty('driveSess');
  if (prev !== null && String(prev) === String(p.sessFin)) return 'unchanged';
  var f = folder_(), name = 'Training_backup_' + today_() + '.json';
  var it = f.getFilesByName(name);
  if (it.hasNext()) it.next().setContent(p.backup); else f.createFile(name, p.backup, 'application/json');
  var files = [], all = f.getFiles();
  while (all.hasNext()) { var x = all.next(); if (/^Training_backup_/.test(x.getName())) files.push(x); }
  files.sort(function (a, b) { return a.getName() < b.getName() ? 1 : -1; });
  for (var i = DRIVE_KEEP; i < files.length; i++) files[i].setTrashed(true);
  props.setProperty('driveSess', String(p.sessFin));
  return 'saved ' + name;
}

/* ---------- input tabs ---------- */

function inputTabs_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(function (s) {
    return s.getName().indexOf(PFX) === 0;
  });
}

function cut_(s) { return String(s).substring(0, 31); }

function makeInput_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var have = inputTabs_();
  if (have.length && !p.overwrite) return { ok: true, exists: true, tabs: have.length };
  var inp = p.input || {};
  var progs = inp.programs || {};
  var made = 0;
  Object.keys(progs).forEach(function (id) {
    var pr = progs[id];
    var name = cut_(PFX + pr.name);
    var sh = ss.getSheetByName(name);
    if (sh) sh.clear(); else sh = ss.insertSheet(name);
    var nw = pr.weeks.length;
    var head = ['Session', 'Block', 'Id', 'Exercise', 'Unit'];
    pr.weeks.forEach(function (w) { head.push(w + ' sets×reps', w + ' kg'); });
    var width = head.length;
    var out = [], bands = [], lastS = null;
    pr.rows.forEach(function (r) {
      if (r[0] !== lastS) {
        var b = []; for (var i = 0; i < width; i++) b.push('');
        b[3] = 'Session ' + r[0];
        out.push(b); bands.push(out.length);
        lastS = r[0];
      }
      var x = r.slice(); while (x.length < width) x.push('');
      out.push(x);
    });
    sh.getRange(1, 1).setValue(pr.name).setFontWeight('bold').setFontSize(13);
    sh.getRange(1, 6).setValue('id:' + id).setFontColor(GREY);
    sh.getRange(2, 1, 1, width).setValues([head]).setFontWeight('bold').setBackground(TEAL).setFontColor('#ffffff');
    if (out.length) {
      var rg = sh.getRange(3, 1, out.length, width);
      rg.setNumberFormat('@');
      for (var wi = 0; wi < nw; wi++) sh.getRange(3, 7 + wi * 2, out.length, 1).setNumberFormat('0.##');
      rg.setValues(out);
      bands.forEach(function (b) { sh.getRange(2 + b, 1, 1, width).setBackground(TEAL_L).setFontWeight('bold'); });
      sh.getRange(3, 6, out.length, nw * 2).setHorizontalAlignment('center');
    }
    sh.getRange(3, 1, Math.max(out.length, 1), 3).setFontColor(GREY);
    sh.setFrozenRows(2);
    sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 210); sh.setColumnWidth(3, 70); sh.setColumnWidth(4, 260); sh.setColumnWidth(5, 90);
    made++;
  });
  // year plan
  var pl = inp.plan || { blocks: [], trips: [] };
  var psh = ss.getSheetByName(PFX + 'Plan');
  if (psh) psh.clear(); else psh = ss.insertSheet(PFX + 'Plan');
  var rows = [];
  (pl.blocks || []).forEach(function (b) { rows.push(['Block', b[0], '', b[1], b[2], '', '']); });
  (pl.trips || []).forEach(function (t) { rows.push(['Trip', t[3], t[0], t[1], t[2], t[4], t[5]]); });
  psh.getRange(1, 1).setValue('Year plan').setFontWeight('bold').setFontSize(13);
  psh.getRange(2, 1, 1, 7).setValues([['Kind', 'Type / kind', 'Name', 'From', 'To / weeks', 'Maybe', 'Lead weeks']]).setFontWeight('bold').setBackground(TEAL).setFontColor('#ffffff');
  if (rows.length) {
    psh.getRange(3, 1, rows.length, 7).setNumberFormat('@').setValues(rows);
  }
  psh.setFrozenRows(2);
  made++;
  return { ok: true, tabs: made };
}

function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}

function readInput_() {
  var tabs = inputTabs_();
  if (!tabs.length) return { ok: false, error: 'The sheet has no input yet' };
  var programs = {}, plan = null;
  tabs.forEach(function (sh) {
    var last = sh.getLastRow(), cols = sh.getLastColumn();
    if (last < 3 || cols < 1) return;
    if (sh.getName() === PFX + 'Plan') {
      var v = sh.getRange(3, 1, last - 2, 7).getValues();
      var blocks = [], trips = [];
      v.forEach(function (r) {
        var k = String(r[0]).toLowerCase();
        if (k === 'block' && r[1] && r[3]) blocks.push([String(r[1]), cell_(r[3]), r[4]]);
        else if (k === 'trip' && r[3]) trips.push([String(r[2]), cell_(r[3]), cell_(r[4]), String(r[1]), String(r[5]), r[6]]);
      });
      plan = { blocks: blocks, trips: trips };
      return;
    }
    var idc = String(sh.getRange(1, 6).getValue());
    if (idc.indexOf('id:') !== 0) return;
    var data = sh.getRange(3, 1, last - 2, cols).getValues();
    var rows = data.filter(function (r) { return String(r[0]).trim() !== ''; }).map(function (r) { return r.map(cell_); });
    programs[idc.substring(3)] = { rows: rows };
  });
  return { ok: true, programs: programs, plan: plan };
}

/* ---------- output workbook in Drive ---------- */

function buildOutput_(p) {
  if (!p.out) return 'off';
  var props = PropertiesService.getScriptProperties();
  var prevSess = props.getProperty('outSess'), prevForce = props.getProperty('outForce');
  var forced = p.outForce && String(p.outForce) !== prevForce;
  var changed = p.sessFin != null && String(p.sessFin) !== prevSess;
  var first = !props.getProperty('outId');
  if (!forced && !changed && !first) return 'unchanged';
  var src = inputTabs_().filter(function (s) { return s.getName() !== PFX + 'Plan'; });
  if (!src.length) return 'no input tabs yet';
  var used = p.used || {};
  var name = 'Train_output_' + today_();
  var out = null, oid = props.getProperty('outId');
  if (oid) { try { out = SpreadsheetApp.openById(oid); } catch (e) { out = null; } }
  if (!out) {
    out = SpreadsheetApp.create(name);
    var file = DriveApp.getFileById(out.getId());
    folder_().addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    props.setProperty('outId', out.getId());
  }
  out.rename(name);
  var keep = {};
  src.forEach(function (sh) {
    var idc = String(sh.getRange(1, 6).getValue());
    if (idc.indexOf('id:') !== 0) return;
    var id = idc.substring(3), prog = String(sh.getRange(1, 1).getValue());
    var last = sh.getLastRow(), cols = sh.getLastColumn();
    var vals = sh.getRange(1, 1, last, cols).getValues();
    var u = used[id] || {};
    var greens = [];
    for (var r = 2; r < vals.length; r++) {
      var row = vals[r], iid = String(row[2]);
      if (!String(row[0]).trim() || !iid) continue;
      for (var wi = 0; (6 + wi * 2) < cols; wi++) {
        var key = wi + '|' + iid;
        if (u[key] != null) { row[6 + wi * 2] = u[key]; greens.push([r + 1, 7 + wi * 2]); }
      }
    }
    vals[0][5] = '';
    vals[0][0] = prog + ' · kg used where logged';
    var tn = cut_('Out · ' + prog);
    keep[tn] = 1;
    var t = out.getSheetByName(tn);
    if (t) t.clear(); else t = out.insertSheet(tn);
    t.getRange(1, 1, vals.length, cols).setNumberFormat('@').setValues(vals);
    t.getRange(1, 1).setFontWeight('bold').setFontSize(13);
    t.getRange(2, 1, 1, cols).setFontWeight('bold').setBackground(TEAL).setFontColor('#ffffff');
    for (var r2 = 2; r2 < vals.length; r2++) if (!String(vals[r2][0]).trim() && String(vals[r2][3]).indexOf('Session') === 0) t.getRange(r2 + 1, 1, 1, cols).setBackground(TEAL_L).setFontWeight('bold');
    t.getRange(3, 1, Math.max(vals.length - 2, 1), 3).setFontColor(GREY);
    greens.forEach(function (g) { t.getRange(g[0], g[1]).setBackground(GREEN_L).setFontWeight('bold'); });
    t.getRange(3, 6, Math.max(vals.length - 2, 1), cols - 5).setHorizontalAlignment('center');
    t.setFrozenRows(2);
    t.setColumnWidth(1, 60); t.setColumnWidth(2, 210); t.setColumnWidth(3, 70); t.setColumnWidth(4, 260); t.setColumnWidth(5, 90);
  });
  var lg = out.getSheetByName('Sets log') || out.insertSheet('Sets log'); keep['Sets log'] = 1;
  lg.clear();
  var lrows = [['Date', 'Program', 'Week', 'Session', 'Exercise', 'Prescribed', 'Sets done', 'Sets planned', 'kg', 'Minutes', 'Pain', 'Notes', 'Readiness']].concat(p.log || []);
  var lw = 13;
  lrows = lrows.map(function (r) { var x = r.slice(0, lw); while (x.length < lw) x.push(''); return x; });
  lg.getRange(1, 1, lrows.length, lw).setValues(lrows);
  lg.getRange(1, 1, 1, lw).setFontWeight('bold').setBackground(TEAL).setFontColor('#ffffff');
  lg.setFrozenRows(1);
  out.getSheets().forEach(function (s) { if (!keep[s.getName()]) { try { out.deleteSheet(s); } catch (e) {} } });
  SpreadsheetApp.flush();
  // Excel copy next to it
  var xname = name + '.xlsx';
  var blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + out.getId() + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(xname);
  var f = folder_();
  var old = f.getFiles();
  while (old.hasNext()) { var o = old.next(); if (/^Train_output_.*\.xlsx$/.test(o.getName())) o.setTrashed(true); }
  f.createFile(blob);
  props.setProperty('outSess', String(p.sessFin));
  if (p.outForce) props.setProperty('outForce', String(p.outForce));
  var info = { name: xname, url: out.getUrl(), at: new Date().toISOString() };
  props.setProperty('outInfo', JSON.stringify(info));
  return 'saved ' + xname;
}

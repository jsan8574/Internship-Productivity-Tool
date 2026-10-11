/**
 * Verify Desk — Google Apps Script backend
 * --------------------------------------------------------------
 * Insurance Benefits & Eligibility Verification productivity tool
 * for interns. The Google Sheet this script is bound to is the
 * single source of truth; the web app reads and writes it directly.
 *
 * SETUP (one time, ~10 minutes)
 *  1. Upload "Internship Insurance Verification.xlsx" to Google Drive and
 *     open it with Google Sheets (File > Save as Google Sheets).
 *  2. Extensions > Apps Script. Replace Code.gs with this file.
 *     Click + > HTML, name it exactly "Index", paste Index.html.
 *  3. Save, select the function "setup" in the toolbar, click Run, and
 *     approve the permissions. This adds the tool columns and the
 *     Roster / Activity Log tabs, and makes you a Manager.
 *  4. Open the "Roster" tab and add every intern: NAME, EMAIL (their
 *     company Google account), ROLE = Intern (or Manager / QA), ACTIVE = Yes.
 *  5. Deploy > New deployment > type "Web app".
 *       Execute as:      Me
 *       Who has access:  Anyone within your organization
 *     Copy the web app URL and share it with interns. Interns do NOT need
 *     access to the Sheet itself — they only ever see rows assigned to them.
 *  6. After any later code change: Deploy > Manage deployments > Edit >
 *     Version: New version > Deploy (the URL stays the same).
 */

// ---------- Configuration ----------
var CFG = {
  VERIF_SHEET: 'Pt Verif',
  ROSTER_SHEET: 'Roster',
  LOG_SHEET: 'Activity Log',
  TZ: Session.getScriptTimeZone() || 'Asia/Manila'
};

// Original template columns (kept in this exact order) + tool columns appended.
var COLS = {
  verifDate:   'DATE OF VERIFICATION',
  internName:  'INTERN NAME',
  status:      'APPROVAL STATUS',
  provider:    'PROVIDER NAME',
  patient:     'PATIENT NAME',
  dob:         'DOB',
  tos:         'TYPE OF SERVICE',
  dos:         'DOS',
  plan:        'INSURANCE PLAN',
  contracted:  'CONTRACTED (YES OR NO)',
  eligibility: 'ELIGIBILITY (YES OR NO)',
  hospice:     'Patient in hospice? (Yes or No)',
  caseName:    'Case Name Updated? (Yes or No)',
  notes:       'Notes Updated? (Yes or No)',
  comment:     'Comment',
  qa:          'QA',
  qaComment:   'QA Comment',
  // ---- tool columns ----
  id:          'RECORD ID',
  assignedOn:  'ASSIGNED DATE',
  internEmail: 'INTERN EMAIL',
  work:        'WORK STATUS',
  source:      'VERIFIED VIA',
  memberId:    'MEMBER ID',
  reference:   'REFERENCE # / REP',
  mode:        'GUIDANCE MODE',
  startedAt:   'STARTED AT',
  completedAt: 'COMPLETED AT',
  qaBy:        'QA BY',
  updatedBy:   'LAST UPDATED BY'
};
var COL_ORDER = ['verifDate','internName','status','provider','patient','dob','tos','dos','plan',
  'contracted','eligibility','hospice','caseName','notes','comment','qa','qaComment',
  'id','assignedOn','internEmail','work','source','memberId','reference','mode','startedAt','completedAt','qaBy','updatedBy'];

var APPROVAL_STATUSES = ['Good to Schedule','Provider out of Network','Policy Not Active','Deceased',
  'Approved for Self Pay','No self pay agreement','Ok for Group Billing','Frequency Requirement Not Met',
  'TEBRA INACTIVE','ON HOLD','No TCM Coverage'];
var WORK_STATUSES = ['Assigned','In Progress','Completed','Escalated'];
var YES_NO = ['Yes','No'];

// Fields an intern may write. Manager-provided fields are read-only for interns.
var INTERN_FIELDS = ['dos','plan','contracted','eligibility','hospice','caseName','notes','comment',
  'status','source','memberId','reference','mode'];

// ---------- Web app entry ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Verify Desk')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Verify Desk')
    .addItem('Run setup / repair columns', 'setup')
    .addItem('Assign IDs to rows added directly in the sheet', 'menuAdopt')
    .addToUi();
}

function menuAdopt() {
  var n = adoptNewRows_();
  SpreadsheetApp.getActive().toast(n + ' row(s) picked up by the tool.');
}

// ---------- Setup ----------
function setup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CFG.VERIF_SHEET) || ss.insertSheet(CFG.VERIF_SHEET);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var norm = headers.map(normHeader_);
  COL_ORDER.forEach(function (key) {
    if (norm.indexOf(normHeader_(COLS[key])) === -1) {
      var col = firstEmptyHeaderCol_(sh);
      sh.getRange(1, col).setValue(COLS[key]);
      norm[col - 1] = normHeader_(COLS[key]);
    }
  });
  var map = colMap_(sh);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#123b3b').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  // Keep text-like columns as plain text so Sheets does not reformat DOB / dates.
  ['dob','dos','assignedOn','id','memberId','reference','startedAt','completedAt','verifDate'].forEach(function (k) {
    sh.getRange(2, map[k], sh.getMaxRows() - 1, 1).setNumberFormat('@');
  });
  // Dropdowns
  setList_(sh, map.status, APPROVAL_STATUSES);
  setList_(sh, map.work, WORK_STATUSES);
  ['contracted','eligibility','hospice','caseName','notes'].forEach(function (k) { setList_(sh, map[k], YES_NO); });
  setList_(sh, map.qa, ['PASSED','FAILED']);
  setList_(sh, map.source, ['TEBRA','Payer portal','Payer call','Clearinghouse']);

  // Roster
  var ro = ss.getSheetByName(CFG.ROSTER_SHEET);
  if (!ro) {
    ro = ss.insertSheet(CFG.ROSTER_SHEET);
    ro.getRange(1, 1, 1, 4).setValues([['NAME','EMAIL','ROLE','ACTIVE']]).setFontWeight('bold');
    var me = Session.getActiveUser().getEmail();
    ro.getRange(2, 1, 1, 4).setValues([[me.split('@')[0], me, 'Manager', 'Yes']]);
    setList_(ro, 3, ['Intern','Manager','QA']);
    setList_(ro, 4, YES_NO);
    ro.setFrozenRows(1);
  }
  // Activity log
  var lg = ss.getSheetByName(CFG.LOG_SHEET);
  if (!lg) {
    lg = ss.insertSheet(CFG.LOG_SHEET);
    lg.getRange(1, 1, 1, 5).setValues([['TIMESTAMP','USER','ACTION','RECORD ID','DETAIL']]).setFontWeight('bold');
    lg.setFrozenRows(1);
  }
  adoptNewRows_();
  return 'Setup complete';
}

// ---------- Single API dispatcher (called from the page) ----------
function api(action, payload) {
  payload = payload || {};
  var me = currentUser_();
  if (!me.role) throw new Error('Your Google account (' + me.email + ') is not on the Roster yet. Ask your manager to add you.');
  switch (action) {
    case 'bootstrap': return bootstrap_(me);
    case 'queue':     return queue_(me, payload);
    case 'record':    return getRecord_(me, payload.id);
    case 'start':     return startRecord_(me, payload.id);
    case 'save':      return saveRecord_(me, payload);
    case 'upload':    requireManager_(me); return upload_(me, payload);
    case 'reassign':  requireManager_(me); return reassign_(me, payload);
    case 'qa':        requireQA_(me); return qa_(me, payload);
    case 'summary':   requireManager_(me); return summary_(payload);
    case 'stats':     return stats_(me, payload);
    default: throw new Error('Unknown action: ' + action);
  }
}

function bootstrap_(me) {
  adoptNewRows_();
  return {
    me: me,
    roster: roster_().filter(function (r) { return r.active; }),
    statuses: APPROVAL_STATUSES,
    today: today_()
  };
}

// ---------- Reads ----------
function queue_(me, p) {
  var all = readAll_();
  var date = p.date || '';
  var rows = all.filter(function (r) {
    if (me.role === 'Intern' && r.internEmail.toLowerCase() !== me.email) return false;
    if (p.email && r.internEmail.toLowerCase() !== p.email.toLowerCase()) return false;
    if (p.needsQA) return r.work === 'Completed' && !r.qa;
    if (date && r.assignedOn !== date) {
      // Always carry forward unfinished work from earlier days.
      return !!p.includeOpen && r.work !== 'Completed' && r.assignedOn < date;
    }
    return true;
  });
  return rows;
}

function getRecord_(me, id) {
  var r = findById_(id);
  if (!r) throw new Error('Record not found. It may have been removed from the sheet.');
  assertCanSee_(me, r);
  return r;
}

function stats_(me, p) {
  var email = (me.role === 'Intern' || !p.email) ? me.email : p.email.toLowerCase();
  var mine = readAll_().filter(function (r) { return r.internEmail.toLowerCase() === email; });
  var done = mine.filter(function (r) { return r.work === 'Completed' || /^fail/i.test(r.qa); });
  var qaDone = done.filter(function (r) { return r.qa; });
  var passed = qaDone.filter(function (r) { return /^pass/i.test(r.qa); });
  var fails = {};
  qaDone.filter(function (r) { return /^fail/i.test(r.qa); }).forEach(function (r) {
    var k = (r.qaComment || 'No comment').trim(); fails[k] = (fails[k] || 0) + 1;
  });
  var guided = done.filter(function (r) { return r.mode === 'Guided'; }).length;
  return {
    completed: done.length, guided: guided, express: done.length - guided,
    qaReviewed: qaDone.length, qaPassed: passed.length,
    topFails: Object.keys(fails).map(function (k) { return { reason: k, count: fails[k] }; })
      .sort(function (a, b) { return b.count - a.count; }).slice(0, 5)
  };
}

function summary_(p) {
  var date = p.date || today_();
  var rows = readAll_().filter(function (r) { return r.assignedOn === date || (r.work !== 'Completed' && r.assignedOn < date); });
  var by = {};
  roster_().filter(function (x) { return x.active && x.role === 'Intern'; }).forEach(function (x) {
    by[x.email] = { name: x.name, email: x.email, assigned: 0, inProgress: 0, completed: 0, escalated: 0 };
  });
  var statusCounts = {};
  rows.forEach(function (r) {
    var k = r.internEmail.toLowerCase() || '(unassigned)';
    if (!by[k]) by[k] = { name: r.internName || k, email: k, assigned: 0, inProgress: 0, completed: 0, escalated: 0 };
    by[k].assigned++;
    if (r.work === 'In Progress') by[k].inProgress++;
    if (r.work === 'Completed') by[k].completed++;
    if (r.work === 'Escalated') by[k].escalated++;
    if (r.status) statusCounts[r.status] = (statusCounts[r.status] || 0) + 1;
  });
  return { date: date, interns: Object.keys(by).map(function (k) { return by[k]; }), statusCounts: statusCounts, rows: rows };
}

// ---------- Writes ----------
function startRecord_(me, id) {
  return withLock_(function () {
    var r = findById_(id);
    assertCanEdit_(me, r);
    var patch = {};
    if (r.work === 'Assigned' || !r.work) patch.work = 'In Progress';
    if (!r.startedAt) patch.startedAt = nowStamp_();
    if (Object.keys(patch).length) { writePatch_(r._row, patch, me); log_(me, 'start', id, ''); }
    return findById_(id);
  });
}

function saveRecord_(me, p) {
  return withLock_(function () {
    var r = findById_(p.id);
    assertCanEdit_(me, r);
    var patch = {};
    Object.keys(p.fields || {}).forEach(function (k) {
      if (INTERN_FIELDS.indexOf(k) !== -1) patch[k] = String(p.fields[k] == null ? '' : p.fields[k]).slice(0, 2000);
    });
    if (patch.status && APPROVAL_STATUSES.indexOf(patch.status) === -1) throw new Error('Unknown approval status: ' + patch.status);
    if (p.complete) {
      var missing = ['status','plan','eligibility','caseName','notes'].filter(function (k) { return !(patch[k] || r[k]); });
      if (missing.length) throw new Error('Fill these before completing: ' + missing.map(function (k) { return COLS[k]; }).join(', '));
      patch.work = 'Completed';
      patch.completedAt = nowStamp_();
      patch.verifDate = today_();
      patch.internName = me.name;
    } else if (p.escalate) {
      patch.work = 'Escalated';
    } else if (r.work === 'Assigned') {
      patch.work = 'In Progress';
    }
    if (r.qa && p.complete) { patch.qa = ''; } // re-submitted after a QA fail → back into the QA queue
    writePatch_(r._row, patch, me);
    log_(me, p.complete ? 'complete' : (p.escalate ? 'escalate' : 'save'), p.id, patch.status || '');
    return findById_(p.id);
  });
}

function upload_(me, p) {
  var rows = p.rows || [];
  if (!rows.length) throw new Error('No rows to upload.');
  var date = p.date || today_();
  var rosterByEmail = {};
  roster_().forEach(function (x) { rosterByEmail[x.email] = x; });
  return withLock_(function () {
    var sh = sheet_();
    var map = colMap_(sh);
    var width = sh.getLastColumn();
    var out = rows.map(function (r) {
      var line = new Array(width).fill('');
      var email = String(r.internEmail || '').toLowerCase();
      var who = rosterByEmail[email];
      set_(line, map, 'provider', r.provider);
      set_(line, map, 'patient', r.patient);
      set_(line, map, 'dob', r.dob);
      set_(line, map, 'tos', r.tos);
      set_(line, map, 'dos', r.dos || '');
      set_(line, map, 'internName', who ? who.name : '');
      set_(line, map, 'internEmail', email);
      set_(line, map, 'assignedOn', date);
      set_(line, map, 'work', 'Assigned');
      set_(line, map, 'id', newId_());
      set_(line, map, 'updatedBy', me.email);
      return line;
    });
    var start = lastDataRow_(sh) + 1;
    sh.getRange(start, 1, out.length, width).setNumberFormat('@').setValues(out);
    log_(me, 'upload', '', out.length + ' rows for ' + date);
    return { added: out.length, date: date };
  });
}

function reassign_(me, p) {
  var who = roster_().filter(function (x) { return x.email === String(p.email).toLowerCase(); })[0];
  if (!who) throw new Error('That intern is not on the Roster.');
  return withLock_(function () {
    (p.ids || []).forEach(function (id) {
      var r = findById_(id);
      if (r) writePatch_(r._row, { internEmail: who.email, internName: who.name }, me);
    });
    log_(me, 'reassign', (p.ids || []).join(','), who.email);
    return { moved: (p.ids || []).length };
  });
}

function qa_(me, p) {
  return withLock_(function () {
    var r = findById_(p.id);
    if (!r) throw new Error('Record not found.');
    var result = /^pass/i.test(p.result) ? 'PASSED' : 'FAILED';
    var patch = { qa: result, qaComment: p.comment || '', qaBy: me.email };
    if (result === 'FAILED') patch.work = 'In Progress'; // send back to the intern to fix
    writePatch_(r._row, patch, me);
    log_(me, 'qa', p.id, result + (p.comment ? ' — ' + p.comment : ''));
    return findById_(p.id);
  });
}

// ---------- Helpers ----------
function currentUser_() {
  var email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  var r = roster_().filter(function (x) { return x.email === email && x.active; })[0];
  return { email: email, name: r ? r.name : email.split('@')[0], role: r ? r.role : '' };
}
function roster_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.ROSTER_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues()
    .filter(function (r) { return r[1]; })
    .map(function (r) {
      var role = /manager/i.test(r[2]) ? 'Manager' : (/qa/i.test(r[2]) ? 'QA' : 'Intern');
      return { name: r[0] || r[1].split('@')[0], email: r[1].trim().toLowerCase(), role: role, active: !/^no/i.test(r[3]) };
    });
}
function requireManager_(me) { if (me.role !== 'Manager') throw new Error('Only managers can do this.'); }
function requireQA_(me) { if (me.role !== 'Manager' && me.role !== 'QA') throw new Error('Only QA reviewers or managers can do this.'); }
function assertCanSee_(me, r) {
  if (me.role === 'Intern' && r.internEmail.toLowerCase() !== me.email) throw new Error('This record is assigned to someone else.');
}
function assertCanEdit_(me, r) {
  if (!r) throw new Error('Record not found.');
  if (me.role === 'Intern' && r.internEmail.toLowerCase() !== me.email) throw new Error('This record is assigned to someone else.');
}

function sheet_() { return SpreadsheetApp.getActive().getSheetByName(CFG.VERIF_SHEET); }
function normHeader_(h) { return String(h || '').toLowerCase().replace(/[^a-z0-9#]/g, ''); }
function colMap_(sh) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(normHeader_);
  var map = {};
  COL_ORDER.forEach(function (k) {
    var i = headers.indexOf(normHeader_(COLS[k]));
    if (i === -1) throw new Error('Column "' + COLS[k] + '" is missing. Run setup from the Verify Desk menu.');
    map[k] = i + 1;
  });
  return map;
}
function firstEmptyHeaderCol_(sh) {
  var h = sh.getRange(1, 1, 1, sh.getMaxColumns()).getDisplayValues()[0];
  for (var i = 0; i < h.length; i++) if (!h[i]) return i + 1;
  sh.insertColumnAfter(h.length); return h.length + 1;
}
function lastDataRow_(sh) {
  var map = colMap_(sh);
  var n = sh.getLastRow(); if (n < 2) return 1;
  var vals = sh.getRange(2, map.patient, n - 1, 1).getDisplayValues();
  for (var i = vals.length - 1; i >= 0; i--) if (vals[i][0]) return i + 2;
  return 1;
}
function readAll_() {
  var sh = sheet_(); var map = colMap_(sh);
  var last = lastDataRow_(sh); if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getDisplayValues();
  var out = [];
  vals.forEach(function (row, i) {
    if (!row[map.patient - 1] || !row[map.id - 1]) return;
    var o = { _row: i + 2 };
    COL_ORDER.forEach(function (k) { o[k] = row[map[k] - 1]; });
    if (!o.work) o.work = 'Assigned';
    out.push(o);
  });
  return out;
}
function findById_(id) { return readAll_().filter(function (r) { return r.id === id; })[0] || null; }
function writePatch_(rowNum, patch, me) {
  var sh = sheet_(); var map = colMap_(sh);
  patch.updatedBy = me.email;
  Object.keys(patch).forEach(function (k) { if (map[k]) sh.getRange(rowNum, map[k]).setValue(patch[k]); });
}
function set_(line, map, key, val) { line[map[key] - 1] = val == null ? '' : String(val); }
function setList_(sh, col, list) {
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build());
}
// Rows the manager typed straight into the sheet: give them an ID, assignee email, and date.
function adoptNewRows_() {
  var sh = sheet_(); if (!sh) return 0;
  var map; try { map = colMap_(sh); } catch (e) { return 0; }
  var last = lastDataRow_(sh); if (last < 2) return 0;
  var rng = sh.getRange(2, 1, last - 1, sh.getLastColumn());
  var vals = rng.getDisplayValues();
  var byName = {}; roster_().forEach(function (x) { byName[x.name.toLowerCase()] = x; });
  var n = 0;
  vals.forEach(function (row, i) {
    if (row[map.patient - 1] && !row[map.id - 1]) {
      var r = i + 2;
      sh.getRange(r, map.id).setValue(newId_());
      var who = byName[String(row[map.internName - 1]).toLowerCase().trim()];
      if (who && !row[map.internEmail - 1]) sh.getRange(r, map.internEmail).setValue(who.email);
      if (!row[map.assignedOn - 1]) sh.getRange(r, map.assignedOn).setValue(today_());
      if (!row[map.work - 1]) sh.getRange(r, map.work).setValue(row[map.status - 1] ? 'Completed' : 'Assigned');
      n++;
    }
  });
  return n;
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The sheet is busy. Try again in a few seconds.');
  try { return fn(); } finally { lock.releaseLock(); }
}
function log_(me, action, id, detail) {
  var lg = SpreadsheetApp.getActive().getSheetByName(CFG.LOG_SHEET);
  if (lg) lg.appendRow([nowStamp_(), me.email, action, id || '', detail || '']);
}
function newId_() { return 'IV-' + Utilities.getUuid().slice(0, 8).toUpperCase(); }
function today_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd'); }
function nowStamp_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm'); }

/**
 * מערכת נוכחות למועדון טניס — צד שרת (Google Apps Script)
 * ------------------------------------------------------------
 * סקריפט זה משמש כ"שרת" (API) עבור אפליקציית הנוכחות, מחובר לגיליון Google Sheets
 * ומחזיר תשובות בפורמט JSON.
 *
 * לשוניות (נוצרות/משודרגות אוטומטית בהרצה הראשונה):
 *   Coaches     : id | name
 *   Groups      : id | name | day | time | coachId | slots      (slots = JSON של אימונים שבועיים)
 *   Members     : id | name | groupId | phone                   (groupId = שדה ישן, לתאימות)
 *   Attendance  : date | groupId | memberId | status | markedBy | timestamp
 *   Memberships : memberId | groupId                            (שיוך רב-לרב חניך↔קבוצות)
 *   Sessions    : date | groupId | status | note | markedBy | timestamp   (ביטול מפגשים)
 *
 * פריסה: Deploy → New deployment → Web app, "Execute as: Me", "Who has access: Anyone".
 * לאחר כל שינוי בקוד יש לפרוס מחדש: Deploy → Manage deployments → עריכה → New version → Deploy.
 */

// ====== הגדרות לשוניות ======
var SHEETS = {
  Coaches:     ['id', 'name'],
  Groups:      ['id', 'name', 'day', 'time', 'coachId', 'slots'],
  Members:     ['id', 'name', 'groupId', 'phone', 'trial'],
  Attendance:  ['date', 'groupId', 'memberId', 'status', 'markedBy', 'timestamp'],
  Memberships: ['memberId', 'groupId'],
  Sessions:    ['date', 'groupId', 'status', 'note', 'markedBy', 'timestamp', 'makeup'],
  Absences:    ['date', 'groupId', 'memberId', 'note', 'by', 'timestamp'],
  Users:       ['id', 'name', 'role', 'password', 'refId']
};

// ====== נקודות כניסה ======

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  return route(params.action || 'ping', params);
}

function doPost(e) {
  var params = {};
  try {
    if (e && e.postData && e.postData.contents) params = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'גוף הבקשה אינו JSON תקין: ' + err });
  }
  return route(params.action || 'ping', params);
}

function route(action, p) {
  try {
    ensureSheets_();
    migrateMemberships_();
    seedUsers_();
    switch (action) {
      case 'login':           return jsonOut({ ok: true, data: login_(p) });
      case 'getUsers':        return jsonOut({ ok: true, data: getUsers_() });
      case 'addUser':         return jsonOut({ ok: true, data: addUser_(p) });
      case 'setUserPassword': return jsonOut({ ok: true, data: setUserPassword_(p) });
      case 'deleteUser':      return jsonOut({ ok: true, data: deleteUser_(p) });

      case 'ping':          return jsonOut({ ok: true, message: 'המערכת פעילה', time: new Date().toISOString() });
      case 'getAll':        return jsonOut({ ok: true, data: getAll_() });

      case 'addCoach':      return jsonOut({ ok: true, data: addCoach_(p) });
      case 'deleteCoach':   return jsonOut({ ok: true, data: deleteCoach_(p) });

      case 'addGroup':      return jsonOut({ ok: true, data: addGroup_(p) });
      case 'editGroup':     return jsonOut({ ok: true, data: editGroup_(p) });
      case 'deleteGroup':   return jsonOut({ ok: true, data: deleteGroup_(p) });

      case 'addMember':     return jsonOut({ ok: true, data: addMember_(p) });
      case 'editMember':    return jsonOut({ ok: true, data: editMember_(p) });
      case 'deleteMember':  return jsonOut({ ok: true, data: deleteMember_(p) });

      case 'saveAttendance':return jsonOut({ ok: true, data: saveAttendance_(p) });
      case 'getHistory':    return jsonOut({ ok: true, data: getHistory_(p) });

      case 'cancelSession': return jsonOut({ ok: true, data: cancelSession_(p) });
      case 'restoreSession':return jsonOut({ ok: true, data: restoreSession_(p) });
      case 'getSessions':   return jsonOut({ ok: true, data: getSessions_(p) });

      case 'reportAbsence':   return jsonOut({ ok: true, data: reportAbsence_(p) });
      case 'unreportAbsence': return jsonOut({ ok: true, data: unreportAbsence_(p) });
      case 'getAbsences':     return jsonOut({ ok: true, data: getAbsences_(p) });

      default:              return jsonOut({ ok: false, error: 'פעולה לא מוכרת: ' + action });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// ====== עזרי תשתית ======

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ensureSheets_() {
  var ss = ss_();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    var headers = SHEETS[name];
    var firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    var needHeader = false;
    for (var i = 0; i < headers.length; i++) { if (firstRow[i] !== headers[i]) { needHeader = true; break; } }
    if (needHeader) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  });
  ['Sheet1', 'גיליון1'].forEach(function (n) {
    var extra = ss.getSheetByName(n);
    if (extra && !SHEETS[n] && ss.getSheets().length > 1 && extra.getLastRow() === 0) {
      try { ss.deleteSheet(extra); } catch (e) {}
    }
  });
}

/** שדרוג חד-פעמי: המרת שדה groupId הישן מ-Members לשורות ב-Memberships. */
function migrateMemberships_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('mm_migrated') === '1') return;
  var memSheet = sheet_('Members');
  var vals = memSheet.getDataRange().getValues();
  var mship = sheet_('Memberships');
  if (vals.length >= 2 && mship.getLastRow() <= 1) {
    var headers = vals[0];
    var idIdx = headers.indexOf('id'), gIdx = headers.indexOf('groupId');
    if (idIdx >= 0 && gIdx >= 0) {
      var toAdd = [];
      for (var r = 1; r < vals.length; r++) {
        var mid = vals[r][idIdx], gid = vals[r][gIdx];
        if (mid && gid) toAdd.push([mid, gid]);
      }
      if (toAdd.length) mship.getRange(mship.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
    }
  }
  props.setProperty('mm_migrated', '1');
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) { ensureSheets_(); sh = ss_().getSheetByName(name); }
  return sh;
}

function readAll_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0], rows = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {}, empty = true;
    for (var c = 0; c < headers.length; c++) {
      var v = values[r][c];
      obj[headers[c]] = v;
      if (v !== '' && v !== null) empty = false;
    }
    if (!empty) { obj._row = r + 1; rows.push(obj); }
  }
  return rows;
}

function newId_(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

/** נרמול שעה ל-"HH:mm" (24 שעות). תומך גם בערך Date שהגיליון יצר. */
function fmtTimeGs_(v) {
  if (v === '' || v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'HH:mm');
  }
  var s = String(v), m = s.match(/(\d{1,2}):(\d{2})/);
  return m ? (('0' + Number(m[1])).slice(-2) + ':' + m[2]) : s;
}

/** נרמול תאריך ל-"yyyy-MM-dd". תומך גם בערך Date שהגיליון יצר. */
function fmtDateGs_(v) {
  if (v === '' || v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v), m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return s;
}

function findRowById_(name, id) {
  var sh = sheet_(name), values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(id)) return r + 1;
  }
  return -1;
}

function deleteRowsWhere_(name, col, value) {
  var sh = sheet_(name), values = sh.getDataRange().getValues();
  if (values.length < 2) return;
  var colIdx = values[0].indexOf(col);
  if (colIdx < 0) return;
  for (var r = values.length - 1; r >= 1; r--) {
    if (String(values[r][colIdx]) === String(value)) sh.deleteRow(r + 1);
  }
}

// ====== שיוכי חניך-קבוצה ======

function setMemberships_(memberId, groupIds) {
  deleteRowsWhere_('Memberships', 'memberId', memberId);
  var sh = sheet_('Memberships');
  var rows = (groupIds || []).filter(function (g) { return g; }).map(function (g) { return [memberId, g]; });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
}

// ====== לו"ז שבועי (slots) ======

function parseSlots_(v) {
  if (!v) return [];
  try {
    var a = (typeof v === 'string') ? JSON.parse(v) : v;
    if (Object.prototype.toString.call(a) !== '[object Array]') return [];
    return a.filter(function (s) { return s && (s.day !== undefined) && s.time; });
  } catch (e) { return []; }
}

function normalizeSlots_(p) {
  var slots = parseSlots_(p.slots);
  if (!slots.length && (p.day !== undefined && p.day !== null && p.day !== '')) {
    slots = [{ day: p.day, time: p.time || '' }];
  }
  if (!slots.length) slots = [{ day: '0', time: '16:00' }];
  return slots.map(function (s) { return { day: String(s.day), time: fmtTimeGs_(s.time), court: (s.court != null ? String(s.court) : '') }; });
}

// ====== getAll ======

function getAll_() {
  var byMember = {};
  readAll_('Memberships').forEach(function (x) {
    var k = String(x.memberId);
    (byMember[k] = byMember[k] || []).push(String(x.groupId));
  });

  return {
    coaches: readAll_('Coaches').map(function (c) { return { id: String(c.id), name: c.name }; }),
    groups:  readAll_('Groups').map(function (g) {
      var slots = parseSlots_(g.slots);
      if (!slots.length) slots = [{ day: String(g.day), time: fmtTimeGs_(g.time) }];
      slots = slots.map(function (s) { return { day: String(s.day), time: fmtTimeGs_(s.time), court: (s.court != null ? String(s.court) : '') }; });
      return { id: String(g.id), name: g.name, day: slots[0].day, time: slots[0].time, coachId: String(g.coachId), slots: slots };
    }),
    members: readAll_('Members').map(function (m) {
      return { id: String(m.id), name: m.name, phone: m.phone ? String(m.phone) : '', trial: truthy_(m.trial), groupIds: byMember[String(m.id)] || [] };
    })
  };
}

// ====== מאמנים ======

function addCoach_(p) {
  if (!p.name) throw new Error('חסר שם מאמן');
  var id = newId_('c');
  sheet_('Coaches').appendRow([id, p.name]);
  return { id: id, name: p.name };
}

function deleteCoach_(p) {
  var row = findRowById_('Coaches', p.id);
  if (row < 0) throw new Error('מאמן לא נמצא');
  sheet_('Coaches').deleteRow(row);
  return { id: p.id, deleted: true };
}

// ====== קבוצות ======

function addGroup_(p) {
  if (!p.name) throw new Error('חסר שם קבוצה');
  var id = newId_('g'), slots = normalizeSlots_(p);
  sheet_('Groups').appendRow([id, p.name, slots[0].day, slots[0].time, p.coachId || '', JSON.stringify(slots)]);
  return { id: id, name: p.name, day: slots[0].day, time: slots[0].time, coachId: String(p.coachId || ''), slots: slots };
}

function editGroup_(p) {
  var row = findRowById_('Groups', p.id);
  if (row < 0) throw new Error('קבוצה לא נמצאה');
  var slots = normalizeSlots_(p);
  sheet_('Groups').getRange(row, 1, 1, 6).setValues([[p.id, p.name, slots[0].day, slots[0].time, p.coachId || '', JSON.stringify(slots)]]);
  return { id: String(p.id), name: p.name, day: slots[0].day, time: slots[0].time, coachId: String(p.coachId || ''), slots: slots };
}

function deleteGroup_(p) {
  var row = findRowById_('Groups', p.id);
  if (row < 0) throw new Error('קבוצה לא נמצאה');
  sheet_('Groups').deleteRow(row);
  deleteRowsWhere_('Memberships', 'groupId', p.id);   // הסרת שיוכים בלבד; החניכים נשמרים
  return { id: p.id, deleted: true };
}

// ====== חניכים ======

function truthy_(v) { return v === true || v === 1 || v === 'TRUE' || v === 'true' || v === '1'; }

function addMember_(p) {
  if (!p.name) throw new Error('חסר שם חניך');
  var id = newId_('m');
  var trial = truthy_(p.trial);
  sheet_('Members').appendRow([id, p.name, '', p.phone || '', trial ? 'TRUE' : '']);
  setMemberships_(id, p.groupIds || []);
  return { id: id, name: p.name, phone: p.phone || '', trial: trial, groupIds: (p.groupIds || []).map(String) };
}

function editMember_(p) {
  var row = findRowById_('Members', p.id);
  if (row < 0) throw new Error('חניך לא נמצא');
  var trial = truthy_(p.trial);
  sheet_('Members').getRange(row, 1, 1, 5).setValues([[p.id, p.name, '', p.phone || '', trial ? 'TRUE' : '']]);
  setMemberships_(p.id, p.groupIds || []);
  return { id: String(p.id), name: p.name, phone: p.phone || '', trial: trial, groupIds: (p.groupIds || []).map(String) };
}

function deleteMember_(p) {
  var row = findRowById_('Members', p.id);
  if (row < 0) throw new Error('חניך לא נמצא');
  sheet_('Members').deleteRow(row);
  deleteRowsWhere_('Memberships', 'memberId', p.id);
  return { id: p.id, deleted: true };
}

// ====== נוכחות ======

/** upsert של נוכחות למפגש: { date, groupId, markedBy, records:[{memberId,status}] } */
function saveAttendance_(p) {
  if (!p.date || !p.groupId) throw new Error('חסר תאריך או קבוצה');
  var records = p.records || [];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_('Attendance'), values = sh.getDataRange().getValues();
    for (var r = values.length - 1; r >= 1; r--) {
      if (fmtDateGs_(values[r][0]) === String(p.date) && String(values[r][1]) === String(p.groupId)) sh.deleteRow(r + 1);
    }
    var ts = new Date().toISOString();
    var rows = records.map(function (rec) { return [p.date, p.groupId, rec.memberId, rec.status, p.markedBy || '', ts]; });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    return { date: p.date, groupId: String(p.groupId), saved: rows.length };
  } finally {
    lock.releaseLock();
  }
}

/** שליפת היסטוריה: { from, to, groupId, memberId } */
function getHistory_(p) {
  p = p || {};
  return readAll_('Attendance').map(function (a) {
    return {
      date: fmtDateGs_(a.date), groupId: String(a.groupId), memberId: String(a.memberId),
      status: String(a.status), markedBy: String(a.markedBy || ''), timestamp: String(a.timestamp || '')
    };
  }).filter(function (a) {
    if (p.from && a.date < String(p.from)) return false;
    if (p.to && a.date > String(p.to)) return false;
    if (p.groupId && a.groupId !== String(p.groupId)) return false;
    if (p.memberId && a.memberId !== String(p.memberId)) return false;
    return true;
  });
}

// ====== סטטוס מפגש (ביטול/החזרה) ======

/** ביטול מפגש: { date, groupId, note, by } */
function cancelSession_(p) {
  if (!p.date || !p.groupId) throw new Error('חסר תאריך או קבוצה');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_('Sessions'), vals = sh.getDataRange().getValues();
    for (var r = vals.length - 1; r >= 1; r--) {
      if (fmtDateGs_(vals[r][0]) === String(p.date) && String(vals[r][1]) === String(p.groupId)) sh.deleteRow(r + 1);
    }
    sh.appendRow([p.date, p.groupId, 'canceled', p.note || '', p.by || '', new Date().toISOString(), p.makeup || '']);
    return { date: p.date, groupId: String(p.groupId), status: 'canceled', note: p.note || '', makeup: p.makeup || '' };
  } finally {
    lock.releaseLock();
  }
}

/** ביטול הביטול: { date, groupId } */
function restoreSession_(p) {
  if (!p.date || !p.groupId) throw new Error('חסר תאריך או קבוצה');
  var sh = sheet_('Sessions'), vals = sh.getDataRange().getValues();
  for (var r = vals.length - 1; r >= 1; r--) {
    if (fmtDateGs_(vals[r][0]) === String(p.date) && String(vals[r][1]) === String(p.groupId)) sh.deleteRow(r + 1);
  }
  return { date: p.date, groupId: String(p.groupId), restored: true };
}

/** שליפת סטטוסי מפגשים: { from, to, groupId } */
function getSessions_(p) {
  p = p || {};
  return readAll_('Sessions').map(function (a) {
    return {
      date: fmtDateGs_(a.date), groupId: String(a.groupId), status: String(a.status),
      note: String(a.note || ''), markedBy: String(a.markedBy || ''), timestamp: String(a.timestamp || ''),
      makeup: a.makeup ? fmtDateGs_(a.makeup) : ''
    };
  }).filter(function (a) {
    if (p.from && a.date < String(p.from)) return false;
    if (p.to && a.date > String(p.to)) return false;
    if (p.groupId && a.groupId !== String(p.groupId)) return false;
    return true;
  });
}

// ====== היעדרויות מתאמנים (דיווח מראש) ======

/** דיווח היעדרות: { date, groupId, memberId, note, by } */
function reportAbsence_(p) {
  if (!p.date || !p.groupId || !p.memberId) throw new Error('חסרים פרטים');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_('Absences'), vals = sh.getDataRange().getValues();
    for (var r = vals.length - 1; r >= 1; r--) {
      if (fmtDateGs_(vals[r][0]) === String(p.date) && String(vals[r][1]) === String(p.groupId) && String(vals[r][2]) === String(p.memberId)) sh.deleteRow(r + 1);
    }
    sh.appendRow([p.date, p.groupId, p.memberId, p.note || '', p.by || '', new Date().toISOString()]);
    return { date: p.date, groupId: String(p.groupId), memberId: String(p.memberId), reported: true };
  } finally {
    lock.releaseLock();
  }
}

/** ביטול דיווח היעדרות: { date, groupId, memberId } */
function unreportAbsence_(p) {
  if (!p.date || !p.groupId || !p.memberId) throw new Error('חסרים פרטים');
  var sh = sheet_('Absences'), vals = sh.getDataRange().getValues();
  for (var r = vals.length - 1; r >= 1; r--) {
    if (fmtDateGs_(vals[r][0]) === String(p.date) && String(vals[r][1]) === String(p.groupId) && String(vals[r][2]) === String(p.memberId)) sh.deleteRow(r + 1);
  }
  return { removed: true };
}

/** שליפת היעדרויות: { from, to, groupId, memberId } */
function getAbsences_(p) {
  p = p || {};
  return readAll_('Absences').map(function (a) {
    return {
      date: fmtDateGs_(a.date), groupId: String(a.groupId), memberId: String(a.memberId),
      note: String(a.note || ''), by: String(a.by || ''), timestamp: String(a.timestamp || '')
    };
  }).filter(function (a) {
    if (p.from && a.date < String(p.from)) return false;
    if (p.to && a.date > String(p.to)) return false;
    if (p.groupId && a.groupId !== String(p.groupId)) return false;
    if (p.memberId && a.memberId !== String(p.memberId)) return false;
    return true;
  });
}

// ====== משתמשים / התחברות (Users) ======

function findByName_(sheetName, name) {
  var n = String(name).trim();
  return readAll_(sheetName).filter(function (x) { return String(x.name).trim() === n; })[0] || null;
}

// נרמול שם + זיהוי "חן" סובלני (רווחים / מילים נוספות כמו "רויזמן"), התאמה לפי מילה שלמה
function normName_(s) { return String(s == null ? '' : s).trim().replace(/\s+/g, ' '); }
function isChen_(name) { return normName_(name).split(' ').indexOf('חן') >= 0; }

/** זריעה חד-פעמית של משתמשי ברירת המחדל (מקושרים למאמנים/חניכים קיימים, בלי למחוק נתונים). */
function seedUsers_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('users_seeded_v3') === '1') return;
  var ush = sheet_('Users');
  var users = readAll_('Users');
  function hasUser(name, role) { return users.some(function (u) { return String(u.name).trim() === name && String(u.role) === role; }); }

  if (!hasUser('מנהל', 'manager')) ush.appendRow([newId_('u'), 'מנהל', 'manager', 'Admin2026', '']);

  [['טל', 'Tal1406'], ['אריה', 'Lion2026']].forEach(function (pair) {
    if (hasUser(pair[0], 'coach')) return;
    var coach = findByName_('Coaches', pair[0]);
    var cid = coach ? String(coach.id) : newId_('c');
    if (!coach) sheet_('Coaches').appendRow([cid, pair[0]]);
    ush.appendRow([newId_('u'), pair[0], 'coach', pair[1], cid]);
  });

  // ----- חן: קישור למתאמן קיים בשם "חן" (סובלני), Chen2026, ללא כפילויות -----
  var chenMems = readAll_('Members').filter(function (m) { return isChen_(m.name); });
  // שומר מועדף: מי ששמו בדיוק "חן"; אחרת הראשון; אם אין בכלל — ניצור "חן"
  var keeper = chenMems.filter(function (m) { return normName_(m.name) === 'חן'; })[0] || chenMems[0] || null;
  if (!keeper) {
    var kid = newId_('m'); sheet_('Members').appendRow([kid, 'חן', '', '']); keeper = { id: kid, name: 'חן' };
  }
  var keeperId = String(keeper.id);
  // מחיקת חברי־"חן" כפולים (כמו "חן רויזמן" מזריעה קודמת) — מוחקים רק אינם ה-keeper, כולל השיוכים שלהם
  chenMems.forEach(function (m) {
    if (String(m.id) === keeperId) return;
    deleteRowsWhere_('Memberships', 'memberId', m.id);
    var mr = findRowById_('Members', m.id); if (mr > 0) sheet_('Members').deleteRow(mr);
  });
  // משתמשי־מתאמן של "חן": שומרים אחד ומקשרים ל-keeper עם Chen2026, מוחקים את השאר
  var chenUsers = readAll_('Users').filter(function (u) { return String(u.role) === 'trainee' && isChen_(u.name); });
  if (chenUsers.length) {
    var keep = chenUsers[0];
    chenUsers.slice(1).forEach(function (u) { var ur = findRowById_('Users', u.id); if (ur > 0) sheet_('Users').deleteRow(ur); });
    var kr = findRowById_('Users', keep.id);
    if (kr > 0) {
      var ush2 = sheet_('Users');
      ush2.getRange(kr, 2, 1, 1).setValue(keeper.name);   // שם
      ush2.getRange(kr, 4, 1, 1).setValue('Chen2026');    // סיסמה
      ush2.getRange(kr, 5, 1, 1).setValue(keeperId);      // refId
    }
  } else {
    ush.appendRow([newId_('u'), keeper.name, 'trainee', 'Chen2026', keeperId]);
  }

  props.setProperty('users_seeded_v3', '1');
}

/** אימות התחברות. { name, mode:'staff'|'trainee', password } → זהות המשתמש או שגיאה. */
function login_(p) {
  var name = String(p.name || '').trim();
  if (!name) throw new Error('נא לבחור שם');
  var mode = String(p.mode || 'staff');
  var pass = String(p.password == null ? '' : p.password);
  var roles = (mode === 'trainee') ? ['trainee'] : ['manager', 'coach'];
  var u = readAll_('Users').filter(function (x) { return String(x.name).trim() === name && roles.indexOf(String(x.role)) >= 0; })[0];
  if (!u) throw new Error('משתמש לא נמצא');
  var stored = String(u.password == null ? '' : u.password);
  if (stored !== '' && stored !== pass) throw new Error('סיסמה שגויה');
  var out = { name: u.name, role: String(u.role) };
  if (u.role === 'coach') out.coachId = String(u.refId || '');
  if (u.role === 'trainee') out.memberId = String(u.refId || '');
  return out;
}

/** רשימת משתמשים ללא סיסמאות (רק דגל hasPassword). */
function getUsers_() {
  return readAll_('Users').map(function (u) {
    return { id: String(u.id), name: u.name, role: String(u.role), refId: String(u.refId || ''), hasPassword: String(u.password == null ? '' : u.password) !== '' };
  });
}

/** הוספת משתמש. { name, role, password, refId? } — יוצר מאמן/חניך אם צריך. */
function addUser_(p) {
  if (!p.name) throw new Error('חסר שם');
  var role = String(p.role || 'trainee');
  var refId = String(p.refId || '');
  if (role === 'coach' && !refId) { refId = newId_('c'); sheet_('Coaches').appendRow([refId, p.name]); }
  if (role === 'trainee' && !refId) { refId = newId_('m'); sheet_('Members').appendRow([refId, p.name, '', '']); }
  var id = newId_('u');
  sheet_('Users').appendRow([id, p.name, role, p.password || '', refId]);
  return { id: id, name: p.name, role: role, refId: refId, hasPassword: !!p.password };
}

/** קביעת/איפוס סיסמה. { id, password } (ריק = ללא סיסמה) */
function setUserPassword_(p) {
  var row = findRowById_('Users', p.id);
  if (row < 0) throw new Error('משתמש לא נמצא');
  sheet_('Users').getRange(row, 4, 1, 1).setValue(p.password || '');
  return { id: p.id, hasPassword: !!p.password };
}

function deleteUser_(p) {
  var row = findRowById_('Users', p.id);
  if (row < 0) throw new Error('משתמש לא נמצא');
  sheet_('Users').deleteRow(row);
  return { id: p.id, deleted: true };
}

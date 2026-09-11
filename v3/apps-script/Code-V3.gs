/**
 * MEMBER REGISTER V3 — Apps Script backend  (schema-corrected build)
 * Cornerstone Church Hua Hin
 * SCHEMA_VERSION 3.0 · BACKEND_BUILD v3-dev-02
 *
 * Script Properties required:
 *   V3_SPREADSHEET_ID   17j5hFBtoUk4gxz6ykYLKA7ve2fw9VUE8h_PpzPXdFYM
 *   SCHEMA_VERSION      3.0
 *   BACKEND_BUILD       v3-dev-02
 *   TOKEN_PEPPER        (random string — session/password hashing)
 *
 * This build reads the ACTUAL current sheet schema:
 *   - Settings uses SettingKey / SettingValue
 *   - Members / Attendance / Decisions may still be LEGACY-shaped (Thai headers)
 *   - EngagementLog, UnmatchedRecords, ImportLog use their real column names
 *
 * Two validation stages:
 *   validateLegacySchema_()  — enough to RUN MIGRATIONS
 *   validateRuntimeSchema_() — strict normalized V3, required for normal operation
 * Migrations are never blocked by runtime validation.
 */

var TZ = 'Asia/Bangkok';

/* ============================================================
 * 1. SPREADSHEET ACCESS
 * ============================================================ */

function props_() { return PropertiesService.getScriptProperties(); }

function getV3Spreadsheet_() {
  var id = props_().getProperty('V3_SPREADSHEET_ID');
  if (!id) throw new Error('SCHEMA: V3_SPREADSHEET_ID is not set in Script Properties.');
  return SpreadsheetApp.openById(id);
}

function getV3Sheet_(name) {
  var sh = getV3Spreadsheet_().getSheetByName(name);
  if (!sh) throw new Error('SCHEMA: missing sheet "' + name + '"');
  return sh;
}

var _tableCache = {};
function readTable_(name, force) {
  if (!force && _tableCache[name]) return _tableCache[name];
  var sh = getV3Sheet_(name);
  var lastCol = Math.max(1, sh.getLastColumn());
  var lastRow = Math.max(1, sh.getLastRow());
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (r.join('') === '') continue;
    var o = { _row: i + 1 };
    for (var c = 0; c < headers.length; c++)
      if (headers[c] && !(headers[c] in o)) o[headers[c]] = r[c];   // duplicate header: first wins
    rows.push(o);
  }
  return (_tableCache[name] = { name: name, headers: headers, rows: rows, sheet: sh });
}

function invalidateCache_(name) { if (name) delete _tableCache[name]; else _tableCache = {}; _mode = null; }

function hasHeader_(t, name) { return t.headers.indexOf(name) >= 0; }

/** First header present from a candidate list, else null. */
function pickHeader_(t, candidates) {
  for (var i = 0; i < candidates.length; i++) if (hasHeader_(t, candidates[i])) return candidates[i];
  return null;
}

/** Read a value from a row using header aliases (V3 name first, legacy names after). */
function field_(row, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var k = candidates[i];
    if (row[k] !== undefined && String(row[k]).trim() !== '') return row[k];
  }
  return '';
}

/** Ensure a column exists; creates it at the right edge when allowed.
 *  Duplicate legacy headers are left untouched — indexOf resolves to the first. */
function ensureColumn_(t, name, create) {
  var i = t.headers.indexOf(name);
  if (i >= 0) return i + 1;
  if (!create) return -1;
  var col = t.headers.length + 1;
  t.sheet.getRange(1, col).setValue(name);
  t.headers.push(name);
  return col;
}

/* ============================================================
 * 2. HEADER ALIAS MAP — legacy compatibility layer
 * ============================================================ */

var ALIAS = {
  Members: {
    MemberID:   ['MemberID', 'รหัส', 'รหัสสมาชิก', 'memberKey'],
    FullName:   ['FullName', 'ชื่อ สกุล', 'ชื่อ-สกุล', 'ชื่อสกุล', 'ชื่อ'],
    StatusRaw:  ['StatusCode', 'สถานะ', 'สถานภาพ', 'Status'],
    BirthDate:  ['BirthDate', 'วันเกิด', 'DOB', 'วันเดือนปีเกิด'],
    Gender:     ['Gender', 'เพศ', 'Sex'],
    Nickname:   ['Nickname', 'ชื่อเล่น'],
    Phone:      ['Phone', 'โทรศัพท์', 'เบอร์โทร', 'โทร', 'Tel'],
    Email:      ['Email', 'อีเมล'],
    Address:    ['Address', 'ที่อยู่'],
    Note:       ['Note', 'หมายเหตุ', 'หมายเหต'],
    JoinDate:   ['JoinDate', 'วันที่เข้าเป็นสมาชิก', 'วันเข้าร่วม']
  },
  Attendance: {
    MemberID: ['MemberID', 'รหัส', 'รหัสสมาชิก'],
    Name:     ['MemberName', 'ชื่อ สกุล', 'ชื่อ-สกุล', 'ชื่อสกุล', 'Name', 'ชื่อ'],
    Date:     ['AttendanceDate', 'Date', 'วันที่', 'วันที่นมัสการ', 'Timestamp', 'วันเวลา']
  },
  /* NOTE: the legacy Attendance date header is not a recognisable name, so the
     date column is resolved by detectAttendanceDateColumn_() (value sampling),
     never by header alias alone. */
  Decisions: {
    MemberID:     ['MemberID', 'memberKey'],
    MemberName:   ['MemberNameSnapshot', 'memberName'],
    Action:       ['Action', 'decision'],
    DecisionDate: ['DecisionDate', 'at'],
    RecordedBy:   ['RecordedBy', 'by'],
    Note:         ['Note', 'note'],
    ReasonCodes:  ['ReasonCodes'],
    NextReview:   ['NextReviewDate'],
    Active:       ['Active']
  },
  Giving: {
    MemberID: ['MemberID', 'รหัส'],
    Name:     ['MemberName', 'ชื่อ สกุล', 'ชื่อ-สกุล'],
    Date:     ['GivingDate', 'Date', 'วันที่'],
    Amount:   ['Amount', 'จำนวนเงิน', 'ยอด']
  },
  Accounts: {
    Username:     ['Username', 'user'],
    PasswordHash: ['PasswordHash'],
    LegacyHash:   ['passHash'],
    Salt:         ['salt'],
    Role:         ['Role', 'role'],
    Active:       ['Active', 'active'],
    DisplayName:  ['DisplayName'],
    UpdatedAt:    ['UpdatedAt', 'updatedAt', 'updated'],
    Note:         ['Note', 'note']
  }
};

function mField_(row, key)  { return field_(row, ALIAS.Members[key]); }
function aField_(row, key)  { return field_(row, ALIAS.Attendance[key]); }
function dField_(row, key)  { return field_(row, ALIAS.Decisions[key]); }
function accField_(row, key) { return field_(row, ALIAS.Accounts[key]); }

/* ------------------------------------------------------------
 * Attendance date-column detection (header alias → value sampling)
 * ------------------------------------------------------------ */

/** Strictly date-shaped? Avoids treating IDs/numbers as dates. */
function looksLikeDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return !isNaN(v);
  var s = String(v == null ? '' : v).trim();
  if (!s) return false;
  if (!/^\d{1,4}[-\/.]\d{1,2}[-\/.]\d{1,4}([ T]\d{1,2}:\d{2}(:\d{2})?)?$/.test(s)) return false;
  return !!parseDate_(s);
}

var MIN_DATE_SAMPLE = 5;
var DATE_CONFIDENCE = 0.9;

var _attDateCol = null;
/**
 * Resolves the real Attendance date column.
 *   { header, method: 'ALIAS'|'SAMPLED'|'NONE'|'AMBIGUOUS', confidence, candidates[] }
 * Never guesses: two or more equally date-like columns → AMBIGUOUS with the list.
 */
function detectAttendanceDateColumn_(force) {
  if (_attDateCol && !force) return _attDateCol;
  var t = readTable_('Attendance', force);

  var alias = pickHeader_(t, ALIAS.Attendance.Date);
  if (alias) return (_attDateCol = { header: alias, method: 'ALIAS', confidence: 1,
                                     candidates: [{ header: alias, ratio: 1 }] });

  var sample = t.rows.slice(0, 300);
  var scored = [];
  t.headers.forEach(function (h) {
    if (!h) return;
    if (ALIAS.Attendance.Name.indexOf(h) >= 0 || ALIAS.Attendance.MemberID.indexOf(h) >= 0) return;
    var total = 0, ok = 0;
    sample.forEach(function (r) {
      var v = r[h];
      if (v === '' || v == null) return;
      total++;
      if (looksLikeDate_(v)) ok++;
    });
    if (total >= MIN_DATE_SAMPLE) scored.push({ header: h, sampled: total, dateLike: ok, ratio: ok / total });
  });
  scored.sort(function (a, b) { return b.ratio - a.ratio || b.sampled - a.sampled; });
  var strong = scored.filter(function (s) { return s.ratio >= DATE_CONFIDENCE; });

  if (!strong.length)
    return (_attDateCol = { header: null, method: 'NONE', confidence: 0, candidates: scored });
  if (strong.length > 1)
    return (_attDateCol = { header: null, method: 'AMBIGUOUS', confidence: strong[0].ratio,
                            candidates: strong });
  return (_attDateCol = { header: strong[0].header, method: 'SAMPLED',
                          confidence: strong[0].ratio, candidates: scored.slice(0, 5) });
}

function requireAttendanceDateColumn_() {
  var d = detectAttendanceDateColumn_();
  if (d.header) return d.header;
  var list = (d.candidates || []).map(function (c) {
    return c.header + ' (' + Math.round((c.ratio || 0) * 100) + '%)';
  }).join(', ');
  throw new Error('ATTENDANCE: ' + (d.method === 'AMBIGUOUS'
    ? 'พบคอลัมน์วันที่มากกว่าหนึ่งคอลัมน์ — ไม่เดา'
    : 'ไม่พบคอลัมน์วันที่ที่เชื่อถือได้') + ' · ผู้สมัคร: [' + (list || 'ไม่มี') + ']');
}

/* ============================================================
 * 3. SCHEMA VALIDATION — TWO STAGES
 * ============================================================ */

/** Stage 1 — the minimum needed for MIGRATIONS to run against legacy data. */
var LEGACY_SCHEMA = {
  Members:    { anyOf: [ALIAS.Members.MemberID, ALIAS.Members.FullName] },
  Attendance: { anyOf: [ALIAS.Attendance.Name.concat(ALIAS.Attendance.MemberID)] },
  Decisions:  { anyOf: [ALIAS.Decisions.MemberID.concat(ALIAS.Decisions.MemberName), ALIAS.Decisions.Action] },
  Settings:   { anyOf: [['SettingKey', 'Key'], ['SettingValue', 'Value']] },
  StatusTypes:{ anyOf: [['StatusCode'], ['DisplayName']] }
};

/** Stage 2 — strict normalized V3, required for the live app. */
var RUNTIME_SCHEMA = {
  Members:         ['MemberID', 'FullName', 'StatusCode'],
  Attendance:      ['MemberID'],
  Giving:          [],
  EngagementLog:   ['LogID', 'MemberID', 'EngagementDate', 'EngagementType', 'Active'],
  StatusTypes:     ['StatusCode', 'DisplayName'],
  MembershipMeta:  ['MemberID'],
  Decisions:       ['DecisionID', 'MemberID', 'Action', 'DecisionDate', 'Active'],
  DecisionHistory: ['HistoryID', 'DecisionID', 'MemberID', 'Action'],
  Accounts:        ['Username', 'PasswordHash', 'Role', 'Active'],
  AuditLog:        [],
  Settings:        ['SettingKey', 'SettingValue'],
  ImportLog:       ['ImportID', 'Status'],
  UnmatchedRecords:['UnmatchedID', 'EntityType', 'MatchStatus']
};

function sheetNames_() {
  return getV3Spreadsheet_().getSheets().map(function (s) { return s.getName(); });
}

function validateLegacySchema_() {
  var present = sheetNames_(), problems = [];
  Object.keys(LEGACY_SCHEMA).forEach(function (name) {
    if (present.indexOf(name) < 0) { problems.push('missing sheet: ' + name); return; }
    var t = readTable_(name);
    LEGACY_SCHEMA[name].anyOf.forEach(function (group) {
      if (!pickHeader_(t, group)) problems.push(name + ': needs one of [' + group.join(', ') + ']');
    });
  });
  var att = null;
  if (present.indexOf('Attendance') >= 0) {
    att = detectAttendanceDateColumn_(true);
    if (!att.header) problems.push('Attendance: ไม่สามารถระบุคอลัมน์วันที่ได้ (' + att.method + ')');
  }
  return { stage: 'LEGACY', ok: problems.length === 0, problems: problems, attendanceDate: att };
}

function validateRuntimeSchema_() {
  var present = sheetNames_(), problems = [];
  Object.keys(RUNTIME_SCHEMA).forEach(function (name) {
    if (present.indexOf(name) < 0) { problems.push('missing sheet: ' + name); return; }
    var t = readTable_(name);
    RUNTIME_SCHEMA[name].forEach(function (h) {
      if (!hasHeader_(t, h)) problems.push(name + ': missing column "' + h + '"');
    });
  });
  return {
    stage: 'RUNTIME', ok: problems.length === 0, problems: problems,
    schemaVersion: props_().getProperty('SCHEMA_VERSION') || null,
    backendBuild: props_().getProperty('BACKEND_BUILD') || null
  };
}

/** Runtime ops fail closed. Migrations must NOT call this. */
function assertRuntimeSchema_() {
  var v = validateRuntimeSchema_();
  if (!v.ok) throw new Error('RUNTIME SCHEMA NOT READY — ' + v.problems.join(' | ') +
    ' · run the migrations first (migrateLegacyStatuses → migrateAttendanceToMemberID → migrateLegacyDecisions).');
  return v;
}

/**
 * COMPATIBILITY vs NORMALIZED.
 * The web app runs in either mode — compatibility is NOT an error, it simply means
 * the legacy-shaped sheets are being read through the alias layer and members
 * without a permanent MemberID use a temporary in-memory row key.
 */
var _mode = null;
function runtimeMode_(force) {
  if (_mode && !force) return _mode;
  var v = validateRuntimeSchema_();
  return (_mode = {
    mode: v.ok ? 'NORMALIZED' : 'COMPATIBILITY',
    normalized: v.ok,
    pending: v.problems
  });
}

/** Writing a decision needs the V3 decision columns; reading does not. */
function assertDecisionsWritable_() {
  var t = readTable_('Decisions');
  var missing = RUNTIME_SCHEMA.Decisions.filter(function (h) { return !hasHeader_(t, h); });
  if (missing.length) throw new Error('DECISIONS ยังไม่พร้อมบันทึกมติ — ขาดคอลัมน์: ' +
    missing.join(', ') + ' · ให้รัน migrateLegacyDecisions(false) ก่อน');
  return true;
}

/** Migrations fail closed on their own, lighter, requirement. */
function assertLegacySchema_() {
  var v = validateLegacySchema_();
  if (!v.ok) throw new Error('LEGACY SCHEMA INVALID — ' + v.problems.join(' | '));
  return v;
}

/* ============================================================
 * 4. SETTINGS  (SettingKey / SettingValue / Active)
 * ============================================================ */

var SETTING_DEFAULTS = {
  SYSTEM_NAME: 'ระบบทะเบียนสมาชิก',
  TIMEZONE: TZ,
  LOCALE: 'th-TH',
  COUNCIL_REVIEW_YEARS: 1,
  OUTSIDE_COUNCIL_REVIEW_YEARS: 2,
  FULL2_ABSENCE_MONTHS: 12,
  FULL2_REVIEW_MONTHS: 36,
  FOLLOWUP_WARNING_MONTHS: 6,
  STRONG_CONTINUITY_PERCENT: 50,
  STRONG_MIN_VISITS_12M: 10,
  STRONG_MIN_MEMBERSHIP_YEARS: 1,
  DEFAULT_NEXT_REVIEW_MONTHS: 12,
  REQUIRE_DECISION_REASON: true,
  GIVING_AMOUNT_VISIBLE: false,
  MEMBERS_SOURCE_MODE: 'V3',
  SUNDAY_SCHOOL_MIN_AGE: 3,
  SUNDAY_SCHOOL_MAX_AGE: 10,
  YOUTH_MIN_AGE: 11,
  YOUTH_MAX_AGE: 24,
  ADULT_MIN_AGE: 25,
  ADULT_MAX_AGE: 59,
  SENIOR_MIN_AGE: 60
};

function settingsCols_(t) {
  return {
    key: pickHeader_(t, ['SettingKey', 'Key']),
    value: pickHeader_(t, ['SettingValue', 'Value'])
  };
}

var _settings = null;
function loadSettings_(force) {
  if (_settings && !force) return _settings;
  var t = readTable_('Settings', force);
  var c = settingsCols_(t);
  if (!c.key || !c.value) throw new Error('SCHEMA: Settings needs SettingKey / SettingValue');
  var out = {};
  Object.keys(SETTING_DEFAULTS).forEach(function (k) { out[k] = SETTING_DEFAULTS[k]; });
  t.rows.forEach(function (r) {
    var k = String(r[c.key] || '').trim();
    if (!k) return;
    if (hasHeader_(t, 'Active') && r.Active !== '' && r.Active != null && !truthy_(r.Active)) return;
    out[k] = coerceSetting_(r[c.value], r.ValueType);
  });
  return (_settings = out);
}

function coerceSetting_(v, type) {
  var s = String(v).trim();
  var t = String(type || '').toUpperCase();
  if (t === 'NUMBER' || t === 'INTEGER' || (!t && s !== '' && !isNaN(s))) return Number(s);
  if (t === 'BOOLEAN' || /^(true|false)$/i.test(s)) return /^true$/i.test(s);
  return v;
}

function getSetting_(key) {
  var s = loadSettings_();
  return (key in s) ? s[key] : (key in SETTING_DEFAULTS ? SETTING_DEFAULTS[key] : null);
}

/** Validate age bands from an arbitrary settings object (used for pre-write checks). */
function validateAgeBandsFrom_(s) {
  var bands = [
    { code: 'INFANT',        min: 0,                               max: Number(s.SUNDAY_SCHOOL_MIN_AGE) - 1 },
    { code: 'SUNDAY_SCHOOL', min: Number(s.SUNDAY_SCHOOL_MIN_AGE), max: Number(s.SUNDAY_SCHOOL_MAX_AGE) },
    { code: 'YOUTH',         min: Number(s.YOUTH_MIN_AGE),         max: Number(s.YOUTH_MAX_AGE) },
    { code: 'ADULT',         min: Number(s.ADULT_MIN_AGE),         max: Number(s.ADULT_MAX_AGE) },
    { code: 'SENIOR',        min: Number(s.SENIOR_MIN_AGE),        max: Infinity }
  ];
  var problems = [];
  bands.forEach(function (b) {
    if (isNaN(b.min) || isNaN(b.max)) problems.push(b.code + ': ค่าไม่ใช่ตัวเลข');
  });
  for (var i = 0; i < bands.length; i++) {
    if (bands[i].min > bands[i].max) problems.push(bands[i].code + ': min > max');
    if (i > 0) {
      var gap = bands[i].min - bands[i - 1].max;
      if (gap > 1) problems.push('ช่วงอายุขาดหาย ระหว่าง ' + bands[i - 1].code + ' และ ' + bands[i].code);
      if (gap < 1) problems.push('ช่วงอายุซ้อนทับ ระหว่าง ' + bands[i - 1].code + ' และ ' + bands[i].code);
    }
  }
  return { ok: problems.length === 0, problems: problems, bands: bands };
}

function validateAgeBands_() { return validateAgeBandsFrom_(loadSettings_()); }

/* ============================================================
 * 5. STATUS TYPES (label master)
 * ============================================================ */

var _statusTypes = null;
function loadStatusTypes_(force) {
  if (_statusTypes && !force) return _statusTypes;
  var t = readTable_('StatusTypes', force);
  var map = {};
  t.rows.forEach(function (r) {
    var code = String(r.StatusCode || '').trim();
    if (!code) return;
    map[code] = {
      StatusCode: code,
      DisplayName: r.DisplayName || code,
      ShortName: r.ShortName || r.DisplayName || code,
      StatusClass: r.StatusClass || '',
      Active: truthy_(r.Active),
      CanVote: truthy_(r.CanVote),
      CanNominate: truthy_(r.CanNominate),
      CountAsCompletedMember: truthy_(r.CountAsCompletedMember),
      RegistrarGroup: r.RegistrarGroup || '',
      DashboardVisible: truthy_(r.DashboardVisible),
      SortOrder: Number(r.SortOrder) || 0,
      ColorHex: r.ColorHex || '',
      RequiresBoardDecision: truthy_(r.RequiresBoardDecision),
      Description: r.Description || ''
    };
  });
  return (_statusTypes = map);
}

function truthy_(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'ใช่';
}

function statusLabel_(code) {
  var st = loadStatusTypes_()[code];
  return st ? st.DisplayName : code;
}

/* ============================================================
 * 6. LEGACY STATUS NORMALIZATION
 * ============================================================ */

var LEGACY_STATUS_MAP = {
  'สมบูรณ์ 1': 'FULL1', 'สมบูรณ์1': 'FULL1',
  'สมาชิกสมบูรณ์ 1': 'FULL1', 'สมาชิกสมบูรณ์1': 'FULL1',
  'สมบูรณ์ 2': 'FULL2', 'สมบูรณ์2': 'FULL2',
  'สมาชิกสมบูรณ์ 2': 'FULL2', 'สมาชิกสมบูรณ์2': 'FULL2',
  'สำรอง': 'RESERVE', 'สมาชิกสำรอง': 'RESERVE',
  'ผู้ร่วมนมัสการ': 'ATTENDER', 'ผู้ร่วมนมัสการ/ยังไม่ระบุสถานภาพ': 'ATTENDER',
  'ลาออก': 'RESIGNED', 'ย้ายสังกัด': 'TRANSFERRED',
  'ให้ออกโดยมติธรรมกิจ': 'REMOVED', 'เสียชีวิต': 'DECEASED',
  'รอพิจารณาเป็นสมาชิกสมบูรณ์': 'PENDING_REVIEW'
};

function normalizeLegacyStatus_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return 'UNSPECIFIED';
  if (loadStatusTypes_()[s]) return s;
  var key = s.replace(/\s+/g, ' ');
  if (LEGACY_STATUS_MAP[key]) return LEGACY_STATUS_MAP[key];
  var compact = key.replace(/\s+/g, '');
  if (LEGACY_STATUS_MAP[compact]) return LEGACY_STATUS_MAP[compact];
  return 'UNSPECIFIED';
}

/* ============================================================
 * 7. DATES & AGE
 * ============================================================ */

function parseDate_(v) {
  if (!v && v !== 0) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v) ? null : v;
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) {
    var yy = Number(m[1]); if (yy > 2400) yy -= 543;
    return new Date(yy, Number(m[2]) - 1, Number(m[3]));
  }
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
  if (m) {
    var y = Number(m[3]); if (y > 2400) y -= 543;
    return new Date(y, Number(m[2]) - 1, Number(m[1]));
  }
  var d = new Date(s);
  return isNaN(d) ? null : d;
}

function today_() { return new Date(Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd')); }
function isoDate_(d) { return d ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : ''; }
function nowIso_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

function calculateAge_(birth, ref) {
  var b = parseDate_(birth);
  if (!b) return null;
  var t = ref || today_();
  var age = t.getFullYear() - b.getFullYear();
  var m = t.getMonth() - b.getMonth();
  if (m < 0 || (m === 0 && t.getDate() < b.getDate())) age--;
  return (age < 0 || age > 130) ? null : age;
}

function monthsBetween_(from, to) {
  if (!from) return null;
  var t = to || today_();
  var m = (t.getFullYear() - from.getFullYear()) * 12 + (t.getMonth() - from.getMonth());
  if (t.getDate() < from.getDate()) m--;
  return m;
}

/* ============================================================
 * 8. AGE / MINISTRY GROUP ENGINE
 * ============================================================ */

var GROUP_LABELS = {
  INFANT: 'เด็กเล็ก',
  SUNDAY_SCHOOL: 'เด็กรวี / รวีวารศึกษา',
  YOUTH: 'อนุชน',
  ADULT: 'ผู้ใหญ่',
  ADULT_MALE: 'บุรุษ',
  ADULT_FEMALE: 'สตรี',
  ADULT_UNKNOWN: 'ผู้ใหญ่ · ยังไม่ระบุกลุ่ม',
  SENIOR: 'ผู้สูงวัย',
  UNKNOWN: 'ยังไม่ระบุกลุ่มอายุ'
};

function normalizeGender_(g) {
  var s = String(g == null ? '' : g).trim().toLowerCase();
  if (!s) return null;
  if (['m', 'male', 'ช', 'ชาย', 'นาย'].indexOf(s) >= 0) return 'M';
  if (['f', 'female', 'ญ', 'หญิง', 'นาง', 'นางสาว'].indexOf(s) >= 0) return 'F';
  return null;                                   // never guess gender from a Thai name
}

function calculateAgeGroup_(age, gender) {
  if (age == null) return { code: 'UNKNOWN', label: GROUP_LABELS.UNKNOWN, ageBand: 'UNKNOWN' };
  var s = loadSettings_();
  if (age < Number(s.SUNDAY_SCHOOL_MIN_AGE))
    return { code: 'INFANT', label: GROUP_LABELS.INFANT, ageBand: 'INFANT' };
  if (age <= Number(s.SUNDAY_SCHOOL_MAX_AGE))
    return { code: 'SUNDAY_SCHOOL', label: GROUP_LABELS.SUNDAY_SCHOOL, ageBand: 'SUNDAY_SCHOOL' };
  if (age >= Number(s.YOUTH_MIN_AGE) && age <= Number(s.YOUTH_MAX_AGE))
    return { code: 'YOUTH', label: GROUP_LABELS.YOUTH, ageBand: 'YOUTH' };
  if (age >= Number(s.SENIOR_MIN_AGE))
    return { code: 'SENIOR', label: GROUP_LABELS.SENIOR, ageBand: 'SENIOR' };
  if (age >= Number(s.ADULT_MIN_AGE) && age <= Number(s.ADULT_MAX_AGE)) {
    var g = normalizeGender_(gender);
    if (g === 'M') return { code: 'ADULT_MALE', label: GROUP_LABELS.ADULT_MALE, ageBand: 'ADULT' };
    if (g === 'F') return { code: 'ADULT_FEMALE', label: GROUP_LABELS.ADULT_FEMALE, ageBand: 'ADULT' };
    return { code: 'ADULT_UNKNOWN', label: GROUP_LABELS.ADULT_UNKNOWN, ageBand: 'ADULT' };
  }
  return { code: 'UNKNOWN', label: GROUP_LABELS.UNKNOWN, ageBand: 'UNKNOWN' };
}

/* ============================================================
 * 9. INDEXES (alias-aware — work on legacy or normalized sheets)
 * ============================================================ */

function normName_(s) { return String(s == null ? '' : s).trim().replace(/\s+/g, ' '); }

/**
 * Working key for a member row.
 *   EXISTING_ID  — the legacy รหัส / MemberID, the real permanent key
 *   LEGACY_ROW   — 'LEGACY-ROW-<row>', temporary, in memory only, never written
 */
function memberKey_(r) {
  var id = String(mField_(r, 'MemberID')).trim();
  return id ? { key: id, keyType: 'EXISTING_ID' }
            : { key: 'LEGACY-ROW-' + r._row, keyType: 'LEGACY_ROW' };
}

function buildMemberIndex_() {
  var t = readTable_('Members');
  var byId = {}, byKey = {}, byName = {}, dupNames = {};
  t.rows.forEach(function (r) {
    var k = memberKey_(r);
    r._key = k.key; r._keyType = k.keyType;
    byKey[k.key] = r;
    if (k.keyType === 'EXISTING_ID') byId[k.key] = r;
    var n = normName_(mField_(r, 'FullName'));
    if (n) {
      if (byName[n]) dupNames[n] = true;
      byName[n] = r;
    }
  });
  return { table: t, byId: byId, byKey: byKey, byName: byName,
           dupNames: dupNames, rows: t.rows };
}

/**
 * Attendance keyed by the member working key.
 * Uses Attendance.MemberID when present; otherwise falls back to the member's
 * normalized FullName, and ONLY when that name is unique. Ambiguous or unknown
 * names are left unresolved — never guessed.
 */
function buildAttendanceIndex_(memberIdx) {
  var t = readTable_('Attendance');
  var dateHeader = requireAttendanceDateColumn_();
  var idHeader = pickHeader_(t, ALIAS.Attendance.MemberID);
  var nameHeader = pickHeader_(t, ALIAS.Attendance.Name);
  var byMember = {}, unresolved = 0;

  t.rows.forEach(function (r) {
    var d = parseDate_(r[dateHeader]);
    if (!d) return;
    var key = null;
    var id = idHeader ? String(r[idHeader] || '').trim() : '';
    if (id && memberIdx && memberIdx.byId[id]) key = id;
    if (!key && nameHeader && memberIdx) {
      var n = normName_(r[nameHeader]);
      if (n && !memberIdx.dupNames[n] && memberIdx.byName[n]) key = memberIdx.byName[n]._key;
    }
    if (!key) { unresolved++; return; }
    (byMember[key] = byMember[key] || []).push(d);
  });
  Object.keys(byMember).forEach(function (k) { byMember[k].sort(function (a, b) { return b - a; }); });
  byMember._unresolved = unresolved;
  return byMember;
}

function getLastWorshipDate_(attIndex, memberId) {
  var list = attIndex[memberId];
  return list && list.length ? list[0] : null;
}

function countVisitsSince_(attIndex, memberId, months) {
  var list = attIndex[memberId] || [];
  var cut = today_(); cut.setMonth(cut.getMonth() - months);
  return list.filter(function (d) { return d >= cut; }).length;
}

function buildMetaIndex_() {
  var t = readTable_('MembershipMeta');
  var by = {};
  t.rows.forEach(function (r) {
    var id = String(r.MemberID || '').trim();
    if (!id) return;
    if (r.Active === '' || r.Active == null || truthy_(r.Active)) by[id] = r;
  });
  return by;
}

function buildDecisionIndex_() {
  var t = readTable_('Decisions');
  var by = {};
  t.rows.forEach(function (r) {
    var id = String(dField_(r, 'MemberID')).trim();
    if (!id) return;
    if (hasHeader_(t, 'Active') && r.Active !== '' && r.Active != null && !truthy_(r.Active)) return;
    var d = parseDate_(dField_(r, 'DecisionDate'));
    var prev = by[id];
    var pd = prev ? parseDate_(dField_(prev, 'DecisionDate')) : null;
    if (!prev || (d && (!pd || d > pd))) by[id] = r;
  });
  return by;
}

function getActiveDecision_(decIndex, memberId) { return decIndex[memberId] || null; }

/* ============================================================
 * 10. STATUS ENGINE — Base / System / Effective
 * ============================================================ */

function calculateSystemStatus_(ctx, member) {
  var s = ctx.settings;
  var id = member._key || memberKey_(member).key;
  var base = normalizeLegacyStatus_(mField_(member, 'StatusRaw'));
  var last = getLastWorshipDate_(ctx.attendance, id);
  var monthsAbsent = last ? monthsBetween_(last) : null;
  var flags = [], explain = [], system = base;

  if (['TRANSFERRED', 'RESIGNED', 'REMOVED', 'DECEASED'].indexOf(base) >= 0) {
    return { base: base, system: base, monthsAbsent: monthsAbsent, lastWorship: last,
             flags: flags, explanation: ['สถานภาพทะเบียนคงที่: ' + statusLabel_(base)] };
  }

  if (base === 'FULL1' || base === 'FULL1_BOARD' || base === 'FULL2') {
    var limit = Number(s.FULL2_ABSENCE_MONTHS);
    if (monthsAbsent == null) {
      system = (base === 'FULL1_BOARD') ? 'FULL2' : base;
      flags.push('NO_ATTENDANCE_DATA');
      explain.push('ไม่พบข้อมูลการนมัสการ — ยังไม่สามารถคำนวณได้');
    } else if (monthsAbsent >= limit) {
      system = 'FULL2';
      explain.push('ขาดการนมัสการ ' + monthsAbsent + ' เดือน (เกินเกณฑ์ ' + limit + ' เดือน)');
      if (monthsAbsent >= Number(s.FULL2_REVIEW_MONTHS)) {
        flags.push('REVIEW_FOR_RESERVE');
        explain.push('ขาดนมัสการถึงเกณฑ์ ' + s.FULL2_REVIEW_MONTHS +
          ' เดือน · ควรนำเข้าพิจารณาเป็นสมาชิกสำรอง (ต้องมีมติธรรมกิจ ไม่เปลี่ยนอัตโนมัติ)');
      }
    } else {
      system = 'FULL1';
      explain.push('นมัสการล่าสุด ' + monthsAbsent + ' เดือนที่ผ่านมา — อยู่ในเกณฑ์สมาชิกสมบูรณ์ 1');
      if (monthsAbsent >= Number(s.FOLLOWUP_WARNING_MONTHS)) {
        flags.push('FOLLOWUP_6M');
        explain.push('เข้าเกณฑ์ติดตาม ' + s.FOLLOWUP_WARNING_MONTHS +
          ' เดือน (การติดตามเท่านั้น ไม่เปลี่ยนสถานภาพหรือสิทธิ)');
      }
    }
  } else if (['ATTENDER', 'UNSPECIFIED', 'RESERVE', 'PENDING_REVIEW'].indexOf(base) >= 0) {
    var meta = ctx.meta[id] || {};
    var origin = String(meta.OriginType || '').trim().toUpperCase();
    var intent = String(meta.MembershipIntent || '').trim().toUpperCase();
    var since = parseDate_(meta.IntentDate || mField_(member, 'JoinDate'));
    var years = since ? monthsBetween_(since) / 12 : null;
    var needYears = origin === 'COUNCIL' ? Number(s.COUNCIL_REVIEW_YEARS)
                  : origin === 'OUTSIDE_COUNCIL' ? Number(s.OUTSIDE_COUNCIL_REVIEW_YEARS) : null;

    if (!origin || origin === 'UNKNOWN' || !intent || intent === 'UNKNOWN' || years == null) {
      system = base;
      flags.push('INFO_INCOMPLETE');
      explain.push('ข้อมูลสำหรับพิจารณายังไม่ครบ');
    } else if (intent === 'NO') {
      system = base;
      flags.push('NO_MEMBERSHIP_INTENT');
      explain.push('ยังไม่แสดงความประสงค์เป็นสมาชิกสมบูรณ์');
    } else if (years >= needYears) {
      system = 'PENDING_REVIEW';
      explain.push('ครบเกณฑ์เวลา ' + needYears +
        ' ปี — รอพิจารณาเป็นสมาชิกสมบูรณ์ (ต้องได้รับการอนุมัติ ไม่เลื่อนอัตโนมัติ)');
    } else {
      system = base;
      explain.push('ยังไม่ครบเกณฑ์เวลา (' + years.toFixed(1) + ' / ' + needYears + ' ปี)');
    }
    if (base === 'RESERVE' && monthsAbsent != null && monthsAbsent >= Number(s.FOLLOWUP_WARNING_MONTHS))
      flags.push('FOLLOWUP_6M');
  }

  return { base: base, system: system, monthsAbsent: monthsAbsent, lastWorship: last,
           flags: flags, explanation: explain };
}

/**
 * A KEEP_FULL1_BY_BOARD decision belongs to the absence period it was made for.
 * If the member returned to worship after the decision date, that period ended —
 * a new absence requires a NEW decision. Old decisions never silently reactivate.
 */
function decisionStillApplies_(decision, sysRes) {
  if (!decision) return false;
  var dDate = parseDate_(dField_(decision, 'DecisionDate'));
  if (!dDate) return false;
  var action = String(dField_(decision, 'Action')).trim();
  if (action !== 'KEEP_FULL1_BY_BOARD') return true;
  if (sysRes.system !== 'FULL2') return false;
  if (sysRes.lastWorship && dDate < sysRes.lastWorship) return false;
  var next = parseDate_(dField_(decision, 'NextReview'));
  if (next && next < today_()) return false;
  return true;
}

function calculateEffectiveStatus_(sysRes, decision) {
  var eff = sysRes.system, applied = null;
  var explain = sysRes.explanation.slice();
  var action = decision ? String(dField_(decision, 'Action')).trim() : '';

  if (decisionStillApplies_(decision, sysRes)) {
    applied = decision;
    switch (action) {
      case 'KEEP_FULL1_BY_BOARD':
        eff = 'FULL1_BOARD';
        explain.push('มติธรรมกิจ: คงสิทธิสมาชิกสมบูรณ์ 1 (' +
          isoDate_(parseDate_(dField_(decision, 'DecisionDate'))) + ')');
        break;
      case 'APPROVE_FULL1':   eff = 'FULL1'; break;
      case 'MOVE_TO_RESERVE': eff = 'RESERVE'; break;
      case 'TRANSFER':        eff = 'TRANSFERRED'; break;
      case 'RESIGN':          eff = 'RESIGNED'; break;
      case 'REMOVE_BY_BOARD': eff = 'REMOVED'; break;
      case 'MARK_DECEASED':   eff = 'DECEASED'; break;
      case 'KEEP_PENDING':    eff = 'PENDING_REVIEW'; break;
      default: applied = null;
    }
  } else if (action === 'KEEP_FULL1_BY_BOARD') {
    explain.push('มีมติธรรมกิจเดิมในประวัติ แต่ไม่มีผลกับช่วงการขาดนมัสการปัจจุบัน — ต้องมีมติใหม่');
  }

  var st = loadStatusTypes_()[eff] || {};
  return { effective: eff, activeDecision: applied,
           canVote: !!st.CanVote, canNominate: !!st.CanNominate, explanation: explain };
}

function buildContext_(skipRuntimeCheck) {
  var mode = runtimeMode_(true);
  var members = buildMemberIndex_();
  return {
    mode: mode.mode,
    settings: loadSettings_(),
    statusTypes: loadStatusTypes_(),
    members: members,
    attendance: buildAttendanceIndex_(members),
    meta: buildMetaIndex_(),
    decisions: buildDecisionIndex_()
  };
}

function evaluateMember_(ctx, m) {
  var k = m._key ? { key: m._key, keyType: m._keyType } : memberKey_(m);
  var id = k.key;
  var sys = calculateSystemStatus_(ctx, m);
  var dec = getActiveDecision_(ctx.decisions, id);
  var eff = calculateEffectiveStatus_(sys, dec);
  var age = calculateAge_(mField_(m, 'BirthDate'));
  var grp = calculateAgeGroup_(age, mField_(m, 'Gender'));
  return {
    MemberID: k.keyType === 'EXISTING_ID' ? id : '',
    MemberKey: id,
    KeyType: k.keyType,
    MemberName: mField_(m, 'FullName'),
    RegistryStatus: statusLabel_(sys.base),
    BaseStatusCode: sys.base,
    LastWorshipDate: isoDate_(sys.lastWorship),
    MonthsSinceLastWorship: sys.monthsAbsent,
    SystemResult: sys.system,
    SystemResultLabel: statusLabel_(sys.system),
    SystemFlags: sys.flags,
    ActiveDecision: dec ? {
      DecisionID: dec.DecisionID || '',
      Action: dField_(dec, 'Action'),
      DecisionDate: isoDate_(parseDate_(dField_(dec, 'DecisionDate'))),
      ReasonCodes: String(dField_(dec, 'ReasonCodes') || '').split(',').filter(String),
      NextReviewDate: isoDate_(parseDate_(dField_(dec, 'NextReview'))),
      RecordedBy: dField_(dec, 'RecordedBy')
    } : null,
    EffectiveResult: eff.effective,
    EffectiveResultLabel: statusLabel_(eff.effective),
    CanVote: eff.canVote,
    CanNominate: eff.canNominate,
    Age: age,
    CalculatedMinistryGroup: grp.code,
    MinistryGroupLabel: grp.label,
    Explanation: eff.explanation
  };
}

/* ============================================================
 * 11. ROLES & AUTHORIZATION
 * ============================================================ */

var ROLE_ADMIN = 'ADMIN', ROLE_EDITOR = 'EDITOR', ROLE_BOARD = 'BOARD', ROLE_REGISTRAR = 'REGISTRAR';

function normalizeRole_(r) {
  var s = String(r || '').trim().toLowerCase();
  if (['admin', 'ผู้ดูแลระบบ'].indexOf(s) >= 0) return ROLE_ADMIN;
  if (['editor', 'ผู้แก้ไข'].indexOf(s) >= 0) return ROLE_EDITOR;
  // legacy 'viewer' was the board/ธรรมกิจ account type — NOT the registrar
  if (['board', 'viewer', 'ธรรมกิจ', 'pastorate'].indexOf(s) >= 0) return ROLE_BOARD;
  if (['registrar', 'ฝ่ายทะเบียน'].indexOf(s) >= 0) return ROLE_REGISTRAR;
  return null;
}

var PERMISSIONS = {
  ADMIN:     { read: true, editMember: true,  decide: true,  settings: true,  accounts: true,  migrate: true,  money: true },
  EDITOR:    { read: true, editMember: true,  decide: false, settings: false, accounts: false, migrate: false, money: false },
  BOARD:     { read: true, editMember: false, decide: true,  settings: false, accounts: false, migrate: false, money: false },
  REGISTRAR: { read: true, editMember: false, decide: false, settings: false, accounts: false, migrate: false, money: false }
};

function authorizeRole_(session, capability) {
  var p = PERMISSIONS[session.role];
  if (!p || !p[capability]) throw new Error('DENIED: บทบาท ' + session.role + ' ไม่มีสิทธิ์ ' + capability);
  return true;
}

var REGISTRAR_FIELDS = ['MemberID', 'MemberKey', 'KeyType', 'MemberName', 'EffectiveResult',
                        'EffectiveResultLabel', 'RegistrarGroup', 'MinistryGroupLabel'];

function registrarGroup_(effCode) {
  var st = loadStatusTypes_()[effCode] || {};
  if (st.RegistrarGroup) return st.RegistrarGroup;
  if (effCode === 'FULL1' || effCode === 'FULL1_BOARD') return 'สมาชิกสมบูรณ์ 1';
  if (effCode === 'FULL2') return 'สมาชิกสมบูรณ์ 2';
  return 'ผู้ร่วมนมัสการ / ยังไม่เป็นสมาชิกสมบูรณ์';
}

/** The single place that decides what leaves the server. */
function sanitizeMemberForRole_(ev, raw, role) {
  if (role === ROLE_REGISTRAR) {
    var out = {};
    ev.RegistrarGroup = registrarGroup_(ev.EffectiveResult);
    REGISTRAR_FIELDS.forEach(function (f) { out[f] = ev[f]; });
    return out;                       // no notes, reasons, DOB, phone, flags, money
  }
  var o = {};
  Object.keys(ev).forEach(function (k) { o[k] = ev[k]; });
  o.BirthDate = isoDate_(parseDate_(mField_(raw, 'BirthDate')));
  o.Gender    = mField_(raw, 'Gender');
  o.Phone     = mField_(raw, 'Phone');
  o.Nickname  = mField_(raw, 'Nickname');
  o.Note      = mField_(raw, 'Note');
  if (role === ROLE_EDITOR) { delete o.ActiveDecision; delete o.Explanation; }
  return o;
}

/* ============================================================
 * 12. SESSIONS
 * ============================================================ */

function hash_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    s + '|' + (props_().getProperty('TOKEN_PEPPER') || ''));
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* ---- legacy (V2) password scheme -----------------------------------
 * Recovered verbatim from deploy-v2/apps-script/Code.gs:
 *   acc = salt + '|' + pass
 *   repeat 5000×:  acc = base64(HMAC-SHA256(acc, PEPPER))
 * PEPPER is a Script Property of the V2 project. Copy it into this project as
 * LEGACY_PEPPER (or keep the same PEPPER key) to keep existing logins working.
 * ------------------------------------------------------------------- */
var LEGACY_HASH_ROUNDS = 5000;

function legacyPepper_() {
  return props_().getProperty('LEGACY_PEPPER') || props_().getProperty('PEPPER') || '';
}

function legacyHash_(pass, salt) {
  var pep = legacyPepper_();
  if (!pep) return null;                       // cannot verify — never guess
  var acc = String(salt) + '|' + String(pass);
  for (var i = 0; i < LEGACY_HASH_ROUNDS; i++)
    acc = Utilities.base64Encode(Utilities.computeHmacSha256Signature(acc, pep));
  return acc;
}

function safeEq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= (a.charCodeAt(i) ^ b.charCodeAt(i));
  return diff === 0;
}

/**
 * V3 PasswordHash first, legacy passHash+salt as fallback.
 * Returns { ok, method: 'V3'|'LEGACY'|'NONE', reason }.
 * Never logs or returns any hash, salt, pepper or plaintext.
 */
function verifyPassword_(row, password) {
  var v3 = String(accField_(row, 'PasswordHash') || '').trim();
  if (v3 && safeEq_(hash_(String(password)), v3)) return { ok: true, method: 'V3' };

  var legacy = String(accField_(row, 'LegacyHash') || '').trim();
  var salt = String(accField_(row, 'Salt') || '');
  if (legacy && salt) {
    var h = legacyHash_(password, salt);
    if (h === null) return { ok: false, method: 'NONE', reason: 'LEGACY_PEPPER_MISSING' };
    if (safeEq_(h, legacy)) return { ok: true, method: 'LEGACY' };
    return { ok: false, method: 'LEGACY', reason: 'BAD_PASSWORD' };
  }
  if (!v3) return { ok: false, method: 'NONE', reason: 'PASSWORD_RESET_REQUIRED' };
  return { ok: false, method: 'V3', reason: 'BAD_PASSWORD' };
}

/** Silently re-hash a verified legacy password into the V3 format. No plaintext stored. */
function upgradePasswordHash_(row, password) {
  try {
    var t = readTable_('Accounts');
    var col = ensureColumn_(t, 'PasswordHash', true);
    if (col < 0) return false;
    t.sheet.getRange(row._row, col).setValue(hash_(String(password)));
    var ua = t.headers.indexOf('UpdatedAt');
    if (ua >= 0) t.sheet.getRange(row._row, ua + 1).setValue(nowIso_());
    invalidateCache_('Accounts');
    return true;
  } catch (e) { return false; }
}

function login_(username, password) {
  var t = readTable_('Accounts', true);
  var u = String(username || '').trim().toLowerCase();
  var row = null;
  t.rows.forEach(function (r) {
    if (String(accField_(r, 'Username') || '').trim().toLowerCase() === u) row = r;
  });
  if (!row) throw new Error('DENIED: บัญชีไม่ถูกต้องหรือถูกปิดใช้งาน');
  var activeRaw = accField_(row, 'Active');
  if (activeRaw !== '' && activeRaw != null && !truthy_(activeRaw))
    throw new Error('DENIED: บัญชีไม่ถูกต้องหรือถูกปิดใช้งาน');

  var v = verifyPassword_(row, password);
  if (!v.ok) {
    appendAuditLog_({ actor: String(accField_(row, 'Username')), action: 'LOGIN_FAIL',
                      entity: 'Accounts', note: v.reason || 'BAD_PASSWORD' });
    if (v.reason === 'PASSWORD_RESET_REQUIRED')
      throw new Error('DENIED: PASSWORD_RESET_REQUIRED — บัญชีนี้ต้องตั้งรหัสผ่านใหม่');
    if (v.reason === 'LEGACY_PEPPER_MISSING')
      throw new Error('DENIED: ยังไม่ได้ตั้งค่า LEGACY_PEPPER — ไม่สามารถตรวจสอบรหัสผ่านเดิมได้');
    throw new Error('DENIED: รหัสผ่านไม่ถูกต้อง');
  }
  if (v.method === 'LEGACY') upgradePasswordHash_(row, password);

  var role = normalizeRole_(accField_(row, 'Role'));
  if (!role) throw new Error('DENIED: บทบาทไม่ถูกต้อง');
  var uname = String(accField_(row, 'Username')).trim();
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('sess:' + token,
    JSON.stringify({ username: uname, role: role, at: nowIso_() }), 21600);
  appendAuditLog_({ actor: uname, action: 'LOGIN', entity: 'Accounts',
                    note: role + ' · ' + v.method });
  return { token: token, username: uname, role: role,
           displayName: accField_(row, 'DisplayName') || uname,
           passwordUpgraded: v.method === 'LEGACY' };
}

function requireSession_(token) {
  var raw = CacheService.getScriptCache().get('sess:' + String(token || ''));
  if (!raw) throw new Error('DENIED: เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return JSON.parse(raw);
}

/* ============================================================
 * 13. APPEND HELPERS · AUDIT · HISTORY
 * ============================================================ */

function appendRow_(sheetName, obj) {
  var t = readTable_(sheetName);
  var row = t.headers.map(function (h) { return (h in obj) ? obj[h] : ''; });
  t.sheet.appendRow(row);
  invalidateCache_(sheetName);
}

function appendAuditLog_(o) {
  try {
    var t = readTable_('AuditLog');
    var payload = {
      AuditID: Utilities.getUuid(),
      At: nowIso_(), Timestamp: nowIso_(),
      Actor: o.actor || '', ActorUser: o.actor || '',
      Action: o.action || '', EntityType: o.entity || '', EntityID: o.entityId || '',
      Before: o.before == null ? '' : JSON.stringify(o.before),
      After: o.after == null ? '' : JSON.stringify(o.after),
      BeforeValue: o.before == null ? '' : JSON.stringify(o.before),
      AfterValue: o.after == null ? '' : JSON.stringify(o.after),
      Note: o.note || '', Source: 'WEB', Active: true
    };
    var row = t.headers.map(function (h) { return (h in payload) ? payload[h] : ''; });
    t.sheet.appendRow(row);
    invalidateCache_('AuditLog');
  } catch (e) { /* audit must never block the operation */ }
}

function appendDecisionHistory_(h) {
  appendRow_('DecisionHistory', {
    HistoryID: Utilities.getUuid(),
    DecisionID: h.DecisionID || '', MemberID: h.MemberID || '',
    MemberNameSnapshot: h.MemberNameSnapshot || '', Action: h.Action || '',
    PreviousStatusCode: h.PreviousStatusCode || '', NewStatusCode: h.NewStatusCode || '',
    SystemResultAtDecision: h.SystemResultAtDecision || '',
    EffectiveResultAfter: h.EffectiveResultAfter || '',
    DecisionDate: h.DecisionDate || '', ReasonCodes: h.ReasonCodes || '',
    Note: h.Note || '', NextReviewDate: h.NextReviewDate || '',
    RecordedBy: h.RecordedBy || '', RecordedAt: nowIso_(),
    Source: h.Source || 'WEB', Active: true
  });
}

/* ============================================================
 * 14. MUTATIONS
 * ============================================================ */

/** MemberID is the permanent primary key — it is NOT in any editable list. */
var EDITABLE_MEMBER_FIELDS = {
  EDITOR: ['FullName', 'Nickname', 'BirthDate', 'Gender', 'Phone', 'Email', 'Address', 'Note'],
  ADMIN:  ['FullName', 'Nickname', 'BirthDate', 'Gender', 'Phone', 'Email', 'Address', 'Note',
           'StatusCode']
};
var PROTECTED_MEMBER_FIELDS = ['MemberID', 'รหัส', 'รหัสสมาชิก'];

function updateMember_(session, memberKey, patch) {
  authorizeRole_(session, 'editMember');
  var allow = EDITABLE_MEMBER_FIELDS[session.role] || [];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    invalidateCache_('Members');
    var t = readTable_('Members');
    var target = String(memberKey || '').trim();
    var row = null;
    t.rows.forEach(function (r) { if (memberKey_(r).key === target) row = r; });
    if (!row) throw new Error('ไม่พบสมาชิก ' + target);

    var before = {}, after = {};
    Object.keys(patch).forEach(function (k) {
      if (PROTECTED_MEMBER_FIELDS.indexOf(k) >= 0)
        throw new Error('DENIED: ไม่อนุญาตให้แก้ไข MemberID ผ่านการแก้ไขปกติ');
      if (allow.indexOf(k) < 0) return;                        // strict allowlist
      var header = pickHeader_(t, ALIAS.Members[k] || [k]) || k;
      var c = t.headers.indexOf(header);
      if (c < 0) return;
      before[k] = row[header];
      after[k] = patch[k];
      t.sheet.getRange(row._row, c + 1).setValue(patch[k]);
    });
    if (!Object.keys(after).length) throw new Error('ไม่มีฟิลด์ที่อนุญาตให้แก้ไข');

    var uc = t.headers.indexOf('UpdatedAt');
    if (uc >= 0) t.sheet.getRange(row._row, uc + 1).setValue(nowIso_());
    var ub = t.headers.indexOf('UpdatedBy');
    if (ub >= 0) t.sheet.getRange(row._row, ub + 1).setValue(session.username);

    invalidateCache_('Members');
    appendAuditLog_({ actor: session.username, action: 'UPDATE_MEMBER', entity: 'Members',
                      entityId: target, before: before, after: after });
    return { ok: true, memberKey: target, updated: after, mode: runtimeMode_(true).mode };
  } finally { lock.releaseLock(); }
}

function saveDecision_(session, d) {
  authorizeRole_(session, 'decide');
  assertDecisionsWritable_();
  if (truthy_(getSetting_('REQUIRE_DECISION_REASON')) && !String(d.ReasonCodes || '').trim())
    throw new Error('ต้องระบุเหตุผลของมติ');
  if (!d.MemberID) throw new Error('ต้องระบุ MemberID');
  if (!d.Action) throw new Error('ต้องระบุประเภทของมติ');
  if (!d.DecisionDate) throw new Error('ต้องระบุวันที่มติ');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    invalidateCache_();
    var ctx = buildContext_();
    var mid = String(d.MemberID).trim();
    var m = ctx.members.byId[mid];
    if (!m) throw new Error(ctx.members.byKey[mid]
      ? 'สมาชิกรายนี้ยังไม่มี MemberID ถาวร — ต้องรัน migrateMembersToV3(false) ก่อนจึงบันทึกมติได้'
      : 'ไม่พบสมาชิก ' + mid);
    var sys = calculateSystemStatus_(ctx, m);
    var prevDec = getActiveDecision_(ctx.decisions, mid);

    var dt = readTable_('Decisions');
    if (prevDec) {
      var ac = dt.headers.indexOf('Active');
      if (ac >= 0) dt.sheet.getRange(prevDec._row, ac + 1).setValue(false);
    }

    var decisionId = Utilities.getUuid();
    appendRow_('Decisions', {
      DecisionID: decisionId, MemberID: mid, MemberNameSnapshot: mField_(m, 'FullName'),
      Action: d.Action, PreviousStatusCode: sys.base, NewStatusCode: d.NewStatusCode || '',
      SystemResultAtDecision: sys.system, DecisionDate: d.DecisionDate,
      ReasonCodes: d.ReasonCodes || '', Note: d.Note || '',
      NextReviewDate: d.NextReviewDate || defaultNextReview_(d.DecisionDate),
      RecordedBy: session.username, RecordedAt: nowIso_(), Source: 'WEB', Active: true
    });
    invalidateCache_('Decisions');

    var eff = calculateEffectiveStatus_(sys, buildDecisionIndex_()[mid]);
    appendDecisionHistory_({
      DecisionID: decisionId, MemberID: mid, MemberNameSnapshot: mField_(m, 'FullName'),
      Action: d.Action, PreviousStatusCode: sys.base, NewStatusCode: d.NewStatusCode || '',
      SystemResultAtDecision: sys.system, EffectiveResultAfter: eff.effective,
      DecisionDate: d.DecisionDate, ReasonCodes: d.ReasonCodes || '', Note: d.Note || '',
      NextReviewDate: d.NextReviewDate || '', RecordedBy: session.username
    });
    appendAuditLog_({ actor: session.username, action: 'SAVE_DECISION', entity: 'Decisions',
                      entityId: decisionId, after: d });
    invalidateCache_();
    return { ok: true, DecisionID: decisionId, EffectiveResult: eff.effective };
  } finally { lock.releaseLock(); }
}

function defaultNextReview_(from) {
  var d = parseDate_(from) || today_();
  d.setMonth(d.getMonth() + Number(getSetting_('DEFAULT_NEXT_REVIEW_MONTHS') || 12));
  return isoDate_(d);
}

function clearDecision_(session, memberId, note) {
  authorizeRole_(session, 'decide');
  assertDecisionsWritable_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    invalidateCache_();
    var ctx = buildContext_();
    var mid = String(memberId).trim();
    var dec = getActiveDecision_(ctx.decisions, mid);
    if (!dec) throw new Error('ไม่มีมติที่ใช้งานอยู่');
    var dt = readTable_('Decisions');
    var ac = dt.headers.indexOf('Active');
    if (ac >= 0) dt.sheet.getRange(dec._row, ac + 1).setValue(false);
    appendDecisionHistory_({
      DecisionID: dec.DecisionID || '', MemberID: mid,
      MemberNameSnapshot: mField_(ctx.members.byId[mid] || {}, 'FullName'),
      Action: 'CLEAR_DECISION', DecisionDate: isoDate_(today_()),
      Note: note || '', RecordedBy: session.username
    });
    appendAuditLog_({ actor: session.username, action: 'CLEAR_DECISION', entity: 'Decisions',
                      entityId: dec.DecisionID || '', note: note || '' });
    invalidateCache_();
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/** Validates the RESULTING configuration before writing. Invalid → nothing is saved. */
function saveSetting_(session, key, value, valueType) {
  authorizeRole_(session, 'settings');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    invalidateCache_('Settings'); _settings = null;
    var t = readTable_('Settings');
    var c = settingsCols_(t);
    if (!c.key || !c.value) throw new Error('SCHEMA: Settings needs SettingKey / SettingValue');

    var row = null;
    t.rows.forEach(function (r) { if (String(r[c.key]).trim() === String(key).trim()) row = r; });
    if (row && hasHeader_(t, 'EditableByAdmin') && row.EditableByAdmin !== '' &&
        row.EditableByAdmin != null && !truthy_(row.EditableByAdmin))
      throw new Error('DENIED: ค่านี้ไม่อนุญาตให้แก้ไขผ่านหน้าเว็บ');

    // pre-write validation of the proposed configuration
    var proposed = {};
    var current = loadSettings_(true);
    Object.keys(current).forEach(function (k) { proposed[k] = current[k]; });
    proposed[key] = coerceSetting_(value, valueType || (row ? row.ValueType : ''));
    var bands = validateAgeBandsFrom_(proposed);
    if (!bands.ok) throw new Error('ค่าที่ตั้งไม่ถูกต้อง (ไม่ได้บันทึก): ' + bands.problems.join(' | '));

    var before = row ? row[c.value] : null;
    if (row) {
      t.sheet.getRange(row._row, t.headers.indexOf(c.value) + 1).setValue(value);
      var ua = t.headers.indexOf('UpdatedAt');
      if (ua >= 0) t.sheet.getRange(row._row, ua + 1).setValue(nowIso_());
      var ubx = t.headers.indexOf('UpdatedBy');
      if (ubx >= 0) t.sheet.getRange(row._row, ubx + 1).setValue(session.username);
    } else {
      var payload = { ValueType: valueType || 'TEXT', Category: 'RULE', Description: '',
                      EditableByAdmin: true, Active: true,
                      UpdatedBy: session.username, UpdatedAt: nowIso_() };
      payload[c.key] = key; payload[c.value] = value;
      appendRow_('Settings', payload);
    }
    invalidateCache_('Settings'); _settings = null;
    appendAuditLog_({ actor: session.username, action: 'SAVE_SETTING', entity: 'Settings',
                      entityId: key, before: before, after: value });
    return { ok: true, key: key, value: value, ageBands: validateAgeBands_() };
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * 15. READ APIS
 * ============================================================ */

var REGISTRAR_GROUPS = [
  { GroupCode: 'FULL1_GROUP',    DisplayName: 'สมาชิกสมบูรณ์ 1', SortOrder: 1, includes: ['FULL1', 'FULL1_BOARD'] },
  { GroupCode: 'FULL2_GROUP',    DisplayName: 'สมาชิกสมบูรณ์ 2', SortOrder: 2, includes: ['FULL2'] },
  { GroupCode: 'ATTENDER_GROUP', DisplayName: 'ผู้ร่วมนมัสการ / ยังไม่เป็นสมาชิกสมบูรณ์', SortOrder: 3, includes: null }
];

function registrarGroupCode_(effCode) {
  for (var i = 0; i < REGISTRAR_GROUPS.length; i++) {
    var g = REGISTRAR_GROUPS[i];
    if (g.includes && g.includes.indexOf(effCode) >= 0) return g.GroupCode;
  }
  return 'ATTENDER_GROUP';
}

function listMembers_(session) {
  var ctx = buildContext_();
  var out = [];
  ctx.members.rows.forEach(function (r) {
    out.push(sanitizeMemberForRole_(evaluateMember_(ctx, r), r, session.role));
  });
  return { members: out, statusTypes: publicStatusTypes_(session),
           settings: publicSettings_(session), mode: ctx.mode,
           unresolvedAttendanceRows: ctx.attendance._unresolved || 0,
           generatedAt: nowIso_() };
}

/** Registrar receives only the three broad group labels — no internal status list. */
function publicStatusTypes_(session) {
  if (session && session.role === ROLE_REGISTRAR)
    return REGISTRAR_GROUPS.map(function (g) {
      return { StatusCode: g.GroupCode, DisplayName: g.DisplayName,
               ShortName: g.DisplayName, SortOrder: g.SortOrder, Active: true,
               DashboardVisible: true };
    });
  var m = loadStatusTypes_(), out = [];
  Object.keys(m).forEach(function (k) {
    var s = m[k];
    out.push({ StatusCode: s.StatusCode, DisplayName: s.DisplayName, ShortName: s.ShortName,
               StatusClass: s.StatusClass, Active: s.Active, SortOrder: s.SortOrder,
               ColorHex: s.ColorHex, DashboardVisible: s.DashboardVisible,
               RegistrarGroup: s.RegistrarGroup, CountAsCompletedMember: s.CountAsCompletedMember });
  });
  out.sort(function (a, b) { return a.SortOrder - b.SortOrder; });
  return out;
}

function publicSettings_(session) {
  var s = loadSettings_();
  if (session.role === ROLE_ADMIN) return s;
  var out = {};
  ['SYSTEM_NAME', 'TIMEZONE', 'LOCALE', 'FULL2_ABSENCE_MONTHS', 'FULL2_REVIEW_MONTHS',
   'FOLLOWUP_WARNING_MONTHS', 'COUNCIL_REVIEW_YEARS', 'OUTSIDE_COUNCIL_REVIEW_YEARS',
   'SUNDAY_SCHOOL_MIN_AGE', 'SUNDAY_SCHOOL_MAX_AGE', 'YOUTH_MIN_AGE', 'YOUTH_MAX_AGE',
   'ADULT_MIN_AGE', 'ADULT_MAX_AGE', 'SENIOR_MIN_AGE'].forEach(function (k) { out[k] = s[k]; });
  return out;
}

function dashboardCounts_(session) {
  var ctx = buildContext_();

  // Registrar: three broad groups only. No status breakdown, no analytics, no decisions.
  if (session.role === ROLE_REGISTRAR) {
    var g = { FULL1_GROUP: 0, FULL2_GROUP: 0, ATTENDER_GROUP: 0 };
    ctx.members.rows.forEach(function (m) {
      g[registrarGroupCode_(evaluateMember_(ctx, m).EffectiveResult)]++;
    });
    return {
      groups: REGISTRAR_GROUPS.map(function (x) {
        return { GroupCode: x.GroupCode, DisplayName: x.DisplayName,
                 SortOrder: x.SortOrder, Count: g[x.GroupCode] };
      }),
      mode: ctx.mode,
      generatedAt: nowIso_()
    };
  }

  var c = {}, analytics = { FOLLOWUP_6M: 0, ABSENT_1Y: 0, ABSENT_2Y: 0, ABSENT_3Y: 0,
                            ABSENT_4Y_PLUS: 0, STRONG: 0, REVIEW_FOR_RESERVE: 0 };
  ctx.members.rows.forEach(function (m) {
    var ev = evaluateMember_(ctx, m);
    c[ev.EffectiveResult] = (c[ev.EffectiveResult] || 0) + 1;
    if (ev.SystemFlags.indexOf('FOLLOWUP_6M') >= 0) analytics.FOLLOWUP_6M++;
    if (ev.SystemFlags.indexOf('REVIEW_FOR_RESERVE') >= 0) analytics.REVIEW_FOR_RESERVE++;
    var ma = ev.MonthsSinceLastWorship;
    if (ma != null) {
      if (ma >= 48) analytics.ABSENT_4Y_PLUS++;
      else if (ma >= 36) analytics.ABSENT_3Y++;
      else if (ma >= 24) analytics.ABSENT_2Y++;
      else if (ma >= 12) analytics.ABSENT_1Y++;
    }
    if (countVisitsSince_(ctx.attendance, ev.MemberKey, 12) >= Number(ctx.settings.STRONG_MIN_VISITS_12M))
      analytics.STRONG++;
  });
  c.TOTAL_COMPLETED = (c.FULL1 || 0) + (c.FULL1_BOARD || 0) + (c.FULL2 || 0);
  return { statusCounts: c, analytics: analytics, mode: ctx.mode, generatedAt: nowIso_() };
}

/** READ-ONLY preview. Writes nothing. Works in compatibility mode. */
function previewStatusEngine(limit) {
  var ctx = buildContext_();
  return ctx.members.rows.slice(0, limit || 20).map(function (m) { return evaluateMember_(ctx, m); });
}

/* ============================================================
 * 16. IMPORT LOG / UNMATCHED RECORDS  (real schemas)
 * ============================================================ */

var IMPORT_STATUS = { RUNNING: 'RUNNING', SUCCESS: 'SUCCESS', PARTIAL: 'PARTIAL',
                      FAILED: 'FAILED', CANCELLED: 'CANCELLED' };

/** Configured code values — do not extend without updating the Sheet's validation. */
var IMPORT_TYPES = ['INITIAL', 'FULL', 'DELTA', 'FINAL_CUTOVER', 'REIMPORT'];
var ENTITY_TYPES = ['MEMBERS', 'ATTENDANCE', 'GIVING', 'PHOTOS', 'DECISIONS', 'OTHER'];

function assertCode_(list, v, what) {
  if (list.indexOf(v) < 0) throw new Error('CODE: ค่า ' + what + ' ไม่ถูกต้อง: ' + v);
  return v;
}

/** Detail that is not a valid code goes into Notes / BatchID, never into a code column. */
function startImport_(importType, entityType, target, dryRun, runBy, detail) {
  assertCode_(IMPORT_TYPES, importType, 'ImportType');
  assertCode_(ENTITY_TYPES, entityType, 'EntityType');
  var id = Utilities.getUuid();
  var batch = (detail ? detail + '-' : '') + Utilities.getUuid().slice(0, 8);
  appendRow_('ImportLog', {
    ImportID: id, BatchID: batch, ImportType: importType, EntityType: entityType,
    SourceSystem: 'LEGACY_V2', TargetSheet: target, DryRun: !!dryRun,
    StartedAt: nowIso_(), Status: IMPORT_STATUS.RUNNING, RunBy: runBy || 'script',
    Notes: detail || '', Active: true
  });
  return { ImportID: id, BatchID: batch, detail: detail || '' };
}

function finishImport_(importId, stats, status, notes) {
  var t = readTable_('ImportLog', true);
  var row = null;
  t.rows.forEach(function (r) { if (String(r.ImportID).trim() === importId) row = r; });
  if (!row) return;
  var set = {
    FinishedAt: nowIso_(), RowsRead: stats.read, RowsInserted: stats.inserted || 0,
    RowsUpdated: stats.matched || 0, RowsSkipped: stats.skipped || 0,
    RowsUnmatched: stats.unmatched || 0, ErrorCount: stats.errors || 0,
    Status: status, Notes: notes || ''
  };
  Object.keys(set).forEach(function (h) {
    var c = t.headers.indexOf(h);
    if (c >= 0) t.sheet.getRange(row._row, c + 1).setValue(set[h]);
  });
  invalidateCache_('ImportLog');
}

function addUnmatched_(u) {
  appendRow_('UnmatchedRecords', {
    UnmatchedID: Utilities.getUuid(),
    ImportID: u.ImportID || '', BatchID: u.BatchID || '',
    EntityType: assertCode_(ENTITY_TYPES, u.EntityType, 'EntityType'),
    SourceSystem: 'LEGACY_V2',
    SourceSheet: u.SourceSheet || '', SourceRow: u.SourceRow || '',
    SourceRecordID: u.SourceRecordID || '', SourceName: u.SourceName || '',
    SourceDate: u.SourceDate || '', SourceAmount: u.SourceAmount || '',
    MatchReason: u.MatchReason || '', MatchStatus: 'OPEN',
    SuggestedMemberID: u.SuggestedMemberID || '', MatchedMemberID: '',
    ResolutionAction: '', ResolutionNote: '', ResolvedBy: '', ResolvedAt: '',
    CreatedAt: nowIso_(), Active: true
  });
}

function importStatusFor_(stats) {
  if (stats.errors) return IMPORT_STATUS.FAILED;
  if (stats.unmatched) return IMPORT_STATUS.PARTIAL;
  return IMPORT_STATUS.SUCCESS;
}

/* ============================================================
 * 17. MIGRATIONS  (legacy-tolerant · DRY RUN by default)
 * ============================================================ */

/**
 * STEP 1 — normalize member identity + status.
 * Reads legacy Thai headers (รหัส / ชื่อ สกุล / สถานะ / วันเกิด) and writes the
 * normalized MemberID / FullName / StatusCode / BirthDate columns, creating them
 * on apply if they don't exist. Never edits the legacy columns.
 */
function newMemberId_() { return 'V3M-' + Utilities.getUuid(); }

function migrateMembersToV3(dryRun) {
  dryRun = (dryRun !== false);
  assertLegacySchema_();
  var imp = startImport_('INITIAL', 'MEMBERS', 'Members', dryRun, actor_(), 'MEMBERS_NORMALIZE');
  var t = readTable_('Members', true);
  var stats = { read: 0, matched: 0, inserted: 0, skipped: 0, unmatched: 0, errors: 0 };
  var preview = [], warnings = [], unknownStatusRows = [], seenIds = {};
  var report = { total: 0, existingMemberID: 0, willCreateMemberID: 0, createdMemberID: 0,
                 duplicateMemberID: 0, blankName: 0, recognizedStatus: 0, unknownStatus: 0 };

  var cId     = ensureColumn_(t, 'MemberID',  !dryRun);
  var cName   = ensureColumn_(t, 'FullName',  !dryRun);
  var cStatus = ensureColumn_(t, 'StatusCode', !dryRun);
  var cBirth  = ensureColumn_(t, 'BirthDate', !dryRun);

  // pre-scan: every legacy รหัส already in use, so generated IDs cannot collide
  t.rows.forEach(function (r) {
    var e = String(mField_(r, 'MemberID')).trim();
    if (e) seenIds[e] = (seenIds[e] || 0) + 1;
  });

  var assigned = {};
  t.rows.forEach(function (r) {
    stats.read++; report.total++;
    var existing = String(mField_(r, 'MemberID')).trim();
    var name = normName_(mField_(r, 'FullName'));
    var rawStatus = mField_(r, 'StatusRaw');
    var code = normalizeLegacyStatus_(rawStatus);
    var birth = parseDate_(mField_(r, 'BirthDate'));
    if (!name) report.blankName++;

    var id = existing, idAction;
    if (existing) {
      if (seenIds[existing] > 1 && assigned[existing]) {        // second+ occurrence
        report.duplicateMemberID++;
        stats.unmatched++;
        warnings.push('MemberID ซ้ำ: ' + existing + ' (แถว ' + r._row + ')');
        if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
          EntityType: 'MEMBERS', SourceSheet: 'Members', SourceRow: r._row,
          SourceRecordID: existing, SourceName: name, MatchReason: 'DUPLICATE_MEMBER_ID' });
        return;                                                  // never overwrite a duplicate
      }
      report.existingMemberID++;
      idAction = 'KEEP_EXISTING';
      assigned[existing] = true;
    } else {
      idAction = dryRun ? 'WILL_BE_CREATED' : 'CREATED';
      if (dryRun) {
        report.willCreateMemberID++;
        id = '';
      } else {
        do { id = newMemberId_(); } while (seenIds[id]);          // verify uniqueness
        seenIds[id] = 1; assigned[id] = true;
        report.createdMemberID++; report.willCreateMemberID++;
      }
    }

    // Unknown legacy status is never guessed: raw text is preserved in the legacy
    // column, StatusCode becomes UNSPECIFIED (needs review), and a review item is filed.
    if (code === 'UNSPECIFIED' && String(rawStatus || '').trim()) {
      report.unknownStatus++;
      stats.unmatched++;
      unknownStatusRows.push({ row: r._row, MemberID: id || '(จะสร้างใหม่)',
                               FullName: name, rawStatus: String(rawStatus) });
      warnings.push('สถานะเดิมไม่รู้จัก: "' + String(rawStatus) + '" · ' + (name || 'แถว ' + r._row) +
                    ' → ตั้งเป็น UNSPECIFIED รอการพิจารณา');
      if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
        EntityType: 'MEMBERS', SourceSheet: 'Members', SourceRow: r._row,
        SourceRecordID: id, SourceName: String(rawStatus), MatchReason: 'UNKNOWN_STATUS_TEXT',
        SuggestedMemberID: id });
    } else {
      report.recognizedStatus++;
      stats.matched++;
    }

    if (preview.length < 50)
      preview.push({ row: r._row, MemberID: id || null, idAction: idAction, FullName: name,
                     status: { from: rawStatus, to: code }, birth: isoDate_(birth) });

    if (!dryRun) {
      t.sheet.getRange(r._row, cId).setValue(id);                 // legacy รหัส untouched
      if (name) t.sheet.getRange(r._row, cName).setValue(name);
      t.sheet.getRange(r._row, cStatus).setValue(code);
      if (birth) t.sheet.getRange(r._row, cBirth).setValue(isoDate_(birth));
      stats.inserted++;
    }
  });

  report.readyForApply = report.duplicateMemberID === 0 && report.blankName === 0;
  finishImport_(imp.ImportID, stats, dryRun ? IMPORT_STATUS.SUCCESS : importStatusFor_(stats),
                (dryRun ? 'DRY RUN — no data written · ' : '') +
                'existing=' + report.existingMemberID +
                ' new=' + (dryRun ? report.willCreateMemberID : report.createdMemberID) +
                ' unknownStatus=' + report.unknownStatus);
  invalidateCache_('Members');
  return { dryRun: dryRun, importId: imp.ImportID, report: report, stats: stats,
           unknownStatusRows: unknownStatusRows,
           warnings: warnings.slice(0, 50), preview: preview };
}

/** Alias kept for continuity with the previous build. */
function migrateLegacyStatuses(dryRun) { return migrateMembersToV3(dryRun); }

/**
 * STEP 2 — attach MemberID to Attendance.
 * Runs BEFORE Attendance has a MemberID column: detects the real date column,
 * matches on ชื่อ สกุล only when exactly one unambiguous member matches,
 * creates the MemberID column at apply time. Duplicates / no-match → UnmatchedRecords.
 */
function migrateAttendanceToMemberID(dryRun) {
  dryRun = (dryRun !== false);
  assertLegacySchema_();
  var imp = startImport_('INITIAL', 'ATTENDANCE', 'Attendance', dryRun, actor_(), 'ATTENDANCE_MEMBERID');
  var idx = buildMemberIndex_();
  var t = readTable_('Attendance', true);
  var nameHeader = pickHeader_(t, ALIAS.Attendance.Name);
  var detected = detectAttendanceDateColumn_(true);
  var dateHeader = detected.header;
  var col = ensureColumn_(t, 'MemberID', !dryRun);
  var stats = { read: 0, matched: 0, inserted: 0, skipped: 0, unmatched: 0, errors: 0 };
  var sample = [];

  if (!dateHeader) {
    finishImport_(imp.ImportID, stats, IMPORT_STATUS.FAILED,
                  'date column ' + detected.method + ': ' + JSON.stringify(detected.candidates));
    requireAttendanceDateColumn_();                      // throws with the candidate list
  }
  if (!nameHeader && col < 0) {
    finishImport_(imp.ImportID, stats, IMPORT_STATUS.FAILED, 'no name column to match on');
    throw new Error('SCHEMA: Attendance มีทั้ง MemberID และคอลัมน์ชื่อไม่ครบ — ไม่สามารถจับคู่ได้');
  }

  t.rows.forEach(function (r) {
    stats.read++;
    if (String(aField_(r, 'MemberID')).trim()) { stats.skipped++; return; }
    var n = normName_(nameHeader ? r[nameHeader] : '');
    var d = parseDate_(dateHeader ? r[dateHeader] : '');
    var reason = !n ? 'NO_NAME'
               : idx.dupNames[n] ? 'DUPLICATE_NAME'
               : !idx.byName[n] ? 'NO_MATCH'
               : !String(mField_(idx.byName[n], 'MemberID')).trim() ? 'MEMBER_HAS_NO_ID' : null;
    if (reason) {
      stats.unmatched++;
      if (sample.length < 25) sample.push({ row: r._row, name: n, reason: reason });
      if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
        EntityType: 'ATTENDANCE', SourceSheet: 'Attendance', SourceRow: r._row,
        SourceName: n, SourceDate: isoDate_(d), MatchReason: reason });
      return;
    }
    stats.matched++;
    if (!dryRun) t.sheet.getRange(r._row, col).setValue(String(mField_(idx.byName[n], 'MemberID')).trim());
  });

  finishImport_(imp.ImportID, stats, dryRun ? IMPORT_STATUS.SUCCESS : importStatusFor_(stats),
                dryRun ? 'DRY RUN — no data written' : '');
  invalidateCache_('Attendance');
  _attDateCol = null;
  return { dryRun: dryRun, importId: imp.ImportID, nameColumn: nameHeader,
           dateColumn: dateHeader, dateColumnMethod: detected.method,
           dateColumnConfidence: detected.confidence, stats: stats, unmatchedSample: sample };
}

/**
 * STEP 3 — convert legacy Decisions (memberKey|memberName|decision|by|at|note)
 * into the V3 structure. Legacy rows are never deleted; V3 columns are added and
 * filled alongside them, and every converted row is mirrored into DecisionHistory.
 */
function migrateLegacyDecisions(dryRun) {
  dryRun = (dryRun !== false);
  assertLegacySchema_();
  var imp = startImport_('INITIAL', 'DECISIONS', 'Decisions', dryRun, actor_(), 'DECISIONS_V3');
  var idx = buildMemberIndex_();
  var t = readTable_('Decisions', true);
  var stats = { read: 0, matched: 0, inserted: 0, skipped: 0, unmatched: 0, errors: 0 };
  var preview = [], converted = [];

  // every column saveDecision_ writes must exist after this migration
  var V3_DECISION_COLUMNS = ['DecisionID', 'MemberID', 'MemberNameSnapshot', 'Action',
    'PreviousStatusCode', 'NewStatusCode', 'SystemResultAtDecision', 'DecisionDate',
    'ReasonCodes', 'Note', 'NextReviewDate', 'RecordedBy', 'RecordedAt', 'Source', 'Active'];
  var C = {};
  V3_DECISION_COLUMNS.forEach(function (h) { C[h] = ensureColumn_(t, h, !dryRun); });

  t.rows.forEach(function (r) {
    stats.read++;
    if (String(r.DecisionID || '').trim() && String(r.MemberID || '').trim()) { stats.skipped++; return; }

    var key = String(dField_(r, 'MemberID')).trim();       // memberKey or MemberID
    var name = normName_(dField_(r, 'MemberName'));
    var member = key && idx.byId[key] ? idx.byId[key]
               : (name && !idx.dupNames[name] && idx.byName[name] &&
                  String(mField_(idx.byName[name], 'MemberID')).trim() ? idx.byName[name] : null);

    if (!member) {
      stats.unmatched++;
      if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
        EntityType: 'DECISIONS', SourceSheet: 'Decisions', SourceRow: r._row,
        SourceRecordID: key, SourceName: name,
        SourceDate: isoDate_(parseDate_(dField_(r, 'DecisionDate'))),
        MatchReason: !key && !name ? 'NO_IDENTIFIER'
                   : (name && idx.dupNames[name]) ? 'DUPLICATE_NAME' : 'NO_MATCH' });
      return;
    }

    var mid = String(mField_(member, 'MemberID')).trim();
    var action = mapLegacyDecisionAction_(dField_(r, 'Action'));
    var date = parseDate_(dField_(r, 'DecisionDate'));
    var decisionId = String(r.DecisionID || '').trim() || Utilities.getUuid();

    if (action === 'OTHER' && String(dField_(r, 'Action') || '').trim()) {
      stats.unmatched++;
      if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
        EntityType: 'DECISIONS', SourceSheet: 'Decisions', SourceRow: r._row,
        SourceRecordID: mid, SourceName: String(dField_(r, 'Action')),
        MatchReason: 'UNKNOWN_DECISION_ACTION', SuggestedMemberID: mid });
    } else {
      stats.matched++;
    }

    if (!date) {
      stats.unmatched++;
      if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
        EntityType: 'DECISIONS', SourceSheet: 'Decisions', SourceRow: r._row,
        SourceRecordID: mid, SourceName: String(dField_(r, 'DecisionDate') || ''),
        MatchReason: 'UNKNOWN_DECISION_DATE', SuggestedMemberID: mid });
    }

    converted.push({ row: r._row, decisionId: decisionId, memberId: mid,
                     name: mField_(member, 'FullName'), action: action, date: date,
                     by: dField_(r, 'RecordedBy') || 'LEGACY', note: dField_(r, 'Note') });

    if (preview.length < 50) preview.push({
      row: r._row, MemberID: mid, name: mField_(member, 'FullName'),
      action: { from: dField_(r, 'Action'), to: action }, date: isoDate_(date)
    });

    if (!dryRun) {
      t.sheet.getRange(r._row, C.DecisionID).setValue(decisionId);
      t.sheet.getRange(r._row, C.MemberID).setValue(mid);
      t.sheet.getRange(r._row, C.MemberNameSnapshot).setValue(mField_(member, 'FullName'));
      t.sheet.getRange(r._row, C.Action).setValue(action);
      if (date) t.sheet.getRange(r._row, C.DecisionDate).setValue(isoDate_(date));
      t.sheet.getRange(r._row, C.RecordedBy).setValue(dField_(r, 'RecordedBy') || 'LEGACY');
      t.sheet.getRange(r._row, C.Note).setValue(dField_(r, 'Note'));
      t.sheet.getRange(r._row, C.Source).setValue('MIGRATION');
      if (!String(r.RecordedAt || '').trim())
        t.sheet.getRange(r._row, C.RecordedAt).setValue(isoDate_(date) || nowIso_());
      t.sheet.getRange(r._row, C.Active).setValue(false);   // resolved in the pass below
      appendDecisionHistory_({
        DecisionID: decisionId, MemberID: mid, MemberNameSnapshot: mField_(member, 'FullName'),
        Action: action, DecisionDate: isoDate_(date), Note: dField_(r, 'Note'),
        RecordedBy: dField_(r, 'RecordedBy') || 'LEGACY', Source: 'MIGRATION'
      });
      stats.inserted++;
    }
  });

  /* Exactly one Active decision per member: the latest dated row.
     CLEAR_DECISION as the latest row means the member has no active decision.
     Everything else stays Active=false — the full trail lives in DecisionHistory. */
  var latest = {}, activeRows = [];
  converted.forEach(function (c) {
    var cur = latest[c.memberId];
    if (!cur) { latest[c.memberId] = c; return; }
    var a = c.date ? c.date.getTime() : -1;
    var b = cur.date ? cur.date.getTime() : -1;
    if (a > b || (a === b && c.row > cur.row)) latest[c.memberId] = c;
  });
  Object.keys(latest).forEach(function (mid) {
    var c = latest[mid];
    if (c.action === 'CLEAR_DECISION' || !c.date) return;
    activeRows.push({ memberId: mid, row: c.row, action: c.action, date: isoDate_(c.date) });
    if (!dryRun) t.sheet.getRange(c.row, C.Active).setValue(true);
  });

  finishImport_(imp.ImportID, stats, dryRun ? IMPORT_STATUS.SUCCESS : importStatusFor_(stats),
                (dryRun ? 'DRY RUN — no data written · ' : '') +
                'active decisions after migration: ' + activeRows.length);
  invalidateCache_('Decisions');
  return { dryRun: dryRun, importId: imp.ImportID, stats: stats, preview: preview,
           activeCount: activeRows.length, activeSample: activeRows.slice(0, 25) };
}

var LEGACY_DECISION_ACTIONS = {
  'คงสมาชิกสมบูรณ์ 1': 'KEEP_FULL1_BY_BOARD', 'คงสิทธิสมาชิกสมบูรณ์ 1': 'KEEP_FULL1_BY_BOARD',
  'มติธรรมกิจ': 'KEEP_FULL1_BY_BOARD', 'keep_full1': 'KEEP_FULL1_BY_BOARD',
  'อนุมัติเป็นสมาชิกสมบูรณ์': 'APPROVE_FULL1', 'approve': 'APPROVE_FULL1',
  'รอพิจารณา': 'KEEP_PENDING', 'pending': 'KEEP_PENDING',
  'ย้ายเป็นสมาชิกสำรอง': 'MOVE_TO_RESERVE', 'reserve': 'MOVE_TO_RESERVE',
  'ย้ายสังกัด': 'TRANSFER', 'transfer': 'TRANSFER',
  'ลาออก': 'RESIGN', 'resign': 'RESIGN',
  'ให้ออกโดยมติธรรมกิจ': 'REMOVE_BY_BOARD', 'remove': 'REMOVE_BY_BOARD',
  'เสียชีวิต': 'MARK_DECEASED', 'deceased': 'MARK_DECEASED',
  'ทบทวน': 'REVIEW', 'review': 'REVIEW',
  'ยกเลิกมติ': 'CLEAR_DECISION', 'clear': 'CLEAR_DECISION'
};

var V3_DECISION_ACTIONS = ['CREATE_DECISION', 'CHANGE_STATUS', 'KEEP_FULL1_BY_BOARD',
  'APPROVE_FULL1', 'KEEP_PENDING', 'MOVE_TO_RESERVE', 'TRANSFER', 'RESIGN',
  'REMOVE_BY_BOARD', 'MARK_DECEASED', 'REVIEW', 'CLEAR_DECISION', 'OTHER'];

function mapLegacyDecisionAction_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return 'OTHER';
  if (V3_DECISION_ACTIONS.indexOf(s.toUpperCase()) >= 0) return s.toUpperCase();
  var k = s.replace(/\s+/g, ' ');
  if (LEGACY_DECISION_ACTIONS[k]) return LEGACY_DECISION_ACTIONS[k];
  if (LEGACY_DECISION_ACTIONS[k.toLowerCase()]) return LEGACY_DECISION_ACTIONS[k.toLowerCase()];
  return 'OTHER';
}

function actor_() {
  try { return Session.getActiveUser().getEmail() || 'script'; } catch (e) { return 'script'; }
}

/**
 * STEP 4 — normalize Accounts.
 * Legacy headers (user | passHash | salt | role | active | updatedAt | note |
 * pwVersion | passwordHint | pwChangedAt, including the duplicate pwChangedAt)
 * are PRESERVED untouched. Only the normalized V3 columns are added/filled:
 *   Username, PasswordHash, Role, Active, DisplayName, UpdatedAt
 * PasswordHash is left EMPTY — legacy hashes cannot be converted without the
 * plaintext, so verifyPassword_() falls back to the legacy scheme and the row
 * self-upgrades on the next successful login.
 * EntityType uses the configured code OTHER; the detail lives in Notes/BatchID.
 */
function migrateAccountsToV3(dryRun) {
  dryRun = (dryRun !== false);
  var t = readTable_('Accounts', true);
  var uCol = pickHeader_(t, ALIAS.Accounts.Username);
  var rCol = pickHeader_(t, ALIAS.Accounts.Role);
  if (!uCol) throw new Error('SCHEMA: Accounts ต้องมีคอลัมน์ user หรือ Username');

  var imp = startImport_('INITIAL', 'OTHER', 'Accounts', dryRun, actor_(), 'ACCOUNTS_NORMALIZE');
  var stats = { read: 0, matched: 0, inserted: 0, skipped: 0, unmatched: 0, errors: 0 };
  var preview = [], warnings = [], seen = {};

  var C = {};
  ['Username', 'PasswordHash', 'Role', 'Active', 'DisplayName', 'UpdatedAt']
    .forEach(function (h) { C[h] = ensureColumn_(t, h, !dryRun); });

  var dupHeaders = {};
  t.headers.forEach(function (h, i) {
    if (!h) return;
    if (t.headers.indexOf(h) !== i && dupHeaders[h] === undefined) dupHeaders[h] = true;
  });
  if (Object.keys(dupHeaders).length)
    warnings.push('พบหัวคอลัมน์ซ้ำใน Accounts: ' + Object.keys(dupHeaders).join(', ') +
                  ' · เก็บข้อมูลเดิมไว้ ใช้ค่าแรกสุดเท่านั้น');
  if (!legacyPepper_())
    warnings.push('ยังไม่ได้ตั้ง LEGACY_PEPPER · รหัสผ่านเดิมจะตรวจสอบไม่ได้ → PASSWORD_RESET_REQUIRED');

  t.rows.forEach(function (r) {
    stats.read++;
    var uname = String(accField_(r, 'Username') || '').trim();
    if (!uname) { stats.skipped++; return; }
    var key = uname.toLowerCase();
    if (seen[key]) {
      stats.unmatched++;
      warnings.push('ชื่อผู้ใช้ซ้ำ: ' + uname + ' (แถว ' + r._row + ')');
      if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
        EntityType: 'OTHER', SourceSheet: 'Accounts', SourceRow: r._row,
        SourceRecordID: uname, SourceName: uname, MatchReason: 'DUPLICATE_USERNAME' });
      return;
    }
    seen[key] = true;

    var rawRole = accField_(r, 'Role');
    var role = normalizeRole_(rawRole);
    var hasV3 = !!String(accField_(r, 'PasswordHash') || '').trim();
    var hasLegacy = !!String(accField_(r, 'LegacyHash') || '').trim() &&
                    !!String(accField_(r, 'Salt') || '').trim();
    var pwState = hasV3 ? 'V3_HASH'
                : (hasLegacy ? (legacyPepper_() ? 'LEGACY_VERIFIABLE' : 'PASSWORD_RESET_REQUIRED')
                             : 'PASSWORD_RESET_REQUIRED');

    if (!role) {
      stats.unmatched++;
      warnings.push('บทบาทไม่รู้จัก: "' + rawRole + '" (' + uname + ')');
      if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
        EntityType: 'OTHER', SourceSheet: 'Accounts', SourceRow: r._row,
        SourceRecordID: uname, SourceName: String(rawRole),
        MatchReason: 'UNKNOWN_ROLE' });
    } else {
      stats.matched++;
    }
    if (pwState === 'PASSWORD_RESET_REQUIRED') {
      warnings.push('ต้องตั้งรหัสผ่านใหม่: ' + uname);
      if (!dryRun) addUnmatched_({ ImportID: imp.ImportID, BatchID: imp.BatchID,
        EntityType: 'OTHER', SourceSheet: 'Accounts', SourceRow: r._row,
        SourceRecordID: uname, SourceName: uname, MatchReason: 'PASSWORD_RESET_REQUIRED' });
    }

    var activeRaw = accField_(r, 'Active');
    var active = (activeRaw === '' || activeRaw == null) ? true : truthy_(activeRaw);

    preview.push({ row: r._row, Username: uname, roleFrom: rawRole, roleTo: role || 'UNKNOWN',
                   Active: active, password: pwState });

    if (!dryRun) {
      t.sheet.getRange(r._row, C.Username).setValue(uname);
      if (role) t.sheet.getRange(r._row, C.Role).setValue(role);
      t.sheet.getRange(r._row, C.Active).setValue(active);
      if (!String(accField_(r, 'DisplayName') || '').trim())
        t.sheet.getRange(r._row, C.DisplayName).setValue(uname);
      var ts = accField_(r, 'UpdatedAt');
      t.sheet.getRange(r._row, C.UpdatedAt).setValue(ts ? String(ts) : nowIso_());
      // PasswordHash intentionally left as-is: never fabricated from passHash
      stats.inserted++;
    }
  });

  finishImport_(imp.ImportID, stats, dryRun ? IMPORT_STATUS.SUCCESS : importStatusFor_(stats),
                (dryRun ? 'DRY RUN — no data written · ' : '') + 'ACCOUNTS_NORMALIZE · ' +
                warnings.length + ' warning(s)');
  invalidateCache_('Accounts');
  return { dryRun: dryRun, importId: imp.ImportID, stats: stats,
           legacyPepperPresent: !!legacyPepper_(),
           duplicateHeaders: Object.keys(dupHeaders), warnings: warnings, preview: preview };
}

/* ============================================================
 * 18. WEB APP ENTRY POINTS  (GET = read only · POST = mutations)
 * ============================================================ */

var GET_ACTIONS  = ['health', 'me', 'members', 'dashboard', 'statusTypes', 'settings', 'previewEngine'];
var POST_ACTIONS = ['login', 'updateMember', 'saveDecision', 'clearDecision', 'saveSetting',
                    'migrateMembers', 'migrateAttendance', 'migrateStatuses', 'migrateDecisions',
                    'migrateAccounts'];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'health';
  if (GET_ACTIONS.indexOf(action) < 0)
    return json_({ ok: false, error: 'METHOD_NOT_ALLOWED: "' + action + '" ต้องเรียกผ่าน POST' });
  return handle_(action, p, 'GET');
}

function doPost(e) {
  var body = {};
  try {
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
  } catch (err) { return json_({ ok: false, error: 'BAD_JSON_BODY' }); }
  var p = {};
  Object.keys((e && e.parameter) || {}).forEach(function (k) { p[k] = e.parameter[k]; });
  Object.keys(body).forEach(function (k) { p[k] = body[k]; });
  var action = p.action || '';
  if (GET_ACTIONS.indexOf(action) < 0 && POST_ACTIONS.indexOf(action) < 0)
    return json_({ ok: false, error: 'UNKNOWN_ACTION: ' + action });
  return handle_(action, p, 'POST');
}

function asObject_(v) {
  if (!v) return {};
  if (typeof v === 'object') return v;
  try { return JSON.parse(v); } catch (e) { return {}; }
}

function handle_(action, p, method) {
  try {
    if (action === 'health') {
      var rm = runtimeMode_(true);
      return json_({
        ok: true, mode: rm.mode, normalized: rm.normalized, pending: rm.pending,
        legacy: validateLegacySchema_(), runtime: validateRuntimeSchema_(),
        ageBands: safeBands_(), at: nowIso_()
      });
    }
    if (action === 'login') return json_({ ok: true, data: login_(p.username, p.password) });

    var session = requireSession_(p.token);

    switch (action) {
      case 'me':          return json_({ ok: true, data: session });
      case 'members':     return json_({ ok: true, data: listMembers_(session) });
      case 'dashboard':   return json_({ ok: true, data: dashboardCounts_(session) });
      case 'statusTypes': return json_({ ok: true, data: publicStatusTypes_(session) });
      case 'settings':    return json_({ ok: true, data: publicSettings_(session) });

      case 'updateMember':
        return json_({ ok: true, data: updateMember_(session, p.memberKey || p.memberId, asObject_(p.patch)) });
      case 'saveDecision':
        return json_({ ok: true, data: saveDecision_(session, asObject_(p.decision)) });
      case 'clearDecision':
        return json_({ ok: true, data: clearDecision_(session, p.memberId, p.note) });
      case 'saveSetting':
        return json_({ ok: true, data: saveSetting_(session, p.key, p.value, p.valueType) });

      case 'previewEngine':
        authorizeRole_(session, 'migrate');
        return json_({ ok: true, data: previewStatusEngine(Number(p.limit) || 20) });
      case 'migrateMembers':
      case 'migrateStatuses':
        authorizeRole_(session, 'migrate');
        return json_({ ok: true, data: migrateMembersToV3(p.dryRun !== false && p.dryRun !== 'false') });
      case 'migrateAttendance':
        authorizeRole_(session, 'migrate');
        return json_({ ok: true, data: migrateAttendanceToMemberID(p.dryRun !== false && p.dryRun !== 'false') });
      case 'migrateDecisions':
        authorizeRole_(session, 'migrate');
        return json_({ ok: true, data: migrateLegacyDecisions(p.dryRun !== false && p.dryRun !== 'false') });
      case 'migrateAccounts':
        authorizeRole_(session, 'migrate');
        return json_({ ok: true, data: migrateAccountsToV3(p.dryRun !== false && p.dryRun !== 'false') });

      default:
        return json_({ ok: false, error: 'UNKNOWN_ACTION: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function safeBands_() { try { return validateAgeBands_(); } catch (e) { return { ok: false, problems: [String(e)] }; } }

/* ============================================================
 * 19. SETUP / DIAGNOSTIC HELPERS  (run manually from the editor)
 * ============================================================ */

/** Safe to run FIRST. Reads only — reports both schema stages and the detected headers. */
function diagnose() {
  var out = {
    legacy: validateLegacySchema_(),
    runtime: validateRuntimeSchema_(),
    detected: {}
  };
  ['Members', 'Attendance', 'Decisions', 'Settings', 'StatusTypes',
   'EngagementLog', 'UnmatchedRecords', 'ImportLog', 'Accounts'].forEach(function (n) {
    try { out.detected[n] = readTable_(n, true).headers; }
    catch (e) { out.detected[n] = String(e.message || e); }
  });
  out.attendanceDateColumn = (function () {
    try { return detectAttendanceDateColumn_(true); }
    catch (e) { return { header: null, method: 'ERROR', error: String(e.message || e) }; }
  })();
  out.attendanceNameColumn = (function () {
    try { return pickHeader_(readTable_('Attendance'), ALIAS.Attendance.Name); }
    catch (e) { return null; }
  })();
  out.ageBands = safeBands_();
  out.accounts = (function () {
    try {
      var t = readTable_('Accounts', true);
      var dup = [];
      t.headers.forEach(function (h, i) {
        if (h && t.headers.indexOf(h) !== i && dup.indexOf(h) < 0) dup.push(h);
      });
      var roles = {};
      t.rows.forEach(function (r) {
        var raw = String(accField_(r, 'Role') || '');
        var k = raw + ' → ' + (normalizeRole_(raw) || 'UNKNOWN');
        roles[k] = (roles[k] || 0) + 1;
      });
      return {
        legacyDetected: {
          username: pickHeader_(t, ALIAS.Accounts.Username),
          legacyHash: pickHeader_(t, ALIAS.Accounts.LegacyHash),
          salt: pickHeader_(t, ALIAS.Accounts.Salt),
          role: pickHeader_(t, ALIAS.Accounts.Role),
          active: pickHeader_(t, ALIAS.Accounts.Active)
        },
        runtimeReady: RUNTIME_SCHEMA.Accounts.every(function (h) { return hasHeader_(t, h); }),
        duplicateHeaders: dup,
        accountCount: t.rows.length,
        roleMapping: roles,
        legacyPepperPresent: !!legacyPepper_(),
        tokenPepperPresent: !!props_().getProperty('TOKEN_PEPPER')
      };
    } catch (e) { return { error: String(e.message || e) }; }
  })();
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

/** Writes the V3 rule/age rows into Settings using SettingKey / SettingValue. */
function setupSettingsRows() {
  var t = readTable_('Settings', true);
  var c = settingsCols_(t);
  if (!c.key || !c.value) throw new Error('SCHEMA: Settings needs SettingKey / SettingValue');
  var have = {};
  t.rows.forEach(function (r) { have[String(r[c.key]).trim()] = true; });
  var wanted = [
    ['SUNDAY_SCHOOL_MIN_AGE', 3, 'NUMBER', 'AGE', 'อายุต่ำสุดกลุ่มรวีวารศึกษา'],
    ['SUNDAY_SCHOOL_MAX_AGE', 10, 'NUMBER', 'AGE', 'อายุสูงสุดกลุ่มรวีวารศึกษา'],
    ['YOUTH_MIN_AGE', 11, 'NUMBER', 'AGE', 'อายุต่ำสุดกลุ่มอนุชน'],
    ['YOUTH_MAX_AGE', 24, 'NUMBER', 'AGE', 'อายุสูงสุดกลุ่มอนุชน'],
    ['ADULT_MIN_AGE', 25, 'NUMBER', 'AGE', 'อายุต่ำสุดกลุ่มผู้ใหญ่'],
    ['ADULT_MAX_AGE', 59, 'NUMBER', 'AGE', 'อายุสูงสุดกลุ่มผู้ใหญ่'],
    ['SENIOR_MIN_AGE', 60, 'NUMBER', 'AGE', 'อายุต่ำสุดกลุ่มผู้สูงวัย'],
    ['FULL2_ABSENCE_MONTHS', 12, 'NUMBER', 'RULE', 'ขาดนมัสการกี่เดือนจึงเป็นสมาชิกสมบูรณ์ 2'],
    ['FULL2_REVIEW_MONTHS', 36, 'NUMBER', 'RULE', 'ขาดนมัสการกี่เดือนจึงควรนำเข้าพิจารณาเป็นสำรอง'],
    ['FOLLOWUP_WARNING_MONTHS', 6, 'NUMBER', 'RULE', 'เกณฑ์ติดตาม (ไม่เปลี่ยนสถานภาพ)'],
    ['COUNCIL_REVIEW_YEARS', 1, 'NUMBER', 'RULE', 'ปีที่ครบเกณฑ์พิจารณา (สภาคริสตจักร)'],
    ['OUTSIDE_COUNCIL_REVIEW_YEARS', 2, 'NUMBER', 'RULE', 'ปีที่ครบเกณฑ์พิจารณา (นอกสภา)'],
    ['DEFAULT_NEXT_REVIEW_MONTHS', 12, 'NUMBER', 'RULE', 'กำหนดทบทวนมติถัดไป'],
    ['REQUIRE_DECISION_REASON', true, 'BOOLEAN', 'RULE', 'บังคับระบุเหตุผลของมติ'],
    ['GIVING_AMOUNT_VISIBLE', false, 'BOOLEAN', 'SECURITY', 'แสดงจำนวนเงินถวาย'],
    ['STRONG_MIN_VISITS_12M', 10, 'NUMBER', 'ANALYTICS', 'จำนวนครั้งขั้นต่ำใน 12 เดือน'],
    ['STRONG_CONTINUITY_PERCENT', 50, 'NUMBER', 'ANALYTICS', 'เปอร์เซ็นต์ความต่อเนื่อง'],
    ['STRONG_MIN_MEMBERSHIP_YEARS', 1, 'NUMBER', 'ANALYTICS', 'ปีขั้นต่ำการเป็นสมาชิก']
  ];
  var added = 0;
  wanted.forEach(function (w) {
    if (have[w[0]]) return;
    var payload = { ValueType: w[2], Category: w[3], Description: w[4],
                    EditableByAdmin: true, Active: true,
                    UpdatedBy: 'setup', UpdatedAt: nowIso_() };
    payload[c.key] = w[0]; payload[c.value] = w[1];
    appendRow_('Settings', payload);
    added++;
  });
  invalidateCache_('Settings'); _settings = null;
  return { added: added, ageBands: validateAgeBands_() };
}

/** Print a password hash to paste into Accounts.PasswordHash. */
function makePasswordHash(plain) { var h = hash_(String(plain)); Logger.log(h); return h; }

/** Convenience: all three migrations as DRY RUN, in the correct order. */
function dryRunAllMigrations() {
  var out = {
    step1_members: migrateMembersToV3(true),
    step2_attendance: migrateAttendanceToMemberID(true),
    step3_decisions: migrateLegacyDecisions(true),
    step4_accounts: migrateAccountsToV3(true)
  };
  out.summary = {
    members: out.step1_members.report,
    attendance: out.step2_attendance.stats,
    decisions: out.step3_decisions.stats,
    accounts: out.step4_accounts.stats,
    unknownStatusRows: out.step1_members.unknownStatusRows,
    warnings: [].concat(out.step1_members.warnings, out.step4_accounts.warnings)
  };
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

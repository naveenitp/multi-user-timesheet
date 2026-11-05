// Multi-user timesheet backend (separate project from the single-user version).
//
//  This spreadsheet is the CONTROL sheet. It holds:
//    Users        - accounts (passwords are stored as salted hashes)
//    Masters      - business names, products, status, source (shared, edited by admins only)
//    SQL Library  - admin only
//  Every person gets their OWN Google Sheet file ("Timesheet - Name") with their Timesheet tab,
//  month tabs and Templates. People never receive the file ids, the Masters sheet or the SQL Library.
//
//  Set up: 1) edit ADMIN_USER / ADMIN_PASS below  2) run setupAdmin once  3) Deploy > New deployment > Web app
//          (Execute as: Me, Who has access: Anyone)  4) paste the Web app URL in the page (Server address).

const ADMIN_USER = 'admin';
const ADMIN_PASS = 'ChangeMe#2026';

const HEAD = ['ID','SL.no','Month','Date','Code','Source','Product','Transaction Start Time','Transaction End time','Total Duration','CONTACT NO','Business Name','Issue/Request','Ticket No','Status','Remarks'];
const KEYS = ['id','sl','month','date','code','src','prod','s','e','dur','contact','biz','issue','ticket','stat','rem'];
const WIDTHS = [40,48,78,125,64,84,110,70,70,76,110,190,100,86,100,330];
const ALIGN  = ['c','c','c','c','c','c','l','c','c','c','c','l','c','c','c','l'];
const HDR = 4, FIRST = 5, N = HEAD.length;          // month tabs: title rows 1-2, header row 4, data from row 5
const MFIRST = 2;                                    // Timesheet tab: header row 1, data from row 2
const MHEAD = ['Type','Name','Code'];
const MON  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const LONG = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DEFAULTS = { stat: ['Open','In Progress','Resolved','Closed'], src: ['Call','Email','WhatsApp','Chat'] };
const NAVY = '#1f3864', BLUE = '#2f5597', PALE = '#eaf0fa', LINE = '#d0d7e2';

// ---------- helpers ----------
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function str_(v, n) { return String(v == null ? '' : v).slice(0, n || 2000); }
function v_(v, k) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz_(), k === 'date' ? 'yyyy-MM-dd' : 'HH:mm');
  return String(v);
}
function isMonth_(n) { return /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{4}$/.test(n); }
function monthKey_(n) { const p = n.split(' '); return +p[1] * 12 + MON.indexOf(p[0]); }
function monthName_(d) { const p = String(d).split('-'); return MON[+p[1] - 1] + ' ' + p[0]; }
function monthSheets_() {
  return ss_().getSheets().filter(s => isMonth_(s.getName())).sort((a, b) => monthKey_(a.getName()) - monthKey_(b.getName()));
}

// ---------- styling ----------
function style_(sh, name) {
  const p = name.split(' '), title = 'Timesheet  |  ' + LONG[MON.indexOf(p[0])] + ' ' + p[1];
  sh.getRange(HDR, 1, 1, N).setValues([HEAD]);                         // header first, so data rows are safe even if styling fails
  const step = f => { try { f(); } catch (e) { Logger.log('style: ' + e); } };
  const rows = sh.getMaxRows() - FIRST + 1;
  step(() => sh.setHiddenGridlines(true));
  step(() => sh.getRange(1, 1, sh.getMaxRows(), N).setFontFamily('Calibri').setFontSize(10).setVerticalAlignment('top'));
  step(() => sh.getRange(FIRST, 1, rows, N).setNumberFormat('@'));      // keep dates and times as typed
  step(() => { const t = sh.getRange(1, 2, 1, N - 1).merge(); t.setValue(title).setBackground(NAVY).setFontColor('#ffffff').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle'); sh.setRowHeight(1, 38); });
  step(() => {
    sh.getRange(2, 2, 1, N - 1).setBackground(PALE).setVerticalAlignment('middle');
    sh.setRowHeight(2, 28);
    const lab = (c, t) => sh.getRange(2, c).setValue(t).setFontColor('#5b6b86').setFontWeight('bold').setHorizontalAlignment('right');
    const val = (c, f) => sh.getRange(2, c).setFormula(f).setFontColor(NAVY).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('left');
    lab(2, 'Entries');    val(3, '=COUNTA(D5:D)');
    lab(5, 'Total time'); val(6, '=LET(m,ARRAYFORMULA(SUM(IFERROR(VALUE(LEFT(J5:J,FIND(":",J5:J)-1)),0)*60+IFERROR(VALUE(MID(J5:J,FIND(":",J5:J)+1,2)),0))),INT(m/60)&"h "&TEXT(MOD(m,60),"00")&"m")');
    lab(9, 'Resolved');   val(10, '=COUNTIF(O5:O,"*esolv*")+COUNTIF(O5:O,"*losed*")');
    lab(12, 'Open');      val(13, '=COUNTIF(O5:O,"Open")');
    sh.setRowHeight(3, 8);
  });
  step(() => {
    sh.getRange(HDR, 2, 1, N - 1).setBackground(BLUE).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, true, true, true, true, false, LINE, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(HDR, 34);
  });
  step(() => { WIDTHS.forEach((w, i) => sh.setColumnWidth(i + 1, w)); sh.hideColumns(1); });
  step(() => ALIGN.forEach((a, i) => { if (i) sh.getRange(FIRST, i + 1, rows, 1).setHorizontalAlignment(a === 'c' ? 'center' : 'left'); }));
  step(() => { sh.getRange(FIRST, 12, rows, 1).setWrap(true); sh.getRange(FIRST, 16, rows, 1).setWrap(true); });
  step(() => sh.setFrozenRows(HDR));
  step(() => {
    const st = sh.getRange(FIRST, 15, rows, 1), band = sh.getRange(FIRST, 2, rows, N - 1);
    const R = SpreadsheetApp.newConditionalFormatRule;
    sh.setConditionalFormatRules([
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O5),"resolv|clos|done|complet")').setBackground('#c6efce').setFontColor('#006100').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O5),"progress|pending|hold|wait|wip|working")').setBackground('#ffeb9c').setFontColor('#7f6000').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O5),"open|new")').setBackground('#ffc7ce').setFontColor('#9c0006').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=AND($D5<>"",ISEVEN(ROW()))').setBackground('#f6f8fc').setRanges([band]).build()
    ]);
  });
  step(() => sh.setTabColor(NAVY));
}
function styleMasters_(sh) {
  const step = f => { try { f(); } catch (e) { Logger.log('style: ' + e); } };
  step(() => sh.getRange(1, 1, 1, 3).setBackground(BLUE).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center'));
  step(() => { [110, 230, 380].forEach((w, i) => sh.setColumnWidth(i + 1, w)); sh.setFrozenRows(1); sh.setTabColor('#f59e0b'); });
  step(() => sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setFontColor('#6b7280').setFontSize(9));
}

// ---------- sheets ----------
function monthSheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  try {
    const earlier = monthSheets_().filter(s => s.getName() !== name && monthKey_(s.getName()) < monthKey_(name)).length;
    ss.setActiveSheet(sh); ss.moveActiveSheet(earlier + 1 + (ss.getSheetByName('Timesheet') ? 1 : 0));
  } catch (e) {}
  style_(sh, name);
  return sh;
}
function mastersSheet_() {
  const ss = ctl_();
  let sh = ss.getSheetByName('Masters');
  if (!sh) {
    sh = ss.insertSheet('Masters');
    sh.getRange(1, 1, sh.getMaxRows(), 3).setNumberFormat('@');
    sh.getRange(1, 1, 1, 3).setValues([MHEAD]);
    styleMasters_(sh);
  }
  return sh;
}
function find_(id) {
  for (const sh of monthSheets_()) {
    const n = sh.getLastRow() - HDR;
    if (n < 1) continue;
    const ids = sh.getRange(FIRST, 1, n, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return { sh: sh, row: FIRST + i };
  }
  return null;
}
function renumber_(sh) {
  const n = sh.getLastRow() - HDR;
  if (n > 0) sh.getRange(FIRST, 2, n, 1).setValues(Array.from({ length: n }, (_, i) => [i + 1]));
}
function ensureRows_(sh, rowNo) {
  const need = rowNo - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), Math.max(need, 100));
}
function border_(sh, row, rows) {
  try { sh.getRange(row, 2, rows, N - 1).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID); } catch (e) {}
}

// ---------- Timesheet tab (all data) ----------
function styleMaster_(sh) {
  const step = f => { try { f(); } catch (e) { Logger.log('style: ' + e); } };
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  step(() => sh.getRange(1, 1, 1, N).setBackground(BLUE).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, false, LINE, SpreadsheetApp.BorderStyle.SOLID));
  step(() => sh.setRowHeight(1, 34));
  step(() => sh.getRange(1, 1, sh.getMaxRows(), N).setFontFamily('Calibri').setFontSize(10).setVerticalAlignment('top'));
  step(() => { WIDTHS.forEach((w, i) => sh.setColumnWidth(i + 1, i === 0 ? 70 : w)); sh.setFrozenRows(1); });
  step(() => sh.getRange(MFIRST, 1, rows, 1).setFontColor('#9aa3b5').setFontSize(8));
  step(() => ALIGN.forEach((a, i) => { if (i) sh.getRange(MFIRST, i + 1, rows, 1).setHorizontalAlignment(a === 'c' ? 'center' : 'left'); }));
  step(() => { sh.getRange(MFIRST, 12, rows, 1).setWrap(true); sh.getRange(MFIRST, 16, rows, 1).setWrap(true); });
  step(() => {
    const st = sh.getRange(MFIRST, 15, rows, 1), band = sh.getRange(MFIRST, 2, rows, N - 1), R = SpreadsheetApp.newConditionalFormatRule;
    sh.setConditionalFormatRules([
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O2),"resolv|clos|done|complet")').setBackground('#c6efce').setFontColor('#006100').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O2),"progress|pending|hold|wait|wip|working")').setBackground('#ffeb9c').setFontColor('#7f6000').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O2),"open|new")').setBackground('#ffc7ce').setFontColor('#9c0006').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=AND($D2<>"",ISEVEN(ROW()))').setBackground('#f6f8fc').setRanges([band]).build()
    ]);
  });
  step(() => sh.setTabColor('#16a34a'));
}
function tsSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName('Timesheet');
  if (!sh) {
    const def = ss.getSheets().length === 1 ? ss.getSheets()[0] : null;
    if (def && def.getLastRow() === 0) { def.setName('Timesheet'); sh = def; } else sh = ss.insertSheet('Timesheet', 0);
  }
  if (!String(sh.getRange(1, 1).getValue())) {
    sh.getRange(1, 1, 1, N).setValues([HEAD]);
    sh.getRange(MFIRST, 1, Math.max(sh.getMaxRows() - 1, 1), N).setNumberFormat('@');
    styleMaster_(sh);
  }
  return sh;
}
function tsFind_(ts, id) {
  const n = ts.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = ts.getRange(MFIRST, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === id) return MFIRST + i;
  return 0;
}
function tsRows_(ts) {
  const n = ts.getLastRow() - 1;
  if (n < 1) return [];
  const head = ts.getRange(1, 1, 1, Math.max(ts.getLastColumn(), 1)).getValues()[0].map(String);
  const idx = HEAD.map((h, i) => head.indexOf(h) >= 0 ? head.indexOf(h) : i);
  return ts.getRange(MFIRST, 1, n, Math.max(ts.getLastColumn(), N)).getValues()
    .filter(r => String(r[idx[0]]))
    .map(r => { const o = {}; KEYS.forEach((k, i) => o[k] = v_(r[idx[i]], k)); return o; });
}
function tsRenumber_(ts) {
  const n = ts.getLastRow() - 1;
  if (n > 0) ts.getRange(MFIRST, 2, n, 1).setValues(Array.from({ length: n }, (_, i) => [i + 1]));
}

// ---------- entries ----------
function upsert_(r) {
  if (!r || !r.id || !/^\d{4}-\d{2}-\d{2}$/.test(String(r.date || ''))) throw 'Invalid entry.';
  const id = str_(r.id, 40), row = {};
  KEYS.forEach(k => row[k] = str_(r[k]));
  row.id = id;
  row.month = Utilities.formatDate(new Date(r.date + 'T00:00:00'), tz_(), 'MMM yyyy');
  // 1) the Timesheet tab holds every entry
  const ts = tsSheet_();
  let at = tsFind_(ts, id);
  if (!at) at = Math.max(ts.getLastRow() + 1, MFIRST);
  row.sl = at - 1;
  ensureRows_(ts, at);
  ts.getRange(at, 1, 1, N).setNumberFormat('@').setValues([KEYS.map(k => row[k])]);
  try { ts.getRange(at, 1, 1, N).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID); } catch (e) {}
  // 2) and the month tab shows it too
  mirror_(row);
}
function mirror_(row) {
  const target = monthName_(row.date), found = find_(row.id);
  let sh, rowNo;
  if (found && found.sh.getName() === target) { sh = found.sh; rowNo = found.row; }
  else {
    if (found) { found.sh.deleteRow(found.row); renumber_(found.sh); }   // the date moved to another month
    sh = monthSheet_(target); rowNo = Math.max(sh.getLastRow() + 1, FIRST);
  }
  const out = {}; KEYS.forEach(k => out[k] = row[k]);
  out.sl = rowNo - HDR;
  ensureRows_(sh, rowNo);
  sh.getRange(rowNo, 1, 1, N).setNumberFormat('@').setValues([KEYS.map(k => out[k])]);
  border_(sh, rowNo, 1);
}
function del_(id) {
  id = String(id);
  const ts = tsSheet_(), at = tsFind_(ts, id);
  if (at) { ts.deleteRow(at); tsRenumber_(ts); }
  const f = find_(id);
  if (f) { f.sh.deleteRow(f.row); renumber_(f.sh); }
}

// ---------- masters (business names, products, ...) ----------
function masters_() {
  const ms = mastersSheet_();
  if (ms.getLastRow() < 2) {
    const seed = [];
    for (const t in DEFAULTS) DEFAULTS[t].forEach(n => seed.push([t, n, '']));
    ms.getRange(2, 1, seed.length, 3).setValues(seed);
  }
  const m = { prod: [], biz: [], stat: [], src: [] };
  ms.getRange(2, 1, ms.getLastRow() - 1, 3).getValues().forEach(r => {
    const t = r[0]; if (!m[t]) return;
    let x = String(r[1]);
    if (t === 'biz') x = { name: x, code: String(r[2]) };
    m[t].push(x);
  });
  return m;
}
function setMasters_(masters) {
  const ms = mastersSheet_(), rows = [];
  ['prod', 'biz', 'stat', 'src'].forEach(t => (masters && masters[t] || []).slice(0, 2000).forEach(x => {
    if (t === 'biz') rows.push([t, str_(x.name, 200), str_(x.code, 30)]);
    else rows.push([t, str_(x, 200), '']);
  }));
  if (ms.getLastRow() > 1) ms.getRange(2, 1, ms.getLastRow() - 1, 3).clearContent();
  if (rows.length) ms.getRange(2, 1, rows.length, 3).setValues(rows);
}

// ---------- SQL Library ----------
const SHEAD = ['ID','Title','Category','Description','SQL','Parameters (JSON)','Expected output','Precautions','Favourite','Updated'];
function styleSql_(sh) {
  const step = f => { try { f(); } catch (e) { Logger.log('style: ' + e); } };
  step(() => sh.getRange(1, 1, 1, SHEAD.length).setBackground(BLUE).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true));
  step(() => sh.setRowHeight(1, 34));
  step(() => sh.getRange(1, 1, sh.getMaxRows(), SHEAD.length).setFontFamily('Calibri').setFontSize(10).setVerticalAlignment('top').setWrap(true));
  step(() => sh.getRange(2, 5, sh.getMaxRows() - 1, 1).setFontFamily('Consolas').setFontSize(9));
  step(() => { [70, 200, 140, 260, 480, 240, 260, 260, 80, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w)); sh.setFrozenRows(1); sh.setTabColor('#7c3aed'); });
}
function sqlSheet_() {
  const ss = ctl_();
  let sh = ss.getSheetByName('SQL Library');
  if (!sh) {
    sh = ss.insertSheet('SQL Library');
    sh.getRange(1, 1, 1, SHEAD.length).setValues([SHEAD]);
    sh.getRange(2, 1, sh.getMaxRows() - 1, SHEAD.length).setNumberFormat('@');
    styleSql_(sh);
  }
  return sh;
}
function sqlRows_() {
  const sh = sqlSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, SHEAD.length).getValues().filter(r => String(r[0])).map(r => {
    let pars = [];
    try { pars = JSON.parse(String(r[5] || '[]')); } catch (e) {}
    return { id: String(r[0]), title: String(r[1]), cat: String(r[2]), desc: String(r[3]), sql: String(r[4]), pars: pars, exp: String(r[6]), pre: String(r[7]), fav: String(r[8]) === 'yes', upd: +r[9] || 0 };
  });
}
function saveSql_(q) {
  if (!q || !q.id || !String(q.title || '').trim() || !String(q.sql || '').trim()) throw 'Invalid query.';
  const sh = sqlSheet_(), id = str_(q.id, 40), n = sh.getLastRow() - 1;
  let at = 0;
  if (n > 0) { const ids = sh.getRange(2, 1, n, 1).getValues(); for (let i = 0; i < n; i++) if (String(ids[i][0]) === id) { at = 2 + i; break; } }
  if (!at) at = Math.max(sh.getLastRow() + 1, 2);
  const pars = (Array.isArray(q.pars) ? q.pars : []).slice(0, 40).map(p => ({ n: str_(p.n, 80), v: str_(p.v, 300), d: str_(p.d, 300) }));
  ensureRows_(sh, at);
  sh.getRange(at, 1, 1, SHEAD.length).setNumberFormat('@').setValues([[id, str_(q.title, 200), str_(q.cat, 100), str_(q.desc, 3000), str_(q.sql, 40000), JSON.stringify(pars), str_(q.exp, 3000), str_(q.pre, 3000), q.fav ? 'yes' : 'no', String(Date.now())]]);
  try { sh.getRange(at, 1, 1, SHEAD.length).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID); } catch (e) {}
}
function delSql_(id) {
  const sh = sqlSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === String(id)) { sh.deleteRow(2 + i); return; }
}


// ---------- accounts, sessions and access control ----------
const UHEAD = ['Username','Name','Role','Hash','Salt','Active','SheetId','Email','Shared','Created'];
const TOKEN_TTL = 21600;   // 6 hours (the maximum CacheService allows)
const MAX_FAILS = 5;       // wrong passwords before a 15 minute lock
let ME = null, USER_SS = null;   // set for every request from the verified login, never from the request body

function ctl_() { return SpreadsheetApp.getActive(); }
function ss_() {
  if (!ME) throw 'No user.';
  if (!USER_SS) USER_SS = SpreadsheetApp.openById(ME.sheetId);
  return USER_SS;
}
function usersSheet_() {
  const ss = ctl_();
  let sh = ss.getSheetByName('Users');
  if (!sh) {
    sh = ss.insertSheet('Users');
    sh.getRange(1, 1, sh.getMaxRows(), UHEAD.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, UHEAD.length).setValues([UHEAD]);
    try { sh.getRange(1, 1, 1, UHEAD.length).setBackground(BLUE).setFontColor('#ffffff').setFontWeight('bold'); sh.setFrozenRows(1); sh.setTabColor('#dc2626'); } catch (e) {}
  }
  return sh;
}
function users_() {
  const sh = usersSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, UHEAD.length).getValues().map((r, i) => ({
    row: i + 2, u: String(r[0]), name: String(r[1]), role: String(r[2]), hash: String(r[3]), salt: String(r[4]),
    active: String(r[5]) === 'yes', sheetId: String(r[6]), email: String(r[7]), shared: String(r[8]) === 'yes'
  }));
}
function hash_(pw, salt) {
  let h = salt + '|' + pw;
  for (let i = 0; i < 1000; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
function same_(a, b) {
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function ver_(u) { return PropertiesService.getScriptProperties().getProperty('ver_' + u) || '0'; }
function bump_(u) { PropertiesService.getScriptProperties().setProperty('ver_' + u, String(+ver_(u) + 1)); }
function newToken_(u) {
  const t = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('t_' + t, JSON.stringify({ u: u, v: ver_(u) }), TOKEN_TTL);
  return t;
}
function auth_(token) {
  if (!token) throw 'auth';
  const raw = CacheService.getScriptCache().get('t_' + token);
  if (!raw) throw 'auth';
  const s = JSON.parse(raw), me = users_().find(x => x.u === s.u);
  if (!me || !me.active || s.v !== ver_(s.u)) throw 'auth';
  return me;
}
function need_() { if (!ME || ME.role !== 'admin') throw 'Admins only.'; }
function login_(b) {
  const u = String(b.user || '').toLowerCase().trim(), pass = String(b.pass || ''), cache = CacheService.getScriptCache(), k = 'f_' + u;
  if (+cache.get(k) >= MAX_FAILS) throw 'Too many failed attempts. Try again in 15 minutes.';
  const me = users_().find(x => x.u === u);
  const ok = me ? (same_(me.hash, hash_(pass, me.salt)) && me.active) : (hash_(pass, 'x'), false);
  if (!ok) { cache.put(k, String(+cache.get(k) + 1), 900); throw 'Incorrect username or password.'; }
  cache.remove(k);
  return { token: newToken_(me.u), user: { username: me.u, name: me.name || me.u, role: me.role } };
}
function setPw_(me, pw) {
  if (String(pw || '').length < 8) throw 'Password must be at least 8 characters.';
  const salt = Utilities.getUuid();
  usersSheet_().getRange(me.row, 4, 1, 2).setValues([[hash_(pw, salt), salt]]);
  bump_(me.u);   // signs out every session of this person
}
function createUserSheet_(u, name) {
  const ns = SpreadsheetApp.create('Timesheet - ' + (name || u));
  const keepMe = ME, keepSS = USER_SS;
  ME = { sheetId: ns.getId() }; USER_SS = ns;
  try { tsSheet_(); tplSheet_(); } finally { ME = keepMe; USER_SS = keepSS; }
  return ns.getId();
}
function addUser_(u, name, pw, role) {
  u = String(u || '').toLowerCase().trim();
  if (!/^[a-z0-9._-]{3,30}$/.test(u)) throw 'Username must be 3-30 characters: letters, numbers, dot, dash or underscore.';
  if (String(pw || '').length < 8) throw 'Password must be at least 8 characters.';
  if (role !== 'admin' && role !== 'user') throw 'Invalid role.';
  if (users_().some(x => x.u === u)) throw 'That username already exists.';
  const sheetId = createUserSheet_(u, str_(name, 60));
  const sh = usersSheet_(), salt = Utilities.getUuid();
  sh.getRange(sh.getLastRow() + 1, 1, 1, UHEAD.length).setValues([[u, str_(name, 60) || u, role, hash_(pw, salt), salt, 'yes', sheetId, '', 'no', new Date().toISOString()]]);
}
// Run once from the editor after editing ADMIN_USER / ADMIN_PASS at the top.
function setupAdmin() {
  if (users_().some(x => x.u === ADMIN_USER)) { Logger.log('Admin already exists. Use resetAdminPassword() if you forgot the password.'); return; }
  addUser_(ADMIN_USER, 'Administrator', ADMIN_PASS, 'admin');
  mastersSheet_(); sqlSheet_();
  Logger.log('Admin created. Now deploy the web app.');
}
// Run from the editor if the admin password is lost: sets it back to ADMIN_PASS and clears any lock.
function resetAdminPassword() {
  const me = users_().find(x => x.u === ADMIN_USER);
  if (!me) { setupAdmin(); return; }
  setPw_(me, ADMIN_PASS);
  usersSheet_().getRange(me.row, 6).setValue('yes');
  CacheService.getScriptCache().remove('f_' + ADMIN_USER);
}

// ---------- each person's templates (kept in their own file) ----------
function tplSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName('Templates');
  if (!sh) {
    sh = ss.insertSheet('Templates');
    sh.getRange(1, 1, sh.getMaxRows(), 2).setNumberFormat('@');
    sh.getRange(1, 1, 1, 2).setValues([['Name', 'Template (JSON)']]);
    try { sh.getRange(1, 1, 1, 2).setBackground(BLUE).setFontColor('#ffffff').setFontWeight('bold'); sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 420); sh.setFrozenRows(1); sh.setTabColor('#f59e0b'); } catch (e) {}
  }
  return sh;
}
function tplRows_() {
  const sh = tplSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const out = [];
  sh.getRange(2, 1, n, 2).getValues().forEach(r => { try { const t = JSON.parse(String(r[1])); if (t && t.name) out.push(t); } catch (e) {} });
  return out;
}
function setTpl_(list) {
  const sh = tplSheet_(), rows = (Array.isArray(list) ? list : []).slice(0, 300).map(t => [str_(t.name, 100), JSON.stringify(t).slice(0, 4000)]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

// ---------- web app ----------
function load_() {
  const ts = tsSheet_(), admin = ME.role === 'admin';
  const res = { user: { username: ME.u, name: ME.name || ME.u, role: ME.role }, masters: masters_(), tpl: tplRows_(), rows: tsRows_(ts) };
  if (admin) res.sql = sqlRows_();                                   // SQL Library is never sent to normal users
  if (admin || ME.shared) res.link = ss_().getUrl() + '#gid=' + ts.getSheetId();   // the file is private unless an admin shared it
  return res;
}
function doGet() {   // never returns any data
  let ready = false;
  try { ready = users_().length > 0; } catch (e) {}
  return out_({ ok: true, service: 'timesheet-multiuser', ready: ready });
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  ME = null; USER_SS = null;
  try {
    const b = JSON.parse(e.postData.contents);
    let res = { ok: true };
    if (b.action === 'login') res = login_(b);
    else {
      ME = auth_(b.token);
      switch (b.action) {
        case 'logout': CacheService.getScriptCache().remove('t_' + b.token); break;
        case 'load': res = load_(); break;
        case 'upsertRow': upsert_(b.row); break;
        case 'deleteRow': del_(b.id); break;
        case 'setTemplates': setTpl_(b.tpl); break;
        case 'changePassword':
          if (!same_(ME.hash, hash_(String(b.old || ''), ME.salt))) throw 'Current password is incorrect.';
          setPw_(ME, b.pass); res = { token: newToken_(ME.u) }; break;
        case 'setMasters': need_(); setMasters_(b.masters); break;
        case 'saveSql': need_(); saveSql_(b.q); break;
        case 'deleteSql': need_(); delSql_(b.id); break;
        case 'listUsers': need_(); res = { users: users_().map(x => ({ username: x.u, name: x.name, role: x.role, active: x.active, shared: x.shared, email: x.email, sheetUrl: 'https://docs.google.com/spreadsheets/d/' + x.sheetId })) }; break;
        case 'addUser': need_(); addUser_(b.username, b.name, b.pass, b.role); break;
        case 'resetPassword': { need_(); const t = users_().find(x => x.u === String(b.username)); if (!t) throw 'User not found.'; setPw_(t, b.pass); break; }
        case 'setActive': {
          need_();
          const t = users_().find(x => x.u === String(b.username));
          if (!t) throw 'User not found.';
          if (t.u === ME.u) throw 'You cannot disable your own account.';
          usersSheet_().getRange(t.row, 6).setValue(b.active ? 'yes' : 'no');
          bump_(t.u); break;
        }
        case 'shareSheet': {   // optional: give a person VIEW access to their own file
          need_();
          const t = users_().find(x => x.u === String(b.username)), mail = String(b.email || '').trim();
          if (!t) throw 'User not found.';
          if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mail)) throw 'Enter a valid Google account email.';
          SpreadsheetApp.openById(t.sheetId).addViewer(mail);
          usersSheet_().getRange(t.row, 8, 1, 2).setValues([[mail, 'yes']]);
          break;
        }
        default: throw 'Unknown action.';
      }
    }
    return out_(res);
  } catch (err) {
    return out_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

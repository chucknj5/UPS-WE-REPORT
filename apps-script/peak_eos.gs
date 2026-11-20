// ===== EOS REPORT (SmartSuite EoS table -> Google Sheets "EOS" tab) =====
// Added to the Peak / New Hires Apps Script. Uses the same Script property SMARTSUITE_TOKEN.
// Routes:  ?action=eos&months=6      pull from SmartSuite (last N months, or months=all), save to the Sheet, return rows
//          ?action=eossheet&months=6 read the saved rows from the Sheet only (fast, no SmartSuite call)
// Optional Script properties: EOS_SHEET_ID (default: the uAttend W.E. Report Sheet), EOS_TABLE.
var EOS_TABLE_ = '6581ee1301a2c2a21aaf079b';
var EOS_SHEET_DEFAULT_ = '1ZBeXa_IoMUi1nqtZLqGj7u9uacQ-5LOOBhT1uSzvWEU';
var EOS_HEAD_ = ['DATE', 'WE', 'SITE #', 'SITE', 'DEPARTMENT', 'LINE', 'SHIFT', 'PLANNED HC', 'START REQ', 'START ACT',
  'FULL REQ', 'FULL ACT', 'PLAN VOL', 'ACT VOL', 'PLAN HRS', 'ACT HRS', 'HANDSHAKE HRS', 'LEADS', 'NEW HIRES', 'RECORD ID', 'PULLED AT'];
var EOS_F_ = {date: 'sf3aca5406', shift: 'sab66e2c02', dept: 's40cfd13c7', siteNum: 's6e4c69c67', pick: 's8e49f69c9',
  planHC: 'sa2c3594bc', startReq: 's06d362c35', startAct: 's1faa34858', fullReq: 'sce0e96ea1', fullAct: 'sa22fe0c65',
  planVol: 's15fa89dd6', actVol: 's7e5a0529c', planHrs: 'sba41cac69', actHrs: 's8100fa62a', hsHrs: 'sdac4ce208',
  leads: 's36e93aecd', nh: 'sdeea9499e'};

function ssFetch_(path, body) {
  var P = PropertiesService.getScriptProperties();
  var tok = P.getProperty('SMARTSUITE_TOKEN');
  if (!tok) throw new Error('Set script property SMARTSUITE_TOKEN');
  var acct = P.getProperty('SMARTSUITE_ACCOUNT') || 'sz7b6s9h';
  var opt = {method: body ? 'post' : 'get', contentType: 'application/json', muteHttpExceptions: true,
    headers: {Authorization: 'Token ' + tok, 'ACCOUNT-ID': acct}};
  if (body) opt.payload = JSON.stringify(body);
  for (var a = 0; a < 4; a++) {
    var res = UrlFetchApp.fetch('https://app.smartsuite.com/api/v1/' + path, opt);
    var c = res.getResponseCode();
    if (c === 200) return JSON.parse(res.getContentText());
    if (c === 429 || c >= 500) { Utilities.sleep(1500 * (a + 1)); continue; }
    var e = new Error('SmartSuite replied ' + c + ': ' + res.getContentText().slice(0, 200)); e.code = c; throw e;
  }
  throw new Error('SmartSuite kept failing (rate limit)');
}
// list every record of a table; with a date filter when dateField+since are given (falls back to filtering here if SmartSuite refuses)
function ssList_(table, dateField, since, keep) {
  var filt = {operator: 'and', fields: []};
  if (dateField && since) filt.fields.push({field: dateField, comparison: 'is_on_or_after', value: {date_mode: 'exact_date', date_mode_value: since}});
  var out = [], offset = 0, total = 1e9, serverFilter = !!(dateField && since);
  while (offset < total) {
    var j;
    try { j = ssFetch_('applications/' + table + '/records/list/?limit=1000&offset=' + offset, {sort: [], filter: filt}); }
    catch (e) {
      if (serverFilter && e.code && e.code < 500 && offset === 0) { serverFilter = false; filt = {operator: 'and', fields: []}; continue; }
      throw e;
    }
    var items = j.items || [];
    items.forEach(function(x) { out.push(keep ? keep(x) : x); });
    total = j.total || out.length;
    if (!items.length) break;
    offset += items.length;
  }
  if (dateField && since && !serverFilter) out = out.filter(function(x) { return dateOf_(x[dateField]) >= since; });
  return out;
}
function dateOf_(v) {
  if (!v) return '';
  if (typeof v === 'string') return v.slice(0, 10);
  var d = v.date || (v.from_date && (v.from_date.date || v.from_date)) || '';
  return String(d).slice(0, 10);
}
function sinceMonths_(months) {
  if (!months || String(months).toLowerCase() === 'all') return '';
  var d = new Date(); d.setMonth(d.getMonth() - Number(months));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function eosChoices_() {
  var t = ssFetch_('applications/' + (PropertiesService.getScriptProperties().getProperty('EOS_TABLE') || EOS_TABLE_) + '/');
  var m = {};
  (t.structure || []).forEach(function(f) {
    if (f.params && Array.isArray(f.params.choices)) { var o = m[f.slug] = {}; f.params.choices.forEach(function(c) { o[c.value] = String(c.label).trim(); }); }
  });
  return m;
}
function eosLine_(dept) {
  var u = String(dept || '').toUpperCase();
  if (/ROADIE/.test(u)) return 'ROADIES';
  if (/SORT/.test(u)) return 'SORT';
  if (/TRANSF/.test(u)) return 'TRANSFER';
  if (/PROC|PRS/.test(u)) return 'PROCESSING';
  if (/TRANSPORT/.test(u)) return 'TRANSPORTATION';
  return u || 'OTHER';
}
function eosShift_(shift, dept) {
  if (shift) return String(shift).toUpperCase();
  var u = String(dept || '').toUpperCase();
  if (/MID/.test(u)) return 'MID';
  if (/NIGHT|PM/.test(u)) return 'PM';
  if (/DAY|AM/.test(u)) return 'AM';
  return '';
}
function weOf_(ds) {
  var d = new Date(Number(ds.slice(0, 4)), Number(ds.slice(5, 7)) - 1, Number(ds.slice(8, 10)), 12);
  d.setDate(d.getDate() + (6 - d.getDay() + 7) % 7);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function eosRow_(x, ch, stamp) {
  var ds = dateOf_(x[EOS_F_.date]); if (!ds) return null;
  var lab = function(f) { var v = x[EOS_F_[f]]; if (Array.isArray(v)) v = v[0]; if (v && typeof v === 'object') v = v.value; return v == null ? '' : ((ch[EOS_F_[f]] || {})[v] || String(v)); };
  var n = function(f) { var v = x[EOS_F_[f]]; return v === '' || v == null || isNaN(Number(v)) ? '' : Number(v); };
  var dept = lab('dept'), num = String(x[EOS_F_.siteNum] || lab('pick') || '').trim();
  var t = String(x.title || ''), k = dept ? t.indexOf(' - ' + dept + ' - ') : -1;
  var site = k > 0 ? t.slice(0, k) : (num ? 'Site ' + num : 'Unassigned');
  if (!num) { var m = site.match(/Site\s+([\d.]+)/); num = m ? m[1] : ''; }
  return [ds, weOf_(ds), num, site, dept, eosLine_(dept), eosShift_(lab('shift'), dept), n('planHC'), n('startReq'), n('startAct'),
    n('fullReq'), n('fullAct'), n('planVol'), n('actVol'), n('planHrs'), n('actHrs'), n('hsHrs'), n('leads'), n('nh'), x.id || '', stamp];
}
function eosSheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('EOS_SHEET_ID') || EOS_SHEET_DEFAULT_;
  var ss = SpreadsheetApp.openById(id), sh = ss.getSheetByName('EOS');
  if (!sh) { sh = ss.insertSheet('EOS'); sh.getRange(1, 1, 1, EOS_HEAD_.length).setValues([EOS_HEAD_]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
// pull from SmartSuite and save: rows on/after the cut-off are replaced, older rows in the Sheet are kept
function eosPull_(months) {
  var since = sinceMonths_(months), ch = eosChoices_(), stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var table = PropertiesService.getScriptProperties().getProperty('EOS_TABLE') || EOS_TABLE_;
  var rows = ssList_(table, EOS_F_.date, since, function(x) { return eosRow_(x, ch, stamp); }).filter(function(r) { return r; });
  rows.sort(function(a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : String(a[3]).localeCompare(String(b[3])); });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = eosSheet_(), last = sh.getLastRow(), keep = [];
    if (since && last > 1) keep = sh.getRange(2, 1, last - 1, EOS_HEAD_.length).getDisplayValues().filter(function(r) { return r[0] && r[0] < since; });
    var all = keep.concat(rows);
    if (last > 1) sh.getRange(2, 1, last - 1, EOS_HEAD_.length).clearContent();
    sh.getRange(1, 1, 1, EOS_HEAD_.length).setValues([EOS_HEAD_]).setFontWeight('bold');
    if (all.length) {
      sh.getRange(2, 1, all.length, 3).setNumberFormat('@');
      sh.getRange(2, 1, all.length, EOS_HEAD_.length).setValues(all);
    }
    sh.getRange(1, EOS_HEAD_.length + 2).setValue('Last pull: ' + stamp + ' · ' + (since ? 'since ' + since : 'ALL history') + ' · ' + rows.length + ' rows (' + all.length + ' in tab)');
  } finally { lock.releaseLock(); }
  return {head: EOS_HEAD_, rows: rows, since: since, pulled: stamp, total: rows.length};
}
function eosRead_(months) {
  var since = sinceMonths_(months), sh = eosSheet_(), last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, EOS_HEAD_.length).getDisplayValues().filter(function(r) { return r[0] && (!since || r[0] >= since); }) : [];
  return {head: EOS_HEAD_, rows: rows, since: since, pulled: String(sh.getRange(1, EOS_HEAD_.length + 2).getDisplayValue() || '').replace(/^Last pull:\s*/, '')};
}
function eosRoute_(p) {
  try { return json_(p.action === 'eossheet' ? eosRead_(p.months || 6) : eosPull_(p.months || 6)); }
  catch (err) { return json_({error: String(err.message || err)}); }
}
// for a time trigger: keeps the last 6 months fresh in the Sheet
function pullEosToSheet() { eosPull_(6); }

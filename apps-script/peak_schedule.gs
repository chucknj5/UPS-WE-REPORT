// ===== SCHEDULE (SmartSuite Schedule table -> Google Sheets "SCHEDULE" tab) =====
// Needs the EOS add-on (ssFetch_) in the same project and Script property SMARTSUITE_TOKEN.
// Routes:  ?action=schedule        pull from SmartSuite (everyone not Archived / DNR), save to the Sheet, return rows
//          ?action=schedulesheet   read the saved rows from the Sheet only (fast)
var SCH_TABLE_ = '661eaa3616bc65334b7db617', SITES_TABLE_ = '6581f43c43df46072a8b4aee';
var SCH_HEAD_ = ['RECORD ID', 'NAME', 'SITE', 'SITE #', 'SHIFT', 'DEPARTMENT', 'POSITION', 'ROLE', 'DAYS OFF', 'WORKED LAST 2 WKS', 'PULLED AT'];
var SCH_F_ = {name: 's6c931bfef', site: 'sfb97e61cd', shift: 'sb7fa31096', dept: 's2ec926238', pos: 'se7150efad', role: 's4759494cd', off: 's663f8bf01', w2: 'se7dac4158'};

function schSites_() {
  var m = {}, offset = 0, total = 1e9;
  while (offset < total) {
    var j = ssFetch_('applications/' + SITES_TABLE_ + '/records/list/?limit=500&offset=' + offset, {sort: [], filter: {operator: 'and', fields: []}});
    (j.items || []).forEach(function(x) { m[x.id] = String(x.title || '').trim(); });
    total = j.total || 0; if (!(j.items || []).length) break; offset += j.items.length;
  }
  return m;
}
function schPull_() {
  var table = PropertiesService.getScriptProperties().getProperty('SCH_TABLE') || SCH_TABLE_;
  var ch = {}, st = ssFetch_('applications/' + table + '/');
  (st.structure || []).forEach(function(f) { if (f.params && Array.isArray(f.params.choices)) { var o = ch[f.slug] = {}; f.params.choices.forEach(function(c) { o[c.value] = String(c.label).trim(); }); } });
  var sites = schSites_(), stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  // skip Archived / DNR on the SmartSuite side when we can
  var skip = Object.keys(ch[SCH_F_.shift] || {}).filter(function(k) { return /^(archived|dnr)$/i.test(ch[SCH_F_.shift][k]); });
  var filt = {operator: 'and', fields: skip.length ? [{field: SCH_F_.shift, comparison: 'is_none_of', value: skip}] : []};
  var lab = function(x, f) { var v = x[SCH_F_[f]]; if (Array.isArray(v)) return v.map(function(i) { return (ch[SCH_F_[f]] || {})[i] || i; }).join(', '); if (v && typeof v === 'object') v = v.value; return v == null ? '' : ((ch[SCH_F_[f]] || {})[v] || String(v)); };
  var rows = [], offset = 0, total = 1e9;
  while (offset < total) {
    var j;
    try { j = ssFetch_('applications/' + table + '/records/list/?limit=200&offset=' + offset, {sort: [], filter: filt}); }
    catch (e) { if (filt.fields.length && offset === 0 && e.code && e.code < 500) { filt = {operator: 'and', fields: []}; continue; } throw e; }
    var items = j.items || [];
    items.forEach(function(x) {
      var shift = lab(x, 'shift'); if (/^(archived|dnr)$/i.test(shift)) return;
      var nm = x[SCH_F_.name]; nm = nm && typeof nm === 'object' ? (nm.sys_root || [nm.first_name, nm.last_name].filter(String).join(' ')) : (nm || '');
      var sid = x[SCH_F_.site]; sid = Array.isArray(sid) ? sid[0] : sid;
      var site = sites[sid] || '', mm = site.match(/Site\s+([\d.]+)/);
      var w2 = x[SCH_F_.w2];
      rows.push([x.id, String(nm || x.title || '').replace(/\s+/g, ' ').trim(), site, mm ? mm[1] : '', shift, lab(x, 'dept'), lab(x, 'pos'), lab(x, 'role'),
        lab(x, 'off'), w2 === true || /^(true|yes)$/i.test(String(w2)) ? 'Yes' : 'No', stamp]);
    });
    total = j.total || rows.length; if (!items.length) break; offset += items.length;
  }
  rows.sort(function(a, b) { return String(a[2]).localeCompare(String(b[2])) || String(a[4]).localeCompare(String(b[4])) || String(a[1]).localeCompare(String(b[1])); });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = schSheet_(), last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, SCH_HEAD_.length).clearContent();
    sh.getRange(1, 1, 1, SCH_HEAD_.length).setValues([SCH_HEAD_]).setFontWeight('bold');
    if (rows.length) { sh.getRange(2, 1, rows.length, SCH_HEAD_.length).setNumberFormat('@'); sh.getRange(2, 1, rows.length, SCH_HEAD_.length).setValues(rows); }
    sh.getRange(1, SCH_HEAD_.length + 2).setValue('Last pull: ' + stamp + ' · ' + rows.length + ' people (Archived / DNR left out)');
  } finally { lock.releaseLock(); }
  return {head: SCH_HEAD_, rows: rows, pulled: stamp};
}
function schSheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('EOS_SHEET_ID') || EOS_SHEET_DEFAULT_;
  var ss = SpreadsheetApp.openById(id), sh = ss.getSheetByName('SCHEDULE');
  if (!sh) { sh = ss.insertSheet('SCHEDULE'); sh.getRange(1, 1, 1, SCH_HEAD_.length).setValues([SCH_HEAD_]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function schRead_() {
  var sh = schSheet_(), last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, SCH_HEAD_.length).getDisplayValues().filter(function(r) { return r[0]; }) : [];
  return {head: SCH_HEAD_, rows: rows, pulled: String(sh.getRange(1, SCH_HEAD_.length + 2).getDisplayValue() || '').replace(/^Last pull:\s*/, '').split(' · ')[0]};
}
function scheduleRoute_(p) {
  try { return json_(p.action === 'schedulesheet' ? schRead_() : schPull_()); }
  catch (err) { return json_({error: String(err.message || err)}); }
}
// for a time trigger: keeps the SCHEDULE tab fresh
function pullScheduleToSheet() { schPull_(); }

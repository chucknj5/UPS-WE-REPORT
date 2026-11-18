/**
 * uAttend → Google Sheets sync for the W.E. Report Builder
 * ------------------------------------------------------------------
 * Paste this whole file into the Google Sheet:  Extensions ▸ Apps Script
 * (replace anything in Code.gs), Save, reload the Sheet, then use the
 * new "uAttend" menu:  1) Setup  2) Test API  3) Pull yesterday
 * 4) Install daily triggers  5) Deploy as Web App (see bottom of file).
 *
 * Data lands in the TIMECARD tab — one row per employee per department per
 * day, same columns as uAttend's Reports ▸ Timecard Report export.
 * Sources:
 *   API   – pulled from the uAttend API (punch report + timecards)
 *   CSV   – Timecard Report CSVs dropped in the Drive folder
 *           "uAttend Timecard Drops" or emailed to this Gmail
 * Re-pulling a day replaces that day (per account), never duplicates.
 */

const TAB = 'TIMECARD', LOGTAB = 'LOG';
const HEAD = ['DATE', 'END DATE', 'ACCOUNT', 'EMPLOYEE NAME', 'PAYROLL ID', 'FIRST NAME', 'LAST NAME', 'DEPARTMENT NAME',
  'REG', 'REG Pay', 'OT1 Hours', 'OT1 Pay', 'OT2 Hours', 'OT2 Pay', 'TOTAL Hours', 'Total Pay', 'SOURCE', 'PULLED AT', 'FLAG'];
const API = 'https://api.workwelltech.com';
const SHEET_ID = '1ZBeXa_IoMUi1nqtZLqGj7u9uacQ-5LOOBhT1uSzvWEU';   // uAttend Timecard Data (W.E. Report)
function ss_() { try { const a = SpreadsheetApp.getActive(); if (a) return a; } catch (e) {} return SpreadsheetApp.openById(SHEET_ID); }
function ui_() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } }
const P = PropertiesService.getScriptProperties();

/* ================= ONE-TIME SETUP (run from the Apps Script editor) =================
 * Put the account name + API key below, click Run ▸ SETUP_ONCE, then blank the key again.
 * The key is stored in Script Properties (not in the Sheet or the web page).
 */
function SETUP_ONCE() {
  const NAME = 'UPS Time Clock';   // exactly as on the top of uAttend's Timecard Report
  const KEY = '';                  // paste the uAttend API key here for the first run only
  ensureSheets_();
  if (KEY) {
    const list = accounts_().filter(a => a.name.toUpperCase() !== NAME.toUpperCase());
    list.push({ name: NAME, key: KEY }); P.setProperty('ACCOUNTS', JSON.stringify(list));
  }
  if (!P.getProperty('WEB_TOKEN')) P.setProperty('WEB_TOKEN', Utilities.getUuid().slice(0, 8).toUpperCase());
  if (!P.getProperty('PULL_BACK_DAYS')) P.setProperty('PULL_BACK_DAYS', '2');
  dropFolder_();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pullYesterday').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('importDrops').timeBased().everyHours(1).create();
  let test = '';
  accounts_().forEach(a => { try { const d = dayStr_(new Date(Date.now() - 864e5));
    const j = call_(a.key, '/reports/punch', { StartDate: d, EndDate: d, UsePaging: true, PageSize: 5, PageNumber: 0 });
    const b = j.Body || j; test += '✓ ' + a.name + ' connected (' + (b.Total || 0) + ' punches ' + d + ')\n'; }
    catch (e) { test += '✗ ' + a.name + ': ' + e.message + '\n'; } });
  console.log(test); log_('SETUP', test);
  console.log(showSettings());
  pullYesterday();
  console.log('Pulled. See the LOG tab in the Sheet.');
}

/* ================= MENU / SETUP ================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('uAttend')
    .addItem('1) Setup (API key / account / web token)', 'setup')
    .addItem('2) Test API connection', 'testApi')
    .addSeparator()
    .addItem('Pull YESTERDAY now', 'pullYesterday')
    .addItem('Pull a date or date range…', 'pullPrompt')
    .addItem('Import Timecard Report CSVs (Drive folder + Gmail)', 'importDrops')
    .addSeparator()
    .addItem('Install daily 8 AM pull + hourly CSV import', 'installTriggers')
    .addItem('Remove triggers', 'removeTriggers')
    .addItem('Show settings', 'showSettings')
    .addToUi();
}

function accounts_() { try { return JSON.parse(P.getProperty('ACCOUNTS') || '[]'); } catch (e) { return []; } }

function setup() {
  const ui = SpreadsheetApp.getUi();
  ensureSheets_();
  const accts = accounts_();
  const n = ui.prompt('uAttend account name',
    'Use the exact name at the top of your uAttend Timecard Report export (e.g. "UPS Time Clock").\n' +
    'Same name = API rows and CSV rows for that account replace each other.\nCurrent: ' +
    (accts.map(a => a.name).join(', ') || '(none)'), ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK || !n.getResponseText().trim()) return;
  const name = n.getResponseText().trim();
  const k = ui.prompt('uAttend API key for "' + name + '"', 'Paste the x-api-key from WorkWell / uAttend support.', ui.ButtonSet.OK_CANCEL);
  if (k.getSelectedButton() !== ui.Button.OK || !k.getResponseText().trim()) return;
  const list = accts.filter(a => a.name.toUpperCase() !== name.toUpperCase());
  list.push({ name: name, key: k.getResponseText().trim() });
  P.setProperty('ACCOUNTS', JSON.stringify(list));
  if (!P.getProperty('WEB_TOKEN')) P.setProperty('WEB_TOKEN', Utilities.getUuid().slice(0, 8).toUpperCase());
  if (!P.getProperty('PULL_BACK_DAYS')) P.setProperty('PULL_BACK_DAYS', '2');
  dropFolder_();
  showSettings();
}

function showSettings() {
  const accts = accounts_();
  const msg = (
    'Accounts: ' + (accts.map(a => a.name + ' (key …' + a.key.slice(-4) + ')').join(', ') || 'none — run Setup') +
    '\n\nWEB TOKEN (paste into the W.E. Report ☁ UATTEND box): ' + (P.getProperty('WEB_TOKEN') || 'run Setup') +
    '\nDaily pull re-pulls the last ' + (P.getProperty('PULL_BACK_DAYS') || '2') + ' day(s) to catch timecard edits.' +
    '\nDrive drop folder: "uAttend Timecard Drops" (id ' + (P.getProperty('DROP_FOLDER_ID') || '—') + ')' +
    '\nTime zone: ' + Session.getScriptTimeZone());
  console.log(msg); const u = ui_(); if (u) u.alert('uAttend settings', msg, u.ButtonSet.OK); return msg;
}

function ensureSheets_() {
  const ss = ss_();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.getSheets()[0].getName() === 'Sheet1' && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(TAB) : ss.insertSheet(TAB);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#351c15').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
  }
  let lg = ss.getSheetByName(LOGTAB);
  if (!lg) { lg = ss.insertSheet(LOGTAB); lg.appendRow(['WHEN', 'WHAT', 'DETAIL']); lg.getRange(1, 1, 1, 3).setFontWeight('bold'); }
  return sh;
}
function log_(what, detail) {
  try { ss_().getSheetByName(LOGTAB).appendRow([new Date(), what, String(detail).slice(0, 4000)]); } catch (e) {}
  console.log(what + ' ' + detail);
}

/* ================= uAttend API ================= */
function call_(key, path, body) {
  for (let attempt = 0; attempt < 5; attempt++) {
    const r = UrlFetchApp.fetch(API + path, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key }, payload: JSON.stringify(body)
    });
    const c = r.getResponseCode();
    if (c === 429 || c >= 500) { Utilities.sleep(1500 * (attempt + 1)); continue; }
    if (c !== 200) throw new Error(path + ' HTTP ' + c + ': ' + r.getContentText().slice(0, 300));
    return JSON.parse(r.getContentText());
  }
  throw new Error(path + ' kept failing (rate limit / server error)');
}
// many timecard calls in parallel, with retry for 429s
function callMany_(key, path, bodies) {
  const out = new Array(bodies.length);
  let todo = bodies.map((b, i) => i);
  for (let attempt = 0; attempt < 5 && todo.length; attempt++) {
    const next = [];
    for (let s = 0; s < todo.length; s += 15) {
      const idx = todo.slice(s, s + 15);
      const res = UrlFetchApp.fetchAll(idx.map(i => ({
        url: API + path, method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': key }, payload: JSON.stringify(bodies[i])
      })));
      res.forEach((r, j) => {
        const c = r.getResponseCode();
        if (c === 200) { try { out[idx[j]] = JSON.parse(r.getContentText()); } catch (e) { out[idx[j]] = null; } }
        else if (c === 429 || c >= 500) next.push(idx[j]);
        else out[idx[j]] = { _error: c + ' ' + r.getContentText().slice(0, 200) };
      });
      Utilities.sleep(250);
    }
    todo = next;
    if (todo.length) Utilities.sleep(2000 * (attempt + 1));
  }
  return out;
}

const hm_ = s => {           // "8:41" → 8.6833 ; "-0:15" → -0.25 ; "8.5" → 8.5
  if (s === null || s === undefined || s === '') return 0;
  s = String(s).trim();
  if (s.indexOf(':') < 0) return parseFloat(s) || 0;
  const neg = s[0] === '-'; const p = s.replace('-', '').split(':');
  const v = (parseInt(p[0], 10) || 0) + (parseInt(p[1], 10) || 0) / 60;
  return neg ? -v : v;
};
const num_ = v => { const n = parseFloat(String(v === undefined || v === null ? '' : v).replace(/[$,]/g, '')); return isNaN(n) ? 0 : n; };
const r2_ = v => Math.round(v * 100) / 100;
const isoDay_ = v => { if (!v) return ''; const s = String(v); const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[0];
  const m2 = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m2 ? m2[3] + '-' + ('0' + m2[1]).slice(-2) + '-' + ('0' + m2[2]).slice(-2) : ''; };

/** Pull one day for one account → rows in HEAD order */
function pullDay_(acct, day) {
  const key = acct.key, stamp = new Date();
  // 1) punch report: who punched that day (and dept / open punches)
  const users = {}, deptName = {}; let page = 0, got = 0, total = 1;
  while (got < total && page < 100) {
    const j = call_(key, '/reports/punch', { StartDate: day, EndDate: day, UsePaging: true, PageSize: 500, PageNumber: page });
    const b = j.Body || j; const items = b.PunchReportLineItems || [];
    total = b.Total || items.length; got += items.length; page++;
    items.forEach(it => {
      if (it.DepartmentCode && it.DepartmentName) deptName[String(it.DepartmentCode).trim().toUpperCase()] = it.DepartmentName;
      if (isoDay_(it.PunchDate || it.InDate) !== day) return;
      const u = users[it.UserId] = users[it.UserId] || { first: it.FirstName || '', last: it.LastName || '', payroll: it.PayrollNumber || '', punches: [] };
      u.punches.push({ dept: it.DepartmentName || it.DepartmentCode || '', tot: hm_(it.Tot), open: !it.OutTime });
    });
    if (!items.length) break;
  }
  const ids = Object.keys(users);
  if (!ids.length) return [];
  // 2) each user's timecard (reg / OT / pay split by department)
  const tcs = callMany_(key, '/timecards', ids.map(id => ({ UserId: Number(id), Date: day })));
  const rows = [];
  ids.forEach((id, n) => {
    const u = users[id], j = tcs[n];
    const anyOpen = u.punches.some(p => p.open);
    const byDept = {};
    let ok = false;
    if (j && !j._error) {
      const b = j.Body || j, tc = b.Timecard || b;
      const first = tc.FirstName || b.FirstName || u.first, last = tc.LastName || b.LastName || u.last;
      u.first = first; u.last = last; u.payroll = tc.PayrollNumber || b.PayrollNumber || u.payroll;
      (tc.DepartmentTransfers || []).forEach(x => { if (x.departmentCode && x.departmentName) deptName[String(x.departmentCode).trim().toUpperCase()] = x.departmentName; });
      (tc.Workdays || []).forEach(w => {
        if (isoDay_(w.Day3 || w.Day2) !== day) return;
        (w.PunchPairs || []).forEach(pp => {
          const code = Number(pp.Paycode);
          if (code === 6 || code === 7) return;               // break / lunch
          const raw = String(pp.DepartmentName || '').trim();   // the API gives the department CODE here (e.g. "05 PRS")
          const d = deptName[raw.toUpperCase()] || raw || (u.punches[0] && u.punches[0].dept) || '';
          const x = byDept[d] = byDept[d] || { reg: 0, ot1: 0, ot2: 0, regPay: 0, totPay: 0 };
          x.reg += hm_(pp.RegularHours); x.ot1 += hm_(pp.Overtime1Hours); x.ot2 += hm_(pp.Overtime2Hours);
          // pay per punch: TotalGrossPay is often "0", so add up the pieces
          const rp = num_(pp.RegularGrossPay), op = num_(pp.Overtime1GrossPay) + num_(pp.Overtime2GrossPay),
            xp = num_(pp.VacationGrossPay) + num_(pp.HolidayGrossPay) + num_(pp.SickGrossPay) + num_(pp.OtherGrossPay);
          x.regPay += rp; x.totPay += Math.max(rp + op + xp, num_(pp.TotalGrossPay));
          x.o1p = (x.o1p || 0) + num_(pp.Overtime1GrossPay); x.o2p = (x.o2p || 0) + num_(pp.Overtime2GrossPay);
          ok = true;
        });
      });
    }
    if (!ok) { // fall back to raw punches (no OT split / pay) so nobody goes missing
      u.punches.forEach(p => { const x = byDept[p.dept] = byDept[p.dept] || { reg: 0, ot1: 0, ot2: 0, regPay: 0, totPay: 0 }; x.reg += p.tot; });
    }
    Object.keys(byDept).forEach(d => {
      const x = byDept[d], tot = x.reg + x.ot1 + x.ot2;
      const otPay = Math.max(0, x.totPay - x.regPay), otH = x.ot1 + x.ot2;
      const flag = [anyOpen ? 'OPEN PUNCH' : '', !ok ? (j && j._error ? 'TIMECARD ERROR ' + j._error : 'PUNCH ONLY - NO PAY') : ''].filter(Boolean).join('; ');
      rows.push([day, day, acct.name, (u.first + ' ' + u.last).trim(), String(u.payroll || ''), u.first, u.last, d,
        r2_(x.reg), r2_(x.regPay), r2_(x.ot1), r2_(x.o1p || (otH ? otPay * x.ot1 / otH : 0)), r2_(x.ot2), r2_(x.o2p || (otH ? otPay * x.ot2 / otH : 0)),
        r2_(tot), r2_(x.totPay), 'API', stamp, flag]);
    });
  });
  return rows;
}

/** Replace all rows for (date,end,account) with new rows */
function writeDay_(day, end, account, rows) {
  const sh = ensureSheets_();
  const last = sh.getLastRow();
  if (last > 1) {
    const keys = sh.getRange(2, 1, last - 1, 3).getDisplayValues();
    const kill = [];
    keys.forEach((k, i) => { if (k[0] === day && k[1] === end && k[2].toUpperCase() === account.toUpperCase()) kill.push(i + 2); });
    for (let i = kill.length - 1; i >= 0;) {   // delete contiguous blocks bottom-up
      let j = i; while (j > 0 && kill[j - 1] === kill[j] - 1) j--;
      sh.deleteRows(kill[j], kill[i] - kill[j] + 1); i = j - 1;
    }
  }
  if (rows.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, 2).setNumberFormat('@');
    sh.getRange(start, 1, rows.length, HEAD.length).setValues(rows);
  }
}

function pullDates_(days, accountName) {
  const accts = accounts_().filter(a => !accountName || a.name.toUpperCase() === String(accountName).toUpperCase());
  if (!accts.length) throw new Error('No uAttend account set up — run uAttend ▸ Setup');
  const done = [];
  accts.forEach(a => days.forEach(d => {
    try { const rows = pullDay_(a, d); writeDay_(d, d, a.name, rows); done.push(a.name + ' ' + d + ': ' + rows.length + ' rows'); log_('API pull', a.name + ' ' + d + ' → ' + rows.length + ' rows'); }
    catch (e) { done.push(a.name + ' ' + d + ': ERROR ' + e.message); log_('API ERROR', a.name + ' ' + d + ' ' + e.message); }
  }));
  return done;
}
const dayStr_ = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
function daysBetween_(a, b) { const out = []; let d = new Date(a + 'T12:00:00'); const e = new Date(b + 'T12:00:00');
  while (d <= e && out.length < 62) { out.push(dayStr_(d)); d = new Date(d.getTime() + 864e5); } return out; }

function pullYesterday() {
  const back = Math.max(1, parseInt(P.getProperty('PULL_BACK_DAYS') || '2', 10));
  const days = []; for (let i = back; i >= 1; i--) days.push(dayStr_(new Date(Date.now() - i * 864e5)));
  const res = pullDates_(days);
  try { ss_().toast(res.join('\n'), 'uAttend pull', 10); } catch (e) {}
}
function pullPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Pull from uAttend', 'Date (2026-11-17) or range (2026-11-09 to 2026-11-15)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().match(/(\d{4}-\d{2}-\d{2})(?:\D+(\d{4}-\d{2}-\d{2}))?/);
  if (!m) return ui.alert('Use yyyy-mm-dd');
  ui.alert(pullDates_(daysBetween_(m[1], m[2] || m[1])).join('\n'));
}
function testApi() {
  const ui = SpreadsheetApp.getUi(), accts = accounts_();
  if (!accts.length) return ui.alert('Run Setup first.');
  const day = dayStr_(new Date(Date.now() - 864e5));
  ui.alert(accts.map(a => {
    try { const j = call_(a.key, '/reports/punch', { StartDate: day, EndDate: day, UsePaging: true, PageSize: 5, PageNumber: 0 });
      const b = j.Body || j; return '✓ ' + a.name + ': connected — ' + (b.Total || (b.PunchReportLineItems || []).length) + ' punches on ' + day; }
    catch (e) { return '✗ ' + a.name + ': ' + e.message; }
  }).join('\n'));
}

/* ================= CSV IMPORT (Reports ▸ Timecard Report) ================= */
function dropFolder_() {
  let id = P.getProperty('DROP_FOLDER_ID'), f = null;
  try { if (id) f = DriveApp.getFolderById(id); } catch (e) { f = null; }
  if (!f) { const it = DriveApp.getFoldersByName('uAttend Timecard Drops'); f = it.hasNext() ? it.next() : DriveApp.createFolder('uAttend Timecard Drops'); P.setProperty('DROP_FOLDER_ID', f.getId()); }
  return f;
}
function parseTimecardCsv_(text, fallbackAccount) {
  const a = Utilities.parseCsv(text.replace(/^﻿/, ''));
  const hi = a.findIndex(r => String(r[0]).trim().toUpperCase() === 'EMPLOYEE NAME');
  if (hi < 0) return null;
  let account = fallbackAccount || '', start = '', end = '';
  for (let i = 0; i < hi; i++) {
    const line = a[i].join(' ');
    const t = line.match(/Timecard Report\s*-\s*(.+)/i); if (t) account = t[1].trim();
    const ds = line.match(/\d{1,2}\/\d{1,2}\/\d{4}/g); if (ds && !start) { start = isoDay_(ds[0]); end = isoDay_(ds[1] || ds[0]); }
  }
  if (!start) return null;
  const H = a[hi].map(h => String(h).trim().toUpperCase());
  const c = n => H.indexOf(n);
  const g = (r, n) => c(n) >= 0 ? r[c(n)] : '';
  const stamp = new Date(), rows = [];
  for (let i = hi + 1; i < a.length; i++) {
    const r = a[i]; const nm = String(r[0] || '').replace(/\s+/g, ' ').trim();
    if (!nm || nm.toUpperCase() === 'TOTAL') continue;
    const reg = num_(g(r, 'REG')), o1 = num_(g(r, 'OT1 HOURS')), o2 = num_(g(r, 'OT2 HOURS'));
    const tot = g(r, 'TOTAL HOURS') !== '' ? num_(g(r, 'TOTAL HOURS')) : reg + o1 + o2;
    rows.push([start, end, account, nm, String(g(r, 'EMPLOYEE PAYROLL ID') || ''), String(g(r, 'FIRST NAME')).trim(), String(g(r, 'LAST NAME')).trim(), String(g(r, 'DEPARTMENT NAME')).trim(),
      reg, num_(g(r, 'REG PAY')), o1, num_(g(r, 'OT1 PAY')), o2, num_(g(r, 'OT2 PAY')), tot, num_(g(r, 'TOTAL PAY')), 'CSV', stamp, '']);
  }
  return { start: start, end: end, account: account, rows: rows };
}
function importDrops() {
  const res = [];
  // Drive folder
  const f = dropFolder_();
  let done = null; const it = f.getFolders(); while (it.hasNext()) { const x = it.next(); if (x.getName() === 'Imported') done = x; }
  if (!done) done = f.createFolder('Imported');
  const files = f.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (!/\.csv$/i.test(file.getName()) && file.getMimeType() !== 'text/csv') continue;
    const p = parseTimecardCsv_(file.getBlob().getDataAsString());
    if (!p) { res.push('skipped ' + file.getName() + ' (not a Timecard Report)'); continue; }
    writeDay_(p.start, p.end, p.account, p.rows); file.moveTo(done);
    res.push(p.account + ' ' + p.start + (p.end !== p.start ? '→' + p.end : '') + ': ' + p.rows.length + ' rows (' + file.getName() + ')');
  }
  // Gmail (uAttend can email reports) — optional, skipped quietly if Gmail isn't authorized
  try {
    const label = GmailApp.getUserLabelByName('uattend-imported') || GmailApp.createLabel('uattend-imported');
    const q = P.getProperty('GMAIL_QUERY') || 'has:attachment filename:csv "Timecard" -label:uattend-imported newer_than:14d';
    GmailApp.search(q, 0, 50).forEach(th => {
      th.getMessages().forEach(m => m.getAttachments().forEach(at => {
        if (!/\.csv$/i.test(at.getName())) return;
        const p = parseTimecardCsv_(at.getDataAsString());
        if (p) { writeDay_(p.start, p.end, p.account, p.rows); res.push('email: ' + p.account + ' ' + p.start + ': ' + p.rows.length + ' rows'); }
      }));
      th.addLabel(label);
    });
  } catch (e) { /* Gmail not authorized / not used */ }
  res.forEach(x => log_('CSV import', x));
  try { ss_().toast(res.join('\n') || 'No new CSVs found', 'uAttend import', 10); } catch (e) {}
  return res;
}

/* ================= TRIGGERS ================= */
function installTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('pullYesterday').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('importDrops').timeBased().everyHours(1).create();
  const u = ui_(); if (u) u.alert('Installed: pull from uAttend every day ~8 AM (' + Session.getScriptTimeZone() + '), and import dropped/emailed Timecard Report CSVs every hour.');
}
function removeTriggers() { ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t)); }

/* ================= WEB APP (feeds the W.E. Report HTML) =================
 * Deploy ▸ New deployment ▸ type "Web app" ▸ Execute as: Me ▸ Who has access: Anyone
 * Copy the /exec URL into the W.E. Report ☁ UATTEND box with the WEB TOKEN.
 * The token keeps the data private: requests without it get nothing.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    if (!P.getProperty('WEB_TOKEN') || p.token !== P.getProperty('WEB_TOKEN')) throw new Error('Bad or missing token');
    const act = p.action || 'data';
    if (act === 'pull') {
      // live pull = yesterday + today only (each day takes ~1-2 min; Google stops a web call at 6 min).
      // Older days are pulled by the 8 AM job and just read from the Sheet.
      const today = dayStr_(new Date()), yday = dayStr_(new Date(Date.now() - 864e5));
      const days = daysBetween_(p.start, p.end || p.start).filter(d => d >= yday && d <= today);
      const done = days.length ? pullDates_(days, p.account || '') : ['(nothing to pull live — older days come from the daily 8 AM pull)'];
      out = { ok: true, pulled: done, data: readRange_(p.start, p.end || p.start, p.account) };
    } else if (act === 'dates') {
      out = { ok: true, dates: listDates_() };
    } else {
      out = { ok: true, data: readRange_(p.start, p.end || p.start, p.account) };
    }
  } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  const json = JSON.stringify(out);
  if (p.callback && /^[\w$.]+$/.test(p.callback))
    return ContentService.createTextOutput(p.callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function listDates_() {
  const sh = ensureSheets_(), last = sh.getLastRow(); if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, 3).getDisplayValues(), m = {};
  v.forEach(r => { const k = r[0] + '|' + r[1] + '|' + r[2]; m[k] = (m[k] || 0) + 1; });
  return Object.keys(m).sort().map(k => { const s = k.split('|'); return { start: s[0], end: s[1], account: s[2], rows: m[k] }; });
}
/** rows whose DATE is within start..end (and END DATE ≤ end) for an account */
function readRange_(start, end, account) {
  const sh = ensureSheets_(), last = sh.getLastRow(); if (last < 2) return { head: HEAD, rows: [] };
  const k = sh.getRange(2, 1, last - 1, 3).getDisplayValues();
  let lo = -1, hi = -1;
  k.forEach((r, i) => { if (r[0] >= start && r[0] <= end && r[1] <= end && (!account || r[2].toUpperCase() === String(account).toUpperCase())) { if (lo < 0) lo = i; hi = i; } });
  if (lo < 0) return { head: HEAD, rows: [] };
  const block = sh.getRange(lo + 2, 1, hi - lo + 1, HEAD.length).getDisplayValues();
  const rows = block.filter(r => r[0] >= start && r[0] <= end && r[1] <= end && (!account || r[2].toUpperCase() === String(account).toUpperCase()))
    .map(r => [r[0], r[1], r[2], r[3], r[4], r[7], num_(r[8]), num_(r[10]), num_(r[12]), num_(r[14]), num_(r[15]), r[16], r[18]]);
  return { head: ['DATE', 'END', 'ACCOUNT', 'NAME', 'PAYROLL', 'DEPT', 'REG', 'OT1', 'OT2', 'TOTAL', 'PAY', 'SOURCE', 'FLAG'], rows: rows };
}

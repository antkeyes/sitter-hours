/** @OnlyCurrentDoc */

/**
 * Sitter Hours: Google Apps Script backend for the clock-in app.
 *
 * How to install (full guide in README.md):
 *   1. Open the Google Sheet -> Extensions -> Apps Script.
 *   2. Replace everything in Code.gs with this file and click Save.
 *   3. Pick "setup" in the function dropdown and click Run (approve the permissions).
 *   4. Deploy -> New deployment -> Web app
 *        Execute as: Me    Who has access: Anyone
 *   5. Back in the spreadsheet: Sitter Hours menu -> Get app links.
 *
 * Keep this file plain ASCII (use \u escapes) so copy/paste can't garble it.
 *
 * Two secret keys control access:
 *   - the sitter link can do everything (clock in/out, edit shifts, payments)
 *   - the parent link can only view, and add/delete the payments the parent logged
 */

const SHIFTS = 'Shifts';
const PAYMENTS = 'Payments';
const SUMMARY = 'Summary';
const RATE_NAME = 'HourlyRate';
const NAME_NAME = 'SitterName';
const DEFAULT_RATE = 25;
const MAX_SHIFT_HOURS = 24;

const SHIFT_HEADERS = ['ID', 'Date', 'Clock In', 'Clock Out', 'Hours', 'Rate', 'Earned', 'Note'];
const PAYMENT_HEADERS = ['ID', 'Date', 'Amount', 'Note', 'Added By'];

const SITTER = 'sitter';
const PARENT = 'parent';
const PARENT_ACTIONS = ['state', 'addPayment', 'deletePayment'];

/* ------------------------------------------------------------------ */
/* One-time setup: run this from the Apps Script editor.               */
/* Safe to run again; it never deletes data.                           */
/* ------------------------------------------------------------------ */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Summary tab (reuse the blank "Sheet1" a new spreadsheet comes with).
  let summary = ss.getSheetByName(SUMMARY);
  if (!summary) {
    const sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() === 0) {
      summary = sheets[0].setName(SUMMARY);
    } else {
      summary = ss.insertSheet(SUMMARY, 0);
    }
  }
  const shifts = ss.getSheetByName(SHIFTS) || ss.insertSheet(SHIFTS);
  const payments = ss.getSheetByName(PAYMENTS) || ss.insertSheet(PAYMENTS);

  // Shifts tab
  shifts.getRange(1, 1, 1, SHIFT_HEADERS.length).setValues([SHIFT_HEADERS])
    .setFontWeight('bold').setBackground('#FDE8E1');
  shifts.setFrozenRows(1);
  shifts.getRange('B:B').setNumberFormat('ddd, mmm d, yyyy');
  shifts.getRange('C:D').setNumberFormat('h:mm am/pm');
  shifts.getRange('E:E').setNumberFormat('0.00');
  shifts.getRange('F:G').setNumberFormat('$#,##0.00');
  shifts.setColumnWidth(2, 150);
  shifts.setColumnWidth(8, 260);
  shifts.hideColumns(1);

  // Payments tab
  payments.getRange(1, 1, 1, PAYMENT_HEADERS.length).setValues([PAYMENT_HEADERS])
    .setFontWeight('bold').setBackground('#E1F5EC');
  payments.setFrozenRows(1);
  payments.getRange('B:B').setNumberFormat('ddd, mmm d, yyyy');
  payments.getRange('C:C').setNumberFormat('$#,##0.00');
  payments.setColumnWidth(2, 150);
  payments.setColumnWidth(4, 260);
  payments.hideColumns(1);

  // Summary tab: only built the first time, so later edits are kept.
  if (!ss.getRangeByName(RATE_NAME)) {
    summary.getRange('A1').setValue('Babysitting Hours').setFontSize(16).setFontWeight('bold');
    const rows = [
      ['Hourly rate', DEFAULT_RATE],
      ["Sitter's first name", 'Katarina'],
      ['', ''],
      ['Hours worked', '=SUM(Shifts!E2:E)'],
      ['Total earned', '=SUM(Shifts!G2:G)'],
      ['Total paid', '=SUM(Payments!C2:C)'],
      ['Still owed', '=B7-B8'],
      ['Unpaid hours (approx.)', '=IF(B3>0, B9/B3, 0)'],
    ];
    summary.getRange(3, 1, rows.length, 2).setValues(rows);
    summary.getRange('B3').setNumberFormat('$#,##0.00');
    summary.getRange('B3:B4').setBackground('#FFF6D6');
    summary.getRange('B6').setNumberFormat('0.00');
    summary.getRange('B7:B9').setNumberFormat('$#,##0.00');
    summary.getRange('B10').setNumberFormat('0.00');
    summary.getRange('A9:B9').setFontWeight('bold');
    summary.setColumnWidth(1, 190);
    summary.setColumnWidth(2, 110);
    ss.setNamedRange(RATE_NAME, summary.getRange('B3'));
    ss.setNamedRange(NAME_NAME, summary.getRange('B4'));
  }

  ensureKeys_();
  ss.setActiveSheet(summary);
  ss.toast('All set! Now deploy it as a web app, then use the Sitter Hours menu \u2192 Get app links.', 'Sitter Hours', 10);
}

/* ------------------------------------------------------------------ */
/* Spreadsheet menu: links for each phone, and turning links off       */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sitter Hours')
    .addItem('Get app links', 'showLinks')
    .addSeparator()
    .addItem('Turn off the parent link (make a new one)', 'resetParentLink')
    .addItem("Turn off the sitter's link (make a new one)", 'resetSitterLink')
    .addToUi();
}

function showLinks() {
  const html = HtmlService.createHtmlOutput(LINKS_HTML).setWidth(560).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sitter Hours: app links');
}

function resetParentLink() { resetKey_('PARENT_KEY', 'parent'); }
function resetSitterLink() { resetKey_('SITTER_KEY', "sitter's"); }

function resetKey_(prop, who) {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('Make a new ' + who + ' link?',
    'The old ' + who + ' link will stop working right away. You will need to send the new one.',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().setProperty(prop, newKey_());
  showLinks();
}

/** Called from the links dialog. */
function linkInfo() {
  ensureKeys_();
  const p = PropertiesService.getScriptProperties();
  let apiUrl = p.getProperty('API_URL') || '';
  if (!apiUrl) {
    try {
      const u = ScriptApp.getService().getUrl() || '';
      if (/\/exec$/.test(u)) apiUrl = u;
    } catch (e) { /* not deployed yet */ }
  }
  return {
    appUrl: p.getProperty('APP_URL') || '',
    apiUrl: apiUrl,
    sitterKey: p.getProperty('SITTER_KEY'),
    parentKey: p.getProperty('PARENT_KEY'),
  };
}

/** Called from the links dialog. */
function saveLinkSettings(appUrl, apiUrl) {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('APP_URL', String(appUrl || '').trim().replace(/#.*$/, ''));
  p.setProperty('API_URL', String(apiUrl || '').trim());
  return linkInfo();
}

function ensureKeys_() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SITTER_KEY')) p.setProperty('SITTER_KEY', newKey_());
  if (!p.getProperty('PARENT_KEY')) p.setProperty('PARENT_KEY', newKey_());
}

function newKey_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}

/* ------------------------------------------------------------------ */
/* Web app endpoints                                                   */
/* ------------------------------------------------------------------ */

function doGet(e) {
  try {
    const role = auth_(e && e.parameter && e.parameter.key);
    return json_({ ok: true, state: state_(role) });
  } catch (err) {
    return json_(fail_(err));
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const role = auth_(req.key);
    if (role === PARENT && PARENT_ACTIONS.indexOf(req.action) < 0) {
      throw new Error("The parent view can't do that.");
    }
    lock.waitLock(15000);
    handle_(req, role);
    SpreadsheetApp.flush();
    return json_({ ok: true, state: state_(role) });
  } catch (err) {
    return json_(fail_(err));
  } finally {
    lock.releaseLock();
  }
}

function auth_(key) {
  const p = PropertiesService.getScriptProperties();
  const sitter = p.getProperty('SITTER_KEY');
  const parent = p.getProperty('PARENT_KEY');
  if (key && sitter && key === sitter) return SITTER;
  if (key && parent && key === parent) return PARENT;
  const err = new Error("This link doesn't work anymore. Ask for a new one.");
  err.code = 'auth';
  throw err;
}

function fail_(err) {
  return { ok: false, error: err.message, code: err.code || '' };
}

function handle_(req, role) {
  switch (req.action) {
    case 'state': return;
    case 'clockIn': return clockIn_(req);
    case 'clockOut': return clockOut_();
    case 'saveShift': return saveShift_(req);
    case 'deleteShift': return deleteRow_(SHIFTS, req.id);
    case 'addPayment': return addPayment_(req, role);
    case 'deletePayment': return deletePayment_(req.id, role);
    default: throw new Error('Unknown action: ' + req.action);
  }
}

/* ------------------------------------------------------------------ */
/* Actions                                                             */
/* ------------------------------------------------------------------ */

function clockIn_(req) {
  if (findOpen_()) throw new Error("You're already clocked in.");
  appendShift_(new Date(), null, req.note);
}

function clockOut_() {
  const open = findOpen_();
  if (!open) throw new Error("You're not clocked in right now.");
  sheet_(SHIFTS).getRange(open.row, 4).setValue(new Date());
}

/** Add a missed shift or fix an existing one. `in`/`out` are epoch milliseconds. */
function saveShift_(req) {
  const start = new Date(Number(req.in));
  const end = req.out ? new Date(Number(req.out)) : null;
  if (isNaN(start)) throw new Error('Please pick a start time.');
  if (end) {
    if (isNaN(end)) throw new Error('Please pick an end time.');
    if (end <= start) throw new Error('The end time has to be after the start time.');
    if ((end - start) / 36e5 > MAX_SHIFT_HOURS) {
      throw new Error('That shift is over ' + MAX_SHIFT_HOURS + ' hours. Double-check the dates.');
    }
  } else {
    const open = findOpen_();
    if (open && open.id !== req.id) throw new Error('You already have a shift in progress.');
  }

  if (req.id) {
    const row = findRow_(SHIFTS, req.id);
    const sh = sheet_(SHIFTS);
    sh.getRange(row, 3, 1, 2).setValues([[start, end || '']]);
    sh.getRange(row, 8).setValue(clean_(req.note));
  } else {
    appendShift_(start, end, req.note);
  }
}

function addPayment_(req, role) {
  const amount = Math.round(Number(req.amount) * 100) / 100;
  if (!(amount > 0)) throw new Error('Please enter an amount.');
  if (amount > 100000) throw new Error('That amount looks too big. Double-check it.');
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const date = req.date ? Utilities.parseDate(String(req.date), tz, 'yyyy-MM-dd') : new Date();
  const sh = sheet_(PAYMENTS);
  sh.getRange(sh.getLastRow() + 1, 1, 1, PAYMENT_HEADERS.length).setValues([[
    newId_(), date, amount, clean_(req.note), role === PARENT ? 'Parent' : 'Sitter',
  ]]);
}

function deletePayment_(id, role) {
  const row = findRow_(PAYMENTS, id);
  if (role === PARENT) {
    const by = String(sheet_(PAYMENTS).getRange(row, 5).getValue());
    if (by !== 'Parent') throw new Error('You can only delete payments you added.');
  }
  sheet_(PAYMENTS).deleteRow(row);
}

function deleteRow_(name, id) {
  sheet_(name).deleteRow(findRow_(name, id));
}

/* ------------------------------------------------------------------ */
/* Reading                                                             */
/* ------------------------------------------------------------------ */

function state_(role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const shifts = readShifts_();
  const payments = readPayments_(tz);

  let hours = 0, earned = 0, paid = 0;
  shifts.forEach(function (s) { hours += s.hours || 0; earned += s.earned || 0; });
  payments.forEach(function (p) { paid += p.amount; });
  const open = shifts.filter(function (s) { return !s.out; })[0] || null;

  shifts.sort(function (a, b) { return b.in - a.in; });
  payments.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  shifts.forEach(function (s) { delete s.row; });

  const nameRange = ss.getRangeByName(NAME_NAME);
  return {
    role: role,
    sitterName: nameRange ? String(nameRange.getValue() || '').trim() : '',
    rate: rate_(),
    open: open ? { id: open.id, in: open.in } : null,
    totals: {
      hours: round_(hours),
      earned: round_(earned),
      paid: round_(paid),
      owed: round_(earned - paid),
    },
    shifts: shifts.slice(0, 100),
    payments: payments.slice(0, 100),
    sheetUrl: ss.getUrl(),
  };
}

function readShifts_() {
  return rows_(SHIFTS, SHIFT_HEADERS.length)
    .filter(function (x) { return x.r[2] instanceof Date; })
    .map(function (x) {
      const r = x.r;
      return {
        id: String(r[0] || 'row' + x.row),
        row: x.row,
        in: r[2].getTime(),
        out: r[3] instanceof Date ? r[3].getTime() : null,
        hours: num_(r[4]),
        rate: num_(r[5]),
        earned: num_(r[6]),
        note: String(r[7] || ''),
      };
    });
}

function readPayments_(tz) {
  return rows_(PAYMENTS, PAYMENT_HEADERS.length)
    .filter(function (x) { return typeof x.r[2] === 'number'; })
    .map(function (x) {
      const r = x.r;
      return {
        id: String(r[0] || 'row' + x.row),
        date: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : '',
        amount: r[2],
        note: String(r[3] || ''),
        by: String(r[4] || ''),
      };
    });
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function appendShift_(start, end, note) {
  const sh = sheet_(SHIFTS);
  const n = sh.getLastRow() + 1;
  sh.getRange(n, 1, 1, SHIFT_HEADERS.length).setValues([[
    newId_(),
    '=IF(C' + n + '="","",INT(C' + n + '))',
    start,
    end || '',
    '=IF(OR(C' + n + '="",D' + n + '=""),"",ROUND((D' + n + '-C' + n + ')*24,2))',
    rate_(),
    '=IF(E' + n + '="","",ROUND(E' + n + '*N(F' + n + '),2))',
    clean_(note),
  ]]);
}

function findOpen_() {
  return readShifts_().filter(function (s) { return !s.out; })[0] || null;
}

function findRow_(name, id) {
  if (!id) throw new Error('Missing id.');
  if (/^row\d+$/.test(id)) return Number(id.slice(3));
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) return i + 2;
    }
  }
  throw new Error("Couldn't find that entry. It may have been deleted.");
}

function rows_(name, width) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, width).getValues().map(function (r, i) {
    return { r: r, row: i + 2 };
  });
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('The "' + name + '" tab is missing. Run setup() in Apps Script.');
  return sh;
}

function rate_() {
  const r = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(RATE_NAME);
  if (!r) throw new Error('Hourly rate not found. Run setup() in Apps Script.');
  return num_(r.getValue());
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 10);
}

/** Keep notes as plain text: a note starting with "=" shouldn't become a formula. */
function clean_(s) {
  s = String(s || '').trim().slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function num_(v) {
  return typeof v === 'number' && isFinite(v) ? v : 0;
}

function round_(v) {
  return Math.round(v * 100) / 100;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* "Get app links" dialog                                              */
/* ------------------------------------------------------------------ */

const LINKS_HTML = `<!doctype html>
<html><head><base target="_top">
<style>
  body { font: 14px/1.45 -apple-system, "Google Sans", Roboto, Arial, sans-serif; color: #2B2522; margin: 0; padding: 4px 6px; }
  h3 { font-size: 14px; margin: 18px 0 6px; }
  label { display: block; font-size: 12px; font-weight: 600; color: #8C817A; margin: 10px 0 4px; }
  input { width: 100%; box-sizing: border-box; font: inherit; padding: 8px 10px; border: 1px solid #DDD3CC; border-radius: 8px; }
  .row { display: flex; gap: 8px; }
  .row input { flex: 1; background: #FBF7F4; font-size: 12px; }
  button { font: inherit; font-weight: 600; border: 0; border-radius: 8px; padding: 8px 14px; cursor: pointer; background: #EFE7E1; }
  button.go { background: #22A86E; color: #fff; }
  .card { border: 1px solid #F1E6DE; border-radius: 12px; padding: 10px 12px 12px; margin-top: 12px; }
  .card.sitter { border-left: 4px solid #F0645A; }
  .card.parent { border-left: 4px solid #22A86E; }
  .muted { color: #8C817A; font-size: 12px; }
  .warn { color: #C2410C; font-size: 13px; margin-top: 10px; }
</style></head>
<body>
  <div class="muted">Fill these in once. They're saved for next time.</div>
  <label>App address (your GitHub Pages URL)</label>
  <input id="appUrl" placeholder="https://your-name.github.io/sitter-hours/">
  <label>Web app URL (Deploy &rarr; Manage deployments, ends in /exec)</label>
  <input id="apiUrl" placeholder="https://script.google.com/macros/s/&hellip;/exec">
  <div style="margin-top:10px"><button class="go" id="save">Save &amp; make links</button></div>
  <div id="warn" class="warn"></div>
  <div id="links" style="display:none">
    <div class="card sitter">
      <h3 style="margin-top:0">&#x1F9E1; Sitter's link</h3>
      <div class="row"><input id="sitterLink" readonly><button data-copy="sitterLink">Copy</button></div>
      <div class="muted">Clock in/out, edit shifts, record payments.</div>
    </div>
    <div class="card parent">
      <h3 style="margin-top:0">&#x1F49A; Parent's link</h3>
      <div class="row"><input id="parentLink" readonly><button data-copy="parentLink">Copy</button></div>
      <div class="muted">Sees hours and what's owed, and can log payments. Can't change shifts.</div>
    </div>
    <p class="muted">Text each person their link. They open it in Safari, then Share &rarr; Add to Home Screen.
      Don't send the sitter's link to the parent.</p>
  </div>
<script>
  const $ = (id) => document.getElementById(id);
  function show(info) {
    $('appUrl').value = info.appUrl || '';
    $('apiUrl').value = info.apiUrl || '';
    const ok = info.appUrl && /\\/exec$/.test(info.apiUrl || '');
    $('warn').textContent = ok ? '' : (info.apiUrl && !/\\/exec$/.test(info.apiUrl)
      ? 'The web app URL should end in /exec.' : 'Fill in both addresses to get the links.');
    $('links').style.display = ok ? 'block' : 'none';
    if (!ok) return;
    const base = info.appUrl + '#api=' + info.apiUrl + '&key=';
    $('sitterLink').value = base + info.sitterKey;
    $('parentLink').value = base + info.parentKey;
  }
  google.script.run.withSuccessHandler(show).linkInfo();
  $('save').onclick = () => {
    $('save').textContent = 'Saving...';
    google.script.run.withSuccessHandler((info) => { $('save').textContent = 'Save & make links'; show(info); })
      .saveLinkSettings($('appUrl').value, $('apiUrl').value);
  };
  document.querySelectorAll('[data-copy]').forEach((b) => b.onclick = () => {
    const input = $(b.dataset.copy);
    input.select();
    document.execCommand('copy');
    b.textContent = 'Copied!';
    setTimeout(() => (b.textContent = 'Copy'), 1500);
  });
</script>
</body></html>`;

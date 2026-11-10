/**
 * Wequiock Presbyterian Church – Website Updater
 * ------------------------------------------------
 * Lets anyone add Announcements, Calendar Events, and an Urgent Notice
 * to the church website just by using a simple form in Google Sheets.
 *
 * The website reads this sheet automatically. No website editing needed.
 */

const SHEET_ID = '1OZSm_6CDh8Zp0DJEqQDOn4BXsgArk-SiPwiLsV8MQXE';
const TZ = 'America/Chicago';

// The PIN people type on the update page. Change it to anything you like,
// then share it only with the people who should post updates.
const ADMIN_PIN = '1888';

const TAB_HELP   = 'START HERE';
const TAB_ANN    = 'Announcements';
const TAB_EVENTS = 'Calendar Events';
const TAB_NOTICE = 'Urgent Notice';

const GREEN = '#2f4a3a';
const GOLD  = '#b08a3e';
const CREAM = '#faf6ee';

const ANN_HEADERS = ['Show on Website?', 'Date Posted', 'Title', 'Message', 'Take Down On (optional)'];
const EV_HEADERS  = ['Show on Website?', 'Date', 'Time', 'Event Name', 'Details (optional)', 'How Often'];
const HOW_OFTEN   = ['One time only', 'Every week'];

/* =========================================================
   MENU
   ========================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⛪ Church Website')
    .addItem('➕ Add an Announcement', 'showAnnouncementForm')
    .addItem('📅 Add a Calendar Event', 'showEventForm')
    .addItem('⚠️ Post or Remove an Urgent Notice', 'showNoticeForm')
    .addSeparator()
    .addItem('🔗 Get the easy update page link', 'showUpdateLink')
    .addItem('❓ How to use this', 'showHelp')
    .addSeparator()
    .addItem('🔧 Set up the sheet (first time only)', 'setupSheet')
    .addToUi();
}

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

/* =========================================================
   ONE-TIME SETUP: builds every tab, column, and format
   ========================================================= */
function setupSheet() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  ss.rename('Wequiock Church Website');

  buildHelpTab_(ss);
  buildAnnouncementsTab_(ss);
  buildEventsTab_(ss);
  buildNoticeTab_(ss);

  // Remove the blank default tab if it is still there and empty
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  // Put tabs in a friendly order
  [TAB_HELP, TAB_ANN, TAB_EVENTS, TAB_NOTICE].forEach((name, i) => {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(TAB_HELP));

  try {
    SpreadsheetApp.getUi().alert('All set!',
      'Your sheet is ready. Use the "⛪ Church Website" menu at the top to add announcements and events.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* ran from the editor, no pop-up needed */ }
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(GREEN).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12)
    .setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 42);
  sh.setFrozenRows(1);
}

function buildAnnouncementsTab_(ss) {
  const sh = getOrCreate_(ss, TAB_ANN);
  const isNew = sh.getLastRow() <= 1;
  styleHeader_(sh, ANN_HEADERS);

  const rows = Math.max(sh.getMaxRows(), 300);
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());

  sh.getRange(2, 1, rows - 1, 1).insertCheckboxes();
  sh.getRange(2, 2, rows - 1, 1).setNumberFormat('mmm d, yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate()
      .setAllowInvalid(false).setHelpText('Type a date, like 11/9/2026').build());
  sh.getRange(2, 5, rows - 1, 1).setNumberFormat('mmm d, yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate()
      .setAllowInvalid(false).setHelpText('Optional. The announcement disappears after this date.').build());
  sh.getRange(2, 1, rows - 1, ANN_HEADERS.length).setFontSize(12).setVerticalAlignment('top');
  sh.getRange(2, 4, rows - 1, 1).setWrap(true);

  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 130); sh.setColumnWidth(3, 260);
  sh.setColumnWidth(4, 480); sh.setColumnWidth(5, 170);

  sh.getRange('A1').setNote('Check the box to show it on the website. Uncheck to hide it (nothing is deleted).');
  sh.getRange('E1').setNote('Optional. Leave blank to keep the announcement up until you uncheck it.');

  if (isNew) {
    const today = new Date();
    sh.getRange(2, 1, 2, 5).setValues([
      [true, today, 'Welcome, visitors!', 'Guests are always welcome. Stay after worship for coffee and fellowship.', ''],
      [true, today, 'Weather cancellations', 'If services are cancelled for snow or bad weather, we will post an Urgent Notice at the top of the website.', '']
    ]);
  }
  addBanding_(sh, ANN_HEADERS.length);
}

function buildEventsTab_(ss) {
  const sh = getOrCreate_(ss, TAB_EVENTS);
  const isNew = sh.getLastRow() <= 1;
  styleHeader_(sh, EV_HEADERS);

  const rows = Math.max(sh.getMaxRows(), 300);
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());

  sh.getRange(2, 1, rows - 1, 1).insertCheckboxes();
  sh.getRange(2, 2, rows - 1, 1).setNumberFormat('ddd, mmm d, yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate()
      .setAllowInvalid(false).setHelpText('Type a date, like 12/24/2026').build());
  sh.getRange(2, 3, rows - 1, 1).setNumberFormat('@'); // keep times as plain text, e.g. 10:00 a.m.
  sh.getRange(2, 6, rows - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(HOW_OFTEN, true)
      .setAllowInvalid(false).build());
  sh.getRange(2, 1, rows - 1, EV_HEADERS.length).setFontSize(12).setVerticalAlignment('top');
  sh.getRange(2, 5, rows - 1, 1).setWrap(true);

  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 170); sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 280); sh.setColumnWidth(5, 380); sh.setColumnWidth(6, 140);

  sh.getRange('C1').setNote('Type it the way you want it shown, like 10:00 a.m. or 6:30 p.m.');
  sh.getRange('F1').setNote('"Every week" repeats on the same day of the week as the Date, starting on that date.');

  if (isNew) {
    sh.getRange(2, 1, 3, 6).setValues([
      [true, new Date(2026, 0, 4), '10:00 a.m.', 'Worship & Sunday School', 'Everyone is welcome.', 'Every week'],
      [true, new Date(2026, 10, 26), '', 'Thanksgiving Day', '', 'One time only'],
      [true, new Date(2026, 11, 24), '', 'Christmas Eve Candlelight Service', 'Carols and scripture by candlelight.', 'One time only']
    ]);
  }
  addBanding_(sh, EV_HEADERS.length);
}

function buildNoticeTab_(ss) {
  const sh = getOrCreate_(ss, TAB_NOTICE);
  const isNew = sh.getLastRow() === 0;

  sh.getRange('A1:B1').merge().setValue('URGENT NOTICE  (shows in a red bar at the top of every page)')
    .setBackground('#7a3b2e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13);
  sh.setRowHeight(1, 40);
  sh.getRange('A2:A4').setValues([['Show it?'], ['Message'], ['Take Down On (optional)']])
    .setFontWeight('bold').setFontSize(12).setBackground(CREAM).setVerticalAlignment('top');
  sh.getRange('B2').insertCheckboxes();
  sh.getRange('B3').setWrap(true).setFontSize(13);
  sh.getRange('B4').setNumberFormat('mmm d, yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 560); sh.setRowHeight(3, 80);
  sh.getRange('A6').setValue('Example: "Church is cancelled this Sunday because of snow. Stay safe!"')
    .setFontColor('#6b665c').setFontStyle('italic');
  if (isNew) sh.getRange('B2').setValue(false);
}

function buildHelpTab_(ss) {
  const sh = getOrCreate_(ss, TAB_HELP);
  sh.clear();
  const lines = [
    ['HOW TO UPDATE THE CHURCH WEBSITE'],
    [''],
    ['The easy way: use the "⛪ Church Website" menu at the very top of this page.'],
    [''],
    ['➕  Add an Announcement  →  type a title and message, click Save.'],
    ['📅  Add a Calendar Event  →  pick a date, type the event, click Save.'],
    ['⚠️  Urgent Notice  →  for snow days or last-minute changes. Shows a red bar on the website.'],
    [''],
    ['Your changes appear on the website within about a minute. Just refresh the page.'],
    [''],
    ['TO HIDE SOMETHING:  go to its tab and uncheck the box in the "Show on Website?" column.'],
    ['TO FIX A TYPO:  go to its tab, click the cell, and type over it.'],
    [''],
    ['Please do not rename the tabs or the column titles. The website looks for them by name.'],
    [''],
    ['Questions? Call Rev. Steve Fewell at 920-655-3733.']
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines).setFontSize(13).setWrap(true);
  sh.getRange('A1').setFontSize(18).setFontWeight('bold').setFontColor(GREEN);
  sh.getRange('A3').setFontWeight('bold');
  sh.getRange('A11:A14').setFontColor('#7a3b2e');
  sh.setColumnWidth(1, 760);
  sh.setHiddenGridlines(true);
}

function addBanding_(sh, cols) {
  sh.getBandings().forEach(b => b.remove());
  sh.getRange(1, 1, sh.getMaxRows(), cols)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(GREEN).setFirstRowColor('#ffffff').setSecondRowColor(CREAM);
}

/* =========================================================
   EASY POP-UP FORMS
   ========================================================= */
function showAnnouncementForm() {
  showForm_('Add an Announcement', 'saveAnnouncement', [
    { id: 'title',   label: 'Title',   type: 'text', required: true, placeholder: 'Example: Thanksgiving Potluck' },
    { id: 'message', label: 'Message', type: 'textarea', required: true, placeholder: 'Type the announcement here' },
    { id: 'posted',  label: 'Date Posted', type: 'date', value: today_() },
    { id: 'until',   label: 'Take it down after (optional)', type: 'date' }
  ]);
}

function showEventForm() {
  showForm_('Add a Calendar Event', 'saveEvent', [
    { id: 'name',  label: 'Event Name', type: 'text', required: true, placeholder: 'Example: Christmas Eve Service' },
    { id: 'date',  label: 'Date', type: 'date', required: true, value: today_() },
    { id: 'time',  label: 'Time (optional)', type: 'text', placeholder: 'Example: 6:00 p.m.' },
    { id: 'details', label: 'Details (optional)', type: 'textarea' },
    { id: 'often', label: 'How Often', type: 'select', options: HOW_OFTEN }
  ]);
}

function showNoticeForm() {
  const sh = ss_().getSheetByName(TAB_NOTICE);
  const current = sh ? String(sh.getRange('B3').getValue() || '') : '';
  const showing = sh ? sh.getRange('B2').getValue() === true : false;
  showForm_('Urgent Notice', 'saveNotice', [
    { id: 'message', label: 'Urgent message (shows in a red bar at the top of the website)', type: 'textarea', value: current,
      placeholder: 'Example: Church is cancelled this Sunday because of snow. Stay safe!' },
    { id: 'until', label: 'Take it down after (optional)', type: 'date' },
    { id: 'show', label: 'What do you want to do?', type: 'select',
      options: ['Show this notice on the website', 'Remove the notice from the website'],
      value: showing ? 'Show this notice on the website' : '' }
  ]);
}

function showUpdateLink() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('Deploy the web app first: Deploy, then New deployment, then Web app.'); return; }
  const link = url;
  const html = HtmlService.createHtmlOutput(
    '<div style="font:16px/1.6 Arial,sans-serif;padding:6px 10px">' +
    '<p>Share this link with anyone who posts updates. They also need the PIN.</p>' +
    '<p><a href="' + link + '" target="_blank" style="word-break:break-all;font-size:15px">' + link + '</a></p>' +
    '<p style="color:#6b665c">Tip: open it on a phone and choose "Add to Home Screen" so it works like an app.</p></div>'
  ).setWidth(480).setHeight(240);
  ui.showModalDialog(html, 'Easy update page');
}

function showHelp() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font:16px/1.6 Arial,sans-serif;padding:6px 10px">' +
    '<p><b>➕ Add an Announcement</b><br>Type a title and a message, then click Save.</p>' +
    '<p><b>📅 Add a Calendar Event</b><br>Pick the date, type the event name, then click Save.</p>' +
    '<p><b>⚠️ Urgent Notice</b><br>For snow days or last-minute changes. It shows a red bar at the top of the website. Use the same menu item to remove it.</p>' +
    '<p><b>To hide something:</b> open its tab and uncheck the "Show on Website?" box.</p>' +
    '<p>Changes show up on the website within about a minute. Refresh the page to see them.</p></div>'
  ).setWidth(460).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'How to use this');
}

function showForm_(title, saveFn, fields) {
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  const inputs = fields.map(f => {
    const req = f.required ? ' required' : '';
    const ph = f.placeholder ? ` placeholder="${esc(f.placeholder)}"` : '';
    let input;
    if (f.type === 'textarea') input = `<textarea id="${f.id}"${req}${ph}>${esc(f.value)}</textarea>`;
    else if (f.type === 'select') input = `<select id="${f.id}">` +
      f.options.map(o => `<option${o === f.value ? ' selected' : ''}>${esc(o)}</option>`).join('') + '</select>';
    else input = `<input id="${f.id}" type="${f.type}" value="${esc(f.value)}"${req}${ph}>`;
    return `<label for="${f.id}">${esc(f.label)}</label>${input}`;
  }).join('');

  const html = `
  <style>
    body{font-family:Arial,sans-serif;font-size:16px;margin:0;padding:6px 10px;color:#2b2a26}
    label{display:block;font-weight:bold;margin:14px 0 6px}
    input,select,textarea{width:100%;box-sizing:border-box;font-size:17px;padding:10px;border:2px solid #cfc6b1;border-radius:6px}
    textarea{height:110px;resize:vertical}
    button{margin-top:20px;width:100%;padding:14px;font-size:18px;font-weight:bold;color:#fff;background:${GREEN};border:none;border-radius:8px;cursor:pointer}
    button:disabled{opacity:.6}
    #msg{margin-top:14px;font-weight:bold}
  </style>
  <form id="f">${inputs}<button id="b" type="submit">Save</button><div id="msg"></div></form>
  <script>
    const ids=${JSON.stringify(fields.map(f => f.id))};
    document.getElementById('f').addEventListener('submit',function(e){
      e.preventDefault();
      const data={}; ids.forEach(id=>data[id]=document.getElementById(id).value);
      const b=document.getElementById('b'), m=document.getElementById('msg');
      b.disabled=true; b.textContent='Saving...';
      google.script.run
        .withSuccessHandler(function(t){ m.style.color='${GREEN}'; m.textContent=t; b.textContent='Saved ✓'; setTimeout(function(){google.script.host.close()},1800); })
        .withFailureHandler(function(err){ m.style.color='#a33'; m.textContent='Something went wrong: '+err.message; b.disabled=false; b.textContent='Save'; })
        ['${saveFn}'](data);
    });
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(480).setHeight(fields.length * 95 + 160), title);
}

/* ---------- save handlers (called by the forms) ---------- */
function saveAnnouncement(d) {
  const sh = ss_().getSheetByName(TAB_ANN);
  if (!sh) throw new Error('Run "Set up the sheet" first.');
  const row = [true, parseDate_(d.posted) || new Date(), clean_(d.title), clean_(d.message), parseDate_(d.until) || ''];
  addRow_(sh, row);
  return 'Announcement saved! It will be on the website in about a minute.';
}

function saveEvent(d) {
  const sh = ss_().getSheetByName(TAB_EVENTS);
  if (!sh) throw new Error('Run "Set up the sheet" first.');
  const date = parseDate_(d.date);
  if (!date) throw new Error('Please pick a date.');
  addRow_(sh, [true, date, clean_(d.time), clean_(d.name), clean_(d.details), d.often || HOW_OFTEN[0]]);
  return 'Event saved! It will be on the calendar in about a minute.';
}

function saveNotice(d) {
  const sh = ss_().getSheetByName(TAB_NOTICE);
  if (!sh) throw new Error('Run "Set up the sheet" first.');
  const show = String(d.show).indexOf('Show') === 0 && clean_(d.message) !== '';
  sh.getRange('B2').setValue(show);
  sh.getRange('B3').setValue(clean_(d.message));
  sh.getRange('B4').setValue(parseDate_(d.until) || '');
  return show ? 'Urgent notice is now showing on the website.' : 'Urgent notice removed from the website.';
}

/** Puts the new entry in the first empty row (the website sorts everything by date). */
function addRow_(sh, row) {
  const col = sh.getName() === TAB_EVENTS ? 4 : 3;
  const last = Math.max(sh.getMaxRows(), 2);
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  let r = vals.findIndex(v => String(v[0]).trim() === '');
  r = r === -1 ? last + 1 : r + 2;
  if (r > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), 50);
    sh.getRange(2, 1, 1, row.length).copyTo(sh.getRange(sh.getMaxRows() - 49, 1, 50, row.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sh.getRange(2, 1, 1, row.length).copyTo(sh.getRange(sh.getMaxRows() - 49, 1, 50, row.length), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  sh.getRange(r, 1, 1, row.length).setValues([row]);
}

/* =========================================================
   WEB APP: the website reads this
   ========================================================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  // Data for the website:  <web app URL>?data=1
  // (The website asks for this behind the scenes.)
  if (p.data || p.callback) {
    const json = JSON.stringify(getWebsiteData_());
    if (p.callback && /^[A-Za-z_$][\w$]*$/.test(p.callback)) {
      return ContentService.createTextOutput(p.callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
  // Everyone else gets the easy update page (just the plain web app link).
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Update the Wequiock Church Website')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebsiteData_() {
  const ss = ss_();
  const todayStr = today_();
  const fmt = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const stillUp = until => !(until instanceof Date) || fmt(until) >= todayStr;

  // Announcements
  const announcements = [];
  const a = ss.getSheetByName(TAB_ANN);
  if (a && a.getLastRow() > 1) {
    a.getRange(2, 1, a.getLastRow() - 1, ANN_HEADERS.length).getValues().forEach(r => {
      const [show, posted, title, message, until] = r;
      if (show !== true || !String(title).trim() || !stillUp(until)) return;
      announcements.push({
        date: posted instanceof Date ? fmt(posted) : '',
        title: String(title).trim(),
        message: String(message).trim()
      });
    });
    announcements.sort((x, y) => (y.date || '').localeCompare(x.date || ''));
  }

  // Events
  const events = [];
  const ev = ss.getSheetByName(TAB_EVENTS);
  if (ev && ev.getLastRow() > 1) {
    ev.getRange(2, 1, ev.getLastRow() - 1, EV_HEADERS.length).getValues().forEach(r => {
      const [show, date, time, name, details, often] = r;
      if (show !== true || !(date instanceof Date) || !String(name).trim()) return;
      events.push({
        date: fmt(date),
        time: time instanceof Date ? Utilities.formatDate(time, TZ, 'h:mm a') : String(time).trim(),
        title: String(name).trim(),
        details: String(details).trim(),
        weekly: often === 'Every week'
      });
    });
  }

  // Urgent notice
  let notice = null;
  const n = ss.getSheetByName(TAB_NOTICE);
  if (n) {
    const show = n.getRange('B2').getValue() === true;
    const msg = String(n.getRange('B3').getValue() || '').trim();
    const until = n.getRange('B4').getValue();
    if (show && msg && stillUp(until)) notice = msg;
  }

  return { announcements, events, notice, updated: new Date().toISOString() };
}

/* =========================================================
   UPDATE PAGE (index.html) – server functions
   ========================================================= */
function checkPin_(pin) {
  if (String(pin || '').trim() !== String(ADMIN_PIN)) throw new Error('That PIN is not right. Please try again.');
}

function webCheckPin(pin) { checkPin_(pin); return true; }

function webSaveAnnouncement(pin, d) { checkPin_(pin); return saveAnnouncement(d); }
function webSaveEvent(pin, d)        { checkPin_(pin); return saveEvent(d); }
function webSaveNotice(pin, d)       { checkPin_(pin); return saveNotice(d); }

/** Everything currently showing on the website, with row numbers so it can be hidden. */
function webList(pin) {
  checkPin_(pin);
  const ss = ss_();
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, TZ, 'MMM d, yyyy') : '';
  const todayStr = today_();
  const out = { announcements: [], events: [], notice: '' };

  const a = ss.getSheetByName(TAB_ANN);
  if (a && a.getLastRow() > 1) {
    a.getRange(2, 1, a.getLastRow() - 1, ANN_HEADERS.length).getValues().forEach((r, i) => {
      if (r[0] === true && String(r[2]).trim()) out.announcements.push({ row: i + 2, title: String(r[2]), date: fmt(r[1]) });
    });
  }
  const ev = ss.getSheetByName(TAB_EVENTS);
  if (ev && ev.getLastRow() > 1) {
    ev.getRange(2, 1, ev.getLastRow() - 1, EV_HEADERS.length).getValues().forEach((r, i) => {
      if (r[0] !== true || !String(r[3]).trim() || !(r[1] instanceof Date)) return;
      const weekly = r[5] === 'Every week';
      if (!weekly && Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') < todayStr) return; // skip past one-time events
      out.events.push({ row: i + 2, title: String(r[3]), date: weekly ? 'Every ' + Utilities.formatDate(r[1], TZ, 'EEEE') : fmt(r[1]), time: String(r[2] || '') });
    });
  }
  const n = ss.getSheetByName(TAB_NOTICE);
  if (n && n.getRange('B2').getValue() === true) out.notice = String(n.getRange('B3').getValue() || '');
  return out;
}

/** Hides an item from the website (unchecks its box; nothing is deleted). */
function webHide(pin, type, row) {
  checkPin_(pin);
  const sh = ss_().getSheetByName(type === 'event' ? TAB_EVENTS : TAB_ANN);
  row = Number(row);
  if (!sh || row < 2 || row > sh.getLastRow()) throw new Error('Could not find that item.');
  sh.getRange(row, 1).setValue(false);
  return 'Removed from the website.';
}

/* ---------- helpers ---------- */
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function clean_(s) { return String(s == null ? '' : s).trim(); }
function parseDate_(s) {
  if (!s) return null;
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}

/** Run this from the editor to check what the website will receive. */
function testWebsiteData() { Logger.log(JSON.stringify(getWebsiteData_(), null, 2)); }

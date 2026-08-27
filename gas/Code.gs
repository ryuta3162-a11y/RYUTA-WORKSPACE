/**
 * RYUTA Workspace — サーバ側
 * スプレッドシートIDは 1 本にまとめる想定（日報ブックに WorkspaceSync シートを追加する形を推奨）
 */
var WS_CONFIG = {
  /** タスク同期を書き込むスプレッドシートID（URLの /d/ と /edit/ の間） */
  SPREADSHEET_ID: '1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q',
  /** 同期用シート名（なければ自動作成） */
  SYNC_SHEET_NAME: 'WorkspaceSync',
  /** 個人 TODO（1行=1タスク）。なければ自動作成 */
  TASKS_SHEET_NAME: 'Tasks',
  TASKS_HEADER: [
    'id',
    'title',
    'bucket',
    'due_date',
    'priority',
    'url',
    'status',
    'done_at',
    'created_at',
    'updated_at',
    'note',
    'period_key',
  ],
  /** 日付列・JSON列・所感（1行目）。既存4列シートは初回保存時に E 列が追記されます */
  HEADER_ROW: ['date', 'active_json', 'done_json', 'updated_at', 'kansou'],
};

/**
 * Webアプリのエントリ
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.api) {
    return handleApiGet_(e);
  }
  var html = HtmlService.createTemplateFromFile('index').evaluate();
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  html.setTitle('RYUTA Workspace');
  return html;
}

function doPost(e) {
  return handleApiPost_(e);
}

/**
 * HTML から CSS/JS を分割している場合に使用（今回は index 単体なら未使用で可）
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

/** Vercel 連携用。スクリプトプロパティ WS_API_TOKEN を設定すると必須 */
function getApiToken_() {
  return PropertiesService.getScriptProperties().getProperty('WS_API_TOKEN') || '';
}

function applyWsApiToken(token) {
  var value = String(token || '').trim();
  if (!value) return { ok: false, message: 'empty' };
  PropertiesService.getScriptProperties().setProperty('WS_API_TOKEN', value);
  return { ok: true };
}

function isApiAuthorized_(e, body) {
  var expected = getApiToken_();
  if (!expected) return true;
  var token = '';
  if (e && e.parameter && e.parameter.token) token = String(e.parameter.token);
  if (body && body.token) token = String(body.token);
  return token === expected;
}

function unauthorized_() {
  return jsonOutput_({ ok: false, message: 'Unauthorized' });
}

/** 今日のカレンダー + Workspace 同期データ（Vercel / AI 用） */
function getDayContextForApi_() {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var tasks = loadWorkspaceTasksFromSheet();
  return {
    ok: true,
    date: today,
    timezone: tz,
    calendarEvents: getTodayCalendarEvents_(),
    workspace: tasks,
  };
}

function getTodayCalendarEvents_() {
  try {
    var start = new Date();
    start.setHours(0, 0, 0, 0);
    var end = new Date();
    end.setHours(23, 59, 59, 999);
    var events = CalendarApp.getDefaultCalendar().getEvents(start, end);
    var out = [];
    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      out.push({
        title: ev.getTitle(),
        start: Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), 'HH:mm'),
        end: Utilities.formatDate(ev.getEndTime(), Session.getScriptTimeZone(), 'HH:mm'),
        isAllDay: ev.isAllDayEvent(),
      });
    }
    return out;
  } catch (err) {
    console.error('Calendar:', err);
    return [];
  }
}

function handleApiGet_(e) {
  try {
    var api = e && e.parameter ? String(e.parameter.api || '') : '';
    if (api === 'status') {
      return jsonOutput_({ ok: true, service: 'ryuta-workspace-gas', version: 'v2-vercel' });
    }
    if (api === 'dayContext') {
      if (!isApiAuthorized_(e, null)) return unauthorized_();
      return jsonOutput_(getDayContextForApi_());
    }
    if (api === 'personalTasks') {
      if (!isApiAuthorized_(e, null)) return unauthorized_();
      return jsonOutput_(listPersonalTasksForApi_());
    }
    if (api === 'partnerMails') {
      if (!isApiAuthorized_(e, null)) return unauthorized_();
      var emails = String((e.parameter && e.parameter.emails) || '');
      var label = String((e.parameter && e.parameter.label) || '');
      return jsonOutput_(searchPartnerMailsForApi_(emails, label));
    }
    if (api === 'vendorMail') {
      return jsonOutput_(syncOneVendorFromGmail_(
        String((e.parameter && e.parameter.company) || ''),
        String((e.parameter && e.parameter.email) || '')
      ));
    }
    if (api === 'vendorDiscover') {
      return jsonOutput_(discoverVendorEmails_(String((e.parameter && e.parameter.company) || '')));
    }
    if (api === 'keidoPreview') {
      var packed = buildKeidoTableHtmlSafe_();
      return jsonOutput_({
        ok: !!(packed && packed.html),
        length: packed && packed.html ? packed.html.length : 0,
        sample: packed && packed.html ? String(packed.html).slice(0, 400) : '',
        err: packed && packed.err ? packed.err : '',
        sheet: packed && packed.sheet ? packed.sheet : '',
        id: packed && packed.id ? packed.id : '',
      });
    }
    if (api === 'dashboard') {
      var tasks = loadWorkspaceTasksFromSheet();
      var unread = getUnreadEmailCount();
      return jsonOutput_({
        ok: true,
        unreadCount: unread,
        tasks: tasks,
      });
    }
    return jsonOutput_({ ok: false, message: 'Unknown GET api: ' + api });
  } catch (e2) {
    return jsonOutput_({ ok: false, message: String(e2.message || e2) });
  }
}

function handleApiPost_(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var api = String((body && body.api) || '');
    var needsAuth =
      api === 'saveTasks' ||
      api === 'createDailyDraft' ||
      api === 'previewDailyReport' ||
      api === 'polishKansou' ||
      api === 'vendorSync' ||
      api === 'personalTasks';
    if (needsAuth && !isApiAuthorized_(e, body)) return unauthorized_();

    if (api === 'saveTasks') {
      return jsonOutput_(
        saveWorkspaceTasksToSheet(
          JSON.stringify(body.active || []),
          JSON.stringify(body.done || []),
          String(body.kansou || '')
        )
      );
    }
    if (api === 'createDailyDraft') {
      return jsonOutput_(
        createDailyReportFromWorkspace(
          JSON.stringify(body.active || []),
          JSON.stringify(body.done || []),
          String(body.kansou || '')
        )
      );
    }
    if (api === 'previewDailyReport') {
      return jsonOutput_(
        previewDailyReport(
          JSON.stringify(body.active || []),
          JSON.stringify(body.done || []),
          String(body.kansou || '')
        )
      );
    }
    if (api === 'polishKansou') {
      return jsonOutput_(polishKansouWithGemini(String(body.text || '')));
    }
    if (api === 'vendorSync') {
      return jsonOutput_(syncVendorCasesFromGmail_());
    }
    if (api === 'personalTasks') {
      return jsonOutput_(handlePersonalTasksPost_(body));
    }
    return jsonOutput_({ ok: false, message: 'Unknown POST api: ' + api });
  } catch (e3) {
    return jsonOutput_({ ok: false, message: String(e3.message || e3) });
  }
}

/**
 * Gmail 受信トレイ未読スレッド数（MAIL オーブのバッジ用）
 */
function getUnreadEmailCount() {
  try {
    return GmailApp.search('is:unread in:inbox').length;
  } catch (err) {
    console.error('Gmail取得エラー:', err);
    return 0;
  }
}

function quoteGmailTerm_(value) {
  var s = String(value || '').trim();
  if (!s) return '';
  if (/[\s\/:]/.test(s)) return '"' + s.replace(/"/g, '') + '"';
  return s;
}

/**
 * 取引先メール／Gmailラベルからスレッドを取得
 */
function searchPartnerMailsForApi_(emailsCsv, label) {
  try {
    var parts = [];
    var labelQ = quoteGmailTerm_(label);
    if (labelQ) parts.push('label:' + labelQ);
    var emailParts = [];
    String(emailsCsv || '')
      .split(/[,\n;]+/)
      .forEach(function (raw) {
        var email = String(raw || '').trim();
        if (!email) return;
        emailParts.push('from:' + email);
        emailParts.push('to:' + email);
      });
    if (emailParts.length) parts.push('(' + emailParts.join(' OR ') + ')');
    if (!parts.length) {
      return { ok: false, message: 'メールアドレスか Gmail ラベルを入れてください' };
    }
    var query = parts.join(' ') + ' newer_than:365d';
    var threads = GmailApp.search(query, 0, 12);
    var items = [];
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var msg = thread.getMessages()[thread.getMessageCount() - 1];
      items.push({
        id: String(thread.getId()),
        subject: msg ? String(msg.getSubject() || '') : '',
        from: msg ? String(msg.getFrom() || '') : '',
        date: msg
          ? Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyy/MM/dd')
          : '',
        url: thread.getPermalink(),
        snippet: msg ? String(msg.getPlainBody() || '').replace(/\s+/g, ' ').slice(0, 80) : '',
      });
    }
    return { ok: true, query: query, items: items };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

var VENDOR_COMPANIES_ = [
  { name: 'SEKAI', keys: ['SEKAI', 'セカイ'], domains: ['sekai.co.jp'], contacts: [
    { name: '志賀', email: 'a-shiga@sekai.co.jp' },
    { name: 'SEKAI team', email: 'team@sekai.co.jp' },
  ] },
  { name: 'Lifefitness', keys: ['Lifefitness', 'Life Fitness', 'ライフフィットネス'], domains: ['lifefitness.com'], contacts: [
    { name: 'Life Fitness CS', email: 'customerservice.jp@lifefitness.com' },
  ] },
  { name: 'LIXIL', keys: ['LIXIL', 'リクシル'], domains: ['lixil.com'], contacts: [
    { name: 'LIXILお問い合わせ', email: 'lxlcc-answer105@lixil.com' },
  ] },
  { name: 'BICS', keys: ['BICS', 'ビックス'] },
  { name: '鳳商事', keys: ['鳳商事'], domains: ['ohtori-s.co.jp'], contacts: [
    { name: '齋藤 翔', email: 'm-saito@ohtori-s.co.jp' },
    { name: '首都圏支店', email: 'shutoken@ohtori-s.co.jp' },
  ] },
  { name: 'アイリスオーヤマ', keys: ['アイリスオーヤマ', 'IRIS'], domains: ['irisohyama.co.jp'] },
  { name: 'KH', keys: ['KH', 'gracene'], domains: ['gracene.com'], contacts: [
    { name: '劉震宇', email: 'info@gracene.com' },
  ] },
  { name: '藤ビル', keys: ['藤ビル'], domains: ['fujibuil.co.jp'], contacts: [
    { name: '小笹', email: 't-ozasa@fujibuil.co.jp' },
  ] },
  { name: 'メトス', keys: ['メトス', 'METOS'], domains: ['metos.co.jp'], contacts: [
    { name: '菅原 睦', email: 's.sugawara@metos.co.jp' },
    { name: '三原 宏次', email: 'k.mihara@metos.co.jp' },
  ] },
  { name: 'プロアバンセ', keys: ['プロアバンセ', 'proavance', '小森'], domains: ['proavance.co.jp'], contacts: [
    { name: '小森', email: 'maintenance@proavance.co.jp' },
    { name: '角田', email: 's.tsunoda@proavance.co.jp' },
  ] },
  { name: 'technogym', keys: ['technogym', 'TechnoGym', 'テクノジム'], domains: ['technogym.com'] },
];

function vendorContactEmails_(vendor) {
  var list = (vendor && vendor.contacts) || [];
  var emails = [];
  for (var i = 0; i < list.length; i++) {
    var em = sanitizeEmail_(list[i].email);
    if (em && emails.indexOf(em) === -1) emails.push(em);
  }
  return emails;
}

function getOrCreateVendorSpreadsheet_() {
  return openWorkspaceSpreadsheet_();
}

function getOrCreateVendorSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var first = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (String(first[0] || '') !== headers[0]) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function extractEmailAddresses_(text) {
  var out = [];
  var re = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  var m;
  var seen = {};
  while ((m = re.exec(String(text || '')))) {
    var email = m[0].toLowerCase();
    if (seen[email]) continue;
    seen[email] = true;
    out.push(email);
  }
  return out;
}

function isOwnEmail_(email, myEmail) {
  email = String(email || '').toLowerCase();
  myEmail = String(myEmail || '').toLowerCase();
  if (!email) return false;
  if (myEmail && email === myEmail) return true;
  return /okamoto-group\.co\.jp$/.test(email);
}

function matchVendorName_(text) {
  var hay = String(text || '');
  for (var i = 0; i < VENDOR_COMPANIES_.length; i++) {
    var keys = VENDOR_COMPANIES_[i].keys;
    for (var k = 0; k < keys.length; k++) {
      if (hay.indexOf(keys[k]) !== -1) return VENDOR_COMPANIES_[i].name;
    }
  }
  return '';
}

function vendorQuery_(keys, domains) {
  var parts = [];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    parts.push(key.indexOf(' ') !== -1 ? '"' + key + '"' : key);
  }
  var q = '(' + parts.join(' OR ') + ')';
  if (domains && domains.length) {
    for (var d = 0; d < domains.length; d++) {
      q += ' OR from:@' + domains[d] + ' OR to:@' + domains[d];
    }
  }
  return q + ' newer_than:5m';
}

function sanitizeEmail_(email) {
  email = String(email || '').toLowerCase().trim();
  if (!/^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/.test(email)) return '';
  return email;
}

function isNippoSubject_(subject) {
  return /日報/.test(String(subject || ''));
}

function threadExternalEmails_(thread, myEmail) {
  var messages = thread.getMessages();
  var out = [];
  var seen = {};
  for (var i = 0; i < messages.length; i++) {
    var text = String(messages[i].getFrom() || '') + ' ' + String(messages[i].getTo() || '') + ' ' + String(messages[i].getCc() || '');
    extractEmailAddresses_(text).forEach(function (email) {
      if (isOwnEmail_(email, myEmail) || seen[email]) return;
      seen[email] = true;
      out.push(email);
    });
  }
  return out;
}

function loadVendorEmailsFromSheet_(company) {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName('業者アドレス');
    if (!sh || sh.getLastRow() < 2) return [];
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    var emails = [];
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0] || '') !== company) continue;
      var em = sanitizeEmail_(values[i][1]);
      if (em && emails.indexOf(em) === -1) emails.push(em);
    }
    return emails;
  } catch (err) {
    return [];
  }
}

function vendorEmailsQuery_(emails, myEmail) {
  var parts = [];
  var n = Math.min(emails.length, 8);
  for (var i = 0; i < n; i++) {
    var email = emails[i];
    if (myEmail) {
      parts.push('(from:' + email + ' to:' + myEmail + ')');
      parts.push('(from:' + myEmail + ' to:' + email + ')');
    } else {
      parts.push('from:' + email);
      parts.push('to:' + email);
    }
  }
  return '(' + parts.join(' OR ') + ') newer_than:5m';
}

function packVendorItem_(row, vendor, query, mode, externals) {
  return {
    company: vendor.name,
    emails: (externals && externals.length ? externals : (row.emails ? row.emails.split(/,\s*/) : [])).join(', '),
    lastDate: row.date || '',
    direction: row.direction || '',
    subject: row.subject || '',
    snippet: row.snippet || '',
    from: row.from || '',
    to: row.to || '',
    permalink: row.permalink || '',
    threadId: row.threadId || '',
    messages: row.messages || [],
    messageCount: row.messageCount || 0,
    contacts: vendor.contacts || [],
    query: query,
    mode: mode,
  };
}

function vendorPinnedQuery_(email, myEmail) {
  email = sanitizeEmail_(email);
  myEmail = sanitizeEmail_(myEmail);
  if (!email) return '';
  if (myEmail) {
    return '(from:' + email + ' to:' + myEmail + ') OR (from:' + myEmail + ' to:' + email + ') newer_than:5m';
  }
  return '(from:' + email + ' OR to:' + email + ') newer_than:5m';
}

function gmailOpenUrl_(thread, company) {
  var id = '';
  try {
    id = String(thread.getId() || '');
  } catch (ignored) {}
  var permalink = '';
  try {
    permalink = String(thread.getPermalink() || '');
  } catch (ignored2) {}
  var uiId = id;
  var matched = permalink.match(/#(?:inbox|all|sent|search\/[^/]+)\/([A-Za-z0-9:_-]+)/);
  if (matched && matched[1]) uiId = matched[1];
  if (company && uiId) {
    return 'https://mail.google.com/mail/u/0/#search/' + encodeURIComponent(company) + '/' + uiId;
  }
  if (uiId) {
    return 'https://mail.google.com/mail/u/0/#all/' + uiId;
  }
  if (id) {
    return 'https://mail.google.com/mail/u/0/?fs=1&tf=cv&search=all&th=' + encodeURIComponent(id);
  }
  return permalink;
}

function cleanMailBody_(raw) {
  var text = String(raw || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  text = text.replace(/[\u200B-\u200D\uFE0E\uFE0F]/g, '');
  text = text.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '');
  text = text.replace(/[●○■□▪▫◆◇★☆☑☐☒⬛⬜⬤◉◎▢]/g, '');
  text = text.replace(/[ \t]+/g, ' ');
  text = text.replace(/[・]/g, '\n');
  var lines = text.split('\n');
  var out = [];
  for (var i = 0; i < lines.length; i++) {
    var t = lines[i].trim();
    if (/^--\s*$/.test(t) || /^_{4,}/.test(t) || /^-{4,}/.test(t) || /^[━─=]{3,}/.test(t)) break;
    if (i > 4 && /^(株式会社|TEL[:：]|Tel[:：]|〒|FAX[:：]|This e-?mail is confidential)/i.test(t)) break;
    if (i > 4 && /(配信停止|このメールに心当たり|Copyright)/.test(t)) break;
    out.push(lines[i]);
  }
  return out.join('\n').trim().slice(0, 2500);
}

function summarizeMessage_(msg, myEmail) {
  var from = String(msg.getFrom() || '');
  var to = String(msg.getTo() || '');
  var cc = String(msg.getCc() || '');
  var subject = String(msg.getSubject() || '');
  var body = cleanMailBody_(msg.getPlainBody());
  var date = msg.getDate() || new Date();
  var fromEmails = extractEmailAddresses_(from);
  var toEmails = extractEmailAddresses_(to + ' ' + cc);
  var counterpart = [];
  fromEmails.concat(toEmails).forEach(function (email) {
    if (!isOwnEmail_(email, myEmail) && counterpart.indexOf(email) === -1) counterpart.push(email);
  });
  var sent = fromEmails.some(function (email) { return isOwnEmail_(email, myEmail); });
  return {
    date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    dateObj: date,
    direction: sent ? '送信' : '受信',
    from: from,
    to: to,
    subject: subject,
    snippet: body,
    emails: counterpart.join(', '),
  };
}

function summarizeThread_(thread, myEmail, company) {
  var messages = thread.getMessages();
  var last = messages[messages.length - 1];
  var row = last ? summarizeMessage_(last, myEmail) : {
    date: '',
    dateObj: new Date(0),
    direction: '',
    from: '',
    to: '',
    subject: '',
    snippet: '',
    emails: '',
  };
  var history = [];
  var start = Math.max(0, messages.length - 4);
  for (var i = start; i < messages.length; i++) {
    history.push(summarizeMessage_(messages[i], myEmail));
  }
  row.messages = history;
  row.messageCount = messages.length;
  row.threadId = String(thread.getId() || '');
  row.permalink = gmailOpenUrl_(thread, company);
  return row;
}

function syncOneVendorFromGmail_(name, pinnedEmail) {
  try {
    var vendor = null;
    for (var i = 0; i < VENDOR_COMPANIES_.length; i++) {
      if (VENDOR_COMPANIES_[i].name === name) {
        vendor = VENDOR_COMPANIES_[i];
        break;
      }
    }
    if (!vendor) return { ok: false, message: '不明な会社です: ' + name };

    var myEmail = '';
    try {
      myEmail = String(Session.getActiveUser().getEmail() || '').toLowerCase();
    } catch (ignored) {}

    pinnedEmail = sanitizeEmail_(pinnedEmail);
    var contactEmails = vendorContactEmails_(vendor);
    var sheetEmails = pinnedEmail ? [] : loadVendorEmailsFromSheet_(vendor.name);
    var mode = 'keyword';
    var query = vendorQuery_(vendor.keys, vendor.domains);
    if (pinnedEmail) {
      mode = 'pin';
      query = vendorPinnedQuery_(pinnedEmail, myEmail);
    } else if (contactEmails.length) {
      mode = 'contact';
      query = vendorEmailsQuery_(contactEmails, myEmail);
    } else if (sheetEmails.length) {
      mode = 'address';
      query = vendorEmailsQuery_(sheetEmails, myEmail);
    }
    var threads = GmailApp.search(query, 0, 8);
    var items = [];
    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var subject0 = '';
      try { subject0 = String(thread.getFirstMessageSubject() || ''); } catch (ignoredSub) {}
      if (isNippoSubject_(subject0)) continue;
      var externals = threadExternalEmails_(thread, myEmail);
      if (!externals.length) continue;
      var row = summarizeThread_(thread, myEmail, vendor.name);
      if (isNippoSubject_(row.subject)) continue;
      items.push(packVendorItem_(row, vendor, query, mode, externals));
    }
    items.sort(function (a, b) { return String(b.lastDate).localeCompare(String(a.lastDate)); });
    return {
      ok: true,
      items: items,
      item: items[0] || null,
      mode: mode,
      contacts: vendor.contacts || [],
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function discoverVendorEmails_(name) {
  try {
    var vendor = null;
    for (var i = 0; i < VENDOR_COMPANIES_.length; i++) {
      if (VENDOR_COMPANIES_[i].name === name) {
        vendor = VENDOR_COMPANIES_[i];
        break;
      }
    }
    if (!vendor) return { ok: false, message: '不明な会社です: ' + name };

    var myEmail = '';
    try {
      myEmail = String(Session.getActiveUser().getEmail() || '').toLowerCase();
    } catch (ignored) {}

    var threads = GmailApp.search(vendorQuery_(vendor.keys, vendor.domains), 0, 15);
    var counts = {};
    var samples = {};
    for (var t = 0; t < threads.length; t++) {
      var subject0 = '';
      try { subject0 = String(threads[t].getFirstMessageSubject() || ''); } catch (ignored2) {}
      if (isNippoSubject_(subject0)) continue;
      var externals = threadExternalEmails_(threads[t], myEmail);
      for (var e = 0; e < externals.length; e++) {
        var email = externals[e];
        counts[email] = (counts[email] || 0) + 1;
        if (!samples[email]) samples[email] = subject0;
      }
    }

    var found = [];
    Object.keys(counts).forEach(function (email) {
      found.push({ email: email, count: counts[email], sample: samples[email] || '' });
    });
    found.sort(function (a, b) { return b.count - a.count; });

    var spreadsheetUrl = '';
    try {
      var ss = openWorkspaceSpreadsheet_();
      var headers = ['company', 'email', 'count', 'sampleSubject', 'updatedAt'];
      var sh = getOrCreateVendorSheet_(ss, '業者アドレス', headers);
      var last = sh.getLastRow();
      if (last > 1) {
        var existing = sh.getRange(2, 1, last - 1, 1).getValues();
        for (var r = existing.length - 1; r >= 0; r--) {
          if (String(existing[r][0] || '') === vendor.name) sh.deleteRow(r + 2);
        }
      }
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      if (found.length) {
        var values = found.map(function (row) {
          return [vendor.name, row.email, row.count, row.sample, now];
        });
        sh.getRange(sh.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
      }
      spreadsheetUrl = ss.getUrl();
    } catch (sheetErr) {
      spreadsheetUrl = '';
    }

    return {
      ok: true,
      company: vendor.name,
      emails: found,
      spreadsheetUrl: spreadsheetUrl,
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function syncVendorCasesFromGmail_() {
  try {
    var myEmail = '';
    try {
      myEmail = String(Session.getActiveUser().getEmail() || '').toLowerCase();
    } catch (ignored) {}
    var latest = {};
    var logs = [];

    for (var i = 0; i < VENDOR_COMPANIES_.length; i++) {
      var vendor = VENDOR_COMPANIES_[i];
      var threads = GmailApp.search(vendorQuery_(vendor.keys, vendor.domains), 0, 10);
      for (var t = 0; t < threads.length; t++) {
        var row = summarizeThread_(threads[t], myEmail, vendor.name);
        row.company = vendor.name;
        logs.push(row);
        if (!latest[vendor.name] || row.dateObj > latest[vendor.name].dateObj) {
          latest[vendor.name] = row;
        }
      }
    }

    var sentThreads = GmailApp.search('in:sent newer_than:365d', 0, 40);
    for (var s = 0; s < sentThreads.length; s++) {
      var sentRow = summarizeThread_(sentThreads[s], myEmail);
      var company = matchVendorName_(
        sentRow.subject + ' ' + sentRow.to + ' ' + sentRow.from + ' ' + sentRow.snippet
      );
      if (!company) continue;
      sentRow.company = company;
      logs.push(sentRow);
      if (!latest[company] || sentRow.dateObj > latest[company].dateObj) {
        latest[company] = sentRow;
      }
    }

    var cases = VENDOR_COMPANIES_.map(function (vendor) {
      var hit = latest[vendor.name];
      return {
        company: vendor.name,
        emails: hit ? hit.emails : '',
        lastDate: hit ? hit.date : '',
        direction: hit ? hit.direction : '',
        subject: hit ? hit.subject : '',
        snippet: hit ? hit.snippet : '',
        from: hit ? hit.from : '',
        permalink: hit ? hit.permalink : '',
      };
    });

    var spreadsheetUrl = '';
    try {
      var ss = getOrCreateVendorSpreadsheet_();
      var caseHeaders = ['company', 'emails', 'lastDate', 'direction', 'subject', 'snippet', 'from', 'permalink', 'updatedAt'];
      var logHeaders = ['date', 'company', 'direction', 'from', 'to', 'subject', 'snippet', 'permalink'];
      var caseSheet = getOrCreateVendorSheet_(ss, 'CASE-LOG', caseHeaders);
      var logSheet = getOrCreateVendorSheet_(ss, 'CASE-LOG Mail', logHeaders);
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      var caseRows = cases.map(function (item) {
        return [item.company, item.emails, item.lastDate, item.direction, item.subject, item.snippet, item.from, item.permalink, now];
      });
      if (caseSheet.getLastRow() > 1) {
        caseSheet.getRange(2, 1, caseSheet.getLastRow() - 1, caseHeaders.length).clearContent();
      }
      if (caseRows.length) caseSheet.getRange(2, 1, caseRows.length, caseHeaders.length).setValues(caseRows);
      logs.sort(function (a, b) { return b.dateObj - a.dateObj; });
      var logRows = logs.slice(0, 200).map(function (row) {
        return [row.date, row.company, row.direction, row.from, row.to, row.subject, row.snippet, row.permalink];
      });
      if (logSheet.getLastRow() > 1) {
        logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logHeaders.length).clearContent();
      }
      if (logRows.length) logSheet.getRange(2, 1, logRows.length, logHeaders.length).setValues(logRows);
      spreadsheetUrl = ss.getUrl();
    } catch (sheetErr) {
      spreadsheetUrl = '';
    }

    return { ok: true, spreadsheetUrl: spreadsheetUrl, cases: cases };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function loadVendorCasesForApi_() {
  try {
    var ss = getOrCreateVendorSpreadsheet_();
    var sh = ss.getSheetByName('CASE-LOG');
    if (!sh || sh.getLastRow() < 2) {
      return { ok: true, spreadsheetId: ss.getId(), spreadsheetUrl: ss.getUrl(), cases: [] };
    }
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
    var cases = values.map(function (r) {
      return {
        company: String(r[0] || ''),
        emails: String(r[1] || ''),
        lastDate: String(r[2] || ''),
        direction: String(r[3] || ''),
        subject: String(r[4] || ''),
        snippet: String(r[5] || ''),
        from: String(r[6] || ''),
        permalink: String(r[7] || ''),
      };
    });
    return { ok: true, spreadsheetUrl: ss.getUrl(), cases: cases };
  } catch (err) {
    return { ok: true, cases: [], message: String(err && err.message ? err.message : err) };
  }
}

/**
 * Workspace のタスクをシートに保存（当日行を upsert）
 * @param {string} activeJson - JSON.stringify 済みの active タスク配列
 * @param {string} doneJson   - JSON.stringify 済みの done タスク配列
 * @param {string} kansouText - 所感（プレーンテキスト）。省略可
 * @return {{ ok: boolean, message?: string, row?: number }}
 */
function saveWorkspaceTasksToSheet(activeJson, doneJson, kansouText) {
  try {
    if (WS_CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      return { ok: false, message: 'SPREADSHEET_ID を Code.gs の WS_CONFIG に設定してください。' };
    }
    if (kansouText === undefined || kansouText === null) kansouText = '';

    var ss = openWorkspaceSpreadsheet_();
    var sheet = getOrCreateSyncSheet_(ss);
    ensureWorkspaceSyncHeader_(sheet);

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var now = new Date();

    var data = sheet.getDataRange().getValues();

    var rowIndex = -1;
    for (var r = 1; r < data.length; r++) {
      var cell = data[r][0];
      if (cell && formatDateCell_(cell) === today) {
        rowIndex = r + 1;
        break;
      }
    }

    var row = [today, activeJson, doneJson, now, kansouText];
    if (rowIndex === -1) {
      sheet.appendRow(row);
      rowIndex = sheet.getLastRow();
    } else {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    }

    return { ok: true, row: rowIndex };
  } catch (e) {
    console.error(e);
    return { ok: false, message: String(e.message || e) };
  }
}

/** 1行目を 5 列構成にそろえる（既存4列のみのとき E1 に kansou を追加） */
function ensureWorkspaceSyncHeader_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.getRange(1, 1, 1, WS_CONFIG.HEADER_ROW.length).setValues([WS_CONFIG.HEADER_ROW]);
    return;
  }
  var h = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (h[0] !== 'date' && h[0] !== WS_CONFIG.HEADER_ROW[0]) {
    sheet.clear();
    sheet.getRange(1, 1, 1, WS_CONFIG.HEADER_ROW.length).setValues([WS_CONFIG.HEADER_ROW]);
    return;
  }
  if (lastCol < WS_CONFIG.HEADER_ROW.length) {
    var remain = WS_CONFIG.HEADER_ROW.length - lastCol;
    // getRange(開始列, 列数) は「書き込む列数」と一致させる必要がある
    sheet.getRange(1, lastCol + 1, 1, remain).setValues([
      WS_CONFIG.HEADER_ROW.slice(lastCol),
    ]);
  }
}

/**
 * 画面ロード時に当日分を復元（localStorage とマージするか上書きするかは index 側で選択）
 */
function loadWorkspaceTasksFromSheet() {
  try {
    if (WS_CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      return { ok: false, message: 'SPREADSHEET_ID 未設定' };
    }
    var ss = openWorkspaceSpreadsheet_();
    var sheet = ss.getSheetByName(WS_CONFIG.SYNC_SHEET_NAME);
    if (!sheet) {
      return { ok: true, active: [], done: [], kansou: '', hasTodayRow: false };
    }

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (formatDateCell_(data[r][0]) === today) {
        var active = parseJsonSafe_(data[r][1]);
        var done = parseJsonSafe_(data[r][2]);
        var kansou = data[r][4] != null ? String(data[r][4]) : '';
        return {
          ok: true,
          active: active,
          done: done,
          kansou: kansou,
          hasTodayRow: true,
        };
      }
    }
    return { ok: true, active: [], done: [], kansou: '', hasTodayRow: false };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function getOrCreateSyncSheet_(ss) {
  var sh = ss.getSheetByName(WS_CONFIG.SYNC_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(WS_CONFIG.SYNC_SHEET_NAME);
  }
  return sh;
}

function formatDateCell_(cell) {
  if (Object.prototype.toString.call(cell) === '[object Date]') {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(cell).trim();
}

function parseJsonSafe_(s) {
  if (s === '' || s == null) return [];
  try {
    var v = typeof s === 'string' ? JSON.parse(s) : s;
    return Array.isArray(v) ? v : [];
  } catch (ignored) {
    return [];
  }
}

var PERSONAL_BUCKETS_ = {
  today: 1,
  week: 1,
  month: 1,
  waiting: 1,
  followup: 1,
};

function handlePersonalTasksPost_(body) {
  var action = String((body && body.action) || 'upsert');
  if (action === 'list') return listPersonalTasksForApi_();
  if (action === 'upsert') return upsertPersonalTask_(body && body.task);
  if (action === 'delete') {
    var id = String((body && body.id) || (body && body.task && body.task.id) || '');
    return deletePersonalTask_(id);
  }
  if (action === 'seedSample') return seedPersonalTasksSample_();
  if (action === 'seedFromMail') return seedPersonalTasksFromMail_();
  if (action === 'previewGyomuFromMail') return previewGyomuFromMail_();
  if (action === 'clearAll') return clearAllPersonalTasks_();
  return { ok: false, message: 'Unknown personalTasks action: ' + action };
}

/**
 * 過去の業務日報メール【業務内容】を集計して返す（確認用）
 */
function previewGyomuFromMail_() {
  try {
    var packed = collectGyomuFromPastReports_();
    return {
      ok: true,
      lines: packed.lines,
      counts: packed.counts,
      scanned: packed.scanned,
      sources: packed.sources,
    };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/**
 * 過去日報メールの業務内容から、簡潔な個人 TODO を作り直す（sample 行は入れ替え）
 */
function seedPersonalTasksFromMail_() {
  try {
    var packed = collectGyomuFromPastReports_();
    var lines = packed.lines || [];
    if (!lines.length) {
      return {
        ok: false,
        message: '過去日報から業務内容を取得できませんでした',
        scanned: packed.scanned,
      };
    }

    var cleared = clearSamplePersonalTasks_();
    var ymd = todayYmd_();
    var weekEnd = weekSundayYmd_(ymd);
    var monthEnd = monthEndYmd_(ymd);
    var created = 0;
    var samples = [];

    // よく出る順：今日 → 今週 → 今月 に簡潔タイトルで配分
    for (var i = 0; i < lines.length && i < 16; i++) {
      var title = shortenGyomuTitle_(lines[i]);
      if (!title) continue;
      var bucket = 'today';
      var due = ymd;
      var priority = 'mid';
      if (i >= 5 && i < 10) {
        bucket = 'week';
        due = weekEnd;
      } else if (i >= 10) {
        bucket = 'month';
        due = monthEnd;
      }
      if (i < 2) priority = 'high';
      if (i >= 13) priority = 'low';
      samples.push({ title: title, bucket: bucket, due_date: due, priority: priority });
    }

    // 定例の短い枠だけ残す（メールに無くても必要なもの）
    samples.push({
      title: '日報作成',
      bucket: 'today',
      due_date: ymd,
      priority: 'high',
      url: 'https://ryuta-workspace.vercel.app/workspace.html',
    });
    samples.push({
      title: 'メール確認',
      bucket: 'today',
      due_date: ymd,
      priority: 'high',
      url: 'https://mail.google.com/mail/u/0/#inbox',
    });

    // 重複タイトル除去（先勝ち）
    var seen = {};
    var uniq = [];
    for (var s = 0; s < samples.length; s++) {
      var key = samples[s].title;
      if (seen[key]) continue;
      seen[key] = 1;
      uniq.push(samples[s]);
    }

    for (var c = 0; c < uniq.length; c++) {
      var raw = uniq[c];
      var t = normalizePersonalTask_({
        id: 'sample_' + Utilities.getUuid().slice(0, 8),
        title: raw.title,
        bucket: raw.bucket,
        due_date: raw.due_date || '',
        priority: raw.priority || 'mid',
        url: raw.url || '',
        status: 'open',
        note: 'sample',
      });
      var res = upsertPersonalTask_(t);
      if (res && res.ok) created++;
    }

    return {
      ok: true,
      created: created,
      cleared: cleared,
      fromMail: lines.slice(0, 20),
      scanned: packed.scanned,
      date: ymd,
    };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

/**
 * 会話ベースの仮サンプル（メール取得できないときのフォールバック）
 */
function seedPersonalTasksSample_() {
  try {
    clearSamplePersonalTasks_();
    var ymd = todayYmd_();
    var weekEnd = weekSundayYmd_(ymd);
    var monthEnd = monthEndYmd_(ymd);
    var samples = [
      { title: '日報作成', bucket: 'today', due_date: ymd, priority: 'high' },
      { title: 'メール確認', bucket: 'today', due_date: ymd, priority: 'high' },
      { title: '清掃業務', bucket: 'today', due_date: ymd, priority: 'mid' },
      { title: 'お客様対応', bucket: 'today', due_date: ymd, priority: 'mid' },
      { title: '事務作業', bucket: 'today', due_date: ymd, priority: 'mid' },
      { title: '入会退会確認', bucket: 'week', due_date: weekEnd, priority: 'high' },
      { title: '稟議進捗確認', bucket: 'week', due_date: weekEnd, priority: 'mid' },
      { title: 'PT予約確認', bucket: 'week', due_date: weekEnd, priority: 'mid' },
      { title: '月次OP確認', bucket: 'month', due_date: monthEnd, priority: 'high' },
      { title: '返答待ち案件', bucket: 'waiting', due_date: weekEnd, priority: 'mid' },
      { title: '継続フォロー', bucket: 'followup', due_date: weekEnd, priority: 'low' },
    ];
    var created = 0;
    for (var i = 0; i < samples.length; i++) {
      var raw = samples[i];
      var t = normalizePersonalTask_({
        id: 'sample_' + Utilities.getUuid().slice(0, 8),
        title: raw.title,
        bucket: raw.bucket,
        due_date: raw.due_date || '',
        priority: raw.priority || 'mid',
        status: 'open',
        note: 'sample',
      });
      var res = upsertPersonalTask_(t);
      if (res && res.ok) created++;
    }
    return { ok: true, created: created, date: ymd };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function clearSamplePersonalTasks_() {
  var ss = openWorkspaceSpreadsheet_();
  var sheet = getOrCreateTasksSheet_(ss);
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var headerLen = WS_CONFIG.TASKS_HEADER.length;
  var values = sheet.getRange(2, 1, last - 1, headerLen).getValues();
  var removed = 0;
  for (var i = values.length - 1; i >= 0; i--) {
    var id = String(values[i][0] || '');
    var note = String(values[i][10] || '');
    if (id.indexOf('sample_') === 0 || note === 'sample') {
      sheet.deleteRow(i + 2);
      removed++;
    }
  }
  return removed;
}

/** Tasks シートの全タスクを削除（ヘッダは残す） */
function clearAllPersonalTasks_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sheet = getOrCreateTasksSheet_(ss);
    var last = sheet.getLastRow();
    var cleared = 0;
    if (last >= 2) {
      cleared = last - 1;
      sheet.deleteRows(2, cleared);
    }
    return { ok: true, cleared: cleared };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function shortenGyomuTitle_(line) {
  var s = String(line || '')
    .replace(/^[\s・\-–—●○◆■□→]+/, '')
    .replace(/\s+/g, '')
    .trim();
  if (!s) return '';
  // 日付だけ・ノイズは除外
  if (/^\d{1,2}日$/.test(s)) return '';
  if (/^\d+分\d*本/.test(s)) return '';
  if (/^(計|合計)/.test(s)) return '';
  if (/日報作成|二重カウント|sample/i.test(s)) return '';
  // 長すぎる説明は先頭だけ（業務内容として短く）
  if (s.length > 16) s = s.slice(0, 16);
  return s;
}

function collectGyomuFromPastReports_() {
  var counts = {};
  var sources = [];
  var scanned = 0;
  var subjectPrefix =
    'EAST運営本部 関東運営ブロック' + REPORT_CONFIG.MY_TEAM + '業務日報　' + REPORT_CONFIG.MY_NAME;
  var legacyPrefix = 'EAST運営本部 関東運営ブロック 第7エリア T2　' + REPORT_CONFIG.MY_NAME;
  var queries = [
    'subject:"' + subjectPrefix + '" newer_than:90d',
    'subject:"' + legacyPrefix + '" newer_than:90d',
    'subject:"業務日報" subject:"' + REPORT_CONFIG.MY_NAME + '" newer_than:90d',
  ];

  for (var q = 0; q < queries.length; q++) {
    var threads = GmailApp.search(queries[q], 0, 40);
    for (var t = 0; t < threads.length; t++) {
      var messages = threads[t].getMessages();
      for (var m = 0; m < messages.length; m++) {
        scanned++;
        var body = messages[m].getPlainBody() || '';
        var lines = extractGyomuLinesFromBody_(body);
        if (!lines.length) continue;
        sources.push({
          date: Utilities.formatDate(messages[m].getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          subject: messages[m].getSubject(),
          count: lines.length,
        });
        for (var i = 0; i < lines.length; i++) {
          var key = shortenGyomuTitle_(lines[i]);
          if (!key) continue;
          counts[key] = (counts[key] || 0) + 1;
        }
      }
    }
  }

  // WorkspaceSync の done_json も補助
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sync = ss.getSheetByName(WS_CONFIG.SYNC_SHEET_NAME);
    if (sync && sync.getLastRow() >= 2) {
      var data = sync.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        var done = parseJsonSafe_(data[r][2]);
        for (var d = 0; d < done.length; d++) {
          var item = done[d];
          var title = typeof item === 'string' ? item : item && (item.title || item.text);
          var short = shortenGyomuTitle_(title);
          if (!short) continue;
          counts[short] = (counts[short] || 0) + 1;
        }
      }
      sources.push({ date: '', subject: 'WorkspaceSync', count: Object.keys(counts).length });
    }
  } catch (ignored) {}

  var lines = Object.keys(counts).sort(function (a, b) {
    return counts[b] - counts[a] || a.localeCompare(b, 'ja');
  });

  return { lines: lines, counts: counts, scanned: scanned, sources: sources.slice(0, 15) };
}

function extractGyomuLinesFromBody_(body) {
  var text = String(body || '').replace(/\r\n/g, '\n');
  var start = text.search(/【業務内容】/);
  if (start < 0) start = text.search(/業務内容/);
  if (start < 0) return [];
  var slice = text.slice(start);
  var endMatch = slice.search(/\n【(?:所感|PT実績|本日の所感)/);
  if (endMatch > 0) slice = slice.slice(0, endMatch);
  var rawLines = slice.split('\n');
  var out = [];
  for (var i = 0; i < rawLines.length; i++) {
    var line = String(rawLines[i] || '').trim();
    if (!line) continue;
    if (/【業務内容】|業務内容/.test(line) && line.length < 12) continue;
    if (/^【/.test(line)) break;
    line = line.replace(/^[\s・\-–—●○◆■□→]+/, '').trim();
    if (!line || line.length < 2) continue;
    if (/お元気様|ご確認|経堂数値|PT実績|計\s*\d+分/.test(line)) continue;
    if (/^\d{1,2}日$/.test(line)) continue;
    out.push(line);
  }
  return out;
}

function listPersonalTasksForApi_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sheet = getOrCreateTasksSheet_(ss);
    var headerLen = WS_CONFIG.TASKS_HEADER.length;
    var last = sheet.getLastRow();
    if (last < 2) return { ok: true, tasks: [], carried: 0, date: todayYmd_() };

    var values = sheet.getRange(2, 1, last - 1, headerLen).getValues();
    var ymd = todayYmd_();
    var now = nowStamp_();
    var carried = 0;
    var tasks = [];
    var changed = false;
    for (var i = 0; i < values.length; i++) {
      var t = personalTaskFromRow_(values[i]);
      if (!t.id) continue;
      if (carryPersonalTask_(t, ymd, now)) {
        values[i] = personalTaskToRow_(t);
        carried++;
        changed = true;
      }
      tasks.push(t);
    }
    if (changed) {
      sheet.getRange(2, 1, values.length, headerLen).setValues(values);
    }
    return { ok: true, tasks: tasks, carried: carried, date: ymd };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function upsertPersonalTask_(raw) {
  try {
    var t = normalizePersonalTask_(raw || {});
    if (!t.title) return { ok: false, message: 'title required' };
    var ss = openWorkspaceSpreadsheet_();
    var sheet = getOrCreateTasksSheet_(ss);
    var headerLen = WS_CONFIG.TASKS_HEADER.length;
    var rowIndex = findPersonalTaskRow_(sheet, t.id);
    var row = personalTaskToRow_(t);
    if (rowIndex === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(rowIndex, 1, 1, headerLen).setValues([row]);
    }
    return { ok: true, task: t };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function deletePersonalTask_(id) {
  try {
    id = String(id || '').trim();
    if (!id) return { ok: false, message: 'id required' };
    var ss = openWorkspaceSpreadsheet_();
    var sheet = getOrCreateTasksSheet_(ss);
    var rowIndex = findPersonalTaskRow_(sheet, id);
    if (rowIndex !== -1) sheet.deleteRow(rowIndex);
    return { ok: true, id: id };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  }
}

function findPersonalTaskRow_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return i + 2;
  }
  return -1;
}

function getOrCreateTasksSheet_(ss) {
  var sh = ss.getSheetByName(WS_CONFIG.TASKS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(WS_CONFIG.TASKS_SHEET_NAME);
    sh.getRange(1, 1, 1, WS_CONFIG.TASKS_HEADER.length).setValues([WS_CONFIG.TASKS_HEADER]);
    sh.setFrozenRows(1);
    return sh;
  }
  ensureTasksHeader_(sh);
  return sh;
}

function ensureTasksHeader_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var h = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (h[0] !== 'id') {
    if (sheet.getLastRow() > 0) {
      sheet.insertRowBefore(1);
    }
    sheet.getRange(1, 1, 1, WS_CONFIG.TASKS_HEADER.length).setValues([WS_CONFIG.TASKS_HEADER]);
    sheet.setFrozenRows(1);
    return;
  }
  if (lastCol < WS_CONFIG.TASKS_HEADER.length) {
    var remain = WS_CONFIG.TASKS_HEADER.length - lastCol;
    sheet.getRange(1, lastCol + 1, 1, remain).setValues([WS_CONFIG.TASKS_HEADER.slice(lastCol)]);
  }
}

function personalTaskFromRow_(row) {
  return {
    id: String(row[0] || ''),
    title: String(row[1] || ''),
    bucket: normalizePersonalBucket_(row[2]),
    due_date: formatDateCell_(row[3]),
    priority: normalizePersonalPriority_(row[4]),
    url: String(row[5] || ''),
    status: String(row[6] || 'open') === 'done' ? 'done' : 'open',
    done_at: formatDateCell_(row[7]),
    created_at: String(row[8] || ''),
    updated_at: String(row[9] || ''),
    note: String(row[10] || ''),
    period_key: String(row[11] || ''),
  };
}

function personalTaskToRow_(t) {
  return [
    t.id,
    t.title,
    t.bucket,
    t.due_date || '',
    t.priority,
    t.url || '',
    t.status,
    t.done_at || '',
    t.created_at || '',
    t.updated_at || '',
    t.note || '',
    t.period_key || '',
  ];
}

function normalizePersonalTask_(raw) {
  var ymd = todayYmd_();
  var now = nowStamp_();
  var t = {
    id: String(raw.id || '').trim() || Utilities.getUuid(),
    title: String(raw.title || '').trim(),
    bucket: normalizePersonalBucket_(raw.bucket),
    due_date: String(raw.due_date || '').trim(),
    priority: normalizePersonalPriority_(raw.priority),
    url: String(raw.url || '').trim(),
    status: String(raw.status || 'open') === 'done' ? 'done' : 'open',
    done_at: String(raw.done_at || '').trim(),
    created_at: String(raw.created_at || now),
    updated_at: now,
    note: String(raw.note || '').trim(),
    period_key: String(raw.period_key || ''),
  };
  if (t.status === 'done' && !t.done_at) t.done_at = ymd;
  if (t.status !== 'done') t.done_at = t.done_at || '';
  if (!t.period_key) t.period_key = personalPeriodKey_(t.bucket, ymd);
  return t;
}

function normalizePersonalBucket_(v) {
  var s = String(v || 'today').trim();
  return PERSONAL_BUCKETS_[s] ? s : 'today';
}

function normalizePersonalPriority_(v) {
  var s = String(v || 'mid').trim();
  if (s === 'high' || s === 'low' || s === 'mid') return s;
  return 'mid';
}

function carryPersonalTask_(t, ymd, now) {
  if (!t || t.status === 'done') return false;
  if (t.bucket !== 'today' && t.bucket !== 'week' && t.bucket !== 'month') return false;
  var nextKey = personalPeriodKey_(t.bucket, ymd);
  var due = String(t.due_date || '');
  var dueStale = due && due < ymd;
  var keyStale = t.period_key !== nextKey;
  if (!keyStale && !dueStale) return false;
  t.period_key = nextKey;
  t.due_date = personalCarryDue_(t.bucket, due, ymd);
  t.updated_at = now;
  return true;
}

function personalPeriodKey_(bucket, ymd) {
  if (bucket === 'today') return ymd;
  if (bucket === 'week') return 'W' + weekMondayYmd_(ymd);
  if (bucket === 'month') return String(ymd).slice(0, 7);
  return '';
}

function personalCarryDue_(bucket, due, ymd) {
  if (due && due >= ymd) return due;
  if (bucket === 'today') return ymd;
  if (bucket === 'week') return weekSundayYmd_(ymd);
  if (bucket === 'month') return monthEndYmd_(ymd);
  return due || '';
}

function todayYmd_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function nowStamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function ymdToDate_(ymd) {
  var parts = String(ymd || '').split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function dateToYmd_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function weekMondayYmd_(ymd) {
  var d = ymdToDate_(ymd);
  var day = d.getDay();
  var diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  return dateToYmd_(d);
}

function weekSundayYmd_(ymd) {
  var d = ymdToDate_(weekMondayYmd_(ymd));
  d.setDate(d.getDate() + 6);
  return dateToYmd_(d);
}

function monthEndYmd_(ymd) {
  var parts = String(ymd || '').split('-');
  var d = new Date(Number(parts[0]), Number(parts[1]), 0);
  return dateToYmd_(d);
}

/**
 * 日報メール（V10 相当）— 宛先・署名はここで編集
 */
var REPORT_CONFIG = {
  RECIPIENT:
    'y_east_staff@okamoto-group.co.jp, yamauchieastnippou_transfer@okamoto-group.co.jp, jf-kyoudou@okamoto-group.co.jp',
  CC: 'k-takakuwa@okamoto-group.co.jp, m-akiyama@okamoto-group.co.jp, t-doi@okamoto-group.co.jp',
  MY_TEAM: '7-2',
  MY_NAME: '日下竜太',
  SPREADSHEET_ID: '14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w',
  REPORT_SHEET_NAME: '日報',
  START_ROW: 8,
  START_COL: 2,
  NUM_ROWS: 33,
  NUM_COLS: 8,
};

// 所感と一緒に共有する空きスケジュールリンク
var AVAILABILITY_URL = 'https://calendar.app.google/egS9bVik7BJTZ8Jb9';

/**
 * 画面から渡した業務・所感でプレビュー（Gmail 下書きは作らない）
 * @param {string} activeJson
 * @param {string} doneJson
 * @param {string} kansouText
 */
function previewDailyReport(activeJson, doneJson, kansouText) {
  try {
    var doneTasks = parseDoneTasksInput_(doneJson);
    var pkg = buildDailyReportPackage_(doneTasks, String(kansouText || ''));
    if (!pkg.ok) return pkg;
    return {
      ok: true,
      subject: pkg.subject,
      previewText: pkg.previewText,
    };
  } catch (e) {
    console.error(e);
    return { ok: false, message: String(e.message || e) };
  }
}

/**
 * シート保存 + Gmail 下書きを 1 回で実行（index の「下書き作成」用）
 */
function createDailyReportFromWorkspace(activeJson, doneJson, kansouText) {
  try {
    var active = parseJsonSafe_(activeJson);
    var doneTasks = parseDoneTasksInput_(doneJson);
    var kansou = String(kansouText || '');
    try {
      saveWorkspaceTasksToSheet(
        JSON.stringify(active),
        JSON.stringify(doneTasks),
        kansou
      );
    } catch (ignoredSave) {}
    return createDailyReportDraftWithData_(doneTasks, kansou);
  } catch (e) {
    console.error(e);
    return { ok: false, message: String(e.message || e) };
  }
}

/**
 * Workspace の「Done」を【業務内容】に反映し、Gmail 下書きを作成する（シート当日行から読む）
 * @return {{ ok: boolean, subject?: string, message?: string }}
 */
function createDailyReportDraft() {
  try {
    if (WS_CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      return { ok: false, message: 'SPREADSHEET_ID を設定してください。' };
    }
    var ss = openWorkspaceSpreadsheet_();
    return createDailyReportDraftWithData_(getTodayDoneTasksFromSync_(ss), getTodayKansouFromSync_(ss));
  } catch (e) {
    console.error(e);
    return { ok: false, message: String(e.message || e) };
  }
}

function createDailyReportDraftWithData_(doneTasks, kansouRaw) {
  var pkg = buildDailyReportPackage_(doneTasks, kansouRaw);
  if (!pkg.ok) return pkg;
  GmailApp.createDraft(REPORT_CONFIG.RECIPIENT, pkg.subject, '', {
    htmlBody: pkg.htmlBody,
    cc: REPORT_CONFIG.CC,
  });
  return { ok: true, subject: pkg.subject };
}

function parseDoneTasksInput_(input) {
  if (Array.isArray(input)) return input;
  return parseJsonSafe_(input);
}

function colToA1_(n) {
  var s = '';
  var num = Number(n);
  while (num > 0) {
    var m = (num - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

function reportRangeA1_() {
  var r1 = REPORT_CONFIG.START_ROW;
  var c1 = REPORT_CONFIG.START_COL;
  var r2 = r1 + REPORT_CONFIG.NUM_ROWS - 1;
  var c2 = c1 + REPORT_CONFIG.NUM_COLS - 1;
  return REPORT_CONFIG.REPORT_SHEET_NAME + '!' + colToA1_(c1) + r1 + ':' + colToA1_(c2) + r2;
}

/** SpreadsheetApp.openById が匿名Webアプリで拒否される場合の代替 */
function sheetsApiGetJson_(url) {
  var token = ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
  });
  var code = res.getResponseCode();
  var text = res.getContentText() || '';
  if (code < 200 || code >= 300) {
    throw new Error('Sheets API ' + code + ' ' + String(text).slice(0, 180));
  }
  return JSON.parse(text);
}

function parseSimpleCsv_(text) {
  var lines = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  var rows = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!line) continue;
    var cells = [];
    var cur = '';
    var inQ = false;
    for (var k = 0; k < line.length; k++) {
      var ch = line.charAt(k);
      if (inQ) {
        if (ch === '"') {
          if (line.charAt(k + 1) === '"') {
            cur += '"';
            k++;
          } else inQ = false;
        } else cur += ch;
      } else if (ch === '"') inQ = true;
      else if (ch === ',') {
        cells.push(cur);
        cur = '';
      } else cur += ch;
    }
    cells.push(cur);
    rows.push(cells);
  }
  return rows;
}

function rowsToKeidoTableHtml_(rows) {
  if (!rows || !rows.length) return '';
  var html = '';
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || [];
    var joined = row.join('').trim();
    if (!joined && i > 15) continue;
    html += '<tr>';
    var cols = Math.max(REPORT_CONFIG.NUM_COLS, row.length);
    for (var j = 0; j < cols; j++) {
      var val = j < row.length ? row[j] : '';
      var displayVal = val === '' || val == null ? '&nbsp;' : escapeHtml_(val).replace(/\n/g, '<br>');
      html +=
        '<td style="border: none; padding: 3px 12px 3px 0; background-color: transparent; color: #1f2937; text-align: left; vertical-align: middle; font-size: 10pt; white-space: nowrap;">' +
        displayVal +
        '</td>';
    }
    html += '</tr>';
  }
  return html;
}

function buildKeidoTableHtmlViaGviz_() {
  var a1 =
    colToA1_(REPORT_CONFIG.START_COL) +
    REPORT_CONFIG.START_ROW +
    ':' +
    colToA1_(REPORT_CONFIG.START_COL + REPORT_CONFIG.NUM_COLS - 1) +
    (REPORT_CONFIG.START_ROW + REPORT_CONFIG.NUM_ROWS - 1);
  var url =
    'https://docs.google.com/spreadsheets/d/' +
    WS_CONFIG.SPREADSHEET_ID +
    '/gviz/tq?tqx=out:csv&sheet=' +
    encodeURIComponent(REPORT_CONFIG.REPORT_SHEET_NAME) +
    '&range=' +
    encodeURIComponent(a1);
  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  });
  if (res.getResponseCode() !== 200) throw new Error('gviz ' + res.getResponseCode());
  var text = res.getContentText() || '';
  if (/<!DOCTYPE html>|<html/i.test(text)) throw new Error('gviz html');
  return rowsToKeidoTableHtml_(parseSimpleCsv_(text));
}

function buildKeidoTableHtmlViaApi_() {
  var range = reportRangeA1_();
  var url =
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(WS_CONFIG.SPREADSHEET_ID) +
    '/values/' +
    encodeURIComponent(range) +
    '?valueRenderOption=FORMATTED_VALUE';
  var json = sheetsApiGetJson_(url);
  return rowsToKeidoTableHtml_(json.values || []);
}

function listSheetNames_(ss) {
  var names = [];
  try {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) names.push(sheets[i].getName());
  } catch (ignored) {}
  return names;
}

function findReportSheet_(ss) {
  if (!ss) return null;
  var preferred = [REPORT_CONFIG.REPORT_SHEET_NAME, '日報', 'data', 'Data', 'dashboard'];
  for (var i = 0; i < preferred.length; i++) {
    var sh = ss.getSheetByName(preferred[i]);
    if (sh) return sh;
  }
  var sheets = ss.getSheets();
  for (var j = 0; j < sheets.length; j++) {
    var name = sheets[j].getName();
    if (/日報|経堂/.test(name)) return sheets[j];
  }
  return null;
}

function openReportSpreadsheet_() {
  var ids = [];
  if (REPORT_CONFIG.SPREADSHEET_ID) ids.push(REPORT_CONFIG.SPREADSHEET_ID);
  if (WS_CONFIG.SPREADSHEET_ID && ids.indexOf(WS_CONFIG.SPREADSHEET_ID) === -1) {
    ids.push(WS_CONFIG.SPREADSHEET_ID);
  }
  var lastErr = '';
  for (var i = 0; i < ids.length; i++) {
    try {
      var ss = SpreadsheetApp.openById(ids[i]);
      var sh = findReportSheet_(ss);
      if (sh) return { ss: ss, sheet: sh, id: ids[i], names: listSheetNames_(ss) };
      lastErr = ids[i] + ' sheets=' + listSheetNames_(ss).join(',');
    } catch (err) {
      lastErr = String(err && err.message ? err.message : err);
    }
  }
  throw new Error(lastErr || '日報シートが見つかりません');
}

function buildKeidoTableHtmlSafe_() {
  try {
    var found = openReportSpreadsheet_();
    var html = buildKeidoTableHtml_(found.sheet);
    if (html) return { html: html, err: '', sheet: found.sheet.getName(), id: found.id };
    return { html: '', err: 'empty table on ' + found.sheet.getName(), sheet: found.sheet.getName(), id: found.id };
  } catch (err) {
    return { html: '', err: String(err && err.message ? err.message : err) };
  }
}

function buildDailyReportPackage_(doneTasks, kansouRaw) {
  var gyomuHtml = buildGyomuNaiyoHtml_(doneTasks);
  var kansouBlockHtml = buildKansouHtml_(kansouRaw);

  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'M月d日');
  var subjectPrefix = 'EAST運営本部 関東運営ブロック' + REPORT_CONFIG.MY_TEAM + '業務日報　' + REPORT_CONFIG.MY_NAME;
  var legacySubjectPrefix = 'EAST運営本部 関東運営ブロック 第7エリア T2　' + REPORT_CONFIG.MY_NAME;
  var subject = subjectPrefix + '　' + formattedDate;

  var lastPTResult = fetchLastPTResultFromGmail_([subjectPrefix, legacySubjectPrefix]);
  var packed = buildKeidoTableHtmlSafe_();
  var tableHtml = packed && packed.html ? packed.html : '';
  if (!tableHtml) {
    tableHtml = '<tr><td style="padding:6px 0;color:#64748b;font-size:10pt;">（経堂数値を取得できませんでした）</td></tr>';
  }

  var baseFontSize = '10.5pt';
  var themeColor = '#1f2937';

  var htmlBody =
    '<div style="font-family: \'Helvetica Neue\', Arial, \'Hiragino Kaku Gothic ProN\', \'Hiragino Sans\', Meiryo, sans-serif; font-size: ' +
    baseFontSize +
    '; color: #2d3748; line-height: 1.6; max-width: 800px;">' +
    '<div style="margin-bottom: 20px; padding: 8px 10px; border:1px solid #e2e8f0; border-radius:8px; background:#f8fafc;">お元気様です。<br>本日の業務日報でございます。<br>ご確認をお願いいたします。</div>' +
    '<div style="margin-bottom: 8px; border-bottom: 1px solid #cbd5e1; display: inline-block; padding-right: 20px;">' +
    '<span style="font-weight: bold; color: ' +
    themeColor +
    '; font-size: 11pt; letter-spacing: 0.05em;">【経堂数値】</span></div>' +
    '<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; width: auto; margin-bottom: 30px; border: none;">' +
    tableHtml +
    '</table>' +
    '<div style="margin-top: 30px;">' +
    '<div style="font-weight: bold; color: ' +
    themeColor +
    '; font-size: 11pt; margin-bottom: 8px; border-bottom: 1px solid #e2e8f0;">【業務内容】</div>' +
    '<div style="padding: 8px 10px 18px 10px; border:1px solid #e5e7eb; border-radius:8px; background:#ffffff;">' +
    gyomuHtml +
    '</div>' +
    '<div style="font-weight: bold; color: ' +
    themeColor +
    '; font-size: 11pt; margin: 16px 0 8px 0; border-bottom: 1px solid #e2e8f0;">【所感】</div>' +
    '<div style="padding: 8px 10px 18px 10px; min-height: 30px; border:1px solid #e5e7eb; border-radius:8px; background:#ffffff;">' +
    kansouBlockHtml +
    '</div>' +
    '<div style="font-weight: bold; color: ' +
    themeColor +
    '; font-size: 11pt; margin: 16px 0 8px 0; border-bottom: 1px solid #e2e8f0;">【PT実績】</div>' +
    '<div style="padding: 8px 10px 18px 10px; border:1px solid #e5e7eb; border-radius:8px; background:#ffffff;">' +
    escapeHtml_(lastPTResult) +
    '</div>' +
    '<div style="border-top: 1px dashed #cbd5e1; border-bottom: 1px dashed #cbd5e1; padding: 15px 0; margin-bottom: 30px; text-align: center;">' +
    '<div style="font-weight: bold; color: #64748b; margin-bottom: 5px; font-size: 8.5pt; letter-spacing: 0.1em;">2026年度 オカモトグループスローガン</div>' +
    '<div style="color: ' +
    themeColor +
    '; font-weight: bold; font-size: 11pt;">「信頼で つなぐ未来と 地域の輪」</div></div>' +
    '<div style="line-height: 1.7; color: #4a5568;">' +
    '<div style="font-size: 9pt;">EAST運営本部　関東運営ブロック　' +
    REPORT_CONFIG.MY_TEAM +
    '</div>' +
    '<div style="margin: 5px 0;"><strong style="font-size: 11pt; color: ' +
    themeColor +
    ';">JOYFIT24経堂</strong></div>' +
    '<div style="margin-bottom: 10px;">' +
    '<span style="font-size: 12.5pt; font-weight: bold; color: #1a202c;">日下　竜太</span>' +
    '<span style="color: #718096; font-size: 9.5pt; margin-left: 8px;">Ryuta Kusaka</span></div>' +
    '<div style="font-size: 9pt; color: #718096;">' +
    '〒156-0052 東京都世田谷区経堂5-23-13<br>' +
    'TEL：03-6804-4100 / FAX：03-6804-4103' +
    '<br><a href="' +
    AVAILABILITY_URL +
    '" style="display:inline-block; margin-top:6px; color:#1d4ed8; text-decoration:none; font-weight:700; border:1px solid #93c5fd; background:#eff6ff; padding:3px 8px; border-radius:999px;" target="_blank">空きスケジュールはこちら</a>' +
    '</div></div></div></div>';

  var gyomuPlain = (doneTasks && doneTasks.length)
    ? doneTasks.map(function (t) { return '・' + t; }).join('\n')
    : '（業務内容が空です）';
  var kansouPlain = normalizeKansouText_(kansouRaw) || '（所感が空です）';
  var previewText =
    '件名: ' +
    subject +
    '\n\n【業務内容】\n' +
    gyomuPlain +
    '\n\n【所感】\n' +
    kansouPlain +
    '\n\n【PT実績】\n' +
    lastPTResult +
    '\n\n※経堂数値はメール本文の表をご確認ください。';

  return { ok: true, subject: subject, htmlBody: htmlBody, previewText: previewText };
}

function getTodayDoneTasksFromSync_(ss) {
  var sheet = ss.getSheetByName(WS_CONFIG.SYNC_SHEET_NAME);
  if (!sheet) return [];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (formatDateCell_(data[r][0]) === today) {
      return parseJsonSafe_(data[r][2]);
    }
  }
  return [];
}

function getTodayKansouFromSync_(ss) {
  var sheet = ss.getSheetByName(WS_CONFIG.SYNC_SHEET_NAME);
  if (!sheet) return '';
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (formatDateCell_(data[r][0]) === today) {
      return data[r][4] != null ? String(data[r][4]) : '';
    }
  }
  return '';
}

function buildKansouHtml_(raw) {
  var normalized = normalizeKansouText_(raw);
  if (!normalized) {
    return '<br>';
  }
  return escapeHtml_(normalized).replace(/\r\n/g, '\n').replace(/\n/g, '<br>');
}

/**
 * 旧フォーマット混在時でも「所感」だけを残す。
 * - 【業務・対応内容】ブロックは削除
 * - 【本日の所感・感想】/【所感】見出しは削除
 */
function normalizeKansouText_(raw) {
  if (raw == null) return '';
  var text = String(raw).replace(/\r\n/g, '\n');
  text = text.replace(/【業務・対応内容】[\s\S]*?(?=【本日の所感・感想】|【所感】|$)/g, '');
  text = text.replace(/【本日の所感・感想】/g, '');
  text = text.replace(/【所感】/g, '');
  text = text.replace(/^\s+|\s+$/g, '');
  text = text.replace(/\n{3,}/g, '\n\n');
  return text;
}

function buildGyomuNaiyoHtml_(tasks) {
  if (!tasks || tasks.length === 0) {
    return '・<span style="color:#94a3b8;">（業務内容欄が空です。入力してから再度お試しください）</span>';
  }
  var parts = [];
  for (var i = 0; i < tasks.length; i++) {
    parts.push('・' + escapeHtml_(tasks[i]));
  }
  return parts.join('<br>');
}

function escapeHtml_(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function fetchLastPTResultFromGmail_(subjectPrefixes) {
  var lastPTResult = '計　60分0本　30分0本';
  try {
    var prefixes = Array.isArray(subjectPrefixes) ? subjectPrefixes : [String(subjectPrefixes || '')];
    var now = new Date();
    for (var p = 0; p < prefixes.length; p++) {
      var prefix = prefixes[p];
      if (!prefix) continue;
      var threads = GmailApp.search('subject:"' + prefix + '"', 0, 20);
      for (var t = 0; t < threads.length; t++) {
        var thread = threads[t];
        var messages = thread.getMessages();
        for (var i = messages.length - 1; i >= 0; i--) {
          var msg = messages[i];
          if (msg.getDate().toDateString() === now.toDateString()) continue;
          var body = msg.getPlainBody();
          var match = body.match(/計\s+\d+分\d+本\s+\d+分\d+本/);
          if (match) {
            lastPTResult = match[0];
            break;
          }
        }
        if (lastPTResult !== '計　60分0本　30分0本') break;
      }
      if (lastPTResult !== '計　60分0本　30分0本') break;
    }
  } catch (e) {
    console.error('PT実績取得:', e);
  }
  return lastPTResult;
}

function buildKeidoTableHtml_(sheet) {
  var startRow = REPORT_CONFIG.START_ROW;
  var startCol = REPORT_CONFIG.START_COL;
  var numRows = REPORT_CONFIG.NUM_ROWS;
  var numCols = REPORT_CONFIG.NUM_COLS;

  var range = sheet.getRange(startRow, startCol, numRows, numCols);
  var values = range.getDisplayValues();
  var fontColors = range.getFontColors();
  var backgrounds = range.getBackgrounds();
  var hAligns = range.getHorizontalAlignments();
  var vAligns = range.getVerticalAlignments();
  var fontWeights = range.getFontWeights();

  var merges = range.getMergedRanges();
  var mergeMap = [];
  for (var ri = 0; ri < numRows; ri++) {
    mergeMap[ri] = [];
    for (var ci = 0; ci < numCols; ci++) {
      mergeMap[ri][ci] = { skip: false, rowSpan: 1, colSpan: 1 };
    }
  }

  merges.forEach(function (rng) {
    var mRow = rng.getRow() - startRow;
    var mCol = rng.getColumn() - startCol;
    var mNumRows = rng.getNumRows();
    var mNumCols = rng.getNumColumns();
    if (mRow >= 0 && mRow < numRows && mCol >= 0 && mCol < numCols) {
      mergeMap[mRow][mCol].rowSpan = mNumRows;
      mergeMap[mRow][mCol].colSpan = mNumCols;
      for (var r = 0; r < mNumRows; r++) {
        for (var c = 0; c < mNumCols; c++) {
          if (r === 0 && c === 0) continue;
          if (mRow + r < numRows && mCol + c < numCols) {
            mergeMap[mRow + r][mCol + c].skip = true;
          }
        }
      }
    }
  });

  var tableFontSize = '10pt';
  var htmlBody = '';
  for (var i = 0; i < numRows; i++) {
    var rowValues = values[i].join('').trim();
    var isRowEmpty =
      rowValues === '' &&
      backgrounds[i].every(function (b) {
        return b === '#ffffff' || b === 'white';
      });
    if (isRowEmpty && i > 15) continue;

    htmlBody += '<tr>';
    for (var j = 0; j < numCols; j++) {
      if (mergeMap[i][j].skip) continue;
      var val = values[i][j];
      var displayVal = val === '' ? '&nbsp;' : escapeHtml_(val).replace(/\n/g, '<br>');
      var bg = backgrounds[i][j];
      var color = fontColors[i][j];
      var hAlign = hAligns[i][j];
      var vAlign = vAligns[i][j];
      var weight = fontWeights[i][j];
      var spanAttr = '';
      if (mergeMap[i][j].rowSpan > 1) spanAttr += ' rowspan="' + mergeMap[i][j].rowSpan + '"';
      if (mergeMap[i][j].colSpan > 1) spanAttr += ' colspan="' + mergeMap[i][j].colSpan + '"';
      htmlBody +=
        '<td' +
        spanAttr +
        ' style="border: none; padding: 3px 12px 3px 0; background-color: ' +
        (bg === '#ffffff' || bg === 'white' ? 'transparent' : bg) +
        '; color: ' +
        color +
        '; text-align: ' +
        hAlign +
        '; vertical-align: ' +
        vAlign +
        '; font-weight: ' +
        weight +
        '; font-size: ' +
        tableFontSize +
        '; white-space: nowrap;">' +
        displayVal +
        '</td>';
    }
    htmlBody += '</tr>';
  }
  return htmlBody;
}

/**
 * Workspace用のスプレッドシートを開く。
 * openById が権限で失敗した場合は、バインド先スプレッドシートへフォールバックを試行。
 */
function authorizeWorkspaceAccess() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  var ss = SpreadsheetApp.openById(WS_CONFIG.SPREADSHEET_ID);
  return { ok: true, name: ss.getName() };
}

function openWorkspaceSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(WS_CONFIG.SPREADSHEET_ID);
  } catch (err) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        return active;
      }
    } catch (ignored) {}
    throw new Error(
      'Workspaceシートにアクセスできません。' +
        'Webアプリの実行ユーザー/権限を再確認してください。' +
        ' detail=' +
        String((err && err.message) || err)
    );
  }
}

/**
 * Gemini で所感を校閲（敬語・誤字・分量）。
 * スクリプトのプロパティに GEMINI_API_KEY を設定（Google AI Studio で発行可）。
 */
function polishKansouWithGemini(rawText) {
  try {
    if (!rawText || String(rawText).trim() === '') {
      return { ok: false, message: 'テキストを入力してください。' };
    }
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) {
      return {
        ok: false,
        message:
          'GEMINI_API_KEY が未設定です。プロジェクトの設定 → スクリプトのプロパティ にキーを追加してください。',
      };
    }
    var prompt =
      '以下はフィットネス施設のスタッフ日報「所感」欄の下書きです。ビジネスメール向けの敬語に整え、誤字脱字を修正し、300文字以内で簡潔にまとめてください。事実と意味は変えないでください。出力は所感の本文のみ（説明・見出し・引用符は不要）。\n\n' +
      String(rawText);
    var models = [
      'gemini-2.5-flash-lite',
      'gemini-3.1-flash-lite',
      'gemini-2.5-flash',
      'gemini-3.5-flash'
    ];
    var lastMessage = '返答を取得できませんでした。';
    for (var m = 0; m < models.length; m++) {
      var url =
        'https://generativelanguage.googleapis.com/v1beta/models/' +
        encodeURIComponent(models[m]) +
        ':generateContent?key=' +
        encodeURIComponent(key);
      var res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.4, maxOutputTokens: 1024 }
        }),
        muteHttpExceptions: true
      });
      var code = res.getResponseCode();
      var raw = res.getContentText() || '';
      var json = null;
      try {
        json = raw ? JSON.parse(raw) : null;
      } catch (parseErr) {
        lastMessage = 'Gemini の応答が JSON ではありません（' + code + '）';
        continue;
      }
      if (code !== 200) {
        lastMessage = (json && json.error && json.error.message) || ('API エラー（コード ' + code + '）');
        if (code === 404 || /not found|deprecated|INVALID_ARGUMENT/i.test(String(lastMessage))) {
          continue;
        }
        return { ok: false, message: lastMessage };
      }
      var text = extractGeminiText_(json);
      if (text) {
        return { ok: true, text: String(text).trim() };
      }
      lastMessage = '返答を取得できませんでした。';
    }
    return { ok: false, message: lastMessage };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function extractGeminiText_(json) {
  try {
    var parts = json.candidates[0].content.parts || [];
    var out = [];
    for (var i = 0; i < parts.length; i++) {
      if (parts[i] && parts[i].thought) continue;
      var t = parts[i] && parts[i].text ? String(parts[i].text).trim() : '';
      if (t) out.push(t);
    }
    return out.join('\n').trim();
  } catch (e) {
    return '';
  }
}

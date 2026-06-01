/**
 * RYUTA Workspace — サーバ側
 * スプレッドシートIDは 1 本にまとめる想定（日報ブックに WorkspaceSync シートを追加する形を推奨）
 */
var WS_CONFIG = {
  /** タスク同期を書き込むスプレッドシートID（URLの /d/ と /edit/ の間） */
  SPREADSHEET_ID: '14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w',
  /** 同期用シート名（なければ自動作成） */
  SYNC_SHEET_NAME: 'WorkspaceSync',
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
      api === 'polishKansou';
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

/**
 * 日報メール（V10 相当）— 宛先・署名はここで編集
 */
var REPORT_CONFIG = {
  RECIPIENT:
    'y_east_staff@okamoto-group.co.jp, yamauchieastnippou_transfer@okamoto-group.co.jp, jf-kyoudou@okamoto-group.co.jp',
  CC: 'k-takakuwa@okamoto-group.co.jp, m-akiyama@okamoto-group.co.jp, t-doi@okamoto-group.co.jp',
  // 26年度運用: テリトリー呼称は「エリア-番号」表記（例: 7-2）
  MY_TEAM: '7-2',
  MY_NAME: '日下竜太',
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
    var saveResult = saveWorkspaceTasksToSheet(
      JSON.stringify(active),
      JSON.stringify(doneTasks),
      kansou
    );
    if (!saveResult.ok) return saveResult;
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

function buildDailyReportPackage_(doneTasks, kansouRaw) {
  if (WS_CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    return { ok: false, message: 'SPREADSHEET_ID を設定してください。' };
  }
  var ss = openWorkspaceSpreadsheet_();
  var reportSheet = ss.getSheetByName(REPORT_CONFIG.REPORT_SHEET_NAME);
  if (!reportSheet) {
    return { ok: false, message: '「' + REPORT_CONFIG.REPORT_SHEET_NAME + '」シートが見つかりません。' };
  }

  var gyomuHtml = buildGyomuNaiyoHtml_(doneTasks);
  var kansouBlockHtml = buildKansouHtml_(kansouRaw);

  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'M月d日');
  var subjectPrefix = 'EAST運営本部 関東運営ブロック' + REPORT_CONFIG.MY_TEAM + '業務日報　' + REPORT_CONFIG.MY_NAME;
  var legacySubjectPrefix = 'EAST運営本部 関東運営ブロック 第7エリア T2　' + REPORT_CONFIG.MY_NAME;
  var subject = subjectPrefix + '　' + formattedDate;

  var lastPTResult = fetchLastPTResultFromGmail_([subjectPrefix, legacySubjectPrefix]);
  var tableHtml = buildKeidoTableHtml_(reportSheet);

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
function openWorkspaceSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(WS_CONFIG.SPREADSHEET_ID);
  } catch (err) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        // バインド型プロジェクトならこれで継続可能。
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
    var url =
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' +
      encodeURIComponent(key);
    var body = {
      contents: [
        {
          parts: [
            {
              text:
                '以下はフィットネス施設のスタッフ日報「所感」欄の下書きです。ビジネスメール向けの敬語に整え、誤字脱字を修正し、300文字以内で簡潔にまとめてください。事実と意味は変えないでください。出力は所感の本文のみ（説明・見出し・引用符は不要）。\n\n' +
                String(rawText),
            },
          ],
        },
      ],
    };
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    var code = res.getResponseCode();
    var json = JSON.parse(res.getContentText());
    if (code !== 200) {
      return {
        ok: false,
        message: (json.error && json.error.message) || 'API エラー（コード ' + code + '）',
      };
    }
    var text =
      json.candidates &&
      json.candidates[0] &&
      json.candidates[0].content &&
      json.candidates[0].content.parts &&
      json.candidates[0].content.parts[0] &&
      json.candidates[0].content.parts[0].text;
    if (!text) {
      return { ok: false, message: '返答を取得できませんでした。' };
    }
    return { ok: true, text: String(text).trim() };
  } catch (e) {
    console.error(e);
    return { ok: false, message: String(e.message || e) };
  }
}

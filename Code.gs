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

/** 受付状況表（nippo）の Web API。更新本体は受付状況表側のトリガーで非同期実行される */
var RECEPTION_REFRESH_URL_ =
  'https://script.google.com/macros/s/AKfycbyQzrG0awDLjYqoBELKF78nphmcycUJRdEQRNq2i3SCjYulMIIgB5dspSnYRF-jfslK/exec';
var RECEPTION_REFRESH_TOKEN_ = 'kyodo-ws-refresh-7f3c91';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('今日の作業')
    .addItem('トップを表示', 'hubShowHome_')
    .addItem('見た目を整える', 'applyFourColorFromMenu')
    .addItem('未納を再取得', 'fillUnpaidFromMenu')
    .addItem('チェックを元へ反映', 'flushUnpaidFromMenu')
    .addToUi();
  ui.createMenu('数値更新')
    .addItem('受付状況表の数値を更新（入会・退会・OP）', 'refreshReceptionNumbersFromMenu')
    .addItem('前回の更新時刻を確認', 'showReceptionRefreshStatus')
    .addToUi();
  try {
    var unpaid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNPAID_SHEET_);
    if (unpaid) unpaid.showRows(1, unpaid.getMaxRows());
  } catch (eShow) {}
}

function onEdit(e) {
  try { masterMonthTabSelect_(e); } catch (err) {}
  try { unpaidOnEditSimple_(e); } catch (errU) {}
  try { joinListOnEditSimple_(e); } catch (errJ) {}
}

function onSelectionChange(e) {
  try { masterMonthTabSelect_(e); } catch (err) {}
}

function callReceptionRefreshApi_(api) {
  var res = UrlFetchApp.fetch(
    RECEPTION_REFRESH_URL_ + '?api=' + api + '&token=' + encodeURIComponent(RECEPTION_REFRESH_TOKEN_),
    { muteHttpExceptions: true, followRedirects: true }
  );
  try {
    return JSON.parse(res.getContentText());
  } catch (e) {
    return { ok: false, message: 'HTTP ' + res.getResponseCode() };
  }
}

function refreshReceptionNumbersFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var before = callReceptionRefreshApi_('refreshStatus');
  var beforeAt = (before && before.lastRefreshed) || '';
  var r = callReceptionRefreshApi_('refreshNumbers');
  if (!r || !r.ok) {
    ss.toast((r && r.message) || '更新を開始できませんでした', '数値更新', 10);
    return r;
  }
  ss.toast('更新中です…（1〜2分）完了したらお知らせします', '数値更新', 150);

  // メニュー実行は最大6分。5分まで完了を待つ
  var deadline = Date.now() + 5 * 60 * 1000;
  while (Date.now() < deadline) {
    Utilities.sleep(10000);
    var s = callReceptionRefreshApi_('refreshStatus');
    if (s && s.ok && s.error) {
      ss.toast('更新でエラーが出ました: ' + s.error, '数値更新', 15);
      return s;
    }
    if (s && s.ok && !s.pending && s.lastRefreshed && s.lastRefreshed !== beforeAt) {
      SpreadsheetApp.flush();
      ss.toast('更新が完了しました（' + s.lastRefreshed + '）', '数値更新', 15);
      return s;
    }
  }
  ss.toast('まだ終わっていません。少ししてから「前回の更新時刻を確認」で見てください。', '数値更新', 15);
  return r;
}

function showReceptionRefreshStatus() {
  var r = callReceptionRefreshApi_('refreshStatus');
  var msg = !r || !r.ok
    ? '状態を取得できませんでした'
    : r.pending
      ? '更新中です。少しお待ちください。'
      : '前回の更新: ' + (r.lastRefreshed || '記録なし');
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '数値更新', 8);
  return r;
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

/** 集約ブックのシート一覧（整理用） */
function listWorkspaceSheets_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sheets = ss.getSheets();
    var out = [];
    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      out.push({
        name: sh.getName(),
        gid: sh.getSheetId(),
        rows: sh.getLastRow(),
        cols: sh.getLastColumn(),
        hidden: sh.isSheetHidden()
      });
    }
    return {
      ok: true,
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      title: ss.getName(),
      sheets: out
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/**
 * deta からリンクを吸い上げ、モノトーンの「URL一覧」に整形。
 * 不要シート（MEMO / dashboard / シート4 / deta）は削除。
 */
function rebuildUrlIndexSheet_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sourceNames = ['deta', 'DETA', 'data', 'Data'];
    var source = null;
    for (var i = 0; i < sourceNames.length; i++) {
      source = ss.getSheetByName(sourceNames[i]);
      if (source) break;
    }
    if (!source) {
      return { ok: false, message: 'deta シートが見つかりません' };
    }

    var links = extractLinksFromSheet_(source);
    var index = ss.getSheetByName('URL一覧');
    if (index) {
      ss.deleteSheet(index);
    }
    index = ss.insertSheet('URL一覧', 0);
    styleUrlIndexSheet_(index, links);

    var drop = ['MEMO', 'dashboard', 'シート4', 'deta', 'DETA'];
    var deleted = [];
    for (var d = 0; d < drop.length; d++) {
      var doomed = ss.getSheetByName(drop[d]);
      if (!doomed) continue;
      // 最後の1枚は消せないので、URL一覧以外が残っているときだけ削除
      if (ss.getSheets().length <= 1) break;
      ss.deleteSheet(doomed);
      deleted.push(drop[d]);
    }

    return {
      ok: true,
      linkCount: links.length,
      deleted: deleted,
      sheet: 'URL一覧',
      spreadsheetUrl: ss.getUrl(),
      links: links
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function extractLinksFromSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var rich = sheet.getRange(1, 1, lastRow, lastCol).getRichTextValues();
  var formulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
  var found = [];
  var seen = {};

  function pushLink(url, label, note) {
    url = String(url || '').trim();
    if (!url) return;
    if (!/^https?:\/\//i.test(url)) {
      if (/^docs\.google\.com\//i.test(url) || /^drive\.google\.com\//i.test(url)) {
        url = 'https://' + url;
      } else {
        return;
      }
    }
    var key = url.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    label = String(label || '').replace(/\s+/g, ' ').trim();
    if (!label || label === url) label = guessLinkLabel_(url);
    found.push({
      label: label,
      url: url,
      kind: classifyLinkKind_(url, label),
      note: String(note || '').trim()
    });
  }

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var text = values[r][c];
      var formula = formulas[r][c];
      var rt = rich[r][c];

      if (formula) {
        var hm = String(formula).match(/HYPERLINK\s*\(\s*"([^"]+)"\s*(?:,\s*"([^"]*)")?\s*\)/i);
        if (hm) pushLink(hm[1], hm[2] || text, '');
      }

      if (rt) {
        try {
          var runs = rt.getRuns ? rt.getRuns() : [];
          if (runs && runs.length) {
            for (var k = 0; k < runs.length; k++) {
              var runUrl = runs[k].getLinkUrl && runs[k].getLinkUrl();
              if (runUrl) pushLink(runUrl, runs[k].getText() || text, '');
            }
          } else if (rt.getLinkUrl) {
            var one = rt.getLinkUrl();
            if (one) pushLink(one, text, '');
          }
        } catch (ignoredRt) {}
      }

      var cell = String(text == null ? '' : text);
      var urlMatches = cell.match(/https?:\/\/[^\s<>"']+/gi) || [];
      for (var u = 0; u < urlMatches.length; u++) {
        var cleaned = urlMatches[u].replace(/[),．。]+$/g, '');
        pushLink(cleaned, cell.replace(urlMatches[u], '').trim() || cleaned, '');
      }
    }
  }

  // タイトルだけの行で、同じ行にURLが無いものは「メモ」として残さない（URLのみ方針）
  found.sort(function (a, b) {
    if (a.kind !== b.kind) return String(a.kind).localeCompare(String(b.kind));
    return String(a.label).localeCompare(String(b.label), 'ja');
  });
  return found;
}

function guessLinkLabel_(url) {
  try {
    var u = String(url || '');
    if (/docs\.google\.com\/spreadsheets/i.test(u)) return 'Google スプレッドシート';
    if (/docs\.google\.com\/document/i.test(u)) return 'Google ドキュメント';
    if (/docs\.google\.com\/presentation/i.test(u)) return 'Google スライド';
    if (/drive\.google\.com/i.test(u)) return 'Google ドライブ';
    if (/vercel\.app/i.test(u)) return 'Vercel App';
    if (/script\.google\.com/i.test(u)) return 'Google Apps Script';
    var host = u.replace(/^https?:\/\//i, '').split('/')[0];
    return host || u;
  } catch (e) {
    return url;
  }
}

function classifyLinkKind_(url, label) {
  var u = String(url || '').toLowerCase();
  var l = String(label || '');
  if (/シフト|shift|キンタイ|カレンダー/.test(l) || /shift/i.test(u)) return 'SHIFT';
  if (/pt|パーソナル|予約/.test(l.toLowerCase())) return 'PT';
  if (/口コミ|review|リプクル/.test(l) || /review/i.test(u)) return 'REVIEW';
  if (/qa|qanda|未収|unpaid/.test(l.toLowerCase()) || /qa|unpaid/i.test(u)) return 'OPS';
  if (/todo|タスク|ダッシュボード|dashboard/.test(l.toLowerCase())) return 'WORK';
  if (/docs\.google\.com|drive\.google\.com/.test(u)) return 'DOCS';
  if (/vercel\.app|script\.google\.com/.test(u)) return 'APP';
  return 'LINK';
}

/** 追加販促ブック（スタッフ入力）→ 見た目整形 + Workspace へ IMPORTRANGE */
var PROMO_SOURCE_ID_ = '1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E';
/** 学校関係者割 Googleフォーム回答。元ブックは読み取りのみ */
var SCHOOL_FORM_SOURCE_ID_ = '1mmG_xM1WoWFKgpmOl5obsKXo_0GjWLnAnLY9hTGanVg';
var SCHOOL_FORM_SHEET_ = 'フォームの回答 1';
var SCHOOL_FORM_COLS_ = 16;
var SCHOOL_WS_SHEET_ = '学割';
/** 件数上限ではない。フォームが増えた分を QUERY が全部こぼせる行数 */
var SCHOOL_MIN_ROWS_ = 3000;

function setupPromoImport_() {
  try {
    var source = SpreadsheetApp.openById(PROMO_SOURCE_ID_);
    var dest = openWorkspaceSpreadsheet_();
    var styled = [];
    var imported = [];
    var sheets = source.getSheets();

    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      if (/学校関係者/.test(String(sh.getName() || ''))) continue;
      var meta = stylePromoSheetKeepValues_(sh);
      styled.push(meta);

      var destName = '販促_' + String(sh.getName()).replace(/[\\\/\?\*\[\]]/g, '').slice(0, 80);
      var existing = dest.getSheetByName(destName);
      if (existing) dest.deleteSheet(existing);
      var mirror = dest.insertSheet(destName);
      styleImportMirrorSheet_(mirror, source.getId(), sh.getName(), meta.usedCols, meta.headers);
      imported.push({ name: destName, source: sh.getName(), cols: meta.usedCols, rows: meta.usedRows });
    }

    // ハブは作らず、ミラーだけ（余計な説明行なし）
    var oldHub = dest.getSheetByName('追加販促');
    if (oldHub) {
      try { dest.deleteSheet(oldHub); } catch (eHub) {}
    }

    // 作成時に残った空のデフォルトシートを除去
    var leftovers = dest.getSheets();
    for (var j = leftovers.length - 1; j >= 0; j--) {
      var nm = leftovers[j].getName();
      if (/^シート\d+$/.test(nm) && leftovers[j].getLastRow() === 0) {
        try {
          if (dest.getSheets().length > 1) dest.deleteSheet(leftovers[j]);
        } catch (eDel) {}
      }
    }

    return {
      ok: true,
      sourceId: PROMO_SOURCE_ID_,
      sourceTitle: source.getName(),
      sourceUrl: source.getUrl(),
      styled: styled,
      imported: imported,
      workspaceUrl: dest.getUrl(),
      note: '初回は Workspace 側で IMPORTRANGE の「アクセスを許可」が必要な場合があります'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/** 値は一切変えず、見た目と余分な空列のみ整理 */
function stylePromoSheetKeepValues_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var usedCols = 0;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c] || '').trim() !== '') usedCols = c + 1;
  }
  if (usedCols < 1) usedCols = 1;

  // 実際にデータがある最終行
  var usedRows = 1;
  if (lastRow > 1) {
    var vals = sheet.getRange(2, 1, lastRow - 1, usedCols).getDisplayValues();
    for (var r = 0; r < vals.length; r++) {
      var rowHas = false;
      for (var k = 0; k < vals[r].length; k++) {
        if (String(vals[r][k] || '').trim() !== '') {
          rowHas = true;
          break;
        }
      }
      if (rowHas) usedRows = r + 2;
    }
  }

  sheet.setHiddenGridlines(true);
  sheet.setTabColor(hubTabColorFor_(sheet.getName()));
  sheet.setFrozenRows(1);

  var dn = dnTheme_();
  var headerRange = sheet.getRange(1, 1, 1, usedCols);
  headerRange
    .setBackground(dn.ink)
    .setFontColor(dn.paper)
    .setFontFamily('Noto Sans JP').setFontStyle('italic')
    .setFontSize(10)
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBorder(false, false, true, false, false, false, dn.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(1, 32);

  if (usedRows >= 2) {
    var body = sheet.getRange(2, 1, usedRows - 1, usedCols);
    body
      .setBackground(dn.paper)
      .setFontColor(dn.ink)
      .setFontFamily('Noto Sans JP').setFontStyle('italic')
      .setFontSize(10)
      .setVerticalAlignment('middle');
  }

  // 空列を隠す（削除するとフォーム連携が壊れることがあるので非表示）
  var maxCols = sheet.getMaxColumns();
  for (var col = 1; col <= usedCols; col++) {
    try { sheet.showColumns(col); } catch (eShow) {}
    sheet.setColumnWidth(col, col === 1 ? 150 : 140);
  }
  if (maxCols > usedCols) {
    try { sheet.hideColumns(usedCols + 1, maxCols - usedCols); } catch (eHide) {}
  }

  // URLっぽい列は少し広げる / 日時は表示形式のみ
  for (var h = 0; h < usedCols; h++) {
    var title = String(headers[h] || '');
    if (/リンク|URL|画像|写真/i.test(title)) sheet.setColumnWidth(h + 1, 260);
    if (/日時|申請/.test(title)) {
      sheet.setColumnWidth(h + 1, 160);
      if (usedRows >= 2) {
        sheet.getRange(2, h + 1, usedRows - 1, 1).setNumberFormat('yyyy/mm/dd HH:mm');
      }
    }
    if (/メール|mail/i.test(title)) sheet.setColumnWidth(h + 1, 200);
  }

  return {
    name: sheet.getName(),
    usedCols: usedCols,
    usedRows: usedRows,
    headers: headers.slice(0, usedCols)
  };
}

function styleImportMirrorSheet_(sheet, sourceId, sourceSheetName, usedCols, headers) {
  sheet.clear();
  try { sheet.clearConditionalFormatRules(); } catch (eClr) {}
  try { sheet.getDataRange().clearDataValidations(); } catch (eVal) {}
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(hubTabColorFor_(sheet.getName()));
  sheet.setFrozenRows(1);

  var cols = Math.max(Number(usedCols) || 1, 1);
  var endCol = columnLetter_(cols);
  // 余計な見出しなし。元シートと同じ範囲をそのまま表示
  var formula =
    '=IMPORTRANGE("' +
    sourceId +
    '","' +
    sourceSheetName.replace(/"/g, '""') +
    '!A:' +
    endCol +
    '")';
  sheet.getRange(1, 1).setFormula(formula);

  for (var c = 1; c <= cols; c++) {
    sheet.setColumnWidth(c, c === 1 ? 160 : 140);
  }

  var hdrs = headers || [];
  var checkCols = [];
  for (var h = 0; h < hdrs.length; h++) {
    var title = String(hdrs[h] || '');
    if (/リンク|URL|画像|写真/i.test(title)) sheet.setColumnWidth(h + 1, 260);
    if (/日時|申請|入会日|タイムスタンプ/.test(title)) {
      sheet.setColumnWidth(h + 1, 160);
    }
    if (/メールアドレス|mail/i.test(title) && !/レクチャー|アンケート|付与/.test(title)) {
      sheet.setColumnWidth(h + 1, 220);
    }
    if (isCheckboxHeader_(title)) {
      checkCols.push(h + 1);
      sheet.setColumnWidth(h + 1, 120);
    }
  }

  // 枠線なし・チェック列は下まで事前適用（行が増えても FALSE 文字に戻らない）
  formatImportMirrorChrome_(sheet, cols, checkCols, hdrs);
}

/**
 * IMPORTRANGE はそのまま。枠線削除＋チェック列は「実データ行＋余裕」だけに適用。
 * （最大行まで付けると空行にチェックが浮かんで見た目がずれる）
 */
function formatImportMirrorChrome_(sheet, cols, checkCols, headers) {
  cols = Math.max(Number(cols) || Math.max(sheet.getLastColumn(), 1), 1);
  var maxR = Math.max(sheet.getMaxRows(), 2);
  var maxC = Math.max(sheet.getMaxColumns(), cols);

  try {
    sheet.getRange(1, 1, maxR, maxC).setBorder(false, false, false, false, false, false);
  } catch (eBorder) {}
  try {
    sheet.getRange(1, 1, maxR, maxC).clearDataValidations();
  } catch (eValAll) {}
  try {
    sheet.clearConditionalFormatRules();
  } catch (eRules) {}

  var hdrs = headers;
  if (!hdrs || !hdrs.length) {
    try {
      hdrs = sheet.getRange(1, 1, 1, cols).getDisplayValues()[0];
    } catch (eH) {
      hdrs = [];
    }
  }

  if (!checkCols || !checkCols.length) {
    checkCols = [];
    for (var h = 0; h < hdrs.length; h++) {
      if (isCheckboxHeader_(hdrs[h])) checkCols.push(h + 1);
    }
  }

  var lastData = findImportLastDataRow_(sheet, cols);
  var buffer = 80; // 新規行がしばらく増えてもカバー
  var endRow = Math.min(Math.max(lastData + buffer, 2), maxR);
  var bodyRows = endRow - 1;

  try {
    var dnChrome = dnTheme_();
    sheet.getRange(1, 1, 1, cols)
      .setBackground(dnChrome.ink)
      .setFontColor(dnChrome.paper)
      .setFontFamily('Noto Sans JP').setFontStyle('italic')
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(false, false, true, false, false, false, dnChrome.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(1, 32);
    if (bodyRows >= 1) {
      sheet.getRange(2, 1, bodyRows, cols)
        .setBackground(dnChrome.paper)
        .setFontColor(dnChrome.ink)
        .setFontFamily('Noto Sans JP').setFontStyle('italic')
        .setFontSize(10)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('center');
    }
  } catch (eStyle) {}

  for (var d = 0; d < hdrs.length; d++) {
    var title = String(hdrs[d] || '');
    if (/日時|申請|入会日|タイムスタンプ/.test(title) && bodyRows >= 1) {
      try {
        sheet.getRange(2, d + 1, bodyRows, 1).setNumberFormat('yyyy/mm/dd HH:mm');
      } catch (eFmt) {}
    }
  }

  var applied = [];
  if (checkCols.length && bodyRows >= 1) {
    try {
      var rules = [];
      for (var i = 0; i < checkCols.length; i++) {
        var col = checkCols[i];
        var colLetter = columnLetter_(col);
        var range = sheet.getRange(2, col, bodyRows, 1);
        range.setDataValidation(
          SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build()
        );
        range.setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.setColumnWidth(col, 140);
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied('=' + colLetter + '2=TRUE')
            .setBackground(dnTheme_().ink)
            .setFontColor(dnTheme_().paper)
            .setBold(true)
            .setRanges([range])
            .build()
        );
        applied.push({
          col: col,
          header: String(hdrs[col - 1] || ''),
          rows: bodyRows,
          lastData: lastData
        });
      }
      sheet.setConditionalFormatRules(rules);
    } catch (eCheck) {}
  }

  hubStampType_(sheet, endRow, cols);
  return {
    ok: true,
    sheet: sheet.getName(),
    cols: cols,
    lastData: lastData,
    endRow: endRow,
    bodyRows: bodyRows,
    checkCols: applied,
    bordersCleared: true
  };
}

/** A〜先頭数列で、表示値がある最終行を返す（ヘッダー=1） */
function findImportLastDataRow_(sheet, cols) {
  var scanCols = Math.min(Math.max(Number(cols) || 3, 1), 12);
  var maxScan = Math.min(Math.max(sheet.getMaxRows(), 2), 3000);
  var vals = sheet.getRange(1, 1, maxScan, scanCols).getDisplayValues();
  var last = 1;
  for (var r = 1; r < vals.length; r++) {
    var nonempty = false;
    for (var c = 0; c < vals[r].length; c++) {
      if (String(vals[r][c] == null ? '' : vals[r][c]).trim() !== '') {
        nonempty = true;
        break;
      }
    }
    if (nonempty) last = r + 1;
  }
  return last;
}

/**
 * 経堂ブロック D/E ＋ FIT365ブロック I/J など、チェック対象列を決める。
 * ヘッダー判定を優先し、不足時は経堂=D/E・FIT365=I/J を補完。
 */
function resolveJoinListCheckCols_(headers, cols) {
  var out = [];
  var seen = {};
  function add(col) {
    col = Number(col);
    if (!col || col < 1 || (cols && col > cols) || seen[col]) return;
    seen[col] = true;
    out.push(col);
  }
  for (var h = 0; h < (headers || []).length; h++) {
    if (isCheckboxHeader_(headers[h])) add(h + 1);
  }
  // 経堂（A1=IMPORTRANGE A:E）
  add(4);
  add(5);
  // FIT365（F1=IMPORTRANGE … → I/J が D/E 相当）
  if (cols >= 10) {
    add(9);
    add(10);
  }
  return out;
}

/** 見た目は触らず、枠線とフィルタだけ掃除（手動整形を維持） */
function tidyImportMirrorKeepLook_(sheet) {
  try {
    var a1 = '';
    try {
      a1 = String(sheet.getRange(1, 1).getFormula() || '');
    } catch (eF) {}
    var maxR = Math.max(sheet.getMaxRows(), 2);
    var maxC = Math.max(sheet.getMaxColumns(), 1);
    try {
      sheet.getRange(1, 1, maxR, maxC).setBorder(false, false, false, false, false, false);
    } catch (eB) {}
    removeSheetFilterSafe_(sheet);
    return {
      ok: true,
      sheet: sheet.getName(),
      importrange: /IMPORTRANGE/i.test(a1),
      bordersCleared: true,
      filterRemoved: true,
      lookPreserved: true
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function removeSheetFilterSafe_(sheet) {
  try {
    var f = sheet.getFilter();
    if (f) f.remove();
  } catch (e) {}
}

/** 入会者一覧：経堂・FIT365とも A:E を一塊で値コピー（行が増えても名前とチェックがずれない） */
function formatJoinListMirrorKeepImport_() {
  return setupJoinListLinked_();
}

function isCheckboxHeader_(title) {
  var t = String(title || '');
  if (!t) return false;
  if (/メールアドレス|email/i.test(t) && !/レクチャー|アンケート|付与/.test(t)) return false;
  return /アンケート|付与済|ポイント付与|レクチャーメール|送信済|済フラグ|アンケート済/.test(t);
}

function stylePromoHubSheet_(sheet, imported) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(hubTabColorFor_(sheet.getName()));
  sheet.setFrozenRows(1);

  var dn = dnTheme_();
  var rows = [['#', 'KIND', 'MIRROR SHEET', 'SOURCE SHEET', 'COLS', 'ROWS']];
  for (var i = 0; i < imported.length; i++) {
    rows.push([
      i + 1,
      'PROMO',
      imported[i].name,
      imported[i].source,
      imported[i].cols,
      imported[i].rows
    ]);
  }
  sheet.getRange(1, 1, rows.length, 6).setValues(rows);
  sheet.getRange(1, 1, 1, 6)
    .setBackground(dn.ink)
    .setFontColor(dn.paper)
    .setFontFamily('Noto Sans JP').setFontStyle('italic')
    .setFontSize(10)
    .setFontWeight('bold');
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 6)
      .setBackground(dn.paper)
      .setFontColor(dn.ink)
      .setFontFamily('Noto Sans JP').setFontStyle('italic')
      .setFontSize(10);
  }
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 72);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 64);
  sheet.setColumnWidth(6, 64);
  try {
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 6) sheet.deleteColumns(7, maxCols - 6);
  } catch (e) {}
}

function columnLetter_(n) {
  var s = '';
  var num = Number(n);
  while (num > 0) {
    var m = (num - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s || 'A';
}

/** 任意ブックのシート名・ヘッダーを覗く（集約設計用） */
function inspectSpreadsheetBook_(id) {
  try {
    if (!id) return { ok: false, message: 'id required' };
    var ss = SpreadsheetApp.openById(id);
    var sheets = ss.getSheets();
    var out = [];
    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var lastCol = Math.max(sh.getLastColumn(), 1);
      var lastRow = Math.max(sh.getLastRow(), 0);
      var headers = [];
      if (lastRow >= 1) {
        headers = sh.getRange(1, 1, 1, Math.min(lastCol, 40)).getDisplayValues()[0];
      }
      var sample = [];
      if (lastRow >= 2) {
        sample = sh.getRange(2, 1, 1, Math.min(lastCol, 12)).getDisplayValues()[0];
      }
      out.push({
        name: sh.getName(),
        gid: sh.getSheetId(),
        rows: lastRow,
        cols: lastCol,
        hidden: sh.isSheetHidden(),
        headers: headers,
        sampleRow2: sample
      });
    }
    return {
      ok: true,
      id: id,
      title: ss.getName(),
      url: ss.getUrl(),
      sheetCount: out.length,
      sheets: out
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err), id: id };
  }
}

/**
 * 制限付き共有でも診断できるように、Workspace 内の IMPORTRANGE 切れと
 * 既知の元スプシへの openById 可否をまとめて返す。
 */
function diagnoseImportsHealth_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var myEmail = '';
    try {
      myEmail = Session.getEffectiveUser().getEmail();
    } catch (eMail) {
      myEmail = '';
    }

    var targets = [
      '経堂マスタ',
      '経営マスタ',
      '【経堂】会員動向',
      '見学体験申請',
      '口コミ_経堂',
      'マシンレクチャー申込',
      '入会者一覧＋自動メール管理',
      '販促_乗り換え',
      '販促_紹介・ペア入会',
      '学割',
      '未納管理',
      '未納管理_推移',
      '経堂_入会',
      '経堂_退会',
      '経堂_OP'
    ];
    var sheetReports = [];
    for (var t = 0; t < targets.length; t++) {
      var sh = ss.getSheetByName(targets[t]);
      if (!sh) {
        sheetReports.push({ name: targets[t], exists: false });
        continue;
      }
      sheetReports.push(inspectSheetImportHealth_(sh));
    }

    var sources = [
      { key: 'workspace', id: WS_CONFIG.SPREADSHEET_ID, label: 'RYUTA Workspace' },
      { key: 'reception', id: '14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w', label: '経堂　受付状況表' },
      { key: 'kengaku', id: '1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY', label: '経堂　見学・体験フォーム' },
      { key: 'promo', id: PROMO_SOURCE_ID_, label: 'JOYFIT24経堂追加販促' },
      { key: 'review', id: REVIEW_SOURCE_ID_, label: 'EAST口コミ回答者' },
      { key: 'machine', id: MACHINE_SOURCE_ID_, label: 'マシンレクチャー・自動メール' },
      { key: 'school', id: SCHOOL_FORM_SOURCE_ID_, label: '学校関係者割フォーム' },
      { key: 'fit365', id: '1BbExBUCfyq1cfNqw4TvlwUriL-AfvghU9XT6McdzGTQ', label: 'FIT365 入会者一覧' }
    ];
    var sourceReports = [];
    for (var s = 0; s < sources.length; s++) {
      sourceReports.push(probeSourceAccess_(sources[s]));
    }

    var brokenSheets = [];
    for (var b = 0; b < sheetReports.length; b++) {
      var r = sheetReports[b];
      if (r.exists === false) continue;
      if (r.refErrorCount > 0 || r.importErrorCount > 0) brokenSheets.push(r.name);
    }
    var blockedSources = [];
    for (var c = 0; c < sourceReports.length; c++) {
      if (!sourceReports[c].ok) blockedSources.push(sourceReports[c].label || sourceReports[c].id);
    }

    return {
      ok: true,
      executedAs: myEmail,
      workspaceId: ss.getId(),
      workspaceUrl: ss.getUrl(),
      summary: {
        healthy: brokenSheets.length === 0 && blockedSources.length === 0,
        brokenSheets: brokenSheets,
        blockedSources: blockedSources
      },
      sheets: sheetReports,
      sources: sourceReports,
      tip:
        '制限付きのままでよい。blockedSources / brokenSheets に出た元スプシへ ' +
        (myEmail || 'GAS実行アカウント') +
        ' を閲覧者以上で共有し、Workspace で「アクセスを許可」を押す。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function inspectSheetImportHealth_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var maxRows = Math.min(lastRow, 80);
  var maxCols = Math.min(lastCol, 30);
  var formulas = sheet.getRange(1, 1, maxRows, maxCols).getFormulas();
  var displays = sheet.getRange(1, 1, maxRows, maxCols).getDisplayValues();

  var importFormulas = [];
  var refErrors = [];
  var importErrors = [];
  var sampleValues = [];

  for (var r = 0; r < formulas.length; r++) {
    for (var c = 0; c < formulas[r].length; c++) {
      var f = String(formulas[r][c] || '');
      var d = String(displays[r][c] || '');
      var a1 = columnLetter_(c + 1) + String(r + 1);
      if (/IMPORTRANGE/i.test(f)) {
        if (importFormulas.length < 8) {
          importFormulas.push({ cell: a1, formula: f.slice(0, 220), display: d.slice(0, 80) });
        }
      }
      if (/#REF!/i.test(d)) {
        if (refErrors.length < 12) refErrors.push({ cell: a1, display: d, formula: f.slice(0, 180) });
      }
      if (/#N\/A|#ERROR!|読み込めません|You don't have permission|アクセス/i.test(d)) {
        if (importErrors.length < 12) importErrors.push({ cell: a1, display: d, formula: f.slice(0, 180) });
      }
    }
  }

  for (var sr = 0; sr < Math.min(displays.length, 5); sr++) {
    sampleValues.push(displays[sr].slice(0, 8));
  }

  return {
    name: sheet.getName(),
    exists: true,
    hidden: sheet.isSheetHidden(),
    rows: lastRow,
    cols: lastCol,
    importrangeCount: importFormulas.length,
    refErrorCount: refErrors.length,
    importErrorCount: importErrors.length,
    importrangeSamples: importFormulas,
    refErrors: refErrors,
    importErrors: importErrors,
    sampleValues: sampleValues
  };
}

function probeSourceAccess_(src) {
  try {
    var book = SpreadsheetApp.openById(src.id);
    return {
      ok: true,
      key: src.key,
      id: src.id,
      label: src.label,
      title: book.getName(),
      url: book.getUrl(),
      sheetCount: book.getSheets().length
    };
  } catch (err) {
    return {
      ok: false,
      key: src.key,
      id: src.id,
      label: src.label,
      message: String(err && err.message ? err.message : err)
    };
  }
}

/** 指定シートの値・数式を深掘り（マスタ診断用） */
function inspectNamedSheetDeep_(name, maxRows, maxCols) {
  try {
    name = String(name || '').trim();
    if (!name) return { ok: false, message: 'name required' };
    var ss = openWorkspaceSpreadsheet_();
    var sheet = ss.getSheetByName(name);
    if (!sheet) return { ok: false, message: 'sheet not found: ' + name };
    maxRows = Math.max(1, Math.min(Number(maxRows) || 40, 120));
    maxCols = Math.max(1, Math.min(Number(maxCols) || 20, 40));
    var lastRow = Math.max(sheet.getLastRow(), 1);
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var rows = Math.min(lastRow, maxRows);
    var cols = Math.min(lastCol, maxCols);
    var formulas = sheet.getRange(1, 1, rows, cols).getFormulas();
    var displays = sheet.getRange(1, 1, rows, cols).getDisplayValues();
    var values = sheet.getRange(1, 1, rows, cols).getValues();

    var hits = [];
    var waiting = [];
    var refs = [];
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        var d = String(displays[r][c] == null ? '' : displays[r][c]);
        var f = String(formulas[r][c] || '');
        var a1 = columnLetter_(c + 1) + String(r + 1);
        if (/読み込み待ち|#REF!|#N\/A|#ERROR!|読み込めません/i.test(d)) {
          waiting.push({ cell: a1, display: d, formula: f.slice(0, 240) });
        }
        if (/#REF!/i.test(d)) refs.push({ cell: a1, display: d, formula: f.slice(0, 240) });
        if (/会員動向|IMPORTRANGE|入会計画|入会実績|数値/i.test(f + d)) {
          if (hits.length < 40) {
            hits.push({
              cell: a1,
              display: d.slice(0, 80),
              formula: f.slice(0, 220),
              valueType: typeof values[r][c]
            });
          }
        }
      }
    }

    // 会員動向シートがあれば、ヘッダー行と「入会計画」行も返す
    var trend = null;
    var trendSheet = ss.getSheetByName('【経堂】会員動向');
    if (trendSheet) {
      var tr = Math.min(Math.max(trendSheet.getLastRow(), 1), 30);
      var tc = Math.min(Math.max(trendSheet.getLastColumn(), 1), 16);
      trend = {
        rows: trendSheet.getLastRow(),
        cols: trendSheet.getLastColumn(),
        top: trendSheet.getRange(1, 1, tr, tc).getDisplayValues(),
        formulasTopLeft: trendSheet.getRange(1, 1, Math.min(tr, 5), Math.min(tc, 8)).getFormulas()
      };
    }

    return {
      ok: true,
      sheet: name,
      rows: lastRow,
      cols: lastCol,
      inspected: { rows: rows, cols: cols },
      waitingOrErrors: waiting,
      refErrors: refs,
      relatedHits: hits,
      memberTrend: trend
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

﻿/**
 * 制限付き共有でもマスタ数値が生きるよう、受付状況表の「日報」を
 * Workspace 内「経堂_受付ライブ」へ値コピーし、経堂マスタの参照を差し替える。
 * あわせて入会者一覧の TRUE/FALSE をチェックボックス化する。
 */
var RECEPTION_SOURCE_ID_ = '14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w';
var RECEPTION_LIVE_SHEET_ = '経堂_受付ライブ';
var JOIN_LIST_SHEET_ = '入会者一覧＋自動メール管理';

function repairRestrictedImports_() {
  try {
    var out = {
      receptionLive: syncReceptionLiveSheet_(),
      masterPatch: patchMasterReceptionFormulas_(),
      masterActuals: fixMasterActualFormulas_(),
      joinMirrors: syncReceptionJoinMirrors_(),
      joinList: formatJoinListMirrorKeepImport_(),
      kengakuProbe: probeSourceAccess_({
        key: 'kengaku',
        id: '1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY',
        label: '見学・体験フォーム'
      })
    };
    ensureReceptionLiveTrigger_();
    return {
      ok: true,
      executedAs: (function () {
        try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
      })(),
      results: out,
      tip:
        '経堂マスタの当月実績は「経堂_受付ライブ」（GAS同期）を参照します。' +
        '見学フォームの #REF! が残る場合は AB6 でアクセス許可を1回押してください。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/** 受付状況表の「日報」シートを値コピー（IMPORTRANGE 不要） */
function syncReceptionLiveSheet_() {
  var srcBook = SpreadsheetApp.openById(RECEPTION_SOURCE_ID_);
  var src = srcBook.getSheetByName('日報');
  if (!src) return { ok: false, message: '受付状況表に「日報」シートがありません' };

  var dest = openWorkspaceSpreadsheet_();
  var live = dest.getSheetByName(RECEPTION_LIVE_SHEET_);
  if (!live) live = dest.insertSheet(RECEPTION_LIVE_SHEET_);
  live.clear();

  var lastRow = Math.max(src.getLastRow(), 40);
  var lastCol = Math.max(src.getLastColumn(), 10);
  var values = src.getRange(1, 1, lastRow, lastCol).getValues();
  live.getRange(1, 1, lastRow, lastCol).setValues(values);
  try { live.hideSheet(); } catch (eHide) {}

  return {
    ok: true,
    source: srcBook.getName() + ' / 日報',
    rows: lastRow,
    cols: lastCol,
    sample: {
      C9: values[8][2],
      C11: values[10][2],
      C13: values[12][2],
      C15: values[14][2],
      D14: values[13][3]
    }
  };
}

/**
 * 経堂マスタ内の IMPORTRANGE($AB$1,"日報!…") を
 * 『経堂_受付ライブ』参照へ置換。
 */
function patchMasterReceptionFormulas_() {
  var dest = openWorkspaceSpreadsheet_();
  var sh = dest.getSheetByName('経堂マスタ');
  if (!sh) return { ok: false, message: '経堂マスタなし' };

  var lastRow = Math.max(sh.getLastRow(), 1);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var range = sh.getRange(1, 1, lastRow, lastCol);
  var formulas = range.getFormulas();
  var changed = [];

  for (var r = 0; r < formulas.length; r++) {
    for (var c = 0; c < formulas[r].length; c++) {
      var f = String(formulas[r][c] || '');
      if (!f) continue;
      var next = f;

      next = next.replace(
        /IMPORTRANGE\(\s*\$AB\$1\s*,\s*"日報!([^"]+)"\s*\)/gi,
        "'" + RECEPTION_LIVE_SHEET_ + "'!$1"
      );
      next = next.replace(
        /IMPORTRANGE\(\s*\$AB\$1\s*,\s*'日報!([^']+)'\s*\)/gi,
        "'" + RECEPTION_LIVE_SHEET_ + "'!$1"
      );

      if (/TEXT\(\$AB\$5,"yymm"\)=TEXT\(TODAY\(\),"yymm"\)/i.test(next) &&
          /経堂_受付ライブ/.test(next)) {
        // フォールバックは fixMasterActualFormulas_ で安全に書く
      }

      if (next !== f) {
        formulas[r][c] = next;
        changed.push({
          cell: columnLetter_(c + 1) + String(r + 1),
          from: f.slice(0, 160),
          to: next.slice(0, 160)
        });
      }
    }
  }

  if (changed.length) range.setFormulas(formulas);

  try {
    var ab4 = String(sh.getRange('AB4').getFormula() || '');
    if (/IMPORTRANGE/i.test(ab4) && /日報/i.test(ab4)) {
      sh.getRange('AB4').setFormula("='" + RECEPTION_LIVE_SHEET_ + "'!B1");
      changed.push({ cell: 'AB4', to: "='" + RECEPTION_LIVE_SHEET_ + "'!B1" });
    }
  } catch (eAb) {}

  // F8 の案内文を更新（固定文字）
  try {
    var f8 = String(sh.getRange('F8').getDisplayValue() || '');
    if (/読み込み待ち/.test(f8)) {
      sh.getRange('F8').setValue('');
      changed.push({ cell: 'F8', to: '(cleared waiting label)' });
    }
  } catch (eF8) {}

  return { ok: true, changed: changed.length, samples: changed.slice(0, 12) };
}

function patchCurrentMonthFallback_(formula) {
  var f = String(formula || '');
  if (!/IFERROR\(\s*N\([^)]+\)\s*,\s*\)/.test(f)) return f;
  var idx = f.match(/IFERROR\(\s*(INDEX\(\s*'【経堂】会員動向'[\s\S]*?\))\s*,\s*\)/i);
  if (!idx) return f;
  // replace の第2引数では $1/$2 が特殊なので、関数置換を使う
  var indexExpr = idx[1];
  return f.replace(/IFERROR\(\s*N\(([^)]+)\)\s*,\s*\)/, function (_m, nExpr) {
    return 'IFERROR(N(' + nExpr + '),' + indexExpr + ')';
  });
}

/** 壊れた F18/F20 を正しい当月ライブ参照へ直す */
function fixMasterActualFormulas_() {
  var dest = openWorkspaceSpreadsheet_();
  var sh = dest.getSheetByName('経堂マスタ');
  if (!sh) return { ok: false, message: '経堂マスタなし' };

  // 基準月 AB5 と受付ID AB1 を修復
  try {
    if (!String(sh.getRange('AB1').getDisplayValue() || '').trim()) {
      sh.getRange('AB1').setValue(RECEPTION_SOURCE_ID_);
    }
  } catch (eAb1) {}

  try {
    var b2 = String(sh.getRange('B2').getDisplayValue() || '');
    var a3 = String(sh.getRange('A3').getDisplayValue() || '');
    var ym = b2 || a3;
    var m = ym.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
    if (m) {
      sh.getRange('AB5').setFormula('=DATE(' + m[1] + ',' + m[2] + ',1)');
    } else {
      // 今日の月初へフォールバック
      sh.getRange('AB5').setFormula('=DATE(YEAR(TODAY()),MONTH(TODAY()),1)');
    }
  } catch (eAb5) {}

  // 当月実績はライブ優先。失敗時のみ会員動向へ。
  var f18 =
    '=IFERROR(N(\'' + RECEPTION_LIVE_SHEET_ + '\'!C13),' +
    'IFERROR(INDEX(\'【経堂】会員動向\'!$C$2:$N$115,MATCH("入 会  | 実績/見込",\'【経堂】会員動向\'!$B$2:$B$115,0),MATCH(MONTH(EDATE($AB$5,0)),\'【経堂】会員動向\'!$C$1:$N$1,0)),))';

  var f20 =
    '=IFERROR(N(\'' + RECEPTION_LIVE_SHEET_ + '\'!C15),' +
    'IFERROR(INDEX(\'【経堂】会員動向\'!$C$2:$N$115,MATCH("解 除  | 実績/見込",\'【経堂】会員動向\'!$B$2:$B$115,0),MATCH(MONTH(EDATE($AB$5,0)),\'【経堂】会員動向\'!$C$1:$N$1,0)),))';

  sh.getRange('F18').setFormula(f18);
  sh.getRange('F20').setFormula(f20);

  // G17（着地見込）が空なら、当月実績 F18 を見る簡易式を確認
  try {
    var g17f = String(sh.getRange('G17').getFormula() || '');
    var g17v = String(sh.getRange('G17').getDisplayValue() || '');
    if (!g17f && !g17v) {
      // 既存レイアウトに合わせ、着地見込=実績（F18）を入れる
      sh.getRange('G17').setFormula('=IFERROR(F18,)');
    }
    var g19f = String(sh.getRange('G19').getFormula() || '');
    var g19v = String(sh.getRange('G19').getDisplayValue() || '');
    if (!g19f && !g19v) {
      sh.getRange('G19').setFormula('=IFERROR(F20,)');
    }
  } catch (eG) {}

  return {
    ok: true,
    F18: sh.getRange('F18').getDisplayValue(),
    F20: sh.getRange('F20').getDisplayValue(),
    liveC13: (function () {
      var live = dest.getSheetByName(RECEPTION_LIVE_SHEET_);
      return live ? live.getRange('C13').getDisplayValue() : null;
    })()
  };
}

function peekMasterKeyCells_() {
  var dest = openWorkspaceSpreadsheet_();
  var sh = dest.getSheetByName('経堂マスタ');
  if (!sh) return { ok: false, message: '経堂マスタなし' };
  var live = dest.getSheetByName(RECEPTION_LIVE_SHEET_);
  var trend = dest.getSheetByName('【経堂】会員動向');
  var cells = ['A3', 'B2', 'AB1', 'AB4', 'AB5', 'AB6', 'B5', 'D5', 'F5', 'H5', 'J5', 'F8', 'F17', 'F18', 'F19', 'F20', 'G17', 'G19', 'H18', 'H20'];
  var out = {};
  for (var i = 0; i < cells.length; i++) {
    var a1 = cells[i];
    out[a1] = {
      display: sh.getRange(a1).getDisplayValue(),
      formula: String(sh.getRange(a1).getFormula() || '')
    };
  }
  return {
    ok: true,
    cells: out,
    live: live ? {
      C9: live.getRange('C9').getDisplayValue(),
      C11: live.getRange('C11').getDisplayValue(),
      C13: live.getRange('C13').getDisplayValue(),
      C15: live.getRange('C15').getDisplayValue(),
      D14: live.getRange('D14').getDisplayValue()
    } : null,
    trendA1: trend ? String(trend.getRange('A1').getFormula() || trend.getRange('A1').getDisplayValue()) : null,
    trendB5: trend ? trend.getRange('B5').getDisplayValue() : null,
    trendRowLabels: trend ? trend.getRange('B1:B12').getDisplayValues() : null
  };
}

function syncReceptionJoinMirrors_() {
  var srcBook = SpreadsheetApp.openById(RECEPTION_SOURCE_ID_);
  var dest = openWorkspaceSpreadsheet_();
  var specs = [
    { dest: '経堂_入会', source: '入会・退会_データ', range: 'A1:E' },
    { dest: '経堂_退会', source: '入会・退会_データ', range: 'G1:L' },
    { dest: '経堂_OP', source: 'OP集計', range: 'A1:E18' }
  ];
  var out = [];
  for (var i = 0; i < specs.length; i++) {
    var sp = specs[i];
    var src = srcBook.getSheetByName(sp.source);
    if (!src) {
      out.push({ name: sp.dest, ok: false, message: 'missing source ' + sp.source });
      continue;
    }
    var parsed = parseA1Range_(sp.range);
    var values = src.getRange(parsed.r1, parsed.c1, parsed.numRows, parsed.numCols).getValues();
    var sh = dest.getSheetByName(sp.dest);
    if (!sh) sh = dest.insertSheet(sp.dest);
    if (/IMPORTRANGE/i.test(String(sh.getRange(1, 1).getFormula() || ''))) {
      out.push({ name: sp.dest, ok: true, skipped: 'IMPORTRANGE のため上書きしない' });
      continue;
    }
    sh.clear();
    sh.getRange(1, 1, values.length, values[0].length).setValues(values);
    try { sh.hideSheet(); } catch (eH) {}
    out.push({ name: sp.dest, ok: true, rows: values.length, cols: values[0].length });
  }
  return { ok: true, sheets: out };
}

function parseA1Range_(a1) {
  var m = String(a1 || '').match(/^([A-Z]+)(\d+):([A-Z]+)(\d*)$/i);
  if (!m) return { r1: 1, c1: 1, numRows: 50, numCols: 5 };
  var c1 = letterToColumn_(m[1]);
  var r1 = Number(m[2]);
  var c2 = letterToColumn_(m[3]);
  var r2 = m[4] ? Number(m[4]) : Math.max(r1, 500);
  return {
    r1: r1,
    c1: c1,
    numRows: Math.max(1, r2 - r1 + 1),
    numCols: Math.max(1, c2 - c1 + 1)
  };
}

function letterToColumn_(letters) {
  var s = String(letters || '').toUpperCase();
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

/**
 * 入会者一覧を元スプシから値同期し、チェック列をチェックボックス化。
 * IMPORTRANGE だと入力規則が効かない／切れるため値同期にする。
 */
function syncJoinListWithCheckboxes_() {
  var srcBook = SpreadsheetApp.openById(MACHINE_SOURCE_ID_);
  var src = srcBook.getSheetByName(JOIN_LIST_SHEET_);
  if (!src) return { ok: false, message: 'source missing: ' + JOIN_LIST_SHEET_ };

  var lastRow = Math.max(src.getLastRow(), 1);
  var lastCol = Math.max(src.getLastColumn(), 1);
  var values = src.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];

  var dest = openWorkspaceSpreadsheet_();
  var sh = dest.getSheetByName(JOIN_LIST_SHEET_);
  if (!sh) sh = dest.insertSheet(JOIN_LIST_SHEET_);
  sh.clear();
  try { sh.clearConditionalFormatRules(); } catch (e0) {}
  try { sh.getDataRange().clearDataValidations(); } catch (e1) {}
  sh.getRange(1, 1, values.length, values[0].length).setValues(values);

  var dnJoin = dnTheme_();
  sh.getRange(1, 1, 1, lastCol)
    .setBackground(dnJoin.ink)
    .setFontColor(dnJoin.paper)
    .setFontFamily('Noto Sans JP').setFontStyle('italic')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(false, false, true, false, false, false, dnJoin.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  if (values.length >= 2) {
    sh.getRange(2, 1, values.length - 1, lastCol)
      .setBackground(dnJoin.paper)
      .setFontColor(dnJoin.ink)
      .setFontFamily('Noto Sans JP').setFontStyle('italic')
      .setFontSize(10)
      .setHorizontalAlignment('center');
  }

  var checkCols = [];
  for (var h = 0; h < headers.length; h++) {
    if (isCheckboxHeader_(headers[h])) checkCols.push(h + 1);
  }
  if (!checkCols.length) {
    if (lastCol >= 4) checkCols.push(4);
    if (lastCol >= 5) checkCols.push(5);
  }

  var applied = [];
  if (values.length >= 2 && checkCols.length) {
    for (var i = 0; i < checkCols.length; i++) {
      var col = checkCols[i];
      var body = [];
      for (var r = 1; r < values.length; r++) {
        var raw = values[r][col - 1];
        if (raw === true || raw === false) body.push([raw]);
        else {
          var s = String(raw == null ? '' : raw).trim().toUpperCase();
          if (s === 'TRUE') body.push([true]);
          else if (s === 'FALSE' || s === '') body.push([false]);
          else body.push([false]);
        }
      }
      var range = sh.getRange(2, col, body.length, 1);
      range.setValues(body);
      range.setDataValidation(
        SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build()
      );
      applied.push({ col: col, header: String(headers[col - 1] || ''), rows: body.length });
    }
  }

  return {
    ok: true,
    rows: values.length,
    cols: lastCol,
    checkCols: applied,
    sourceUrl: srcBook.getUrl()
  };
}

function ensureReceptionLiveTrigger_() {
  try {
    var exists = false;
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'syncReceptionLiveTriggered_') {
        exists = true;
        break;
      }
    }
    if (!exists) {
      ScriptApp.newTrigger('syncReceptionLiveTriggered_')
        .timeBased()
        .everyHours(1)
        .create();
    }
  } catch (err) {}
}

function syncReceptionLiveTriggered_() {
  try { syncReceptionLiveSheet_(); } catch (e1) {}
  try { syncReceptionJoinMirrors_(); } catch (e2) {}
  try { joinListRefresh_(); } catch (e3) {}
  try { unpaidFillTrendAll_(); } catch (e4) {}
}



/** EAST口コミ → Workspace（読み取り専用 / IMPORTRANGE+QUERY） */
var REVIEW_SOURCE_ID_ = '13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM';
var REVIEW_JOYFIT_SHEET_ = '回答シート_JOYFIT';
var REVIEW_IMPORT_COLS_ = 23; // A:W（ポイント付与済・付与日時まで）
var REVIEW_GRANT_APP_URL_ =
  'https://script.google.com/a/macros/okamoto-group.co.jp/s/AKfycbwu1eUxJzePa494p-343axfwgUcnHATf-db7FKw806rXZQsHn_ea0uHc6415yw-RZ80/exec';

/**
 * EAST側は一切変更しない。
 * Workspace に QUERY(IMPORTRANGE(...)) を置き、経堂(storeId=kyodo)だけライブ同期。
 * A1 からデータ（timestamp が1列目）。付与アプリURLは URL一覧の先頭へ。
 */
function setupReviewImport_() {
  try {
    var dest = openWorkspaceSpreadsheet_();
    var destName = '口コミ_経堂';
    var mirror = dest.getSheetByName(destName);
    if (!mirror) {
      mirror = dest.insertSheet(destName, 0);
    } else {
      mirror.clear();
      try { mirror.clearConditionalFormatRules(); } catch (e0) {}
      try { mirror.getDataRange().clearDataValidations(); } catch (e1) {}
    }

    var cols = REVIEW_IMPORT_COLS_;
    var endCol = columnLetter_(cols);
    var maxCols = mirror.getMaxColumns();
    if (maxCols < cols) mirror.insertColumnsAfter(maxCols, cols - maxCols);

    // A1 からライブ取得（timestamp = 1列目）
    var formula =
      '=QUERY(IMPORTRANGE("' +
      REVIEW_SOURCE_ID_ +
      '","' +
      REVIEW_JOYFIT_SHEET_.replace(/"/g, '""') +
      '!A:' +
      endCol +
      '"),"select * where Col2 = \'kyodo\'",1)';

    mirror.setHiddenGridlines(true);
    mirror.setTabColor(hubTabColorFor_(destName));
    mirror.setFrozenRows(1);
    mirror.getRange(1, 1).setFormula(formula);

    for (var c = 1; c <= cols; c++) mirror.setColumnWidth(c, 110);
    mirror.setColumnWidth(1, 150);
    mirror.setColumnWidth(3, 140);
    mirror.setColumnWidth(5, 100);
    mirror.setColumnWidth(9, 200);
    mirror.setColumnWidth(12, 220);
    mirror.setColumnWidth(13, 160);
    mirror.setColumnWidth(14, 200);
    mirror.setColumnWidth(15, 260);
    mirror.setColumnWidth(16, 200);
    mirror.setColumnWidth(22, 110);
    mirror.setColumnWidth(23, 130);

    try {
      var dnRev = dnTheme_();
      mirror.getRange(1, 1, 1, cols)
        .setBackground(dnRev.ink)
        .setFontColor(dnRev.paper)
        .setFontFamily('Noto Sans JP').setFontStyle('italic')
        .setFontSize(10)
        .setFontWeight('bold');
      mirror.setRowHeight(1, 32);
      var lastBody = Math.min(mirror.getMaxRows(), 500);
      if (lastBody >= 2) {
        mirror.getRange(2, 1, lastBody - 1, cols)
          .setBackground(dnRev.paper)
          .setFontColor(dnRev.ink)
          .setFontFamily('Noto Sans JP').setFontStyle('italic')
          .setFontSize(10)
          .setVerticalAlignment('middle');
        mirror.getRange('A2:A' + lastBody).setNumberFormat('yyyy/mm/dd HH:mm');
        mirror.getRange('W2:W' + lastBody).setNumberFormat('yyyy/mm/dd HH:mm');
        var rules = [];
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied('=$V2=TRUE')
            .setBackground(dnRev.cream)
            .setRanges([mirror.getRange('V2:V' + lastBody)])
            .build()
        );
        mirror.setConditionalFormatRules(rules);
      }
    } catch (eStyle) {}

    // 付与アプリURLは URL一覧の一番上（データ1行目）へ
    upsertReviewGrantUrlIndexTop_(dest);

    var leftovers = dest.getSheets();
    for (var j = leftovers.length - 1; j >= 0; j--) {
      var nm = leftovers[j].getName();
      if (/^シート\d+$/.test(nm) && leftovers[j].getLastRow() === 0) {
        try {
          if (dest.getSheets().length > 1) dest.deleteSheet(leftovers[j]);
        } catch (eDel) {}
      }
    }

    removeReviewSyncTriggers_();

    return {
      ok: true,
      mode: 'importrange-query',
      sourceId: REVIEW_SOURCE_ID_,
      sourceSheet: REVIEW_JOYFIT_SHEET_,
      filter: "Col2 = 'kyodo'",
      importRange: 'A:' + endCol,
      grantAppUrl: REVIEW_GRANT_APP_URL_,
      destSheet: destName,
      formula: formula,
      workspaceUrl: dest.getUrl(),
      note: '口コミ_経堂は A1=timestamp。付与アプリURLは URL一覧の先頭。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/** URL一覧の先頭行（2行目）に口コミ付与アプリを置く */
function upsertReviewGrantUrlIndexTop_(ss) {
  try {
    var sh = ss.getSheetByName('URL一覧');
    if (!sh) return;

    var last = Math.max(sh.getLastRow(), 1);
    // 既存の同一URL行を削除
    if (last >= 2) {
      var urls = sh.getRange(2, 4, last - 1, 1).getDisplayValues();
      for (var i = urls.length - 1; i >= 0; i--) {
        if (String(urls[i][0] || '').indexOf('AKfycbwu1eUxJzePa494p-343axfwgUcnHATf-db7FKw806rXZQsHn_ea0uHc6415yw-RZ80') !== -1) {
          sh.deleteRow(i + 2);
        }
      }
    }

    // ヘッダーが無ければ作る
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, 4).setValues([['#', 'KIND', 'TITLE', 'URL']]);
      var dnUrlHead = dnTheme_();
      sh.getRange(1, 1, 1, 4)
        .setBackground(dnUrlHead.ink)
        .setFontColor(dnUrlHead.paper)
        .setFontFamily('Noto Sans JP').setFontStyle('italic')
        .setFontSize(10)
        .setFontWeight('bold');
    }

    // 2行目に挿入して一番上へ
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 4).setValues([[
      1,
      'REVIEW',
      '口コミ付与アプリ（EAST）',
      REVIEW_GRANT_APP_URL_
    ]]);
    var dnUrlRow = dnTheme_();
    sh.getRange(2, 1, 1, 4)
      .setBackground(dnUrlRow.paper)
      .setFontColor(dnUrlRow.ink)
      .setFontFamily('Noto Sans JP').setFontStyle('italic')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sh.getRange(2, 4).setFontColor(dnUrlRow.blood);
    try {
      sh.getRange(2, 4).setFormula(
        '=HYPERLINK("' + REVIEW_GRANT_APP_URL_ + '","' + REVIEW_GRANT_APP_URL_ + '")'
      );
    } catch (eLink) {}

    // # を振り直す
    var endRow = sh.getLastRow();
    if (endRow >= 2) {
      var nums = [];
      for (var n = 1; n <= endRow - 1; n++) nums.push([n]);
      sh.getRange(2, 1, nums.length, 1).setValues(nums);
    }
  } catch (err) {}
}

/** 互換API */
function syncReviewKyodo_() {
  return setupReviewImport_();
}

function joinListToBool_(v) {
  if (v === true) return true;
  if (v === false) return false;
  var s = String(v == null ? '' : v).trim().toUpperCase();
  return s === 'TRUE' || s === '☑' || s === '1' || s === 'CHECKED';
}

var FIT365_JOIN_SOURCE_ID_ = '1BbExBUCfyq1cfNqw4TvlwUriL-AfvghU9XT6McdzGTQ';
var JOIN_LIST_BODY_START_ = 3;
var JOIN_LIST_BLOCK_COLS_ = 5;

function joinListOnEditSimple_(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (!sh || sh.getName() !== JOIN_LIST_SHEET_) return;
  if (e.range.getRow() < JOIN_LIST_BODY_START_) return;
  var lastCol = e.range.getColumn() + e.range.getNumColumns() - 1;
  if (e.range.getColumn() > 10 || lastCol < 1) return;
  joinListQueueRange_(sh, e);
}

function joinListOldValueAt_(e) {
  if (!e) return '';
  if (e.range && e.range.getNumRows() === 1 && e.range.getNumColumns() === 1) {
    return String(e.oldValue == null ? '' : e.oldValue);
  }
  return '';
}

function joinListQueueRange_(sh, e) {
  var range = e.range;
  var q = unpaidQueueSheet_(sh.getParent());
  var last = Math.max(q.getLastRow(), 1);
  var rows = [];
  var r0 = range.getRow();
  var nR = range.getNumRows();
  var c0 = range.getColumn();
  var lastCol = c0 + range.getNumColumns() - 1;
  var seen = {};
  var r;
  for (r = 0; r < nR; r++) {
    var destRow = r0 + r;
    if (destRow < JOIN_LIST_BODY_START_) continue;
    var blocks = [];
    if (c0 <= 5) blocks.push({ start: 1, key: 'kyodo', emailCol: 3 });
    if (lastCol >= 6 && c0 <= 10) blocks.push({ start: 6, key: 'fit365', emailCol: 8 });
    var b;
    for (b = 0; b < blocks.length; b++) {
      var block = blocks[b];
      var id = destRow + ':' + block.start;
      if (seen[id]) continue;
      seen[id] = true;
      var oldEmail = '';
      if (c0 <= block.emailCol && lastCol >= block.emailCol) oldEmail = joinListOldValueAt_(e);
      rows.push([
        new Date(),
        'join',
        block.key,
        destRow,
        block.start,
        JSON.stringify({ oldEmail: oldEmail }),
        'queued'
      ]);
    }
  }
  if (rows.length) q.getRange(last + 1, 1, rows.length, 7).setValues(rows);
}

function joinListCheckboxRule_() {
  return SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
}

function joinListPaintChrome_(sh) {
  if (!sh) return;
  var t = dnTheme_();
  var kyodoUrl = 'https://docs.google.com/spreadsheets/d/' + MACHINE_SOURCE_ID_ + '/edit';
  var fitUrl = 'https://docs.google.com/spreadsheets/d/' + FIT365_JOIN_SOURCE_ID_ + '/edit';
  hubPaintOpenSourceCell_(sh, 1, 1, 3, { url: kyodoUrl, label: '元のシートを開く ↗' });
  hubPaintOpenSourceCell_(sh, 1, 6, 3, { url: fitUrl, label: 'FIT365 のシートを開く ↗' });
  try { sh.getRange('D1:E1').breakApart(); } catch (e1) {}
  try { sh.getRange('I1:J1').breakApart(); } catch (e2) {}
  sh.getRange('D1:E1').merge().setValue('変更は自動で元シートへ')
    .setFontSize(8).setFontColor(t.ash).setFontWeight('normal')
    .setBackground(t.paper).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setNote('チェック・名前・メール・入会日を変えると、未納のゲートストップと同じく元のスプシへ自動で送ります。メニューの「チェックを元へ反映」は予備です。');
  sh.getRange('I1:J1').merge().setValue('変更は自動で元シートへ')
    .setFontSize(8).setFontColor(t.ash).setFontWeight('normal')
    .setBackground(t.paper).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setNote('チェック・名前・メール・入会日を変えると、FIT365 の元スプシへ自動で送ります。メニューの「チェックを元へ反映」は予備です。');
}

/**
 * 経堂 A:E / FIT365 F:J を元スプシから一塊で値コピーする。
 * IMPORTRANGE の名前 + 別置きチェックだと、元に新規行が入った瞬間に位置がずれる。
 */
function setupJoinListLinked_() {
  try {
    var dest = openWorkspaceSpreadsheet_();
    var sh = dest.getSheetByName(JOIN_LIST_SHEET_);
    if (!sh) sh = dest.insertSheet(JOIN_LIST_SHEET_);
    if (sh.getMaxColumns() < 10) sh.insertColumnsAfter(sh.getMaxColumns(), 10 - sh.getMaxColumns());
    try { sh.getRange(1, 1, 1, 10).breakApart(); } catch (eB) {}
    try { ensureUnpaidEditTrigger_(); } catch (eT) {}
    var pulled = joinListRefresh_(sh);
    try { removeSheetFilterSafe_(sh); } catch (eFil) {}
    var lecture = dest.getSheetByName('マシンレクチャー申込');
    var lectureOut = lecture ? tidyImportMirrorKeepLook_(lecture) : null;
    return {
      ok: true,
      pulled: pulled,
      snap: joinListSnapshot_(sh),
      machineLecture: lectureOut
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function joinListRefresh_(sh) {
  if (!sh) {
    var ss = openWorkspaceSpreadsheet_();
    sh = ss.getSheetByName(JOIN_LIST_SHEET_);
  }
  if (!sh) return { ok: false, message: 'missing join list' };
  if (sh.getMaxColumns() < 10) sh.insertColumnsAfter(sh.getMaxColumns(), 10 - sh.getMaxColumns());
  joinListPaintChrome_(sh);
  var kyodo = joinListCopyBlock_(sh, MACHINE_SOURCE_ID_, 1);
  var fit = joinListCopyBlock_(sh, FIT365_JOIN_SOURCE_ID_, 6);
  styleJoinListLinked_(sh, kyodo.rows || 0, fit.rows || 0);
  try { removeSheetFilterSafe_(sh); } catch (eFil) {}
  return { ok: !!(kyodo.ok && fit.ok), kyodo: kyodo, fit365: fit, snap: joinListSnapshot_(sh) };
}

function joinListCopyBlock_(destSh, srcId, destStartCol) {
  var src = SpreadsheetApp.openById(srcId);
  var srcSh = src.getSheetByName(JOIN_LIST_SHEET_);
  if (!srcSh) return { ok: false, message: 'no source sheet' };
  var srcLast = Math.max(srcSh.getLastRow(), 1);
  var srcCols = Math.min(Math.max(srcSh.getLastColumn(), 1), JOIN_LIST_BLOCK_COLS_);
  var raw = srcSh.getRange(1, 1, srcLast, JOIN_LIST_BLOCK_COLS_).getValues();
  var headers = raw[0];
  while (headers.length < JOIN_LIST_BLOCK_COLS_) headers.push('');
  var body = [];
  var i;
  for (i = 1; i < raw.length; i++) {
    var row = raw[i];
    var has = false;
    var c;
    for (c = 0; c < JOIN_LIST_BLOCK_COLS_; c++) {
      if (String(row[c] == null ? '' : row[c]).trim() !== '') has = true;
    }
    if (!has) continue;
    body.push([
      row[0],
      row[1],
      row[2],
      joinListToBool_(row[3]),
      joinListToBool_(row[4])
    ]);
  }
  var need = JOIN_LIST_BODY_START_ - 1 + Math.max(body.length, 1);
  if (destSh.getMaxRows() < need) destSh.insertRowsAfter(destSh.getMaxRows(), need - destSh.getMaxRows());
  var oldLast = Math.max(destSh.getLastRow(), JOIN_LIST_BODY_START_);
  var clearFrom = 2;
  var clearN = Math.max(oldLast, need) - clearFrom + 1;
  destSh.getRange(clearFrom, destStartCol, clearN, JOIN_LIST_BLOCK_COLS_).clearContent();
  destSh.getRange(clearFrom, destStartCol, clearN, JOIN_LIST_BLOCK_COLS_).clearDataValidations();
  destSh.getRange(2, destStartCol, 1, JOIN_LIST_BLOCK_COLS_).setValues([headers.slice(0, JOIN_LIST_BLOCK_COLS_)]);
  if (body.length) {
    destSh.getRange(JOIN_LIST_BODY_START_, destStartCol, body.length, JOIN_LIST_BLOCK_COLS_).setValues(body);
    destSh.getRange(JOIN_LIST_BODY_START_, destStartCol, body.length, 1).setNumberFormat('yyyy/mm/dd');
    var checkRng = destSh.getRange(JOIN_LIST_BODY_START_, destStartCol + 3, body.length, 2);
    checkRng.setDataValidation(joinListCheckboxRule_());
    checkRng.setHorizontalAlignment('center');
  }
  return {
    ok: true,
    srcCols: srcCols,
    rows: body.length,
    true1: body.filter(function (x) { return x[3]; }).length,
    true2: body.filter(function (x) { return x[4]; }).length
  };
}

function joinListCountBlock_(sh, startCol) {
  var last = Math.max(sh.getLastRow(), 2);
  if (last < JOIN_LIST_BODY_START_) return 0;
  var n = last - 2;
  var vals = sh.getRange(JOIN_LIST_BODY_START_, startCol, n, 3).getDisplayValues();
  var count = 0;
  var i;
  for (i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() || String(vals[i][1] || '').trim() || String(vals[i][2] || '').trim()) {
      count = i + 1;
    }
  }
  return count;
}

function styleJoinListLinked_(sh, kyodoRows, fitRows) {
  if (!sh) return;
  var t = dnTheme_();
  kyodoRows = Math.max(0, Number(kyodoRows) || 0);
  fitRows = Math.max(0, Number(fitRows) || 0);
  if (!kyodoRows) kyodoRows = joinListCountBlock_(sh, 1);
  if (!fitRows) fitRows = joinListCountBlock_(sh, 6);
  var bodyRows = Math.max(kyodoRows, fitRows, 1);
  var paintRows = Math.max(bodyRows, Math.min(sh.getMaxRows() - 2, bodyRows + 40));
  try { sh.clearConditionalFormatRules(); } catch (e0) {}
  try { sh.getRange(1, 1, sh.getMaxRows(), 10).setBorder(false, false, false, false, false, false); } catch (eB) {}
  hubType_(sh.getRange(3, 1, paintRows, 10))
    .setBackground(t.paper).setFontColor(t.ink).setFontSize(10).setFontWeight('normal')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(false);
  sh.getRange(3, 2, paintRows, 1).setHorizontalAlignment('left');
  sh.getRange(3, 7, paintRows, 1).setHorizontalAlignment('left');
  sh.getRange(3, 3, paintRows, 1).setHorizontalAlignment('left');
  sh.getRange(3, 8, paintRows, 1).setHorizontalAlignment('left');
  hubType_(sh.getRange('A2:J2'))
    .setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(false, false, true, false, false, false, t.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange('D1:E1').setBackground(t.paper).setFontColor(t.ash).setFontWeight('normal').setFontSize(8).setWrap(true);
  sh.getRange('I1:J1').setBackground(t.paper).setFontColor(t.ash).setFontWeight('normal').setFontSize(8).setWrap(true);
  sh.setFrozenRows(2);
  sh.setHiddenGridlines(true);
  sh.setTabColor(t.ink);
  sh.setRowHeight(1, 28);
  sh.setRowHeight(2, 48);
  try { sh.setRowHeightsForced(3, paintRows, 26); } catch (eH) {}
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidths(4, 2, 118);
  sh.setColumnWidth(6, 150);
  sh.setColumnWidth(7, 110);
  sh.setColumnWidth(8, 220);
  sh.setColumnWidths(9, 2, 118);
  var rules = [];
  if (kyodoRows > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=D3=TRUE')
        .setBackground(t.ink)
        .setFontColor(t.paper)
        .setBold(true)
        .setRanges([sh.getRange(3, 4, kyodoRows, 2)])
        .build()
    );
  }
  if (fitRows > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=I3=TRUE')
        .setBackground(t.ink)
        .setFontColor(t.paper)
        .setBold(true)
        .setRanges([sh.getRange(3, 9, fitRows, 2)])
        .build()
    );
  }
  if (rules.length) sh.setConditionalFormatRules(rules);
}

function joinListSnapshot_(sh) {
  if (!sh) return null;
  var cell = function (a1) {
    var r = sh.getRange(a1);
    var dv = r.getDataValidation();
    return {
      v: r.getValue(),
      d: String(r.getDisplayValue() || ''),
      f: String(r.getFormula() || ''),
      check: dv ? String(dv.getCriteriaType()) : ''
    };
  };
  return {
    last: sh.getLastRow(),
    kyodoRows: joinListCountBlock_(sh, 1),
    fitRows: joinListCountBlock_(sh, 6),
    a2: cell('A2'),
    f2: cell('F2'),
    a3: cell('A3'),
    b3: cell('B3'),
    e3: cell('E3'),
    e9: cell('E9'),
    f3: cell('F3'),
    g3: cell('G3'),
    i3: cell('I3'),
    j3: cell('J3'),
    eTail: cell('E' + (joinListCountBlock_(sh, 1) + 4)),
    jTail: cell('J' + (joinListCountBlock_(sh, 6) + 4)),
    a2imp: /IMPORTRANGE/i.test(String(sh.getRange('A2').getFormula() || '')),
    f2imp: /IMPORTRANGE/i.test(String(sh.getRange('F2').getFormula() || ''))
  };
}

function joinListPullChecks_(sh) {
  sh = sh || openWorkspaceSpreadsheet_().getSheetByName(JOIN_LIST_SHEET_);
  return joinListRefresh_(sh);
}

function joinListPushChecks_(ss) {
  ss = ss || openWorkspaceSpreadsheet_();
  var sh = ss.getSheetByName(JOIN_LIST_SHEET_);
  if (!sh) return { ok: false, message: 'missing join list' };
  var kyodo = joinListPushBlock_(sh, 1, MACHINE_SOURCE_ID_);
  var fit = joinListPushBlock_(sh, 6, FIT365_JOIN_SOURCE_ID_);
  return { ok: !!(kyodo.ok && fit.ok), kyodo: kyodo, fit365: fit };
}

function joinListSrcEmailMap_(srcSh) {
  var srcLast = Math.max(srcSh.getLastRow(), 2);
  var srcEmails = srcSh.getRange(2, 3, srcLast - 1, 1).getDisplayValues();
  var map = {};
  var i;
  for (i = 0; i < srcEmails.length; i++) {
    var k = String(srcEmails[i][0] || '').trim().toLowerCase();
    if (k) map[k] = i + 2;
  }
  return map;
}

function joinListRowPayload_(row) {
  return [
    row[0],
    row[1],
    row[2],
    joinListToBool_(row[3]),
    joinListToBool_(row[4])
  ];
}

function joinListPushOneRow_(destSh, destRow, destStartCol, srcId, oldEmail) {
  if (!destSh || destRow < JOIN_LIST_BODY_START_) return { ok: true, pushed: 0 };
  var row = destSh.getRange(destRow, destStartCol, 1, 5).getValues()[0];
  var emailNow = String(row[2] == null ? '' : row[2]).trim().toLowerCase();
  var lookup = String(oldEmail || '').trim().toLowerCase() || emailNow;
  if (!lookup || lookup === 'メールアドレス') return { ok: true, pushed: 0 };
  var src = SpreadsheetApp.openById(srcId);
  var srcSh = src.getSheetByName(JOIN_LIST_SHEET_);
  if (!srcSh) return { ok: false, message: 'no source sheet' };
  var map = joinListSrcEmailMap_(srcSh);
  var srcRow = map[lookup] || (emailNow && emailNow !== lookup ? map[emailNow] : 0);
  if (!srcRow) return { ok: true, pushed: 0, missed: 1 };
  srcSh.getRange(srcRow, 1, 1, 5).setValues([joinListRowPayload_(row)]);
  return { ok: true, pushed: 1 };
}

function joinListPushEdited_(e) {
  if (!e || !e.range) return { ok: true, pushed: 0 };
  var sh = e.range.getSheet();
  var r0 = e.range.getRow();
  var nR = e.range.getNumRows();
  var c0 = e.range.getColumn();
  var lastCol = c0 + e.range.getNumColumns() - 1;
  var pushed = 0;
  var missed = 0;
  var r;
  for (r = 0; r < nR; r++) {
    var destRow = r0 + r;
    if (destRow < JOIN_LIST_BODY_START_) continue;
    if (c0 <= 5) {
      var oldK = (c0 <= 3 && lastCol >= 3) ? joinListOldValueAt_(e) : '';
      var ky = joinListPushOneRow_(sh, destRow, 1, MACHINE_SOURCE_ID_, oldK);
      pushed += (ky && ky.pushed) || 0;
      missed += (ky && ky.missed) || 0;
    }
    if (lastCol >= 6 && c0 <= 10) {
      var oldF = (c0 <= 8 && lastCol >= 8) ? joinListOldValueAt_(e) : '';
      var fit = joinListPushOneRow_(sh, destRow, 6, FIT365_JOIN_SOURCE_ID_, oldF);
      pushed += (fit && fit.pushed) || 0;
      missed += (fit && fit.missed) || 0;
    }
  }
  return { ok: true, pushed: pushed, missed: missed };
}

function joinListFlushQueue_(ss) {
  ss = ss || openWorkspaceSpreadsheet_();
  var q = ss.getSheetByName(UNPAID_QUEUE_SHEET_);
  if (!q || q.getLastRow() < 2) return { ok: true, pushed: 0 };
  var n = q.getLastRow() - 1;
  var data = q.getRange(2, 1, n, 7).getValues();
  var sh = ss.getSheetByName(JOIN_LIST_SHEET_);
  if (!sh) return { ok: false, message: 'missing join list' };
  var pushed = 0;
  var missed = 0;
  var i;
  for (i = 0; i < data.length; i++) {
    if (String(data[i][1] || '') !== 'join') continue;
    if (String(data[i][6] || '') !== 'queued') continue;
    var destRow = Number(data[i][3]);
    var startCol = Number(data[i][4]) || (String(data[i][2] || '') === 'fit365' ? 6 : 1);
    var payload = {};
    try { payload = JSON.parse(data[i][5]); } catch (eJ) { payload = {}; }
    var srcId = startCol === 6 ? FIT365_JOIN_SOURCE_ID_ : MACHINE_SOURCE_ID_;
    var r = joinListPushOneRow_(sh, destRow, startCol, srcId, payload && payload.oldEmail);
    if (r && r.pushed) {
      q.getRange(i + 2, 7).setValue('done');
      pushed += 1;
    } else if (r && r.missed) {
      q.getRange(i + 2, 7).setValue('missed');
      missed += 1;
    } else {
      q.getRange(i + 2, 7).setValue('done');
    }
  }
  return { ok: true, pushed: pushed, missed: missed };
}

function joinListPushBlock_(destSh, destStartCol, srcId) {
  var src = SpreadsheetApp.openById(srcId);
  var srcSh = src.getSheetByName(JOIN_LIST_SHEET_);
  if (!srcSh) return { ok: false, message: 'no source sheet' };
  var n = joinListCountBlock_(destSh, destStartCol);
  if (n < 1) return { ok: true, pushed: 0 };
  var block = destSh.getRange(JOIN_LIST_BODY_START_, destStartCol, n, 5).getValues();
  var map = joinListSrcEmailMap_(srcSh);
  var pushed = 0;
  var missed = 0;
  var i;
  for (i = 0; i < block.length; i++) {
    var em = String(block[i][2] == null ? '' : block[i][2]).trim().toLowerCase();
    if (!em || em === 'メールアドレス') continue;
    var srcRow = map[em];
    if (!srcRow) {
      missed += 1;
      continue;
    }
    srcSh.getRange(srcRow, 1, 1, 5).setValues([joinListRowPayload_(block[i])]);
    pushed += 1;
  }
  return { ok: true, pushed: pushed, missed: missed };
}

/** マシンレクチャー／入会者一覧 → Workspace（入会者一覧は値コピーで名前とチェックを一塊同期） */
var MACHINE_SOURCE_ID_ = '1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8';

function setupMachineImport_() {
  try {
    var source = SpreadsheetApp.openById(MACHINE_SOURCE_ID_);
    var dest = openWorkspaceSpreadsheet_();
    var sheets = source.getSheets();
    var imported = [];

    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var name = String(sh.getName());
      var lastCol = Math.max(sh.getLastColumn(), 1);
      var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      var usedCols = 0;
      for (var c = 0; c < headers.length; c++) {
        if (String(headers[c] || '').trim() !== '') usedCols = c + 1;
      }
      if (usedCols < 1) usedCols = lastCol;

      var existing = dest.getSheetByName(name);
      if (existing) dest.deleteSheet(existing);
      var mirror = dest.insertSheet(name);
      styleImportMirrorSheet_(mirror, source.getId(), name, usedCols, headers.slice(0, usedCols));
      imported.push({ name: name, cols: usedCols, rows: Math.max(sh.getLastRow(), 0) });
    }

    var leftovers = dest.getSheets();
    for (var j = leftovers.length - 1; j >= 0; j--) {
      var nm = leftovers[j].getName();
      if (/^シート\d+$/.test(nm) && leftovers[j].getLastRow() === 0) {
        try {
          if (dest.getSheets().length > 1) dest.deleteSheet(leftovers[j]);
        } catch (eDel) {}
      }
    }

    return {
      ok: true,
      sourceId: MACHINE_SOURCE_ID_,
      sourceTitle: source.getName(),
      sourceUrl: source.getUrl(),
      imported: imported,
      joinList: setupJoinListLinked_(),
      workspaceUrl: dest.getUrl(),
      note: 'シート名そのまま / 入会者一覧は経堂・FIT365とも名前とチェックを一塊で同期。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function removeReviewSyncTriggers_() {
  try {
    var handlers = { syncReviewKyodoTriggered_: 1, syncReviewKyodo_: 1 };
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      var fn = triggers[i].getHandlerFunction();
      if (handlers[fn]) ScriptApp.deleteTrigger(triggers[i]);
    }
  } catch (err) {
    // script.scriptapp 未許可でも放置でOK
  }
}

function styleUrlIndexSheet_(sheet, links) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(hubTabColorFor_('URL一覧'));

  var dn = dnTheme_();
  var header = [['#', 'KIND', 'TITLE', 'URL']];
  sheet.getRange(1, 1, 1, 4).setValues(header);
  sheet.getRange(1, 1, 1, 4)
    .setBackground(dn.ink)
    .setFontColor(dn.paper)
    .setFontFamily('Noto Sans JP').setFontStyle('italic')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  if (!links.length) {
    sheet.getRange(2, 1, 1, 4).setValues([['', '', 'リンクなし', '']]);
    sheet.setColumnWidths(1, 1, 48);
    sheet.setColumnWidths(2, 1, 88);
    sheet.setColumnWidths(3, 1, 360);
    sheet.setColumnWidths(4, 1, 520);
    return;
  }

  var rows = links.map(function (item, idx) {
    return [idx + 1, item.kind, item.label, item.url];
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);

  var body = sheet.getRange(2, 1, rows.length, 4);
  body
    .setBackground(dn.paper)
    .setFontColor(dn.ink)
    .setFontFamily('Noto Sans JP').setFontStyle('italic')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  for (var i = 0; i < rows.length; i++) {
    if (i % 2 === 1) {
      sheet.getRange(i + 2, 1, 1, 4).setBackground(dn.cream);
    }
  }

  sheet.getRange(2, 1, rows.length, 1).setFontColor(dn.ash).setHorizontalAlignment('right');
  sheet.getRange(2, 2, rows.length, 1).setFontColor(dn.ash).setHorizontalAlignment('center');
  sheet.getRange(2, 4, rows.length, 1).setFontColor(dn.blood);

  // URL をクリッカブルに
  for (var r = 0; r < rows.length; r++) {
    var url = rows[r][3];
    sheet.getRange(r + 2, 4).setFormula('=HYPERLINK("' + String(url).replace(/"/g, '""') + '","' + String(url).replace(/"/g, '""') + '")');
  }

  sheet.setColumnWidths(1, 1, 48);
  sheet.setColumnWidths(2, 1, 88);
  sheet.setColumnWidths(3, 1, 360);
  sheet.setColumnWidths(4, 1, 560);
  sheet.setFrozenRows(1);
  sheet.setRowHeights(1, 1, 32);
  if (rows.length) sheet.setRowHeights(2, rows.length, 28);

  // 余白列を使わない（A-Dのみ）
  try {
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 4) sheet.deleteColumns(5, maxCols - 4);
  } catch (ignoredCols) {}
}

/** 今日のカレンダー + Workspace 同期データ（Vercel / AI 用） */
function getDayContextForApi_(dateYmd) {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var ymd = String(dateYmd || '').trim() || today;
  var tasks = loadWorkspaceTasksFromSheet();
  return {
    ok: true,
    date: ymd,
    today: today,
    timezone: tz,
    calendarEvents: getCalendarEventsForYmd_(ymd),
    workspace: tasks,
  };
}

function getTodayCalendarEvents_() {
  return getCalendarEventsForYmd_(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  );
}

/**
 * 指定日の予定。自分のアカウント（デフォルトカレンダー）に登録されている
 * 予定のみを対象にする。共有された他人のカレンダーは読み込まない。
 */
function getCalendarEventsForYmd_(ymd) {
  try {
    var tz = Session.getScriptTimeZone();
    var day = String(ymd || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) return [];
    var start = Utilities.parseDate(day + ' 00:00:00', tz, 'yyyy-MM-dd HH:mm:ss');
    var end = Utilities.parseDate(day + ' 23:59:59', tz, 'yyyy-MM-dd HH:mm:ss');

    var cal = CalendarApp.getDefaultCalendar();
    if (!cal) return [];
    var events = cal.getEvents(start, end) || [];

    var out = [];
    var seen = {};
    for (var i = 0; i < events.length; i++) {
      var item = packCalendarEvent_(events[i], tz);
      if (!item) continue;
      var key = item.title + '|' + item.start + '|' + item.end + '|' + item.isAllDay;
      if (seen[key]) continue;
      seen[key] = 1;
      out.push(item);
    }
    out.sort(function (a, b) {
      if (a.isAllDay !== b.isAllDay) return a.isAllDay ? -1 : 1;
      return String(a.start).localeCompare(String(b.start));
    });
    return out;
  } catch (err) {
    console.error('Calendar:', err);
    return [];
  }
}

function packCalendarEvent_(ev, tz) {
  if (!ev) return null;
  var pick = function (fn) {
    try {
      var v = fn();
      return v == null ? '' : String(v);
    } catch (err) {
      return '';
    }
  };
  var isAllDay = false;
  try {
    isAllDay = ev.isAllDayEvent();
  } catch (err) {
    isAllDay = false;
  }
  var description = cleanEventDescription_(pick(function () { return ev.getDescription(); }));
  var location = pick(function () { return ev.getLocation(); });
  var guests = [];
  try {
    var list = ev.getGuestList(true) || [];
    for (var i = 0; i < list.length && i < 20; i++) {
      guests.push(list[i].getName() || list[i].getEmail());
    }
  } catch (err) {
    guests = [];
  }
  return {
    id: pick(function () { return ev.getId(); }),
    title: pick(function () { return ev.getTitle(); }),
    start: isAllDay ? '' : Utilities.formatDate(ev.getStartTime(), tz, 'HH:mm'),
    end: isAllDay ? '' : Utilities.formatDate(ev.getEndTime(), tz, 'HH:mm'),
    isAllDay: isAllDay,
    location: location,
    description: description.length > 800 ? description.slice(0, 800) : description,
    guests: guests,
    organizer: pick(function () { return ev.getCreators().join(', '); }),
    myStatus: pick(function () { return ev.getMyStatus(); }),
    meetUrl: extractMeetingUrl_(location + '\n' + description),
  };
}

/** 同期用マーカーなど、表示に不要な行を説明文から除く */
function cleanEventDescription_(raw) {
  return String(raw || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]*>/g, ' ')
    .split('\n')
    .filter(function (line) {
      var s = line.trim();
      if (!s) return false;
      if (/^\[SYNC_KEY:/.test(s)) return false;
      if (/^(原文|同期)\s*[:：]/.test(s)) return false;
      return true;
    })
    .join(' / ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** 説明・場所から Meet / Zoom / Teams 等の会議URLを拾う */
function extractMeetingUrl_(text) {
  var s = String(text || '');
  if (!s) return '';
  var patterns = [
    /https:\/\/meet\.google\.com\/[a-z0-9\-]+/i,
    /https:\/\/[a-z0-9.\-]*zoom\.us\/j\/[^\s<>"']+/i,
    /https:\/\/teams\.microsoft\.com\/[^\s<>"']+/i,
    /https:\/\/[a-z0-9.\-]*webex\.com\/[^\s<>"']+/i,
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = s.match(patterns[i]);
    if (m) return m[0];
  }
  return '';
}

function handleApiGet_(e) {
  try {
    var api = e && e.parameter ? String(e.parameter.api || '') : '';
    if (api === 'status') {
      return jsonOutput_({ ok: true, service: 'ryuta-workspace-gas', version: 'v2-hub-links' });
    }
    if (api === 'listSheets') {
      return jsonOutput_(listWorkspaceSheets_());
    }
    if (api === 'diagnoseImports') {
      return jsonOutput_(diagnoseImportsHealth_());
    }
    if (api === 'inspectSheet') {
      return jsonOutput_(
        inspectNamedSheetDeep_(
          String((e.parameter && e.parameter.name) || ''),
          Number((e.parameter && e.parameter.rows) || 40),
          Number((e.parameter && e.parameter.cols) || 20)
        )
      );
    }
    if (api === 'setupUnpaidView') {
      return jsonOutput_(setupUnpaidView_());
    }
    if (api === 'fillUnpaidNow') {
      return jsonOutput_(fillUnpaidNow_());
    }
    if (api === 'fillUnpaidTrendNow') {
      return jsonOutput_(unpaidFillTrendAll_());
    }
    if (api === 'flushUnpaidQueue') {
      var ssFlush = openWorkspaceSpreadsheet_();
      var joinQ = null;
      try { joinQ = joinListFlushQueue_(ssFlush); } catch (eJQ) { joinQ = { ok: false, message: String(eJQ) }; }
      var joinPush = joinListPushChecks_(ssFlush);
      var joinRef = null;
      try { joinRef = joinListRefresh_(ssFlush.getSheetByName(JOIN_LIST_SHEET_)); } catch (eJR) {}
      return jsonOutput_({
        unpaid: unpaidFlushQueue_(ssFlush),
        joinQueue: joinQ,
        joinList: joinPush,
        joinRefresh: joinRef
      });
    }
    if (api === 'styleUnpaidNow') {
      return jsonOutput_(styleUnpaidNow_());
    }
    if (api === 'restyleHubLook') {
      return jsonOutput_(restyleHubLook_());
    }
    if (api === 'setupHubHome') {
      return jsonOutput_(setupHubHome_());
    }
    if (api === 'setupMasterKpiCharts') {
      return jsonOutput_(setupMasterKpiCharts_());
    }
    if (api === 'setupMasterIntroKpi') {
      return jsonOutput_(setupMasterIntroKpi_());
    }
    if (api === 'setupSchoolDiscountImport') {
      return jsonOutput_(setupSchoolDiscountImport_());
    }
    if (api === 'hubOpen') {
      return jsonOutput_(hubOpenNamed_(String((e.parameter && e.parameter.name) || '')));
    }
    if (api === 'hubClose') {
      return jsonOutput_(hubCloseWork_());
    }
    if (api === 'peekExternal') {
      try {
        var exBook = SpreadsheetApp.openById(String(e.parameter.id || ''));
        var exName = String(e.parameter.name || '');
        if (!exName) {
          return jsonOutput_({
            ok: true,
            title: exBook.getName(),
            sheets: exBook.getSheets().map(function (s) {
              return { name: s.getName(), rows: s.getLastRow(), cols: s.getLastColumn(), hidden: s.isSheetHidden() };
            })
          });
        }
        var exSh = exBook.getSheetByName(exName);
        var exRg = exSh.getRange(String(e.parameter.range || 'A1:Z5'));
        return jsonOutput_({
          ok: true,
          values: exRg.getDisplayValues(),
          merged: exRg.getMergedRanges().map(function (m) { return m.getA1Notation(); }),
          frozenRows: exSh.getFrozenRows()
        });
      } catch (eEx) {
        return jsonOutput_({ ok: false, message: String(eEx.message || eEx) });
      }
    }
    if (api === 'peekRows') {
      return jsonOutput_(peekSheetRows_(
        String((e.parameter && e.parameter.name) || ''),
        Number((e.parameter && e.parameter.tail) || 8)
      ));
    }
    if (api === 'readRange') {
      try {
        var rrSh = openWorkspaceSpreadsheet_().getSheetByName(String(e.parameter.name || ''));
        var rr = rrSh.getRange(String(e.parameter.range || 'A1:C3'));
        return jsonOutput_({
          ok: true,
          values: rr.getDisplayValues(),
          formulas: rr.getFormulas(),
          merged: rr.getMergedRanges().map(function (m) { return m.getA1Notation(); }),
          timeZone: rrSh.getParent().getSpreadsheetTimeZone(),
          maxRows: rrSh.getMaxRows(),
          maxCols: rrSh.getMaxColumns()
        });
      } catch (eRR) {
        return jsonOutput_({ ok: false, message: String(eRR.message || eRR) });
      }
    }
    if (api === 'listTriggers') {
      return jsonOutput_({
        ok: true,
        triggers: ScriptApp.getProjectTriggers().map(function (t) {
          return { fn: t.getHandlerFunction(), type: String(t.getEventType()) };
        })
      });
    }
    if (api === 'styleMembershipMirrors') {
      return jsonOutput_(styleMembershipMirrors_());
    }
    if (api === 'setupKengakuJoinLive') {
      var kConfirm = String((e.parameter && e.parameter.confirm) || '');
      if (kConfirm === 'repair') return jsonOutput_(repairKengakuMirror_());
      return jsonOutput_(setupKengakuJoinLive_(kConfirm));
    }
    if (api === 'funnelPreview') {
      return jsonOutput_(previewTourToJoinFunnel_(Number((e.parameter && e.parameter.days) || 60)));
    }
    if (api === 'containerCheck') {
      var boundSs = null;
      try { boundSs = SpreadsheetApp.getActiveSpreadsheet(); } catch (eC) {}
      return jsonOutput_({ ok: true, bound: !!boundSs, id: boundSs ? boundSs.getId() : '' });
    }
    if (api === 'formatJoinList') {
      return jsonOutput_(formatJoinListMirrorKeepImport_());
    }
    if (api === 'repairRestrictedImports') {
      return jsonOutput_(repairRestrictedImports_());
    }
    if (api === 'fixMasterActuals') {
      return jsonOutput_(fixMasterActualFormulas_());
    }
    if (api === 'peekMaster') {
      return jsonOutput_(peekMasterKeyCells_());
    }
    if (api === 'rebuildUrlIndex') {
      return jsonOutput_(rebuildUrlIndexSheet_());
    }
    if (api === 'setupPromoImport') {
      return jsonOutput_(setupPromoImport_());
    }
    if (api === 'inspectBook') {
      return jsonOutput_(inspectSpreadsheetBook_(String((e.parameter && e.parameter.id) || '')));
    }
    if (api === 'setupReviewImport') {
      return jsonOutput_(setupReviewImport_());
    }
    if (api === 'syncReviewKyodo') {
      return jsonOutput_(syncReviewKyodo_());
    }
    if (api === 'setupMachineImport') {
      return jsonOutput_(setupMachineImport_());
    }
    if (api === 'dayContext') {
      if (!isApiAuthorized_(e, null)) return unauthorized_();
      return jsonOutput_(getDayContextForApi_(e && e.parameter ? e.parameter.date : ''));
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
        String((e.parameter && e.parameter.email) || ''),
        String((e.parameter && e.parameter.since) || '')
      ));
    }
    if (api === 'vendorDiscover') {
      return jsonOutput_(discoverVendorEmails_(String((e.parameter && e.parameter.company) || '')));
    }
    if (api === 'vendorFiles') {
      if (!isApiAuthorized_(e, null)) return unauthorized_();
      return jsonOutput_(listVendorFiles_(String((e.parameter && e.parameter.threadId) || '')));
    }
    if (api === 'vendorFile') {
      if (!isApiAuthorized_(e, null)) return unauthorized_();
      return jsonOutput_(getVendorFile_(
        String((e.parameter && e.parameter.threadId) || ''),
        String((e.parameter && e.parameter.messageId) || ''),
        String((e.parameter && e.parameter.index) || '0')
      ));
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
    { name: '添田', email: 'e-soeta@sekai.co.jp' },
    { name: '渋谷メーリス', email: 'shibuya_sekai@sekai.co.jp' },
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
    { name: '本部受注', email: 'honbu-order@ohtori-s.co.jp' },
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
  return q;
}

function withMailWindow_(query, since) {
  var q = String(query || '').replace(/\s+newer_than:\S+/g, '').replace(/\s+after:\S+/g, '').trim();
  var s = String(since || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    var d = new Date(s.slice(0, 10) + 'T00:00:00+09:00');
    d.setDate(d.getDate() - 1);
    var y = d.getFullYear();
    var m = d.getMonth() + 1;
    var day = d.getDate();
    var mm = (m < 10 ? '0' : '') + m;
    var dd = (day < 10 ? '0' : '') + day;
    return q + ' after:' + y + '/' + mm + '/' + dd;
  }
  if (/^\d+d$/.test(s) || /^\d+m$/.test(s)) return q + ' newer_than:' + s;
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
  return '(' + parts.join(' OR ') + ')';
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

function isUsefulAttachment_(att) {
  var name = String(att.getName() || '');
  var type = String(att.getContentType() || '').toLowerCase();
  var size = 0;
  try { size = Number(att.getSize() || 0); } catch (ignored) {}
  if (/\.(xlsx|xls|xlsm|csv|pdf|png|jpe?g|gif|webp|docx|doc|pptx|ppt)$/i.test(name)) return true;
  if (type.indexOf('pdf') >= 0 || type.indexOf('spreadsheet') >= 0 || type.indexOf('excel') >= 0) return true;
  if (type.indexOf('officedocument') >= 0) return true;
  if (type.indexOf('image/') === 0 && size >= 20000) return true;
  return false;
}

function listVendorFiles_(threadId) {
  try {
    threadId = String(threadId || '').trim();
    if (!threadId) return { ok: false, message: 'threadId がありません' };
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return { ok: false, message: 'スレッドが見つかりません' };
    var messages = thread.getMessages();
    var files = [];
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var atts = msg.getAttachments();
      for (var a = 0; a < atts.length; a++) {
        var att = atts[a];
        if (!isUsefulAttachment_(att)) continue;
        var size = 0;
        try { size = Number(att.getSize() || 0); } catch (ignoredSize) {}
        files.push({
          messageId: String(msg.getId() || ''),
          index: a,
          name: String(att.getName() || '添付'),
          type: String(att.getContentType() || ''),
          size: size
        });
      }
    }
    return { ok: true, files: files };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function getVendorFile_(threadId, messageId, index) {
  try {
    var idx = parseInt(String(index || '0'), 10);
    if (isNaN(idx) || idx < 0) return { ok: false, message: 'index が不正です' };
    var msg = null;
    if (messageId) {
      msg = GmailApp.getMessageById(String(messageId));
    } else if (threadId) {
      var thread = GmailApp.getThreadById(String(threadId));
      var messages = thread ? thread.getMessages() : [];
      msg = messages.length ? messages[messages.length - 1] : null;
    }
    if (!msg) return { ok: false, message: 'メールが見つかりません' };
    var atts = msg.getAttachments();
    if (idx >= atts.length) return { ok: false, message: '添付が見つかりません' };
    var att = atts[idx];
    var bytes = att.getBytes();
    if (bytes.length > 6 * 1024 * 1024) {
      return { ok: false, message: '6MB超です。Gmailで開いてください。' };
    }
    return {
      ok: true,
      name: String(att.getName() || '添付'),
      type: String(att.getContentType() || 'application/octet-stream'),
      data: Utilities.base64Encode(bytes)
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function vendorPinnedQuery_(email, myEmail) {
  email = sanitizeEmail_(email);
  myEmail = sanitizeEmail_(myEmail);
  if (!email) return '';
  if (myEmail) {
    return '(from:' + email + ' to:' + myEmail + ') OR (from:' + myEmail + ' to:' + email + ')';
  }
  return '(from:' + email + ' OR to:' + email + ')';
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

function syncOneVendorFromGmail_(name, pinnedEmail, since) {
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
    query = withMailWindow_(query, since);
    var limit = /^\d{4}-\d{2}-\d{2}/.test(String(since || '')) || /^\d+d$/.test(String(since || '')) ? 5 : 8;
    var threads = GmailApp.search(query, 0, limit);
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

    var threads = GmailApp.search(withMailWindow_(vendorQuery_(vendor.keys, vendor.domains), '5m'), 0, 15);
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
      var threads = GmailApp.search(withMailWindow_(vendorQuery_(vendor.keys, vendor.domains), '5m'), 0, 10);
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

var UNPAID_SOURCE_ID_ = '10vpQRDfTdwx_Wb7JaSm3lZCkTk8msLyf8ggAHhI1shI';
var UNPAID_SHEET_ = '未納管理';
var HUB_HOME_SHEET_ = 'トップ';

/** 4色：地 #EDEDED / 墨 #171717 / 灰 #444444 / 赤 #DA0037。赤は未回収と強調だけ */
function dnTheme_() {
  return {
    ink: '#171717',
    paper: '#EDEDED',
    cream: '#EDEDED',
    blood: '#DA0037',
    apple: '#DA0037',
    ash: '#444444',
    line: '#444444',
    ghost: '#EDEDED'
  };
}

function hubFontFamily_() {
  return 'Noto Sans JP';
}

function hubType_(range) {
  return range.setFontFamily(hubFontFamily_()).setFontStyle('italic');
}

/** 空セルも含めて斜体にしておく。あとに入力した文字も同じ書体になる */
function hubStampType_(sh, rows, cols) {
  if (!sh) return;
  var maxR = sh.getMaxRows();
  var maxC = sh.getMaxColumns();
  var r = Math.min(Math.max(Number(rows) || Math.max(sh.getLastRow(), 80) + 80, 2), maxR);
  var c = Math.min(Math.max(Number(cols) || Math.max(sh.getLastColumn(), 8), 1), Math.min(maxC, 40));
  try { hubType_(sh.getRange(1, 1, r, c)); } catch (e0) {}
}

function hubTabColorFor_(name) {
  var t = dnTheme_();
  if (name.indexOf('未納') === 0) return t.blood;
  if (/backup|シート\d+/.test(name) || name === 'Tasks' || name === 'WorkspaceSync') return t.ash;
  return t.ink;
}

function applyHubTabColors_(ss) {
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    try { sheets[i].setTabColor(hubTabColorFor_(sheets[i].getName())); } catch (e0) {}
  }
}

function hubIsAlwaysHidden_(name) {
  if (name === HUB_HOME_SHEET_) return false;
  if (/^経堂_/.test(name)) return true;
  if (/backup/i.test(name)) return true;
  if (/^シート\d+$/.test(name)) return true;
  if (name === 'Tasks' || name === 'WorkspaceSync') return true;
  if (name === '未納_同期' || name === '未納_推移_グラフ元') return true;
  return false;
}

function hubIsWorkSheet_(name) {
  return !!name && name !== HUB_HOME_SHEET_ && !hubIsAlwaysHidden_(name);
}

var HUB_KEY_BASE_ = 50;
var HUB_KEY_SPAN_ = 8;

function hubEnsureKeyCols_(sh) {
  var need = HUB_KEY_BASE_ + HUB_KEY_SPAN_ + 2;
  if (sh.getMaxColumns() < need) {
    sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  }
}

function hubHideInternalCols_(sh) {
  hubEnsureKeyCols_(sh);
  try { sh.showColumns(1, 8); } catch (eShow) {}
  var start = 9;
  var n = sh.getMaxColumns() - start + 1;
  if (n > 0) {
    try { sh.hideColumns(start, n); } catch (eHide) {}
  }
}

function hubWriteKey_(sh, row, col, value) {
  var t = dnTheme_();
  sh.getRange(row, HUB_KEY_BASE_ + col)
    .setValue(value)
    .setFontColor(t.paper)
    .setBackground(t.paper)
    .setFontSize(1)
    .setFontWeight('normal')
    .setHorizontalAlignment('left');
}

function hubUi_() {
  var t = dnTheme_();
  return {
    bg: t.paper,
    card: t.paper,
    ink: t.ink,
    mute: t.ash,
    line: t.ink,
    onBg: t.ink,
    onFg: t.paper,
    shut: t.ink,
    rail: t.ink
  };
}

function hubCatalog_(ss) {
  var items = [
    { group: '数字', name: '経堂マスタ', title: 'マスタ' },
    { group: '数字', name: '【経堂】会員動向', title: '動向' },
    { group: '未納', name: '未納管理', title: '今月' },
    { group: '未納', name: '未納管理_推移', title: '推移' },
    { group: '現場', name: '見学体験申請', title: '見学' },
    { group: '現場', name: '学割', title: '学割' },
    { group: '現場', name: '口コミ_経堂', title: '口コミ' },
    { group: '現場', name: 'マシンレクチャー申込', title: 'レクチャー' },
    { group: '現場', name: '入会者一覧＋自動メール管理', title: '入会者' }
  ];
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (/^販促_/.test(n) && n.indexOf('backup') === -1 && n.indexOf('学校') === -1) {
      items.push({ group: '販促', name: n, title: hubShortTitle_(n.replace(/^販促_/, '')) });
    }
  }
  return items.filter(function (it) {
    return !!ss.getSheetByName(it.name);
  });
}

function hubShortTitle_(title) {
  if (/乗換/.test(title)) return '乗り換え';
  if (/紹介|ペア/.test(title)) return '紹介';
  if (/学校/.test(title)) return '学校';
  if (/ラグビー/.test(title)) return 'ラグビー';
  if (/6ヶ月|6カ月/.test(title)) return '6ヶ月';
  return title;
}

function hubBookUrl_(id) {
  return 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
}

function hubSourceForName_(name) {
  if (name === '経堂マスタ' || name.indexOf('会員動向') !== -1 || /^経堂_/.test(name)) {
    return { url: hubBookUrl_(RECEPTION_SOURCE_ID_), label: '元のシートを開く ↗' };
  }
  if (name.indexOf('未納') === 0) {
    return { url: hubBookUrl_(UNPAID_SOURCE_ID_), label: '元の未納管理ドライブを開く ↗' };
  }
  if (name.indexOf('見学') === 0) {
    return { url: hubBookUrl_(KENGAKU_SOURCE_ID_), label: '元のシートを開く ↗' };
  }
  if (name.indexOf('口コミ') === 0) {
    return { url: hubBookUrl_(REVIEW_SOURCE_ID_), label: '元のシートを開く ↗' };
  }
  if (name.indexOf('マシン') === 0 || name.indexOf('入会者一覧') === 0) {
    return { url: hubBookUrl_(MACHINE_SOURCE_ID_), label: '元のシートを開く ↗' };
  }
  if (name === '学割' || name.indexOf('学校') !== -1) {
    return { url: schoolFormOpenUrl_(), label: '元のシートを開く ↗' };
  }
  if (name.indexOf('販促_') === 0) {
    return { url: hubBookUrl_(PROMO_SOURCE_ID_), label: '元のシートを開く ↗' };
  }
  return null;
}

function hubPaintOpenSourceCell_(sh, row, col, cols, src) {
  var t = dnTheme_();
  var solid = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  cols = Math.max(Number(cols) || 1, 1);
  try { sh.getRange(row, col, 1, cols).breakApart(); } catch (e0) {}
  var rng = sh.getRange(row, col, 1, cols);
  if (cols > 1) rng = rng.merge();
  rng.setFormula('=HYPERLINK("' + String(src.url).replace(/"/g, '""') + '","' + String(src.label).replace(/"/g, '""') + '")')
    .setBackground(t.ink)
    .setFontColor(t.paper)
    .setFontFamily(hubFontFamily_())
    .setFontStyle('italic')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, t.blood, solid);
  sh.setRowHeight(row, 28);
}

function hubEnsureSourceBanner_(sh) {
  if (!sh) return;
  var n = sh.getName();
  if (n === HUB_HOME_SHEET_ || n === UNPAID_SHEET_ || n === '経堂マスタ' || n === '学割' || n === '販促_学校関係者') return;
  if (n === JOIN_LIST_SHEET_ || n === UNPAID_TREND_SHEET_) return;
  if (hubIsAlwaysHidden_(n) && n !== UNPAID_TREND_SHEET_) return;
  var src = hubSourceForName_(n);
  if (!src) return;
  var f1 = String(sh.getRange(1, 1).getFormula() || '');
  var d1 = String(sh.getRange(1, 1).getDisplayValue() || '');
  if (/元のシートを開く|元の未納管理ドライブ/.test(f1 + d1)) {
    var span = Math.min(Math.max(sh.getLastColumn(), 4), 6);
    hubPaintOpenSourceCell_(sh, 1, 1, span, src);
    return;
  }
  if (n === UNPAID_TREND_SHEET_) return;
  if (n === '見学体験申請') {
    hubPaintOpenSourceCell_(sh, 1, 12, 2, src);
    sh.setColumnWidth(12, 168);
    sh.setColumnWidth(13, 28);
    return;
  }
  if (n.indexOf('会員動向') !== -1) {
    hubPaintOpenSourceCell_(sh, 1, 16, 2, src);
    sh.setColumnWidth(16, 168);
    sh.setColumnWidth(17, 28);
    return;
  }
  if (/IMPORTRANGE|QUERY\(/.test(f1) || d1) {
    sh.insertRowBefore(1);
    sh.setFrozenRows(sh.getFrozenRows() + 1);
  }
  hubPaintOpenSourceCell_(sh, 1, 1, Math.min(Math.max(sh.getLastColumn(), 4), 6), src);
}

function hubSourceLinks_() {
  return [
    { title: '経堂　受付状況表 ↗', url: 'https://docs.google.com/spreadsheets/d/14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w/edit' },
    { title: '26年度未納管理ドライブ【経堂】 ↗', url: 'https://docs.google.com/spreadsheets/d/' + UNPAID_SOURCE_ID_ + '/edit' },
    { title: '経堂　見学・体験フォーム ↗', url: 'https://docs.google.com/spreadsheets/d/1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY/edit' },
    { title: 'JOYFIT24経堂追加販促 ↗', url: 'https://docs.google.com/spreadsheets/d/1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E/edit' },
    { title: '学校関係者割フォーム ↗', url: hubBookUrl_(SCHOOL_FORM_SOURCE_ID_) },
    { title: 'EAST口コミ回答者 ↗', url: 'https://docs.google.com/spreadsheets/d/13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM/edit' },
    { title: '20分マシンレクチャー・自動送信メール ↗', url: 'https://docs.google.com/spreadsheets/d/1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8/edit' },
    { title: '口コミ付与アプリ ↗', url: REVIEW_GRANT_APP_URL_ }
  ];
}

function hubPaintTile_(home, row, col, on) {
  var u = hubUi_();
  home.getRange(row, col)
    .setBackground(on ? u.onBg : u.card)
    .setFontColor(on ? u.onFg : u.ink);
}

function hubApplyFocus_(ss, sheetName) {
  if (!sheetName) return hubCloseWork_();
  var target = ss.getSheetByName(sheetName);
  if (target) {
    try { target.showSheet(); } catch (e0) {}
  }
  return { ok: true, open: sheetName };
}

function hubCloseWork_() {
  var ss = openWorkspaceSpreadsheet_();
  var home = ss.getSheetByName(HUB_HOME_SHEET_);
  if (home) {
    try { home.showSheet(); } catch (e0) {}
  }
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (n === HUB_HOME_SHEET_ || hubIsAlwaysHidden_(n)) continue;
    if (!sheets[i].isSheetHidden()) {
      try { sheets[i].hideSheet(); } catch (e1) {}
    }
  }
  if (home) {
    hubResetTiles_(home);
    hubHideInternalCols_(home);
    ss.setActiveSheet(home);
  }
  return { ok: true, open: HUB_HOME_SHEET_ };
}

function hubResetTiles_(home) {
  var last = Math.max(home.getLastRow(), 4);
  var keys = home.getRange(1, HUB_KEY_BASE_ + 1, last, 8).getDisplayValues();
  var r;
  var c;
  for (r = 0; r < keys.length; r++) {
    for (c = 0; c < keys[r].length; c++) {
      if (String(keys[r][c] || '').indexOf('SHEET:') === 0) {
        hubPaintTile_(home, r + 1, c + 1, false);
      }
    }
  }
}

function hubShowHome_() {
  var ss = openWorkspaceSpreadsheet_();
  var home = ss.getSheetByName(HUB_HOME_SHEET_);
  if (!home) return setupHubHome_();
  try { home.showSheet(); } catch (e0) {}
  ss.setActiveSheet(home);
  return { ok: true, open: HUB_HOME_SHEET_ };
}

function hubOpenNamed_(name) {
  var ss = openWorkspaceSpreadsheet_();
  var n = String(name || '').trim();
  if (!n) return { ok: false, message: 'name が空です' };
  if (!ss.getSheetByName(n)) return { ok: false, message: 'シートがありません: ' + n };
  return hubApplyFocus_(ss, n);
}

function hubRefreshStatus_(home, ss) {
  var last = Math.max(home.getLastRow(), 4);
  var keys = home.getRange(1, HUB_KEY_BASE_ + 1, last, 8).getDisplayValues();
  var r;
  var c;
  for (r = 0; r < keys.length; r++) {
    for (c = 0; c < keys[r].length; c++) {
      var key = String(keys[r][c] || '');
      if (key.indexOf('SHEET:') !== 0) continue;
      var target = ss.getSheetByName(key.slice(6));
      hubPaintTile_(home, r + 1, c + 1, !!(target && !target.isSheetHidden()));
    }
  }
}

function handleHubHomeSelect_(e) {
  if (!e || !e.range) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== HUB_HOME_SHEET_) return;
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var key = String(sh.getRange(row, HUB_KEY_BASE_ + col).getValue() || '');
  if (!key || key.indexOf('SHEET:') !== 0) {
    hubRefreshStatus_(sh, sh.getParent());
    return;
  }
  var name = key.slice(6);
  var ss = sh.getParent();
  var target = ss.getSheetByName(name);
  if (!target) return;
  try { sh.getRange(1, 1).activate(); } catch (eSel) {}
  try { target.showSheet(); } catch (e1) {}
  hubPaintTile_(sh, row, col, true);
  ss.setActiveSheet(target);
}

function handleHubHomeEdit_(e) {}

function hubPaintHomeTitle_(sh) {
  var u = hubUi_();
  var t = dnTheme_();
  var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  try { sh.getRange(1, 1, 1, 6).breakApart(); } catch (e0) {}
  hubType_(sh.getRange(1, 1, 1, 6).merge())
    .setValue('経堂')
    .setFontSize(22)
    .setFontWeight('bold')
    .setFontColor(u.ink)
    .setBackground(u.bg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, t.blood, medium);
  sh.setRowHeight(1, 48);
}

function restyleHubHomeLook_(sh) {
  var u = hubUi_();
  var t = dnTheme_();
  var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  sh.setTabColor(t.ink);
  var last = Math.max(sh.getLastRow(), 24);
  var rows = Math.min(last, 40);
  hubType_(sh.getRange(1, 1, rows, 8)).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(1, 1, rows, 7).setBackground(u.bg).setFontColor(u.ink);
  sh.getRange(1, 8, rows, 1).setBackground(u.rail).setFontColor(u.bg);
  var vals = sh.getRange(1, 1, rows, 8).getDisplayValues();
  var forms = sh.getRange(1, 1, rows, 8).getFormulas();
  var groups = { '数字': 1, '未納': 1, '現場': 1, '販促': 1 };
  var r;
  var c;
  for (r = 0; r < vals.length; r++) {
    var g = String(vals[r][0] || '');
    if (groups[g]) {
      sh.getRange(r + 1, 1)
        .setBackground(g === '未納' ? t.blood : u.ink)
        .setFontColor(u.bg)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
    for (c = 1; c <= 5; c++) {
      var f = String(forms[r][c] || '');
      var v = String(vals[r][c] || '');
      var prev = r > 0 ? String(forms[r - 1][c] || '') + String(vals[r - 1][c] || '') : '';
      if (/元のシートを開く|元の未納管理ドライブ/.test(f + v)) {
        var prevName = r > 0 ? String(vals[r - 1][c] || '') : '';
        sh.getRange(r + 1, c + 1)
          .setBackground(/未納/.test(prevName) ? t.blood : t.ash)
          .setFontColor(t.paper)
          .setFontWeight('bold')
          .setFontSize(9)
          .setHorizontalAlignment('center')
          .setBorder(false, true, false, true, false, false, t.ink, medium);
      } else if (/#gid=/.test(f)) {
        sh.getRange(r + 1, c + 1)
          .setBackground(u.card)
          .setFontColor(u.ink)
          .setFontWeight('bold')
          .setHorizontalAlignment('center')
          .setBorder(true, true, false, true, false, false, t.ink, medium);
      } else if (!v && !f && /元のシートを開く|元の未納管理ドライブ/.test(prev)) {
        sh.getRange(r + 1, c + 1).setBackground(t.ink).setFontColor(t.ink)
          .setBorder(false, true, true, true, false, false, t.blood, medium);
      }
    }
  }
  hubPaintHomeTitle_(sh);
  sh.getRange(1, 8).setBackground(u.rail).setFontColor(u.bg).setFontWeight('bold')
    .setHorizontalAlignment('center');
}

function setupHubHome_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName(HUB_HOME_SHEET_);
    if (!sh) sh = ss.insertSheet(HUB_HOME_SHEET_, 0);
    sh.showSheet();
    hubEnsureKeyCols_(sh);
    var maxR = sh.getMaxRows();
    var maxC = sh.getMaxColumns();
    try { sh.getRange(1, 1, maxR, maxC).breakApart(); } catch (eBr) {}
    try { sh.getRange(1, 1, maxR, maxC).clearDataValidations(); } catch (eVal) {}
    try { sh.clearConditionalFormatRules(); } catch (e0) {}
    sh.clear();
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(1);
    try { sh.getRange(1, 1, maxR, maxC).clearDataValidations(); } catch (eVal2) {}

    var u = hubUi_();
    var t = dnTheme_();
    var items = hubCatalog_(ss);
    var links = hubSourceLinks_();
    var groups = [];
    var seen = {};
    var i;
    for (i = 0; i < items.length; i++) {
      if (!seen[items[i].group]) {
        seen[items[i].group] = [];
        groups.push(items[i].group);
      }
      seen[items[i].group].push(items[i]);
    }

    hubType_(sh.getRange(1, 1, 40, 8))
      .setBackground(u.bg)
      .setFontColor(u.ink)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setBorder(false, false, false, false, false, false)
      .setFontWeight('normal')
      .setWrap(true);

    hubPaintHomeTitle_(sh);

    var row = 2;
    var g;
    var solid = SpreadsheetApp.BorderStyle.SOLID;
    var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    for (g = 0; g < groups.length; g++) {
      var group = groups[g];
      var list = seen[group];
      var nameRow = row;
      var srcRow = row + 1;
      var shadowRow = row + 2;
      try { sh.getRange(nameRow, 1, 3, 1).breakApart(); } catch (eMg) {}
      sh.getRange(nameRow, 1, 3, 1)
        .merge()
        .setValue(group)
        .setFontSize(12)
        .setFontWeight('bold')
        .setFontStyle('italic')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBackground(group === '未納' ? t.blood : u.ink)
        .setFontColor(u.bg)
        .setBorder(true, true, true, true, false, false, t.ink, medium);
      var p;
      for (p = 0; p < list.length; p++) {
        var col = 2 + p;
        var it = list[p];
        var target = ss.getSheetByName(it.name);
        if (!target) continue;
        sh.getRange(nameRow, col)
          .setFormula('=HYPERLINK("#gid=' + target.getSheetId() + '","' + String(it.name).replace(/"/g, '""') + '")')
          .setFontSize(11)
          .setFontWeight('bold')
          .setFontStyle('italic')
          .setWrap(true)
          .setBackground(u.card)
          .setFontColor(u.ink)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setBorder(true, true, false, true, false, false, t.ink, medium);
        hubWriteKey_(sh, nameRow, col, 'SHEET:' + it.name);
        var src = hubSourceForName_(it.name);
        if (src) {
          sh.getRange(srcRow, col)
            .setFormula('=HYPERLINK("' + String(src.url).replace(/"/g, '""') + '","元のシートを開く ↗")')
            .setFontSize(9)
            .setFontWeight('bold')
            .setFontStyle('italic')
            .setWrap(false)
            .setBackground(it.group === '未納' ? t.blood : t.ash)
            .setFontColor(t.paper)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle')
            .setBorder(false, true, false, true, false, false, t.ink, medium);
          hubWriteKey_(sh, srcRow, col, 'SRC:' + it.name);
        }
        sh.getRange(shadowRow, col)
          .setBackground(t.ink)
          .setFontColor(t.ink)
          .setBorder(false, true, true, true, false, false, t.blood, medium);
        hubWriteKey_(sh, shadowRow, col, 'SHADOW');
      }
      sh.setRowHeight(nameRow, 52);
      sh.setRowHeight(srcRow, 28);
      sh.setRowHeight(shadowRow, 6);
      row += 3;
      sh.setRowHeight(row, 12);
      row += 1;
    }

    sh.getRange(1, 7, 40, 1).setBackground(u.bg).setBorder(false, false, false, false, false, false);
    sh.getRange(1, 8).setValue('引用元').setFontSize(11).setFontWeight('bold').setFontStyle('italic')
      .setBackground(u.rail).setFontColor(u.bg).setHorizontalAlignment('center')
      .setBorder(false, false, true, false, false, false, t.blood, medium);
    var c;
    for (c = 0; c < links.length; c++) {
      var lr = 2 + c;
      var title = String(links[c].title).replace(/"/g, '""');
      var url = String(links[c].url).replace(/"/g, '""');
      sh.getRange(lr, 8)
        .setFormula('=HYPERLINK("' + url + '","' + title + '")')
        .setFontColor(u.bg)
        .setFontSize(10)
        .setFontWeight('bold')
        .setFontStyle('italic')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true)
        .setBackground(u.rail)
        .setBorder(false, false, true, false, false, false, t.ash, solid);
    }

    try {
      var leftover = ss.getSheetByName('URL一覧');
      if (leftover && ss.getSheets().length > 1) ss.deleteSheet(leftover);
    } catch (eDel) {}

    sh.setHiddenGridlines(true);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 64);
    for (i = 2; i <= 6; i++) sh.setColumnWidth(i, 176);
    sh.setColumnWidth(7, 28);
    sh.setColumnWidth(8, 280);
    hubHideInternalCols_(sh);

    restyleHubHomeLook_(sh);
    hubHideInternalCols_(sh);
    applyHubTabColors_(ss);
    return {
      ok: true,
      sheet: HUB_HOME_SHEET_,
      items: items.length,
      visible: ss.getSheets().filter(function (s) { return !s.isSheetHidden(); }).map(function (s) { return s.getName(); })
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function restyleHeaderBody_(sh) {
  var t = dnTheme_();
  var last = Math.max(sh.getLastRow(), 1);
  var cols = Math.max(sh.getLastColumn(), 1);
  var f1 = String(sh.getRange(1, 1).getFormula() || '');
  var d1 = String(sh.getRange(1, 1).getDisplayValue() || '');
  var headerRow = /元のシートを開く|元の未納管理ドライブ/.test(f1 + d1) ? 2 : 1;
  var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  hubType_(sh.getRange(headerRow, 1, 1, cols))
    .setBackground(t.ink)
    .setFontColor(t.paper)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, t.blood, medium);
  if (last > headerRow) {
    hubType_(sh.getRange(headerRow + 1, 1, last - headerRow, cols))
      .setBackground(t.paper)
      .setFontColor(t.ink)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
  hubStampType_(sh, last + 80, cols);
}

function schoolDiscountQueryFormula_() {
  return '=QUERY(IMPORTRANGE("' + SCHOOL_FORM_SOURCE_ID_ +
    '","\'' + SCHOOL_FORM_SHEET_.replace(/'/g, "''") + '\'!A:P"),' +
    '"select * where Col2 = \'JOYFIT24 経堂\' order by Col1 desc",1)';
}

function schoolDiscountHeaders_() {
  return [
    '申請日時', '店舗', '氏名', '連絡先', '区分', '学校名', '卒業見込', '証明書類',
    'LITEオプション', '趣味の期間', '続く理由', 'やめた趣味', 'やめたきっかけ',
    '趣味の金額', '健康の金額', '美容の金額'
  ];
}

function permitImportRange_(ss, donorId) {
  try {
    var resp = UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + ss.getId() +
        '/externaldata/addimportrangepermissions?donorDocId=' + donorId,
      { method: 'post', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
    );
    return String(resp.getResponseCode());
  } catch (eP) {
    return 'error: ' + (eP && eP.message ? eP.message : eP);
  }
}

function setupSchoolDiscountSheet_(ss) {
  var old = ss.getSheetByName('販促_学校関係者');
  var sh = ss.getSheetByName(SCHOOL_WS_SHEET_);
  if (!sh && old) {
    old.setName(SCHOOL_WS_SHEET_);
    sh = old;
  }
  if (!sh) sh = ss.insertSheet(SCHOOL_WS_SHEET_);
  if (old && old.getSheetId() !== sh.getSheetId()) {
    try { ss.deleteSheet(old); } catch (eDel) { try { old.hideSheet(); } catch (eH) {} }
  }
  var formula = schoolDiscountQueryFormula_();
  var f1 = String(sh.getRange(1, 1).getFormula() || '');
  var f2 = String(sh.getRange(2, 1).getFormula() || '');
  var hooked = /1mmG_xM1WoWFKgpmOl5obsKXo/.test(f1 + f2);
  if (!hooked) {
    try { sh.getRange(1, 1, Math.min(sh.getMaxRows(), 20), Math.max(sh.getLastColumn(), SCHOOL_FORM_COLS_)).breakApart(); } catch (e0) {}
    sh.clear();
    try { sh.clearConditionalFormatRules(); } catch (e1) {}
    sh.getRange(1, 1).setFormula(formula);
  } else if (/QUERY\(IMPORTRANGE/.test(f1)) {
    sh.getRange(1, 1).setFormula(formula);
  } else if (/QUERY\(IMPORTRANGE/.test(f2)) {
    sh.getRange(2, 1).setFormula(formula);
  }
  if (sh.getMaxColumns() < SCHOOL_FORM_COLS_) {
    sh.insertColumnsAfter(sh.getMaxColumns(), SCHOOL_FORM_COLS_ - sh.getMaxColumns());
  }
  if (sh.getMaxRows() < SCHOOL_MIN_ROWS_) {
    sh.insertRowsAfter(sh.getMaxRows(), SCHOOL_MIN_ROWS_ - sh.getMaxRows());
  }
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(1);
  sh.setTabColor(hubTabColorFor_(SCHOOL_WS_SHEET_));
  var t = dnTheme_();
  var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  hubType_(sh.getRange(1, 1, 1, SCHOOL_FORM_COLS_))
    .setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, t.blood, medium);
  var bodyRows = Math.max(sh.getMaxRows() - 1, 1);
  hubType_(sh.getRange(2, 1, bodyRows, SCHOOL_FORM_COLS_))
    .setBackground(t.paper).setFontColor(t.ink)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(2, 1, bodyRows, 1).setNumberFormat('yyyy/mm/dd HH:mm');
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 180);
  sh.setColumnWidth(6, 160);
  var c;
  for (c = 5; c <= SCHOOL_FORM_COLS_; c++) {
    if (c !== 6) sh.setColumnWidth(c, 120);
  }
  hubStampType_(sh, 200, SCHOOL_FORM_COLS_);
  try {
    if (sh.getMaxColumns() < SCHOOL_FORM_COLS_ + 1) {
      sh.insertColumnsAfter(sh.getMaxColumns(), SCHOOL_FORM_COLS_ + 1 - sh.getMaxColumns());
    }
    hubPaintOpenSourceCell_(sh, 1, SCHOOL_FORM_COLS_ + 1, 1, {
      url: schoolFormOpenUrl_(),
      label: '元のシートを開く ↗'
    });
    sh.setColumnWidth(SCHOOL_FORM_COLS_ + 1, 168);
  } catch (eSrc) {}
  return sh;
}

function schoolMasterListFormula_() {
  return '=IFERROR(QUERY(\'' + SCHOOL_WS_SHEET_ + '\'!A2:P,"select Col1,Col3,Col6 where Col1 is not null order by Col1 desc",0),"")';
}

function masterIntroListFormula_() {
  return '=IFERROR(QUERY({\'販促_紹介・ペア入会\'!A3:A,\'販促_紹介・ペア入会\'!B3:B,\'販促_紹介・ペア入会\'!F3:F},' +
    '"select Col1,Col2,Col3 where Col1 is not null order by Col1 desc",0),"")';
}

function hubSheetGidUrl_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return hubBookUrl_(ss.getId());
  return hubBookUrl_(ss.getId()) + '#gid=' + sh.getSheetId();
}

function schoolFormOpenUrl_() {
  return hubBookUrl_(SCHOOL_FORM_SOURCE_ID_) + '#gid=667254510';
}

function clearMasterSideCard_(sh, col, w) {
  var lastR = sh.getMaxRows();
  var rng = sh.getRange(1, col, lastR, w);
  try { rng.breakApart(); } catch (e0) {}
  try { rng.clearDataValidations(); } catch (e1) {}
  try { rng.clearContent(); } catch (e2) {}
  try { rng.clearNote(); } catch (e3) {}
  try { rng.setNumberFormat('General'); } catch (e4) {}
}

function paintMasterSideCard_(sh, col, w, titleF, linkF, headers, listF, widths) {
  var t = dnTheme_();
  var hair = SpreadsheetApp.BorderStyle.SOLID;
  var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  clearMasterSideCard_(sh, col, w);
  sh.getRange(1, col, 1, w).merge()
    .setFormula(titleF)
    .setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, t.blood, medium);
  hubType_(sh.getRange(1, col));
  sh.setRowHeight(1, 30);
  sh.getRange(2, col, 1, w).merge()
    .setFormula(linkF)
    .setBackground(t.ash).setFontColor(t.paper).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, true, false, true, false, false, t.ink, hair);
  hubType_(sh.getRange(2, col)).setFontSize(9);
  sh.setRowHeight(2, 22);
  sh.getRange(3, col, 1, w).setValues([headers])
    .setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, t.ash, hair);
  hubType_(sh.getRange(3, col, 1, w)).setFontSize(9);
  sh.setRowHeight(3, 24);
  sh.getRange(4, col).setFormula(listF);
  var body = Math.min(Math.max(sh.getMaxRows() - 3, 1), 400);
  hubType_(sh.getRange(4, col, body, w))
    .setBackground(t.paper).setFontColor(t.ink)
    .setVerticalAlignment('middle');
  sh.getRange(4, col, body, 1).setNumberFormat('yyyy/mm/dd HH:mm')
    .setHorizontalAlignment('center');
  var i;
  for (i = 1; i < w; i++) {
    sh.getRange(4, col + i, body, 1).setHorizontalAlignment(i === w - 1 && w >= 4 ? 'center' : 'left');
  }
  var zebra = Math.min(body, 80);
  for (i = 1; i < zebra; i += 2) {
    sh.getRange(4 + i, col, 1, w).setBackground('#E6E6E6');
  }
  if (widths) {
    for (i = 0; i < widths.length && i < w; i++) sh.setColumnWidth(col + i, widths[i]);
  }
}

function setupMasterSchoolPanel_(sh) {
  return setupMasterMonthPanel_(sh);
}

function masterMonthTabs_() {
  return ['すべて', '見学体験', '紹介', '学割', 'ラグビー割', '乗り換え', '6ヶ月', 'レクチャー', '口コミ', '入会メール', '入会', '退会'];
}

function masterMonthWhere_(monthA1) {
  return 'Col1 >= date \'"&TEXT($' + monthA1 + ',"yyyy-mm-dd")&"\' and Col1 < date \'"&TEXT(EDATE($' + monthA1 + ',1),"yyyy-mm-dd")&"\'';
}

function masterMonthQ_(rangeA1, selectCols, monthA1) {
  return 'QUERY(' + rangeA1 + ',"select ' + selectCols + ' where ' + masterMonthWhere_(monthA1) + ' order by Col1 desc",0)';
}

function masterMonthIfChain_(sel, pairs, fallback) {
  var s = fallback;
  var i;
  for (i = pairs.length - 2; i >= 0; i -= 2) {
    s = 'IF(' + sel + '=' + pairs[i] + ',' + pairs[i + 1] + ',' + s + ')';
  }
  return s;
}

function masterMonthStackFormula_(monthA1) {
  var w = masterMonthWhere_(monthA1);
  var blank = '{"","","",""}';
  var q = function (rng, sel) {
    return 'IFERROR(QUERY(' + rng + ',"select ' + sel + ' where ' + w + '",0),' + blank + ')';
  };
  var parts = [
    q("'見学体験申請'!A2:D", "Col1,'見学体験',Col3,Col4"),
    q("'販促_紹介・ペア入会'!A3:G", "Col1,'紹介',Col6,Col7"),
    q("'学割'!A2:D", "Col1,'学割',Col3,Col4"),
    q("'販促_ラグビー割'!A3:E", "Col1,'ラグビー割',Col3,Col5"),
    q("'販促_乗り換え'!A3:D", "Col1,'乗り換え',Col3,Col4"),
    q("'販促_6ヶ月継続'!A3:D", "Col1,'6ヶ月',Col3,Col4"),
    q("'マシンレクチャー申込'!A3:D", "Col1,'レクチャー',Col2,Col4"),
    q("'口コミ_経堂'!A3:F", "Col1,'口コミ',Col5,Col6"),
    q("'入会者一覧＋自動メール管理'!A3:C", "Col1,'入会メール',Col2,Col3"),
    q("'経堂_入会'!A2:F", "Col1,'入会',Col2,Col6"),
    q("'経堂_退会'!A2:C", "Col1,'退会',Col2,Col3")
  ].join(',');
  return 'IFERROR(QUERY(VSTACK(' + parts + '),"select Col1,Col2,Col3,Col4 where Col1 is not null order by Col1 desc",0),"")';
}

function masterMonthQueryFormula_(monthA1, selA1) {
  var sel = '$' + selA1;
  var q = function (inner) {
    return 'IFERROR(' + inner + ',"")';
  };
  return '=' + masterMonthIfChain_(sel, [
    '"見学体験"', q(masterMonthQ_("'見学体験申請'!A2:J", 'Col1,Col2,Col3,Col4,Col5,Col8,Col9,Col10', monthA1)),
    '"紹介"', q(masterMonthQ_("'販促_紹介・ペア入会'!A3:G", 'Col1,Col2,Col3,Col5,Col6,Col7', monthA1)),
    '"学割"', q(masterMonthQ_("'学割'!A2:P", '*', monthA1)),
    '"ラグビー割"', q(masterMonthQ_("'販促_ラグビー割'!A3:F", 'Col1,Col3,Col4,Col5,Col6', monthA1)),
    '"乗り換え"', q(masterMonthQ_("'販促_乗り換え'!A3:E", 'Col1,Col3,Col4,Col5', monthA1)),
    '"6ヶ月"', q(masterMonthQ_("'販促_6ヶ月継続'!A3:D", 'Col1,Col3,Col4', monthA1)),
    '"レクチャー"', q(masterMonthQ_("'マシンレクチャー申込'!A3:G", 'Col1,Col2,Col3,Col4,Col5,Col6,Col7', monthA1)),
    '"口コミ"', q(masterMonthQ_("'口コミ_経堂'!A3:W", 'Col1,Col5,Col6,Col9,Col10,Col22', monthA1)),
    '"入会メール"', q(masterMonthQ_("'入会者一覧＋自動メール管理'!A3:E", 'Col1,Col2,Col3,Col4,Col5', monthA1)),
    '"入会"', q(masterMonthQ_("'経堂_入会'!A2:F", 'Col1,Col2,Col3,Col4,Col6', monthA1)),
    '"退会"', q(masterMonthQ_("'経堂_退会'!A2:D", 'Col1,Col2,Col3,Col4', monthA1))
  ], masterMonthStackFormula_(monthA1));
}

function masterMonthHeaderFormula_(selA1) {
  var sel = '$' + selA1;
  var school = '{"' + schoolDiscountHeaders_().join('","') + '"}';
  return '=' + masterMonthIfChain_(sel, [
    '"見学体験"', '{"日時","区分","名前","メール","電話","希望日","時刻","入会"}',
    '"紹介"', '{"日時","種別","紹介者","紹介者電話","被紹介者","被紹介者電話"}',
    '"学割"', school,
    '"ラグビー割"', '{"日時","名前","フリガナ","電話","メール"}',
    '"乗り換え"', '{"日時","名前","移籍元","地名"}',
    '"6ヶ月"', '{"日時","名前","同意"}',
    '"レクチャー"', '{"日時","名前","年齢","メール","希望","予約日","予約時間"}',
    '"口コミ"', '{"日時","氏名","会員番号","メール","来店日","付与"}',
    '"入会メール"', '{"入会日","名前","メール","アンケート","レクチャーメール"}',
    '"入会"', '{"日時","氏名","入会月","区分","メール"}',
    '"退会"', '{"日時","氏名","退会月","区分"}'
  ], '{"日時","種別","名前","連絡先"}');
}

function masterMonthLinkFormula_(ss, selA1) {
  var sel = '$' + selA1;
  var gid = function (name) {
    return hubSheetGidUrl_(ss, name);
  };
  return '=' + masterMonthIfChain_(sel, [
    '"学割"', 'HYPERLINK("' + gid('学割') + '","学割シートを開く ↗")',
    '"見学体験"', 'HYPERLINK("' + gid('見学体験申請') + '","見学体験申請を開く ↗")',
    '"紹介"', 'HYPERLINK("' + gid('販促_紹介・ペア入会') + '","紹介シートを開く ↗")',
    '"ラグビー割"', 'HYPERLINK("' + gid('販促_ラグビー割') + '","ラグビー割を開く ↗")',
    '"乗り換え"', 'HYPERLINK("' + gid('販促_乗り換え') + '","乗り換えを開く ↗")',
    '"6ヶ月"', 'HYPERLINK("' + gid('販促_6ヶ月継続') + '","6ヶ月継続を開く ↗")',
    '"レクチャー"', 'HYPERLINK("' + gid('マシンレクチャー申込') + '","レクチャーを開く ↗")',
    '"口コミ"', 'HYPERLINK("' + REVIEW_GRANT_APP_URL_ + '","口コミ付与アプリ")',
    '"入会メール"', 'HYPERLINK("' + gid('入会者一覧＋自動メール管理') + '","入会メールを開く ↗")',
    '"入会"', 'HYPERLINK("' + gid('経堂_入会') + '","入会シートを開く ↗")',
    '"退会"', 'HYPERLINK("' + gid('経堂_退会') + '","退会シートを開く ↗")'
  ], 'HYPERLINK("' + hubBookUrl_(ss.getId()) + '#gid=989823202","トップでシートを探す ↗")');
}

function masterMonthTabSelect_(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (!sh || sh.getName() !== '経堂マスタ') return;
  if (e.range.getRow() !== 2) return;
  var start = masterFindRowLabel_(sh, 1, '当月の申請');
  if (!start) return;
  var col = e.range.getColumn();
  var tabs = masterMonthTabs_();
  if (col < start || col > start + tabs.length) return;
  var raw = String(e.range.getDisplayValue() || '');
  var hit = '';
  var i;
  for (i = 0; i < tabs.length; i++) {
    if (raw === tabs[i]) hit = tabs[i];
  }
  if (!hit) return;
  var sel = sh.getRange(2, start);
  if (String(sel.getDisplayValue() || '') !== hit) sel.setValue(hit);
  masterPaintMonthTabs_(sh, start);
}

function masterPaintMonthTabs_(sh, start) {
  var t = dnTheme_();
  var tabs = masterMonthTabs_();
  var sel = String(sh.getRange(2, start).getDisplayValue() || 'すべて');
  var hair = SpreadsheetApp.BorderStyle.SOLID;
  var helperCol = 0;
  try { helperCol = sh.getRange(masterMonthCellA1_(sh)).getColumn(); } catch (eH) {}
  var lastChip = start + tabs.length;
  if (helperCol && lastChip >= helperCol) lastChip = helperCol - 1;
  var i;
  for (i = 0; i < tabs.length; i++) {
    var col = start + 1 + i;
    if (col > lastChip) break;
    var cell = sh.getRange(2, col);
    var on = tabs[i] === sel;
    cell.setValue(tabs[i]);
    hubType_(cell)
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBackground(on ? t.ink : t.cream)
      .setFontColor(on ? t.paper : t.ink)
      .setBorder(true, true, true, true, false, false, t.ash, hair);
  }
}

function setupMasterMonthPanel_(sh) {
  var kpiEnd = masterFindRowLabel_(sh, 4, '当月学割');
  if (!kpiEnd) kpiEnd = masterFindRowLabel_(sh, 4, '当月紹介');
  var start = kpiEnd ? kpiEnd + 2 : 16;
  var tabs = masterMonthTabs_();
  var need = Math.max(16, tabs.length + 2);
  var monthA1 = masterMonthCellA1_(sh);
  var helperCol = sh.getRange(monthA1).getColumn();
  if (helperCol > start && helperCol - start < need) {
    sh.insertColumnsBefore(helperCol, need - (helperCol - start));
    monthA1 = masterMonthCellA1_(sh);
    helperCol = sh.getRange(monthA1).getColumn();
  }
  var width = helperCol > start ? helperCol - start : need;
  try { sh.showColumns(start, Math.max(width, 1)); } catch (eShow) {}
  var prev = '';
  try { prev = String(sh.getRange(2, start).getDisplayValue() || ''); } catch (ePrev) {}
  try { sh.getRange(1, start, 3, width).breakApart(); } catch (e0) {}
  try { sh.getRange(1, start, sh.getMaxRows(), width).clearDataValidations(); } catch (e1) {}
  try { sh.getRange(1, start, 3, width).clearContent(); } catch (e2) {}
  try { sh.getRange(4, start, Math.min(Math.max(sh.getMaxRows() - 3, 1), 400), width).clearContent(); } catch (e3) {}
  var t = dnTheme_();
  var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  var selA1 = columnLetter_(start) + '2';
  var ss = sh.getParent();
  if (tabs.indexOf(prev) === -1) prev = 'すべて';
  sh.getRange(1, start, 1, Math.max(width - 1, 1)).merge()
    .setFormula('="当月の申請　"&TEXT($' + monthA1 + ',"yyyy年m月")&"　"&$' + selA1)
    .setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, t.blood, medium);
  hubType_(sh.getRange(1, start));
  sh.setRowHeight(1, 30);
  sh.getRange(1, start + width - 1)
    .setFormula(masterMonthLinkFormula_(ss, selA1))
    .setBackground(t.ash).setFontColor(t.paper).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  hubType_(sh.getRange(1, start + width - 1)).setFontSize(9);
  sh.getRange(2, start).setValue(prev);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(tabs, true).setAllowInvalid(false).build();
  sh.getRange(2, start).setDataValidation(rule);
  hubType_(sh.getRange(2, start))
    .setFontSize(9).setFontWeight('bold')
    .setBackground(t.cream).setFontColor(t.ink)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  masterPaintMonthTabs_(sh, start);
  sh.setRowHeight(2, 32);
  sh.getRange(3, start).setFormula(masterMonthHeaderFormula_(selA1));
  hubType_(sh.getRange(3, start, 1, width))
    .setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(9).setWrap(true);
  sh.setRowHeight(3, 28);
  sh.getRange(4, start).setFormula(masterMonthQueryFormula_(monthA1, selA1));
  var body = Math.min(Math.max(sh.getMaxRows() - 3, 1), 400);
  hubType_(sh.getRange(4, start, body, width))
    .setBackground(t.paper).setFontColor(t.ink)
    .setVerticalAlignment('middle');
  sh.getRange(4, start, body, 1).setNumberFormat('yyyy/mm/dd HH:mm').setHorizontalAlignment('center');
  var z;
  for (z = 1; z < Math.min(body, 80); z += 2) {
    sh.getRange(4 + z, start, 1, width).setBackground('#E6E6E6');
  }
  sh.setColumnWidth(start, 118);
  var c;
  for (c = 1; c < width; c++) {
    sh.setColumnWidth(start + c, c === 1 ? 92 : c === 2 ? 150 : c === width - 1 ? 140 : 108);
  }
  return { ok: true, start: start, width: width, monthCell: monthA1, tab: prev };
}

/** 学校関係者割フォーム（経堂だけ）を Workspace へ。元ブックは読み取りのみ */
function setupSchoolDiscountImport_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sheet = setupSchoolDiscountSheet_(ss);
    var permit = permitImportRange_(ss, SCHOOL_FORM_SOURCE_ID_);
    var master = ss.getSheetByName('経堂マスタ');
    var schoolKpi = master ? setupMasterSchoolKpi_() : { ok: false };
    var panel = master ? setupMasterSchoolPanel_(master) : { ok: false, message: '経堂マスタなし' };
    if (master) {
      restyleKyodoMasterLook_(master);
    }
    return {
      ok: true,
      sourceId: SCHOOL_FORM_SOURCE_ID_,
      sourceSheet: SCHOOL_FORM_SHEET_,
      destSheet: SCHOOL_WS_SHEET_,
      permit: permit,
      panel: panel,
      schoolKpi: schoolKpi,
      note: '元のフォーム回答ブックは未変更。B列が JOYFIT24 経堂 の行は件数制限なしで全部（今後の追加分も含む）。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function masterFindRowLabel_(sh, row, needle) {
  var last = Math.max(sh.getLastColumn(), 20);
  var vals = sh.getRange(row, 1, 1, last).getDisplayValues()[0];
  var i;
  for (i = 0; i < vals.length; i++) {
    if (String(vals[i] || '').indexOf(needle) !== -1) return i + 1;
  }
  return 0;
}

function masterMonthCellA1_(sh) {
  var hits = ['AB5', 'AD5', 'AF5', 'AT5', 'AZ5'];
  var i;
  for (i = 0; i < hits.length; i++) {
    var f = String(sh.getRange(hits[i]).getFormula() || '');
    if (/DATE\(VALUE\(LEFT\(\$B\$2|DATE\(YEAR\(TODAY/.test(f)) return hits[i];
  }
  var last = Math.max(sh.getLastColumn(), 28);
  var start = 20;
  var forms = sh.getRange(5, start, 1, last - start + 1).getFormulas()[0];
  for (i = 0; i < forms.length; i++) {
    if (/DATE\(VALUE\(LEFT\(\$B\$2/.test(String(forms[i] || ''))) return columnLetter_(start + i) + '5';
  }
  return 'AB5';
}

function masterPromoListFormula_(monthA1) {
  return '=IFERROR(QUERY({\'販促_乗り換え\'!A3:C;' +
    '{\'販促_紹介・ペア入会\'!A3:A,\'販促_紹介・ペア入会\'!B3:B,\'販促_紹介・ペア入会\'!F3:F};' +
    '{\'販促_学校関係者\'!A2:A,ARRAYFORMULA(IF(\'販促_学校関係者\'!A2:A="","","学校関係者割")),\'販促_学校関係者\'!C2:C};' +
    '\'販促_ラグビー割\'!A3:C;\'販促_6ヶ月継続\'!A3:C},' +
    '"select Col1,Col2,Col3 where Col1 is not null order by Col1 desc",0),"")';
}

function masterSchoolCountFormula_(monthA1) {
  return '=IFERROR(COUNTIFS(\'' + SCHOOL_WS_SHEET_ + '\'!A2:A,">="&$' + monthA1 +
    ',\'' + SCHOOL_WS_SHEET_ + '\'!A2:A,"<"&EDATE($' + monthA1 + ',1)),)';
}

/** 当月紹介の右に、学校関係者割の当月件数。2列カードのまま足す */
function setupMasterSchoolKpi_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName('経堂マスタ');
    if (!sh) return { ok: false, message: '経堂マスタなし' };
    var introCol = masterFindRowLabel_(sh, 4, '当月紹介');
    var schoolKpiCol = masterFindRowLabel_(sh, 4, '当月学割');
    var inserted = false;
    if (!schoolKpiCol) {
      if (!introCol) return { ok: false, message: '当月紹介の位置が見つかりません' };
      sh.insertColumnsAfter(introCol + 1, 2);
      inserted = true;
      schoolKpiCol = introCol + 2;
    }
    var monthA1 = masterMonthCellA1_(sh);
    try { sh.getRange(4, schoolKpiCol, 3, 2).breakApart(); } catch (eBr) {}
    if (introCol) {
      try {
        sh.getRange(4, introCol, 3, 2).copyTo(
          sh.getRange(4, schoolKpiCol, 3, 2),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
      } catch (eFmt) {}
    }
    sh.getRange(4, schoolKpiCol, 1, 2).merge().setValue('当月学割')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sh.getRange(5, schoolKpiCol, 2, 2).merge()
      .setFormula(masterSchoolCountFormula_(monthA1))
      .setNumberFormat('0" 件"')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontWeight('bold');
    if (introCol) {
      try {
        sh.setColumnWidth(schoolKpiCol, sh.getColumnWidth(introCol));
        sh.setColumnWidth(schoolKpiCol + 1, sh.getColumnWidth(introCol + 1));
      } catch (eW) {}
    }
    var titleEnd = schoolKpiCol + 1;
    var title = String(sh.getRange(1, 1).getDisplayValue() || '');
    if (!title || /元のシート/.test(title)) title = 'JOYFIT24経堂マスタ';
    try { sh.getRange(1, 1, 1, titleEnd).breakApart(); } catch (eT) {}
    sh.getRange(1, 1, 1, titleEnd).merge().setValue(title)
      .setHorizontalAlignment('center').setFontWeight('bold');
    return {
      ok: true,
      inserted: inserted,
      schoolKpiCol: schoolKpiCol,
      monthCell: monthA1
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function stripMasterSideConditionalFormats_(sh) {
  var reviewCol = masterFindRowLabel_(sh, 1, '当月の申請');
  if (!reviewCol) return;
  var rules = sh.getConditionalFormatRules();
  var kept = [];
  var i;
  for (i = 0; i < rules.length; i++) {
    var ranges = rules[i].getRanges();
    var clipped = [];
    var g;
    for (g = 0; g < ranges.length; g++) {
      var rng = ranges[g];
      var c1 = rng.getColumn();
      var c2 = c1 + rng.getNumColumns() - 1;
      if (c2 < reviewCol) {
        clipped.push(rng);
      } else if (c1 < reviewCol) {
        clipped.push(sh.getRange(rng.getRow(), c1, rng.getNumRows(), reviewCol - c1));
      }
    }
    if (clipped.length) {
      try { kept.push(rules[i].copy().setRanges(clipped).build()); } catch (eC) {}
    }
  }
  sh.setConditionalFormatRules(kept);
}

function masterIntroCountFormula_(monthA1) {
  return '=IFERROR(COUNTIFS(\'販促_紹介・ペア入会\'!A3:A,">="&$' + monthA1 +
    ',\'販促_紹介・ペア入会\'!A3:A,"<"&EDATE($' + monthA1 + ',1)),)';
}

/** 経堂マスタの当月移籍の右に、紹介・ペア入会の当月件数を出す */
function setupMasterIntroKpi_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName('経堂マスタ');
    if (!sh) return { ok: false, message: '経堂マスタなし' };
    if (!ss.getSheetByName('販促_紹介・ペア入会')) {
      return { ok: false, message: '販促_紹介・ペア入会 がありません' };
    }
    var moveCol = masterFindRowLabel_(sh, 4, '当月移籍');
    var introCol = masterFindRowLabel_(sh, 4, '当月紹介');
    var inserted = false;
    if (!introCol) {
      if (!moveCol) return { ok: false, message: '当月移籍の位置が見つかりません' };
      sh.insertColumnsAfter(moveCol + 1, 2);
      inserted = true;
      introCol = moveCol + 2;
    }
    var monthA1 = masterMonthCellA1_(sh);
    try { sh.getRange(4, introCol, 3, 2).breakApart(); } catch (eBr) {}
    if (moveCol) {
      try {
        sh.getRange(4, moveCol, 3, 2).copyTo(
          sh.getRange(4, introCol, 3, 2),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
      } catch (eFmt) {}
    }
    sh.getRange(4, introCol, 1, 2).merge().setValue('当月紹介')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sh.getRange(5, introCol, 2, 2).merge()
      .setFormula(masterIntroCountFormula_(monthA1))
      .setNumberFormat('0" 件"')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontWeight('bold');
    if (moveCol) {
      try {
        sh.setColumnWidth(introCol, sh.getColumnWidth(moveCol));
        sh.setColumnWidth(introCol + 1, sh.getColumnWidth(moveCol + 1));
      } catch (eW) {}
    }
    var titleEnd = introCol + 1;
    var title = String(sh.getRange(1, 1).getDisplayValue() || '');
    if (!title || /元のシート/.test(title)) title = 'JOYFIT24経堂マスタ';
    try { sh.getRange(1, 1, 1, titleEnd).breakApart(); } catch (eT) {}
    sh.getRange(1, 1, 1, titleEnd).merge().setValue(title)
      .setHorizontalAlignment('center').setFontWeight('bold');

    restyleKyodoMasterLook_(sh);
    SpreadsheetApp.flush();
    return {
      ok: true,
      inserted: inserted,
      introCol: introCol,
      monthCell: monthA1,
      count: sh.getRange(5, introCol).getDisplayValue()
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function restyleKyodoMasterLook_(sh) {
  var t = dnTheme_();
  var aLast = Math.min(Math.max(sh.getLastRow(), 24), 80);
  var labels = sh.getRange(1, 1, aLast, 1).getDisplayValues();
  var headerRow = 16;
  var last = 24;
  var i;
  for (i = 0; i < labels.length; i++) {
    var lab = String(labels[i][0] || '');
    if (lab === '項目') headerRow = i + 1;
    if (lab) last = i + 1;
  }
  last = Math.max(last, 24);
  var start = headerRow + 1;
  var bodyRows = Math.max(last - headerRow, 1);
  var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  var introCol = masterFindRowLabel_(sh, 4, '当月紹介');
  var schoolKpiCol = masterFindRowLabel_(sh, 4, '当月学割');
  var kpiCols = schoolKpiCol ? schoolKpiCol + 1 : (introCol ? introCol + 1 : 11);
  try { stripMasterSideConditionalFormats_(sh); } catch (eCf) {}
  hubType_(sh.getRange(1, 1, last, kpiCols))
    .setFontColor(t.ink)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, kpiCols).setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setBorder(false, false, true, false, false, false, t.blood, medium);
  sh.getRange(2, 1, Math.max(headerRow - 2, 1), kpiCols).setBackground(t.paper);
  sh.getRange(headerRow, 1, 1, Math.min(kpiCols, 10)).setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(false, false, true, false, false, false, t.blood, medium);
  sh.getRange(start, 1, bodyRows, Math.min(kpiCols, 10)).setBackground(t.ghost).setFontColor(t.ink);
  sh.getRange(start, 1, bodyRows, 1).setBackground(t.cream);
  sh.getRange(start, 6, bodyRows, 1).setBackground(t.cream);
  sh.getRange(start, 7, bodyRows, 1).setBackground(t.cream);
  try {
    sh.getRange('F7:K14').setBackground(t.ghost).setFontColor(t.ink);
    sh.getRange('F7').setBackground(t.blood).setFontColor(t.paper);
  } catch (eBrief) {}
  try {
    var sideStart = masterFindRowLabel_(sh, 1, '当月の申請');
    if (sideStart) {
      try { masterPaintMonthTabs_(sh, sideStart); } catch (eTab) {}
      var helperCol = sh.getRange(masterMonthCellA1_(sh)).getColumn();
      var sideW = helperCol > sideStart ? helperCol - sideStart : 16;
      try { sh.showColumns(sideStart, Math.max(sideW, 1)); } catch (eShowSide) {}
      sh.getRange(1, sideStart, 1, Math.max(sideW - 1, 1)).setBackground(t.ink).setFontColor(t.paper);
      sh.getRange(3, sideStart, 1, sideW).setBackground(t.ink).setFontColor(t.paper);
      var body = Math.min(Math.max(Math.max(sh.getLastRow(), 24) - 3, 1), 400);
      hubType_(sh.getRange(4, sideStart, body, sideW))
        .setBackground(t.paper)
        .setFontColor(t.ink)
        .setVerticalAlignment('middle');
      var z;
      for (z = 1; z < Math.min(body, 80); z += 2) {
        sh.getRange(4 + z, sideStart, 1, sideW).setBackground('#E6E6E6');
      }
    }
  } catch (eSide) {}
  if (introCol) {
    try {
      hubType_(sh.getRange(4, introCol, 1, 2)).setFontWeight('bold').setFontSize(11);
      hubType_(sh.getRange(5, introCol, 2, 2)).setFontWeight('bold').setFontSize(18)
        .setNumberFormat('0" 件"');
    } catch (eIntroStyle) {}
  }
  if (schoolKpiCol) {
    try {
      hubType_(sh.getRange(4, schoolKpiCol, 1, 2)).setFontWeight('bold').setFontSize(11);
      hubType_(sh.getRange(5, schoolKpiCol, 2, 2)).setFontWeight('bold').setFontSize(18)
        .setNumberFormat('0" 件"');
    } catch (eSchoolStyle) {}
  }
  hubStampType_(sh, last + 40, kpiCols);
  sh.setTabColor(t.ink);
  try { masterEnsureKpiCharts_(sh); } catch (eCh) {}
}

function masterFindALabel_(sh, needle) {
  var last = Math.min(Math.max(sh.getLastRow(), 24), 80);
  var vals = sh.getRange(1, 1, last, 1).getDisplayValues();
  var i;
  for (i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '') === needle) return i + 1;
  }
  return 0;
}

function masterClearKpiCharts_(sh) {
  var charts = sh.getCharts();
  var i;
  for (i = charts.length - 1; i >= 0; i--) {
    try {
      var info = charts[i].getContainerInfo();
      if (info.getAnchorColumn() >= 11 && info.getAnchorRow() >= 15) sh.removeChart(charts[i]);
    } catch (e0) {}
  }
}

function masterWriteKpiChartSource_(sh) {
  var head = masterFindALabel_(sh, '項目') || 16;
  var blocks = [
    { key: 'join', plan: masterFindALabel_(sh, '入会計画'), act: masterFindALabel_(sh, '入会実績') },
    { key: 'leave', plan: masterFindALabel_(sh, '解除計画'), act: masterFindALabel_(sh, '解除実績') },
    { key: 'start', plan: masterFindALabel_(sh, '月初計画'), act: masterFindALabel_(sh, '月初実績') },
    { key: 'end', plan: masterFindALabel_(sh, '月末計画'), act: masterFindALabel_(sh, '月末実績') }
  ];
  var startCol = 35;
  var need = startCol + 15;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  try { sh.showColumns(startCol, 15); } catch (eShow) {}
  sh.getRange(1, startCol, 8, 15).clearContent();
  var t = dnTheme_();
  var out = [];
  var b;
  for (b = 0; b < blocks.length; b++) {
    var col = startCol + b * 4;
    var plan = blocks[b].plan;
    var act = blocks[b].act;
    if (!plan || !act) {
      out.push({ key: blocks[b].key, ok: false });
      continue;
    }
    sh.getRange(1, col, 1, 3).setValues([['月', '計画', '実績']]);
    var r;
    var forms = [];
    for (r = 0; r < 5; r++) {
      var srcCol = columnLetter_(2 + r);
      forms.push([
        '=' + srcCol + head,
        '=' + srcCol + plan,
        '=' + srcCol + act
      ]);
    }
    sh.getRange(2, col, 5, 3).setFormulas(forms);
    sh.getRange(2, col + 1, 5, 2).setNumberFormat('0');
    out.push({ key: blocks[b].key, ok: true, col: col, plan: plan, act: act });
  }
  sh.getRange(1, startCol, 6, 15).setFontColor(t.paper).setBackground(t.paper).setFontSize(8);
  try { sh.setColumnWidths(startCol, 15, 10); } catch (eW) {}
  return { ok: true, head: head, startCol: startCol, blocks: out };
}

function masterReplaceMomSpark_(sh, head) {
  var t = dnTheme_();
  var last = Math.min(Math.max(sh.getLastRow(), head + 8), 80);
  var labels = sh.getRange(head, 1, last - head + 1, 1).getDisplayValues();
  sh.getRange(head, 10).setValue('5ヶ月').setFontWeight('bold').setHorizontalAlignment('center');
  var i;
  for (i = 1; i < labels.length; i++) {
    var name = String(labels[i][0] || '');
    if (!name) continue;
    var row = head + i;
    var color = /解除実績|退会実績/.test(name) ? t.blood : t.ink;
    sh.getRange(row, 10).setFormula(
      '=IFERROR(SPARKLINE(B' + row + ':F' + row + ',{"charttype","line";"color","' + color + '";"linewidth",2}),)'
    );
  }
}

function masterEnsureKpiCharts_(sh) {
  if (!sh) return { ok: false };
  masterClearKpiCharts_(sh);
  var src = masterWriteKpiChartSource_(sh);
  var head = src.head || 16;
  try { masterReplaceMomSpark_(sh, head); } catch (eSp) {}
  var t = dnTheme_();
  var start = src.startCol || 35;
  var chartTop = head;
  var gap = 12;
  var titles = [
    { title: '入会 計画と実績', colors: [t.ash, t.ink] },
    { title: '退会 計画と実績', colors: [t.ash, t.blood] },
    { title: '月初 計画と実績', colors: [t.ash, t.ink] },
    { title: '月末 計画と実績', colors: [t.ash, t.ink] }
  ];
  SpreadsheetApp.flush();
  var placed = 0;
  var errors = [];
  var i;
  for (i = 0; i < titles.length; i++) {
    var col = start + i * 4;
    var r = unpaidInsertTrendChart_(sh, {
      kind: 'column',
      range: sh.getRange(1, col, 6, 3),
      title: titles[i].title,
      row: chartTop + i * gap,
      col: 11,
      width: 460,
      height: 230,
      colors: titles[i].colors
    });
    if (r && r.ok) placed += 1;
    else errors.push(titles[i].title + ': ' + ((r && r.message) || 'fail'));
  }
  return { ok: placed > 0, charts: placed, errors: errors, head: head, src: src };
}

function setupMasterKpiCharts_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName('経堂マスタ');
    if (!sh) return { ok: false, message: '経堂マスタなし' };
    var out = masterEnsureKpiCharts_(sh);
    return {
      ok: !!(out && out.ok),
      charts: out && out.charts,
      errors: out && out.errors,
      head: out && out.head,
      j16: String(sh.getRange('J16').getDisplayValue() || ''),
      a17: String(sh.getRange('A17').getDisplayValue() || ''),
      ai1: String(sh.getRange(1, 35).getDisplayValue() || ''),
      ai2: String(sh.getRange(2, 35).getDisplayValue() || '')
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function restyleUnpaidTrendLook_(tr) {
  try { unpaidStyleTrendSheet_(tr); } catch (e0) {}
  try { unpaidEnsureTrendCharts_(tr); } catch (e1) {}
}

function restyleUrlIndexLook_(sh) {
  var t = dnTheme_();
  var last = Math.max(sh.getLastRow(), 1);
  var cols = Math.min(Math.max(sh.getLastColumn(), 4), 4);
  sh.getRange(1, 1, 1, cols)
    .setBackground(t.ink)
    .setFontColor(t.paper)
    .setFontFamily('Noto Sans JP').setFontStyle('italic')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(false, false, true, false, false, false, t.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, cols)
      .setBackground(t.paper)
      .setFontColor(t.ink)
      .setFontFamily('Noto Sans JP').setFontStyle('italic');
    var i;
    for (i = 0; i < last - 1; i++) {
      if (i % 2 === 1) sh.getRange(i + 2, 1, 1, cols).setBackground(t.cream);
    }
    sh.getRange(2, 1, last - 1, 1).setFontColor(t.ash);
    sh.getRange(2, 2, last - 1, 1).setFontColor(t.ash);
    if (cols >= 4) sh.getRange(2, 4, last - 1, 1).setFontColor(t.blood);
  }
}

function restyleHubLook_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var unpaid = ss.getSheetByName(UNPAID_SHEET_);
    var trend = ss.getSheetByName(UNPAID_TREND_SHEET_);
    if (unpaid) {
      applyUnpaidScanLook_(unpaid);
      if (trend) applyUnpaidNotes_(unpaid, trend);
    }
    if (trend) restyleUnpaidTrendLook_(trend);
    var master = ss.getSheetByName('経堂マスタ');
    if (master) {
      try { setupMasterIntroKpi_(); } catch (eIntro) {}
      try { setupSchoolDiscountImport_(); } catch (eSch) {}
      restyleKyodoMasterLook_(master);
    }
    var urlSh = ss.getSheetByName('URL一覧');
    if (urlSh) restyleUrlIndexLook_(urlSh);
    if (master) hubEnsureSourceBanner_(master);
    if (trend) hubEnsureSourceBanner_(trend);
    if (master) hubStampType_(master);
    if (unpaid) hubStampType_(unpaid, Math.max(unpaid.getLastRow(), 80) + 80, 16);
    if (trend) hubStampType_(trend);
    var home = ss.getSheetByName(HUB_HOME_SHEET_);
    if (home) {
      restyleHubHomeLook_(home);
      hubStampType_(home, 40, 8);
    }
    var sheets = ss.getSheets();
    var i;
    for (i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var n = sh.getName();
      if (n === HUB_HOME_SHEET_ || n === '経堂マスタ' || n === UNPAID_SHEET_ || n === UNPAID_TREND_SHEET_ || n === 'URL一覧' || n === '学割' || n === '販促_学校関係者') continue;
      if (/backup|シート\d+/.test(n)) continue;
      if (n === JOIN_LIST_SHEET_) {
        try { joinListFlushQueue_(ss); } catch (eJQ) {}
        try { joinListPushChecks_(ss); } catch (eJP) {}
        try { joinListRefresh_(sh); } catch (eJ) {}
        continue;
      }
      if (n === '見学体験申請') {
        try { repairKengakuMirror_(); } catch (eK) {}
        continue;
      }
      if (/^販促_/.test(n) || n === 'マシンレクチャー申込' ||
          n === '口コミ_経堂' || n === '【経堂】会員動向' ||
          /^経堂_/.test(n) || n === 'Tasks' || n === 'WorkspaceSync') {
        try { restyleHeaderBody_(sh); } catch (e1) {}
        try { hubEnsureSourceBanner_(sh); } catch (e2) {}
      }
    }
    applyHubTabColors_(ss);
    var tabs = ss.getSheets().map(function (s) {
      return { name: s.getName(), tab: hubTabColorFor_(s.getName()), hidden: s.isSheetHidden() };
    });
    return { ok: true, tabs: tabs };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/** 25年8月〜27年12月（古い順。推移シート用） */
function unpaidMonthListChrono_() {
  var out = [];
  var y;
  var m;
  for (y = 2025; y <= 2027; y++) {
    for (m = 1; m <= 12; m++) {
      if (y === 2025 && m < 8) continue;
      out.push(String(y % 100) + '年' + m + '月');
    }
  }
  return out;
}

/** ドロップダウン用。いまの月を一番上、あとは新しい順 */
function unpaidMonthOptions_() {
  var now = new Date();
  var cur = Utilities.formatDate(now, 'Asia/Tokyo', 'yy') + '年' +
    Number(Utilities.formatDate(now, 'Asia/Tokyo', 'M')) + '月';
  var items = unpaidMonthListChrono_().slice().reverse();
  var rest = items.filter(function (x) { return x !== cur; });
  return items.indexOf(cur) >= 0 ? [cur].concat(rest) : items;
}

/**
 * 表示用に変換した IMPORTRANGE。元データは変えず、表示だけ
 * TRUE/FALSE → ☑/☐、金額列（見出しに 額・当月分・手数料・繰越・支払 を含む列）→ ¥#,##0、
 * A列の区分（退会後未納貸倒候補・2ヶ月未納など）を下の行まで埋める。
 * 会員名の右に 入会日・入会区分・未納開始・入会→未納 の4列を差し込む（経堂_入会 を氏名で照合。
 * 同名が複数なら未納開始月末までで最新の入会）。
 */
function unpaidViewFormula_() {
  return '=IMPORTRANGE("' + UNPAID_SOURCE_ID_ + '","\'"&B1&"\'!A1:AK")';
}

var UNPAID_TREND_SHEET_ = '未納管理_推移';
var UNPAID_DATA_ROW_ = 5;
var UNPAID_COLS_ = 41;
var UNPAID_NAME_COL_ = 4;
var UNPAID_PAY_COL_ = 10;
var UNPAID_REC_COL_ = 21;

function unpaidShortCatExpr_(cx) {
  return 'IFS(' +
    'REGEXMATCH(' + cx + ',"貸倒|貸し倒"),"貸倒",' +
    'REGEXMATCH(' + cx + ',"JACCS"),"JACCS",' +
    'REGEXMATCH(' + cx + ',"1ヶ月|1ヵ月|1カ月|1か月"),"1ヶ月",' +
    'REGEXMATCH(' + cx + ',"2ヶ月|2ヵ月|2カ月|2か月"),"2ヶ月",' +
    'REGEXMATCH(' + cx + ',"過年度"),"過年度",' +
    'TRUE,' + cx + ')';
}

function unpaidParseAmount_(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  var s = String(v == null ? '' : v).replace(/[¥,\s]/g, '');
  var n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function unpaidStatsFromValues_(vals) {
  var out = { n: 0, sp: 0, st: 0, sr: 0, nr: 0, kone: { pay: 0, rec: 0 }, ktwo: { pay: 0, rec: 0 }, kbad: { pay: 0, rec: 0 }, kjac: { pay: 0, rec: 0 } };
  if (!vals || vals.length < 3) return out;
  var cols = vals[0].length;
  var h = [];
  var c;
  for (c = 0; c < cols; c++) {
    h.push(String(vals[0][c] || '') + String(vals[1][c] || '') + String(vals[2][c] || ''));
  }
  var fc = function (re) {
    var rx = new RegExp(re);
    for (c = 0; c < h.length; c++) {
      if (rx.test(h[c])) return c;
    }
    return -1;
  };
  var mc = fc('会員番号');
  var pc = fc('支払額');
  var tc = fc('総額');
  var uc = fc('回収金額|入金金額|レジ打ち金額');
  var cat = '';
  var r;
  for (r = 0; r < vals.length; r++) {
    var row = vals[r];
    if (row[0]) cat = String(row[0]);
    var mem = mc < 0 ? '' : String(row[mc] == null ? '' : row[mc]).trim();
    var pay = pc < 0 ? 0 : unpaidParseAmount_(row[pc]);
    var tot = tc < 0 ? 0 : unpaidParseAmount_(row[tc]);
    var rec = uc < 0 ? 0 : unpaidParseAmount_(row[uc]);
    if (uc >= 0 && rec === 0) {
      var neigh = String(row[uc + 1] || '') + String(row[uc + 2] || '') + String(row[uc + 3] || '');
      if (/入金/.test(neigh)) rec = pay;
    }
    if (!mem || mem === '会員番号' || mem === '合計' || pay <= 0) continue;
    out.n += 1;
    out.sp += pay;
    out.st += tot;
    out.sr += rec;
    if (rec > 0) out.nr += 1;
    var bump = function (k) {
      k.pay += pay;
      k.rec += rec;
    };
    if (/1ヶ月|1ヵ月|1カ月|1か月/.test(cat)) bump(out.kone);
    else if (/2ヶ月|2ヵ月|2カ月|2か月/.test(cat)) bump(out.ktwo);
    if (/貸倒|貸し倒/.test(cat)) bump(out.kbad);
    if (/JACCS/.test(cat)) bump(out.kjac);
  }
  return out;
}

function unpaidYen_(n) {
  return '¥' + Math.round(Number(n) || 0).toLocaleString('en-US');
}

function unpaidRateText_(pay, rec) {
  if (!pay) return '対象なし';
  return (Math.round((rec / pay) * 1000) / 10).toFixed(1) + '%';
}

function unpaidWriteDashboardValues_(sh, st) {
  sh.getRange('D1:M3').setValues([
    ['未納件数', '未納総額', '回収額', '回収率', '未回収額', '回収済み', '1ヶ月未納 回収率', '2ヶ月未納 回収率', '貸倒候補 回収率', 'JACCS 回収率'],
    [
      st.n + '件',
      unpaidYen_(st.sp),
      unpaidYen_(st.sr),
      st.sp ? unpaidRateText_(st.sp, st.sr) : '-',
      unpaidYen_(st.sp - st.sr),
      st.nr + '件',
      unpaidRateText_(st.kone.pay, st.kone.rec),
      unpaidRateText_(st.ktwo.pay, st.ktwo.rec),
      unpaidRateText_(st.kbad.pay, st.kbad.rec),
      unpaidRateText_(st.kjac.pay, st.kjac.rec)
    ],
    [
      '支払額ベース',
      '手数料込 ' + unpaidYen_(st.st),
      '回収金額の合計',
      '回収額÷未納総額',
      '残り ' + (st.n - st.nr) + '件',
      st.n ? Math.round((st.nr / st.n) * 100) + '%（人数）' : '',
      st.kone.pay ? unpaidYen_(st.kone.rec) + ' / ' + unpaidYen_(st.kone.pay) : '対象なし',
      st.ktwo.pay ? unpaidYen_(st.ktwo.rec) + ' / ' + unpaidYen_(st.ktwo.pay) : '対象なし',
      st.kbad.pay ? unpaidYen_(st.kbad.rec) + ' / ' + unpaidYen_(st.kbad.pay) : '対象なし',
      st.kjac.pay ? unpaidYen_(st.kjac.rec) + ' / ' + unpaidYen_(st.kjac.pay) : '対象なし'
    ]
  ]);
}

function unpaidShortCat_(s) {
  var t = String(s || '');
  if (/貸倒|貸し倒/.test(t)) return '貸倒';
  if (/JACCS/.test(t)) return 'JACCS';
  if (/1ヶ月|1ヵ月|1カ月|1か月/.test(t)) return '1ヶ月';
  if (/2ヶ月|2ヵ月|2カ月|2か月/.test(t)) return '2ヶ月';
  if (/過年度/.test(t)) return '過年度';
  return t;
}

function unpaidPrepareDisplay_(vals) {
  if (!vals || !vals.length) return vals;
  var out = [];
  var cat = '';
  var r;
  var c;
  for (r = 0; r < vals.length; r++) {
    var row = vals[r].slice();
    if (row[0]) cat = String(row[0]);
    else if (r >= 3 && cat) {
      var hasData = false;
      for (c = 1; c < row.length; c++) {
        if (row[c] !== '' && row[c] != null) {
          hasData = true;
          break;
        }
      }
      if (hasData) row[0] = unpaidShortCat_(cat);
    }
    if (r >= 3 && row[0]) row[0] = unpaidShortCat_(row[0]);
    out.push(row);
  }
  return out;
}

function unpaidDetectBoolCols_(vals) {
  var cols = [];
  if (!vals || vals.length < 4) return cols;
  var width = vals[0].length;
  var c;
  var r;
  for (c = 0; c < width; c++) {
    var nBool = 0;
    for (r = 3; r < vals.length; r++) {
      if (typeof vals[r][c] === 'boolean') nBool += 1;
    }
    if (nBool > 0) cols.push(c);
  }
  return cols;
}

function unpaidIsPushCol_(col1, boolCols) {
  if (col1 === UNPAID_REC_COL_) return true;
  if (col1 === UNPAID_NAME_COL_) return true;
  return boolCols.indexOf(col1 - 1) >= 0;
}

function unpaidApplyCheckboxValidations_(destSh, vals) {
  var boolCols = unpaidDetectBoolCols_(vals);
  var r0 = UNPAID_DATA_ROW_ + 2;
  var n = Math.max(vals.length - 2, 1);
  if (r0 + n - 1 > destSh.getMaxRows()) n = destSh.getMaxRows() - r0 + 1;
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
  var i;
  for (i = 0; i < boolCols.length; i++) {
    destSh.getRange(r0, boolCols[i] + 1, n, 1).setDataValidation(rule);
  }
  try {
    PropertiesService.getDocumentProperties().setProperty('unpaidBoolCols', JSON.stringify(boolCols));
  } catch (eP) {}
  return boolCols;
}

function unpaidBoolColsStored_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty('unpaidBoolCols');
    if (raw) return JSON.parse(raw);
  } catch (eP) {}
  return [13, 14, 15, 16, 17, 18, 19];
}

var UNPAID_QUEUE_SHEET_ = '未納_同期';

function unpaidQueueSheet_(ss) {
  ss = ss || openWorkspaceSpreadsheet_();
  var sh = ss.getSheetByName(UNPAID_QUEUE_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(UNPAID_QUEUE_SHEET_);
    sh.hideSheet();
    sh.getRange(1, 1, 1, 7).setValues([['ts', 'action', 'month', 'srcRow', 'srcCol', 'value', 'status']]);
    sh.setTabColor(dnTheme_().ash);
  }
  try { sh.hideSheet(); } catch (eH) {}
  return sh;
}

function unpaidOnEditSimple_(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (!sh || sh.getName() !== UNPAID_SHEET_) return;
  if (e.range.getRow() === 1 && e.range.getColumn() === 2) return;
  if (e.range.getRow() < UNPAID_DATA_ROW_ + 2) return;
  unpaidQueueRange_(sh, e.range);
}

function unpaidQueueRange_(sh, range) {
  var month = String(sh.getRange('B1').getDisplayValue() || '').trim();
  if (!month) return;
  var boolCols = unpaidBoolColsStored_();
  var vals = range.getValues();
  var q = unpaidQueueSheet_(sh.getParent());
  var last = Math.max(q.getLastRow(), 1);
  var rows = [];
  var r;
  var c;
  for (r = 0; r < vals.length; r++) {
    for (c = 0; c < vals[r].length; c++) {
      var destR = range.getRow() + r;
      var destC = range.getColumn() + c;
      if (destR < UNPAID_DATA_ROW_ + 2) continue;
      if (!unpaidIsPushCol_(destC, boolCols)) continue;
      rows.push([
        new Date(),
        'push',
        month,
        destR - UNPAID_DATA_ROW_ + 1,
        destC,
        JSON.stringify(vals[r][c]),
        'queued'
      ]);
    }
  }
  if (!rows.length) return;
  q.getRange(last + 1, 1, rows.length, 7).setValues(rows);
}

function unpaidFlushQueue_(ss) {
  ss = ss || openWorkspaceSpreadsheet_();
  var q = ss.getSheetByName(UNPAID_QUEUE_SHEET_);
  if (!q || q.getLastRow() < 2) return { ok: true, pushed: 0, pending: 0 };
  var n = q.getLastRow() - 1;
  var data = q.getRange(2, 1, n, 7).getValues();
  var pending = [];
  var i;
  for (i = 0; i < data.length; i++) {
    if (String(data[i][6] || '') === 'queued' && String(data[i][1] || '') !== 'join') {
      pending.push({ i: i, row: data[i] });
    }
  }
  if (!pending.length) return { ok: true, pushed: 0, pending: 0 };
  var byMonth = {};
  for (i = 0; i < pending.length; i++) {
    var m = String(pending[i].row[2] || '');
    if (!byMonth[m]) byMonth[m] = [];
    byMonth[m].push(pending[i]);
  }
  var src = SpreadsheetApp.openById(UNPAID_SOURCE_ID_);
  var pushed = 0;
  var errors = 0;
  var monthName;
  for (monthName in byMonth) {
    if (!byMonth.hasOwnProperty(monthName) || !monthName) continue;
    var srcSh = src.getSheetByName(monthName);
    var items = byMonth[monthName];
    var j;
    for (j = 0; j < items.length; j++) {
      var it = items[j];
      var srcRow = Number(it.row[3]);
      var srcCol = Number(it.row[4]);
      var val = it.row[5];
      try { val = JSON.parse(val); } catch (eJ) {}
      try {
        if (!srcSh || srcRow < 3 || srcCol < 1) throw new Error('bad target');
        srcSh.getRange(srcRow, srcCol).setValue(val);
        q.getRange(it.i + 2, 7).setValue('done');
        pushed += 1;
      } catch (eW) {
        q.getRange(it.i + 2, 7).setValue('error: ' + String(eW && eW.message ? eW.message : eW));
        errors += 1;
      }
    }
  }
  SpreadsheetApp.flush();
  return { ok: errors === 0, pushed: pushed, errors: errors, pending: pending.length - pushed };
}

function unpaidReadMonthPack_(src, monthName) {
  var srcSh = src.getSheetByName(monthName);
  if (!srcSh) return null;
  var lastR = Math.min(Math.max(srcSh.getLastRow(), 3), 400);
  var lastC = Math.min(Math.max(srcSh.getLastColumn(), 1), 37);
  return {
    sh: srcSh,
    vals: srcSh.getRange(1, 1, lastR, lastC).getValues(),
    lastR: lastR,
    lastC: lastC
  };
}

function unpaidReadMonthValues_(src, monthName) {
  var pack = unpaidReadMonthPack_(src, monthName);
  return pack ? pack.vals : null;
}

function unpaidFillFromSource_(destSh) {
  var monthName = String(destSh.getRange('B1').getDisplayValue() || '').trim();
  if (!monthName) return { ok: false, message: 'B1 empty' };
  var flushed = { ok: true, pushed: 0 };
  try { flushed = unpaidFlushQueue_(destSh.getParent()); } catch (eQ) { flushed = { ok: false, message: String(eQ) }; }
  var src = SpreadsheetApp.openById(UNPAID_SOURCE_ID_);
  var pack = unpaidReadMonthPack_(src, monthName);
  if (!pack) return { ok: false, message: 'no sheet ' + monthName, flushed: flushed };
  var vals = pack.vals;
  var shown = unpaidPrepareDisplay_(vals);
  destSh.getRange(UNPAID_DATA_ROW_, 1).clearContent();
  destSh.getRange(UNPAID_DATA_ROW_, 1, shown.length, shown[0].length).setValues(shown);
  try {
    pack.sh.getRange(1, 1, pack.lastR, pack.lastC).copyTo(
      destSh.getRange(UNPAID_DATA_ROW_, 1),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );
  } catch (eVal) {}
  var boolCols = unpaidApplyCheckboxValidations_(destSh, vals);
  try {
    destSh.getRange(UNPAID_DATA_ROW_, UNPAID_PAY_COL_, shown.length, 1).setNumberFormat('¥#,##0');
    destSh.getRange(UNPAID_DATA_ROW_, UNPAID_REC_COL_, shown.length, 1).setNumberFormat('¥#,##0');
  } catch (eFmt) {}
  var tailStart = UNPAID_DATA_ROW_ + shown.length;
  var last = destSh.getLastRow();
  if (last >= tailStart) {
    destSh.getRange(tailStart, 1, last - tailStart + 1, Math.min(UNPAID_COLS_, destSh.getMaxColumns())).clearContent();
  }
  var st = unpaidStatsFromValues_(vals);
  unpaidWriteDashboardValues_(destSh, st);
  SpreadsheetApp.flush();
  return {
    ok: true,
    month: monthName,
    rows: vals.length,
    cols: shown[0].length,
    a5: String(destSh.getRange('A5').getDisplayValue() || ''),
    d2: String(destSh.getRange('D2').getDisplayValue() || ''),
    n: st.n,
    sp: st.sp,
    sr: st.sr,
    nr: st.nr,
    checks: boolCols,
    flushed: flushed,
    n7check: (function () {
      var dv = destSh.getRange('N7').getDataValidation();
      return dv ? String(dv.getCriteriaType()) : '';
    })(),
    stats: st
  };
}

function ensureUnpaidEditTrigger_() {
  var ss = openWorkspaceSpreadsheet_();
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'unpaidOnEditInstalled_') return { ok: true, existed: true };
  }
  ScriptApp.newTrigger('unpaidOnEditInstalled_').forSpreadsheet(ss).onEdit().create();
  return { ok: true, created: true };
}

function unpaidOnEditInstalled_(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (!sh) return;
  var name = sh.getName();
  if (name === JOIN_LIST_SHEET_) {
    try { joinListPushEdited_(e); } catch (e1) {}
    try { joinListFlushQueue_(sh.getParent()); } catch (e2) {}
    return;
  }
  if (name !== UNPAID_SHEET_) return;
  if (e.range.getRow() === 1 && e.range.getColumn() === 2) {
    var filledEd = unpaidFillFromSource_(sh);
    styleUnpaidView_(sh);
    styleUnpaidDashboard_(sh);
    unpaidHideNoiseCols_(sh);
    try {
      var trEd = sh.getParent().getSheetByName(UNPAID_TREND_SHEET_);
      if (trEd && filledEd && filledEd.ok && filledEd.stats) unpaidWriteTrendMonth_(trEd, filledEd.month, filledEd.stats);
    } catch (eTr) {}
    return;
  }
  try { unpaidFlushQueue_(sh.getParent()); } catch (eF) {}
}

function fillUnpaidFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || openWorkspaceSpreadsheet_();
  var sh = ss.getSheetByName(UNPAID_SHEET_);
  if (!sh) {
    ss.toast('未納管理シートがありません', '未納', 8);
    return { ok: false };
  }
  try { ensureUnpaidEditTrigger_(); } catch (eT) {}
  var filled = unpaidFillFromSource_(sh);
  try { styleUnpaidView_(sh); styleUnpaidDashboard_(sh); unpaidHideNoiseCols_(sh); } catch (eS) {}
  var trend = null;
  try { trend = unpaidFillTrendAll_(ss); } catch (eTr) { trend = { ok: false, message: String(eTr) }; }
  var join = null;
  try { joinListFlushQueue_(ss); } catch (eJQ) {}
  try { joinListPushChecks_(ss); } catch (eJP) {}
  try { join = joinListRefresh_(ss.getSheetByName(JOIN_LIST_SHEET_)); } catch (eJ) {}
  ss.toast(filled && filled.ok ? ('再取得しました（' + filled.d2 + '）') : String(filled && filled.message ? filled.message : '失敗'), '再取得', 8);
  return { unpaid: filled, trend: trend, joinList: join };
}

function flushUnpaidFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || openWorkspaceSpreadsheet_();
  try { ensureUnpaidEditTrigger_(); } catch (eT) {}
  var r = unpaidFlushQueue_(ss);
  var jq = { ok: true, pushed: 0 };
  try { jq = joinListFlushQueue_(ss); } catch (eJQ) { jq = { ok: false, message: String(eJQ) }; }
  var j = { ok: true, kyodo: { pushed: 0 }, fit365: { pushed: 0 } };
  try { j = joinListPushChecks_(ss); } catch (eJ) { j = { ok: false, message: String(eJ) }; }
  try { joinListRefresh_(ss.getSheetByName(JOIN_LIST_SHEET_)); } catch (eR) {}
  var jp = ((j.kyodo && j.kyodo.pushed) || 0) + ((j.fit365 && j.fit365.pushed) || 0) + ((jq && jq.pushed) || 0);
  var total = (r.pushed || 0) + jp;
  ss.toast(total ? ('元ファイルへ ' + total + ' 件送りました') : (r.errors || !j.ok ? '送れなかったマスがあります' : '送る変更はありません'), '反映', 8);
  return { unpaid: r, joinQueue: jq, joinList: j };
}

/**
 * 月タブ1枚分の集計に使う LET 変数（新旧レイアウト両対応。列は見出しの文字で探す）。
 * ok=会員番号あり・支払額>0 の行、pay=支払額、tot=総額、cat=A列区分の埋め。
 * rec=回収金額。金額欄が空でも右隣3列に「〇〇入金」があれば支払額を回収済みとみなす（旧レイアウトは金額欄を使っていない）。
 */
function unpaidStatsLet_(rangeExpr) {
  return 'd,IMPORTRANGE("' + UNPAID_SOURCE_ID_ + '",' + rangeExpr + '),' +
    'h,BYCOL(CHOOSEROWS(d,1,2,3),LAMBDA(c,TEXTJOIN("",TRUE,c))),' +
    'fc,LAMBDA(re,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(h,re)),0),0)),' +
    'nv,LAMBDA(col,IF(col<1,0,IFERROR(VALUE(REGEXREPLACE(CHOOSECOLS(d,col)&"","[¥,\\s]","")),0))),' +
    'cat,SCAN("",CHOOSECOLS(d,1),LAMBDA(a,x,IF(x&""<>"",x&"",a))),' +
    'mc,fc("会員番号"),' +
    'mem,TRIM(CHOOSECOLS(d,IF(mc<1,1,mc))&""),' +
    'pay,nv(fc("支払額")),tot,nv(fc("総額")),uc,fc("回収金額|入金金額|レジ打ち金額"),recraw,nv(uc),' +
    'rec,IF(uc<1,0,IF(recraw>0,recraw,IF(REGEXMATCH(CHOOSECOLS(d,uc+1)&CHOOSECOLS(d,uc+2)&CHOOSECOLS(d,uc+3)&"","入金"),pay,0))),' +
    'ok,(mc>0)*(mem<>"")*(mem<>"会員番号")*(mem<>"合計")*(pay>0),' +
    'n,SUM(ok),sp,SUM(ok*pay),st,SUM(ok*tot),sr,SUM(ok*rec),nr,SUM(ok*(rec>0)),' +
    'kone,ok*REGEXMATCH(cat,"1ヶ月|1ヵ月|1カ月|1か月"),' +
    'ktwo,ok*REGEXMATCH(cat,"2ヶ月|2ヵ月|2カ月|2か月"),' +
    'kbad,ok*REGEXMATCH(cat,"貸倒|貸し倒"),' +
    'kjac,ok*REGEXMATCH(cat,"JACCS"),';
}

/** 未納管理 D1:M3 のダッシュボード（B1 の月） */
function unpaidDashboardFormula_() {
  return '=ARRAYFORMULA(LET(' + unpaidStatsLet_('"\'"&$B$1&"\'!A1:AK"') +
    'yen,LAMBDA(x,TEXT(x,"¥#,##0")),' +
    'rt,LAMBDA(k,IF(SUM(k)=0,"対象なし",IF(SUM(k*pay)=0,"対象なし",TEXT(SUM(k*rec)/SUM(k*pay),"0.0%")))),' +
    'sub,LAMBDA(k,SUM(k)&"件　"&yen(SUM(k*rec))&" / "&yen(SUM(k*pay))),' +
    'VSTACK({"未納件数","未納総額","回収額","回収率","未回収額","回収済み","1ヶ月未納 回収率","2ヶ月未納 回収率","貸倒候補 回収率","JACCS 回収率"},' +
    'HSTACK(n&"件",yen(sp),yen(sr),IF(sp=0,"-",TEXT(sr/sp,"0.0%")),yen(sp-sr),nr&"件",rt(kone),rt(ktwo),rt(kbad),rt(kjac)),' +
    'HSTACK("支払額ベース","手数料込 "&yen(st),"回収金額の合計","回収額÷未納総額","残り "&(n-nr)&"件",' +
    'IF(n=0,"",TEXT(nr/n,"0%")&"（人数）"),sub(kone),sub(ktwo),sub(kbad),sub(kjac)))))';
}

/** 未納管理_推移 の1行（A列の月）。数値のまま返す */
function unpaidTrendRowFormula_(row) {
  return '=IFERROR(ARRAYFORMULA(LET(' + unpaidStatsLet_('"\'"&$A' + row + '&"\'!A1:AK"') +
    'rq,LAMBDA(k,IF(SUM(k)=0,"対象なし",IF(SUM(k*pay)=0,"対象なし",SUM(k*rec)/SUM(k*pay)))),' +
    'rate,IF(sp=0,"",sr/sp),' +
    'HSTACK(n,sp,sr,rate,IF(rate="","",SPARKLINE(rate,{"charttype","bar";"max",1;"color1","' + dnTheme_().blood + '"})),sp-sr,nr,rq(kone),rq(ktwo),rq(kbad)))),"")';
}

function unpaidHideNoiseCols_(sh) {
  var maxC = sh.getMaxColumns();
  try { sh.showColumns(1, maxC); } catch (eShow) {}
  var hideFrom = 27;
  if (maxC >= hideFrom) {
    try { sh.hideColumns(hideFrom, maxC - hideFrom + 1); } catch (eTail) {}
  }
}

function styleUnpaidView_(sh) {
  var maxR = Math.min(sh.getMaxRows(), Math.max(sh.getLastRow(), UNPAID_DATA_ROW_ + 20));
  var top = UNPAID_DATA_ROW_;
  var solid = SpreadsheetApp.BorderStyle.SOLID;
  var W = UNPAID_COLS_;
  var t = dnTheme_();
  var r0 = top + 2;
  var bodyRows = maxR - r0 + 1;
  var body = sh.getRange(top, 1, maxR - top + 1, W);
  hubType_(body).setBackground(t.paper).setFontColor(t.ink).setFontSize(10).setFontWeight('normal')
    .setVerticalAlignment('middle').setWrap(false).setBorder(false, false, false, false, false, false);
  hubType_(sh.getRange(top, 1, 2, W)).setFontWeight('bold').setBackground(t.ink).setFontColor(t.paper)
    .setFontSize(9).setWrap(true).setHorizontalAlignment('center')
    .setBorder(false, false, true, false, false, false, t.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(r0, 1, bodyRows, W)
    .setBorder(null, null, null, null, null, true, t.line, solid);
  sh.getRange(r0, 2, bodyRows, W - 1).setHorizontalAlignment('center');
  sh.getRange(r0, UNPAID_NAME_COL_, bodyRows, 1).setHorizontalAlignment('left')
    .setFontWeight('bold').setFontSize(11);
  sh.getRange(r0, 1, bodyRows, 1).setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setFontColor(t.ink);
  sh.getRange(r0, UNPAID_PAY_COL_, bodyRows, 1).setFontSize(11).setFontWeight('bold');
  sh.getRange(r0, UNPAID_REC_COL_, bodyRows, 1).setFontWeight('bold');

  var rng = sh.getRange(r0, 1, bodyRows, W);
  var joinRng = sh.getRange(r0, UNPAID_NAME_COL_, bodyRows, 5);
  var catRng = sh.getRange(r0, 1, bodyRows, 1);
  var payRng = sh.getRange(r0, UNPAID_PAY_COL_, bodyRows, 1);
  var recRng = sh.getRange(r0, UNPAID_REC_COL_, bodyRows, 1);
  var name = '$D' + r0 + '&""';
  var hasName = 'AND(' + name + '<>"",' + name + '<>"会員名")';
  var recovered = 'REGEXMATCH($U' + r0 + '&"","[1-9]")';
  var early = 'FALSE';
  var none = 'FALSE';
  var rules = sh.getConditionalFormatRules().filter(function (r) {
    return r.getRanges().every(function (x) { return x.getSheet().getName() !== UNPAID_SHEET_; });
  });
  var add = function (formula, bg, color, bold, range) {
    var b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setRanges([range || rng]);
    if (bg) b.setBackground(bg);
    if (color) b.setFontColor(color);
    if (bold) b.setBold(true);
    rules.push(b.build());
  };
  add('=OR($B' + r0 + '="会員番号",$C' + r0 + '="会員番号")', t.ink, t.paper, true);
  add('=OR($A' + r0 + '="合計",$B' + r0 + '="合計")', t.ink, t.paper, true);
  add('=AND(' + hasName + ',REGEXMATCH($A' + r0 + '&"","^1ヶ月"))', t.blood, t.paper, true, catRng);
  add('=AND(' + hasName + ',REGEXMATCH($A' + r0 + '&"","^2ヶ月"))', t.ink, t.paper, true, catRng);
  add('=AND(' + hasName + ',REGEXMATCH($A' + r0 + '&"","貸倒"))', t.ash, t.paper, true, catRng);
  add('=AND(' + hasName + ',REGEXMATCH($A' + r0 + '&"","JACCS"))', t.ash, t.paper, true, catRng);
  add('=AND(' + hasName + ',REGEXMATCH($A' + r0 + '&"","過年度"))', t.ink, t.paper, true, catRng);
  add('=AND(' + hasName + ',' + recovered + ')', t.paper, t.ash, true, recRng);
  add('=AND(' + hasName + ',' + recovered + ')', t.paper, t.ash, false);
  add('=AND(' + hasName + ',NOT(' + recovered + '))', null, t.blood, true, payRng);
  add('=AND(' + hasName + ',' + early + ')', null, t.blood, true, joinRng);
  add('=AND(' + hasName + ',' + none + ')', null, t.ash, false, joinRng);
  sh.setConditionalFormatRules(rules);
  unpaidHideNoiseCols_(sh);
}

var UNPAID_SCAN_PROP_ = 'unpaidScanV2';
var FOUR_COLOR_PROP_ = 'hubFieldV1';

function applyUnpaidScanLook_(sh) {
  if (!sh) return { ok: false };
  try { sh.showRows(1, sh.getMaxRows()); } catch (eShowR) {}
  try { sh.showSheet(); } catch (eShowS) {}
  if (sh.getMaxColumns() < UNPAID_COLS_) {
    sh.insertColumnsAfter(sh.getMaxColumns(), UNPAID_COLS_ - sh.getMaxColumns());
  }
  try { permitImportRange_(sh.getParent(), UNPAID_SOURCE_ID_); } catch (eP) {}
  try { ensureUnpaidEditTrigger_(); } catch (eT) {}
  var filled = { ok: false };
  try { filled = unpaidFillFromSource_(sh); } catch (eFill) { filled = { ok: false, message: String(eFill) }; }
  SpreadsheetApp.flush();
  styleUnpaidView_(sh);
  styleUnpaidDashboard_(sh);
  sh.setColumnWidth(1, 72);
  sh.setColumnWidth(2, 52);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 148);
  sh.setColumnWidths(5, 4, 88);
  sh.setColumnWidth(9, 110);
  sh.setColumnWidths(10, 2, 92);
  sh.setColumnWidths(12, 5, 100);
  sh.setColumnWidths(14, 7, 44);
  sh.setColumnWidth(25, 104);
  sh.setColumnWidth(26, 92);
  try { sh.setRowHeightsForced(UNPAID_DATA_ROW_ + 2, Math.min(Math.max(sh.getLastRow() - UNPAID_DATA_ROW_ - 1, 10), 400), 24); } catch (eH) {}
  unpaidHideNoiseCols_(sh);
  sh.setFrozenColumns(4);
  sh.setFrozenRows(UNPAID_DATA_ROW_ + 1);
  sh.setHiddenGridlines(true);
  return { ok: true, filled: filled };
}

function maybeApplyFourColorOnce_() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty(FOUR_COLOR_PROP_) === '1') return { ok: true, skipped: true };
  applyFourColorLook_();
  props.setProperty(FOUR_COLOR_PROP_, '1');
  props.setProperty(UNPAID_SCAN_PROP_, '1');
  return { ok: true };
}

function applyFourColorLook_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || openWorkspaceSpreadsheet_();
  try { setupSchoolDiscountSheet_(ss); } catch (eSch) {}
  setupHubHome_();
  return restyleHubLook_();
}

function applyFourColorFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = applyFourColorLook_();
  try { PropertiesService.getDocumentProperties().setProperty(FOUR_COLOR_PROP_, '1'); } catch (eP) {}
  ss.toast('見出しは中央、文字は斜体です。空セルに打った文字も斜体になります', '経堂', 8);
  return r;
}

function maybeApplyUnpaidScanOnce_(optSheet) {
  return maybeApplyFourColorOnce_();
}

function applyUnpaidScanFromMenu() {
  return applyFourColorFromMenu();
}

function styleUnpaidDashboard_(sh) {
  var t = dnTheme_();
  var solid = SpreadsheetApp.BorderStyle.SOLID;
  try { sh.getRange('B1:C1').breakApart(); } catch (e0) {}
  try { sh.getRange('A2:C3').breakApart(); } catch (e1) {}
  try { sh.getRange('A4:C4').breakApart(); } catch (e2) {}
  sh.getRange(1, 1, 4, 37).setBackground(t.paper).setFontColor(t.ink).setVerticalAlignment('middle');
  hubType_(sh.getRange(1, 1, 4, 13));
  sh.getRange('A1').setFontSize(9).setFontColor(t.ash).setHorizontalAlignment('center');
  sh.getRange('B1:C1').merge().setFontSize(12).setFontWeight('bold').setFontStyle('italic').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, null, null, t.blood, solid);
  var unpaidSrc = hubSourceForName_(UNPAID_SHEET_);
  if (unpaidSrc) {
    hubPaintOpenSourceCell_(sh, 2, 1, 3, unpaidSrc);
  } else {
    sh.getRange('A2:C2').merge().setFontSize(9).setHorizontalAlignment('left');
  }
  sh.getRange('A3:C3').merge().setFontSize(9).setHorizontalAlignment('center').setFontStyle('italic')
    .setBackground(t.paper).setFontColor(t.ash).setFontWeight('bold');

  var card = sh.getRange('D1:M3');
  card.setBackground(t.cream).setHorizontalAlignment('center').setWrap(true)
    .setBorder(true, true, true, true, true, false, t.paper, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange('D1:M1').setFontSize(8).setFontColor(t.ash).setFontWeight('bold');
  sh.getRange('D2:M2').setFontSize(14).setFontColor(t.ink).setFontWeight('bold').setWrap(false);
  sh.getRange('D3:M3').setFontSize(8).setFontColor(t.ash).setFontWeight('normal');
  sh.getRange('G1:G3').setBackground(t.blood).setFontColor(t.paper);
  sh.getRange('G1').setFontColor(t.paper);
  sh.getRange('G3').setFontColor(t.paper);
  sh.getRange('A4:C4').merge().setValue('チェック・会員名・回収金額はこちらで編集できます。変えたら元の未納管理ドライブへ自動で送ります。')
    .setFontSize(8).setFontColor(t.ash).setHorizontalAlignment('left').setFontWeight('normal')
    .setBackground(t.paper);

  sh.setRowHeight(1, 24);
  sh.setRowHeight(2, 36);
  sh.setRowHeight(3, 32);
  sh.setRowHeight(4, 20);
}

var UNPAID_TREND_HEAD_ROW_ = 6;
var UNPAID_TREND_CHART_SHEET_ = '未納_推移_グラフ元';
var UNPAID_TREND_CHART_GAP_ROWS_ = 15;

function unpaidTrendRowFromStats_(monthName, st) {
  st = st || { n: 0, sp: 0, sr: 0, nr: 0, kone: { pay: 0, rec: 0 }, ktwo: { pay: 0, rec: 0 }, kbad: { pay: 0, rec: 0 } };
  var rq = function (k) {
    return k && k.pay ? k.rec / k.pay : '';
  };
  return [
    monthName,
    st.n || 0,
    st.sp || 0,
    st.sr || 0,
    st.sp ? st.sr / st.sp : '',
    '',
    (st.sp || 0) - (st.sr || 0),
    st.nr || 0,
    rq(st.kone),
    rq(st.ktwo),
    rq(st.kbad)
  ];
}

function unpaidSparkFormula_(row) {
  return '=IFERROR(IF(N(E' + row + ')<=0,"",SPARKLINE(E' + row + ',{"charttype","bar";"max",1;"color1","' + dnTheme_().blood + '"})),"")';
}

function unpaidSourceMonthTabs_(src) {
  var re = /^(\d{2})年(\d{1,2})月$/;
  var out = [];
  var sheets = src.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    var m = n.match(re);
    if (!m) continue;
    out.push({
      name: n,
      sh: sheets[i],
      y: 2000 + Number(m[1]),
      m: Number(m[2])
    });
  }
  out.sort(function (a, b) { return a.y - b.y || a.m - b.m; });
  return out;
}

function unpaidReadMonthValuesFast_(srcSh) {
  if (!srcSh) return null;
  var lastR = Number(srcSh.getLastRow()) || 0;
  var lastC = Number(srcSh.getLastColumn()) || 0;
  if (lastR < 3 || lastC < 1) return null;
  lastR = Math.min(lastR, 500);
  lastC = Math.min(lastC, 40);
  return srcSh.getRange(1, 1, lastR, lastC).getValues();
}

function unpaidEnsureTrendSheet_(ss) {
  ss = ss || openWorkspaceSpreadsheet_();
  var tr = ss.getSheetByName(UNPAID_TREND_SHEET_);
  if (!tr) tr = ss.insertSheet(UNPAID_TREND_SHEET_);
  if (tr.getMaxColumns() < 12) tr.insertColumnsAfter(tr.getMaxColumns(), 12 - tr.getMaxColumns());
  if (tr.getMaxRows() < 90) tr.insertRowsAfter(tr.getMaxRows(), 90 - tr.getMaxRows());
  try { tr.showSheet(); } catch (eShow) {}
  try { tr.showColumns(1, tr.getMaxColumns()); } catch (eShowC) {}
  return tr;
}

function unpaidEnsureTrendChartSheet_(ss) {
  ss = ss || openWorkspaceSpreadsheet_();
  var sh = ss.getSheetByName(UNPAID_TREND_CHART_SHEET_);
  if (!sh) sh = ss.insertSheet(UNPAID_TREND_CHART_SHEET_);
  if (sh.getMaxColumns() < 14) sh.insertColumnsAfter(sh.getMaxColumns(), 14 - sh.getMaxColumns());
  if (sh.getMaxRows() < 40) sh.insertRowsAfter(sh.getMaxRows(), 40 - sh.getMaxRows());
  try { sh.hideSheet(); } catch (eH) {}
  sh.setTabColor(dnTheme_().ash);
  return sh;
}

function unpaidClearTrendCharts_(tr) {
  var charts = tr.getCharts();
  var i;
  for (i = charts.length - 1; i >= 0; i--) {
    try { tr.removeChart(charts[i]); } catch (e0) {}
  }
}

function unpaidChartStyle_() {
  var t = dnTheme_();
  return {
    backgroundColor: t.paper,
    titleTextStyle: { color: t.ink, fontName: 'Noto Sans JP', italic: true, fontSize: 12, bold: true },
    legend: { position: 'bottom', textStyle: { color: t.ash, fontName: 'Noto Sans JP', fontSize: 9 } },
    hAxis: { textStyle: { color: t.ash, fontSize: 8, fontName: 'Noto Sans JP' }, slantedText: true, slantedTextAngle: 45 },
    vAxis: { textStyle: { color: t.ash, fontSize: 9, fontName: 'Noto Sans JP' }, gridlines: { color: t.paper }, minorGridlines: { count: 0 } },
    chartArea: { left: 56, top: 36, width: '78%', height: '58%' }
  };
}

function unpaidInsertTrendChart_(tr, spec) {
  try {
    var b = tr.newChart()
      .addRange(spec.range)
      .setNumHeaders(1)
      .setPosition(spec.row, spec.col, 8, 8)
      .setOption('title', spec.title)
      .setOption('width', spec.width || 540)
      .setOption('height', spec.height || 280)
      .setOption('colors', spec.colors || ['#171717', '#DA0037', '#444444'])
      .setOption('backgroundColor', '#EDEDED')
      .setOption('legend', { position: 'bottom' })
      .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
      .setOption('useFirstColumnAsDomain', true);
    if (spec.kind === 'line') b = b.asLineChart();
    else b = b.asColumnChart();
    if (spec.vAxis) b = b.setOption('vAxis', spec.vAxis);
    tr.insertChart(b.build());
    return { ok: true };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function unpaidChartNum_(v) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number' && !isNaN(v)) return v;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

function unpaidWriteTrendChartSource_(tr, body) {
  var start = 14;
  if (tr.getMaxColumns() < start + 13) {
    tr.insertColumnsAfter(tr.getMaxColumns(), start + 13 - tr.getMaxColumns());
  }
  try { tr.showColumns(1, tr.getMaxColumns()); } catch (eShow) {}
  var clearTo = Math.max(tr.getMaxRows(), (body && body.length ? body.length : 1) + 2);
  tr.getRange(1, start, clearTo, 14).clearContent();
  var src = unpaidEnsureTrendChartSheet_(tr.getParent());
  src.clear();
  if (!body || !body.length) return { ok: true, rows: 0, sheet: src, start: start };
  var headers = [[
    '年月', '未納総額', '回収額', '',
    '年月', '回収率', '',
    '年月', '未納件数', '',
    '年月', '1ヶ月未納', '2ヶ月未納', '貸倒候補'
  ]];
  var rows = [];
  var i;
  for (i = 0; i < body.length; i++) {
    var r = body[i];
    rows.push([
      r[0], unpaidChartNum_(r[2]), unpaidChartNum_(r[3]), '',
      r[0], unpaidChartNum_(r[4]), '',
      r[0], unpaidChartNum_(r[1]), '',
      r[0], unpaidChartNum_(r[8]), unpaidChartNum_(r[9]), unpaidChartNum_(r[10])
    ]);
  }
  src.getRange(1, 1, 1, 14).setValues(headers);
  src.getRange(2, 1, rows.length, 14).setValues(rows);
  try { src.hideSheet(); } catch (eH) {}
  // 同じシートの見える列にも置く（非表示列だとグラフが空になる）
  var t = dnTheme_();
  tr.getRange(1, start, 1, 14).setValues(headers);
  tr.getRange(2, start, rows.length, 14).setValues(rows);
  tr.getRange(2, start + 1, rows.length, 2).setNumberFormat('0');
  tr.getRange(2, start + 5, rows.length, 1).setNumberFormat('0.000');
  tr.getRange(2, start + 8, rows.length, 1).setNumberFormat('0');
  tr.getRange(2, start + 11, rows.length, 3).setNumberFormat('0.000');
  tr.getRange(1, start, rows.length + 1, 14)
    .setFontColor(t.paper).setBackground(t.paper).setFontSize(8);
  try { tr.setColumnWidths(start, 14, 10); } catch (eW) {}
  return { ok: true, rows: rows.length, sheet: src, start: start };
}

function unpaidTrendTableBounds_(tr) {
  var head = UNPAID_TREND_HEAD_ROW_;
  var scanLast = Math.min(Math.max(tr.getLastRow(), head), 60);
  if (String(tr.getRange(head, 1).getDisplayValue() || '') !== '年月') {
    var r;
    for (r = 1; r <= scanLast; r++) {
      if (String(tr.getRange(r, 1).getDisplayValue() || '') === '年月') {
        head = r;
        break;
      }
    }
  }
  var labels = tr.getRange(head, 1, Math.max(scanLast - head + 1, 1), 1).getDisplayValues();
  var n = 0;
  var i;
  for (i = 1; i < labels.length; i++) {
    if (/^\d{2}年\d{1,2}月$/.test(String(labels[i][0] || ''))) n = i;
    else if (n > 0) break;
  }
  return { head: head, n: n, last: head + n };
}

function unpaidEnsureTrendCharts_(tr) {
  if (!tr) return { ok: false, charts: 0 };
  unpaidClearTrendCharts_(tr);
  var b = unpaidTrendTableBounds_(tr);
  if (!b.n) return { ok: true, charts: 0 };
  var tableBody = tr.getRange(b.head + 1, 1, b.n, 11).getValues();
  var srcPack = unpaidWriteTrendChartSource_(tr, tableBody);
  var start = (srcPack && srcPack.start) || 14;
  var t = dnTheme_();
  var nSrc = Math.max((srcPack && srcPack.rows) || b.n, 1);
  var srcRows = nSrc + 1;
  var chartTop = b.last + 3;
  var gap = UNPAID_TREND_CHART_GAP_ROWS_;
  var need = chartTop + gap * 2 + 2;
  if (tr.getMaxRows() < need) tr.insertRowsAfter(tr.getMaxRows(), need - tr.getMaxRows());
  SpreadsheetApp.flush();
  var pctAxis = { format: '0%', viewWindow: { min: 0, max: 1 } };
  var errors = [];
  var placed = 0;
  var specs = [
    {
      kind: 'column',
      range: tr.getRange(1, start, srcRows, 3),
      title: '未納総額と回収額',
      row: chartTop,
      col: 1,
      colors: [t.ink, t.blood]
    },
    {
      kind: 'line',
      range: tr.getRange(1, start + 4, srcRows, 2),
      title: '回収率',
      row: chartTop,
      col: 7,
      colors: [t.blood],
      vAxis: pctAxis
    },
    {
      kind: 'column',
      range: tr.getRange(1, start + 7, srcRows, 2),
      title: '未納件数',
      row: chartTop,
      col: 1,
      colors: [t.ink]
    },
    {
      kind: 'line',
      range: tr.getRange(1, start + 10, srcRows, 4),
      title: '1ヶ月 / 2ヶ月 / 貸倒 の回収率',
      row: chartTop + gap,
      col: 7,
      colors: [t.ink, t.ash, t.blood],
      vAxis: pctAxis
    }
  ];
  var i;
  for (i = 0; i < specs.length; i++) {
    var r = unpaidInsertTrendChart_(tr, specs[i]);
    if (r && r.ok) placed += 1;
    else errors.push((specs[i].title || '') + ': ' + ((r && r.message) || 'fail'));
  }
  return { ok: placed > 0, charts: placed, chartTop: chartTop, errors: errors, srcRows: nSrc };
}

function unpaidStyleTrendSheet_(tr) {
  if (!tr) return;
  var t = dnTheme_();
  var b = unpaidTrendTableBounds_(tr);
  var paintRows = Math.max(b.last + UNPAID_TREND_CHART_GAP_ROWS_ * 2 + 8, 70);
  var cols = Math.min(tr.getMaxColumns(), 12);
  hubType_(tr.getRange(1, 1, paintRows, cols))
    .setBackground(t.paper).setFontColor(t.ink).setFontSize(10)
    .setFontWeight('normal').setVerticalAlignment('middle').setWrap(false);
  try { tr.getRange(1, 1, paintRows, cols).setBorder(false, false, false, false, false, false); } catch (eB) {}
  hubPaintOpenSourceCell_(tr, 1, 1, 3, {
    url: 'https://docs.google.com/spreadsheets/d/' + UNPAID_SOURCE_ID_ + '/edit',
    label: '元の未納管理ドライブを開く ↗'
  });
  hubType_(tr.getRange(1, 4, 1, 5))
    .setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setFontSize(14).setHorizontalAlignment('center')
    .setBorder(false, false, true, false, false, false, t.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  tr.getRange(1, 9, 1, 3).setBackground(t.paper).setFontColor(t.ash).setFontSize(9)
    .setHorizontalAlignment('right').setFontWeight('normal');
  var i;
  for (i = 0; i < 6; i++) {
    var c1 = i * 2 + 1;
    tr.getRange(2, c1, 1, 2).setBackground(t.ink).setFontColor(t.paper).setFontSize(8)
      .setFontWeight('bold').setHorizontalAlignment('center');
    tr.getRange(3, c1, 1, 2).setBackground(t.paper).setFontColor(t.ink).setFontSize(16)
      .setFontWeight('bold').setHorizontalAlignment('center');
  }
  tr.getRange(4, 1, 1, 12).setBackground(t.paper).setFontColor(t.ash).setFontSize(9)
    .setHorizontalAlignment('left');
  if (b.n > 0) {
    var head = tr.getRange(b.head, 1, 1, 11);
    hubType_(head).setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
      .setFontSize(9).setHorizontalAlignment('center').setWrap(true)
      .setBorder(false, false, true, false, false, false, t.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    tr.getRange(b.head, 6).setBackground(t.blood).setFontColor(t.paper);
    var body = tr.getRange(b.head + 1, 1, b.n, 11);
    hubType_(body).setBackground(t.paper).setFontColor(t.ink).setFontSize(10)
      .setHorizontalAlignment('center');
    tr.getRange(b.head + 1, 2, b.n, 10).setHorizontalAlignment('right');
    tr.getRange(b.head + 1, 1, b.n, 1).setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('@');
    tr.getRange(b.head + 1, 2, b.n, 1).setNumberFormat('0"件"');
    tr.getRange(b.head + 1, 3, b.n, 2).setNumberFormat('¥#,##0');
    tr.getRange(b.head + 1, 5, b.n, 1).setNumberFormat('0.0%');
    tr.getRange(b.head + 1, 7, b.n, 1).setNumberFormat('¥#,##0');
    tr.getRange(b.head + 1, 8, b.n, 1).setNumberFormat('0"件"');
    tr.getRange(b.head + 1, 9, b.n, 3).setNumberFormat('0.0%');
    try { tr.getRange(b.head + 1, 1, b.n, 11).setBorder(null, null, null, null, null, true, t.line, SpreadsheetApp.BorderStyle.SOLID); } catch (eL) {}
    var tz = 'Asia/Tokyo';
    var cur = Utilities.formatDate(new Date(), tz, 'yy') + '年' + Number(Utilities.formatDate(new Date(), tz, 'M')) + '月';
    var names = tr.getRange(b.head + 1, 1, b.n, 1).getDisplayValues();
    var r;
    for (r = 0; r < names.length; r++) {
      if (String(names[r][0]) === cur) {
        tr.getRange(b.head + 1 + r, 1, 1, 11)
          .setBorder(false, false, true, false, false, false, t.blood, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
  tr.getRange('A3:D3').setNumberFormat('@');
  tr.getRange('E3:F3').setNumberFormat('0"件"');
  tr.getRange('G3:H3').setNumberFormat('¥#,##0');
  tr.getRange('I3:J3').setNumberFormat('0.0%');
  tr.getRange('K3:L3').setNumberFormat('¥#,##0');
  tr.setFrozenRows(3);
  tr.setHiddenGridlines(true);
  tr.setTabColor(t.blood);
  tr.setRowHeight(1, 28);
  tr.setRowHeight(2, 22);
  tr.setRowHeight(3, 36);
  tr.setRowHeight(4, 28);
  tr.setRowHeight(5, 12);
  tr.setColumnWidth(1, 92);
  tr.setColumnWidths(2, 4, 100);
  tr.setColumnWidth(6, 150);
  tr.setColumnWidths(7, 5, 100);
  try { if (b.n > 0) tr.setRowHeightsForced(b.head + 1, b.n, 26); } catch (eH) {}
}

function unpaidWriteTrendMonth_(tr, monthName, st) {
  if (!tr || !monthName || !st) return;
  var b = unpaidTrendTableBounds_(tr);
  if (!b.n) return;
  var names = tr.getRange(b.head + 1, 1, b.n, 1).getDisplayValues();
  var i;
  for (i = 0; i < names.length; i++) {
    if (String(names[i][0]) === monthName) {
      var row = b.head + 1 + i;
      tr.getRange(row, 1, 1, 11).setValues([unpaidTrendRowFromStats_(monthName, st)]);
      tr.getRange(row, 6).setFormula(unpaidSparkFormula_(row));
      try {
        var bodyNow = tr.getRange(b.head + 1, 1, b.n, 11).getValues();
        unpaidWriteTrendChartSource_(tr, bodyNow);
      } catch (eC) {}
      return;
    }
  }
}

function unpaidPaintTrend_(tr, body, totals, latest) {
  var t = dnTheme_();
  unpaidClearTrendCharts_(tr);
  var wipeRows = Math.min(tr.getMaxRows(), 90);
  var wipeCols = tr.getMaxColumns();
  try { tr.getRange(1, 1, wipeRows, wipeCols).breakApart(); } catch (e0) {}
  try { tr.getRange(1, 1, wipeRows, wipeCols).clearContent(); } catch (e1) {}
  try { tr.getRange(1, 1, wipeRows, wipeCols).clearDataValidations(); } catch (e2) {}
  var head = UNPAID_TREND_HEAD_ROW_;
  var need = head + Math.max(body.length, 1) + UNPAID_TREND_CHART_GAP_ROWS_ * 2 + 8;
  if (tr.getMaxRows() < need) tr.insertRowsAfter(tr.getMaxRows(), need - tr.getMaxRows());
  var srcUrl = 'https://docs.google.com/spreadsheets/d/' + UNPAID_SOURCE_ID_ + '/edit';
  hubPaintOpenSourceCell_(tr, 1, 1, 3, { url: srcUrl, label: '元の未納管理ドライブを開く ↗' });
  tr.getRange(1, 4, 1, 5).merge().setValue('未納 月別推移');
  var nowText = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  tr.getRange(1, 9, 1, 3).merge().setValue('更新 ' + nowText);
  var latestName = latest && latest.name ? latest.name : '—';
  var ls = latest && latest.st ? latest.st : { n: 0, sp: 0, sr: 0 };
  var latestRate = ls.sp ? ls.sr / ls.sp : '';
  tr.getRange(2, 1, 1, 12).setValues([[
    '対象月', '', '直近月', '', '直近 件数', '', '直近 未納総額', '', '直近 回収率', '', '直近 未回収', ''
  ]]);
  tr.getRange(3, 1, 1, 12).setValues([[
    body.length ? body.length + 'ヶ月' : '0ヶ月', '',
    latestName, '',
    ls.n || 0, '',
    ls.sp || 0, '',
    latestRate, '',
    (ls.sp || 0) - (ls.sr || 0), ''
  ]]);
  var mi;
  for (mi = 0; mi < 6; mi++) {
    try { tr.getRange(2, mi * 2 + 1, 1, 2).merge(); } catch (eM1) {}
    try { tr.getRange(3, mi * 2 + 1, 1, 2).merge(); } catch (eM2) {}
  }
  var totRate = totals.sp ? totals.sr / totals.sp : '';
  tr.getRange(4, 1, 1, 12).merge().setValue(
    '各月タブのスナップショットです。繰越があると月をまたいで二重計上されます。表の下にグラフがあります。空の月タブは出しません。　月合計 ' +
    (totals.n || 0) + '件 / 未納 ' + unpaidYen_(totals.sp) + ' / 回収 ' + unpaidYen_(totals.sr) +
    (totRate === '' ? '' : ' / 合計回収率 ' + (Math.round(totRate * 1000) / 10).toFixed(1) + '%')
  );
  tr.getRange(head, 1, 1, 11).setValues([[
    '年月', '未納件数', '未納総額', '回収額', '回収率', '全体回収率バー', '未回収額', '回収済み件数',
    '1ヶ月未納 回収率', '2ヶ月未納 回収率', '貸倒候補 回収率'
  ]]);
  if (body.length) {
    tr.getRange(head + 1, 1, body.length, 11).setValues(body);
    var spark = [];
    var i;
    for (i = 0; i < body.length; i++) spark.push([unpaidSparkFormula_(head + 1 + i)]);
    tr.getRange(head + 1, 6, body.length, 1).setFormulas(spark);
    tr.getRange(head + body.length + 2, 1, 1, 11).merge()
      .setValue('月ごとの推移グラフ')
      .setFontSize(10).setFontColor(t.ash).setFontWeight('bold')
      .setHorizontalAlignment('left').setBackground(t.paper);
  }
  unpaidWriteTrendChartSource_(tr, body);
  unpaidStyleTrendSheet_(tr);
  if (body.length) {
    var spark2 = [];
    var si;
    for (si = 0; si < body.length; si++) spark2.push([unpaidSparkFormula_(head + 1 + si)]);
    tr.getRange(head + 1, 6, body.length, 1).setFormulas(spark2);
  }
  if (latestRate !== '' && latestRate < 0.5) {
    tr.getRange(3, 9, 1, 2).setBackground(t.blood).setFontColor(t.paper);
  }
  return unpaidEnsureTrendCharts_(tr);
}

function unpaidFillTrendAll_(ss) {
  ss = ss || openWorkspaceSpreadsheet_();
  var tr = unpaidEnsureTrendSheet_(ss);
  try {
    var src = SpreadsheetApp.openById(UNPAID_SOURCE_ID_);
    var months = unpaidSourceMonthTabs_(src);
    var body = [];
    var scanned = [];
    var totals = { n: 0, sp: 0, sr: 0, nr: 0 };
    var latest = null;
    var i;
    for (i = 0; i < months.length; i++) {
      var name = months[i].name;
      var st = { n: 0, sp: 0, sr: 0, nr: 0, kone: { pay: 0, rec: 0 }, ktwo: { pay: 0, rec: 0 }, kbad: { pay: 0, rec: 0 } };
      try {
        var vals = unpaidReadMonthValuesFast_(months[i].sh);
        if (vals) st = unpaidStatsFromValues_(vals);
      } catch (eM) {
        scanned.push({ name: name, ok: false, message: String(eM && eM.message ? eM.message : eM) });
        body.push(unpaidTrendRowFromStats_(name, st));
        continue;
      }
      body.push(unpaidTrendRowFromStats_(name, st));
      totals.n += st.n;
      totals.sp += st.sp;
      totals.sr += st.sr;
      totals.nr += st.nr;
      latest = { name: name, st: st };
      scanned.push({ name: name, ok: true, n: st.n, sp: st.sp, sr: st.sr });
    }
    var painted = unpaidPaintTrend_(tr, body, totals, latest);
    var bNow = unpaidTrendTableBounds_(tr);
    return {
      ok: true,
      months: scanned.length,
      filled: scanned.filter(function (x) { return x.ok; }).length,
      latest: latest ? latest.name : '',
      totals: totals,
      sheet: tr,
      gid: tr.getSheetId(),
      charts: tr.getCharts().length,
      chartTop: painted && painted.chartTop,
      chartErrors: painted && painted.errors,
      srcRows: painted && painted.srcRows,
      headRow: bNow.head,
      tableLast: bNow.last,
      a1: String(tr.getRange('A1').getDisplayValue() || ''),
      d1: String(tr.getRange('D1').getDisplayValue() || ''),
      a6: String(tr.getRange('A6').getDisplayValue() || ''),
      a7: String(tr.getRange('A7').getDisplayValue() || ''),
      c3: String(tr.getRange('C3').getDisplayValue() || ''),
      i3: String(tr.getRange('I3').getDisplayValue() || ''),
      scanned: scanned
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err), sheet: tr };
  }
}

function setupUnpaidTrend_(ss, options) {
  var out = unpaidFillTrendAll_(ss);
  return (out && out.sheet) || unpaidEnsureTrendSheet_(ss);
}

function applyUnpaidNotes_(sh, tr) {
  try {
    sh.getRange('A1').setNote('レジ送信・ゲートストップ・SMS のチェックと会員名は Workspace から変えられます。変えたら元の未納管理ドライブへ自動で送ります。支払額など元データは変えません。');
    sh.getRange('B1').setNote('元ファイルの月タブ名です。いまの月が一番上、あとは新しい順です。');
    sh.getRange('D1').setNote('会員番号があり、支払額が1円以上の行の件数。');
    sh.getRange('E1').setNote('支払額の合計（手数料は含まない）。');
    sh.getRange('F1').setNote('回収金額の合計。空でも右隣に「〇〇入金」があれば支払額を回収済みとみなす。');
    sh.getRange('G1').setNote('回収額÷未納総額（支払額ベース）。このシート上部で血色に反転している数字。');
    sh.getRange('J1').setNote('A列区分が「1ヶ月未納」の行だけを集計した回収率。');
    sh.getRange('K1').setNote('A列区分が「2ヶ月未納」の行だけを集計した回収率。');
    sh.getRange('L1').setNote('A列に「貸倒」または「貸し倒れ」を含む行の回収率。該当者がいなければ対象なし。');
  } catch (eN) {}
  if (!tr) return;
  try {
    tr.getRange(UNPAID_TREND_HEAD_ROW_, 5).setNote('その月全体の回収率（回収額÷未納総額＝支払額ベース）。右のバーと同じ数字です。');
    tr.getRange(UNPAID_TREND_HEAD_ROW_, 6).setNote('左の「回収率」を0〜100%の横棒にしたもの。貸倒や1ヶ月未納の率ではありません。');
    tr.getRange(UNPAID_TREND_HEAD_ROW_, 11).setNote('A列が貸倒／貸し倒れの行の回収率。該当がなければ空欄（グラフが壊れないように「対象なし」は出さない）。');
    tr.getRange('A4').setNote('各月タブの単純合計。繰越があると二重計上の可能性があります。');
  } catch (eT) {}
}

function fillUnpaidNow_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName(UNPAID_SHEET_);
    if (!sh) return { ok: false, message: 'missing 未納管理' };
    try { ensureUnpaidEditTrigger_(); } catch (eT) {}
    var filled = unpaidFillFromSource_(sh);
    try { styleUnpaidView_(sh); } catch (eV) {}
    try { styleUnpaidDashboard_(sh); } catch (eD) {}
    try { unpaidHideNoiseCols_(sh); } catch (eH) {}
    try { sh.setColumnWidths(14, 7, 44); } catch (eW) {}
    var tr = ss.getSheetByName(UNPAID_TREND_SHEET_);
    var trend = null;
    try { trend = unpaidFillTrendAll_(ss); tr = trend && trend.sheet ? trend.sheet : tr; } catch (eTr) {
      trend = { ok: false, message: String(eTr && eTr.message ? eTr.message : eTr) };
      if (tr && filled && filled.ok && filled.stats) {
        try { unpaidWriteTrendMonth_(tr, filled.month, filled.stats); } catch (eOne) {}
      }
    }
    return {
      ok: !!(filled && filled.ok),
      filled: filled,
      trend: trend,
      dash: sh.getRange('D1:M3').getDisplayValues(),
      head: sh.getRange(UNPAID_DATA_ROW_, 1, 12, 12).getDisplayValues()
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function styleUnpaidNow_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName(UNPAID_SHEET_);
    if (!sh) return { ok: false, message: 'missing 未納管理' };
    try { ensureUnpaidEditTrigger_(); } catch (eT) {}
    styleUnpaidView_(sh);
    styleUnpaidDashboard_(sh);
    unpaidHideNoiseCols_(sh);
    sh.setColumnWidth(1, 72);
    sh.setColumnWidth(2, 52);
    sh.setColumnWidth(3, 100);
    sh.setColumnWidth(4, 148);
    sh.setColumnWidths(5, 4, 88);
    sh.setColumnWidth(9, 110);
    sh.setColumnWidths(10, 2, 92);
    sh.setColumnWidths(14, 7, 44);
    sh.setFrozenColumns(4);
    sh.setFrozenRows(UNPAID_DATA_ROW_ + 1);
    sh.setHiddenGridlines(true);
    sh.setTabColor(dnTheme_().blood);
    try { sh.showSheet(); } catch (eShow) {}
    return {
      ok: true,
      d2: String(sh.getRange('D2').getDisplayValue() || ''),
      a5c: sh.getRange('A5:D8').getDisplayValues()
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/**
 * 未納管理ドライブ【経堂】の月タブを、B1 の年月選択で切り替えて表示するシートを作る。
 * 1〜3行目＝選択月の集計、5行目〜＝元シートと同じマス（チェック可）。チェックは元ファイルへ保存。
 */
function setupUnpaidView_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName(UNPAID_SHEET_);
    if (!sh) sh = ss.insertSheet(UNPAID_SHEET_);
    var tz = 'Asia/Tokyo';
    var now = new Date();
    var current = Utilities.formatDate(now, tz, 'yy') + '年' + Number(Utilities.formatDate(now, tz, 'M')) + '月';
    var options = unpaidMonthOptions_();
    var b1 = String(sh.getRange('B1').getDisplayValue() || '');
    if (options.indexOf(b1) < 0) b1 = current;

    if (sh.getMaxRows() < 200) sh.insertRowsAfter(sh.getMaxRows(), 200 - sh.getMaxRows());
    if (sh.getMaxColumns() < UNPAID_COLS_) sh.insertColumnsAfter(sh.getMaxColumns(), UNPAID_COLS_ - sh.getMaxColumns());
    sh.setFrozenRows(0);
    var head = sh.getRange(1, 1, UNPAID_DATA_ROW_ + 1, sh.getMaxColumns());
    head.breakApart();
    head.clear();
    head.clearDataValidations();

    var trend = setupUnpaidTrend_(ss, options);

    sh.getRange('A1').setValue('年月');
    sh.getRange('B1').setNumberFormat('@').setValue(b1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(options, true).setAllowInvalid(false).build());
    sh.getRange('A2').setFormula('=HYPERLINK("https://docs.google.com/spreadsheets/d/' + UNPAID_SOURCE_ID_ + '/edit","元の未納管理ドライブを開く ↗")');
    sh.getRange('A3').setFormula('=HYPERLINK("#gid=' + trend.getSheetId() + '","月別の推移を見る ↗")');
    var filled = { ok: false };
    var triggerErr = '';
    try { ensureUnpaidEditTrigger_(); } catch (eT) { triggerErr = String(eT && eT.message ? eT.message : eT); }
    try { filled = unpaidFillFromSource_(sh); } catch (eFill) { filled = { ok: false, message: String(eFill) }; }
    if (filled && filled.ok && filled.stats) {
      try { unpaidWriteTrendMonth_(trend, filled.month, filled.stats); } catch (eTr) {}
    }

    styleUnpaidView_(sh);
    styleUnpaidDashboard_(sh);
    applyUnpaidNotes_(sh, trend);
    sh.setFrozenRows(UNPAID_DATA_ROW_ + 1);
    sh.setHiddenGridlines(true);
    try { sh.showSheet(); } catch (eShowU) {}
    sh.setColumnWidth(1, 72);
    sh.setColumnWidth(2, 52);
    sh.setRowHeightsForced(UNPAID_DATA_ROW_ + 2, Math.min(Math.max(sh.getLastRow() - UNPAID_DATA_ROW_ - 1, 10), 400), 24);
    ss.setActiveSheet(trend);
    ss.moveActiveSheet(sh.getIndex() + (trend.getIndex() < sh.getIndex() ? 0 : 1));
    ss.setActiveSheet(sh);
    sh.setColumnWidth(3, 100);
    sh.setColumnWidth(4, 148);
    sh.setColumnWidths(5, 4, 88);
    sh.setColumnWidth(9, 110);
    sh.setColumnWidths(10, 2, 92);
    sh.setColumnWidths(12, 5, 100);
    sh.setColumnWidths(14, 7, 44);
    sh.setColumnWidth(25, 104);
    sh.setColumnWidth(26, 92);
    unpaidHideNoiseCols_(sh);
    sh.setFrozenColumns(4);
    sh.setTabColor(dnTheme_().blood);
    applyHubTabColors_(ss);
    try { PropertiesService.getDocumentProperties().setProperty(UNPAID_SCAN_PROP_, '1'); } catch (eProp) {}
    try { PropertiesService.getDocumentProperties().setProperty(FOUR_COLOR_PROP_, '1'); } catch (eProp2) {}
    SpreadsheetApp.flush();
    return {
      ok: true,
      b1: sh.getRange('B1').getDisplayValue(),
      filled: filled,
      triggerErr: triggerErr,
      dash: sh.getRange('D1:M3').getDisplayValues(),
      head: sh.getRange(UNPAID_DATA_ROW_, 1, 14, 10).getDisplayValues(),
      trend: trend.getRange(1, 1, 16, 11).getDisplayValues()
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

var KENGAKU_SOURCE_ID_ = '1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY';
var KENGAKU_JOIN_WINDOW_DAYS_ = 180;

var KENGAKU_JOIN_LAST_ROW_ = 400;

function kengakuImportFormula_() {
  return '=IMPORTRANGE("' + KENGAKU_SOURCE_ID_ + '","見学体験申請!A2:I")';
}

/**
 * 見学体験申請の K 列（入会日）。申込の前日〜180日以内で、氏名かメールが一致する最初の入会日。
 * 経堂_入会（氏名・メール）と入会者一覧（メール）の早い方。J 列はこの K 列から入会／未入会を出す。
 * 元フォームの見出し行は取り込まない（A2 からデータ）。配列は 400 行で止めてシートを膨らませない。
 */
function kengakuJoinDateFormula_() {
  var d = KENGAKU_JOIN_WINDOW_DAYS_;
  var last = KENGAKU_JOIN_LAST_ROW_;
  return '={"入会日";ARRAYFORMULA(LET(' +
    "jd,'経堂_入会'!A2:A," +
    "jn,REGEXREPLACE('経堂_入会'!B2:B&\"\",\"[\\s　]\",\"\")," +
    "jm,LOWER(TRIM('経堂_入会'!F2:F&\"\"))," +
    "ld,'" + JOIN_LIST_SHEET_ + "'!A3:A," +
    "lm,LOWER(TRIM('" + JOIN_LIST_SHEET_ + "'!C3:C&\"\"))," +
    'MAP(A2:A' + last + ',B2:B' + last + ',C2:C' + last + ',D2:D' + last + ',LAMBDA(t,k,n,m,' +
    'IF(OR(t="",NOT(REGEXMATCH(k&"","見学|体験"))),"",IFERROR(LET(' +
    'nn,REGEXREPLACE(n&"","[\\s　]",""),mm,LOWER(TRIM(m&"")),lo,INT(t)-1,hi,t+' + d + ',' +
    'djoin,IFERROR(MIN(FILTER(jd,jd>=lo,jd<=hi,((nn<>"")*(jn=nn)+(mm<>"")*(jm=mm))>0)),0),' +
    'dlist,IF(mm="",0,IFERROR(MIN(FILTER(ld,ld>=lo,ld<=hi,lm=mm)),0)),' +
    'IF(djoin+dlist=0,"",IF(djoin=0,dlist,IF(dlist=0,djoin,MIN(djoin,dlist))))),""))))))}';
}

/** 見学体験申請の J 列（入会／未入会）。K 列に入会日があれば「入会」 */
function kengakuJoinFormula_() {
  var last = KENGAKU_JOIN_LAST_ROW_;
  return '={"入会";ARRAYFORMULA(IF(REGEXMATCH(B2:B' + last + '&"","見学|体験"),IF(K2:K' + last + '<>"","入会","未入会"),""))}';
}

function repairKengakuMirror_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName('見学体験申請');
    if (!sh) return { ok: false, message: 'missing 見学体験申請' };
    if (sh.getMaxColumns() < 12) sh.insertColumnsAfter(sh.getMaxColumns(), 12 - sh.getMaxColumns());
    sh.getRange(2, 1).setFormula(kengakuImportFormula_());
    sh.getRange(1, 10).setFormula(kengakuJoinFormula_());
    sh.getRange(1, 11).setFormula(kengakuJoinDateFormula_());
    var src = hubSourceForName_('見学体験申請');
    if (src) {
      hubPaintOpenSourceCell_(sh, 1, 12, 1, src);
      sh.setColumnWidth(12, 168);
    }
    try { restyleHeaderBody_(sh); } catch (eSt) {}
    sh.getRange(1, 10).setFormula(kengakuJoinFormula_());
    sh.getRange(1, 11).setFormula(kengakuJoinDateFormula_());
    var n = Math.min(Math.max(sh.getMaxRows() - 1, 1), KENGAKU_JOIN_LAST_ROW_);
    sh.getRange(2, 10, n, 2).setHorizontalAlignment('center');
    sh.getRange(2, 11, n, 1).setNumberFormat('yyyy/mm/dd');
    SpreadsheetApp.flush();
    return {
      ok: true,
      a2: String(sh.getRange('A2').getFormula() || ''),
      b2: String(sh.getRange('B2').getDisplayValue() || ''),
      j2: String(sh.getRange('J2').getDisplayValue() || ''),
      k2: String(sh.getRange('K2').getDisplayValue() || ''),
      last: sh.getLastRow()
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/**
 * 経堂_入会 / 経堂_退会（受付状況表 入会・退会_データの IMPORTRANGE）を非表示のデータシートとして白黒に整える。
 * 値・数式には触れない。
 */
function styleMembershipMirrors_() {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var out = [];
    ['経堂_入会', '経堂_退会'].forEach(function (name) {
      var sh = ss.getSheetByName(name);
      if (!sh) { out.push({ name: name, ok: false }); return; }
      if (!/IMPORTRANGE/i.test(String(sh.getRange(1, 1).getFormula() || ''))) {
        out.push({ name: name, ok: false, message: 'A1 が IMPORTRANGE ではないため未処理' });
        return;
      }
      var maxR = sh.getMaxRows();
      var cols = Math.max(sh.getLastColumn(), 1);
      sh.hideSheet();
      sh.setHiddenGridlines(false);
      sh.setFrozenRows(1);
      sh.setTabColor(hubTabColorFor_(name));
      var dnMem = dnTheme_();
      sh.getRange(1, 1, maxR, sh.getMaxColumns())
        .setBorder(false, false, false, false, false, false)
        .setBackground(dnMem.paper)
        .setFontColor(dnMem.ink)
        .setFontFamily('Noto Sans JP').setFontStyle('italic')
        .setFontSize(10)
        .setVerticalAlignment('middle');
      sh.getRange(1, 1, 1, cols)
        .setBackground(dnMem.ink)
        .setFontColor(dnMem.paper)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      sh.setRowHeight(1, 28);
      sh.getRange(2, 1, maxR - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
      sh.setColumnWidth(1, 140);
      sh.setColumnWidth(2, 130);
      sh.setColumnWidth(3, 95);
      sh.setColumnWidth(4, 95);
      if (cols >= 5) sh.setColumnWidth(5, 150);
      if (cols >= 6) sh.setColumnWidth(6, 100);
      out.push({ name: name, ok: true, rows: sh.getLastRow() });
    });
    return { ok: true, sheets: out };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/** 経堂マスタ R4: 見学体験一覧。全件・新しい順 */
function kengakuMasterListFormula_() {
  return "=IFERROR(QUERY({'見学体験申請'!A2:C," +
    "ARRAYFORMULA(SUBSTITUTE(TO_TEXT('見学体験申請'!H2:H),\"-\",\"/\"))," +
    "ARRAYFORMULA(TO_TEXT('見学体験申請'!I2:I))," +
    "'見学体験申請'!J2:J}," +
    "\"select Col1,Col2,Col3,Col4,Col6 where Col1 is not null order by Col1 desc\",0),\"\")";
}

/** 経堂マスタ 見学見出し: 当月の入会数/申込数（率）。一覧自体は全件 */
function kengakuMasterTitleFormula_(monthA1) {
  monthA1 = monthA1 || 'AV5';
  var monthRange =
    "'見学体験申請'!A2:A,\">=\"&$" + monthA1 + ",'見学体験申請'!A2:A,\"<\"&EDATE($" + monthA1 + ",1)";
  var total =
    '(COUNTIFS(' + monthRange + ",'見学体験申請'!B2:B,\"見学\")+COUNTIFS(" + monthRange + ",'見学体験申請'!B2:B,\"体験\"))";
  var joined = 'COUNTIFS(' + monthRange + ",'見学体験申請'!J2:J,\"入会\")";
  return '="見学体験　"&COUNTA(\'見学体験申請\'!A2:A)&"件　当月入会 "&' + joined + '&"/"&' + total +
    '&"（"&IFERROR(TEXT(' + joined + '/' + total + ',"0%"),"-")&"）"';
}

function masterReviewListFormula_() {
  return "=IFERROR(QUERY('口コミ_経堂'!A3:W,\"select Col1,Col5,Col6,Col22 where Col1 is not null order by Col1 desc\",0),\"\")";
}

/**
 * 見学体験申請を元フォームからの IMPORTRANGE に切替え（バックアップ作成）、J列に入会判定、
 * 経堂マスタ R:W に今月の入会率と○列を出す。confirm=yes のときだけ書き込む。
 */
function applyKengakuMasterFormulas_(master) {
  var monthA1 = masterMonthCellA1_(master);
  var col = masterFindRowLabel_(master, 1, '見学体験');
  if (!col) col = 22;
  master.getRange(1, col).setFormula(kengakuMasterTitleFormula_(monthA1));
  master.getRange(4, col).setFormula(kengakuMasterListFormula_());
  return col;
}

function setupKengakuJoinLive_(confirm) {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var sh = ss.getSheetByName('見学体験申請');
    var master = ss.getSheetByName('経堂マスタ');
    if (!sh || !master) return { ok: false, message: 'sheet missing' };
    var plan = {
      kengaku: 'A1:I1 見出しは維持 / A2:I の値→ IMPORTRANGE / J列=入会判定',
      master: 'R1 タイトルに今月の入会率 / R4 の QUERY に入会列を追加（W列に○）/ W3 見出し「入会」'
    };
    if (confirm === 'r1only') {
      var kColR = applyKengakuMasterFormulas_(master);
      SpreadsheetApp.flush();
      return { ok: true, r1: master.getRange(1, kColR).getDisplayValue() };
    }
    if (confirm === 'emailJoin') {
      var joinSh = ss.getSheetByName('経堂_入会');
      var a1 = joinSh ? String(joinSh.getRange('A1').getFormula() || '') : '';
      if (a1.indexOf('入会・退会_データ!A1:E"') >= 0) {
        joinSh.getRange('A1').setFormula(a1.replace('入会・退会_データ!A1:E"', '入会・退会_データ!A1:F"'));
      }
      var leaveSh = ss.getSheetByName('経堂_退会');
      var leaveA1 = leaveSh ? String(leaveSh.getRange('A1').getFormula() || '') : '';
      if (leaveA1.indexOf('入会・退会_データ!G1:L"') >= 0) {
        leaveSh.getRange('A1').setFormula(leaveA1.replace('入会・退会_データ!G1:L"', '入会・退会_データ!G1:M"'));
        leaveSh.setColumnWidth(7, 220);
      }
      var beforeJoined = sh.getRange('J2:J').getDisplayValues().filter(function (r) { return r[0] === '○' || r[0] === '入会'; }).length;
      sh.getRange(1, 10).setFormula(kengakuJoinFormula_());
      SpreadsheetApp.flush();
      if (joinSh) joinSh.setColumnWidth(6, 220);
      return {
        ok: true,
        mirrorA1: joinSh ? joinSh.getRange('A1').getFormula() : null,
        mirrorF1: joinSh ? joinSh.getRange('F1:F3').getDisplayValues() : null,
        joinedBefore: beforeJoined,
        joinedAfter: sh.getRange('J2:J').getDisplayValues().filter(function (r) { return r[0] === '入会'; }).length,
        notJoined: sh.getRange('J2:J').getDisplayValues().filter(function (r) { return r[0] === '未入会'; }).length,
        r1: master.getRange('R1').getDisplayValue()
      };
    }
    if (confirm === 'leaveList') {
      var leaveF = "=IFERROR(QUERY('経堂_退会'!A2:F,\"select Col2, Col1 where Col1 >= date '\"&TEXT(TODAY(),\"yyyy-mm-dd\")&\"' " +
        "and Col1 < date '\"&TEXT(TODAY()+1,\"yyyy-mm-dd\")&\"' and Col6 <> true and Col4 <> '法人会員' " +
        "and Col3 = '\"&TEXT(TODAY(),\"yyyy年m月\")&\"' order by Col1 desc limit 6\",0),\"\")";
      var before = master.getRange('D9').getFormula();
      master.getRange('D9').setFormula(leaveF);
      SpreadsheetApp.flush();
      return { ok: true, before: before, list: master.getRange('B9:E14').getDisplayValues() };
    }
    if (confirm === 'joinDate') {
      sh.getRange(1, 11).setFormula(kengakuJoinDateFormula_());
      sh.getRange(1, 10).setFormula(kengakuJoinFormula_());
      sh.getRange(2, 11, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd').setHorizontalAlignment('center');
      try { sh.getRange(1, 10).copyFormatToRange(sh, 11, 11, 1, 1); } catch (eF) {}
      SpreadsheetApp.flush();
      var jd = sh.getRange('J2:K').getDisplayValues();
      return {
        ok: true,
        joined: jd.filter(function (r) { return r[0] === '入会'; }).length,
        notJoined: jd.filter(function (r) { return r[0] === '未入会'; }).length,
        withDate: jd.filter(function (r) { return r[1]; }).length,
        k1: sh.getRange('K1').getDisplayValue(),
        sample: sh.getRange('A185:K192').getDisplayValues().map(function (r) { return [r[0], r[1], r[9], r[10]]; }),
        r1: master.getRange('R1').getDisplayValue()
      };
    }
    if (confirm === 'joinLabel') {
      sh.getRange(1, 10).setFormula(kengakuJoinFormula_());
      applyKengakuMasterFormulas_(master);
      ['経堂_入会', '経堂_退会'].forEach(function (n) {
        var m = ss.getSheetByName(n);
        if (m) m.hideSheet();
      });
      SpreadsheetApp.flush();
      var jv = sh.getRange('J2:J').getDisplayValues();
      return {
        ok: true,
        joined: jv.filter(function (r) { return r[0] === '入会'; }).length,
        notJoined: jv.filter(function (r) { return r[0] === '未入会'; }).length,
        circles: jv.filter(function (r) { return r[0] === '○'; }).length,
        r1: master.getRange('R1').getDisplayValue(),
        r4: master.getRange('R4:W8').getDisplayValues()
      };
    }
    if (confirm === 'r4fix') {
      master.getRange('R4').setFormula(kengakuMasterListFormula_());
      SpreadsheetApp.flush();
      return { ok: true, rows: master.getRange('R4:W7').getDisplayValues() };
    }
    if (confirm !== 'yes') return { ok: true, dryRun: true, plan: plan };

    var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MMdd_HHmm');
    var backup = sh.copyTo(ss).setName('見学体験申請_backup_' + stamp);
    backup.hideSheet();
    var masterBackup = master.copyTo(ss).setName('経堂マスタ_backup_' + stamp);
    masterBackup.hideSheet();

    var lastRow = Math.max(sh.getLastRow(), 2);
    sh.getRange(2, 1, lastRow - 1, 9).clearContent();
    sh.getRange(2, 1).setFormula(kengakuImportFormula_());
    sh.getRange(1, 10).setFormula(kengakuJoinFormula_());
    sh.getRange(1, 11).setFormula(kengakuJoinDateFormula_());
    sh.getRange(1, 10).copyFormatToRange(sh, 10, 10, 1, 1);
    try { sh.getRange(1, 9).copyFormatToRange(sh, 10, 10, 1, 1); } catch (eFmt) {}
    sh.getRange(2, 10, sh.getMaxRows() - 1, 1).setHorizontalAlignment('center');

    master.getRange('R1').setFormula(kengakuMasterTitleFormula_());
    var r4 = master.getRange('R4').getFormula();
    var r4new = kengakuMasterListFormula_();
    if (r4new !== r4) master.getRange('R4').setFormula(r4new);
    master.getRange('V3').copyFormatToRange(master, 23, 23, 3, 3);
    master.getRange('W3').setValue('入会');
    master.getRange('V4:V40').copyFormatToRange(master, 23, 23, 4, 40);
    master.getRange('W4:W40').setHorizontalAlignment('center');

    SpreadsheetApp.flush();
    return {
      ok: true,
      backups: [backup.getName(), masterBackup.getName()],
      r4Changed: r4new !== r4,
      r1: master.getRange('R1').getDisplayValue(),
      j1: sh.getRange('J1').getDisplayValue(),
      rows: sh.getLastRow()
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/**
 * 読み取り専用の試算: 見学体験申請 → 入会（経堂_入会 の氏名 / 入会者一覧 のメール）
 * 申込日以降 days 日以内に、名前かメールが一致した入会があれば「入会」とみなす。
 */
function previewTourToJoinFunnel_(days) {
  try {
    days = Math.max(7, Math.min(Number(days) || 60, 365));
    var ss = openWorkspaceSpreadsheet_();
    var norm = function (s) {
      return String(s || '').replace(/[\s\u3000]/g, '').normalize('NFKC').toLowerCase();
    };
    var toDate = function (v) {
      if (v instanceof Date) return v;
      var d = new Date(String(v || '').replace(/\//g, '-').replace(' ', 'T'));
      return isNaN(d.getTime()) ? null : d;
    };

    var joins = [];
    var joinSh = ss.getSheetByName('経堂_入会');
    if (joinSh && joinSh.getLastRow() > 1) {
      joinSh.getRange(2, 1, joinSh.getLastRow() - 1, 4).getValues().forEach(function (r) {
        var d = toDate(r[0]);
        if (d && r[1]) joins.push({ at: d, name: norm(r[1]), email: '', kind: String(r[3] || '') });
      });
    }
    var listSh = ss.getSheetByName(JOIN_LIST_SHEET_);
    var collectJoinList = function (startCol) {
      if (!listSh || listSh.getLastRow() < JOIN_LIST_BODY_START_) return;
      var n = listSh.getLastRow() - 2;
      listSh.getRange(JOIN_LIST_BODY_START_, startCol, n, 3).getValues().forEach(function (r) {
        if (String(r[0] || '') === '入会日') return;
        var d = toDate(r[0]);
        if (d && (r[1] || r[2])) joins.push({ at: d, name: norm(r[1]), email: norm(r[2]), kind: '' });
      });
    };
    collectJoinList(1);
    collectJoinList(6);

    var tourSh = ss.getSheetByName('見学体験申請');
    var tours = tourSh.getRange(2, 1, Math.max(tourSh.getLastRow() - 1, 1), 9).getValues();
    var byMonth = {};
    var total = { tours: 0, withName: 0, joined: 0, byName: 0, byEmail: 0, daysSum: 0 };
    var samples = [];
    var limitMs = days * 86400000;
    tours.forEach(function (r) {
      var at = toDate(r[0]);
      var kind = String(r[1] || '');
      if (!at || !/見学|体験/.test(kind)) return;
      var name = norm(r[2]);
      var email = norm(r[3]);
      var ym = Utilities.formatDate(at, 'Asia/Tokyo', 'yyyy/MM');
      var m = byMonth[ym] || (byMonth[ym] = { tours: 0, joined: 0, 見学: 0, 体験: 0, 見学入会: 0, 体験入会: 0 });
      m.tours++; m[kind]++; total.tours++;
      if (!name && !email) return;
      total.withName++;
      var hit = null;
      var how = '';
      for (var i = 0; i < joins.length; i++) {
        var j = joins[i];
        var diff = j.at.getTime() - at.getTime();
        if (diff < -86400000 || diff > limitMs) continue;
        if (email && j.email && email === j.email) { hit = j; how = 'email'; break; }
        if (name && j.name && name === j.name) { hit = j; how = 'name'; }
      }
      if (hit) {
        m.joined++; m[kind + '入会']++; total.joined++;
        if (how === 'email') total.byEmail++; else total.byName++;
        var dd = Math.max(0, Math.round((hit.at.getTime() - at.getTime()) / 86400000));
        total.daysSum += dd;
        if (samples.length < 8) samples.push({ tour: ym, kind: kind, how: how, days: dd, joinKind: hit.kind });
      }
    });
    return {
      ok: true,
      windowDays: days,
      joinsLoaded: joins.length,
      total: total,
      avgDaysToJoin: total.joined ? Math.round(total.daysSum / total.joined * 10) / 10 : null,
      byMonth: byMonth,
      samples: samples
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/** 読み取り専用: 見出し行・先頭2行・末尾N行と A1/F1 の数式を返す（全体把握用） */
function peekSheetRows_(name, tail) {
  try {
    var ss = openWorkspaceSpreadsheet_();
    var names = name ? [name] : ss.getSheets().map(function (s) { return s.getName(); });
    tail = Math.max(1, Math.min(Number(tail) || 8, 40));
    var out = [];
    for (var i = 0; i < names.length; i++) {
      var sh = ss.getSheetByName(names[i]);
      if (!sh) { out.push({ name: names[i], exists: false }); continue; }
      var lastRow = sh.getLastRow();
      var lastCol = Math.min(Math.max(sh.getLastColumn(), 1), 30);
      var head = lastRow >= 1 ? sh.getRange(1, 1, Math.min(lastRow, 3), lastCol).getDisplayValues() : [];
      var tailRows = [];
      if (lastRow > 3) {
        var start = Math.max(4, lastRow - tail + 1);
        tailRows = sh.getRange(start, 1, lastRow - start + 1, lastCol).getDisplayValues();
      }
      var formulas = {};
      var topFormulas = sh.getRange(1, 1, 1, lastCol).getFormulas()[0];
      for (var c = 0; c < topFormulas.length; c++) {
        if (topFormulas[c]) formulas[columnLetter_(c + 1) + '1'] = String(topFormulas[c]).slice(0, 200);
      }
      out.push({
        name: names[i],
        hidden: sh.isSheetHidden(),
        rows: lastRow,
        cols: sh.getLastColumn(),
        topFormulas: formulas,
        head: head,
        tail: tailRows
      });
    }
    return { ok: true, sheets: out };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
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
 * 高負荷・未提供時は候補モデルを順に試す。
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
    var models = [
      'gemini-2.0-flash-lite',
      'gemini-2.0-flash',
      'gemini-2.5-flash-lite',
      'gemini-2.5-flash',
      'gemini-flash-latest',
      'gemini-1.5-flash',
    ];
    var prompt =
      '以下はフィットネス施設のスタッフ日報「所感」欄の下書きです。ビジネスメール向けの敬語に整え、誤字脱字を修正し、300文字以内で簡潔にまとめてください。事実と意味は変えないでください。出力は所感の本文のみ（説明・見出し・引用符は不要）。\n\n' +
      String(rawText);
    var lastMessage = '添削に失敗しました。';
    for (var i = 0; i < models.length; i++) {
      var url =
        'https://generativelanguage.googleapis.com/v1beta/models/' +
        encodeURIComponent(models[i]) +
        ':generateContent?key=' +
        encodeURIComponent(key);
      var body = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.4, maxOutputTokens: 1024 },
      };
      var res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(body),
        muteHttpExceptions: true,
      });
      var code = res.getResponseCode();
      var json = {};
      try {
        json = JSON.parse(res.getContentText());
      } catch (eParse) {
        lastMessage = 'Gemini の応答が不正です（' + code + '）';
        continue;
      }
      if (code !== 200) {
        lastMessage = (json.error && json.error.message) || 'API エラー（コード ' + code + '）';
        if (
          code === 429 ||
          code === 503 ||
          /high demand|resource.?exhausted|unavailable|try again|not found|not supported/i.test(
            lastMessage
          )
        ) {
          continue;
        }
        return { ok: false, message: lastMessage };
      }
      var text =
        json.candidates &&
        json.candidates[0] &&
        json.candidates[0].content &&
        json.candidates[0].content.parts &&
        json.candidates[0].content.parts[0] &&
        json.candidates[0].content.parts[0].text;
      if (text) return { ok: true, text: String(text).trim(), model: models[i] };
      lastMessage = '返答を取得できませんでした。';
    }
    if (/high demand|try again later/i.test(lastMessage)) {
      lastMessage =
        'AIが混み合っています。十数秒待ってからもう一度「添削」を押してください。';
    }
    return { ok: false, message: lastMessage };
  } catch (e) {
    console.error(e);
    return { ok: false, message: String(e.message || e) };
  }
}

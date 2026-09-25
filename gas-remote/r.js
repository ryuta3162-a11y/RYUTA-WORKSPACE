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

function setupPromoImport_() {
  try {
    var source = SpreadsheetApp.openById(PROMO_SOURCE_ID_);
    var dest = openWorkspaceSpreadsheet_();
    var styled = [];
    var imported = [];
    var sheets = source.getSheets();

    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
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
  sheet.setTabColor('#222222');
  sheet.setFrozenRows(1);

  var headerRange = sheet.getRange(1, 1, 1, usedCols);
  headerRange
    .setBackground('#111111')
    .setFontColor('#FFFFFF')
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setWrap(true);
  sheet.setRowHeight(1, 32);

  if (usedRows >= 2) {
    var body = sheet.getRange(2, 1, usedRows - 1, usedCols);
    body
      .setBackground('#FAFAFA')
      .setFontColor('#111111')
      .setFontFamily('Roboto Mono')
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
  sheet.setTabColor('#222222');
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
  var lastBody = Math.min(sheet.getMaxRows(), 1000);
  for (var h = 0; h < hdrs.length; h++) {
    var title = String(hdrs[h] || '');
    if (/リンク|URL|画像|写真/i.test(title)) sheet.setColumnWidth(h + 1, 260);
    if (/日時|申請|入会日|タイムスタンプ/.test(title)) {
      sheet.setColumnWidth(h + 1, 160);
      try {
        sheet.getRange(2, h + 1, lastBody - 1, 1).setNumberFormat('yyyy/mm/dd HH:mm');
      } catch (eFmt) {}
    }
    if (/メールアドレス|mail/i.test(title) && !/レクチャー|アンケート|付与/.test(title)) {
      sheet.setColumnWidth(h + 1, 220);
    }
    // 口コミのポイント列と同様：チェック用途の列
    if (isCheckboxHeader_(title)) {
      checkCols.push(h + 1);
      sheet.setColumnWidth(h + 1, 120);
    }
  }

  // 見た目だけ整える（値は IMPORTRANGE）
  try {
    sheet.getRange(1, 1, 1, cols)
      .setBackground('#111111')
      .setFontColor('#FFFFFF')
      .setFontFamily('Roboto Mono')
      .setFontSize(10)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 32);
    if (lastBody >= 2) {
      sheet.getRange(2, 1, lastBody - 1, cols)
        .setBackground('#FAFAFA')
        .setFontColor('#111111')
        .setFontFamily('Roboto Mono')
        .setFontSize(10)
        .setVerticalAlignment('middle');
    }
  } catch (eStyle) {}

  // TRUE/FALSE をチェックボックス表示＋付与済みは緑（口コミと同じ考え方）
  if (checkCols.length && lastBody >= 2) {
    try {
      var rules = sheet.getConditionalFormatRules() || [];
      for (var i = 0; i < checkCols.length; i++) {
        var col = checkCols[i];
        var colLetter = columnLetter_(col);
        var range = sheet.getRange(2, col, lastBody - 1, 1);
        range.setDataValidation(
          SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build()
        );
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied('=$' + colLetter + '2=TRUE')
            .setBackground('#E8F5E9')
            .setRanges([range])
            .build()
        );
      }
      sheet.setConditionalFormatRules(rules);
    } catch (eCheck) {}
  }
}

function isCheckboxHeader_(title) {
  var t = String(title || '');
  if (!t) return false;
  if (/メールアドレス|email/i.test(t) && !/レクチャー|アンケート|付与/.test(t)) return false;
  return /アンケート|付与済|ポイント付与|レクチャーメール|送信済|済フラグ/.test(t);
}

function stylePromoHubSheet_(sheet, imported) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#000000');
  sheet.setFrozenRows(1);

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
    .setBackground('#111111')
    .setFontColor('#FFFFFF')
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setFontWeight('bold');
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 6)
      .setBackground('#FAFAFA')
      .setFontColor('#111111')
      .setFontFamily('Roboto Mono')
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
function inspectSheetRange_(id, sheetName, a1, wantFormulas, wantFormat) {
  try {
    if (!id || !sheetName || !a1) return { ok: false, message: 'id, sheet, range required' };
    var ss = SpreadsheetApp.openById(id);
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, message: 'sheet not found' };
    var rng = sh.getRange(a1);
    return {
      ok: true,
      sheet: sheetName,
      range: a1,
      values: rng.getDisplayValues(),
      formulas: wantFormulas ? rng.getFormulas() : [],
      fontSizes: wantFormat ? rng.getFontSizes() : [],
      backgrounds: wantFormat ? rng.getBackgrounds() : []
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

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
    mirror.setTabColor('#1a1a1a');
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
      mirror.getRange(1, 1, 1, cols)
        .setBackground('#111111')
        .setFontColor('#FFFFFF')
        .setFontFamily('Roboto Mono')
        .setFontSize(10)
        .setFontWeight('bold');
      mirror.setRowHeight(1, 32);
      var lastBody = Math.min(mirror.getMaxRows(), 500);
      if (lastBody >= 2) {
        mirror.getRange(2, 1, lastBody - 1, cols)
          .setBackground('#FAFAFA')
          .setFontColor('#111111')
          .setFontFamily('Roboto Mono')
          .setFontSize(10)
          .setVerticalAlignment('middle');
        mirror.getRange('A2:A' + lastBody).setNumberFormat('yyyy/mm/dd HH:mm');
        mirror.getRange('W2:W' + lastBody).setNumberFormat('yyyy/mm/dd HH:mm');
        var rules = [];
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied('=$V2=TRUE')
            .setBackground('#E8F5E9')
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
      sh.getRange(1, 1, 1, 4)
        .setBackground('#111111')
        .setFontColor('#FFFFFF')
        .setFontFamily('Roboto Mono')
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
    sh.getRange(2, 1, 1, 4)
      .setBackground('#FAFAFA')
      .setFontColor('#111111')
      .setFontFamily('Roboto Mono')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sh.getRange(2, 4).setFontColor('#1A73E8');
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

/** マシンレクチャー／入会者一覧 → Workspace（同名シート・内容そのまま・IMPORTRANGE） */
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
      workspaceUrl: dest.getUrl(),
      note: 'シート名そのまま / 内容は IMPORTRANGE。初回はアクセス許可が必要な場合あり。元ブック未変更。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

var KENGAKU_SOURCE_ID_ = '1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY';
var KENGAKU_SOURCE_SHEET_ = '見学体験申請';
var KENGAKU_DEST_SHEET_ = '見学体験申請';

function setupKengakuImport_(ss) {
  try {
    var dest = ss || openWorkspaceSpreadsheet_();
    var existing = dest.getSheetByName(KENGAKU_DEST_SHEET_);
    if (existing) dest.deleteSheet(existing);
    var sh = dest.insertSheet(KENGAKU_DEST_SHEET_);
    sh.setHiddenGridlines(true);
    sh.setTabColor('#1B2838');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 9).setValues([[
      'タイムスタンプ', '区分', '名前', 'メール', '電話', '性別', '年代', '希望日', '時刻'
    ]]);
    var src = SpreadsheetApp.openById(KENGAKU_SOURCE_ID_).getSheetByName(KENGAKU_SOURCE_SHEET_);
    if (!src) return { ok: false, message: 'source sheet not found' };
    var lastRow = Math.max(src.getLastRow(), 1);
    var copied = src.getRange(1, 1, lastRow, 9).getDisplayValues();
    sh.getRange(2, 1, copied.length, 9).setValues(copied);
    sh.getRange(1, 1, 1, 9)
      .setBackground('#0F1419')
      .setFontColor('#FFFFFF')
      .setFontFamily('Meiryo')
      .setFontSize(10)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');
    sh.setRowHeight(1, 24);
    var lastBody = Math.min(sh.getMaxRows(), 400);
    sh.getRange(2, 1, lastBody - 1, 9)
      .setBackground('#FFFFFF')
      .setFontColor('#0F1419')
      .setFontFamily('Meiryo')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sh.getRange(2, 1, lastBody - 1, 1).setNumberFormat('yyyy/mm/dd HH:mm');
    sh.getRange(2, 8, lastBody - 1, 1).setNumberFormat('yyyy/mm/dd');
    sh.getRange(2, 9, lastBody - 1, 1).setNumberFormat('hh:mm');
    var widths = [150, 56, 110, 200, 110, 48, 64, 96, 56];
    for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);
    for (var r = 2; r <= Math.min(lastBody, 220); r++) sh.setRowHeight(r, 20);
    return {
      ok: true,
      sheet: KENGAKU_DEST_SHEET_,
      sourceId: KENGAKU_SOURCE_ID_,
      sourceSheet: KENGAKU_SOURCE_SHEET_,
      workspaceUrl: dest.getUrl() + '#gid=' + sh.getSheetId(),
      note: '元ブック未変更。値は setup 時にコピー。経堂マスタ Z6 に IMPORTRANGE あり。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/**
 * ALLDATA … 数字の司令塔（1枚）
 * 名簿・OP明細の全行は載せない。元ブックは読み取りのみ。
 */
var UKETSUKE_SOURCE_ID_ = '14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w';
var ALLDATA_SHEET_NAME_ = '経堂マスタ';
var OP_HELPER_SHEET_ = '経堂_OP';
var OP_LOG_HELPER_SHEET_ = '経堂_OPログ';
var NYUKAI_HELPER_SHEET_ = '経堂_入会';
var TAIKAI_HELPER_SHEET_ = '経堂_退会';
var OP_NAMES_ = [
  '安心サポート', '安心サポートVIP', '水素水', 'オンラインレッスン',
  '体組成計', '契約ロッカー1,500', 'レンタルマット', 'プロテイン12杯',
  'プロテイン無制限', 'プロテイン＋水素水', 'レンタルタオル', 'タンニング',
  'セルフエステ', 'ホットスタジオ', 'ヨガロッカー', 'ピラティスリフォーマー'
];

function setupAllData_() {
  try {
    var dest = openWorkspaceSpreadsheet_();
    ['ALLDATA', ALLDATA_SHEET_NAME_].forEach(function (name) {
      var existing = dest.getSheetByName(name);
      if (existing) dest.deleteSheet(existing);
    });
    setupOpHelperSheets_(dest);
    setupNyukaiHelperSheets_(dest);
    setupKengakuImport_(dest);
    var sh = dest.insertSheet(ALLDATA_SHEET_NAME_, 0);
    styleAllDataSheet_(sh);
    ensureKyodoMasterEditTrigger_(dest);
    var aiTrigger = ensureKyodoAiBriefTrigger_();
    SpreadsheetApp.flush();
    var brief = refreshKyodoAiBrief_();
    return {
      ok: true,
      sheet: ALLDATA_SHEET_NAME_,
      workspaceUrl: dest.getUrl() + '#gid=' + sh.getSheetId(),
      uketsukeId: UKETSUKE_SOURCE_ID_,
      aiBrief: brief,
      aiTrigger: aiTrigger,
      aiAuthUrl: kyodoAiAuthUrl_()
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function monthCountQuery_(sheetName, startExpr, endExpr) {
  var ref = "'" + String(sheetName).replace(/'/g, "''") + "'!A:A";
  return (
    '=IFERROR(COUNTA(QUERY(' + ref +
    ',"select A where A >= date \'"&TEXT(' + startExpr + ',"yyyy-mm-dd")&"\' and A < date \'"&TEXT(' +
    endExpr + ',"yyyy-mm-dd")&"\'",0)),0)'
  );
}

function irSum_(yyCell, a1) {
  return '=IFERROR(SUM(IMPORTRANGE($AB$1,' + yyCell + '&"!' + a1 + '")),)';
}

function irN_(yyCell, a1) {
  return '=IFERROR(N(IMPORTRANGE($AB$1,' + yyCell + '&"!' + a1 + '")),)';
}

function irNippoIfCurrent_(a1) {
  return (
    '=IF(TEXT($AB$5,"yymm")=TEXT(TODAY(),"yymm"),IFERROR(N(IMPORTRANGE($AB$1,"日報!' + a1 + '")),),)'
  );
}

function irNippoSumIfCurrent_(a1) {
  return (
    '=IF(TEXT($AB$5,"yymm")=TEXT(TODAY(),"yymm"),IFERROR(SUM(IMPORTRANGE($AB$1,"日報!' + a1 + '")),),)'
  );
}

function trendAt_(label, monthOffset) {
  var sh = "'" + String(KYODO_TREND_DEST_SHEET_).replace(/'/g, "''") + "'";
  return (
    '=IFERROR(INDEX(' + sh + '!$C$2:$N$115,' +
    'MATCH("' + String(label).replace(/"/g, '""') + '",' + sh + '!$B$2:$B$115,0),' +
    'MATCH(MONTH(EDATE($AB$5,' + monthOffset + ')),' + sh + '!$C$1:$N$1,0)),)'
  );
}

function irSumOff_(monthOffset, a1) {
  return (
    '=IFERROR(SUM(IMPORTRANGE($AB$1,TEXT(EDATE($AB$5,' + monthOffset + '),"yymm")&"!' + a1 + '")),)'
  );
}

function monthCountOff_(sheetName, monthOffset) {
  return monthCountQuery_(
    sheetName,
    'EDATE($AB$5,' + monthOffset + ')',
    'EDATE($AB$5,' + (monthOffset + 1) + ')'
  );
}

function nippoOrTrend_(nippoA1, label, monthOffset) {
  if (monthOffset !== 0) return trendAt_(label, monthOffset);
  var trend = trendAt_(label, 0).replace(/^=/, '');
  return (
    '=IF(TEXT($AB$5,"yymm")=TEXT(TODAY(),"yymm"),IFERROR(N(IMPORTRANGE($AB$1,"日報!' +
    nippoA1 +
    '")),),' +
    trend +
    ')'
  );
}

function paceFrom_(actualA1) {
  return (
    '=IF(TEXT($AB$5,"yymm")<>TEXT(TODAY(),"yymm"),,' +
    'IF(AND(ISNUMBER(' + actualA1 + '),' + actualA1 + '<>"",DAY(TODAY())>0),' +
    'ROUND(' + actualA1 + '/DAY(TODAY())*DAY(EOMONTH($AB$5,0)),0),))'
  );
}

function progressFormula_(r) {
  return (
    '=IF(AND(ISNUMBER(G' + r + '),ISNUMBER(H' + r + '),H' + r + '<>0),G' + r + '/H' + r +
    ',IF(AND(ISNUMBER(F' + r + '),ISNUMBER(H' + r + '),H' + r + '<>0),F' + r + '/H' + r + ',))'
  );
}

function planFromPrev_(r) {
  return '=IF(ISNUMBER(E' + r + '),E' + r + ',)';
}

function setupOpHelperSheets_(ss) {
  var op = ss.getSheetByName(OP_HELPER_SHEET_);
  if (!op) op = ss.insertSheet(OP_HELPER_SHEET_);
  op.clear();
  op.getRange(1, 1).setFormula(
    '=IMPORTRANGE("' + UKETSUKE_SOURCE_ID_ + '","OP集計!A1:E18")'
  );
  try { op.hideSheet(); } catch (e0) {}

  var log = ss.getSheetByName(OP_LOG_HELPER_SHEET_);
  if (!log) log = ss.insertSheet(OP_LOG_HELPER_SHEET_);
  log.clear();
  log.getRange(1, 1).setFormula(
    '=QUERY(IMPORTRANGE("' + UKETSUKE_SOURCE_ID_ + '","OP集計!I:N"),"select Col1, Col3, Col5",1)'
  );
  try { log.hideSheet(); } catch (e1) {}
}

function setupNyukaiHelperSheets_(ss) {
  var enroll = ss.getSheetByName(NYUKAI_HELPER_SHEET_);
  if (!enroll) enroll = ss.insertSheet(NYUKAI_HELPER_SHEET_);
  enroll.clear();
  enroll.getRange(1, 1).setFormula(
    '=IMPORTRANGE("' + UKETSUKE_SOURCE_ID_ + '","入会・退会_データ!A1:E")'
  );
  try { enroll.hideSheet(); } catch (e0) {}

  var withdraw = ss.getSheetByName(TAIKAI_HELPER_SHEET_);
  if (!withdraw) withdraw = ss.insertSheet(TAIKAI_HELPER_SHEET_);
  withdraw.clear();
  withdraw.getRange(1, 1).setFormula(
    '=IMPORTRANGE("' + UKETSUKE_SOURCE_ID_ + '","入会・退会_データ!G1:L")'
  );
  try { withdraw.hideSheet(); } catch (e1) {}
}

function parseUketsukeOpDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v;
  if (typeof v === 'number' && v > 30000) {
    var base = new Date(1899, 11, 30);
    var d0 = new Date(base.getTime() + Math.floor(v) * 86400000);
    return isNaN(d0.getTime()) ? null : d0;
  }
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return null;
}

function uketsukeOpDedupeKey_(row) {
  var d = parseUketsukeOpDate_(row && row[0]);
  var ym = d ? (d.getFullYear() + '-' + String(d.getMonth() + 1)) : '';
  var name = String((row && row[1]) || '').replace(/\s+/g, '');
  var opt = String((row && row[4]) || '').trim();
  var bucket = String((row && row[2]) || '').indexOf('利用停止') !== -1 ? 'stop' : 'start';
  var fallback = name ? '' : String((row && row[5]) || '');
  return ym + '|' + name + '|' + fallback + '|' + bucket + '|' + opt;
}

function compactUketsukeOpDuplicates_() {
  try {
    var ss = SpreadsheetApp.openById(UKETSUKE_SOURCE_ID_);
    var sh = ss.getSheetByName('OP集計');
    if (!sh) return { ok: false, message: 'OP集計 not found' };
    var monthText = String(sh.getRange('B1').getDisplayValue() || sh.getRange('B1').getValue() || '').trim();
    var mm = monthText.match(/^(\d{4})年(\d{1,2})月$/);
    if (!mm) return { ok: false, message: 'OP集計!B1 is not a year-month: ' + monthText };
    var year = parseInt(mm[1], 10);
    var month = parseInt(mm[2], 10) - 1;
    var last = Math.max(sh.getLastRow(), 1);
    var log = sh.getRange(1, 9, last, 6).getValues();
    var header = log[0];
    var other = [];
    var thisMonth = [];
    var i;
    for (i = 1; i < log.length; i++) {
      var row = log[i];
      if (!row[0] && !row[1] && !row[5]) continue;
      var d = parseUketsukeOpDate_(row[0]);
      if (d && d.getFullYear() === year && d.getMonth() === month) thisMonth.push(row);
      else other.push(row);
    }
    var seen = {};
    var unique = [];
    for (i = 0; i < thisMonth.length; i++) {
      var key = uketsukeOpDedupeKey_(thisMonth[i]);
      if (!key || seen[key]) continue;
      seen[key] = true;
      unique.push(thisMonth[i]);
    }
    var out = [header].concat(other).concat(unique);
    out.sort(function (a, b) {
      if (String(a[0]).indexOf('受信') !== -1) return -1;
      if (String(b[0]).indexOf('受信') !== -1) return 1;
      var da = parseUketsukeOpDate_(a[0]);
      var db = parseUketsukeOpDate_(b[0]);
      var ta = da ? da.getTime() : 0;
      var tb = db ? db.getTime() : 0;
      return tb - ta;
    });
    var prevLast = last;
    sh.getRange(1, 9, prevLast, 6).clearContent();
    sh.getRange(1, 9, out.length, 6).setValues(out);
    if (out.length > 1) {
      sh.getRange(2, 9, out.length - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
    }

    var idxByOpt = {};
    for (i = 0; i < OP_NAMES_.length; i++) idxByOpt[OP_NAMES_[i]] = i;
    var counts = OP_NAMES_.map(function () { return { newSignup: 0, opAdd: 0, stop: 0 }; });
    var counted = {};
    for (i = 0; i < unique.length; i++) {
      var k2 = uketsukeOpDedupeKey_(unique[i]);
      if (counted[k2]) continue;
      counted[k2] = true;
      var cat = String(unique[i][2] || '');
      var opt = String(unique[i][4] || '').trim();
      var idx = idxByOpt[opt];
      if (idx === undefined) continue;
      if (cat.indexOf('利用停止') !== -1) counts[idx].stop++;
      else if (cat.indexOf('新規入会') !== -1) counts[idx].newSignup++;
      else if (cat.indexOf('利用開始') !== -1 || cat.indexOf('OP追加') !== -1) counts[idx].opAdd++;
    }
    var bCol = [];
    var cCol = [];
    var dCol = [];
    var eCol = [];
    var summary = [];
    for (i = 0; i < OP_NAMES_.length; i++) {
      var start = counts[i].newSignup + counts[i].opAdd;
      bCol.push([counts[i].newSignup]);
      cCol.push([counts[i].opAdd]);
      dCol.push([start]);
      eCol.push([counts[i].stop]);
      summary.push({
        name: OP_NAMES_[i],
        newSignup: counts[i].newSignup,
        opAdd: counts[i].opAdd,
        start: start,
        stop: counts[i].stop
      });
    }
    sh.getRange(3, 2, OP_NAMES_.length, 1).setValues(bCol);
    sh.getRange(3, 3, OP_NAMES_.length, 1).setValues(cCol);
    sh.getRange(3, 4, OP_NAMES_.length, 1).setValues(dCol);
    sh.getRange(3, 5, OP_NAMES_.length, 1).setValues(eCol);

    var nippo = ss.getSheetByName('日報');
    if (nippo) {
      var labels = nippo.getRange(21, 2, OP_NAMES_.length, 1).getDisplayValues();
      var currentCol = [];
      var stopCol = [];
      for (i = 0; i < OP_NAMES_.length; i++) {
        var lab = String(labels[i][0] || '').trim();
        var cidx = idxByOpt[lab];
        if (cidx === undefined) cidx = i;
        currentCol.push([counts[cidx].newSignup + counts[cidx].opAdd]);
        stopCol.push([counts[cidx].stop]);
      }
      nippo.getRange(21, 4, OP_NAMES_.length, 1).setValues(currentCol);
      nippo.getRange(21, 5, OP_NAMES_.length, 1).setValues(stopCol);
    }

    var membershipView = refreshUketsukeMembershipDisplay_(ss, monthText);
    return {
      ok: true,
      month: monthText,
      before: thisMonth.length,
      after: unique.length,
      enrollShown: membershipView.enroll,
      withdrawShown: membershipView.withdraw,
      options: summary
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function ymLabelOf_(v) {
  var s = String(v || '').trim();
  var m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
  if (m) return m[1] + '年' + parseInt(m[2], 10) + '月';
  return s;
}

function refreshUketsukeMembershipDisplay_(ss, monthText) {
  var mem = ss.getSheetByName('入会・退会');
  var data = ss.getSheetByName('入会・退会_データ');
  if (!mem || !data) return { enroll: 0, withdraw: 0 };
  try { mem.showSheet(); } catch (e1) {}
  mem.getRange('B1').setValue(monthText);
  mem.getRange('G1').setValue(monthText);
  var last = Math.max(data.getLastRow(), 1);
  var enrollRaw = data.getRange(2, 1, last, 5).getValues();
  var withdrawRaw = data.getRange(2, 7, last, 6).getValues();
  var enroll = [];
  var withdraw = [];
  var i;
  for (i = 0; i < enrollRaw.length; i++) {
    if (!enrollRaw[i][1]) continue;
    if (ymLabelOf_(enrollRaw[i][2]) === monthText) enroll.push(enrollRaw[i]);
  }
  for (i = 0; i < withdrawRaw.length; i++) {
    if (!withdrawRaw[i][1]) continue;
    if (ymLabelOf_(withdrawRaw[i][2]) === monthText) withdraw.push(withdrawRaw[i]);
  }
  enroll.sort(function (a, b) {
    var ta = a[0] instanceof Date ? a[0].getTime() : 0;
    var tb = b[0] instanceof Date ? b[0].getTime() : 0;
    return tb - ta;
  });
  withdraw.sort(function (a, b) {
    var ta = a[0] instanceof Date ? a[0].getTime() : 0;
    var tb = b[0] instanceof Date ? b[0].getTime() : 0;
    return tb - ta;
  });
  var n = Math.max(enroll.length, withdraw.length, 1);
  var prev = Math.max(mem.getLastRow() - 2, n);
  mem.getRange(3, 1, prev, 11).clearContent();
  var rows = [];
  for (i = 0; i < n; i++) {
    var row = ['', '', '', '', '', '', '', '', '', '', ''];
    if (i < enroll.length) {
      row[0] = enroll[i][0];
      row[1] = enroll[i][1];
      row[2] = ymLabelOf_(enroll[i][2]);
      row[3] = enroll[i][3];
    }
    if (i < withdraw.length) {
      row[5] = withdraw[i][0];
      row[6] = withdraw[i][1];
      row[7] = ymLabelOf_(withdraw[i][2]);
      row[8] = withdraw[i][3];
      row[9] = withdraw[i][5];
      row[10] = withdraw[i][4];
    }
    rows.push(row);
  }
  if (rows.length) mem.getRange(3, 1, rows.length, 11).setValues(rows);
  return { enroll: enroll.length, withdraw: withdraw.length };
}

function extractNyukaiNameHub_(body) {
  var text = String(body || '');
  var m1 = text.match(/お名前\s*[:：]\s*(.+?)\s*様/);
  if (m1) return String(m1[1]).replace(/^[>\s]+/, '').trim();
  var lines = text.split(/\r?\n/);
  var i;
  for (i = 0; i < Math.min(lines.length, 20); i++) {
    var line = String(lines[i] || '').trim();
    if (!line || /JOYFIT|受付番号|この度は|お支払い|ご利用開始|ご入会内容|会員情報/.test(line)) continue;
    var m2 = line.match(/^(.{1,30}?)\s*様\s*$/);
    if (m2) return String(m2[1]).replace(/^[>\s]+/, '').trim();
  }
  return '';
}

function calcNyukaiYmHub_(body, emailDate) {
  var startMatch = String(body || '').match(/ご利用開始日[の]?\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (startMatch) {
    return parseInt(startMatch[1], 10) + '年' + parseInt(startMatch[2], 10) + '月';
  }
  return emailDate.getFullYear() + '年' + (emailDate.getMonth() + 1) + '月';
}

function detectNyukaiCatHub_(subject, body) {
  var s = String(subject || '');
  var b = String(body || '');
  if (/法人個人月払|法人会員|法人.*月払/.test(b)) return '法人会員';
  if (s.indexOf('【JOYFIT24経堂】') !== -1) return '6ヶ月割';
  if (/ナショナル会員/.test(b)) return '6ヶ月割';
  if (/ご入会ありがとうございます/.test(s) && s.indexOf('【JOYFIT24経堂】') === -1) return '法人会員';
  return '6ヶ月割';
}

function auditNyukaiGmail_() {
  try {
    var q = 'from:info@joyfit-service.jp subject:(ご入会ありがとうございます) after:2026/08/30 before:2026/10/01';
    var threads = GmailApp.search(q, 0, 200);
    var people = [];
    var seen = {};
    var t;
    for (t = 0; t < threads.length; t++) {
      var messages = threads[t].getMessages();
      var m;
      for (m = messages.length - 1; m >= 0; m--) {
        var msg = messages[m];
        var subject = String(msg.getSubject() || '');
        if (subject.indexOf('ご入会') === -1) continue;
        var body = String(msg.getPlainBody() || '');
        if (body.length < 80) body = String(msg.getBody() || '').replace(/<[^>]+>/g, ' ');
        var name = extractNyukaiNameHub_(body);
        var date = msg.getDate();
        var ym = calcNyukaiYmHub_(body, date);
        var cat = detectNyukaiCatHub_(subject, body);
        var key = String(name).replace(/\s+/g, '') + '|' + ym;
        if (!name || seen[key]) continue;
        seen[key] = true;
        people.push({
          name: name,
          ym: ym,
          cat: cat,
          date: Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
          msgId: msg.getId(),
          subject: subject
        });
        break;
      }
    }
    var sept = people.filter(function (p) { return p.ym === '2026年9月'; });
    var other = people.filter(function (p) { return p.ym !== '2026年9月'; });
    return {
      ok: true,
      threads: threads.length,
      parsed: people.length,
      sept: sept.length,
      septPeople: sept,
      otherYm: other
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function backfillEnrollFromOpLog_() {
  try {
    var ss = SpreadsheetApp.openById(UKETSUKE_SOURCE_ID_);
    var op = ss.getSheetByName('OP集計');
    var data = ss.getSheetByName('入会・退会_データ');
    if (!op || !data) return { ok: false, message: 'sheet missing' };
    var monthText = String(op.getRange('B1').getDisplayValue() || '2026年9月').trim();
    var lastOp = Math.max(op.getLastRow(), 1);
    var log = op.getRange(1, 9, lastOp, 6).getValues();
    var lastData = Math.max(data.getLastRow(), 1);
    var enroll = data.getRange(2, 1, lastData, 5).getValues();
    var have = {};
    var i;
    for (i = 0; i < enroll.length; i++) {
      if (!enroll[i][1]) continue;
      if (ymLabelOf_(enroll[i][2]) !== monthText) continue;
      have[String(enroll[i][1]).replace(/\s+/g, '')] = true;
    }
    var added = [];
    var seenNew = {};
    for (i = 1; i < log.length; i++) {
      var row = log[i];
      var d = parseUketsukeOpDate_(row[0]);
      var cat = String(row[2] || '');
      if (cat.indexOf('新規入会') === -1) continue;
      if (!d) continue;
      var ym = d.getFullYear() + '年' + (d.getMonth() + 1) + '月';
      if (ym !== monthText) continue;
      var name = String(row[1] || '').trim();
      var key = name.replace(/\s+/g, '');
      if (!key || have[key] || seenNew[key]) continue;
      seenNew[key] = true;
      added.push([row[0], name, monthText, '6ヶ月割', String(row[5] || '')]);
    }

    var gmail = { ok: false, sept: 0 };
    try { gmail = auditNyukaiGmail_(); } catch (e2) { gmail = { ok: false }; }
    if (gmail && gmail.ok && gmail.septPeople) {
      for (i = 0; i < gmail.septPeople.length; i++) {
        var p = gmail.septPeople[i];
        var nk = String(p.name || '').replace(/\s+/g, '');
        if (!nk || have[nk] || seenNew[nk]) continue;
        seenNew[nk] = true;
        added.push([p.date, p.name, monthText, p.cat || '6ヶ月割', p.msgId || '']);
      }
    }

    if (added.length) {
      var start = lastData + 1;
      data.getRange(start, 3, added.length, 1).setNumberFormat('@');
      data.getRange(start, 1, added.length, 5).setValues(added);
      data.getRange(start, 1, added.length, 1).setNumberFormat('yyyy/mm/dd hh:mm');
    }
    var view = refreshUketsukeMembershipDisplay_(ss, monthText);
    var nippo = ss.getSheetByName('日報');
    if (nippo) {
      nippo.getRange('D10').setFormula('=IFERROR(N(INDEX(INDIRECT("\'"&$B$1&"\'!$I$5:$I$35"),DAY(TODAY())))+N(INDEX(INDIRECT("\'"&$B$1&"\'!$H$5:$H$35"),DAY(TODAY()))),0)');
      nippo.getRange('F10').setFormula('=IFERROR(N(INDEX(INDIRECT("\'"&$B$1&"\'!$J$5:$J$35"),DAY(TODAY()))),0)');
      nippo.getRange('H10').setFormula('=IFERROR(N(INDEX(INDIRECT("\'"&$B$1&"\'!$K$5:$K$35"),DAY(TODAY()))),0)');
      nippo.getRange('D14').setFormula('=IFERROR(N(INDIRECT($B$1&"!H36"))+N(INDIRECT($B$1&"!I36")),0)');
      nippo.getRange('F14').setFormula('=IFERROR(N(INDIRECT($B$1&"!J36")),0)');
      nippo.getRange('H14').setFormula('=IFERROR(N(INDIRECT($B$1&"!K36")),0)');
    }
    var mm = monthText.match(/^(\d{4})年(\d{1,2})月$/);
    var synced = null;
    if (mm) {
      synced = syncMonthlyEnrollCountsSafe_(
        ss, parseInt(mm[1], 10), parseInt(mm[2], 10) - 1, monthText
      );
    }
    return {
      ok: true,
      month: monthText,
      added: added.length,
      addedNames: added.map(function (r) { return r[1]; }),
      enrollShown: view.enroll,
      withdrawShown: view.withdraw,
      daily: synced,
      gmailSept: gmail && gmail.sept
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function syncMonthlyEnrollCountsSafe_(ss, year, month, monthText) {
  var sheetName = String(year % 100).padStart(2, '0') + String(month + 1).padStart(2, '0');
  var sheet = ss.getSheetByName(sheetName);
  var data = ss.getSheetByName('入会・退会_データ');
  if (!sheet || !data) return { updated: false };
  var last = Math.max(data.getLastRow(), 1);
  var rows = data.getRange(2, 1, last, 5).getValues();
  var gen = {};
  var corp = {};
  var d;
  for (d = 1; d <= 31; d++) { gen[d] = 0; corp[d] = 0; }
  var seen = {};
  var i;
  for (i = 0; i < rows.length; i++) {
    if (ymLabelOf_(rows[i][2]) !== monthText) continue;
    var name = String(rows[i][1] || '').replace(/\s+/g, '');
    if (!name) continue;
    var uniq = name + '|' + String(rows[i][3] || '');
    if (seen[uniq]) continue;
    seen[uniq] = true;
    var ts = parseUketsukeOpDate_(rows[i][0]);
    var day = 1;
    if (ts && ts.getFullYear() === year && ts.getMonth() === month) day = ts.getDate();
    var cat = String(rows[i][3] || '');
    if (cat.indexOf('法人') !== -1) corp[day]++;
    else gen[day]++;
  }
  var gVals = [];
  var cVals = [];
  for (d = 1; d <= 31; d++) {
    gVals.push([gen[d] > 0 ? gen[d] : '']);
    cVals.push([corp[d] > 0 ? corp[d] : '']);
  }
  sheet.getRange(5, 5, 31, 1).setValues(gVals);
  sheet.getRange(5, 6, 31, 1).setValues(cVals);
  var genTotal = 0;
  var corpTotal = 0;
  for (d = 1; d <= 31; d++) { genTotal += gen[d]; corpTotal += corp[d]; }
  return {
    updated: true,
    sheet: sheetName,
    general: genTotal,
    corporate: corpTotal
  };
}

function nameKeyHub_(s) {
  return String(s || '').replace(/\s+/g, '');
}

function normalizeOptionNameHub_(rawName) {
  var s = String(rawName || '');
  if (s.indexOf('水素水') !== -1 && s.indexOf('プロテイン') !== -1) return 'プロテイン＋水素水';
  if (s.indexOf('VIP') !== -1 && (s.indexOf('あんしん') !== -1 || s.indexOf('安心') !== -1)) return '安心サポートVIP';
  if (s.indexOf('ボディプランナー') !== -1 || s.indexOf('ボディープランナー') !== -1 || s.indexOf('体組成') !== -1) {
    return '体組成計';
  }
  if (s.indexOf('マットレンタル') !== -1 || s.indexOf('レンタルマット') !== -1) return 'レンタルマット';
  if (s.indexOf('ピラティス') !== -1) return 'ピラティスリフォーマー';
  if (s.indexOf('ロッカー') !== -1 && /(1[,，]?500|１[,，]?５００|1500)/.test(s)) return '契約ロッカー1,500';
  if (s.indexOf('ヨガ') !== -1 && s.indexOf('ロッカー') !== -1) return 'ヨガロッカー';
  if (s.indexOf('レンタルタオル') !== -1 || (s.indexOf('タオル') !== -1 && s.indexOf('レンタル') !== -1)) {
    return 'レンタルタオル';
  }
  if (s.indexOf('ホットスタジオ') !== -1 || s.indexOf('HOTスタジオ') !== -1) return 'ホットスタジオ';
  if (s.indexOf('オンラインレッスン') !== -1 || s.indexOf('オンライン・レッスン') !== -1) return 'オンラインレッスン';
  if (s.indexOf('セルフエステ') !== -1) return 'セルフエステ';
  if (s.indexOf('タンニング') !== -1) return 'タンニング';
  if (s.indexOf('プロテイン') !== -1 && (s.indexOf('無制限') !== -1 || s.indexOf('制限なし') !== -1)) {
    return 'プロテイン無制限';
  }
  if (s.indexOf('プロテイン') !== -1 && (s.indexOf('12') !== -1 || s.indexOf('１２') !== -1)) return 'プロテイン12杯';
  if (s.indexOf('安心サポート') !== -1 || s.indexOf('あんしんサポート') !== -1) return '安心サポート';
  if (s.indexOf('水素水') !== -1) return '水素水';
  var byLen = OP_NAMES_.slice().sort(function (a, b) { return b.length - a.length; });
  var i;
  for (i = 0; i < byLen.length; i++) {
    if (s.indexOf(byLen[i]) !== -1) return byLen[i];
  }
  return '';
}

function extractFeeSectionHub_(body) {
  var text = String(body || '');
  if (!text) return '';
  var start = text.search(/月会費.*内訳|お支払い内容|ご契約中のオプション|お支払い内容は以下/);
  if (start >= 0) {
    var rest = text.substring(start);
    var endRel = rest.search(/APP登録方法|お支払方法|注意事項|ご案内|ご利用規約|クレジットカード|JOYFIT App/);
    return endRel > 0 ? rest.substring(0, endRel) : rest;
  }
  var m = text.search(/[（(]\s*\d{1,2}\s*月分\s*[）)]/);
  if (m < 0) return text;
  return text.substring(Math.max(0, m - 80), Math.min(text.length, m + 1400));
}

function parseOpNamesFromBodyHub_(body) {
  var text = extractFeeSectionHub_(body);
  var found = {};
  var list = [];
  var re = /([^\n\r]{2,40}?)[（(]\s*(\d{1,2})\s*月分\s*[）)][^\n\r]{0,80}?(\d[\d,]*)\s*円/g;
  var m;
  while ((m = re.exec(text)) !== null) {
    var raw = String(m[1] || '').replace(/^[\s\-・･>]+/, '').trim();
    if (!raw || /法人個人月払|法人会員|ナショナル会員|入会金|事務手数料|月会費|小計|合計/.test(raw)) continue;
    var norm = normalizeOptionNameHub_(raw);
    if (!norm || found[norm]) continue;
    found[norm] = true;
    list.push(norm);
  }
  if (list.length < 2) {
    var i;
    for (i = 0; i < OP_NAMES_.length; i++) {
      var name = OP_NAMES_[i];
      if (found[name]) continue;
      if (text.indexOf(name) === -1) continue;
      if (name === '安心サポート' && found['安心サポートVIP']) continue;
      if (name === '水素水' && found['プロテイン＋水素水']) continue;
      found[name] = true;
      list.push(name);
    }
  }
  return list;
}

function searchGmailPagedHub_(query, maxThreads) {
  var out = [];
  var start = 0;
  var cap = Math.max(1, maxThreads || 400);
  while (out.length < cap) {
    var batch = GmailApp.search(query, start, 100);
    if (!batch || !batch.length) break;
    var i;
    for (i = 0; i < batch.length && out.length < cap; i++) out.push(batch[i]);
    if (batch.length < 100) break;
    start += batch.length;
  }
  return out;
}

function gmailAccountHub_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (e0) {
    try { return Session.getEffectiveUser().getEmail() || ''; } catch (e1) { return ''; }
  }
}

function auditOpMailCoverage_() {
  try {
    var ss = SpreadsheetApp.openById(UKETSUKE_SOURCE_ID_);
    var op = ss.getSheetByName('OP集計');
    var data = ss.getSheetByName('入会・退会_データ');
    if (!op || !data) return { ok: false, message: 'sheet missing' };
    var monthText = String(op.getRange('B1').getDisplayValue() || '2026年9月').trim();
    var mm = monthText.match(/^(\d{4})年(\d{1,2})月$/);
    var year = mm ? parseInt(mm[1], 10) : 2026;
    var month = mm ? parseInt(mm[2], 10) - 1 : 8;

    var lastData = Math.max(data.getLastRow(), 1);
    var enrollRows = data.getRange(2, 1, lastData, 5).getValues();
    var enrollByKey = {};
    var enrollList = [];
    var i;
    for (i = 0; i < enrollRows.length; i++) {
      var name = String(enrollRows[i][1] || '').trim();
      if (!name) continue;
      if (ymLabelOf_(enrollRows[i][2]) !== monthText) continue;
      var key = nameKeyHub_(name);
      if (enrollByKey[key]) continue;
      var cat = String(enrollRows[i][3] || '');
      var rec = {
        name: name,
        cat: cat,
        six: cat.indexOf('法人') === -1,
        msgId: String(enrollRows[i][4] || ''),
        startOps: [],
        addOps: [],
        stopOps: []
      };
      enrollByKey[key] = rec;
      enrollList.push(rec);
    }

    var lastOp = Math.max(op.getLastRow(), 1);
    var log = op.getRange(1, 9, lastOp, 6).getValues();
    var startCounts = {};
    var addCounts = {};
    var stopCounts = {};
    var uniqueStart = {};
    var uniqueAdd = {};
    var uniqueStop = {};
    var startPeople = {};
    var addPeople = {};
    var orphanStarts = [];
    var j;
    for (j = 0; j < OP_NAMES_.length; j++) {
      startCounts[OP_NAMES_[j]] = 0;
      addCounts[OP_NAMES_[j]] = 0;
      stopCounts[OP_NAMES_[j]] = 0;
    }
    for (i = 1; i < log.length; i++) {
      var row = log[i];
      var d = parseUketsukeOpDate_(row[0]);
      if (!d || d.getFullYear() !== year || d.getMonth() !== month) continue;
      var opt = String(row[4] || '').trim();
      var cat2 = String(row[2] || '');
      var person = String(row[1] || '').trim();
      var pk = nameKeyHub_(person);
      var ukey = uketsukeOpDedupeKey_(row);
      if (cat2.indexOf('利用停止') !== -1) {
        if (uniqueStop[ukey]) continue;
        uniqueStop[ukey] = true;
        if (stopCounts[opt] !== undefined) stopCounts[opt]++;
        if (enrollByKey[pk]) enrollByKey[pk].stopOps.push(opt);
      } else if (cat2.indexOf('新規入会') !== -1) {
        if (uniqueStart[ukey]) continue;
        uniqueStart[ukey] = true;
        if (startCounts[opt] !== undefined) startCounts[opt]++;
        if (!startPeople[pk]) startPeople[pk] = [];
        startPeople[pk].push(opt);
        if (enrollByKey[pk]) enrollByKey[pk].startOps.push(opt);
        else orphanStarts.push({ name: person, opt: opt });
      } else if (cat2.indexOf('利用開始') !== -1 || cat2.indexOf('OP追加') !== -1) {
        if (uniqueAdd[ukey]) continue;
        uniqueAdd[ukey] = true;
        if (addCounts[opt] !== undefined) addCounts[opt]++;
        if (!addPeople[pk]) addPeople[pk] = [];
        addPeople[pk].push(opt);
        if (enrollByKey[pk]) enrollByKey[pk].addOps.push(opt);
      }
    }

    var packCore = [
      '安心サポートVIP', 'オンラインレッスン', '体組成計', 'レンタルマット',
      'プロテイン＋水素水', 'レンタルタオル', 'ホットスタジオ'
    ];
    var six = enrollList.filter(function (p) { return p.six; });
    var corp = enrollList.filter(function (p) { return !p.six; });
    var packOk = [];
    var packShort = [];
    var packZero = [];
    for (i = 0; i < six.length; i++) {
      var ops = six[i].startOps;
      var uniqOps = {};
      for (j = 0; j < ops.length; j++) uniqOps[ops[j]] = true;
      var n = Object.keys(uniqOps).length;
      var missing = [];
      for (j = 0; j < packCore.length; j++) {
        if (!uniqOps[packCore[j]]) missing.push(packCore[j]);
      }
      if (!uniqOps['タンニング'] && !uniqOps['セルフエステ']) missing.push('タンニングまたはセルフエステ');
      var item = { name: six[i].name, opCount: n, ops: Object.keys(uniqOps), missing: missing };
      if (n === 0) packZero.push(item);
      else if (missing.length) packShort.push(item);
      else packOk.push(item);
    }

    var corpWithOp = corp.filter(function (p) { return p.startOps.length; }).map(function (p) {
      return { name: p.name, opCount: p.startOps.length, ops: p.startOps };
    });
    var signupNotSix = [];
    for (i = 0; i < enrollList.length; i++) {
      if (enrollList[i].six) continue;
      if (enrollList[i].startOps.length) signupNotSix.push({
        name: enrollList[i].name,
        cat: enrollList[i].cat,
        ops: enrollList[i].startOps
      });
    }
    var startOnly = [];
    var pk;
    for (pk in startPeople) {
      if (!startPeople.hasOwnProperty(pk)) continue;
      if (!enrollByKey[pk]) startOnly.push({ name: pk, ops: startPeople[pk] });
      else if (!enrollByKey[pk].six) startOnly.push({
        name: enrollByKey[pk].name,
        cat: enrollByKey[pk].cat,
        ops: startPeople[pk]
      });
    }
    var gmail = {
      account: gmailAccountHub_(),
      enrollThreads: 0,
      enrollParsed: 0,
      enrollSept: 0,
      opChangeThreads: 0,
      enrollPeople: [],
      missingVsSheet: [],
      extraVsSheet: []
    };
    try {
      var enrollQ = 'from:info@joyfit-service.jp subject:ご入会ありがとうございます after:2026/08/30 before:2026/10/01';
      var changeQ = 'from:info@joyfit-service.jp subject:オプションご契約につきまして after:2026/08/30 before:2026/10/01';
      var enrollThreads = searchGmailPagedHub_(enrollQ, 400);
      var changeThreads = searchGmailPagedHub_(changeQ, 400);
      gmail.enrollThreads = enrollThreads.length;
      gmail.opChangeThreads = changeThreads.length;
      var seenG = {};
      for (i = 0; i < enrollThreads.length; i++) {
        var messages = enrollThreads[i].getMessages();
        var midx;
        for (midx = messages.length - 1; midx >= 0; midx--) {
          var msg = messages[midx];
          var subject = String(msg.getSubject() || '');
          if (subject.indexOf('ご入会') === -1) continue;
          var body = String(msg.getPlainBody() || '');
          if (body.length < 80) body = String(msg.getBody() || '').replace(/<[^>]+>/g, ' ');
          var personName = extractNyukaiNameHub_(body);
          if (!personName) continue;
          var date = msg.getDate();
          var ym = calcNyukaiYmHub_(body, date);
          gmail.enrollParsed++;
          if (ym !== monthText) continue;
          var gk = nameKeyHub_(personName);
          if (seenG[gk]) continue;
          seenG[gk] = true;
          var opsG = parseOpNamesFromBodyHub_(body);
          gmail.enrollSept++;
          gmail.enrollPeople.push({
            name: personName,
            cat: detectNyukaiCatHub_(subject, body),
            opCount: opsG.length,
            ops: opsG
          });
          break;
        }
      }
      for (i = 0; i < gmail.enrollPeople.length; i++) {
        if (!enrollByKey[nameKeyHub_(gmail.enrollPeople[i].name)]) {
          gmail.extraVsSheet.push(gmail.enrollPeople[i].name);
        }
      }
      for (i = 0; i < six.length; i++) {
        if (!seenG[nameKeyHub_(six[i].name)]) gmail.missingVsSheet.push(six[i].name);
      }
    } catch (gerr) {
      gmail.error = String(gerr && gerr.message ? gerr.message : gerr);
    }

    var startTotals = {};
    for (j = 0; j < OP_NAMES_.length; j++) {
      startTotals[OP_NAMES_[j]] = {
        newSignup: startCounts[OP_NAMES_[j]],
        opAdd: addCounts[OP_NAMES_[j]],
        start: startCounts[OP_NAMES_[j]] + addCounts[OP_NAMES_[j]],
        stop: stopCounts[OP_NAMES_[j]]
      };
    }

    return {
      ok: true,
      month: monthText,
      gmailAccount: gmail.account,
      enroll: {
        total: enrollList.length,
        sixMonth: six.length,
        corporate: corp.length,
        corporateWithSignupOp: corpWithOp,
        signupOpNotSixMonth: startOnly
      },
      pack: {
        expectedCore: packCore.length + 1,
        complete: packOk.length,
        short: packShort.length,
        zero: packZero.length,
        shortPeople: packShort,
        zeroPeople: packZero
      },
      opLogPeople: {
        newSignupPeople: Object.keys(startPeople).length,
        opAddPeople: Object.keys(addPeople).length,
        opAddNames: Object.keys(addPeople).map(function (k) {
          return { name: (enrollByKey[k] && enrollByKey[k].name) || k, ops: addPeople[k] };
        }),
        orphanNewSignupRows: orphanStarts.length
      },
      optionStarts: startTotals,
      gmailKusaka: {
        account: gmail.account,
        enrollThreads: gmail.enrollThreads,
        enrollParsed: gmail.enrollParsed,
        enrollSeptUnique: gmail.enrollSept,
        opChangeThreads: gmail.opChangeThreads,
        sixMonthMissingInThisGmail: gmail.missingVsSheet,
        extraInThisGmail: gmail.extraVsSheet,
        error: gmail.error || ''
      },
      note: '店舗Gmailの入会・退会_データ／OPログを正とし、この実行アカウントのGmailは照合用。6ヶ月割はテンプレ8種（コア7+タンニングかセルフエステ）。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function inspectQueryThreadsHub_(query, openMessages, maxThreads) {
  var threads = searchGmailPagedHub_(query, maxThreads || 800);
  var uniqueNames = {};
  var enrollMessages = 0;
  var optionMessages = 0;
  var allMessages = 0;
  var bundled = [];
  var i;
  if (openMessages) {
    for (i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages() || [];
      allMessages += messages.length;
      var names = [];
      var j;
      for (j = 0; j < messages.length; j++) {
        var subject = String(messages[j].getSubject() || '');
        var body = '';
        var name = '';
        if (subject.indexOf('ご入会') !== -1) {
          enrollMessages++;
          body = String(messages[j].getPlainBody() || '');
          if (body.length < 80) body = String(messages[j].getBody() || '').replace(/<[^>]+>/g, ' ');
          name = extractNyukaiNameHub_(body);
        } else if (subject.indexOf('オプションご契約') !== -1) {
          optionMessages++;
          if (!name) {
            body = String(messages[j].getPlainBody() || '');
            var m = body.match(/(.{1,20}?)様/);
            if (m) name = String(m[1]).replace(/^[>\s]+/, '').trim();
          }
        }
        if (name) {
          uniqueNames[nameKeyHub_(name)] = name;
          names.push(name);
        }
      }
      if (messages.length > 1) {
        bundled.push({
          messages: messages.length,
          names: names
        });
      }
    }
  }
  return {
    query: query,
    threads: threads.length,
    messagesOpened: openMessages ? allMessages : null,
    enrollMessages: openMessages ? enrollMessages : null,
    optionMessages: openMessages ? optionMessages : null,
    uniquePeople: openMessages ? Object.keys(uniqueNames).length : null,
    threadsWithMultipleMessages: openMessages ? bundled.length : null,
    bundledSample: openMessages ? bundled.slice(0, 12) : []
  };
}

function auditGmailThreadShape_() {
  try {
    var septEnroll = inspectQueryThreadsHub_(
      'from:info@joyfit-service.jp subject:ご入会ありがとうございます after:2026/08/30 before:2026/10/01',
      true,
      400
    );
    var septOption = inspectQueryThreadsHub_(
      'from:info@joyfit-service.jp subject:オプションご契約につきまして after:2026/08/30 before:2026/10/01',
      true,
      400
    );
    var allEnroll = inspectQueryThreadsHub_(
      'from:info@joyfit-service.jp subject:ご入会ありがとうございます',
      false,
      800
    );
    var allOption = inspectQueryThreadsHub_(
      'from:info@joyfit-service.jp subject:オプションご契約につきまして',
      false,
      800
    );
    return {
      ok: true,
      account: gmailAccountHub_(),
      meaning: {
        thread: 'Gmailの1行＝会話の束。左の info 7 は7通入っているという意味',
        previous25: '前回の25は after/before 付き検索のスレッド数で、1通ずつではない。さらに各スレッドの先頭1通しか見ていなかった',
        screenshotRows: '画面右上の 572行 / 784行 は期間なし検索の会話行数'
      },
      septemberEnroll: septEnroll,
      septemberOption: septOption,
      allTimeEnrollThreads: allEnroll.threads,
      allTimeOptionThreads: allOption.threads
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function parseOptionContractHub_(date, body, msgId) {
  var results = [];
  var nameMatch = String(body || '').match(/([^\r\n]{1,40}?)\s*様/);
  var name = nameMatch ? String(nameMatch[1]).replace(/^[>\s]+/, '').trim() : '';
  var re = /[（(【\[](利用開始|利用停止)[）)】\]]\s*([^\r\n]+)/g;
  var m;
  while ((m = re.exec(String(body || ''))) !== null) {
    var status = m[1] === '利用開始' ? '利用開始(OP追加)' : '利用停止';
    var raw = String(m[2]).trim();
    if (!raw || /について|場合でも|^合計|^小計|^の場合|^円/.test(raw)) continue;
    var norm = normalizeOptionNameHub_(raw);
    if (!norm) continue;
    results.push([date, name, status, raw, norm, msgId]);
  }
  return results;
}

function countGmailLabelThreadsHub_(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) return { name: labelName, exists: false, threads: 0 };
  var n = 0;
  var start = 0;
  while (n < 3000) {
    var batch = label.getThreads(start, 100);
    if (!batch || !batch.length) break;
    n += batch.length;
    if (batch.length < 100) break;
    start += batch.length;
  }
  return { name: labelName, exists: true, threads: n };
}

function rebuildFromUnbundledMail_(dryRun) {
  try {
    var ss = SpreadsheetApp.openById(UKETSUKE_SOURCE_ID_);
    var data = ss.getSheetByName('入会・退会_データ');
    var op = ss.getSheetByName('OP集計');
    if (!data || !op) return { ok: false, message: 'sheet missing' };
    var monthText = String(op.getRange('B1').getDisplayValue() || '2026年9月').trim();
    var packCore = [
      '安心サポートVIP', 'オンラインレッスン', '体組成計', 'レンタルマット',
      'プロテイン＋水素水', 'レンタルタオル', 'ホットスタジオ'
    ];

    var lastData = Math.max(data.getLastRow(), 1);
    var enrollRows = data.getRange(2, 1, lastData, 5).getValues();
    var sheetByKey = {};
    var sheetSept = [];
    var i;
    for (i = 0; i < enrollRows.length; i++) {
      var sName = String(enrollRows[i][1] || '').trim();
      if (!sName) continue;
      var sYm = ymLabelOf_(enrollRows[i][2]);
      var sKey = nameKeyHub_(sName) + '|' + sYm;
      sheetByKey[sKey] = {
        name: sName,
        ym: sYm,
        cat: String(enrollRows[i][3] || '')
      };
      if (sYm === monthText) sheetSept.push(sheetByKey[sKey]);
    }

    var enrollQ = 'from:info@joyfit-service.jp subject:ご入会ありがとうございます after:2026/08/30 before:2026/10/01';
    var changeQ = 'from:info@joyfit-service.jp subject:オプションご契約につきまして after:2026/08/30 before:2026/10/01';
    var enrollThreads = searchGmailPagedHub_(enrollQ, 400);
    var changeThreads = searchGmailPagedHub_(changeQ, 400);

    var gmailPeople = [];
    var seenG = {};
    var enrollMessages = 0;
    var t;
    for (t = 0; t < enrollThreads.length; t++) {
      var messages = enrollThreads[t].getMessages() || [];
      var mi;
      for (mi = 0; mi < messages.length; mi++) {
        var msg = messages[mi];
        var subject = String(msg.getSubject() || '');
        if (subject.indexOf('ご入会') === -1) continue;
        enrollMessages++;
        var body = String(msg.getPlainBody() || '');
        if (body.length < 80) body = String(msg.getBody() || '').replace(/<[^>]+>/g, ' ');
        var personName = extractNyukaiNameHub_(body);
        if (!personName) continue;
        var date = msg.getDate();
        var ym = calcNyukaiYmHub_(body, date);
        var cat = detectNyukaiCatHub_(subject, body);
        var gKey = nameKeyHub_(personName) + '|' + ym;
        if (seenG[gKey]) continue;
        seenG[gKey] = true;
        var ops = parseOpNamesFromBodyHub_(body);
        gmailPeople.push({
          name: personName,
          ym: ym,
          cat: cat,
          six: cat.indexOf('法人') === -1,
          date: date,
          msgId: msg.getId(),
          ops: ops,
          opCount: ops.length,
          inSheet: !!sheetByKey[gKey]
        });
      }
    }

    var septPeople = gmailPeople.filter(function (p) { return p.ym === monthText; });
    var six = septPeople.filter(function (p) { return p.six; });
    var corp = septPeople.filter(function (p) { return !p.six; });
    var missing = septPeople.filter(function (p) { return !p.inSheet; });
    var packShort = [];
    var packOk = 0;
    for (i = 0; i < six.length; i++) {
      var uniq = {};
      for (t = 0; t < six[i].ops.length; t++) uniq[six[i].ops[t]] = true;
      var miss = [];
      for (t = 0; t < packCore.length; t++) {
        if (!uniq[packCore[t]]) miss.push(packCore[t]);
      }
      if (!uniq['タンニング'] && !uniq['セルフエステ']) miss.push('タンニングまたはセルフエステ');
      if (miss.length) packShort.push({ name: six[i].name, opCount: six[i].opCount, missing: miss, inSheet: six[i].inSheet });
      else packOk++;
    }

    var changeRows = [];
    var changeMessages = 0;
    for (t = 0; t < changeThreads.length; t++) {
      var cmsgs = changeThreads[t].getMessages() || [];
      for (i = 0; i < cmsgs.length; i++) {
        var csub = String(cmsgs[i].getSubject() || '');
        if (csub.indexOf('オプションご契約') === -1) continue;
        changeMessages++;
        var cbody = String(cmsgs[i].getPlainBody() || '');
        if (cbody.length < 80) cbody = String(cmsgs[i].getBody() || '').replace(/<[^>]+>/g, ' ');
        changeRows = changeRows.concat(parseOptionContractHub_(cmsgs[i].getDate(), cbody, cmsgs[i].getId()));
      }
    }

    var labels = [
      countGmailLabelThreadsHub_('入会メール/一般会員'),
      countGmailLabelThreadsHub_('入会メール/法人会員'),
      countGmailLabelThreadsHub_('退会メール/一般会員'),
      countGmailLabelThreadsHub_('退会メール/法人会員'),
      countGmailLabelThreadsHub_('オプションメール/追加・停止')
    ];

    var written = { enrollAdded: 0, opAdded: 0, compact: null };
    if (!dryRun) {
      var addEnroll = missing.map(function (p) {
        return [p.date, p.name, p.ym, p.cat, p.msgId];
      });
      if (addEnroll.length) {
        var start = lastData + 1;
        data.getRange(start, 3, addEnroll.length, 1).setNumberFormat('@');
        data.getRange(start, 1, addEnroll.length, 5).setValues(addEnroll);
        data.getRange(start, 1, addEnroll.length, 1).setNumberFormat('yyyy/mm/dd hh:mm');
        written.enrollAdded = addEnroll.length;
      }

      var lastOp = Math.max(op.getLastRow(), 1);
      var log = op.getRange(1, 9, lastOp, 6).getValues();
      var seenOp = {};
      for (i = 1; i < log.length; i++) {
        var k = uketsukeOpDedupeKey_(log[i]);
        if (k) seenOp[k] = true;
      }
      var newOp = [];
      for (i = 0; i < six.length; i++) {
        for (t = 0; t < six[i].ops.length; t++) {
          var row = [six[i].date, six[i].name, '利用開始(新規入会)', six[i].ops[t], six[i].ops[t], six[i].msgId];
          var dk = uketsukeOpDedupeKey_(row);
          if (seenOp[dk]) continue;
          seenOp[dk] = true;
          newOp.push(row);
        }
      }
      for (i = 0; i < changeRows.length; i++) {
        var ck = uketsukeOpDedupeKey_(changeRows[i]);
        if (!ck || seenOp[ck]) continue;
        seenOp[ck] = true;
        newOp.push(changeRows[i]);
      }
      if (newOp.length) {
        op.getRange(lastOp + 1, 9, newOp.length, 6).setValues(newOp);
        written.opAdded = newOp.length;
      }
      written.compact = compactUketsukeOpDuplicates_();
      var mm = monthText.match(/^(\d{4})年(\d{1,2})月$/);
      if (mm) {
        syncMonthlyEnrollCountsSafe_(ss, parseInt(mm[1], 10), parseInt(mm[2], 10) - 1, monthText);
      }
    }

    return {
      ok: true,
      dryRun: !!dryRun,
      meaning: {
        previous74: '入会・退会_データの2026年9月ユニーク人数（6ヶ月割69+法人5）。法人以外がキャンペーン一般会員',
        gmail82: '受信日8/30-10/1の入会メール通数。入会月は本文の利用開始日',
        pack: '一般会員（6ヶ月割）にはテンプレ約7〜8種。法人は原則セットなし'
      },
      labels: labels,
      gmail: {
        enrollThreads: enrollThreads.length,
        enrollMessages: enrollMessages,
        uniqueNameMonth: gmailPeople.length,
        septUnique: septPeople.length,
        septSixMonth: six.length,
        septCorporate: corp.length,
        optionThreads: changeThreads.length,
        optionMessages: changeMessages,
        optionParsedRows: changeRows.length
      },
      sheetSept: {
        total: sheetSept.length,
        sixMonth: sheetSept.filter(function (p) { return String(p.cat).indexOf('法人') === -1; }).length,
        corporate: sheetSept.filter(function (p) { return String(p.cat).indexOf('法人') !== -1; }).length
      },
      missingFromSheet: missing.map(function (p) {
        return { name: p.name, ym: p.ym, cat: p.cat, opCount: p.opCount, ops: p.ops };
      }),
      packFromMail: {
        complete: packOk,
        short: packShort.length,
        shortPeople: packShort
      },
      otherYmInWindow: gmailPeople.filter(function (p) { return p.ym !== monthText; }).map(function (p) {
        return { name: p.name, ym: p.ym, cat: p.cat };
      }),
      written: written
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function opStartAt_(optionName, monthOffset) {
  var log = "'" + OP_LOG_HELPER_SHEET_ + "'";
  var name = String(optionName).replace(/"/g, '""');
  var counted =
    'COUNTIFS(' + log + '!$A:$A,">="&EDATE($AB$5,' + monthOffset + '),' +
    log + '!$A:$A,"<"&EDATE($AB$5,' + (monthOffset + 1) + '),' +
    log + '!$C:$C,"' + name + '",' +
    log + '!$B:$B,"*利用開始*")';
  if (monthOffset !== 0) return '=' + counted;
  var live =
    'IFERROR(INDEX(\'' + OP_HELPER_SHEET_ + '\'!$D$3:$D$18,MATCH("' + name +
    '",\'' + OP_HELPER_SHEET_ + '\'!$A$3:$A$18,0)),' + counted + ')';
  return '=IF(TEXT($AB$5,"yymm")=TEXT(TODAY(),"yymm"),' + live + ',' + counted + ')';
}

function styleAllDataSheet_(sheet) {
  var ss = sheet.getParent();
  var maxR = sheet.getMaxRows();
  var maxC = sheet.getMaxColumns();
  sheet.clear();
  try { sheet.clearConditionalFormatRules(); } catch (e0) {}
  try { sheet.getRange(1, 1, maxR, maxC).breakApart(); } catch (e1) {}
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#111111');
  sheet.setFrozenRows(6);

  var ink = '#111111';
  var paper = '#FFFFFF';
  var zebra = '#F5F5F5';
  var soft = '#E8E8E8';
  var line = '#D4D4D4';
  var mute = '#666666';

  var now = new Date();
  var monthList = [];
  var jpYm = function (d) {
    return d.getFullYear() + '年' + (d.getMonth() + 1) + '月';
  };
  for (var i = 0; i < 24; i++) {
    monthList.push(jpYm(new Date(now.getFullYear(), now.getMonth() - i, 1)));
  }

  var needCols = 30;
  if (sheet.getMaxColumns() < needCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols - sheet.getMaxColumns());
  }
  sheet.getRange('AB1').setValue(UKETSUKE_SOURCE_ID_);
  sheet.getRange('AB5').setFormula('=DATE(VALUE(LEFT($B$2,4)),VALUE(REGEXEXTRACT($B$2,"年(\\d+)月")),1)');
  sheet.getRange('AB2').setFormula('=TEXT($AB$5,"yymm")');
  sheet.getRange('AB3').setFormula('=TEXT(EDATE($AB$5,-1),"yymm")');
  sheet.getRange('AB4').setFormula('=IMPORTRANGE($AB$1,"日報!B1")');
  sheet.getRange('AB6').setFormula(
    '=IMPORTRANGE("' + KENGAKU_SOURCE_ID_ + '","' + KENGAKU_SOURCE_SHEET_ + '!A1")'
  );
  try { sheet.hideColumns(28, 3); } catch (eHide) {}

  sheet.getRange('A1').setValue('経堂マスタ');
  sheet.getRange('A1:K1').merge();
  sheet.getRange('A2').setValue('年月');
  sheet.getRange('B2').setNumberFormat('@');
  sheet.getRange('B2').setValue(jpYm(new Date(now.getFullYear(), now.getMonth(), 1)));
  sheet.getRange('B2:C2').merge();
  sheet.getRange('B2').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(monthList, true)
      .setAllowInvalid(false)
      .build()
  );
  sheet.getRange('E2').setValue('今日');
  sheet.getRange('F2:G2').merge();
  sheet.getRange('F2').setFormula('=TEXT(TODAY(),"yyyy年m月d日")');

  var headerRow = 16;
  sheet.getRange(headerRow, 1).setValue('項目');
  sheet.getRange(headerRow, 2).setFormula('=TEXT(EDATE($AB$5,-4),"yyyy年m月")');
  sheet.getRange(headerRow, 3).setFormula('=TEXT(EDATE($AB$5,-3),"yyyy年m月")');
  sheet.getRange(headerRow, 4).setFormula('=TEXT(EDATE($AB$5,-2),"yyyy年m月")');
  sheet.getRange(headerRow, 5).setFormula('=TEXT(EDATE($AB$5,-1),"yyyy年m月")');
  sheet.getRange(headerRow, 6).setFormula('=TEXT($AB$5,"yyyy年m月")');
  sheet.getRange(headerRow, 7).setValue('着地見込');
  sheet.getRange(headerRow, 8).setValue('計画');
  sheet.getRange(headerRow, 9).setValue('進捗');
  sheet.getRange(headerRow, 10).setValue('対前月');

  var mOff = [-4, -3, -2, -1, 0];
  var five = function (builder) {
    return mOff.map(function (off) { return builder(off); });
  };

  var start = 17;
  var specs = [
    { name: '入会計画', vals: five(function (o) { return trendAt_('入会計画', o); }) },
    { name: '入会実績', vals: five(function (o) { return nippoOrTrend_('C13', '入 会  | 実績/見込', o); }), plan: '入会計画', pace: true },
    { name: '解除計画', vals: five(function (o) { return trendAt_('解除計画', o); }) },
    { name: '解除実績', vals: five(function (o) { return nippoOrTrend_('C15', '解 除  | 実績/見込', o); }), plan: '解除計画', pace: true },
    { name: '月初計画', vals: five(function (o) { return trendAt_('月初計画', o); }) },
    { name: '月初実績', vals: five(function (o) { return trendAt_('月初会員数 | 実績/見込', o); }), plan: '月初計画' },
    { name: '月末計画', vals: five(function (o) { return trendAt_('月末計画', o); }) },
    { name: '月末実績', vals: five(function (o) { return trendAt_('月末会員数 | 実績/見込', o); }), plan: '月末計画', paceMonthEnd: true },
    { name: '純増', vals: five(function (o) { return trendAt_('純増', o); }), paceNet: true, planFromPrev: true },
    { name: '休会', vals: five(function (o) { return trendAt_('休会', o); }), pace: true, planFromPrev: true },
    { name: '紹介', vals: five(function (o) { return trendAt_('紹介', o); }), pace: true, planFromPrev: true }
  ];
  var opFirst = specs.length;
  OP_NAMES_.forEach(function (opName) {
    specs.push({
      name: opName,
      op: true,
      pace: true,
      planFromPrev: true,
      vals: five(function (o) { return opStartAt_(opName, o); })
    });
  });
  var opLast = specs.length - 1;
  var opSumVals = five(function (o) {
    var col = String.fromCharCode(70 + o);
    return '=IFERROR(SUM(' + col + (start + opFirst) + ':' + col + (start + opLast) + '),)';
  });
  specs.push({ name: 'OP合計', vals: opSumVals, pace: true, planFromPrev: true, opTotal: true });
  specs = specs.concat([
    { name: '口コミ', vals: five(function (o) { return monthCountOff_('口コミ_経堂', o); }), pace: true, planFromPrev: true },
    { name: '見学体験', vals: five(function (o) { return monthCountOff_('見学体験申請', o); }), pace: true, planFromPrev: true },
    { name: 'レクチャー', vals: five(function (o) { return monthCountOff_('マシンレクチャー申込', o); }), pace: true, planFromPrev: true },
    { name: '販促乗換', vals: five(function (o) { return monthCountOff_('販促_乗り換え', o); }), pace: true, planFromPrev: true },
    { name: '販促ペア', vals: five(function (o) { return monthCountOff_('販促_紹介・ペア入会', o); }), pace: true, planFromPrev: true },
    { name: '学校関係者', vals: five(function (o) { return monthCountOff_('販促_学校関係者', o); }), pace: true, planFromPrev: true },
    { name: 'ラグビー割', vals: five(function (o) { return monthCountOff_('販促_ラグビー割', o); }), pace: true, planFromPrev: true },
    { name: '6ヶ月継続', vals: five(function (o) { return monthCountOff_('販促_6ヶ月継続', o); }), pace: true, planFromPrev: true }
  ]);

  var enrollRow = start + 1;
  var cancelRow = start + 3;
  var beginRow = start + 5;
  var body = specs.map(function (spec, idx) {
    var r = start + idx;
    var pace = '';
    if (spec.pace) pace = paceFrom_('F' + r);
    if (spec.paceNet) {
      pace = '=IF(AND(ISNUMBER(G' + enrollRow + '),ISNUMBER(G' + cancelRow + ')),G' + enrollRow + '-G' + cancelRow + ',)';
    }
    if (spec.paceMonthEnd) {
      pace =
        '=IF(TEXT($AB$5,"yymm")<>TEXT(TODAY(),"yymm"),,' +
        'IF(AND(ISNUMBER(F' + beginRow + '),ISNUMBER(G' + enrollRow + '),ISNUMBER(G' + cancelRow + ')),' +
        'F' + beginRow + '+G' + enrollRow + '-G' + cancelRow + ',))';
    }
    var planCell = '';
    if (spec.plan) planCell = trendAt_(spec.plan, 0);
    else if (spec.planFromPrev) planCell = planFromPrev_(r);
    var progress = (spec.plan || spec.planFromPrev) ? progressFormula_(r) : '';
    var delta = '=IF(AND(ISNUMBER(F' + r + '),ISNUMBER(E' + r + ')),F' + r + '-E' + r + ',)';
    return [spec.name].concat(spec.vals).concat([pace, planCell, progress, delta]);
  });
  sheet.getRange(start, 1, body.length, 10).setValues(body);

  var last = start + body.length - 1;
  var table = sheet.getRange(1, 1, last, 10);
  table
    .setFontFamily('Meiryo')
    .setFontColor(ink)
    .setVerticalAlignment('middle');

  sheet.getRange('A1:K1')
    .setBackground(ink)
    .setFontColor(paper)
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 32);

  sheet.getRange('A2:K2').setBackground(paper).setFontSize(10);
  sheet.getRange('A2').setFontWeight('bold').setFontColor(mute).setHorizontalAlignment('center');
  sheet.getRange('B2:C2')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setBackground(zebra);
  sheet.setRowHeight(2, 22);
  sheet.getRange('E2')
    .setFontWeight('bold')
    .setFontColor(mute)
    .setHorizontalAlignment('right');
  sheet.getRange('F2:G2')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setBackground(soft);

  sheet.getRange(headerRow, 1, 1, 10)
    .setBackground(ink)
    .setFontColor(paper)
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(headerRow, 22);

  sheet.getRange(start, 1, body.length, 10)
    .setBackground(paper)
    .setFontSize(10);
  sheet.getRange(start, 1, body.length, 1)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');
  sheet.getRange(start, 2, body.length, 9)
    .setHorizontalAlignment('right')
    .setFontWeight('bold');
  sheet.getRange(start, 2, body.length, 7).setNumberFormat('#,##0');
  sheet.getRange(start, 10, body.length, 1).setNumberFormat('#,##0;-#,##0');
  sheet.getRange(start, 9, body.length, 1).setNumberFormat('0%');
  sheet.getRange(start, 6, body.length, 1).setBackground(zebra);
  sheet.getRange(start, 7, body.length, 1).setBackground(soft);
  sheet.getRange(start, 8, body.length, 1).setBackground(zebra);

  for (var r = 0; r < body.length; r++) {
    if (r % 2 === 1) {
      sheet.getRange(start + r, 1, 1, 5).setBackground(zebra);
      sheet.getRange(start + r, 10).setBackground(zebra);
    }
    sheet.setRowHeight(start + r, 20);
  }

  var opTotalRow = start + opLast + 1;
  sheet.getRange(start + opFirst, 1, OP_NAMES_.length, 1)
    .setFontSize(10)
    .setBackground(ink)
    .setFontColor(paper);
  sheet.getRange(start + opFirst, 2, OP_NAMES_.length, 4).setBackground(zebra);
  sheet.getRange(opTotalRow, 1, 1, 10)
    .setBackground(ink)
    .setFontColor(paper)
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.getRange(opTotalRow, 6).setBackground(soft).setFontColor(ink);
  sheet.getRange(opTotalRow, 7).setBackground(line).setFontColor(ink);
  sheet.getRange(opTotalRow, 8).setBackground(zebra).setFontColor(ink);

  var promoFirst = opTotalRow + 1;
  sheet.getRange(promoFirst, 1, last - promoFirst + 1, 1)
    .setBackground('#333333')
    .setFontColor(paper)
    .setFontSize(10);

  sheet.getRange(headerRow, 1, last - headerRow + 1, 10).setBorder(
    true, true, true, true, true, true,
    line, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('A1:K1').setBorder(
    true, true, true, true, false, false,
    ink, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(headerRow, 1, 1, 10).setBorder(
    true, true, true, true, false, false,
    ink, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(start + opFirst, 1, OP_NAMES_.length + 1, 10).setBorder(
    true, true, true, true, true, true,
    line, SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.setColumnWidth(1, 150);
  var c;
  for (c = 2; c <= 11; c++) sheet.setColumnWidth(c, 80);

  addKyodoTodayBoard_(sheet);
  addKyodoMasterSideLists_(sheet);
  applyKyodoMasterFormats_(sheet, start, last);

  var leftover = ss.getSheets();
  for (var j = leftover.length - 1; j >= 0; j--) {
    var nm = leftover[j].getName();
    if (/^シート\d+$/.test(nm) && leftover[j].getLastRow() === 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(leftover[j]); } catch (eDel) {}
    }
  }
}

function todayMemberWhere_() {
  return (
    'Col1 >= date \'"&TEXT(TODAY(),"yyyy-mm-dd")&"\' and Col1 < date \'"&TEXT(TODAY()+1,"yyyy-mm-dd")&"\''
  );
}

function addKyodoTodayBoard_(sheet) {
  var ink = '#111111';
  var paper = '#FFFFFF';
  var zebra = '#F5F5F5';
  var soft = '#E8E8E8';
  var line = '#D4D4D4';
  var mute = '#666666';
  var irLive = function (a1) {
    return '=IFERROR(N(IMPORTRANGE($AB$1,"日報!' + a1 + '")),)';
  };
  var where = todayMemberWhere_();
  var enrollSh = "'" + NYUKAI_HELPER_SHEET_.replace(/'/g, "''") + "'";
  var withdrawSh = "'" + TAIKAI_HELPER_SHEET_.replace(/'/g, "''") + "'";

  sheet.getRange('A3:K15').setFontFamily('Meiryo').setVerticalAlignment('middle');
  sheet.getRange('A3:K3').setBackground(paper).setFontColor(paper).setFontSize(10).clearContent();
  sheet.setRowHeight(3, 10);

  var cards = [
    { label: '当日入会', formula: irLive('C9') },
    { label: '当日退会', formula: irLive('C11') },
    { label: '当月入会', formula: irLive('C13') },
    { label: '当月末退会', formula: irLive('C15') },
    { label: '当月移籍', formula: irLive('D14') }
  ];
  var i;
  for (i = 0; i < cards.length; i++) {
    var col = 2 + i * 2;
    sheet.getRange(4, col, 1, 2).merge();
    sheet.getRange(4, col).setValue(cards[i].label);
    sheet.getRange(4, col, 1, 2)
      .setBackground(zebra)
      .setFontColor(mute)
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.getRange(5, col, 2, 2).merge();
    sheet.getRange(5, col).setFormula(cards[i].formula);
    sheet.getRange(5, col, 2, 2)
      .setBackground(zebra)
      .setFontColor(ink)
      .setFontSize(20)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setNumberFormat('#,##0" 名"');
    sheet.getRange(4, col, 3, 2).setBorder(
      true, true, true, true, false, false,
      line, SpreadsheetApp.BorderStyle.SOLID
    );
  }
  sheet.getRange('A4:A6').setBackground(paper);
  sheet.setRowHeight(4, 22);
  sheet.setRowHeight(5, 28);
  sheet.setRowHeight(6, 28);

  sheet.getRange('B7:C7').merge();
  sheet.getRange('B7').setValue('今日の入会者');
  sheet.getRange('D7:E7').merge();
  sheet.getRange('D7').setValue('今日の退会者');
  sheet.getRange('B7:E7')
    .setBackground(ink)
    .setFontColor(paper)
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('A7').setBackground(paper);
  sheet.setRowHeight(7, 20);

  sheet.getRange('B8:E8').setValues([['氏名', '時刻', '氏名', '時刻']]);
  sheet.getRange('B8:E8')
    .setBackground(soft)
    .setFontColor(ink)
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('A8').setBackground(paper);
  sheet.setRowHeight(8, 18);

  sheet.getRange('B9').setFormula(
    '=IFERROR(QUERY(' + enrollSh + '!A2:E,"select Col2, Col1 where ' +
    where + ' order by Col1 desc limit 6",0),"")'
  );
  sheet.getRange('D9').setFormula(
    '=IFERROR(QUERY(' + withdrawSh + '!A2:F,"select Col2, Col1 where ' +
    where + ' and Col6 <> true order by Col1 desc limit 6",0),"")'
  );
  sheet.getRange('A9:E14').setBackground(paper).setFontColor(ink).setFontSize(10);
  sheet.getRange('B9:E14').setFontWeight('bold');
  sheet.getRange('B9:B14').setHorizontalAlignment('left');
  sheet.getRange('D9:D14').setHorizontalAlignment('left');
  sheet.getRange('C9:C14').setHorizontalAlignment('center').setNumberFormat('h:mm');
  sheet.getRange('E9:E14').setHorizontalAlignment('center').setNumberFormat('h:mm');
  for (i = 9; i <= 14; i++) {
    if ((i - 9) % 2 === 1) {
      sheet.getRange(i, 2, 1, 4).setBackground(zebra);
    }
    sheet.setRowHeight(i, 20);
  }
  sheet.getRange('B7:C14').setBorder(
    true, true, true, true, true, false,
    line, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('D7:E14').setBorder(
    true, true, true, true, true, false,
    line, SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.getRange('A15:J15').merge();
  sheet.getRange('A15').setValue('5ヶ月データベース');
  sheet.getRange('A15:J15')
    .setBackground(ink)
    .setFontColor(paper)
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('K15').setBackground(paper);
  sheet.setRowHeight(15, 22);
  addKyodoAiBriefPanel_(sheet);
}

function addKyodoAiBriefPanel_(sheet) {
  var ink = '#111111';
  var paper = '#FFFFFF';
  var zebra = '#F5F5F5';
  var soft = '#E8E8E8';
  var line = '#D4D4D4';
  try { sheet.getRange('F7:K14').breakApart(); } catch (e0) {}
  sheet.getRange('F7:K14')
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontColor(ink)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('F7:K7').merge();
  sheet.getRange('F7').setValue('今月の読み');
  sheet.getRange('F7:K7')
    .setBackground(ink)
    .setFontColor(paper)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('F8:K8').merge();
  sheet.getRange('F8').setValue('数値を読んで作成します');
  sheet.getRange('F8:K8')
    .setBackground(soft)
    .setFontColor('#333333')
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet.getRange('F9:F10').merge();
  sheet.getRange('F9').setValue('好調');
  sheet.getRange('F11:F12').merge();
  sheet.getRange('F11').setValue('課題');
  sheet.getRange('F13:F14').merge();
  sheet.getRange('F13').setValue('一手');
  sheet.getRange('F9:F14')
    .setBackground(ink)
    .setFontColor(paper)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('G9:K10').merge();
  sheet.getRange('G11:K12').merge();
  sheet.getRange('G13:K14').merge();
  sheet.getRange('G9').setValue('');
  sheet.getRange('G11').setValue('');
  sheet.getRange('G13').setValue('');
  sheet.getRange('G9:K10').setBackground(paper).setVerticalAlignment('top');
  sheet.getRange('G11:K12').setBackground(zebra).setVerticalAlignment('top');
  sheet.getRange('G13:K14').setBackground(paper).setVerticalAlignment('top');
  sheet.getRange('F7:K14').setBorder(
    true, true, true, true, true, false,
    line, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.setRowHeight(7, 22);
  sheet.setRowHeight(8, 22);
  var r;
  for (r = 9; r <= 14; r++) sheet.setRowHeight(r, 24);
}

function parseKyodoPct_(v) {
  if (typeof v === 'number' && isFinite(v)) return v > 3 ? v / 100 : v;
  var s = String(v == null ? '' : v).replace(/%/g, '').replace(/,/g, '').trim();
  if (!s) return null;
  var n = Number(s);
  if (!isFinite(n)) return null;
  return n > 3 ? n / 100 : n;
}

function collectKyodoKpiSnapshot_(sheet) {
  var last = Math.max(sheet.getLastRow(), 17);
  var table = sheet.getRange(16, 1, last - 15, 10).getDisplayValues();
  var items = [];
  for (var i = 1; i < table.length; i++) {
    var name = String(table[i][0] || '').trim();
    if (!name) continue;
    items.push({
      name: name,
      month: String(table[i][5] || ''),
      pace: String(table[i][6] || ''),
      plan: String(table[i][7] || ''),
      progress: String(table[i][8] || ''),
      vsPrev: String(table[i][9] || ''),
      pct: parseKyodoPct_(table[i][8])
    });
  }
  return {
    month: String(sheet.getRange('B2').getDisplayValue() || ''),
    today: String(sheet.getRange('F2').getDisplayValue() || ''),
    todayEnroll: String(sheet.getRange('B5').getDisplayValue() || ''),
    todayCancel: String(sheet.getRange('D5').getDisplayValue() || ''),
    monthEnroll: String(sheet.getRange('F5').getDisplayValue() || ''),
    monthCancel: String(sheet.getRange('H5').getDisplayValue() || ''),
    monthMove: String(sheet.getRange('J5').getDisplayValue() || ''),
    items: items
  };
}

function findKyodoKpi_(items, name) {
  for (var i = 0; i < items.length; i++) {
    if (items[i].name === name) return items[i];
  }
  return { name: name, month: '', pace: '', plan: '', progress: '', vsPrev: '', pct: null };
}

function kyodoBriefTicker_(snap) {
  var enroll = findKyodoKpi_(snap.items, '入会実績');
  var cancel = findKyodoKpi_(snap.items, '解除実績');
  var op = findKyodoKpi_(snap.items, 'OP合計');
  var net = findKyodoKpi_(snap.items, '純増');
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d H:mm');
  return (
    '入会' + (enroll.progress || '—') +
    '  解除' + (cancel.progress || '—') +
    '  OP' + (op.progress || '—') +
    '  純増' + (net.progress || '—') +
    '    ' + stamp + ' 自動'
  );
}

function rankedKyodoKpi_(items, wantHigh) {
  var ranked = items.filter(function (it) {
    if (it.pct == null) return false;
    if (it.name.indexOf('計画') !== -1) return false;
    if (it.name.indexOf('販促') === 0) return false;
    if (it.name === '口コミ' || it.name === 'レクチャー' || it.name === '学校関係者') return false;
    if (it.name === 'ラグビー割' || it.name === '6ヶ月継続') return false;
    return true;
  });
  ranked.sort(function (a, b) {
    return wantHigh ? b.pct - a.pct : a.pct - b.pct;
  });
  return ranked;
}

function buildKyodoAiBriefFallback_(snap) {
  var enroll = findKyodoKpi_(snap.items, '入会実績');
  var cancel = findKyodoKpi_(snap.items, '解除実績');
  var op = findKyodoKpi_(snap.items, 'OP合計');
  var net = findKyodoKpi_(snap.items, '純増');
  var highs = rankedKyodoKpi_(snap.items, true);
  var lows = rankedKyodoKpi_(snap.items, false);
  var goodBits = [];
  if (enroll.pct != null && enroll.pct >= 1) {
    goodBits.push('入会の着地は' + enroll.progress + '（実績' + enroll.month + '／計画' + enroll.plan + '）。');
  }
  var hi = 0;
  for (var i = 0; i < highs.length && hi < 2; i++) {
    if (highs[i].name === '入会実績' || highs[i].name === '口コミ') continue;
    if (highs[i].pct < 1) break;
    goodBits.push(highs[i].name + 'は' + highs[i].progress + '。');
    hi++;
  }
  if (!goodBits.length) goodBits.push('計画を明確に超えている項目はまだ少ない。');

  var badBits = [];
  if (cancel.pct != null && cancel.pct >= 1) {
    badBits.push('解除の着地は' + cancel.progress + '（実績' + cancel.month + '／計画' + cancel.plan + '）。');
  }
  var lo = 0;
  for (i = 0; i < lows.length && lo < 2; i++) {
    if (lows[i].pct == null || lows[i].pct >= 0.8) continue;
    if (lows[i].name === '解除実績') continue;
    badBits.push(lows[i].name + 'は' + lows[i].progress + '。');
    lo++;
  }
  if (!badBits.length) badBits.push('大きな割れはない。計画付近の項目を維持する。');

  var acts = [];
  if (enroll.pct != null && enroll.pct >= 1.2) {
    acts.push('入会はペース維持。見学・紹介の当日クロージングを外さない。');
  } else if (enroll.pct != null && enroll.pct < 0.9) {
    acts.push('入会が計画を下回る。残日で体験からの当日入会を厚くする。');
  }
  if (cancel.pct != null && cancel.pct >= 1) {
    acts.push('解除が計画以上。継続・休会の代替提案を退会面談で必ず出す。');
  }
  if (op.pct != null && op.pct < 0.95) {
    acts.push('OP着地は' + op.progress + '。入会者のパック欠けをその場で埋める。');
  }
  if (net.pct != null && net.pct < 0.85) {
    acts.push('純増が' + net.progress + '。入会維持より解除抑制を優先する。');
  }
  if (!acts.length) acts.push('数字は計画内。今日の入会者へOPの取りこぼしがないかだけ確認する。');

  var headline = '着地を見て今日動く';
  if (enroll.pct != null && enroll.pct >= 1.2 && cancel.pct != null && cancel.pct < 1) {
    headline = '入会は強い。解除を抑える';
  } else if (enroll.pct != null && enroll.pct >= 1.2) {
    headline = '入会は過熱。取りこぼしを防ぐ';
  } else if (cancel.pct != null && cancel.pct >= 1) {
    headline = '解除が計画を超えている';
  }

  return {
    headline: headline,
    good: goodBits.join(''),
    bad: badBits.join(''),
    action: acts.slice(0, 2).join(''),
    ticker: kyodoBriefTicker_(snap),
    via: 'fallback'
  };
}

function compactKyodoKpiForPrompt_(snap) {
  var focus = {
    '入会実績': 1, '解除実績': 1, '純増': 1, 'OP合計': 1,
    '安心サポートVIP': 1, '水素水': 1, 'オンラインレッスン': 1,
    'セルフエステ': 1, 'タンニング': 1, '口コミ': 1, '見学体験': 1
  };
  var rows = [];
  for (var i = 0; i < snap.items.length; i++) {
    var it = snap.items[i];
    if (!focus[it.name] && (it.pct == null || (it.pct >= 0.85 && it.pct <= 1.15))) continue;
    if (it.name.indexOf('計画') !== -1) continue;
    rows.push(
      it.name + ': 当月' + it.month + ' 着地' + it.pace +
      ' 計画' + it.plan + ' 進捗' + it.progress + ' 対前月' + it.vsPrev
    );
  }
  return rows.join('\n');
}

function parseKyodoAiJson_(raw) {
  var s = String(raw || '').trim();
  s = s.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/, '');
  var start = s.indexOf('{');
  var end = s.lastIndexOf('}');
  if (start < 0 || end <= start) return null;
  try {
    var obj = JSON.parse(s.substring(start, end + 1));
    if (!obj) return null;
    return {
      headline: String(obj.headline || '').replace(/\s+/g, ' ').trim(),
      good: String(obj.good || '').replace(/\s+/g, ' ').trim(),
      bad: String(obj.bad || '').replace(/\s+/g, ' ').trim(),
      action: String(obj.action || '').replace(/\s+/g, ' ').trim()
    };
  } catch (err) {
    return null;
  }
}

function buildKyodoAiBriefWithGemini_(snap) {
  var prompt = [
    'あなたはJOYFIT24経堂の店舗経営参謀です。',
    '与えた数字だけを根拠に、店長が今すぐ動ける短い日本語を返す。',
    '推測で事実を作らない。空欄の指標は触れない。口コミの極端な進捗は件数母数が小さいので過大評価しない。',
    '出力はJSONのみ。キーは headline, good, bad, action。',
    'headline は18字以内の一言。good は好調1〜2点（数字必須、80字以内）。',
    'bad は課題1〜2点（数字必須、80字以内）。action は今日〜今週の具体行動を2文（120字以内）。',
    '',
    '対象月: ' + snap.month + ' / 今日: ' + snap.today,
    '当日入会 ' + snap.todayEnroll + ' / 当日退会 ' + snap.todayCancel +
      ' / 当月入会 ' + snap.monthEnroll + ' / 当月末退会 ' + snap.monthCancel +
      ' / 当月移籍 ' + snap.monthMove,
    compactKyodoKpiForPrompt_(snap)
  ].join('\n');
  var gen = geminiGenerateText_(prompt);
  if (!gen || !gen.ok) return { error: (gen && gen.message) || 'gemini failed' };
  var parsed = parseKyodoAiJson_(gen.text);
  if (!parsed || !parsed.good || !parsed.action) {
    return { error: 'parse failed', raw: String((gen && gen.text) || '').slice(0, 240) };
  }
  return {
    brief: {
      headline: parsed.headline || '着地を見て今日動く',
      good: parsed.good,
      bad: parsed.bad || '',
      action: parsed.action,
      ticker: kyodoBriefTicker_(snap),
      via: 'gemini',
      model: gen.model || ''
    }
  };
}

function writeKyodoAiBrief_(sheet, brief) {
  var title = '今月の読み';
  if (brief.headline) title += '　' + brief.headline;
  sheet.getRange('F7').setValue(title);
  sheet.getRange('F8').setValue(brief.ticker || '');
  sheet.getRange('G9').setValue(brief.good || '');
  sheet.getRange('G11').setValue(brief.bad || '');
  sheet.getRange('G13').setValue(brief.action || '');
}

function refreshKyodoAiBrief_() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (eLock) {
    return { ok: false, message: 'lock timeout' };
  }
  var sheet;
  try {
    var ss = openWorkspaceSpreadsheet_();
    sheet = ss.getSheetByName(ALLDATA_SHEET_NAME_);
    if (!sheet) return { ok: false, message: '経堂マスタ not found' };
    var triggerInfo = ensureKyodoAiBriefTrigger_();
    SpreadsheetApp.flush();
    var snap = collectKyodoKpiSnapshot_(sheet);
    var enroll = findKyodoKpi_(snap.items, '入会実績');
    if (!snap.items.length || (enroll.progress === '' && enroll.month === '')) {
      writeKyodoAiBrief_(sheet, {
        headline: '',
        ticker: '数値の読み込み待ち',
        good: '',
        bad: '',
        action: ''
      });
      return { ok: false, message: 'kpi not ready', via: 'pending' };
    }
    var attempted = buildKyodoAiBriefWithGemini_(snap);
    var brief = (attempted && attempted.brief) || buildKyodoAiBriefFallback_(snap);
    writeKyodoAiBrief_(sheet, brief);
    return {
      ok: true,
      via: brief.via,
      model: brief.model || '',
      headline: brief.headline,
      ticker: brief.ticker,
      geminiError: (attempted && attempted.error) || '',
      aiAuthUrl: kyodoAiAuthUrl_(),
      aiTrigger: triggerInfo
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  } finally {
    try { lock.releaseLock(); } catch (eRel) {}
  }
}

function kyodoAiAuthUrl_() {
  try {
    var info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    if (info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      return info.getAuthorizationUrl() || '';
    }
  } catch (err) {}
  return '';
}

function refreshKyodoAiBriefTriggered_() {
  return refreshKyodoAiBrief_();
}

function ensureKyodoAiBriefTrigger_() {
  try {
    var fn = 'refreshKyodoAiBriefTriggered_';
    var triggers = ScriptApp.getProjectTriggers();
    var count = 0;
    var i;
    for (i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === fn) count++;
    }
    if (count < 1) {
      ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(8).create();
      count++;
    }
    if (count < 2) {
      ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(21).create();
    }
    return { ok: true, count: Math.max(count, 2) };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function applyKyodoMasterFormats_(sheet, start, last) {
  var progress = sheet.getRange(start, 9, last - start + 1, 1);
  var delta = sheet.getRange(start, 10, last - start + 1, 1);
  var rules = [];
  var progressSteps = [
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '>=1.5)', bg: '#111111', fg: '#FFFFFF' },
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '>=1.2)', bg: '#333333', fg: '#FFFFFF' },
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '>=1)', bg: '#666666', fg: '#FFFFFF' },
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '>=0.8)', bg: '#D4D4D4', fg: '#111111' },
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '<0.8)', bg: '#F5F5F5', fg: '#666666' }
  ];
  for (var i = 0; i < progressSteps.length; i++) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(progressSteps[i].formula)
        .setBackground(progressSteps[i].bg)
        .setFontColor(progressSteps[i].fg)
        .setRanges([progress])
        .build()
    );
  }
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#E8E8E8')
      .setFontColor('#111111')
      .setRanges([delta])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#F5F5F5')
      .setFontColor('#666666')
      .setRanges([delta])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$AA4=FALSE')
      .setBackground('#F5F5F5')
      .setFontColor('#666666')
      .setRanges([sheet.getRange('AA4:AA40')])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$AA4=TRUE')
      .setBackground('#111111')
      .setFontColor('#FFFFFF')
      .setRanges([sheet.getRange('AA4:AA40')])
      .build()
  );
  sheet.setConditionalFormatRules(rules);
}

function addKyodoMasterSideLists_(sheet) {
  var monthFilter =
    ' >= date \'"&TEXT($AB$5,"yyyy-mm-dd")&"\' and Col1 < date \'"&TEXT(EDATE($AB$5,1),"yyyy-mm-dd")&"\'';
  var kengaku = sheet.getParent().getSheetByName(KENGAKU_DEST_SHEET_);
  var kengakuLink = kengaku
    ? (sheet.getParent().getUrl() + '#gid=' + kengaku.getSheetId())
    : ('https://docs.google.com/spreadsheets/d/' + KENGAKU_SOURCE_ID_ + '/edit');

  sheet.getRange('L1:N1').merge();
  sheet.getRange('L1').setValue('今月の追加販促');
  sheet.getRange('L1:N2')
    .setFontFamily('Meiryo')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('L1:N1').setBackground('#111111').setFontColor('#FFFFFF').setFontSize(10);
  sheet.getRange('L2:N2').merge();
  sheet.getRange('L2').setFormula(
    '=HYPERLINK("https://docs.google.com/spreadsheets/d/' + PROMO_SOURCE_ID_ + '/edit","追加販促を開く")'
  );
  sheet.getRange('L2:N2').setBackground('#E8E8E8').setFontColor('#111111').setFontSize(10);
  sheet.getRange('L3:N3').setValues([['申請日時', '種別', '名前']]);
  sheet.getRange('L3:N3')
    .setBackground('#111111')
    .setFontColor('#FFFFFF')
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontWeight('bold');
  sheet.getRange('L4').setFormula(
    '=IFERROR(QUERY({' +
    '\'販促_乗り換え\'!A2:C;' +
    '\'販促_紹介・ペア入会\'!A2:C;' +
    '\'販促_ラグビー割\'!A2:C;' +
    '\'販促_6ヶ月継続\'!A2:C' +
    '},"select Col1,Col2,Col3 where Col1' +
    monthFilter +
    ' order by Col1 desc",0),"")'
  );
  sheet.getRange('L4:L40').setNumberFormat('yyyy/mm/dd HH:mm');
  sheet.getRange('L3:N40').setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle');

  sheet.getRange('R1:V1').merge();
  sheet.getRange('R1').setValue('今月の見学体験');
  sheet.getRange('R1:V2')
    .setFontFamily('Meiryo')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('R1:V1').setBackground('#111111').setFontColor('#FFFFFF').setFontSize(10);
  sheet.getRange('R2:V2').merge();
  sheet.getRange('R2').setFormula('=HYPERLINK("' + kengakuLink + '","見学体験申請シート")');
  sheet.getRange('R2:V2').setBackground('#E8E8E8').setFontColor('#111111').setFontSize(10);
  sheet.getRange('R3:V3').setValues([['申請日時', '区分', '名前', '希望日', '時刻']]);
  sheet.getRange('R3:V3')
    .setBackground('#111111')
    .setFontColor('#FFFFFF')
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontWeight('bold');
  sheet.getRange('R4').setFormula(
    '=IFERROR(QUERY(\'' + KENGAKU_DEST_SHEET_ + '\'!A2:I,"select Col1,Col2,Col3,Col8,Col9 where Col1' +
    monthFilter +
    ' order by Col1 desc",0),"")'
  );
  sheet.getRange('R4:R40').setNumberFormat('yyyy/mm/dd HH:mm');
  sheet.getRange('U4:U40').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('V4:V40').setNumberFormat('hh:mm');
  sheet.getRange('R3:V40').setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle');

  sheet.getRange('X1:AA1').merge();
  sheet.getRange('X1').setValue('今月の口コミ');
  sheet.getRange('X1:AA2')
    .setFontFamily('Meiryo')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('X1:AA1').setBackground('#111111').setFontColor('#FFFFFF').setFontSize(10);
  sheet.getRange('X2:AA2').merge();
  sheet.getRange('X2').setFormula(
    '=HYPERLINK("' + REVIEW_GRANT_APP_URL_ + '","口コミ付与アプリ")'
  );
  sheet.getRange('X2:AA2').setBackground('#E8E8E8').setFontColor('#111111').setFontSize(10);
  sheet.getRange('X3:AA3').setValues([['日時', '氏名', '会員番号', '付与']]);
  sheet.getRange('X3:AA3')
    .setBackground('#111111')
    .setFontColor('#FFFFFF')
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontWeight('bold');
  fillKyodoMasterReviews_(sheet);
  sheet.getRange('X4:X40').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('X3:AA40').setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle');

  sheet.setColumnWidth(12, 128);
  sheet.setColumnWidth(13, 96);
  sheet.setColumnWidth(14, 110);
  sheet.setColumnWidth(15, 16);
  sheet.setColumnWidth(16, 16);
  sheet.setColumnWidth(17, 16);
  sheet.setColumnWidth(18, 128);
  sheet.setColumnWidth(19, 52);
  sheet.setColumnWidth(20, 110);
  sheet.setColumnWidth(21, 88);
  sheet.setColumnWidth(22, 52);
  sheet.setColumnWidth(23, 16);
  sheet.setColumnWidth(24, 96);
  sheet.setColumnWidth(25, 110);
  sheet.setColumnWidth(26, 110);
  sheet.setColumnWidth(27, 52);
}

function parseMasterMonthStart_(sheet) {
  var raw = String(sheet.getRange('B2').getDisplayValue() || sheet.getRange('B2').getValue() || '');
  var m = raw.match(/(\d{4})年(\d{1,2})月/);
  if (!m) {
    var now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1);
  }
  return new Date(Number(m[1]), Number(m[2]) - 1, 1);
}

function fillKyodoMasterReviews_(sheet) {
  var start = parseMasterMonthStart_(sheet);
  var end = new Date(start.getFullYear(), start.getMonth() + 1, 1);
  var src = SpreadsheetApp.openById(REVIEW_SOURCE_ID_).getSheetByName(REVIEW_JOYFIT_SHEET_);
  if (!src) return;
  var last = Math.max(src.getLastRow(), 1);
  var vals = src.getRange(1, 1, last, 23).getValues();
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    if (String(row[1] || '') !== 'kyodo') continue;
    var ts = row[0];
    var t = ts instanceof Date ? ts : new Date(ts);
    if (!(t instanceof Date) || isNaN(t.getTime())) continue;
    if (t < start || t >= end) continue;
    var granted = row[21] === true || String(row[21]).toUpperCase() === 'TRUE';
    rows.push({
      ts: t,
      name: row[4] || '',
      code: row[5] || '',
      granted: granted,
      submissionId: String(row[15] || ''),
      eastRow: i + 1
    });
  }
  rows.sort(function (a, b) {
    if (a.granted !== b.granted) return a.granted ? 1 : -1;
    return b.ts.getTime() - a.ts.getTime();
  });
  var maxList = 37;
  if (rows.length > maxList) rows = rows.slice(0, maxList);
  var bodyEnd = 40;
  sheet.getRange(4, 24, bodyEnd - 3, 4).clearContent();
  sheet.getRange(4, 29, bodyEnd - 3, 2).clearContent();
  try { sheet.getRange(4, 27, bodyEnd - 3, 1).clearDataValidations(); } catch (e0) {}
  if (!rows.length) return;
  var display = [];
  var meta = [];
  for (var r = 0; r < rows.length; r++) {
    display.push([rows[r].ts, rows[r].name, rows[r].code, rows[r].granted]);
    meta.push([rows[r].submissionId, rows[r].eastRow]);
  }
  sheet.getRange(4, 24, display.length, 4).setValues(display);
  sheet.getRange(4, 29, meta.length, 2).setValues(meta);
  var aa = sheet.getRange(4, 27, display.length, 1);
  aa.setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build()
  );
}

function ensureKyodoMasterEditTrigger_(ss) {
  try {
    var dest = ss || openWorkspaceSpreadsheet_();
    var fn = 'onKyodoMasterEdit_';
    var triggers = ScriptApp.getProjectTriggers();
    var found = false;
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === fn) found = true;
    }
    if (!found) {
      ScriptApp.newTrigger(fn).forSpreadsheet(dest.getId()).onEdit().create();
    }
  } catch (err) {}
}

function onKyodoMasterEdit_(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== ALLDATA_SHEET_NAME_) return;
    var row = e.range.getRow();
    var col = e.range.getColumn();
    if (row === 2 && (col === 2 || col === 3)) {
      fillKyodoMasterReviews_(sheet);
      refreshKyodoAiBrief_();
      return;
    }
    if (col !== 27 || row < 4) return;
    var granted = e.range.getValue() === true;
    var eastRow = Number(sheet.getRange(row, 30).getValue());
    var submissionId = String(sheet.getRange(row, 29).getValue() || '');
    var code = String(sheet.getRange(row, 26).getValue() || '');
    var src = SpreadsheetApp.openById(REVIEW_SOURCE_ID_).getSheetByName(REVIEW_JOYFIT_SHEET_);
    if (!src) return;
    if (!eastRow || eastRow < 2) {
      eastRow = findEastReviewRow_(src, submissionId, code);
    }
    if (!eastRow) return;
    src.getRange(eastRow, 22).setValue(granted);
    src.getRange(eastRow, 23).setValue(granted ? new Date() : '');
  } catch (err) {}
}

function findEastReviewRow_(src, submissionId, code) {
  var last = Math.max(src.getLastRow(), 1);
  var vals = src.getRange(1, 1, last, 16).getValues();
  var sid = String(submissionId || '');
  var member = String(code || '');
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][1] || '') !== 'kyodo') continue;
    if (sid && String(vals[i][15] || '') === sid) return i + 1;
  }
  if (!member) return 0;
  for (var j = 1; j < vals.length; j++) {
    if (String(vals[j][1] || '') !== 'kyodo') continue;
    if (String(vals[j][5] || '') === member) return j + 1;
  }
  return 0;
}

var KYODO_TREND_SOURCE_ID_ = '1LOOUG97wuiKbhzl0BjJstXgLaaSCZAKNFdD8P3I5x_o';
var KYODO_TREND_SOURCE_SHEET_ = '経堂';
var KYODO_TREND_DEST_SHEET_ = '【経堂】会員動向';

function setupKyodoTrend_() {
  try {
    var dest = openWorkspaceSpreadsheet_();
    var existing = dest.getSheetByName(KYODO_TREND_DEST_SHEET_);
    if (existing) dest.deleteSheet(existing);
    var afterAllData = dest.getSheetByName(ALLDATA_SHEET_NAME_) ? 1 : 0;
    var sh = dest.insertSheet(KYODO_TREND_DEST_SHEET_, afterAllData);
    styleKyodoTrendSheet_(sh);
    return {
      ok: true,
      sheet: KYODO_TREND_DEST_SHEET_,
      sourceSheet: KYODO_TREND_SOURCE_SHEET_,
      range: 'A1:N115',
      workspaceUrl: dest.getUrl() + '#gid=' + sh.getSheetId()
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function styleKyodoTrendSheet_(sheet) {
  sheet.clear();
  try { sheet.clearConditionalFormatRules(); } catch (e0) {}
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#111111');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  sheet.getRange(1, 1).setFormula(
    '=IMPORTRANGE("' + KYODO_TREND_SOURCE_ID_ + '","' + KYODO_TREND_SOURCE_SHEET_ + '!A1:N115")'
  );

  var rows = 115;
  var cols = 14;
  var area = sheet.getRange(1, 1, rows, cols);
  area
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontColor('#111111')
    .setVerticalAlignment('middle')
    .setBackground('#FFFFFF');

  sheet.getRange(1, 1, 1, cols)
    .setBackground('#111111')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 28);

  for (var r = 2; r <= rows; r++) {
    sheet.setRowHeight(r, 22);
    if (r % 2 === 0) {
      sheet.getRange(r, 1, 1, cols).setBackground('#F5F5F5');
    }
  }

  sheet.getRange(1, 1, rows, 2).setFontWeight('bold').setHorizontalAlignment('left');
  sheet.getRange(2, 3, rows - 1, cols - 2).setHorizontalAlignment('right');

  area.setBorder(
    true, true, true, true, true, true,
    '#BDBDBD', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(1, 1, 1, cols).setBorder(
    true, true, true, true, false, false,
    '#111111', SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.setColumnWidth(1, 88);
  sheet.setColumnWidth(2, 168);
  for (var c = 3; c <= cols; c++) sheet.setColumnWidth(c, 72);
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
  sheet.setTabColor('#111111');

  var header = [['#', 'KIND', 'TITLE', 'URL']];
  sheet.getRange(1, 1, 1, 4).setValues(header);
  sheet.getRange(1, 1, 1, 4)
    .setBackground('#111111')
    .setFontColor('#FFFFFF')
    .setFontFamily('Roboto Mono')
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
    .setBackground('#FAFAFA')
    .setFontColor('#111111')
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  // ゼブラ
  for (var i = 0; i < rows.length; i++) {
    if (i % 2 === 1) {
      sheet.getRange(i + 2, 1, 1, 4).setBackground('#EEEEEE');
    }
  }

  sheet.getRange(2, 1, rows.length, 1).setFontColor('#888888').setHorizontalAlignment('right');
  sheet.getRange(2, 2, rows.length, 1).setFontColor('#555555').setHorizontalAlignment('center');
  sheet.getRange(2, 4, rows.length, 1).setFontColor('#1A73E8');

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
      return jsonOutput_({ ok: true, service: 'ryuta-workspace-gas', version: 'v2-vercel' });
    }
    if (api === 'listSheets') {
      return jsonOutput_(listWorkspaceSheets_());
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
    if (api === 'inspectRange') {
      return jsonOutput_(inspectSheetRange_(
        String((e.parameter && e.parameter.id) || ''),
        String((e.parameter && e.parameter.sheet) || ''),
        String((e.parameter && e.parameter.range) || ''),
        String((e.parameter && e.parameter.formulas) || '') === '1',
        String((e.parameter && e.parameter.format) || '') === '1'
      ));
    }
    if (api === 'compactUketsukeOpDuplicates') {
      return jsonOutput_(compactUketsukeOpDuplicates_());
    }
    if (api === 'auditNyukaiGmail') {
      return jsonOutput_(auditNyukaiGmail_());
    }
    if (api === 'backfillEnrollFromOp') {
      return jsonOutput_(backfillEnrollFromOpLog_());
    }
    if (api === 'auditOpMailCoverage') {
      return jsonOutput_(auditOpMailCoverage_());
    }
    if (api === 'auditGmailThreadShape') {
      return jsonOutput_(auditGmailThreadShape_());
    }
    if (api === 'auditUnbundledEnroll') {
      return jsonOutput_(rebuildFromUnbundledMail_(true));
    }
    if (api === 'rebuildFromUnbundledMail') {
      return jsonOutput_(rebuildFromUnbundledMail_(String((e.parameter && e.parameter.dry) || '') === '1'));
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
    if (api === 'setupKengakuImport') {
      return jsonOutput_(setupKengakuImport_());
    }
    if (api === 'setupAllData') {
      return jsonOutput_(setupAllData_());
    }
    if (api === 'refreshKyodoAiBrief') {
      return jsonOutput_(refreshKyodoAiBrief_());
    }
    if (api === 'setupKyodoTrend') {
      return jsonOutput_(setupKyodoTrend_());
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

function geminiGenerateText_(prompt) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    return { ok: false, message: 'GEMINI_API_KEY が未設定です。' };
  }
  var models = ['gemini-2.0-flash', 'gemini-2.5-flash', 'gemini-flash-latest', 'gemini-1.5-flash'];
  var lastErr = '';
  for (var i = 0; i < models.length; i++) {
    try {
      var url =
        'https://generativelanguage.googleapis.com/v1beta/models/' +
        models[i] +
        ':generateContent?key=' +
        encodeURIComponent(key);
      var res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: String(prompt || '') }] }]
        }),
        muteHttpExceptions: true
      });
      var code = res.getResponseCode();
      var json = JSON.parse(res.getContentText() || '{}');
      if (code >= 400) {
        lastErr = (json.error && json.error.message) || ('API エラー（コード ' + code + '）');
        if (code === 404 || code === 400) continue;
        return { ok: false, message: lastErr };
      }
      var text =
        json.candidates &&
        json.candidates[0] &&
        json.candidates[0].content &&
        json.candidates[0].content.parts &&
        json.candidates[0].content.parts[0] &&
        json.candidates[0].content.parts[0].text;
      if (text) return { ok: true, text: String(text).trim(), model: models[i] };
      lastErr = '返答を取得できませんでした。';
    } catch (err) {
      lastErr = String(err && err.message ? err.message : err);
    }
  }
  return { ok: false, message: lastErr || 'Gemini を呼べませんでした' };
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
    var gen = geminiGenerateText_(
      '以下はフィットネス施設のスタッフ日報「所感」欄の下書きです。ビジネスメール向けの敬語に整え、誤字脱字を修正し、300文字以内で簡潔にまとめてください。事実と意味は変えないでください。出力は所感の本文のみ（説明・見出し・引用符は不要）。\n\n' +
      String(rawText)
    );
    if (!gen.ok) return gen;
    return { ok: true, text: gen.text };
  } catch (e) {
    console.error(e);
    return { ok: false, message: String(e.message || e) };
  }
}

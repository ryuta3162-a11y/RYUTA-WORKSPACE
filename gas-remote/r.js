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
function inspectSheetRange_(id, sheetName, a1, wantFormulas) {
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
      formulas: wantFormulas ? rng.getFormulas() : []
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
    setupKengakuImport_(dest);
    var sh = dest.insertSheet(ALLDATA_SHEET_NAME_, 0);
    styleAllDataSheet_(sh);
    return {
      ok: true,
      sheet: ALLDATA_SHEET_NAME_,
      workspaceUrl: dest.getUrl() + '#gid=' + sh.getSheetId(),
      uketsukeId: UKETSUKE_SOURCE_ID_
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
  return '=IFERROR(SUM(IMPORTRANGE($Z$1,' + yyCell + '&"!' + a1 + '")),)';
}

function irN_(yyCell, a1) {
  return '=IFERROR(N(IMPORTRANGE($Z$1,' + yyCell + '&"!' + a1 + '")),)';
}

function irNippoIfCurrent_(a1) {
  return (
    '=IF(TEXT($Z$5,"yymm")=TEXT(TODAY(),"yymm"),IFERROR(N(IMPORTRANGE($Z$1,"日報!' + a1 + '")),),)'
  );
}

function irNippoSumIfCurrent_(a1) {
  return (
    '=IF(TEXT($Z$5,"yymm")=TEXT(TODAY(),"yymm"),IFERROR(SUM(IMPORTRANGE($Z$1,"日報!' + a1 + '")),),)'
  );
}

function trendAt_(label, monthOffset) {
  var sh = "'" + String(KYODO_TREND_DEST_SHEET_).replace(/'/g, "''") + "'";
  return (
    '=IFERROR(INDEX(' + sh + '!$C$2:$N$115,' +
    'MATCH("' + String(label).replace(/"/g, '""') + '",' + sh + '!$B$2:$B$115,0),' +
    'MATCH(MONTH(EDATE($Z$5,' + monthOffset + ')),' + sh + '!$C$1:$N$1,0)),)'
  );
}

function irSumOff_(monthOffset, a1) {
  return (
    '=IFERROR(SUM(IMPORTRANGE($Z$1,TEXT(EDATE($Z$5,' + monthOffset + '),"yymm")&"!' + a1 + '")),)'
  );
}

function monthCountOff_(sheetName, monthOffset) {
  return monthCountQuery_(
    sheetName,
    'EDATE($Z$5,' + monthOffset + ')',
    'EDATE($Z$5,' + (monthOffset + 1) + ')'
  );
}

function nippoOrTrend_(nippoA1, label, monthOffset) {
  if (monthOffset !== 0) return trendAt_(label, monthOffset);
  var trend = trendAt_(label, 0).replace(/^=/, '');
  return (
    '=IF(TEXT($Z$5,"yymm")=TEXT(TODAY(),"yymm"),IFERROR(N(IMPORTRANGE($Z$1,"日報!' +
    nippoA1 +
    '")),),' +
    trend +
    ')'
  );
}

function paceFrom_(actualA1) {
  return (
    '=IF(TEXT($Z$5,"yymm")<>TEXT(TODAY(),"yymm"),,' +
    'IF(AND(ISNUMBER(' + actualA1 + '),' + actualA1 + '<>"",DAY(TODAY())>0),' +
    'ROUND(' + actualA1 + '/DAY(TODAY())*DAY(EOMONTH($Z$5,0)),0),))'
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

function opStartAt_(optionName, monthOffset) {
  var log = "'" + OP_LOG_HELPER_SHEET_ + "'";
  var name = String(optionName).replace(/"/g, '""');
  var counted =
    'COUNTIFS(' + log + '!$A:$A,">="&EDATE($Z$5,' + monthOffset + '),' +
    log + '!$A:$A,"<"&EDATE($Z$5,' + (monthOffset + 1) + '),' +
    log + '!$C:$C,"' + name + '",' +
    log + '!$B:$B,"*利用開始*")';
  if (monthOffset !== 0) return '=' + counted;
  var live =
    'IFERROR(INDEX(\'' + OP_HELPER_SHEET_ + '\'!$D$3:$D$18,MATCH("' + name +
    '",\'' + OP_HELPER_SHEET_ + '\'!$A$3:$A$18,0)),' + counted + ')';
  return '=IF(TEXT($Z$5,"yymm")=TEXT(TODAY(),"yymm"),' + live + ',' + counted + ')';
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
  sheet.setFrozenRows(3);

  var now = new Date();
  var monthList = [];
  var jpYm = function (d) {
    return d.getFullYear() + '年' + (d.getMonth() + 1) + '月';
  };
  for (var i = 0; i < 24; i++) {
    monthList.push(jpYm(new Date(now.getFullYear(), now.getMonth() - i, 1)));
  }

  sheet.getRange('Z1').setValue(UKETSUKE_SOURCE_ID_);
  sheet.getRange('Z5').setFormula('=DATE(VALUE(LEFT($B$2,4)),VALUE(REGEXEXTRACT($B$2,"年(\\d+)月")),1)');
  sheet.getRange('Z2').setFormula('=TEXT($Z$5,"yymm")');
  sheet.getRange('Z3').setFormula('=TEXT(EDATE($Z$5,-1),"yymm")');
  sheet.getRange('Z4').setFormula('=IMPORTRANGE($Z$1,"日報!B1")');
  sheet.getRange('Z6').setFormula(
    '=IMPORTRANGE("' + KENGAKU_SOURCE_ID_ + '","' + KENGAKU_SOURCE_SHEET_ + '!A1")'
  );
  sheet.hideColumns(26, 1);

  sheet.getRange('A1').setValue('経堂マスタ');
  sheet.getRange('A1:J1').merge();
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

  sheet.getRange('A3').setValue('項目');
  sheet.getRange('B3').setFormula('=TEXT(EDATE($Z$5,-4),"yyyy年m月")');
  sheet.getRange('C3').setFormula('=TEXT(EDATE($Z$5,-3),"yyyy年m月")');
  sheet.getRange('D3').setFormula('=TEXT(EDATE($Z$5,-2),"yyyy年m月")');
  sheet.getRange('E3').setFormula('=TEXT(EDATE($Z$5,-1),"yyyy年m月")');
  sheet.getRange('F3').setFormula('=TEXT($Z$5,"yyyy年m月")');
  sheet.getRange('G3').setValue('着地見込');
  sheet.getRange('H3').setValue('計画');
  sheet.getRange('I3').setValue('進捗');
  sheet.getRange('J3').setValue('対前月');

  var mOff = [-4, -3, -2, -1, 0];
  var five = function (builder) {
    return mOff.map(function (off) { return builder(off); });
  };

  var start = 4;
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
        '=IF(TEXT($Z$5,"yymm")<>TEXT(TODAY(),"yymm"),,' +
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
    .setFontColor('#0F1419')
    .setVerticalAlignment('middle');

  sheet.getRange('A1:J1')
    .setBackground('#0F1419')
    .setFontColor('#FFFFFF')
    .setFontSize(15)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 28);

  sheet.getRange('A2:J2').setBackground('#FFFFFF').setFontSize(10);
  sheet.getRange('A2').setFontWeight('bold').setFontColor('#6B7280').setHorizontalAlignment('center');
  sheet.getRange('B2:C2')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setBackground('#E8EEF2');
  sheet.setRowHeight(2, 24);

  sheet.getRange('A3:J3')
    .setBackground('#1B2838')
    .setFontColor('#FFFFFF')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(3, 22);

  sheet.getRange(start, 1, body.length, 10)
    .setBackground('#FFFFFF')
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
  sheet.getRange(start, 6, body.length, 1).setBackground('#E8EEF2');
  sheet.getRange(start, 7, body.length, 1).setBackground('#D6E4F0');
  sheet.getRange(start, 8, body.length, 1).setBackground('#F3F4F6');

  for (var r = 0; r < body.length; r++) {
    if (r % 2 === 1) {
      sheet.getRange(start + r, 1, 1, 5).setBackground('#F4F6F8');
      sheet.getRange(start + r, 10).setBackground('#F4F6F8');
    }
    sheet.setRowHeight(start + r, 20);
  }

  var opTotalRow = start + opLast + 1;
  sheet.getRange(start + opFirst, 1, OP_NAMES_.length, 1)
    .setFontSize(9)
    .setBackground('#1B2838')
    .setFontColor('#FFFFFF');
  sheet.getRange(start + opFirst, 2, OP_NAMES_.length, 4).setBackground('#F8FAFC');
  sheet.getRange(opTotalRow, 1, 1, 10)
    .setBackground('#0F1419')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.getRange(opTotalRow, 6).setBackground('#E8EEF2').setFontColor('#0F1419');
  sheet.getRange(opTotalRow, 7).setBackground('#D6E4F0').setFontColor('#0F1419');
  sheet.getRange(opTotalRow, 8).setBackground('#F3F4F6').setFontColor('#0F1419');

  var promoFirst = opTotalRow + 1;
  sheet.getRange(promoFirst, 1, last - promoFirst + 1, 1)
    .setBackground('#1B2838')
    .setFontColor('#FFFFFF');

  sheet.getRange(1, 1, last, 10).setBorder(
    true, true, true, true, true, true,
    '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('A1:J1').setBorder(
    true, true, true, true, false, false,
    '#0F1419', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('A3:J3').setBorder(
    true, true, true, true, false, false,
    '#1B2838', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange(start + opFirst, 1, OP_NAMES_.length + 1, 10).setBorder(
    true, true, true, true, true, true,
    '#9AA5B1', SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.setColumnWidth(1, 160);
  for (var c = 2; c <= 6; c++) sheet.setColumnWidth(c, 80);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 64);
  sheet.setColumnWidth(9, 64);
  sheet.setColumnWidth(10, 64);
  sheet.setColumnWidth(11, 16);

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

function applyKyodoMasterFormats_(sheet, start, last) {
  var progress = sheet.getRange(start, 9, last - start + 1, 1);
  var delta = sheet.getRange(start, 10, last - start + 1, 1);
  var rules = [];
  var progressSteps = [
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '>=1.5)', bg: '#0F1419', fg: '#FFFFFF' },
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '>=1.2)', bg: '#1B2838', fg: '#FFFFFF' },
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '>=1)', bg: '#D6E4F0', fg: '#0F1419' },
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '>=0.8)', bg: '#E8EEF2', fg: '#0F1419' },
    { formula: '=AND(ISNUMBER(I' + start + '),I' + start + '<0.8)', bg: '#9CA3AF', fg: '#FFFFFF' }
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
      .setBackground('#D6E4F0')
      .setFontColor('#1B2838')
      .setRanges([delta])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#E5E7EB')
      .setFontColor('#4B5563')
      .setRanges([delta])
      .build()
  );
  sheet.setConditionalFormatRules(rules);
}

function addKyodoMasterSideLists_(sheet) {
  var monthFilter =
    ' >= date \'"&TEXT($Z$5,"yyyy-mm-dd")&"\' and Col1 < date \'"&TEXT(EDATE($Z$5,1),"yyyy-mm-dd")&"\'';
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
  sheet.getRange('L1:N1').setBackground('#0F1419').setFontColor('#FFFFFF').setFontSize(11);
  sheet.getRange('L2:N2').merge();
  sheet.getRange('L2').setFormula(
    '=HYPERLINK("https://docs.google.com/spreadsheets/d/' + PROMO_SOURCE_ID_ + '/edit","追加販促を開く")'
  );
  sheet.getRange('L2:N2').setBackground('#E8EEF2').setFontColor('#1B2838').setFontSize(9);
  sheet.getRange('L3:N3').setValues([['申請日時', '種別', '名前']]);
  sheet.getRange('L3:N3')
    .setBackground('#1B2838')
    .setFontColor('#FFFFFF')
    .setFontFamily('Meiryo')
    .setFontSize(9)
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
  sheet.getRange('L3:N40').setFontFamily('Meiryo').setFontSize(9).setVerticalAlignment('middle');
  for (var pr = 4; pr <= 40; pr++) sheet.setRowHeight(pr, 20);

  sheet.getRange('R1:V1').merge();
  sheet.getRange('R1').setValue('今月の見学体験');
  sheet.getRange('R1:V2')
    .setFontFamily('Meiryo')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('R1:V1').setBackground('#0F1419').setFontColor('#FFFFFF').setFontSize(11);
  sheet.getRange('R2:V2').merge();
  sheet.getRange('R2').setFormula('=HYPERLINK("' + kengakuLink + '","見学体験申請シート")');
  sheet.getRange('R2:V2').setBackground('#D6E4F0').setFontColor('#1B2838').setFontSize(9);
  sheet.getRange('R3:V3').setValues([['申請日時', '区分', '名前', '希望日', '時刻']]);
  sheet.getRange('R3:V3')
    .setBackground('#1B2838')
    .setFontColor('#FFFFFF')
    .setFontFamily('Meiryo')
    .setFontSize(9)
    .setFontWeight('bold');
  sheet.getRange('R4').setFormula(
    '=IFERROR(QUERY(\'' + KENGAKU_DEST_SHEET_ + '\'!A2:I,"select Col1,Col2,Col3,Col8,Col9 where Col1' +
    monthFilter +
    ' order by Col1 desc",0),"")'
  );
  sheet.getRange('R4:R40').setNumberFormat('yyyy/mm/dd HH:mm');
  sheet.getRange('U4:U40').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('V4:V40').setNumberFormat('hh:mm');
  sheet.getRange('R3:V40').setFontFamily('Meiryo').setFontSize(9).setVerticalAlignment('middle');

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
        String((e.parameter && e.parameter.formulas) || '') === '1'
      ));
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

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
    .addItem('しまう', 'hubCloseWork_')
    .addItem('トップを表示', 'hubShowHome_')
    .addToUi();
  ui.createMenu('数値更新')
    .addItem('受付状況表の数値を更新（入会・退会・OP）', 'refreshReceptionNumbersFromMenu')
    .addItem('前回の更新時刻を確認', 'showReceptionRefreshStatus')
    .addItem('入会者一覧を元シートと連動し直す', 'relinkJoinListFromMenu')
    .addSeparator()
    .addItem('請求報告の自動連携を有効にする（初回のみ）', 'setupBillingLinkFromMenu')
    .addItem('請求報告を今すぐ取り込む', 'refreshBillingFromMenu')
    .addItem('会員分析を作り直す', 'refreshMemberAnalysisFromMenu')
    .addItem('移籍・復会・紹介を日報へ反映', 'refreshJoinBreakdownFromMenu')
    .addItem('日報の契約・解約をB1の月に連動', 'refreshNippoOpFromMenu')
    .addToUi();
  try {
    linkJoinListLive_(SpreadsheetApp.getActiveSpreadsheet(), false);
  } catch (eLink) {}
  try {
    linkUnpaidFollowup_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (eFollow) {}
  try {
    ensureEnjoyPointLink_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (eEnjoy) {}
  try {
    ensureMasterMemberNo_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (eMemberNo) {}
  try {
    ensureMasterApplyCheckmarks_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (eCheck) {}
  try {
    syncJoinBreakdownIfChanged_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (eBd) {}
}

var JOIN_LIST_FIT365_ID_ = '1BbExBUCfyq1cfNqw4TvlwUriL-AfvghU9XT6McdzGTQ';

function joinListImportFormula_(sourceId) {
  return '=LET(d,IFERROR(IMPORTRANGE("' + sourceId + '","\'入会者一覧＋自動メール管理\'!A1:E"),{"元シート読込エラー","","","",""}),' +
    'f,FILTER(d,(SEQUENCE(ROWS(d))=1)+(INDEX(d,,2)<>"")),' +
    'MAKEARRAY(ROWS(f),5,LAMBDA(r,c,LET(v,INDEX(f,r,c),' +
    'IF(OR(r=1,c<4),v,IF(OR(v=TRUE,UPPER(v&"")="TRUE"),"☑","☐"))))))';
}

/**
 * 入会者一覧（経堂 A:E / FIT365 F:J）を IMPORTRANGE で元シートと常時連動させる。
 * このプロジェクトは script.scriptapp 権限がなく時間トリガーを作れないため、値コピー同期は止まる。
 */
function linkJoinListLive_(ss, force) {
  var sh = ss.getSheetByName(JOIN_LIST_SHEET_);
  if (!sh) return { ok: false, message: 'sheet not found' };
  var a2 = String(sh.getRange('A2').getFormula() || '');
  if (!force && /IMPORTRANGE/i.test(a2)) return { ok: true, skipped: true };
  var last = Math.max(sh.getMaxRows(), 3);
  sh.getRange(2, 1, last - 1, 10).clearContent().clearDataValidations();
  sh.getRange('A2').setFormula(joinListImportFormula_(MACHINE_SOURCE_ID_));
  sh.getRange('F2').setFormula(joinListImportFormula_(JOIN_LIST_FIT365_ID_));
  sh.getRange(3, 1, last - 2, 1).setNumberFormat('yyyy/mm/dd');
  sh.getRange(3, 6, last - 2, 1).setNumberFormat('yyyy/mm/dd');
  sh.getRange(3, 4, last - 2, 2).setHorizontalAlignment('center');
  sh.getRange(3, 9, last - 2, 2).setHorizontalAlignment('center');
  sh.getRange('D1').setValue('元シートと自動連動（編集は元シートで）');
  sh.getRange('I1').setValue('元シートと自動連動（編集は元シートで）');
  return { ok: true, linked: true };
}

function relinkJoinListFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = linkJoinListLive_(ss, true);
  ss.toast(r.ok ? '入会者一覧を元シートと連動しました' : String(r.message), '入会者一覧', 8);
}

function onEdit(e) {}

function onSelectionChange(e) {
  try {
    handleHubHomeSelect_(e);
  } catch (err) {}
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
  ss.toast('更新中です…（2〜4分）完了したらお知らせします', '数値更新', 150);

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
      reloadReceptionImports_(ss);
      ss.toast('更新が完了しました（' + s.lastRefreshed + '）', '数値更新', 15);
      return s;
    }
  }
  ss.toast('まだ終わっていません。少ししてから「前回の更新時刻を確認」で見てください。', '数値更新', 15);
  return r;
}

/**
 * IMPORTRANGE は元が更新されても数分〜数十分キャッシュを返すことがある。
 * 経堂マスタの IMPORTRANGE が参照する元ID（AF1）を書き直して、読み直させる。
 */
function reloadReceptionImports_(ss) {
  var sh = ss.getSheetByName('経堂マスタ');
  if (!sh) return;
  var applyCol = masterApplyCol_(sh);
  var cell = sh.getRange(1, applyCol ? applyCol + 16 : 32);
  var id = cell.getValue();
  if (!id) return;
  cell.setValue('');
  SpreadsheetApp.flush();
  cell.setValue(id);
  SpreadsheetApp.flush();
}

function showReceptionRefreshStatus() {
  var r = callReceptionRefreshApi_('refreshStatus');
  var msg = !r || !r.ok
    ? '状態を取得できませんでした'
    : r.pending
      ? '更新中です。少しお待ちください。'
      : r.error
        ? '前回の更新でエラー: ' + r.error + '（最後に成功: ' + (r.lastRefreshed || '記録なし') + '）'
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
  sheet.setTabColor(hubTabColorFor_(sheet.getName()));
  sheet.setFrozenRows(1);

  var dn = dnTheme_();
  var headerRange = sheet.getRange(1, 1, 1, usedCols);
  headerRange
    .setBackground(dn.ink)
    .setFontColor(dn.paper)
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setWrap(true);
  sheet.setRowHeight(1, 32);

  if (usedRows >= 2) {
    var body = sheet.getRange(2, 1, usedRows - 1, usedCols);
    body
      .setBackground(dn.paper)
      .setFontColor(dn.ink)
      .setFontFamily('Meiryo')
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
      .setFontFamily('Meiryo')
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 32);
    if (bodyRows >= 1) {
      sheet.getRange(2, 1, bodyRows, cols)
        .setBackground(dnChrome.paper)
        .setFontColor(dnChrome.ink)
        .setFontFamily('Meiryo')
        .setFontSize(10)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('left');
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
            .whenFormulaSatisfied('=$' + colLetter + '2=TRUE')
            .setBackground(dnTheme_().cream)
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

/**
 * マシンレクチャー申込：新しい行が入っても同じ背景色になるよう、A列に値がある行を条件付き書式で塗る。
 * 色は手直し済みの A3 の背景色に合わせる。他の見た目（文字・列幅など）は触らない。
 */
function tintLectureRows_() {
  try {
    var sh = openWorkspaceSpreadsheet_().getSheetByName('マシンレクチャー申込');
    if (!sh) return { ok: false, message: 'sheet not found' };
    var bg = String(sh.getRange('A3').getBackground() || '#ffffff');
    if (/^#?f{6}$/i.test(bg.replace('#', '')) || bg === '#ffffff') bg = '#efefef';
    var cols = Math.max(sh.getLastColumn(), 7);
    var rng = sh.getRange(3, 1, sh.getMaxRows() - 2, cols);
    var formula = '=$A3<>""';
    var rules = sh.getConditionalFormatRules().filter(function (r) {
      var c = r.getBooleanCondition();
      return !(c && c.getCriteriaValues && String(c.getCriteriaValues()[0]) === formula);
    });
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
      .setBackground(bg).setRanges([rng]).build());
    sh.setConditionalFormatRules(rules);
    return { ok: true, color: bg, range: rng.getA1Notation(), rules: rules.length };
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

/** 入会者一覧：IMPORTRANGE 維持のまま枠線削除＋チェック列を整列 */
function formatJoinListMirrorKeepImport_() {
  try {
    var dest = openWorkspaceSpreadsheet_();
    var sh = dest.getSheetByName(JOIN_LIST_SHEET_);
    if (!sh) sh = dest.getSheetByName('入会者一覧+自動メール管理');
    if (!sh) return { ok: false, message: 'sheet not found: ' + JOIN_LIST_SHEET_ };

    var a1 = '';
    try {
      a1 = String(sh.getRange(1, 1).getFormula() || '');
    } catch (eF) {}
    if (!/IMPORTRANGE/i.test(a1)) {
      return {
        ok: false,
        message: 'A1 が IMPORTRANGE ではありません（値同期シートの可能性）。数式を確認してください。',
        a1: a1.slice(0, 120)
      };
    }

    var cols = Math.max(sh.getLastColumn(), 5);
    var headers = sh.getRange(1, 1, 1, cols).getDisplayValues()[0];
    var joinChecks = resolveJoinListCheckCols_(headers, cols);
    var out = formatImportMirrorChrome_(sh, cols, joinChecks, headers);
    try {
      removeSheetFilterSafe_(sh);
    } catch (eFil) {}

    // マシンレクチャー申込：手直しした見た目を尊重。枠線・変なフィルタだけ掃除（色や列幅は触らない）
    var lecture = dest.getSheetByName('マシンレクチャー申込');
    var lectureOut = null;
    if (lecture) {
      lectureOut = tidyImportMirrorKeepLook_(lecture);
    }

    return {
      ok: true,
      joinList: out,
      checkCols: joinChecks,
      layout: {
        kyodo: 'A:E（D/E=チェック）',
        fit365: cols >= 10 ? 'F起点（I/J=チェック＝元D/E）' : '未接続または列不足'
      },
      machineLecture: lectureOut,
      tip:
        '経堂 D/E と FIT365 I/J を同じチェックボックス表示にしています。' +
        '値は各 IMPORTRANGE のまま自動更新。操作は元スプシ側。'
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
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
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontWeight('bold');
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 6)
      .setBackground(dn.paper)
      .setFontColor(dn.ink)
      .setFontFamily('Meiryo')
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
    .setFontFamily('Meiryo')
    .setFontSize(10)
    .setFontWeight('bold');
  if (values.length >= 2) {
    sh.getRange(2, 1, values.length - 1, lastCol)
      .setBackground(dnJoin.paper)
      .setFontColor(dnJoin.ink)
      .setFontFamily('Meiryo')
      .setFontSize(10);
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
  try { formatJoinListMirrorKeepImport_(); } catch (e3) {}
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
        .setFontFamily('Meiryo')
        .setFontSize(10)
        .setFontWeight('bold');
      mirror.setRowHeight(1, 32);
      var lastBody = Math.min(mirror.getMaxRows(), 500);
      if (lastBody >= 2) {
        mirror.getRange(2, 1, lastBody - 1, cols)
          .setBackground(dnRev.paper)
          .setFontColor(dnRev.ink)
          .setFontFamily('Meiryo')
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
        .setFontFamily('Meiryo')
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
      .setFontFamily('Meiryo')
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
    .setFontFamily('Meiryo')
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
    .setFontFamily('Meiryo')
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
      return jsonOutput_({ ok: true, service: 'ryuta-workspace-gas', version: 'v2-hub-hide-keys' });
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
    if (api === 'restyleHubLook') {
      return jsonOutput_(restyleHubLook_());
    }
    if (api === 'setupHubHome') {
      return jsonOutput_(setupHubHome_());
    }
    if (api === 'hubOpen') {
      return jsonOutput_(hubOpenNamed_(String((e.parameter && e.parameter.name) || '')));
    }
    if (api === 'hubClose') {
      return jsonOutput_(hubCloseWork_());
    }
    if (api === 'tintLectureRows') {
      return jsonOutput_(tintLectureRows_());
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
    if (api === 'syncJoinBreakdown') {
      return jsonOutput_(syncJoinBreakdown_(openWorkspaceSpreadsheet_()));
    }
    if (api === 'ensureNippoOpByB1') {
      return jsonOutput_(ensureNippoOpByB1_(true));
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
      return jsonOutput_(setupKengakuJoinLive_(String((e.parameter && e.parameter.confirm) || '')));
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

/** デスノート寄りの配色。クリーム紙・墨・血赤。大きくは変えない */
function dnTheme_() {
  return {
    ink: '#140C0C',
    paper: '#F4ECD9',
    cream: '#E9DCC6',
    blood: '#8B1216',
    apple: '#B91C1C',
    ash: '#5C5346',
    line: '#C9B896',
    ghost: '#FBF6EA'
  };
}

function hubTabColorFor_(name) {
  var t = dnTheme_();
  if (name === HUB_HOME_SHEET_) return '#111111';
  if (name === '経堂マスタ') return t.ink;
  if (name.indexOf('未納') === 0 || name === '請求・回収実績') return t.blood;
  if (name.indexOf('見学体験') === 0 && name.indexOf('backup') === -1) return '#6B3A1F';
  if (name.indexOf('販促_') === 0) return '#3F4A28';
  if (name.indexOf('口コミ') === 0) return '#5A1F2A';
  if (name.indexOf('マシンレクチャー') === 0) return '#2C4A3A';
  if (name.indexOf('入会者一覧') === 0) return '#3A3228';
  if (name.indexOf('会員動向') !== -1) return '#4A4038';
  if (name === 'URL一覧') return '#6A6258';
  if (name.indexOf('経堂_') === 0) return '#4A4440';
  if (name === 'Tasks' || name === 'WorkspaceSync') return '#7A746C';
  if (/backup|シート\d+/.test(name)) return '#B0A89C';
  return t.ash;
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
  sh.getRange(row, HUB_KEY_BASE_ + col)
    .setValue(value)
    .setFontColor('#FFFFFF')
    .setBackground('#FFFFFF')
    .setFontSize(1)
    .setFontWeight('normal')
    .setHorizontalAlignment('left');
}

function hubUi_() {
  return {
    bg: '#F3F3F3',
    card: '#FFFFFF',
    ink: '#0A0A0A',
    mute: '#6A6A6A',
    line: '#0A0A0A',
    onBg: '#0A0A0A',
    onFg: '#FFFFFF',
    shut: '#0A0A0A',
    rail: '#0A0A0A'
  };
}

function hubCatalog_(ss) {
  var items = [
    { group: '数字', name: '経堂マスタ', title: 'マスタ' },
    { group: '数字', name: '【経堂】会員動向', title: '動向' },
    { group: '未納', name: '未納管理', title: '今月' },
    { group: '未納', name: '未納管理_推移', title: '推移' },
    { group: '未納', name: '請求・回収実績', title: '請求' },
    { group: '数字', name: '会員分析', title: '会員分析' },
    { group: '現場', name: '見学体験申請', title: '見学' },
    { group: '現場', name: '口コミ_経堂', title: '口コミ' },
    { group: '現場', name: 'マシンレクチャー申込', title: 'レクチャー' },
    { group: '現場', name: '入会者一覧＋自動メール管理', title: '入会者' }
  ];
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (/^販促_/.test(n) && n.indexOf('backup') === -1) {
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

function hubSourceLinks_() {
  return [
    { title: '経堂　受付状況表', url: 'https://docs.google.com/spreadsheets/d/14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w/edit' },
    { title: '26年度未納管理ドライブ【経堂】', url: 'https://docs.google.com/spreadsheets/d/' + UNPAID_SOURCE_ID_ + '/edit' },
    { title: '経堂　見学・体験フォーム', url: 'https://docs.google.com/spreadsheets/d/1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY/edit' },
    { title: 'JOYFIT24経堂追加販促', url: 'https://docs.google.com/spreadsheets/d/1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E/edit' },
    { title: 'EAST口コミ回答者', url: 'https://docs.google.com/spreadsheets/d/13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM/edit' },
    { title: '20分マシンレクチャー・自動送信メール', url: 'https://docs.google.com/spreadsheets/d/1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8/edit' },
    { title: '口コミ付与アプリ', url: REVIEW_GRANT_APP_URL_ },
    { title: ENJOY_POINT_TITLE_, url: ENJOY_POINT_URL_ }
  ];
}

var ENJOY_POINT_URL_ = 'https://main.d5z4bnw4wyrxn.amplifyapp.com/store-settings/basic/points?clubCode=1304';
var ENJOY_POINT_TITLE_ = 'エンジョイポイント付与（口コミ確認後）';

/** 経堂マスタ「当月の申請」の表の先頭列（1行目に「当月の申請」の式がある列）。見つからなければ 0 */
function masterApplyCol_(sh) {
  var f = sh.getRange(1, 1, 1, Math.min(sh.getLastColumn(), 60)).getFormulas()[0];
  for (var c = 0; c < f.length; c++) {
    if (f[c].indexOf('当月の申請') >= 0) return c + 1;
  }
  return 0;
}

/** 経堂マスタ「当月の申請」の QUERY の TRUE/FALSE を ☑/☐ で表示する。包み済みなら何もしない */
function ensureMasterApplyCheckmarks_(ss) {
  var sh = ss.getSheetByName('経堂マスタ');
  if (!sh) return;
  var col = masterApplyCol_(sh);
  if (!col) return;
  var cell = sh.getRange(4, col);
  var f = cell.getFormula();
  if (!f || /^=LET\(src_,/.test(f)) return;
  cell.setFormula('=LET(src_,' + f.slice(1) + ',MAP(src_,LAMBDA(v_,IF(ISLOGICAL(v_),IF(v_,"☑","☐"),IF(v_&""="","",v_)))))');
}

/**
 * 経堂マスタ「当月の申請」の左に会員番号の列を置き、累計入会データの氏名（空白を無視）から会員番号を関数で引く。
 * 同名が複数なら電話番号（下10桁）で絞り、絞れなければ「要確認」。口コミのように会員番号の列がある表はそれを出す。
 */
function masterMemberNoFormula_(col) {
  var a = columnLetter_(col);
  var z = columnLetter_(col + 15);
  var src = "'" + MEMBER_JOIN_SRC_ + "'!";
  return '=LET(h,' + a + '3:' + z + '3,d,' + a + '4:' + z + '300,' +
    'ni,IFERROR(MATCH("被紹介者",h,0),IFERROR(MATCH("氏名",h,0),IFERROR(MATCH("名前",h,0),0))),' +
    'ti,IFERROR(MATCH("被紹介者電話",h,0),IFERROR(MATCH("電話",h,0),IFERROR(MATCH("連絡先",h,0),0))),' +
    'ki,IFERROR(MATCH("会員番号",h,0),0),' +
    'n,COUNTA(' + a + '4:' + a + '300),' +
    'no,' + src + 'F2:F,' +
    'nd,HSTACK(no,' + src + 'M2:M),' +
    'nm,ARRAYFORMULA(REGEXREPLACE(' + src + 'G2:G&"","[\\s　]","")),' +
    'tl,ARRAYFORMULA(RIGHT(REGEXREPLACE(' + src + 'L2:L&"","\\D",""),9)),' +
    'IF(n=0,"",MAP(SEQUENCE(n),LAMBDA(i,IF(ki>0,INDEX(d,i,ki),' +
    'LET(k,IF(ni=0,"",REGEXREPLACE(INDEX(d,i,ni)&"","[\\s　]","")),' +
    't,IF(ti=0,"",RIGHT(REGEXREPLACE(INDEX(d,i,ti)&"","\\D",""),9)),' +
    'c,IF(k="",0,IFERROR(ROWS(FILTER(no,nm=k)),0)),' +
    'fp,IF(OR(c<2,LEN(t)<9),"",IFERROR(SORT(FILTER(nd,nm=k,tl=t),2,FALSE),"")),' +
    'IF(c=0,"",IF(c=1,XLOOKUP(k,nm,no),IF(INDEX(fp,1,1)<>"",INDEX(fp,1,1),INDEX(SORT(FILTER(nd,nm=k),2,FALSE),1,1))))))))))';
}

function ensureMasterMemberNo_(ss) {
  var sh = ss.getSheetByName('経堂マスタ');
  if (!sh) return;
  var col = masterApplyCol_(sh);
  if (!col || col < 2) return;
  var memberCol = col - 1;
  if (String(sh.getRange(3, memberCol).getValue()) !== '会員番号') {
    sh.insertColumnBefore(col);
    memberCol = col;
    col = col + 1;
    var last = Math.min(sh.getMaxRows(), 300);
    sh.getRange(3, col, last - 2, 1).copyTo(sh.getRange(3, memberCol, last - 2, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sh.getRange(1, memberCol, 2, 1).setBackground(sh.getRange(1, col).getBackground());
    sh.getRange(3, memberCol).setValue('会員番号');
    sh.getRange(4, memberCol, last - 3, 1).setNumberFormat('0').setHorizontalAlignment('center').setFontWeight('bold');
    sh.setColumnWidth(memberCol, 100);
  }
  var tab = sh.getRange(2, memberCol);
  if (tab.getDataValidation()) {
    tab.clearDataValidations().clearContent().clearFormat()
      .setBackground(sh.getRange(1, col).getBackground());
  }
  var cell = sh.getRange(4, memberCol);
  var f = cell.getFormula();
  var mark = columnLetter_(col) + '3:' + columnLetter_(col + 15) + '3';
  if (f.indexOf(mark) < 0 || f.indexOf('HSTACK(') < 0) cell.setFormula(masterMemberNoFormula_(col));
}

/** トップの引用元リンク（H列）の末尾にエンジョイポイント付与画面を足す。既にあれば何もしない */
function ensureEnjoyPointLink_(ss) {
  var sh = ss.getSheetByName(HUB_HOME_SHEET_);
  if (!sh) return;
  var last = Math.min(sh.getLastRow(), 40);
  if (last < 2) return;
  var f = sh.getRange(1, 8, last, 1).getFormulas();
  var lastLink = 0;
  for (var r = 0; r < f.length; r++) {
    if (f[r][0].indexOf(ENJOY_POINT_URL_) >= 0) return;
    if (/HYPERLINK/i.test(f[r][0])) lastLink = r + 1;
  }
  if (!lastLink) return;
  var target = sh.getRange(lastLink + 1, 8);
  if (String(target.getDisplayValue()) !== '') return;
  var arrow = / ↗"\)$/.test(f[lastLink - 1][0]) ? ' ↗' : '';
  sh.getRange(lastLink, 8).copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  target.setFormula('=HYPERLINK("' + ENJOY_POINT_URL_ + '","' + ENJOY_POINT_TITLE_ + arrow + '")');
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
  if (!key) return;
  if (key === 'CLOSE') {
    hubCloseWork_();
    return;
  }
  if (key.indexOf('SHEET:') !== 0) return;
  var name = key.slice(6);
  var ss = sh.getParent();
  var target = ss.getSheetByName(name);
  if (!target) return;
  var opening = target.isSheetHidden();
  if (opening) {
    try { target.showSheet(); } catch (e1) {}
  } else if (hubIsWorkSheet_(name)) {
    try { target.hideSheet(); } catch (e2) {}
  }
  hubPaintTile_(sh, row, col, opening);
}

function handleHubHomeEdit_(e) {}

function restyleHubHomeLook_(sh) {
  sh.setTabColor('#111111');
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

    sh.getRange(1, 1, 40, 8)
      .setBackground(u.bg)
      .setFontColor(u.ink)
      .setFontFamily('Noto Sans JP')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setBorder(false, false, false, false, false, false)
      .setFontWeight('normal')
      .setWrap(false);

    sh.getRange(1, 1).setValue('経堂').setFontSize(16).setFontWeight('bold').setFontColor(u.ink);
    sh.getRange(1, 6).setValue('しまう').setFontSize(12).setFontWeight('bold')
      .setBackground(u.shut).setFontColor('#FFFFFF');
    hubWriteKey_(sh, 1, 6, 'CLOSE');
    sh.setRowHeight(1, 36);

    var row = 2;
    var g;
    var solid = SpreadsheetApp.BorderStyle.SOLID;
    for (g = 0; g < groups.length; g++) {
      var group = groups[g];
      var list = seen[group];
      sh.getRange(row, 1)
        .setValue(group)
        .setFontSize(12)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground(u.ink)
        .setFontColor(u.bg)
        .setBorder(true, true, true, true, false, false, u.line, solid);
      var p;
      for (p = 0; p < list.length; p++) {
        var col = 2 + p;
        var it = list[p];
        sh.getRange(row, col)
          .setValue(it.title)
          .setFontSize(13)
          .setFontWeight('bold')
          .setWrap(false)
          .setBackground(u.card)
          .setFontColor(u.ink)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setBorder(true, true, true, true, false, false, u.line, solid);
        hubWriteKey_(sh, row, col, 'SHEET:' + it.name);
      }
      sh.setRowHeight(row, 44);
      row += 1;
    }

    sh.getRange(1, 7, 40, 1).setBackground(u.bg).setBorder(false, false, false, false, false, false);
    sh.getRange(1, 8).setValue('引用元').setFontSize(11).setFontWeight('bold')
      .setBackground(u.rail).setFontColor('#FFFFFF').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, u.line, solid);
    var c;
    for (c = 0; c < links.length; c++) {
      var lr = 2 + c;
      var title = String(links[c].title).replace(/"/g, '""');
      var url = String(links[c].url).replace(/"/g, '""');
      sh.getRange(lr, 8)
        .setFormula('=HYPERLINK("' + url + '","' + title + '")')
        .setFontColor('#FFFFFF')
        .setFontSize(10)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle')
        .setWrap(true)
        .setBackground(u.rail)
        .setBorder(true, true, true, true, false, false, '#2A2A2A', solid);
    }

    try {
      var leftover = ss.getSheetByName('URL一覧');
      if (leftover && ss.getSheets().length > 1) ss.deleteSheet(leftover);
    } catch (eDel) {}

    sh.setHiddenGridlines(true);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 64);
    for (i = 2; i <= 6; i++) sh.setColumnWidth(i, 118);
    sh.setColumnWidth(7, 28);
    sh.setColumnWidth(8, 280);
    hubHideInternalCols_(sh);

    restyleHubHomeLook_(sh);
    hubCloseWork_();
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
  sh.getRange(1, 1, 1, cols)
    .setBackground(t.ink)
    .setFontColor(t.paper)
    .setFontWeight('bold')
    .setFontFamily('Meiryo');
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, cols)
      .setBackground(t.paper)
      .setFontColor(t.ink)
      .setFontFamily('Meiryo');
  }
}

function restyleKyodoMasterLook_(sh) {
  var t = dnTheme_();
  var last = Math.max(sh.getLastRow(), 24);
  var maxScan = Math.min(last, 40);
  var labels = sh.getRange(1, 1, maxScan, 1).getDisplayValues();
  var headerRow = 16;
  var i;
  for (i = 0; i < labels.length; i++) {
    if (String(labels[i][0] || '') === '項目') headerRow = i + 1;
  }
  var start = headerRow + 1;
  var bodyRows = Math.max(last - headerRow, 1);
  sh.getRange(1, 1, last, 11).setFontFamily('Meiryo').setFontColor(t.ink);
  sh.getRange(1, 1, 1, 11).setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold');
  sh.getRange(2, 1, Math.max(headerRow - 2, 1), 11).setBackground(t.paper);
  sh.getRange(headerRow, 1, 1, 10).setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold');
  sh.getRange(start, 1, bodyRows, 10).setBackground(t.ghost).setFontColor(t.ink);
  sh.getRange(start, 1, bodyRows, 1).setBackground(t.cream);
  sh.getRange(start, 6, bodyRows, 1).setBackground(t.cream);
  sh.getRange(start, 7, bodyRows, 1).setBackground(t.cream);
  try {
    sh.getRange('F7:K14').setBackground(t.ghost).setFontColor(t.ink);
    sh.getRange('F7').setBackground(t.ink).setFontColor(t.paper);
  } catch (eBrief) {}
  try {
    sh.getRange('L1:N1').setBackground(t.ink).setFontColor(t.paper);
    sh.getRange('R1:V1').setBackground(t.ink).setFontColor(t.paper);
    sh.getRange('X1:AA1').setBackground(t.ink).setFontColor(t.paper);
    sh.getRange('L2:N2').setBackground(t.cream).setFontColor(t.ink);
    sh.getRange('R2:V2').setBackground(t.cream).setFontColor(t.ink);
    sh.getRange('X2:AA2').setBackground(t.cream).setFontColor(t.ink);
    sh.getRange('L3:N3').setBackground(t.ink).setFontColor(t.paper);
    sh.getRange('R3:V3').setBackground(t.ink).setFontColor(t.paper);
    sh.getRange('X3:AA3').setBackground(t.ink).setFontColor(t.paper);
  } catch (eSide) {}
  sh.setTabColor(t.ink);
}

function restyleUnpaidTrendLook_(tr) {
  var t = dnTheme_();
  var last = Math.max(tr.getLastRow(), 3);
  tr.getRange(1, 1, last, 11).setBackground(t.paper).setFontColor(t.ink).setFontFamily('Meiryo');
  tr.getRange(1, 1, 1, 11).setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold');
  tr.getRange(2, 1, 1, 11).setBackground(t.cream).setFontWeight('bold');
  tr.getRange(1, 6, 1, 1).setBackground(t.blood).setFontColor(t.paper);
  tr.setTabColor(t.blood);
  try {
    var body = Math.max(last - 2, 1);
    var rng = tr.getRange(3, 2, body, 1);
    var forms = rng.getFormulas();
    var i;
    var changed = false;
    for (i = 0; i < forms.length; i++) {
      var s = String(forms[i][0] || '');
      if (s.indexOf('#333333') !== -1) {
        forms[i][0] = s.replace(/#333333/g, t.blood);
        changed = true;
      }
    }
    if (changed) rng.setFormulas(forms);
  } catch (eSp) {}
}

function restyleUrlIndexLook_(sh) {
  var t = dnTheme_();
  var last = Math.max(sh.getLastRow(), 1);
  var cols = Math.min(Math.max(sh.getLastColumn(), 4), 4);
  sh.getRange(1, 1, 1, cols)
    .setBackground(t.ink)
    .setFontColor(t.paper)
    .setFontFamily('Meiryo')
    .setFontWeight('bold');
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, cols)
      .setBackground(t.paper)
      .setFontColor(t.ink)
      .setFontFamily('Meiryo');
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
      styleUnpaidView_(unpaid);
      styleUnpaidDashboard_(unpaid);
      if (trend) applyUnpaidNotes_(unpaid, trend);
    }
    if (trend) restyleUnpaidTrendLook_(trend);
    var master = ss.getSheetByName('経堂マスタ');
    if (master) restyleKyodoMasterLook_(master);
    var urlSh = ss.getSheetByName('URL一覧');
    if (urlSh) restyleUrlIndexLook_(urlSh);
    var home = ss.getSheetByName(HUB_HOME_SHEET_);
    if (home) {
      restyleHubHomeLook_(home);
      hubRefreshStatus_(home, ss);
    }
    var sheets = ss.getSheets();
    var i;
    for (i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var n = sh.getName();
      if (n === HUB_HOME_SHEET_ || n === '経堂マスタ' || n === UNPAID_SHEET_ || n === UNPAID_TREND_SHEET_ || n === 'URL一覧') continue;
      if (/backup|シート\d+/.test(n)) continue;
      if (/^販促_/.test(n) || n === 'マシンレクチャー申込' || n === '入会者一覧＋自動メール管理' ||
          n === '口コミ_経堂' || n === '【経堂】会員動向' || n === '見学体験申請' ||
          /^経堂_/.test(n) || n === 'Tasks' || n === 'WorkspaceSync') {
        try { restyleHeaderBody_(sh); } catch (e1) {}
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
  var src = 'IMPORTRANGE("' + UNPAID_SOURCE_ID_ + '","\'"&B1&"\'!A1:AK")';
  return '=IFERROR(LET(d,' + src + ',' +
    'h,BYCOL(CHOOSEROWS(d,1,2,3),LAMBDA(c,TEXTJOIN("",TRUE,c))),' +
    'cat,SCAN("",CHOOSECOLS(d,1),LAMBDA(acc,x,IF(x<>"",x,acc))),' +
    'nc,MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(h,"会員名")),0),' +
    'mc,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(h,"未納対象月")),0),0),' +
    'jd,\'経堂_入会\'!A2:A,jk,\'経堂_入会\'!D2:D,' +
    'jn,ARRAYFORMULA(REGEXREPLACE(\'経堂_入会\'!B2:B&"","[\\s　]","")),' +
    'ex,MAKEARRAY(ROWS(d),4,LAMBDA(r,c,LET(nm,TRIM(IFERROR(INDEX(d,r,nc),"")&""),' +
    'IF(nm="","",IF(nm="会員名",CHOOSE(c,"入会日","入会区分","未納開始","入会から未納"),' +
    'LET(om,IF(mc=0,"",IFERROR(INDEX(d,r,mc),"")&""),' +
    'od,IFERROR(DATE(2000+VALUE(REGEXEXTRACT(om,"(\\d+)年")),VALUE(REGEXEXTRACT(om,"年(\\d+)")),1),""),' +
    'nn,REGEXREPLACE(nm,"[\\s　]",""),' +
    'lim,IF(od="",TODAY()+1,EOMONTH(od,0)+1),' +
    'jdt,IFERROR(MAX(FILTER(jd,jn=nn,jd<lim)),0),' +
    'CHOOSE(c,' +
    'IF(jdt=0,"該当なし",TEXT(jdt,"yyyy/mm/dd")),' +
    'IF(jdt=0,"",IFERROR(INDEX(FILTER(jk,jn=nn,jd=jdt),1),"")),' +
    'IF(od="","",TEXT(od,"yy年m月")),' +
    'IF(OR(jdt=0,od=""),"",((YEAR(od)-YEAR(jdt))*12+MONTH(od)-MONTH(jdt))&"ヶ月")))))))),' +
    'IF(AND(ROWS(d)=1,ISERROR(INDEX(d,1,1))),NA(),' +
    'LET(vw,MAKEARRAY(ROWS(d),COLUMNS(d),LAMBDA(r,c,LET(v,IFERROR(INDEX(d,r,c),""),' +
    'n,IF(ISNUMBER(v),v,IF(REGEXMATCH(v&"","^-?[0-9,]+(\\.[0-9]+)?$"),VALUE(v),"")),' +
    'lab,IF(AND(c>=28,r>1),IFERROR(LET(x,ARRAY_CONSTRAIN(CHOOSECOLS(d,c),r-1,1),CHOOSEROWS(FILTER(x,ISTEXT(x)),-1)),""),INDEX(h,1,c)),' +
    'IF(c=1,IF(INDEX(d,r,3)&INDEX(d,r,4)="",v,INDEX(cat,r,1)),' +
    'IF(ISLOGICAL(v),IF(v,"☑","☐"),' +
    'IF(n="",v,' +
    'IF(REGEXMATCH(lab,"率$"),TEXT(n,"0.0%"),' +
    'IF(AND(REGEXMATCH(lab,"DL|日$"),n>40000),TEXT(n,"m/d"),' +
    'IF(REGEXMATCH(lab,"額|当月分|手数料|繰越|支払|回収$"),TEXT(n,"¥#,##0"),v))))))))),' +
    'HSTACK(CHOOSECOLS(vw,SEQUENCE(1,nc)),ex,CHOOSECOLS(vw,SEQUENCE(1,COLUMNS(vw)-nc,nc+1)))))),' +
    '"「"&B1&"」のシートは元ファイルにまだありません")';
}

var UNPAID_TREND_SHEET_ = '未納管理_推移';
var UNPAID_DATA_ROW_ = 5;
var UNPAID_COLS_ = 41;

/**
 * 月タブ1枚分の集計に使う LET 変数（新旧レイアウト両対応。列は見出しの文字で探す）。
 * ok=会員番号あり・支払額>0 の行、pay=支払額、tot=総額、cat=A列区分の埋め。
 * rec=回収金額。金額欄が空でも右隣3列に「〇〇入金」があれば支払額を回収済みとみなす（旧レイアウトは金額欄を使っていない）。
 */
function unpaidStatsLet_(rangeExpr) {
  return 'd,IMPORTRANGE("' + UNPAID_SOURCE_ID_ + '",' + rangeExpr + '),' +
    'h,BYCOL(CHOOSEROWS(d,1,2,3),LAMBDA(c,TEXTJOIN("",TRUE,c))),' +
    'fc,LAMBDA(re,MATCH(TRUE,REGEXMATCH(h,re),0)),' +
    'nv,LAMBDA(col,IFERROR(VALUE(REGEXREPLACE(CHOOSECOLS(d,col)&"","[¥,\\s]","")),0)),' +
    'cat,SCAN("",CHOOSECOLS(d,1),LAMBDA(a,x,IF(x&""<>"",x&"",a))),' +
    'mem,TRIM(CHOOSECOLS(d,fc("会員番号"))&""),' +
    'pay,nv(fc("支払額")),tot,nv(fc("総額")),uc,fc("回収金額|入金金額|レジ打ち金額"),recraw,nv(uc),' +
    'rec,IF(recraw>0,recraw,IF(REGEXMATCH(CHOOSECOLS(d,uc+1)&CHOOSECOLS(d,uc+2)&CHOOSECOLS(d,uc+3)&"","入金"),pay,0)),' +
    'ok,(mem<>"")*(mem<>"会員番号")*(mem<>"合計")*(pay>0),' +
    'n,SUM(ok),sp,SUM(ok*pay),st,SUM(ok*tot),sr,SUM(ok*rec),nr,SUM(ok*(rec>0)),' +
    'kone,ok*REGEXMATCH(cat,"1ヶ月|1ヵ月|1カ月|1か月"),' +
    'ktwo,ok*REGEXMATCH(cat,"2ヶ月|2ヵ月|2カ月|2か月"),' +
    'kbad,ok*REGEXMATCH(cat,"貸倒|貸し倒"),' +
    'kjac,ok*REGEXMATCH(cat,"JACCS"),';
}

/** 未納管理 D1:M3 のダッシュボード（B1 の月） */
function unpaidDashboardFormula_() {
  return '=IFERROR(ARRAYFORMULA(LET(' + unpaidStatsLet_('"\'"&$B$1&"\'!A1:AK"') +
    'yen,LAMBDA(x,TEXT(x,"¥#,##0")),' +
    'rt,LAMBDA(k,IF(SUM(k)=0,"対象なし",IF(SUM(k*pay)=0,"対象なし",TEXT(SUM(k*rec)/SUM(k*pay),"0.0%")))),' +
    'sub,LAMBDA(k,SUM(k)&"件　"&yen(SUM(k*rec))&" / "&yen(SUM(k*pay))),' +
    'VSTACK({"未納件数","未納総額","回収額","回収率","未回収額","回収済み","1ヶ月未納 回収率","2ヶ月未納 回収率","貸倒候補 回収率","JACCS 回収率"},' +
    'HSTACK(n&"件",yen(sp),yen(sr),IF(sp=0,"-",TEXT(sr/sp,"0.0%")),yen(sp-sr),nr&"件",rt(kone),rt(ktwo),rt(kbad),rt(kjac)),' +
    'HSTACK("支払額ベース","手数料込 "&yen(st),"回収金額の合計","回収額÷未納総額","残り "&(n-nr)&"件",' +
    'IF(n=0,"",TEXT(nr/n,"0%")&"（人数）"),sub(kone),sub(ktwo),sub(kbad),sub(kjac))))),"")';
}

/** 未納管理_推移 の1行（A列の月）。数値のまま返す */
function unpaidTrendRowFormula_(row) {
  return '=IFERROR(ARRAYFORMULA(LET(' + unpaidStatsLet_('"\'"&$A' + row + '&"\'!A1:AK"') +
    'rq,LAMBDA(k,IF(SUM(k)=0,"対象なし",IF(SUM(k*pay)=0,"対象なし",SUM(k*rec)/SUM(k*pay)))),' +
    'rate,IF(sp=0,"",sr/sp),' +
    'HSTACK(n,sp,sr,rate,IF(rate="","",SPARKLINE(rate,{"charttype","bar";"max",1;"color1","#8B1216"})),sp-sr,nr,rq(kone),rq(ktwo),rq(kbad)))),"")';
}

function styleUnpaidView_(sh) {
  var maxR = sh.getMaxRows();
  var top = UNPAID_DATA_ROW_;
  var solid = SpreadsheetApp.BorderStyle.SOLID;

  var W = UNPAID_COLS_;
  var t = dnTheme_();
  var body = sh.getRange(top, 1, maxR - top + 1, W);
  body.setBackground(t.paper).setFontColor(t.ink).setFontSize(10).setFontWeight('normal')
    .setVerticalAlignment('middle').setWrap(false).setBorder(false, false, false, false, false, false);
  sh.getRange(top, 1, 2, W).setFontWeight('bold').setBackground(t.ink).setFontColor(t.paper)
    .setFontSize(9).setWrap(true).setHorizontalAlignment('center');
  sh.getRange(top + 2, 1, maxR - top - 1, W)
    .setBorder(null, null, null, null, null, true, t.line, solid);
  sh.getRange(top + 2, 2, maxR - top - 1, W - 1).setHorizontalAlignment('center');
  sh.getRange(top + 2, 4, maxR - top - 1, 1).setHorizontalAlignment('left');
  sh.getRange(top + 2, 1, maxR - top - 1, 1).setFontWeight('bold').setFontSize(9).setWrap(false)
    .setFontColor(t.ink);

  var rng = sh.getRange(top + 2, 1, maxR - top - 1, W);
  var joinRng = sh.getRange(top + 2, 4, maxR - top - 1, 5);
  var r0 = top + 2;
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
  var bands = [
    ['REGEXMATCH($A' + r0 + '&"","貸倒|貸し倒")', '#E4D3B8'],
    ['REGEXMATCH($A' + r0 + '&"","2ヶ月|2ヵ月|2カ月|3ヶ月|3ヵ月|3カ月")', t.cream],
    ['REGEXMATCH($A' + r0 + '&"","JACCS")', t.ghost],
    ['TRUE', null]
  ];
  var early = 'IFERROR(VALUE(REGEXEXTRACT($H' + r0 + '&"","^(\\d+)ヶ月$"))<=2,FALSE)';
  var none = '$E' + r0 + '="該当なし"';
  bands.forEach(function (b) {
    add('=AND(' + early + ',' + b[0] + ')', b[1], t.apple, true, joinRng);
    add('=AND(' + none + ',' + b[0] + ')', b[1], t.ash, false, joinRng);
  });
  add('=OR($A' + r0 + '="合計",$B' + r0 + '="合計")', t.cream, t.ink, true);
  bands.slice(0, 3).forEach(function (b) { add('=' + b[0], b[1], null, false); });
  sh.setConditionalFormatRules(rules);
}

function styleUnpaidDashboard_(sh) {
  var t = dnTheme_();
  var solid = SpreadsheetApp.BorderStyle.SOLID;
  sh.getRange(1, 1, 4, 37).setBackground(t.paper).setFontColor(t.ink).setVerticalAlignment('middle');
  sh.getRange('A1').setFontSize(9).setFontColor(t.ash).setHorizontalAlignment('right');
  sh.getRange('B1:C1').merge().setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, null, null, t.blood, solid);
  sh.getRange('A2:C2').merge().setFontSize(9).setHorizontalAlignment('left');
  sh.getRange('A3:C3').merge().setFontSize(9).setHorizontalAlignment('left');

  var card = sh.getRange('D1:M3');
  card.setBackground(t.cream).setHorizontalAlignment('center').setWrap(true)
    .setBorder(true, true, true, true, true, false, t.paper, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange('D1:M1').setFontSize(8).setFontColor(t.ash).setFontWeight('bold');
  sh.getRange('D2:M2').setFontSize(14).setFontColor(t.ink).setFontWeight('bold').setWrap(false);
  sh.getRange('D3:M3').setFontSize(8).setFontColor(t.ash).setFontWeight('normal');
  sh.getRange('G1:G3').setBackground(t.blood).setFontColor(t.paper);
  sh.getRange('G1').setFontColor(t.cream);
  sh.getRange('G3').setFontColor(t.cream);

  sh.setRowHeight(1, 24);
  sh.setRowHeight(2, 36);
  sh.setRowHeight(3, 32);
  sh.setRowHeight(4, 10);
}

function setupUnpaidTrend_(ss, options) {
  var tr = ss.getSheetByName(UNPAID_TREND_SHEET_);
  if (!tr) tr = ss.insertSheet(UNPAID_TREND_SHEET_);
  var months = unpaidMonthListChrono_();
  var need = months.length + 2;
  if (tr.getMaxRows() < need) tr.insertRowsAfter(tr.getMaxRows(), need - tr.getMaxRows());
  if (tr.getMaxColumns() < 11) tr.insertColumnsAfter(tr.getMaxColumns(), 11 - tr.getMaxColumns());
  tr.clear();
  tr.getRange(1, 1, 1, 11).setValues([[
    '年月', '未納件数', '未納総額', '回収額', '回収率', '全体回収率バー', '未回収額', '回収済み件数',
    '1ヶ月未納 回収率', '2ヶ月未納 回収率', '貸倒候補 回収率'
  ]]);
  tr.getRange(2, 1, 1, 11).setValues([['累計', '', '', '=SUM(D3:D)', '', '', '', '=SUM(H3:H)', '', '', '']]);
  var rows = months.map(function (m, i) {
    return [m, unpaidTrendRowFormula_(i + 3)];
  });
  tr.getRange(3, 1, rows.length, 1).setNumberFormat('@');
  tr.getRange(3, 1, rows.length, 2).setValues(rows);

  var t = dnTheme_();
  var solid = SpreadsheetApp.BorderStyle.SOLID;
  var all = tr.getRange(1, 1, tr.getMaxRows(), 11);
  all.setBackground(t.paper).setFontColor(t.ink).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
  tr.getRange(3, 1, rows.length, 11).setBorder(null, null, null, null, null, true, t.line, solid);
  tr.getRange(1, 1, 1, 11).setBackground(t.ink).setFontColor(t.paper).setFontWeight('bold')
    .setFontSize(9).setWrap(true).setHorizontalAlignment('center');
  tr.getRange(1, 6, 1, 1).setBackground(t.blood).setFontColor(t.paper);
  tr.getRange(2, 1, 1, 11).setBackground(t.cream).setFontWeight('bold');
  tr.getRange(3, 1, rows.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
  tr.getRange(2, 2, rows.length + 1, 10).setHorizontalAlignment('right');
  tr.getRange(3, 2, rows.length, 1).setNumberFormat('0"件"');
  tr.getRange(2, 3, rows.length + 1, 2).setNumberFormat('¥#,##0');
  tr.getRange(3, 5, rows.length, 1).setNumberFormat('0.0%');
  tr.getRange(3, 7, rows.length, 1).setNumberFormat('¥#,##0');
  tr.getRange(2, 8, rows.length + 1, 1).setNumberFormat('0"件"');
  tr.getRange(3, 9, rows.length, 3).setNumberFormat('0.0%');
  tr.setFrozenRows(2);
  tr.setColumnWidth(1, 90);
  tr.setColumnWidths(2, 4, 105);
  tr.setColumnWidth(6, 180);
  tr.setColumnWidths(7, 5, 105);
  tr.setRowHeight(1, 36);
  tr.setHiddenGridlines(true);
  tr.setTabColor(dnTheme_().blood);
  return tr;
}

function applyUnpaidNotes_(sh, tr) {
  sh.getRange('A1').setNote('☑☐ は表示だけです。クリックしても元ファイルは変わりません。操作は「元の未納管理ドライブ」で。');
  sh.getRange('B1').setNote('元ファイルの月タブ名です。いまの月が一番上、あとは新しい順です。');
  sh.getRange('D1').setNote('会員番号があり、支払額が1円以上の行の件数。');
  sh.getRange('E1').setNote('支払額の合計（手数料は含まない）。');
  sh.getRange('F1').setNote('回収金額の合計。空でも右隣に「〇〇入金」があれば支払額を回収済みとみなす。');
  sh.getRange('G1').setNote('回収額÷未納総額（支払額ベース）。このシート上部で血色に反転している数字。');
  sh.getRange('J1').setNote('A列区分が「1ヶ月未納」の行だけを集計した回収率。');
  sh.getRange('K1').setNote('A列区分が「2ヶ月未納」の行だけを集計した回収率。');
  sh.getRange('L1').setNote('A列に「貸倒」または「貸し倒れ」を含む行の回収率。該当者がいなければ対象なし。');
  sh.getRange('E6').setNote('経堂_入会を氏名（空白なし）で照合。未納開始月末までの最新入会。見つからない場合は該当なし（他店・改姓・2023/10以前など）。');
  sh.getRange('H6').setNote('入会日から未納開始月までの経過月数。2ヶ月以内は赤字。元の未納管理ファイルは触っていません。');
  tr.getRange('E1').setNote('その月全体の回収率（回収額÷未納総額＝支払額ベース）。右のバーと同じ数字です。');
  tr.getRange('F1').setNote('左の「回収率」を0〜100%の横棒にしたもの。貸倒や1ヶ月未納の率ではありません。');
  tr.getRange('K1').setNote('A列が貸倒／貸し倒れの行の回収率。該当者がいなければ「対象なし」。25年11月のように表記が「貸し倒れ」でも集計します。');
  tr.getRange('A2').setNote('各月の回収額の単純合計。繰越があると二重計上の可能性があります。');
}

/**
 * 未納管理ドライブ【経堂】の月タブを、B1 の年月選択で切り替えて表示するシートを作る。
 * 1〜3行目＝選択月の集計、5行目〜＝元シートの表示用コピー。別タブ「未納管理_推移」に全月の集計。
 * 元ファイルには触れない（IMPORTRANGE のみ）。
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
    sh.getRange('A4').setFormula('=HYPERLINK("https://docs.google.com/spreadsheets/d/' + UNPAID_FOLLOWUP_ID_ + '/edit","対応後☑用シートを開く ↗")');
    sh.getRange('D1').setFormula(unpaidDashboardFormula_());
    sh.getRange(UNPAID_DATA_ROW_, 1).setFormula(unpaidViewFormula_());

    var permit = '';
    try {
      var resp = UrlFetchApp.fetch(
        'https://docs.google.com/spreadsheets/d/' + ss.getId() +
          '/externaldata/addimportrangepermissions?donorDocId=' + UNPAID_SOURCE_ID_,
        { method: 'post', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
      );
      permit = String(resp.getResponseCode());
    } catch (eP) {
      permit = 'error: ' + (eP && eP.message ? eP.message : eP);
    }

    styleUnpaidView_(sh);
    styleUnpaidDashboard_(sh);
    applyUnpaidNotes_(sh, trend);
    sh.setFrozenRows(UNPAID_DATA_ROW_ + 1);
    sh.setHiddenGridlines(true);
    sh.setColumnWidth(1, 130);
    sh.setColumnWidth(2, 70);
    sh.setRowHeightsForced(UNPAID_DATA_ROW_ + 2, sh.getMaxRows() - UNPAID_DATA_ROW_ - 1, 22);
    ss.setActiveSheet(trend);
    ss.moveActiveSheet(sh.getIndex() + (trend.getIndex() < sh.getIndex() ? 0 : 1));
    ss.setActiveSheet(sh);
    sh.setColumnWidth(3, 100);
    sh.setColumnWidth(4, 130);
    sh.setColumnWidths(5, 4, 88);
    sh.setColumnWidths(9, 9, 96);
    sh.setColumnWidths(18, 13, 64);
    sh.setColumnWidths(31, UNPAID_COLS_ - 30, 84);
    sh.setFrozenColumns(4);
    sh.setTabColor(dnTheme_().blood);
    applyHubTabColors_(ss);
    SpreadsheetApp.flush();
    return {
      ok: true,
      b1: sh.getRange('B1').getDisplayValue(),
      permit: permit,
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

/**
 * 見学体験申請の K 列（入会日）。申込の前日〜180日以内で、氏名かメールが一致する最初の入会日。
 * 経堂_入会（氏名・メール）と入会者一覧（メール）の早い方。J 列はこの K 列から入会／未入会を出す。
 */
function kengakuJoinDateFormula_() {
  var d = KENGAKU_JOIN_WINDOW_DAYS_;
  return '={"入会日";ARRAYFORMULA(LET(' +
    "jd,'経堂_入会'!A2:A," +
    "jn,REGEXREPLACE('経堂_入会'!B2:B&\"\",\"[\\s　]\",\"\")," +
    "jm,LOWER(TRIM('経堂_入会'!F2:F&\"\"))," +
    "ld,'" + JOIN_LIST_SHEET_ + "'!A2:A," +
    "lm,LOWER(TRIM('" + JOIN_LIST_SHEET_ + "'!C2:C&\"\"))," +
    'MAP(A2:A,B2:B,C2:C,D2:D,LAMBDA(t,k,n,m,' +
    'IF(OR(t="",NOT(REGEXMATCH(k&"","見学|体験"))),"",IFERROR(LET(' +
    'nn,REGEXREPLACE(n&"","[\\s　]",""),mm,LOWER(TRIM(m&"")),lo,INT(t)-1,hi,t+' + d + ',' +
    'djoin,IFERROR(MIN(FILTER(jd,jd>=lo,jd<=hi,((nn<>"")*(jn=nn)+(mm<>"")*(jm=mm))>0)),0),' +
    'dlist,IF(mm="",0,IFERROR(MIN(FILTER(ld,ld>=lo,ld<=hi,lm=mm)),0)),' +
    'IF(djoin+dlist=0,"",IF(djoin=0,dlist,IF(dlist=0,djoin,MIN(djoin,dlist))))),""))))))}';
}

/** 見学体験申請の J 列（入会／未入会）。K 列に入会日があれば「入会」 */
function kengakuJoinFormula_() {
  return '={"入会";ARRAYFORMULA(IF(REGEXMATCH(B2:B&"","見学|体験"),IF(K2:K<>"","入会","未入会"),""))}';
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
        .setFontFamily('Meiryo')
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

/** 経堂マスタ R4: 今月の見学体験一覧。元フォームの希望日・時刻は文字列なので TO_TEXT で揃える */
function kengakuMasterListFormula_() {
  return "=IFERROR(QUERY({'見学体験申請'!A2:C," +
    "ARRAYFORMULA(SUBSTITUTE(TO_TEXT('見学体験申請'!H2:H),\"-\",\"/\"))," +
    "ARRAYFORMULA(TO_TEXT('見学体験申請'!I2:I))," +
    "'見学体験申請'!J2:J}," +
    "\"select Col1,Col2,Col3,Col4,Col5,Col6 where Col1 >= date '\"&TEXT($AB$5,\"yyyy-mm-dd\")&\"' " +
    "and Col1 < date '\"&TEXT(EDATE($AB$5,1),\"yyyy-mm-dd\")&\"' order by Col1 desc\",0),\"\")";
}

/** 経堂マスタ R1: 選択月（AB5）の見学体験 入会数/申込数（率） */
function kengakuMasterTitleFormula_() {
  var monthRange =
    "'見学体験申請'!A2:A,\">=\"&$AB$5,'見学体験申請'!A2:A,\"<\"&EDATE($AB$5,1)";
  var total =
    '(COUNTIFS(' + monthRange + ",'見学体験申請'!B2:B,\"見学\")+COUNTIFS(" + monthRange + ",'見学体験申請'!B2:B,\"体験\"))";
  var joined = 'COUNTIFS(' + monthRange + ",'見学体験申請'!J2:J,\"入会\")";
  return '="今月の見学体験　入会 "&' + joined + '&"/"&' + total +
    '&"（"&IFERROR(TEXT(' + joined + '/' + total + ',"0%"),"-")&"）"';
}

/**
 * 見学体験申請を元フォームからの IMPORTRANGE に切替え（バックアップ作成）、J列に入会判定、
 * 経堂マスタ R:W に今月の入会率と○列を出す。confirm=yes のときだけ書き込む。
 */
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
      master.getRange('R1').setFormula(kengakuMasterTitleFormula_());
      SpreadsheetApp.flush();
      return { ok: true, r1: master.getRange('R1').getDisplayValue() };
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
      master.getRange('R1').setFormula(kengakuMasterTitleFormula_());
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
    sh.getRange(2, 1).setFormula('=IMPORTRANGE("' + KENGAKU_SOURCE_ID_ + '","見学体験申請!A1:I")');
    sh.getRange(1, 10).setFormula(kengakuJoinFormula_());
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
    if (listSh && listSh.getLastRow() > 1) {
      listSh.getRange(2, 1, listSh.getLastRow() - 1, 3).getValues().forEach(function (r) {
        var d = toDate(r[0]);
        if (d && (r[1] || r[2])) joins.push({ at: d, name: norm(r[1]), email: norm(r[2]), kind: '' });
      });
    }

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

/**
 * 請求・回収実績：26年度未納一覧【EAST運営本部】の「経堂」行（C:AH）を月度ごとに表示・入力。
 * 入力セルの変更は installable onEdit で元シートの経堂行へ書き戻す（数式セル・書式には触れない）。
 * 元シートの値は開いた時と 5 分ごとの時間トリガーで取り込み直す。
 */
var BILL_SHEET_ = '請求・回収実績';
var BILL_OLD_SHEET_ = '未納_請求報告';
var BILL_ANALYSIS_ROW_ = 19;
var BILL_ANALYSIS_VER_ = '4';
var BILL_SOURCE_ID_ = '1qFF8HGOlSOczshMI5Vg5iTAgN_iLQ2aemJLp35V3rbA';
var BILL_STORE_ = '経堂';
var BILL_REF_TAB_ = '26年6月度';
var BILL_LAST_COL_ = 34;
var BILL_FIRST_ROW_ = 5;
var BILL_RED_ = '#f4cccc';
var BILL_CALC_ = '#eeeeee';
var BILL_NOTE_ = '白いセルに入力すると元シートの経堂行へすぐ反映　赤＝未入力　グレー＝元シートの自動計算（入力不可）';
var UNPAID_FOLLOWUP_ID_ = '1NvIIRTXC9XCAuib5USFouigkvmM8H2WDBOTWLjfN8oM';

function billMonths_() {
  var out = [];
  var m;
  for (m = 4; m <= 12; m++) out.push('26年' + m + '月度');
  for (m = 1; m <= 3; m++) out.push('27年' + m + '月度');
  return out;
}

function billColumnKeys_(head) {
  var keys = [];
  var seen = {};
  var g = '';
  var s = '';
  for (var c = 2; c < head[0].length; c++) {
    var r1 = String(head[0][c] || '').trim();
    var r2 = String(head[1][c] || '').trim();
    var h = String(head[2][c] || '').replace(/\s+/g, ' ').trim();
    if (r1) {
      g = r1.replace(/[【】]/g, '');
      s = '当月';
    }
    if (r2) s = /翌月/.test(r2) ? '翌月' : '当月';
    if (!h) continue;
    var sub = /規約退会|未納率|入金/.test(h) ? '他' : s;
    var base = g + '|' + sub + '|' + h;
    seen[base] = (seen[base] || 0) + 1;
    keys.push({ key: base + '|' + seen[base], group: g, sub: sub, header: h, col: c + 1 });
  }
  return keys;
}

function billCanon_(src) {
  var ref = src.getSheetByName(BILL_REF_TAB_);
  if (!ref) throw new Error('基準タブがありません: ' + BILL_REF_TAB_);
  return billColumnKeys_(ref.getRange(1, 1, 3, BILL_LAST_COL_).getDisplayValues());
}

function billTabInfo_(tab) {
  var lastCol = Math.max(tab.getLastColumn(), BILL_LAST_COL_);
  var keys = billColumnKeys_(tab.getRange(1, 1, 3, lastCol).getDisplayValues());
  var map = {};
  for (var i = 0; i < keys.length; i++) map[keys[i].key] = keys[i].col;
  var names = tab.getRange(1, 2, Math.max(tab.getLastRow(), 1), 1).getDisplayValues();
  var row = 0;
  for (var r = 3; r < names.length; r++) {
    if (String(names[r][0]).trim() === BILL_STORE_) {
      row = r + 1;
      break;
    }
  }
  return { map: map, row: row, lastCol: lastCol };
}

function billMergeRuns_(sh, row, arr) {
  var start = 1;
  for (var c = 2; c <= arr.length; c++) {
    if (c === arr.length || arr[c] !== arr[start]) {
      if (c - start > 1) sh.getRange(row, start + 1, 1, c - start).merge();
      start = c;
    }
  }
}

function billBuildHeader_(sh, canon) {
  var n = canon.length;
  sh.clear();
  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (eB) {}
  try {
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  } catch (eR) {}
  sh.setHiddenGridlines(true);
  var r2 = [''];
  var r3 = [''];
  var r4 = ['月度'];
  for (var i = 0; i < n; i++) {
    r2.push(canon[i].group);
    r3.push(canon[i].sub === '翌月' ? '翌月振替結果後' : (canon[i].sub === '当月' ? '当月振替結果' : ''));
    r4.push(canon[i].header);
  }
  sh.getRange(2, 1, 3, n + 1).setValues([r2, r3, r4]);
  billMergeRuns_(sh, 2, r2);
  billMergeRuns_(sh, 3, r3);
  sh.getRange(1, 1, 4, n + 1).setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(2, 1, 1, n + 1).setBackground('#000000').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(3, 1, 1, n + 1).setBackground('#424242').setFontColor('#ffffff').setHorizontalAlignment('center');
  sh.getRange(4, 1, 1, n + 1).setBackground('#212121').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.getRange(1, 1)
    .setFormula('=HYPERLINK("https://docs.google.com/spreadsheets/d/' + BILL_SOURCE_ID_ + '/edit","26年度未納一覧【EAST運営本部】を開く ↗")')
    .setFontWeight('bold').setFontColor('#000000');
  sh.getRange(1, 5, 1, 14).merge().setFontColor('#616161').setFontSize(9);
  var months = billMonths_();
  sh.getRange(BILL_FIRST_ROW_, 1, months.length, 1)
    .setValues(months.map(function (m) { return [m]; }))
    .setFontWeight('bold').setBackground('#fafafa').setFontColor('#212121');
  sh.getRange(BILL_FIRST_ROW_, 1, months.length, n + 1)
    .setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(BILL_FIRST_ROW_, 2, months.length, n).setHorizontalAlignment('right');
  sh.setRowHeight(4, 42);
  sh.setColumnWidth(1, 96);
  sh.setColumnWidths(2, n, 84);
  sh.setFrozenRows(4);
  sh.setFrozenColumns(1);
  try {
    sh.getRange(1, 1, 4, n + 1).protect().setDescription('請求報告の見出し').setWarningOnly(true);
  } catch (eP) {}
  sh.getRange(1, n + 2).setValue(canon.map(function (k) { return k.key; }).join('\t'));
  sh.hideColumns(n + 2);
}

function billPull_(ss) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, message: 'busy' };
  try {
    var src = SpreadsheetApp.openById(BILL_SOURCE_ID_);
    var canon = billCanon_(src);
    var n = canon.length;
    var sh = ss.getSheetByName(BILL_SHEET_);
    if (!sh) {
      sh = ss.getSheetByName(BILL_OLD_SHEET_);
      if (sh) sh.setName(BILL_SHEET_);
    }
    if (!sh) {
      var after = ss.getSheetByName(UNPAID_TREND_SHEET_) || ss.getSheetByName(UNPAID_SHEET_);
      sh = ss.insertSheet(BILL_SHEET_, after ? after.getIndex() : ss.getSheets().length);
    }
    var sig = canon.map(function (k) { return k.key; }).join('\t');
    if (String(sh.getRange(1, n + 2).getValue()) !== sig) billBuildHeader_(sh, canon);
    try { sh.setTabColor(hubTabColorFor_(BILL_SHEET_)); } catch (eT) {}
    billEnsureAnalysis_(sh, canon);

    var months = billMonths_();
    var vals = [], bgs = [], fcs = [], nfs = [];
    for (var m = 0; m < months.length; m++) {
      var tab = src.getSheetByName(months[m]);
      var info = tab ? billTabInfo_(tab) : null;
      var rv = null, rf = null, rn = null;
      if (info && info.row) {
        var rg = tab.getRange(info.row, 1, 1, info.lastCol);
        rv = rg.getValues()[0];
        rf = rg.getFormulas()[0];
        rn = rg.getNumberFormats()[0];
      }
      var v = [], b = [], f = [], nf = [];
      for (var i = 0; i < n; i++) {
        var col = info ? info.map[canon[i].key] : 0;
        if (!rv || !col) {
          v.push('');
          b.push('#f5f5f5');
          f.push('#bdbdbd');
          nf.push('General');
          continue;
        }
        var val = rv[col - 1];
        if (typeof val === 'string' && /^#/.test(val)) val = '';
        var isCalc = !!rf[col - 1];
        v.push(val);
        b.push(isCalc ? BILL_CALC_ : (val === '' ? BILL_RED_ : '#ffffff'));
        f.push(isCalc ? '#757575' : '#000000');
        nf.push(rn[col - 1] || 'General');
      }
      vals.push(v);
      bgs.push(b);
      fcs.push(f);
      nfs.push(nf);
    }
    var body = sh.getRange(BILL_FIRST_ROW_, 2, months.length, n);
    body.setNumberFormats(nfs);
    body.setValues(vals);
    body.setBackgrounds(bgs);
    body.setFontColors(fcs);
    sh.getRange(1, 5).setValue('取得 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm') + '（開いた時＋5分ごと）　' + BILL_NOTE_);
    return { ok: true, months: months.length, cols: n };
  } finally {
    lock.releaseLock();
  }
}

function billEnsureAnalysis_(sh, canon) {
  var top = BILL_ANALYSIS_ROW_;
  var months = billMonths_();
  var colOf = function (key) {
    for (var i = 0; i < canon.length; i++) {
      if (canon[i].key === key) return columnLetter_(i + 2);
    }
    throw new Error('列が見つかりません: ' + key);
  };
  var cCnt = colOf('合計|当月|請求件数|1');
  var cAmt = colOf('合計|当月|請求金額|1');
  var amts = sh.getRange(cAmt + BILL_FIRST_ROW_ + ':' + cAmt + (BILL_FIRST_ROW_ + months.length - 1)).getValues();
  var filled = 0;
  amts.forEach(function (v, k) { if (Number(v[0])) filled = k + 1; });
  filled = Math.max(filled, 1);
  var verCell = sh.getRange(2, canon.length + 2);
  var sig = BILL_ANALYSIS_VER_ + ':' + filled;
  if (String(verCell.getValue()) === sig && sh.getCharts().length >= 2) return;
  verCell.setValue(sig);
  var need = top + months.length + 30;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());

  var cRec = colOf('合計|当月|回収金額|1');
  var cUn = colOf('合計|当月|不納金額|1');
  var cUn2 = colOf('合計|翌月|不納金額|1');

  var heads = ['月度', '売上', '請求件数', '客単価', '回収金額', '回収率',
    '未納額', '未納率', '売上 前月比', '翌月振替の不納額（参考）'];
  var nc = heads.length;
  sh.getRange(top, 1, months.length + 2, 12).clear();
  sh.getRange(top, 1).setValue('売上・回収の分析');
  sh.getRange(top, 3, 1, 8).merge()
    .setValue('上の表（合計列）から自動計算。言葉の意味と計算方法は表の下の「用語」');
  sh.getRange(top + 1, 1, 1, nc).setValues([heads]);
  var rows = [];
  for (var i = 0; i < months.length; i++) {
    var s = BILL_FIRST_ROW_ + i;
    var r = top + 2 + i;
    var prev = r - 1;
    rows.push([
      '=$A' + s,
      '=IF(N(' + cAmt + s + ')=0,"",' + cAmt + s + ')',
      '=IF(N(' + cCnt + s + ')=0,"",' + cCnt + s + ')',
      '=IFERROR(B' + r + '/C' + r + ',"")',
      '=IF(B' + r + '="","",' + cRec + s + ')',
      '=IFERROR(E' + r + '/B' + r + ',"")',
      '=IF(B' + r + '="","",' + cUn + s + ')',
      '=IFERROR(G' + r + '/B' + r + ',"")',
      i === 0 ? '' : '=IF(OR(B' + r + '="",B' + prev + '=""),"",B' + r + '/B' + prev + '-1)',
      '=IF(OR(B' + r + '="",N(' + cUn2 + s + ')=0),"",' + cUn2 + s + ')'
    ]);
  }
  var body = sh.getRange(top + 2, 1, months.length, nc);
  body.setFormulas(rows);

  sh.getRange(top, 1, 1, nc).setFontFamily('Meiryo');
  sh.getRange(top, 1).setFontWeight('bold').setFontSize(12).setFontColor('#000000');
  sh.getRange(top, 3).setFontSize(9).setFontColor('#757575');
  sh.getRange(top + 1, 1, 1, nc).setBackground('#212121').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true).setFontFamily('Meiryo').setFontSize(10);
  body.setFontFamily('Meiryo').setFontSize(10).setBackground('#ffffff').setFontColor('#000000')
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(top + 2, 1, months.length, 1).setFontWeight('bold').setBackground('#fafafa');
  sh.getRange(top + 2, 2, months.length, 1).setNumberFormat('¥#,##0').setFontWeight('bold');
  sh.getRange(top + 2, 3, months.length, 1).setNumberFormat('#,##0');
  sh.getRange(top + 2, 4, months.length, 1).setNumberFormat('¥#,##0');
  sh.getRange(top + 2, 5, months.length, 1).setNumberFormat('¥#,##0');
  sh.getRange(top + 2, 6, months.length, 1).setNumberFormat('0.0%').setFontWeight('bold');
  sh.getRange(top + 2, 7, months.length, 1).setNumberFormat('¥#,##0');
  sh.getRange(top + 2, 8, months.length, 1).setNumberFormat('0.00%');
  sh.getRange(top + 2, 9, months.length, 1).setNumberFormat('+0.0%;[Red]-0.0%;0.0%');
  sh.getRange(top + 2, 10, months.length, 1).setNumberFormat('¥#,##0').setFontColor('#9e9e9e');
  sh.getRange(top + 2, 2, months.length, nc - 1).setHorizontalAlignment('right');
  sh.setRowHeight(top + 1, 36);

  var tableEnd = top + 1 + months.length;
  var gRow = tableEnd + 2;
  var terms = [
    ['用語', '意味と計算方法'],
    ['売上', 'その月度に請求した金額の合計（上の表「合計・当月・請求金額」）'],
    ['客単価', '売上 ÷ 請求件数。会員1人に1ヶ月で請求している平均額（月会費＋オプション込み）'],
    ['回収率', '回収金額 ÷ 売上。その月の引き落としで回収できた割合'],
    ['未納率', '未納額 ÷ 売上（元シートの「未納率」と同じ考え方）'],
    ['翌月振替の不納額', '元シート「翌月振替結果後」の不納金額をそのまま表示。当月より件数が多い月があり意味が未確認のため、計算には使っていない']
  ];
  try { sh.getRange(gRow, 1, 8, 12).breakApart(); } catch (eBr) {}
  sh.getRange(gRow, 1, 8, 12).clear();
  for (var t = 0; t < terms.length; t++) {
    sh.getRange(gRow + t, 1).setValue(terms[t][0]);
    sh.getRange(gRow + t, 2, 1, 9).merge().setValue(terms[t][1]);
  }
  sh.getRange(gRow, 1, 1, 10).setBackground('#616161').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(gRow + 1, 1, terms.length - 1, 1).setFontWeight('bold').setBackground('#fafafa');
  sh.getRange(gRow, 1, terms.length, 10).setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  var hdr = top + 1;
  var last = top + 1 + filled;
  var chartRow = gRow + terms.length + 1;
  var sales = sh.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sh.getRange('A' + hdr + ':B' + last))
    .addRange(sh.getRange('D' + hdr + ':D' + last))
    .setNumHeaders(1)
    .setOption('title', '売上（棒・左の目盛り）と客単価（赤線・右の目盛り）')
    .setOption('seriesType', 'bars')
    .setOption('series', {
      0: { type: 'bars', color: '#424242', targetAxisIndex: 0 },
      1: { type: 'line', color: '#c5221f', lineWidth: 2, pointSize: 6, targetAxisIndex: 1, dataLabel: 'value' }
    })
    .setOption('vAxes', {
      0: { format: 'short', viewWindow: { min: 0 }, gridlines: { count: 4 } },
      1: { format: '¥#,##0', viewWindow: { min: 0, max: 12000 }, gridlines: { count: 4 } }
    })
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 620)
    .setOption('height', 300)
    .setPosition(chartRow, 1, 0, 0)
    .build();
  sh.insertChart(sales);
  var rate = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange('A' + hdr + ':A' + last))
    .addRange(sh.getRange('F' + hdr + ':F' + last))
    .setNumHeaders(1)
    .setOption('title', '回収率（回収金額 ÷ 売上）')
    .setOption('series', {
      0: { color: '#000000', lineWidth: 3, pointSize: 6, dataLabel: 'value' }
    })
    .setOption('vAxis', { format: '0%', viewWindow: { min: 0.8, max: 1 }, gridlines: { count: 5 } })
    .setOption('legend', { position: 'none' })
    .setOption('width', 620)
    .setOption('height', 300)
    .setPosition(chartRow, 8, 0, 0)
    .build();
  sh.insertChart(rate);
}

function billParseInput_(raw) {
  if (raw === '' || raw == null) return '';
  if (typeof raw !== 'string') return raw;
  var s = raw.replace(/[,，¥￥円\s]/g, '');
  if (s === '') return '';
  if (/^-?\d+(\.\d+)?%$/.test(s)) return Number(s.slice(0, -1)) / 100;
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  return raw;
}

function billingOnEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== BILL_SHEET_) return;
  var months = billMonths_();
  var r0 = e.range.getRow();
  var c0 = e.range.getColumn();
  var nr = e.range.getNumRows();
  var nc = e.range.getNumColumns();
  var lastRow = BILL_FIRST_ROW_ + months.length - 1;
  if (r0 + nr - 1 < BILL_FIRST_ROW_ || r0 > lastRow || c0 + nc - 1 < 2) return;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var src = SpreadsheetApp.openById(BILL_SOURCE_ID_);
    var canon = billCanon_(src);
    var vals = e.range.getValues();
    var cache = {};
    for (var i = 0; i < nr; i++) {
      var r = r0 + i;
      if (r < BILL_FIRST_ROW_ || r > lastRow) continue;
      var month = months[r - BILL_FIRST_ROW_];
      if (!cache[month]) {
        var tab = src.getSheetByName(month);
        cache[month] = tab ? { tab: tab, info: billTabInfo_(tab) } : { tab: null };
      }
      var t = cache[month];
      for (var j = 0; j < nc; j++) {
        var c = c0 + j;
        if (c < 2 || c > canon.length + 1) continue;
        var cell = sh.getRange(r, c);
        var col = t.tab && t.info.row ? t.info.map[canon[c - 2].key] : 0;
        if (!col) {
          cell.setValue('');
          continue;
        }
        var target = t.tab.getRange(t.info.row, col);
        if (target.getFormula()) {
          cell.setValue(target.getValue());
          cell.setBackground(BILL_CALC_).setFontColor('#757575');
          continue;
        }
        var nv = billParseInput_(vals[i][j]);
        if (String(target.getValue()) !== String(nv)) target.setValue(nv);
        cell.setBackground(nv === '' ? BILL_RED_ : '#ffffff').setFontColor('#000000');
      }
    }
    sh.getRange(1, 5).setValue('反映 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm') + '　' + BILL_NOTE_);
  } finally {
    lock.releaseLock();
  }
}

function billingPullTriggered() {
  var ss = openWorkspaceSpreadsheet_();
  try { billPull_(ss); } catch (e) { console.error(e); }
  try { memberAnalysisIfChanged_(ss); } catch (e2) { console.error(e2); }
  try { ensureEnjoyPointLink_(ss); } catch (e3) { console.error(e3); }
  try { ensureMasterMemberNo_(ss); } catch (e5) { console.error(e5); }
  try { ensureMasterApplyCheckmarks_(ss); } catch (e4) { console.error(e4); }
  try { syncJoinBreakdownIfChanged_(ss); } catch (e6) { console.error(e6); }
}

/**
 * 会員分析：累計入会データ／累計退会データ（手動貼り付け）から継続率・生涯売上・属性別を集計。
 * 元データの行数が変わった時だけ作り直す。客単価は 請求・回収実績 の分析表を参照。
 */
var MEMBER_SHEET_ = '会員分析';
var MEMBER_JOIN_SRC_ = '累計入会データ';
var MEMBER_LEAVE_SRC_ = '累計退会データ';
var MEMBER_ANALYSIS_VER_ = '7';

function memberAnalysisIfChanged_(ss) {
  var j = ss.getSheetByName(MEMBER_JOIN_SRC_);
  var l = ss.getSheetByName(MEMBER_LEAVE_SRC_);
  if (!j || !l) return;
  var sig = MEMBER_ANALYSIS_VER_ + ':' + j.getLastRow() + ':' + l.getLastRow() + ':' + (ss.getSheetByName(MEMBER_SHEET_) ? 1 : 0);
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('MEMBER_ANALYSIS_SIG') === sig) return;
  buildMemberAnalysis_(ss);
  props.setProperty('MEMBER_ANALYSIS_SIG', sig);
}

function memberYm_(s) {
  var m = String(s || '').match(/(\d{4})\D+(\d{1,2})/);
  return m ? Number(m[1]) * 12 + Number(m[2]) - 1 : null;
}

function memberYmLabel_(ym) {
  return Math.floor(ym / 12) + '/' + ('0' + (ym % 12 + 1)).slice(-2);
}

function memberNorm_(s) {
  return String(s || '').normalize('NFKC').replace(/[\s・]/g, '')
    .replace(/[\u3041-\u3096]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) + 0x60); });
}

function memberPhone_(s) {
  var d = String(s || '').replace(/\D/g, '');
  if (d.length === 10 && d.charAt(0) !== '0') d = '0' + d;
  if (d.length < 10 || d.length > 11 || d.charAt(0) !== '0') return '';
  return d.slice(-10);
}

function memberYmd_(s) {
  var m = String(s || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) } : null;
}

function memberMail_(s) {
  var v = String(s || '').trim().toLowerCase();
  return /@/.test(v) ? v : '';
}

/**
 * 日報の数字マスだけ更新する。C/E/G の「移籍」「/復会」「/紹介」は触らない。
 * 14行＝当月、10行＝日報のその日（B1の月が過去ならその月末日）。
 * 移籍・紹介＝Workspace の販促シートを数える（日報の IMPORTRANGE は許可ができず 0 になる）。
 * 復会＝経堂_入会の過去一致。月シート 2609 の INDIRECT は使わない。
 */
var JOIN_BREAKDOWN_VER_ = '3';
var JOIN_BREAKDOWN_NOTE_ = 'Workspace自動（14行=当月、10行=その日。移籍・紹介＝追加販促の申請、復会＝経堂_入会の過去一致）';

function parseNippoYm_(ss) {
  var tz = 'Asia/Tokyo';
  var now = new Date();
  var todayY = Number(Utilities.formatDate(now, tz, 'yyyy'));
  var todayM = Number(Utilities.formatDate(now, tz, 'M'));
  var todayD = Number(Utilities.formatDate(now, tz, 'd'));
  var y = todayY;
  var m = todayM;
  try {
    var nip = SpreadsheetApp.openById(RECEPTION_SOURCE_ID_).getSheetByName('日報');
    var b1 = String(nip ? nip.getRange('B1').getDisplayValue() : '').replace(/\D/g, '');
    if (b1.length === 4) {
      y = 2000 + Number(b1.slice(0, 2));
      m = Number(b1.slice(2, 4));
    } else if (b1.length === 6) {
      y = Number(b1.slice(0, 4));
      m = Number(b1.slice(4, 6));
    }
  } catch (eYm) {}
  if (!y || !m) { y = todayY; m = todayM; }
  return { y: y, m: m, d: todayD, ym: y * 12 + m - 1, todayY: todayY, todayM: todayM, sameMonth: y === todayY && m === todayM };
}

function joinBreakdownSig_(ss) {
  var ym = parseNippoYm_(ss);
  var joinSh = ss.getSheetByName('経堂_入会');
  var moveSh = ss.getSheetByName('販促_乗り換え');
  var introSh = ss.getSheetByName('販促_紹介・ペア入会');
  var nip = '';
  try {
    var nsh = SpreadsheetApp.openById(RECEPTION_SOURCE_ID_).getSheetByName('日報');
    nip = [
      nsh.getRange('B1').getDisplayValue(),
      nsh.getRange('D10').getDisplayValue(), nsh.getRange('F10').getDisplayValue(), nsh.getRange('H10').getDisplayValue(),
      nsh.getRange('D14').getDisplayValue(), nsh.getRange('F14').getDisplayValue(), nsh.getRange('H14').getDisplayValue(),
      nsh.getRange('D21').getFormula() || nsh.getRange('D21').getDisplayValue()
    ].join(':');
  } catch (eSig) {}
  return JOIN_BREAKDOWN_VER_ + ':' + ym.y + '-' + ym.m + '-' + ym.d + ':' +
    (joinSh ? joinSh.getLastRow() : 0) + ':' +
    (moveSh ? moveSh.getLastRow() : 0) + ':' +
    (introSh ? introSh.getLastRow() : 0) + ':' + nip;
}

function syncJoinBreakdownIfChanged_(ss) {
  try {
    ss = ss || openWorkspaceSpreadsheet_();
    var sig = joinBreakdownSig_(ss);
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('JOIN_BREAKDOWN_SIG') === sig) return { ok: true, skipped: true };
    var r = syncJoinBreakdown_(ss);
    try { props.setProperty('JOIN_BREAKDOWN_SIG', joinBreakdownSig_(ss)); } catch (eP) {}
    return r;
  } catch (err) {
    console.error(err);
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function syncJoinBreakdown_(ss) {
  ss = ss || openWorkspaceSpreadsheet_();
  var counts = countMoveIntroRejoin_(ss);
  var nippo = writeNippoMoveIntroRejoin_(counts);
  var op = { ok: false };
  try { op = ensureNippoOpByB1_(false); } catch (eOp) { op = { ok: false, message: String(eOp && eOp.message ? eOp.message : eOp) }; }
  try { noteMemberAnalysisBreakdown_(ss); } catch (eN) {}
  return { ok: true, counts: counts, nippo: nippo, op: op };
}

function refreshJoinBreakdownFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = syncJoinBreakdown_(ss);
  var c = (r && r.counts) || {};
  ss.toast('当月 移籍' + c.move + ' 復会' + c.rejoin + ' 紹介' + c.intro +
    '／当日 移籍' + c.moveToday + ' 復会' + c.rejoinToday + ' 紹介' + c.introToday, '移籍・復会・紹介', 8);
}

function reportDay_(ym) {
  var monthStart = new Date(ym.y, ym.m - 1, 1);
  var monthEnd = new Date(ym.y, ym.m, 0);
  var today = new Date(ym.todayY, ym.todayM - 1, ym.d);
  if (today < monthStart) return monthStart;
  if (today > monthEnd) return monthEnd;
  return today;
}

function countPromo_(sheet, introOnly, y, m, day) {
  var out = { month: 0, day: 0 };
  if (!sheet || sheet.getLastRow() < 2) return out;
  var values = sheet.getDataRange().getDisplayValues();
  var h = -1;
  for (var r = 0; r < Math.min(values.length, 5); r++) {
    if (values[r].some(function (v) { return /申請日時|タイムスタンプ/.test(v); })) { h = r; break; }
  }
  if (h < 0) return out;
  var head = values[h];
  var dateCol = -1;
  head.forEach(function (v, c) { if (dateCol < 0 && /申請日時|タイムスタンプ/.test(v)) dateCol = c; });
  var hasHi = head.some(function (v) { return /被紹介/.test(v); });
  var nameCols = [];
  var skipCols = {};
  head.forEach(function (v, c) {
    if (hasHi && /紹介者/.test(v) && !/被紹介/.test(v)) { skipCols[c] = 1; return; }
    if (introOnly) {
      if (hasHi ? /被紹介.*名前/.test(v) : /名前|登録名|氏名/.test(v)) nameCols.push(c);
    } else if (/名前|登録名|氏名/.test(v) && !/紹介者/.test(v)) {
      nameCols.push(c);
    }
  });
  var seenM = {};
  var seenD = {};
  for (var i = h + 1; i < values.length; i++) {
    var row = values[i];
    var ymd = memberYmd_(row[dateCol]);
    if (!ymd || ymd.y !== y || ymd.m !== m) continue;
    var key = '';
    for (var n = 0; n < nameCols.length && !key; n++) {
      var raw = String(row[nameCols[n]] || '').trim();
      if (!raw || memberPhone_(raw)) continue;
      var nm = memberNorm_(raw);
      if (nm.length >= 2) key = nm;
    }
    if (!key && introOnly) {
      for (var c2 = 0; c2 < row.length && !key; c2++) {
        if (skipCols[c2] || c2 === dateCol) continue;
        var raw2 = String(row[c2] || '').trim();
        if (!raw2 || memberPhone_(raw2) || /JOYFIT|経堂|該当/.test(raw2)) continue;
        var nm2 = memberNorm_(raw2);
        if (nm2.length >= 2 && !/\d{4,}/.test(nm2)) key = nm2;
      }
    }
    if (!key) key = 'row' + i;
    if (!seenM[key]) { seenM[key] = 1; out.month++; }
    if (day && ymd.d === day && !seenD[key]) { seenD[key] = 1; out.day++; }
  }
  return out;
}

function countMoveIntroRejoin_(ss) {
  var ym = parseNippoYm_(ss);
  var day = reportDay_(ym);
  var move = countPromo_(ss.getSheetByName('販促_乗り換え'), false, ym.y, ym.m, day.getDate());
  var intro = countPromo_(ss.getSheetByName('販促_紹介・ペア入会'), true, ym.y, ym.m, day.getDate());
  var rejoin = countRejoinOnly_(ss, ym, day);
  return {
    y: ym.y, m: ym.m, day: day.getDate(),
    move: move.month, intro: intro.month, rejoin: rejoin.month,
    moveToday: move.day, introToday: intro.day, rejoinToday: rejoin.day
  };
}

function countRejoinOnly_(ss, ym, day) {
  ym = ym || parseNippoYm_(ss);
  day = day || reportDay_(ym);
  var histMail = {};
  var histName = {};
  var thisMonth = [];
  var joinSh = ss.getSheetByName('経堂_入会');
  if (joinSh && joinSh.getLastRow() > 1) {
    var jv = joinSh.getRange(2, 1, joinSh.getLastRow() - 1, 6).getDisplayValues();
    for (var i = 0; i < jv.length; i++) {
      var name = memberNorm_(jv[i][1]);
      var mail = memberMail_(jv[i][5]);
      var stamp = String(jv[i][0] || '');
      var rowYm = memberYm_(jv[i][2]);
      if (rowYm == null) rowYm = memberYm_(stamp);
      if (rowYm == null) continue;
      if (rowYm < ym.ym) {
        if (mail) histMail[mail] = 1;
        if (name && name.length >= 2) histName[name] = 1;
      } else if (rowYm === ym.ym) {
        var ymd = memberYmd_(stamp);
        thisMonth.push({
          name: name,
          mail: mail,
          day: ymd ? ymd.d : 0
        });
      }
    }
  }
  var seen = {};
  var month = 0;
  var dayN = 0;
  var reportD = day.getDate();
  for (var t = 0; t < thisMonth.length; t++) {
    var p = thisMonth[t];
    var key = p.mail || p.name;
    if (!key || seen[key]) continue;
    var hit = (p.mail && histMail[p.mail]) || (p.name && p.name.length >= 2 && histName[p.name]);
    if (!hit) continue;
    seen[key] = 1;
    month++;
    if (p.day === reportD) dayN++;
  }
  return { month: month, day: dayN };
}

function writeNippoMoveIntroRejoin_(counts) {
  var sh = SpreadsheetApp.openById(RECEPTION_SOURCE_ID_).getSheetByName('日報');
  if (!sh) return { ok: false, message: '日報なし' };
  var setNum = function (a1, val) {
    var cell = sh.getRange(a1);
    cell.setValue(Number(val));
    cell.setNote(JOIN_BREAKDOWN_NOTE_);
    return { cell: a1, to: val };
  };
  return {
    ok: true,
    D14: setNum('D14', counts.move),
    F14: setNum('F14', counts.rejoin),
    H14: setNum('H14', counts.intro),
    D10: setNum('D10', counts.moveToday),
    F10: setNum('F10', counts.rejoinToday),
    H10: setNum('H10', counts.introToday)
  };
}

function noteMemberAnalysisBreakdown_(ss) {
  var sh = ss.getSheetByName(MEMBER_SHEET_);
  if (!sh) return;
  var b5 = sh.getRange('B5');
  if (String(b5.getNote() || '') !== JOIN_BREAKDOWN_NOTE_) b5.setNote(JOIN_BREAKDOWN_NOTE_);
}

/**
 * 日報の契約・解約は OP集計の「対象月」ではなく、日報 B1（2610 など）の月を数える。
 * OP集計のプルダウンは「ログに1件でもある月」しか出ないので、月初の10月は選べない。
 * 同じブックの OP集計ログ（I=日時 K=区分 M=集計名）を COUNTIFS する。月初列は触らない。
 */
var NIPPO_OP_NOTE_ = '日報B1の月を自動集計（OP集計の対象月は見ない。B1を2611にすれば11月）';

function nippoOpCountFormula_(row, kind) {
  return '=IFERROR(LET('
    + 'bcode,REGEXREPLACE($B$1&"","[^0-9]",""),'
    + 'yy,IF(LEN(bcode)=6,VALUE(LEFT(bcode,4)),2000+VALUE(LEFT(bcode,2))),'
    + 'mm,VALUE(RIGHT(bcode,2)),'
    + 'COUNTIFS(\'OP集計\'!$M:$M,$B' + row
    + ',\'OP集計\'!$K:$K,"' + kind + '"'
    + ',\'OP集計\'!$I:$I,">="&DATE(yy,mm,1)'
    + ',\'OP集計\'!$I:$I,"<"&DATE(yy,mm+1,1))),0)';
}

function nippoOpLastRow_(sh) {
  var last = 21;
  var names = sh.getRange('B21:B40').getDisplayValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0] || '').trim()) last = 21 + i;
  }
  return last;
}

function nippoOpFormulaOk_(formula) {
  var f = String(formula || '');
  return /COUNTIFS/i.test(f) && /OP集計/.test(f) && /\$B\$1/.test(f);
}

function ensureNippoOpByB1_(force) {
  var book = SpreadsheetApp.openById(RECEPTION_SOURCE_ID_);
  var sh = book.getSheetByName('日報');
  var op = book.getSheetByName('OP集計');
  if (!sh) return { ok: false, message: '日報なし' };
  if (!op) return { ok: false, message: 'OP集計なし' };
  if (!force && nippoOpFormulaOk_(sh.getRange('D21').getFormula()) && nippoOpFormulaOk_(sh.getRange('E21').getFormula())) {
    var keep = sh.getRange('D22:F22').getDisplayValues()[0];
    return {
      ok: true,
      skipped: true,
      b1: sh.getRange('B1').getDisplayValue(),
      vipContract: keep[0],
      vipCancel: keep[1],
      vipDelta: keep[2]
    };
  }
  var last = nippoOpLastRow_(sh);
  var n = last - 20;
  var dForms = [];
  var eForms = [];
  var fForms = [];
  for (var row = 21; row <= last; row++) {
    dForms.push([nippoOpCountFormula_(row, '利用開始*')]);
    eForms.push([nippoOpCountFormula_(row, '利用停止')]);
    fForms.push(['=IFERROR(D' + row + '-E' + row + ',0)']);
  }
  sh.getRange(21, 4, n, 1).setFormulas(dForms).setNote(NIPPO_OP_NOTE_);
  sh.getRange(21, 5, n, 1).setFormulas(eForms).setNote(NIPPO_OP_NOTE_);
  sh.getRange(21, 6, n, 1).setFormulas(fForms);
  sh.getRange(21, 4, n, 3).setNumberFormat('0');
  SpreadsheetApp.flush();
  var vip = sh.getRange('D22:F22').getDisplayValues()[0];
  return {
    ok: true,
    skipped: false,
    b1: sh.getRange('B1').getDisplayValue(),
    lastRow: last,
    vipContract: vip[0],
    vipCancel: vip[1],
    vipDelta: vip[2]
  };
}

function refreshNippoOpFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = ensureNippoOpByB1_(true);
  ss.toast(r.ok
    ? ('B1=' + r.b1 + '　VIP 契約' + r.vipContract + ' 解約' + r.vipCancel)
    : String(r.message), '日報オプション', 10);
  return r;
}

function 日報オプションをB1連動() {
  var book = SpreadsheetApp.openById('14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w');
  var sh = book.getSheetByName('日報');
  var opSh = book.getSheetByName('OP集計');
  if (!sh) throw new Error('日報なし');
  if (!opSh) throw new Error('OP集計なし');
  var last = 21;
  var names = sh.getRange('B21:B40').getDisplayValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0] || '').trim()) last = 21 + i;
  }
  var dForms = [];
  var eForms = [];
  var fForms = [];
  for (var row = 21; row <= last; row++) {
    var body = 'LET(bcode,REGEXREPLACE($B$1&"","[^0-9]",""),yy,IF(LEN(bcode)=6,VALUE(LEFT(bcode,4)),2000+VALUE(LEFT(bcode,2))),mm,VALUE(RIGHT(bcode,2)),COUNTIFS(\'OP集計\'!$M:$M,$B'
      + row + ',\'OP集計\'!$K:$K,"KIND",\'OP集計\'!$I:$I,">="&DATE(yy,mm,1),\'OP集計\'!$I:$I,"<"&DATE(yy,mm+1,1)))';
    dForms.push(['=IFERROR(' + body.replace('KIND', '利用開始*') + ',0)']);
    eForms.push(['=IFERROR(' + body.replace('KIND', '利用停止') + ',0)']);
    fForms.push(['=IFERROR(D' + row + '-E' + row + ',0)']);
  }
  var n = last - 20;
  var note = '日報B1の月を自動集計（OP集計の対象月は見ない。B1を2611にすれば11月）';
  sh.getRange(21, 4, n, 1).setFormulas(dForms).setNote(note);
  sh.getRange(21, 5, n, 1).setFormulas(eForms).setNote(note);
  sh.getRange(21, 6, n, 1).setFormulas(fForms);
  sh.getRange(21, 4, n, 3).setNumberFormat('0');
  SpreadsheetApp.flush();
  var b1 = sh.getRange('B1').getDisplayValue();
  var vip = sh.getRange('D22:F22').getDisplayValues()[0];
  var r = { ok: true, b1: b1, vipContract: vip[0], vipCancel: vip[1], vipDelta: vip[2] };
  try { installJoinBreakdownTrigger_(); } catch (eT) {}
  try {
    SpreadsheetApp.getUi().alert(
      '日報B1=' + b1 + ' の月を契約・解約が数えます。\n'
      + '安心サポートVIP　契約' + vip[0] + '　解約' + vip[1] + '　増減' + vip[2] + '\n'
      + 'OP集計の対象月は触らなくて大丈夫です。'
    );
  } catch (eUi) {}
  return r;
}

function 移籍復会紹介を反映() {
  var r = syncJoinBreakdown_(openWorkspaceSpreadsheet_());
  try { installJoinBreakdownTrigger_(); } catch (eT) {}
  var c = (r && r.counts) || {};
  var op = (r && r.op) || {};
  try {
    SpreadsheetApp.getUi().alert(
      '当月　移籍' + c.move + ' 復会' + c.rejoin + ' 紹介' + c.intro + '\n' +
      '当日　移籍' + c.moveToday + ' 復会' + c.rejoinToday + ' 紹介' + c.introToday + '\n' +
      'オプション　B1=' + (op.b1 || '') + '　VIP契約' + (op.vipContract || (op.skipped ? '関数済' : ''))
    );
  } catch (eUi) {}
  return r;
}

function installJoinBreakdownTrigger_() {
  var fn = 'syncJoinBreakdownIfChanged_';
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === fn) return;
  }
  var ss = openWorkspaceSpreadsheet_();
  ScriptApp.newTrigger(fn).timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger(fn).forSpreadsheet(ss).onOpen().create();
}


/**
 * 販促／申込シートの申請者を累計入会データの会員と照合する。
 * 電話番号 → 氏名（漢字・カナ）の順。申請月の前2ヶ月〜後3ヶ月に入会した人だけを候補にする。
 * 「紹介」系は被紹介者の列だけを見る。
 */
function memberMatchPromo_(sheet, members, idx) {
  var values = sheet.getDataRange().getDisplayValues();
  var h = -1;
  for (var r = 0; r < Math.min(values.length, 5); r++) {
    if (values[r].some(function (v) { return /申請日時|タイムスタンプ/.test(v); })) { h = r; break; }
  }
  var out = { apps: 0, list: [], minYm: null };
  if (h < 0) return out;
  var head = values[h];
  var dateCol = -1;
  head.forEach(function (v, c) { if (dateCol < 0 && /申請日時|タイムスタンプ/.test(v)) dateCol = c; });
  var hasHi = head.some(function (v) { return /被紹介/.test(v); });
  var nameCols = [];
  var skipCols = {};
  head.forEach(function (v, c) {
    if (hasHi && /紹介者/.test(v) && !/被紹介/.test(v)) { skipCols[c] = 1; return; }
    if (hasHi ? /被紹介.*名前/.test(v) : /名前|登録名|氏名/.test(v)) nameCols.push(c);
  });
  var seen = {};
  for (var i = h + 1; i < values.length; i++) {
    var row = values[i];
    var appYm = memberYm_(row[dateCol]);
    if (appYm == null) continue;
    out.apps++;
    if (out.minYm == null || appYm < out.minYm) out.minYm = appYm;
    var inWin = function (m) { return m.ym >= appYm - 2 && m.ym <= appYm + 3; };
    var hit = null;
    for (var c = 0; c < row.length && !hit; c++) {
      if (skipCols[c] || c === dateCol) continue;
      var tel = memberPhone_(row[c]);
      if (tel && idx.tel[tel]) hit = idx.tel[tel].filter(inWin)[0] || null;
    }
    for (var n = 0; n < nameCols.length && !hit; n++) {
      var key = memberNorm_(row[nameCols[n]]);
      if (key.length < 2) continue;
      hit = ((idx.name[key] || []).concat(idx.kana[key] || [])).filter(inWin)[0] || null;
      if (!hit) {
        for (var k = 0; k < members.length; k++) {
          var m = members[k];
          if (!inWin(m)) continue;
          if ((m.name.length >= 3 && key.indexOf(m.name) >= 0) || (m.kana.length >= 4 && key.indexOf(m.kana) >= 0)) { hit = m; break; }
        }
      }
    }
    if (hit) {
      var hk = members.indexOf(hit);
      if (!seen[hk]) { seen[hk] = 1; out.list.push(hit); }
    }
  }
  return out;
}

function memberAgeBand_(a) {
  a = Number(a);
  if (!a) return '不明';
  if (a < 20) return '10代';
  if (a < 60) return Math.floor(a / 10) * 10 + '代';
  return '60代以上';
}

function buildMemberAnalysis_(ss) {
  var js = ss.getSheetByName(MEMBER_JOIN_SRC_);
  var ls = ss.getSheetByName(MEMBER_LEAVE_SRC_);
  if (!js || !ls) return { ok: false, message: '累計データのシートがありません' };
  var id = function (v) { return String(v || '').trim().replace(/^0+/, ''); };

  var lv = ls.getRange(2, 6, Math.max(ls.getLastRow() - 1, 1), 13).getDisplayValues();
  var leaves = {};
  var leaveCount = 0;
  var tenureSum = 0;
  for (var i = 0; i < lv.length; i++) {
    if (!/^\d+$/.test(String(lv[i][0]).trim())) continue;
    var t = Number(lv[i][10]) || 0;
    var lym = memberYm_(lv[i][12]);
    leaves[id(lv[i][0])] = { tenure: t, ym: lym };
    leaveCount++;
    tenureSum += t;
  }

  var jv = js.getRange(2, 1, Math.max(js.getLastRow() - 1, 1), 15).getDisplayValues();
  var members = [];
  var nowYm = 0;
  for (var k = 0; k < jv.length; k++) {
    if (!/^\d+$/.test(String(jv[k][5]).trim())) continue;
    var plan = String(jv[k][2]).trim();
    var ym = memberYm_(jv[k][13]);
    if (ym == null) continue;
    var reqYm = memberYm_(jv[k][12]);
    if (reqYm != null && reqYm > nowYm) nowYm = reqYm;
    var lf = leaves[id(jv[k][5])] || null;
    members.push({
      ym: ym, gender: String(jv[k][8]).trim() || '不明', age: memberAgeBand_(jv[k][9]),
      plan: plan || '通常（契約名称なし）', left: lf,
      name: memberNorm_(jv[k][6]), kana: memberNorm_(jv[k][7]), tel: memberPhone_(jv[k][11])
    });
  }
  var recentLeaves = 0;
  Object.keys(leaves).forEach(function (key) {
    var lym2 = leaves[key].ym;
    if (lym2 != null && lym2 <= nowYm && lym2 > nowYm - 12) recentLeaves++;
  });
  var active = members.filter(function (m) { return !m.left; }).length;

  var stats = function (list) {
    var o = { n: list.length, r3: '', r6: '', r12: '', r24: '', now: '', active: 0, avgTen: '' };
    [3, 6, 12, 24].forEach(function (mo) {
      var el = list.filter(function (m) { return nowYm - m.ym >= mo; });
      if (el.length < 10) return;
      var kept = el.filter(function (m) { return !m.left || m.left.tenure > mo; }).length;
      o['r' + mo] = kept / el.length;
    });
    o.active = list.filter(function (m) { return !m.left; }).length;
    if (list.length) o.now = o.active / list.length;
    var lefts = list.filter(function (m) { return m.left; });
    if (lefts.length) o.avgTen = lefts.reduce(function (a, m) { return a + m.left.tenure; }, 0) / lefts.length;
    return o;
  };
  var group = function (list, keyFn) {
    var g = {};
    list.forEach(function (m) {
      var key = keyFn(m);
      (g[key] = g[key] || []).push(m);
    });
    return g;
  };

  var sh = ss.getSheetByName(MEMBER_SHEET_);
  if (!sh) {
    var after = ss.getSheetByName(BILL_SHEET_);
    sh = ss.insertSheet(MEMBER_SHEET_, after ? after.getIndex() : ss.getSheets().length);
  }
  sh.clear();
  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  sh.clearConditionalFormatRules();
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (eB) {}
  try { sh.expandAllRowGroups(); } catch (eG) {}
  for (var gd = 0; gd < 3; gd++) {
    try { sh.getRange(1, 1, sh.getMaxRows(), 1).shiftRowGroupDepth(-1); } catch (eG2) { break; }
  }
  if (sh.getMaxRows() < 220) sh.insertRowsAfter(sh.getMaxRows(), 220 - sh.getMaxRows());
  sh.setHiddenGridlines(true);
  sh.setTabColor('#212121');

  var cols = ['区分', '人数', '3ヶ月継続率', '6ヶ月継続率', '12ヶ月継続率', '24ヶ月継続率', '今も在籍', '平均在籍（退会者）', '生涯売上（退会者）'];
  var nc = cols.length;
  var pct = function (x) { return Math.round(x * 1000) / 10 + '%'; };
  var grey = '#757575';

  sh.getRange('A1').setValue('会員分析').setFontSize(16).setFontWeight('bold');
  sh.getRange('E1').setValue('更新 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm') +
    '　累計データは ' + memberYmLabel_(nowYm) + ' 入会分まで').setFontSize(9).setFontColor(grey);

  // 今月の会員数（受付状況表・日報を IMPORTRANGE で直接参照）
  var nip = function (a1) { return 'IMPORTRANGE("' + RECEPTION_SOURCE_ID_ + '","日報!' + a1 + '")'; };
  var mf = function (a1) { return '"男 "&' + nip('F' + a1) + '&" / 女 "&' + nip('H' + a1); };
  sh.getRange('A2').setFormula('=IFERROR("今月（"&' + nip('B1') + '&"）の会員数","今月の会員数（読み込み中）")')
    .setFontWeight('bold').setFontSize(12);
  sh.getRange('E2').setValue('受付状況表「日報」からリアルタイム').setFontSize(9).setFontColor(grey);
  sh.getRange(3, 1, 1, 8).setValues([['月初会員', '当月入会', '当月退会', '純増', '月末安定会員', '翌月月初会員', '休会', '今月の退会率']]);
  sh.getRange(4, 1, 1, 8).setFormulas([[
    '=IFERROR(' + nip('C12') + ',"")', '=IFERROR(' + nip('C13') + ',"")', '=IFERROR(' + nip('C15') + ',"")',
    '=IFERROR(B4-C4,"")', '=IFERROR(' + nip('C16') + ',"")', '=IFERROR(' + nip('C17') + ',"")',
    '=IFERROR(' + nip('C18') + ',"")', '=IFERROR(C4/A4,"")'
  ]]);
  sh.getRange(5, 1, 1, 8).setFormulas([[
    '=IFERROR(' + mf(12) + ',"")',
    '=IFERROR("移籍"&' + nip('D14') + '&" 復会"&' + nip('F14') + '&" 紹介"&' + nip('H14') + ',"")',
    '', '="入会−退会"',
    '=IFERROR(' + mf(16) + ',"")', '=IFERROR(' + mf(17) + ',"")',
    '', '="退会÷月初"'
  ]]);
  sh.getRange(3, 1, 1, 8).setBackground('#212121').setFontColor('#ffffff').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(4, 1, 1, 8).setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#ffffff');
  sh.getRange(5, 1, 1, 8).setFontSize(9).setFontColor(grey).setHorizontalAlignment('center');
  sh.getRange(3, 1, 3, 8).setBorder(true, true, true, true, true, false, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('A4:G4').setNumberFormat('#,##0"人"');
  sh.getRange('D4').setNumberFormat('+#,##0"人";-#,##0"人"');
  sh.getRange('H4').setNumberFormat('0.00%');
  sh.getRange('A4').setFontColor('#c5221f');

  // 1人あたりのお金と続き方（累計データ）
  sh.getRange('A7').setValue('1人あたりのお金と続き方').setFontWeight('bold').setFontSize(12);
  sh.getRange('E7').setValue('累計入会・退会データから').setFontSize(9).setFontColor(grey);
  sh.getRange(8, 1, 1, 7).setValues([['累計入会', '累計退会', '月の退会率', '平均在籍', '客単価', '生涯売上（実績）', '生涯売上（理論）']]);
  sh.getRange(9, 1, 1, 7).setValues([[members.length, leaveCount, '', leaveCount ? tenureSum / leaveCount : '', '', '', '']]);
  sh.getRange('C9').setFormula('=IFERROR(' + (recentLeaves / 12) + '/A4,"")');
  sh.getRange('E9').setFormula("=IFERROR(AVERAGE('" + BILL_SHEET_ + "'!D" + (BILL_ANALYSIS_ROW_ + 2) + ':D' + (BILL_ANALYSIS_ROW_ + 13) + '),"")');
  sh.getRange('F9').setFormula('=IFERROR(D9*E9,"")');
  sh.getRange('G9').setFormula('=IFERROR(E9/C9,"")');
  sh.getRange(10, 1, 1, 7).setValues([['', '', '直近12ヶ月の平均÷月初会員', '辞めた人の平均', '請求金額÷件数', '平均在籍×客単価', '客単価÷退会率']]);
  sh.getRange(8, 1, 1, 7).setBackground('#616161').setFontColor('#ffffff').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(9, 1, 1, 7).setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#fafafa');
  sh.getRange(10, 1, 1, 7).setFontSize(8).setFontColor(grey).setHorizontalAlignment('center').setWrap(true);
  sh.getRange(8, 1, 3, 7).setBorder(true, true, true, true, true, false, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('A9:B9').setNumberFormat('#,##0"人"');
  sh.getRange('C9').setNumberFormat('0.00%');
  sh.getRange('D9').setNumberFormat('0.0"ヶ月"');
  sh.getRange('E9:G9').setNumberFormat('¥#,##0');
  sh.getRange('F9:G9').setBackground('#000000').setFontColor('#ffffff');

  var gid = sh.getSheetId();
  var menuRow = 12;
  var sections = [];
  var row = 15;
  var rateRanges = [];
  var setTip = function (r, tip, isFormula) {
    var rg = sh.getRange(r, 2, 1, nc - 1).merge();
    if (isFormula) rg.setFormula(tip); else rg.setValue(tip);
    rg.setFontColor('#c5221f').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left');
  };
  var openSection = function (title, menu, note) {
    sh.getRange(row, 1, 1, nc).setBackground('#f5f5f5')
      .setBorder(true, false, false, false, false, false, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, 1).setValue('▶ ' + title).setFontWeight('bold').setFontSize(12);
    sections.push({ title: title, menu: menu, row: row });
    sh.setRowHeight(row, 32);
    row++;
    var bodyStart = row;
    if (note) {
      sh.getRange(row, 1, 1, nc).merge().setValue(note).setFontSize(9).setFontColor(grey).setWrap(true);
      sh.setRowHeight(row, note.length > 70 ? 34 : 22);
      row++;
    }
    return bodyStart;
  };
  var closeSection = function (bodyStart) {
    row++;
    sections[sections.length - 1].body = [bodyStart, row - 1];
  };
  var writeTable = function (title, menu, note, labels, lists, withLtv, tipFn) {
    var titleRow = row;
    var bodyStart = openSection(title, menu, note);
    sh.getRange(row, 1, 1, nc).setValues([cols])
      .setBackground('#212121').setFontColor('#ffffff').setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(row, 34);
    row++;
    var start = row;
    var st = lists.map(function (l) { return stats(l); });
    if (labels.length) {
      var out = labels.map(function (lab, i) {
        var s = st[i];
        return [lab, s.n, s.r3, s.r6, s.r12, s.r24, s.active, s.avgTen, ''];
      });
      sh.getRange(start, 1, out.length, nc).setValues(out)
        .setFontSize(10).setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
      sh.getRange(start, 1, out.length, 1).setFontWeight('bold').setBackground('#fafafa').setWrap(true);
      sh.getRange(start, 2, out.length, 1).setNumberFormat('#,##0');
      sh.getRange(start, 3, out.length, 4).setNumberFormat('0.0%');
      sh.getRange(start, 7, out.length, 1).setNumberFormat('#,##0');
      sh.getRange(start, 8, out.length, 1).setNumberFormat('0.0"ヶ月"');
      if (withLtv) {
        var f = [];
        for (var q = 0; q < out.length; q++) f.push(['=IF(H' + (start + q) + '="","",H' + (start + q) + '*$E$9)']);
        sh.getRange(start, 9, out.length, 1).setFormulas(f).setNumberFormat('¥#,##0');
      }
      rateRanges.push(sh.getRange(start, 3, out.length, 4));
      row = start + out.length;
    }
    closeSection(bodyStart);
    var tip = tipFn ? tipFn(st) : '';
    if (tip) setTip(titleRow, tip, false);
    return { start: start, st: st };
  };
  var extremes = function (labels, st, key, word) {
    var arr = labels.map(function (l, i) { return { l: l, v: st[i][key] }; }).filter(function (x) { return x.v !== ''; });
    if (arr.length < 2) return '';
    arr.sort(function (a, b) { return b.v - a.v; });
    var last = arr[arr.length - 1];
    return word + '　続きやすい：' + arr[0].l + ' ' + pct(arr[0].v) + '　／　続きにくい：' + last.l + ' ' + pct(last.v);
  };

  // お金への影響（請求 × 累計）
  var baseR3 = stats(members.filter(function (m) { return m.ym >= memberYm_('2017/01') && m.ym < memberYm_('2023/01'); })).r3;
  var recentR3 = stats(members.filter(function (m) { return m.ym > nowYm - 15 && m.ym <= nowYm - 3; })).r3;
  var joins12 = members.filter(function (m) { return m.ym > nowYm - 12; }).length;
  var billB = "'" + BILL_SHEET_ + "'!B" + (BILL_ANALYSIS_ROW_ + 2) + ':B' + (BILL_ANALYSIS_ROW_ + 13);
  var moneyTitle = row;
  var moneyBody = openSection('お金への影響', 'お金');
  sh.getRange(row, 1, 1, 3).setValues([['項目', '金額', '計算']]);
  sh.getRange(row, 3, 1, nc - 2).merge();
  sh.getRange(row, 1, 1, nc).setBackground('#212121').setFontColor('#ffffff').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  row++;
  var money = [
    ['月の売上（直近月）', '=IFERROR(ARRAYFORMULA(LOOKUP(2,1/(' + billB + '<>""),' + billB + ')),"")', '請求・回収実績の直近月の請求金額'],
    ['退会で毎月消える売上', '=IFERROR(' + (recentLeaves / 12) + '*$E$9,"")',
      '月平均の退会 ' + (Math.round(recentLeaves / 12 * 10) / 10) + '人 × 客単価'],
    ['入会1人の生涯売上', '=G9', '客単価 ÷ 月の退会率（1人獲得にかけてよい販促費の上限の目安）'],
    ['3ヶ月継続率を戻すと（年）',
      (baseR3 !== '' && recentR3 !== '') ? '=IFERROR(' + joins12 + '*' + Math.max(baseR3 - recentR3, 0) + '*$G$9,"")' : '',
      (recentR3 !== '' ? pct(recentR3) : '-') + '（直近）を ' + (baseR3 !== '' ? pct(baseR3) : '-') + '（2017〜22年）に戻した場合の試算']
  ];
  var money0 = row;
  for (var mi = 0; mi < money.length; mi++) {
    sh.getRange(row, 1).setValue(money[mi][0]);
    if (money[mi][1]) sh.getRange(row, 2).setFormula(money[mi][1]);
    sh.getRange(row, 3, 1, nc - 2).merge().setValue(money[mi][2]);
    row++;
  }
  sh.getRange(money0, 1, money.length, nc).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(money0, 1, money.length, 1).setFontWeight('bold').setBackground('#fafafa');
  sh.getRange(money0, 2, money.length, 1).setNumberFormat('¥#,##0').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(money0, 3, money.length, 1).setFontSize(9).setFontColor(grey);
  sh.getRange(row - 1, 2).setFontColor('#c5221f');
  closeSection(moneyBody);
  setTip(moneyTitle, '=IFERROR("退会で毎月 約"&TEXT(B' + (money0 + 1) + '/10000,"#,##0")&"万円の売上が消えている → その分を新規入会で埋めている","")', true);

  // 販促・レクチャー × 累計
  var idx = { tel: {}, name: {}, kana: {} };
  members.forEach(function (m) {
    if (m.tel) (idx.tel[m.tel] = idx.tel[m.tel] || []).push(m);
    if (m.name) (idx.name[m.name] = idx.name[m.name] || []).push(m);
    if (m.kana) (idx.kana[m.kana] = idx.kana[m.kana] || []).push(m);
  });
  var promos = [
    ['乗り換え', '販促_乗り換え'],
    ['紹介・ペア入会', '販促_紹介・ペア入会'],
    ['学割', '学割'],
    ['ラグビー割', '販促_ラグビー割'],
    ['6ヶ月継続特典', '販促_6ヶ月継続'],
    ['マシンレクチャー', 'マシンレクチャー申込']
  ];
  var promoLabels = [];
  var promoLists = [];
  var promoCounts = [];
  var used = {};
  var minYm = null;
  promos.forEach(function (p) {
    var psh = ss.getSheetByName(p[1]);
    if (!psh) return;
    var res = memberMatchPromo_(psh, members, idx);
    res.list.forEach(function (m) { used[members.indexOf(m)] = 1; });
    if (res.minYm != null && (minYm == null || res.minYm < minYm)) minYm = res.minYm;
    promoLabels.push(p[0]);
    promoLists.push(res.list);
    promoCounts.push(p[0] + ' ' + res.apps + '→' + res.list.length);
  });
  if (minYm != null) {
    var base = members.filter(function (m, ix) { return m.ym >= minYm && !used[ix]; });
    promoLabels.push('販促なし（比較用）');
    promoLists.push(base);
  }
  var promoRes = writeTable('販促別の続き方', '販促別',
    '申請シートの人を入会データと電話番号・氏名で照合（申請→照合：' + promoCounts.join('、') + '）。比較用＝同じ時期に販促なしで入会した人',
    promoLabels, promoLists, true, function (st) {
      var arr = promoLabels.map(function (l, i) { return { l: l.replace('（比較用）', ''), v: st[i].r6 }; })
        .filter(function (x) { return x.v !== '' && x.l !== '6ヶ月継続特典'; });
      if (!arr.length) return '';
      arr.sort(function (a, b) { return b.v - a.v; });
      return '6ヶ月後も続いている割合　' + arr.map(function (x) { return x.l + ' ' + pct(x.v); }).join('　');
    });
  if (promoLists.length) {
    sh.getRange(promoRes.start + promoLists.length - 1, 1, 1, nc).setFontColor(grey).setFontStyle('italic');
  }

  // 季節 × 継続率
  var seasonPool = members.filter(function (m) { return m.ym >= memberYm_('2023/01'); });
  var bySeason = group(seasonPool, function (m) { return m.ym % 12 + 1; });
  var seasonKeys = [];
  for (var sm = 1; sm <= 12; sm++) seasonKeys.push(sm);
  var seasonLabels = seasonKeys.map(function (x) { return x + '月入会'; });
  var thisMonth = Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M'));
  writeTable('入会した月で比べる', '入会した月', '2023年以降の入会。季節で続きやすさが違うかを見る',
    seasonLabels, seasonKeys.map(function (x) { return bySeason[x] || []; }), true, function (st) {
      var avg12 = stats(seasonPool).r12;
      var s12 = st[thisMonth - 1].r12;
      if (s12 === '' || avg12 === '') return extremes(seasonLabels, st, 'r12', '12ヶ月継続率');
      var diff = (s12 - avg12) * 100;
      return '今月（' + thisMonth + '月）入会は' + (diff < -3 ? '続きにくい' : diff > 3 ? '続きやすい' : '平均並み') +
        '（12ヶ月継続 ' + pct(s12) + '／平均 ' + pct(avg12) + '）' +
        (diff < -3 ? ' → 入会後1〜3ヶ月のフォローを厚く' : diff > 3 ? ' → 入会を取りにいく好機' : '');
    });

  var byYear = group(members, function (m) { return Math.floor(m.ym / 12); });
  var years = Object.keys(byYear).sort();
  writeTable('入会した年で比べる', '入会した年', '',
    years.map(function (y) { return y + '年'; }), years.map(function (y) { return byYear[y]; }), true, function () {
      if (baseR3 === '' || recentR3 === '') return '';
      return '3ヶ月後も続いている割合　昔（2017〜22年）' + pct(baseR3) + ' → 直近 ' + pct(recentR3) +
        (recentR3 < baseR3 ? '（下がっている）' : '');
    });

  var monthsList = [];
  for (var mm = nowYm - 35; mm <= nowYm; mm++) monthsList.push(mm);
  var byMonth = group(members, function (m) { return m.ym; });
  var monthLabels = monthsList.map(memberYmLabel_);
  var monthRes = writeTable('月ごとの推移', '月ごとの推移', '直近36ヶ月の入会月ごと。右のグラフと同じ数字',
    monthLabels, monthsList.map(function (x) { return byMonth[x] || []; }), false, function (st) {
      for (var i = st.length - 1; i >= 0; i--) {
        if (st[i].r3 === '') continue;
        var vals = st.filter(function (s) { return s.r3 !== ''; }).map(function (s) { return s.r3; });
        var avg = vals.reduce(function (a, b) { return a + b; }, 0) / vals.length;
        return '最新（' + monthLabels[i] + '入会）の3ヶ月継続率 ' + pct(st[i].r3) + '（36ヶ月の平均 ' + pct(avg) + '）';
      }
      return '';
    });
  var monthStart = monthRes.start;
  var monthEnd = monthStart + monthsList.length - 1;

  var recent = members.filter(function (m) { return m.ym >= memberYm_('2023/01'); });
  var byAge = group(recent, function (m) { return m.age; });
  var ages = Object.keys(byAge).sort();
  writeTable('年代で比べる', '年代', '2023年以降の入会。年齢は現在',
    ages, ages.map(function (a) { return byAge[a]; }), true, function (st) { return extremes(ages, st, 'r6', '6ヶ月継続率'); });
  var byGender = group(recent, function (m) { return m.gender; });
  var genders = Object.keys(byGender).sort();
  writeTable('性別で比べる', '性別', '2023年以降の入会',
    genders, genders.map(function (g) { return byGender[g]; }), true, function (st) {
      return '6ヶ月継続率　' + genders.map(function (g, i) { return st[i].r6 === '' ? '' : g + ' ' + pct(st[i].r6); })
        .filter(String).join('　／　');
    });
  var byPlan = group(recent, function (m) { return m.plan; });
  var plans = Object.keys(byPlan).sort(function (a, b) { return byPlan[b].length - byPlan[a].length; }).slice(0, 12);
  writeTable('契約プランで比べる', 'プラン', '2023年以降の入会・人数の多い12プラン',
    plans, plans.map(function (p) { return byPlan[p]; }), true, function (st) { return extremes(plans, st, 'r6', '6ヶ月継続率'); });

  var rules = rateRanges.map(function (rg) {
    return SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#f4cccc', SpreadsheetApp.InterpolationType.NUMBER, '0.3')
      .setGradientMaxpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setRanges([rg])
      .build();
  });
  sh.setConditionalFormatRules(rules);

  sh.getRange(1, 1, row + 2, nc).setFontFamily('Meiryo');
  sh.getRange(3, 1, row, nc).setVerticalAlignment('middle');
  sh.setColumnWidth(1, 210);
  sh.setColumnWidths(2, nc - 1, 108);
  sh.setRowHeight(3, 28);
  sh.setRowHeight(4, 40);
  sh.setRowHeight(5, 22);
  sh.setRowHeight(8, 28);
  sh.setRowHeight(9, 36);
  sh.setRowHeight(10, 28);
  sh.setFrozenRows(5);

  sh.getRange(menuRow, 1).setValue('▼ 見たい項目をタップ（飛んだ先の行の左にある「＋」で中身が開きます）')
    .setFontWeight('bold').setFontSize(10);
  var links = sections.map(function (s) {
    return '=HYPERLINK("#gid=' + gid + '&range=A' + s.row + '","' + s.menu + '")';
  });
  sh.getRange(menuRow + 1, 1, 1, links.length).setFormulas([links])
    .setFontSize(10).setFontWeight('bold').setFontColor('#c5221f').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#ffffff').setBorder(true, true, true, true, true, false, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(menuRow + 1, 34);
  sh.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  sections.forEach(function (s, i) {
    if (!s.body) return;
    sh.getRange(s.body[0], 1, s.body[1] - s.body[0] + 1, 1).shiftRowGroupDepth(1);
    if (i > 0) sh.getRowGroup(s.body[0], 1).collapse();
  });

  var chart = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .addRange(sh.getRange(monthStart - 1, 1, monthsList.length + 1, 1))
    .addRange(sh.getRange(monthStart - 1, 3, monthsList.length + 1, 1))
    .addRange(sh.getRange(monthStart - 1, 4, monthsList.length + 1, 1))
    .addRange(sh.getRange(monthStart - 1, 5, monthsList.length + 1, 1))
    .setNumHeaders(1)
    .setOption('title', '入会月ごとの継続率（赤＝3ヶ月後・濃い灰＝6ヶ月後・薄い灰＝12ヶ月後）')
    .setOption('series', {
      0: { color: '#c5221f', lineWidth: 3, pointSize: 4 },
      1: { color: '#616161', lineWidth: 2 },
      2: { color: '#bdbdbd', lineWidth: 2 }
    })
    .setOption('vAxis', { format: '0%', viewWindow: { min: 0, max: 1 } })
    .setOption('legend', { position: 'none' })
    .setOption('width', 620)
    .setOption('height', 320)
    .setPosition(menuRow, nc + 2, 0, 0)
    .build();
  sh.insertChart(chart);
  return { ok: true, members: members.length, leaves: leaveCount, active: active, monthEnd: monthEnd };
}

function refreshMemberAnalysisFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = buildMemberAnalysis_(ss);
  try { syncJoinBreakdown_(ss); } catch (eBd) {}
  ss.toast(r.ok ? '会員分析を更新しました' : String(r.message), '会員分析', 5);
}

function linkUnpaidFollowup_(ss) {
  var sh = ss.getSheetByName(UNPAID_SHEET_);
  if (!sh) return;
  var a4 = sh.getRange('A4');
  if (/HYPERLINK/i.test(String(a4.getFormula() || ''))) return;
  sh.getRange('A3').copyTo(a4, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  a4.setFormula('=HYPERLINK("https://docs.google.com/spreadsheets/d/' + UNPAID_FOLLOWUP_ID_ + '/edit","対応後☑用シートを開く ↗")');
}

function setupBillingLinkFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  linkUnpaidFollowup_(ss);
  var r = billPull_(ss);
  var handlers = { billingOnEdit: 1, billingPullTriggered: 1 };
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (handlers[triggers[i].getHandlerFunction()]) ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('billingOnEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('billingPullTriggered').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('billingPullTriggered').forSpreadsheet(ss).onOpen().create();
  var sh = ss.getSheetByName(BILL_SHEET_);
  if (sh) ss.setActiveSheet(sh);
  ss.toast(r.ok ? '請求報告の自動連携を有効にしました' : String(r.message), '請求・回収実績', 8);
}

function refreshBillingFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = billPull_(ss);
  ss.toast(r.ok ? '元シートから取り込みました' : String(r.message), '請求・回収実績', 5);
}

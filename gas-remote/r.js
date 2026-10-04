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
  ui.createMenu('ファイルUP')
  .addItem('月初作業ファイルUP（会員数・オプション・年齢男女のExcel）', 'openGessho3Files')
  .addItem('数値更新ファイルUP（入会・退会手続き一覧表、退会アンケートCSV）', 'openNumbersUpload')
  .addToUi();
  try { hideGesshoSideSheets_(); } catch (eHideSide) { Logger.log(eHideSide); }
  try { ensureNippoMirror_(); } catch (eMirror) { Logger.log(eMirror); }
  try { installTokureiKaiinSheet(); } catch (eTokurei) { Logger.log(eTokurei); }
  try { fixOctoberNippo_(); } catch (eOct) { Logger.log(eOct); }
  try {
    if (PropertiesService.getDocumentProperties().getProperty('CUMULATIVE_LOOK_V') !== 'v1') {
      formatCumulativeSheets_();
    }
  } catch (eLook) { Logger.log(eLook); }
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

/** 追加販促ブック（スタッフ入力）→ 見た目整形 + Workspace へ IMPORTRANGE */
var PROMO_SOURCE_ID_ = '1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E';

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
var NIPPO_MIRROR_SHEET_ = '日報';

/** 受付状況表の日報を Workspace「日報」へ IMPORTRANGE。月初３ファイルの反映がここに出る。 */
function ensureNippoMirror_() {
  var ss = null;
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId() === WS_CONFIG.SPREADSHEET_ID) ss = active;
  } catch (eActive) {}
  if (!ss) ss = openWorkspaceSpreadsheet_();
  var sh = ss.getSheetByName(NIPPO_MIRROR_SHEET_);
  if (!sh) {
    var master = ss.getSheetByName('経堂マスタ');
    var index = master ? master.getIndex() + 1 : 1;
    sh = ss.insertSheet(NIPPO_MIRROR_SHEET_, index);
    sh.setTabColor('#111111');
  }
  var formula = '=IMPORTRANGE("' + RECEPTION_SOURCE_ID_ + '","\'日報\'!A1:I40")';
  var current = String(sh.getRange('A1').getFormula() || '');
  if (current.indexOf('日報') < 0 || current.indexOf(RECEPTION_SOURCE_ID_) < 0) {
    sh.getRange('A1').setFormula(formula);
  }
  sh.getRange('K1').clearContent();
  sh.setColumnWidth(11, 120);
  formatNippoMirrorLook_(sh);
  return { ok: true, gid: sh.getSheetId(), formula: sh.getRange('A1').getFormula() };
}

/** 写しの列幅・見出し・枠。IMPORTRANGEの値は触らない。 */
function formatNippoMirrorLook_(sh) {
  var widths = [32, 200, 72, 56, 72, 64, 64, 64, 52];
  for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);
  sh.setHiddenGridlines(true);
  var area = sh.getRange(1, 1, 40, 9);
  area.setFontFamily('Meiryo');
  area.setFontSize(10);
  area.setFontColor('#111111');
  area.setFontWeight('normal');
  area.setVerticalAlignment('middle');
  area.setWrap(false);
  for (var r = 1; r <= 40; r++) sh.setRowHeight(r, r === 20 ? 26 : 22);

  sh.getRange('B1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sh.getRange(5, 1, 3, 8).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange(9, 2, 10, 1).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange(9, 3, 10, 7).setHorizontalAlignment('center');

  var header = sh.getRange(20, 2, 1, 5);
  header.setFontWeight('bold');
  header.setFontColor('#ffffff');
  header.setBackground('#111111');
  header.setHorizontalAlignment('center');

  sh.getRange(20, 2, 17, 5).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(21, 2, 16, 1).setFontWeight('bold').setHorizontalAlignment('left').setBackground('#ffffff');
  sh.getRange(21, 3, 16, 4).setHorizontalAlignment('center').setNumberFormat('0;-0;0').setBackground('#ffffff');
  sh.getRange(21, 3, 16, 1).setBackground('#f2f2f2');

  sh.getRange('A42').setValue('メニュー「月初３ファイル」で3つのExcelを入れると、この表も同じ数字になります。');
  sh.getRange('A42').setFontColor('#666666').setFontSize(9).setFontWeight('normal');
}

function ensureNippoMirror() {
  var result = ensureNippoMirror_();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + ss.getId() +
        '/externaldata/addimportrangepermissions?donorDocId=' + RECEPTION_SOURCE_ID_,
      {
        method: 'post',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      }
    );
    ss.toast('日報の写しを出しました', '月初３ファイル', 5);
  } catch (eGrant) {
    Logger.log(eGrant);
  }
  return result;
}
var JOIN_LIST_SHEET_ = '入会者一覧＋自動メール管理';

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

/** マシンレクチャー／入会者一覧 → Workspace（同名シート・内容そのまま・IMPORTRANGE） */
var MACHINE_SOURCE_ID_ = '1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8';

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
    if (api === 'exportPdf' && String(e.parameter.token || '') === RECEPTION_REFRESH_TOKEN_) {
      var ssx = openWorkspaceSpreadsheet_();
      var shx = ssx.getSheetByName(String(e.parameter.name || ''));
      if (!shx) return jsonOutput_({ ok: false, message: 'no sheet' });
      var url = 'https://docs.google.com/spreadsheets/d/' + ssx.getId() + '/export?format=pdf&gid=' + shx.getSheetId() +
        '&size=A3&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=false&fzr=false' +
        (e.parameter.range ? '&range=' + encodeURIComponent(String(e.parameter.range)) : '');
      var resx = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (resx.getResponseCode() !== 200) return jsonOutput_({ ok: false, code: resx.getResponseCode(), text: resx.getContentText().slice(0, 200) });
      return jsonOutput_({ ok: true, pdf: Utilities.base64Encode(resx.getBlob().getBytes()) });
    }
    if (api === 'ensureNippoMirror') {
      return jsonOutput_(ensureNippoMirror_());
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
    if (api === 'importLeaveSurvey') {
      if (String(body.token || '') !== RECEPTION_REFRESH_TOKEN_) return unauthorized_();
      return jsonOutput_(importLeaveSurvey_(body.rows || []));
    }
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

var HUB_KEY_BASE_ = 50;
var HUB_KEY_SPAN_ = 8;

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

/** 経堂マスタ B2（年月）は今日の月に自動で切り替える。AG5 以降の月計算はすべて B2 から作られる */
var MASTER_MONTH_FORMULA_ = '=TEXT(TODAY(),"yyyy年m月")';

function ensureMasterMonthAuto_(ss) {
  var sh = ss.getSheetByName('経堂マスタ');
  if (!sh) return;
  var b2 = sh.getRange('B2');
  if (b2.getFormula() === MASTER_MONTH_FORMULA_) return;
  b2.clearDataValidations().setFormula(MASTER_MONTH_FORMULA_);
}

/** 5ヶ月データベース「入会実績」の下に、各月1日〜今日と同じ日までの入会数（経堂_入会）を出す */
var MASTER_SAMEDAY_LABEL_ = '入会 同日比較';

function masterSameDayFormula_(offset) {
  return '=LET(m,EDATE($AG$5,' + offset + '),d,MIN(DAY(TODAY()),DAY(EOMONTH(m,0))),' +
    'IFERROR(COUNTIFS(\'経堂_入会\'!$C:$C,TEXT(m,"yyyy年m月"),\'経堂_入会\'!$A:$A,"<"&(m+d)),))';
}

function ensureMasterSameDayRow_(ss) {
  var sh = ss.getSheetByName('経堂マスタ');
  if (!sh) return;
  var labels = sh.getRange('A1:A80').getDisplayValues();
  var base = -1;
  for (var i = 0; i < labels.length; i++) {
    if (labels[i][0] === '入会実績') { base = i + 1; break; }
  }
  if (base < 0) return;
  var row = base + 1;
  if (labels[base][0] !== MASTER_SAMEDAY_LABEL_) {
    sh.getRange(row, 1, 1, 15).insertCells(SpreadsheetApp.Dimension.ROWS);
    sh.getRange(base, 1, 1, 15).copyTo(sh.getRange(row, 1, 1, 15), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sh.getRange(row, 7, 1, 9).clearContent();
  }
  var f = [];
  for (var k = -4; k <= 0; k++) f.push(masterSameDayFormula_(k));
  if (sh.getRange(row, 2).getFormula() === f[0] && sh.getRange(row, 10).getFormula()) return;
  sh.getRange(row, 1).setValue(MASTER_SAMEDAY_LABEL_)
    .setNote('各月の1日〜今日と同じ日までの入会数（経堂_入会：入会メールの自動カウント）。月途中でも前月までと同じ条件で比べられます');
  sh.getRange(row, 2, 1, 5).setFormulas([f]);
  sh.getRange(row, 10).setFormula('=IF(AND(ISNUMBER(F' + row + '),ISNUMBER(E' + row + ')),F' + row + '-E' + row + ',)');
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

/** トップの作り（版を上げると次の自動実行で作り直す） */
var TOP_LAYOUT_VERSION_ = 'simple-v7';
var TOP_SRC_RECEPTION_ = 'https://docs.google.com/spreadsheets/d/14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w/edit';
var TOP_SRC_UNPAID_ = 'https://docs.google.com/spreadsheets/d/10vpQRDfTdwx_Wb7JaSm3lZCkTk8msLyf8ggAHhI1shI/edit';
var TOP_SRC_TRIAL_ = 'https://docs.google.com/spreadsheets/d/1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY/edit';
var TOP_SRC_SCHOOL_ = 'https://docs.google.com/spreadsheets/d/1mmG_xM1WoWFKgpmOl5obsKXo_0GjWLnAnLY9hTGanVg/edit#gid=667254510';
var TOP_SRC_REVIEW_ = 'https://docs.google.com/spreadsheets/d/13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM/edit';
var TOP_SRC_LECTURE_ = 'https://docs.google.com/spreadsheets/d/1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8/edit';
var TOP_SRC_PROMO_ = 'https://docs.google.com/spreadsheets/d/1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E/edit';

/** sub: 文字列なら元シートURL、{text} ならただの説明 */
function topSections_() {
  return [
    { label: '数字', items: [
      { name: '経堂マスタ', sub: TOP_SRC_RECEPTION_ },
      { name: '日報', sub: TOP_SRC_RECEPTION_ },
      { name: '【経堂】会員動向', sub: TOP_SRC_RECEPTION_ }
    ] },
    { label: '分析', items: [
      { name: '会員分析', sub: { text: '今月の数字（累計データから自動）' } },
      { name: '入会・退会分析', sub: { text: '推移・退会の内訳・継続率（会議用）' } },
      { name: '有料オプション検討', sub: { text: '目標と月初・入会特典の比較' } }
    ] },
    { label: '未納', items: [
      { name: '未納管理', sub: TOP_SRC_UNPAID_ },
      { name: '未納管理_推移', sub: TOP_SRC_UNPAID_ },
      { name: '請求・回収実績', sub: { text: '5分ごとに自動取得' } },
      { name: '規約退会リスト', sub: { text: '赤い名前から自動作成' } }
    ] },
    { label: '現場', items: [
      { name: '見学体験申請', sub: TOP_SRC_TRIAL_ },
      { name: '学割', sub: TOP_SRC_SCHOOL_ },
      { name: '口コミ管理', sub: TOP_SRC_REVIEW_ },
      { name: 'マシンレクチャー申込', sub: TOP_SRC_LECTURE_ },
      { name: '入会者一覧＋自動メール管理', sub: TOP_SRC_LECTURE_ }
    ] },
    { label: '販促', items: [
      { name: '販促_乗り換え', sub: TOP_SRC_PROMO_ },
      { name: '販促_紹介・ペア入会', sub: TOP_SRC_PROMO_ },
      { name: '販促_ラグビー割', sub: TOP_SRC_PROMO_ },
      { name: '販促_6ヶ月継続', sub: TOP_SRC_PROMO_ }
    ] },
    { label: 'データ', items: [
      { name: '累計入会データ', sub: { text: '入会の元データ' } },
      { name: '累計退会データ', sub: { text: '退会の元データ' } }
    ] }
  ];
}

function topExternalLinks_() {
  return [
    ['経堂　受付状況表', TOP_SRC_RECEPTION_],
    ['26年度未納管理ドライブ【経堂】', TOP_SRC_UNPAID_],
    ['未納マニュアル', 'https://docs.google.com/presentation/d/1bcjGsBP2ZsjN8hzZ5Z7C-5ZXDxcuWCFgvqniVyn8gcA/edit'],
    ['経堂　見学・体験フォーム', TOP_SRC_TRIAL_],
    ['学校関係者割フォーム', 'https://docs.google.com/spreadsheets/d/1mmG_xM1WoWFKgpmOl5obsKXo_0GjWLnAnLY9hTGanVg/edit'],
    ['EAST口コミ回答者', TOP_SRC_REVIEW_],
    ['20分マシンレクチャー・自動送信メール', TOP_SRC_LECTURE_],
    ['JOYFIT24経堂追加販促', TOP_SRC_PROMO_],
    ['口コミ付与アプリ', 'https://script.google.com/a/macros/okamoto-group.co.jp/s/AKfycbwu1eUxJzePa494p-343axfwgUcnHATf-db7FKw806rXZQsHn_ea0uHc6415yw-RZ80/exec'],
    [ENJOY_POINT_TITLE_, ENJOY_POINT_URL_],
    ['Google口コミに返信', 'https://business.google.com/reviews']
  ];
}

function topLinkFormula_(url, text) {
  return '=HYPERLINK("' + url + '","' + String(text).replace(/"/g, '""') + '")';
}

/**
 * トップをシンプルなリンク集に作り直す。シート名を押すとそのシートへ移動するだけ。
 * #gid リンクは非表示シートを開けないので、載せたシートは表示のままにする。
 */
function ensureTopSimple_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('TOP_LAYOUT_V') === TOP_LAYOUT_VERSION_) return { ok: true, skipped: true };
  var sh = ss.getSheetByName(HUB_HOME_SHEET_);
  if (!sh) sh = ss.insertSheet(HUB_HOME_SHEET_, 0);

  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#D9D9D9', SOFT = '#F5F5F5', RED = '#B91C1C';
  var sections = topSections_();
  var links = topExternalLinks_();
  var rows = 3 + sections.length * 3;

  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (eB) {}
  sh.clear();
  sh.clearNotes();
  try { sh.showColumns(1, sh.getMaxColumns()); } catch (eC) {}
  if (sh.getMaxColumns() > 8) sh.deleteColumns(9, sh.getMaxColumns() - 8);
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(0);

  sh.setColumnWidth(1, 76);
  for (var c = 2; c <= 6; c++) sh.setColumnWidth(c, 176);
  sh.setColumnWidth(7, 28);
  sh.setColumnWidth(8, 300);

  var all = sh.getRange(1, 1, rows, 8);
  all.setFontFamily('Noto Sans JP').setFontColor(INK).setFontSize(10)
    .setVerticalAlignment('middle').setBackground('#FFFFFF');

  sh.setRowHeight(1, 44);
  sh.getRange('A1:F1').merge().setValue('経堂 ワークスペース')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange('H1').setValue('外部リンク').setFontSize(12).setFontWeight('bold');
  sh.getRange('A1:F1').setBorder(null, null, true, null, null, null, RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange('H1').setBorder(null, null, true, null, null, null, RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(2, 26);
  sh.getRange('A2:F2').merge().setValue('シート名を押すと、そのシートへ移動します')
    .setFontSize(9).setFontColor(MUTE);
  sh.getRange('H2').setValue('元ファイル・フォーム・アプリ').setFontSize(9).setFontColor(MUTE);
  sh.setRowHeight(3, 10);

  var shown = [];
  for (var s = 0; s < sections.length; s++) {
    var top = 4 + s * 3;
    sh.setRowHeight(top, 38);
    sh.setRowHeight(top + 1, 22);
    sh.setRowHeight(top + 2, 14);
    sh.getRange(top, 1, 2, 1).merge().setValue(sections[s].label)
      .setBackground(INK).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center');
    var items = sections[s].items;
    for (var i = 0; i < items.length && i < 5; i++) {
      var col = 2 + i;
      var tile = sh.getRange(top, col);
      var sub = sh.getRange(top + 1, col);
      var target = ss.getSheetByName(items[i].name);
      if (target) {
        if (target.isSheetHidden()) { try { target.showSheet(); } catch (eS) {} }
        tile.setFormula(topLinkFormula_('#gid=' + target.getSheetId(), items[i].name));
        shown.push(items[i].name);
      } else {
        tile.setValue(items[i].name + '（なし）');
      }
      tile.setFontWeight('bold').setFontSize(11).setFontColor(INK).setFontLine('none')
        .setHorizontalAlignment('center').setWrap(true)
        .setBorder(true, true, null, true, null, null, INK, SpreadsheetApp.BorderStyle.SOLID);
      if (typeof items[i].sub === 'string') {
        sub.setFormula(topLinkFormula_(items[i].sub, '元のシート ↗'));
      } else if (items[i].sub) {
        sub.setValue(items[i].sub.text);
      }
      sub.setBackground(SOFT).setFontSize(9).setFontColor(MUTE).setFontLine('none')
        .setHorizontalAlignment('center')
        .setBorder(null, true, true, true, null, null, INK, SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  var slots = [];
  for (var t = 0; t < sections.length; t++) slots.push(4 + t * 3, 5 + t * 3);
  for (var k = 0; k < links.length && k < slots.length; k++) {
    var cell = sh.getRange(slots[k], 8);
    cell.setFormula(topLinkFormula_(links[k][1], links[k][0] + ' ↗'))
      .setFontColor(INK).setFontLine('none').setFontSize(10)
      .setBorder(null, null, true, null, null, null, LINE, SpreadsheetApp.BorderStyle.SOLID);
  }

  try { sh.setTabColor(INK); } catch (eT) {}
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
  props.setProperty('TOP_LAYOUT_V', TOP_LAYOUT_VERSION_);
  return { ok: true, shown: shown };
}

var UNPAID_TREND_SHEET_ = '未納管理_推移';
var UNPAID_DATA_ROW_ = 5;
var UNPAID_COLS_ = 41;

var KENGAKU_SOURCE_ID_ = '1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY';
var KENGAKU_JOIN_WINDOW_DAYS_ = 180;

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
var BILL_ANALYSIS_VER_ = '5';
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
  var months = billMonths_();
  var base = null;
  var lists = [];
  for (var i = 0; i < months.length; i++) {
    var tab = src.getSheetByName(months[i]);
    if (!tab) continue;
    var lastCol = Math.max(tab.getLastColumn(), 3);
    var keys = billColumnKeys_(tab.getRange(1, 1, 3, lastCol).getDisplayValues());
    if (months[i] === BILL_REF_TAB_) base = keys;
    lists.push(keys);
  }
  if (!base) {
    if (!lists.length) throw new Error('基準タブがありません: ' + BILL_REF_TAB_);
    base = lists[0].slice();
  } else {
    base = base.slice();
  }
  var have = {};
  base.forEach(function (k) { have[k.key] = true; });
  lists.forEach(function (keys) {
    for (var i = 0; i < keys.length; i++) {
      if (have[keys[i].key]) continue;
      var insertAt = base.length;
      for (var p = i - 1; p >= 0; p--) {
        for (var b = 0; b < base.length; b++) {
          if (base[b].key === keys[p].key) {
            insertAt = b + 1;
            p = -1;
            break;
          }
        }
      }
      base.splice(insertAt, 0, keys[i]);
      have[keys[i].key] = true;
    }
  });
  return base;
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
  if (sh.getName() === MEMBER_SHEET_ && (e.range.getA1Notation() === 'B1' || e.range.getA1Notation() === 'E1')) {
    memberAnalysisCharts_(sh);
    return;
  }
  if (sh.getName() === UNPAID_SHEET_ && e.range.getA1Notation() === 'B1') {
    syncUnpaidMirror_(sh.getParent(), true);
    return;
  }
  if (sh.getName() === REVIEW_TODO_SHEET_) {
    reviewTodoOnEdit_(e);
    return;
  }
  if (sh.getName() === KAIGI_SHEET_) {
    kaigiOnEdit_(e);
    return;
  }
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
  var t0 = Date.now();
  var trace = [];
  var mark = function (name, fn) {
    var s = Date.now();
    try { var res = fn(); trace.push(name + ' ' + Math.round((Date.now() - s) / 1000) + 's' + (res ? ' ' + JSON.stringify(res).slice(0, 60) : '')); }
    catch (err) { trace.push(name + ' ERR ' + String(err && err.message || err).slice(0, 80)); }
  };
  var writeTrace = function () {
    try {
      var log = ss.getSheetByName('WorkspaceSync');
      if (log) log.getRange('H1:H2').setValues([[Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm:ss')], [trace.join(' / ') + ' / total ' + Math.round((Date.now() - t0) / 1000) + 's']]);
    } catch (eLog) {}
  };
  mark('topStatus', function () { return syncTopRefreshStatus_(ss); });
  mark('reviewTodo', function () { return syncReviewTodo_(ss); });
  mark('unpaidMirror', function () { syncUnpaidMirror_(ss, false); });
  writeTrace();
  try { memberAnalysisIfChanged_(ss); } catch (e2) { console.error(e2); }
  try { billPull_(ss); } catch (e) { console.error(e); }
  try { ensureEnjoyPointLink_(ss); } catch (e3) { console.error(e3); }
  try { ensureMasterMonthAuto_(ss); } catch (e8) { console.error(e8); }
  try { ensureMasterSameDayRow_(ss); } catch (e9) { console.error(e9); }
  try { fixOctoberNippo_(); } catch (e10) { console.error(e10); }
  try { ensureNippoOpeningFemale_(); } catch (e11) { console.error(e11); }
  try { ensureNippoGenderFlow_(); } catch (e12) { console.error(e12); }
  try { ensureMasterMemberNo_(ss); } catch (e5) { console.error(e5); }
  try { ensureMasterApplyCheckmarks_(ss); } catch (e4) { console.error(e4); }
  try { linkJoinListLive_(ss, false); } catch (e6) { console.error(e6); }
  try { linkUnpaidFollowup_(ss); } catch (e7) { console.error(e7); }
  try { syncKiyakuList_(ss, false); } catch (e13) { console.error(e13); }
  try { fixHqSeptLeave_(); } catch (e14) { console.error(e14); }
  try { fixNippoOctKiyaku_(); } catch (e15) { console.error(e15); }
  mark('hqOptHistory', function () { fixHqOptionHistory_(); });
  mark('hqOptNotes', function () { clearHqOptionNotes_(); });
  mark('optPlan', function () { ensureOptionPlanSheet_(ss); });
  mark('joinLeave', function () { return ensureJoinLeaveAnalysis_(ss); });
  mark('joinLeaveHq', function () { return migrateJoinLeaveHq_(ss); });
  mark('kaigi', function () { return migrateKaigiSheet_(ss); });
  mark('kaigiHq', function () { return migrateKaigiHq_(ss); });
  mark('kaigiV2', function () { return rebuildKaigiSheet_(ss); });
  mark('jlLook', function () { return migrateJoinLeaveLook_(ss); });
  mark('jlBars', function () { return migrateJoinLeaveBars_(ss); });
  mark('jl3y', function () { return rebuildJoinLeave3y_(ss); });
  mark('jl3yCh', function () { return fixJoinLeave3yChart_(ss); });
  mark('jl3yCh2', function () { return moveJoinLeaveTrendChartApi_(ss); });
  mark('jl3ySub', function () {
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('JL_3Y_SUB') === 'v1') return { ok: true, skipped: true };
    var sh = ss.getSheetByName(JL_SHEET_);
    sh.getRange('Q5').setFormula('="3年平均 "&TEXT(AVERAGE(AN37,AN25,AN13),"0.0")');
    sh.getRange('W5').setFormula('="3年平均 "&TEXT(AVERAGE(AO37,AO25,AO13),"0.0")');
    props.setProperty('JL_3Y_SUB', 'v1');
    return { ok: true };
  });
  mark('jl3yInk', function () {
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('JL_3Y') !== JL_3Y_V_ || props.getProperty('JL_3Y_INK') === 'v1') return { ok: true, skipped: true };
    var sh = ss.getSheetByName(JL_SHEET_);
    sh.getRange('D10:F22').setFontColor('#111111');
    sh.getRange('I10:K22').setFontColor('#111111');
    sh.getRange('G10:G22').setFontColor('#111111').setFontWeight('bold');
    sh.getRange('L10:L22').setFontColor('#111111').setFontWeight('bold');
    props.setProperty('JL_3Y_INK', 'v1');
    return { ok: true };
  });
  mark('kaigiMon', function () { return addKaigiMonthly_(ss); });
  mark('kaigiFrom', function () { return addKaigiFromMonth_(ss); });
  mark('kaigiSurvey', function () { return addKaigiSurvey_(ss); });
  mark('kaigiByMon', function () { return rebuildKaigiByMonth_(ss); });
  mark('kaigiSurveyW', function () {
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('KAIGI_SURVEY_W') === 'v1') return { ok: true, skipped: true };
    ss.getSheetByName(KAIGI_SHEET_).setColumnWidth(15, 160);
    props.setProperty('KAIGI_SURVEY_W', 'v1');
    return { ok: true };
  });
  mark('kaigiSumFix', function () {
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('KAIGI_SUMFIX') === 'v1') return { ok: true, skipped: true };
    var agg = ss.getSheetByName('分析用_期間集計');
    var from = 'LET(ms,EDATE(s,SEQUENCE(DATEDIF(s,e+1,"M"),1,0)),', to = 'LET(ms,ARRAYFORMULA(EDATE(s,SEQUENCE(DATEDIF(s,e+1,"M"),1,0))),';
    var n = 0;
    ['B', 'C', 'D', 'E'].forEach(function (X) {
      [20, 21].forEach(function (r) {
        var cell = agg.getRange(X + r), f = cell.getFormula();
        if (f.indexOf(from) >= 0) { cell.setFormula(f.split(from).join(to)); n++; }
      });
    });
    props.setProperty('KAIGI_SUMFIX', 'v1');
    return { ok: true, fixed: n };
  });
  mark('top', function () { ensureTopSimple_(ss); });
  writeTrace();
}

/**
 * 会員分析：日報は IMPORTRANGE。年代と男女は累計の入会月・退会月を COUNTIFS で数える。
 * B1 で月末・月初・入会・退会、E1 で当月・半年・昨年比。
 */
var MEMBER_SHEET_ = '会員分析';
var MEMBER_JOIN_SRC_ = '累計入会データ';
var MEMBER_LEAVE_SRC_ = '累計退会データ';
var MEMBER_ANALYSIS_VER_ = '16';

function memberAnalysisNote_(ss, msg) {
  try {
    var mark = ss.getSheetByName('学割');
    if (!mark) return;
    if (!msg) {
      mark.getRange('Z1').clearContent();
      return;
    }
    mark.getRange('Z1').setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss') + ' ' + String(msg).slice(0, 180));
  } catch (eNote) {}
}

function memberAnalysisIfChanged_(ss) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    var j = ss.getSheetByName(MEMBER_JOIN_SRC_);
    var l = ss.getSheetByName(MEMBER_LEAVE_SRC_);
    if (!j || !l) return;
    var sig = MEMBER_ANALYSIS_VER_ + ':' + j.getLastRow() + ':' + l.getLastRow() + ':' + (ss.getSheetByName(MEMBER_SHEET_) ? 1 : 0);
    var props = PropertiesService.getDocumentProperties();
    var sh = ss.getSheetByName(MEMBER_SHEET_);
    if (props.getProperty('MEMBER_ANALYSIS_SIG') === sig) {
      var chartErr = '';
      if (sh && sh.getCharts().length === 0) {
        try { memberAnalysisCharts_(sh); } catch (eChart) { chartErr = String(eChart.message || eChart); }
      }
      if (chartErr) memberAnalysisNote_(ss, 'charts ' + chartErr);
      else if (sh && sh.getCharts().length > 0) memberAnalysisNote_(ss, '');
      else memberAnalysisNote_(ss, 'charts 0');
      return;
    }
    memberAnalysisNote_(ss, 'start');
    buildMemberAnalysis_(ss);
    props.setProperty('MEMBER_ANALYSIS_SIG', sig);
    memberAnalysisNote_(ss, 'done');
    try { memberAnalysisCharts_(ss.getSheetByName(MEMBER_SHEET_)); } catch (eChart2) {}
  } catch (err) {
    memberAnalysisNote_(ss, 'ERR ' + (err && err.message ? err.message : err));
  } finally {
    lock.releaseLock();
  }
}

function setFormulaIfChanged_(range, formula) {
  if (range.getFormula() === formula) return;
  range.setFormula(formula);
}

function ensureMemberCalcColumns_(ss) {
  var join = ss.getSheetByName(MEMBER_JOIN_SRC_);
  var leave = ss.getSheetByName(MEMBER_LEAVE_SRC_);
  if (!join || !leave) return;
  var end = 8000;
  leave.getRange(1, 20).setValue('番号キー（自動）').setFontColor('#757575');
  setFormulaIfChanged_(leave.getRange(2, 20), '=ARRAYFORMULA(IF(F2:F' + end + '="","",TEXT(VALUE(F2:F' + end + '),"0")))');
  leave.getRange(1, 21).setValue('退会月（自動）').setFontColor('#757575');
  setFormulaIfChanged_(leave.getRange(2, 21),
    '=ARRAYFORMULA(IF(R2:R' + end + '="","",IF(ISNUMBER(R2:R' + end + '),DATE(YEAR(R2:R' + end + '),MONTH(R2:R' + end + '),1),IFERROR(DATE(VALUE(LEFT(R2:R' + end + '&"",4)),VALUE(MID(R2:R' + end + '&"",6,2)),1),""))))'
  );
  join.getRange(1, 20).setValue('番号キー（自動）').setFontColor('#757575');
  setFormulaIfChanged_(join.getRange(2, 20), '=ARRAYFORMULA(IF(F2:F' + end + '="","",TEXT(VALUE(F2:F' + end + '),"0")))');
  join.getRange(1, 21).setValue('退会年月（自動）').setFontColor('#757575');
  setFormulaIfChanged_(join.getRange(2, 21),
    '=BYROW(T2:T' + end + ',LAMBDA(k,IF(k="","",IF(COUNTIF(\'累計退会データ\'!T$2:T$' + end + ',k)=0,' +
      'IFERROR(INDEX(\'' + KIYAKU_SHEET_ + '\'!K$4:K$500,MATCH(VALUE(k),\'' + KIYAKU_SHEET_ + '\'!C$4:C$500,0)),""),' +
      'MAXIFS(\'累計退会データ\'!U$2:U$' + end + ',\'累計退会データ\'!T$2:T$' + end + ',k)))))'
  );
  join.getRange(1, 22).setValue('入会月（自動）').setFontColor('#757575');
  setFormulaIfChanged_(join.getRange(2, 22),
    '=ARRAYFORMULA(IF(N2:N' + end + '="",0,IF(ISNUMBER(N2:N' + end + '),DATE(YEAR(N2:N' + end + '),MONTH(N2:N' + end + '),1),IFERROR(DATE(VALUE(LEFT(N2:N' + end + '&"",4)),VALUE(MID(N2:N' + end + '&"",6,2)),1),0))))'
  );
  join.getRange(1, 23).setValue('退会月（自動）').setFontColor('#757575');
  setFormulaIfChanged_(join.getRange(2, 23), '=ARRAYFORMULA(IF((U2:U' + end + '="")+(U2:U' + end + '=0),"",U2:U' + end + '))');
  join.getRange(1, 24).setValue('年齢（自動）').setFontColor('#757575');
  setFormulaIfChanged_(join.getRange(2, 24),
    '=ARRAYFORMULA(IF(J2:J' + end + '="",0,IF(ISNUMBER(J2:J' + end + '),J2:J' + end + ',IFERROR(VALUE(REGEXREPLACE(J2:J' + end + '&"","[^0-9.]","")),0))))'
  );
  try { join.hideColumns(20, 5); } catch (eJ) {}
  try { leave.hideColumns(20, 2); } catch (eL) {}
}

function ensureVisitKind_(ss) {
  var sh = ss.getSheetByName('見学体験申請');
  if (!sh) return;
  sh.getRange(1, 11).setValue('種別').setFontColor('#757575');
  setFormulaIfChanged_(sh.getRange(2, 11),
    '=MAP(B2:B400,E2:E400,LAMBDA(k,p,IF(k="","",IF(OR(p="",AND(COUNTIFS(E$2:E$400,p,B$2:B$400,"見学")=0,COUNTIFS(E$2:E$400,p,B$2:B$400,"体験")=0)),k,IF(AND(COUNTIFS(E$2:E$400,p,B$2:B$400,"見学")>0,COUNTIFS(E$2:E$400,p,B$2:B$400,"体験")>0),"両方",k)))))'
  );
  try { sh.hideColumns(11, 1); } catch (eK) {}
}

function memberSliceFormula_(monthCell, gender, ageMin, ageMax) {
  var s = "'累計入会データ'!";
  var pairs = [
    [s + 'C2:C8000', '"<>法人会員(都度利用)"'],
    [s + 'C2:C8000', '"<>OGF会員"'],
    [s + 'C2:C8000', '"<>ゴールド会員"']
  ];
  if (gender) pairs.push([s + 'I2:I8000', '"' + gender + '"']);
  if (ageMin != null && ageMax != null && ageMin <= 0) {
    pairs.push([s + 'X2:X8000', '">0"']);
    pairs.push([s + 'X2:X8000', '"<' + ageMax + '"']);
  } else if (ageMin != null && ageMax != null && ageMax >= 200) {
    pairs.push([s + 'X2:X8000', '">=' + ageMin + '"']);
  } else if (ageMin != null && ageMax != null) {
    pairs.push([s + 'X2:X8000', '">=' + ageMin + '"']);
    pairs.push([s + 'X2:X8000', '"<' + ageMax + '"']);
  }
  function count(extra) {
    return 'COUNTIFS(' + pairs.concat(extra).map(function (p) { return p[0] + ',' + p[1]; }).join(',') + ')';
  }
  var V = s + 'V2:V8000';
  var W = s + 'W2:W8000';
  var opening = '(' + count([[V, '">0"'], [V, '"<"&' + monthCell], [W, '""']]) + '+' + count([[V, '">0"'], [V, '"<"&' + monthCell], [W, '">="&' + monthCell]]) + ')';
  var enrolled = '(' + count([[V, '">0"'], [V, '"<="&' + monthCell], [W, '""']]) + '+' + count([[V, '">0"'], [V, '"<="&' + monthCell], [W, '">"&' + monthCell]]) + ')';
  var joined = count([[V, monthCell]]);
  var left = count([[W, monthCell]]);
  return '=IF(' + monthCell + '="","",IF($B$1="入会者",' + joined + ',IF($B$1="退会",' + left + ',IF($B$1="月初会員",' + opening + ',' + enrolled + '))))';
}

function buildMemberAnalysis_(ss) {
  ensureMemberCalcColumns_(ss);
  ensureVisitKind_(ss);
  var sh = ss.getSheetByName(MEMBER_SHEET_);
  if (!sh) {
    var afterBill = ss.getSheetByName(BILL_SHEET_);
    sh = ss.insertSheet(MEMBER_SHEET_, afterBill ? afterBill.getIndex() : ss.getSheets().length);
  }
  sh.clear();
  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  sh.clearConditionalFormatRules();
  try { sh.getRange(1, 1, Math.min(sh.getMaxRows(), 80), Math.min(sh.getMaxColumns(), 40)).breakApart(); } catch (eB) {}
  try { sh.getRange(1, 1, 2, 8).clearDataValidations(); } catch (eV) {}
  sh.setHiddenGridlines(true);
  sh.setTabColor('#111111');

  var nip = function (a1) { return 'IMPORTRANGE("' + RECEPTION_SOURCE_ID_ + '","日報!' + a1 + '")'; };
  var mf = function (a1) { return '"男 "&' + nip('F' + a1) + '&" / 女 "&' + nip('H' + a1); };
  var grey = '#757575';
  var bands = [
    ['10代', 0, 20], ['20代', 20, 30], ['30代', 30, 40], ['40代', 40, 50],
    ['50代', 50, 60], ['60代', 60, 70], ['70代以上', 70, 200]
  ];
  var months = ['$AA$1', '$AB$1', '$AC$1', '$AD$1', '$AE$1', '$AF$1'];
  var countCols = [2, 4, 6, 8, 10, 12];
  var pctCols = [3, 5, 7, 9, 11, 13];

  sh.getRange('A1').setValue('会員分析').setFontSize(16).setFontWeight('bold').setVerticalAlignment('middle');
  sh.getRange('B1:C1').merge();
  sh.getRange('B1').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['月末会員', '月初会員', '入会者', '退会'], true).setAllowInvalid(false).build());
  sh.getRange('B1').setValue('月末会員').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#111111').setFontColor('#ffffff').setVerticalAlignment('middle');
  sh.getRange('D1').setValue('比較').setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.getRange('E1:F1').merge();
  sh.getRange('E1').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['当月', '半年', '昨年比'], true).setAllowInvalid(false).build());
  sh.getRange('E1').setValue('半年').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#111111').setFontColor('#ffffff').setVerticalAlignment('middle');
  sh.getRange('H1').setFormula('="更新 "&TEXT(NOW(),"M/d HH:mm")').setFontSize(9).setFontColor(grey);
  sh.getRange('A2').setFormula('=IFERROR("今月（"&' + nip('B1') + '&"）","今月")').setFontWeight('bold').setFontSize(12);

  sh.getRange('AA1').setFormula('=IFERROR(EDATE(DATE(2000+INT(' + nip('B1') + '/100),MOD(' + nip('B1') + ',100),1),IF($B$1="月末会員",-1,0)),"")');
  sh.getRange('AB1').setFormula('=IF($E$1="半年",EDATE($AA$1,-1),IF($E$1="昨年比",EDATE($AA$1,-12),""))');
  sh.getRange('AC1').setFormula('=IF($E$1="半年",EDATE($AA$1,-2),"")');
  sh.getRange('AD1').setFormula('=IF($E$1="半年",EDATE($AA$1,-3),"")');
  sh.getRange('AE1').setFormula('=IF($E$1="半年",EDATE($AA$1,-4),"")');
  sh.getRange('AF1').setFormula('=IF($E$1="半年",EDATE($AA$1,-5),"")');
  sh.getRange('AA1:AF1').setNumberFormat('yyyy/m');

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
  sh.getRange(3, 1, 1, 8).setBackground('#111111').setFontColor('#ffffff').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(4, 1, 1, 8).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(5, 1, 1, 8).setFontSize(9).setFontColor(grey).setHorizontalAlignment('center');
  sh.getRange('A4:G4').setNumberFormat('#,##0');
  sh.getRange('D4').setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange('H4').setNumberFormat('0.00%');

  sh.getRange('A7').setFormula('=IF($B$1="入会者","入会",IF($B$1="退会","退会",IF($B$1="月初会員","月初","月末")))');
  sh.getRange(8, 1, 1, 4).setValues([['', '日報', '累計', '差']]).setBackground('#333333').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A9').setValue('人数');
  sh.getRange('A10').setValue('男性');
  sh.getRange('A11').setValue('女性');
  sh.getRange('B9').setFormula('=IF($B$1="入会者",IFERROR(' + nip('C13') + ',""),IF($B$1="退会",IFERROR(' + nip('C15') + ',""),IF($B$1="月初会員",IFERROR(' + nip('C12') + ',""),IFERROR(' + nip('C16') + ',""))))');
  sh.getRange('B10').setFormula('=IF($B$1="月末会員",IFERROR(' + nip('F16') + ',""),IF($B$1="月初会員",IFERROR(' + nip('F12') + ',""),""))');
  sh.getRange('B11').setFormula('=IF($B$1="月末会員",IFERROR(' + nip('H16') + ',""),IF($B$1="月初会員",IFERROR(' + nip('H12') + ',""),""))');
  sh.getRange('C9').setFormula(memberSliceFormula_('$AA$1', '', null, null));
  sh.getRange('C10').setFormula(memberSliceFormula_('$AA$1', '男', null, null));
  sh.getRange('C11').setFormula(memberSliceFormula_('$AA$1', '女', null, null));
  sh.getRange('D9').setFormula('=IF(OR(B9="",C9=""),"",C9-B9)');
  sh.getRange('D10').setFormula('=IF(OR(B10="",C10=""),"",C10-B10)');
  sh.getRange('D11').setFormula('=IF(OR(B11="",C11=""),"",C11-B11)');
  sh.getRange('A7').setFontWeight('bold').setFontSize(12);
  sh.getRange('B9:D11').setNumberFormat('#,##0');
  sh.getRange('D9:D11').setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange('A9:A11').setFontWeight('bold').setBackground('#fafafa');
  sh.getRange('A9:D11').setBorder(true, true, true, true, true, true, '#eeeeee', SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');
  sh.getRange('A9:A11').setHorizontalAlignment('left');

  function monthHead(cell) {
    return '=IF(' + cell + '="","",TEXT(' + cell + ',"m月"))';
  }
  function pctHead(cell) {
    return '=IF(' + cell + '="","","%")';
  }
  var wide = 1 + months.length * 2;
  function writeBandTable(row, title) {
    var titleRg = sh.getRange(row, 1, 1, wide);
    titleRg.merge();
    titleRg.setFormula(title).setBackground('#111111').setFontColor('#ffffff').setFontWeight('bold').setVerticalAlignment('middle');
    sh.setRowHeight(row, 24);
    var head = ['="年代"'];
    for (var h = 0; h < months.length; h++) {
      head.push(monthHead(months[h]));
      head.push(pctHead(months[h]));
    }
    sh.getRange(row + 1, 1, 1, wide).setFormulas([head]).setBackground('#333333').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    for (var i = 0; i < bands.length; i++) {
      var r = row + 2 + i;
      sh.getRange(r, 1).setValue(bands[i][0]).setFontWeight('bold').setBackground('#fafafa');
      for (var k = 0; k < months.length; k++) {
        var countA1 = columnLetter_(countCols[k]) + r;
        var sumA1 = columnLetter_(countCols[k]) + '$' + (row + 2) + ':' + columnLetter_(countCols[k]) + '$' + (row + 1 + bands.length);
        sh.getRange(r, countCols[k]).setFormula(memberSliceFormula_(months[k], '', bands[i][1], bands[i][2]));
        sh.getRange(r, pctCols[k]).setFormula('=IF(SUM(' + sumA1 + ')=0,"",' + countA1 + '/SUM(' + sumA1 + '))');
      }
    }
    sh.getRange(row + 2, 2, bands.length, wide - 1).setNumberFormat('#,##0');
    for (var pf = 0; pf < pctCols.length; pf++) sh.getRange(row + 2, pctCols[pf], bands.length, 1).setNumberFormat('0%');
    sh.getRange(row + 2, 1, bands.length, wide).setBorder(true, true, true, true, true, true, '#eeeeee', SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');
    sh.getRange(row + 2, 1, bands.length, 1).setHorizontalAlignment('left');
  }
  writeBandTable(13, '=IF($B$1="入会者","入会",IF($B$1="退会","退会",IF($B$1="月初会員","月初","月末")))&"の年代"');

  var gTitle = sh.getRange(23, 1, 1, wide);
  gTitle.merge();
  gTitle.setFormula('=IF($B$1="入会者","入会",IF($B$1="退会","退会",IF($B$1="月初会員","月初","月末")))&"の男女"')
    .setBackground('#111111').setFontColor('#ffffff').setFontWeight('bold');
  sh.setRowHeight(23, 24);
  var gHead = ['="性別"'];
  for (var gh = 0; gh < months.length; gh++) {
    gHead.push(monthHead(months[gh]));
    gHead.push(pctHead(months[gh]));
  }
  sh.getRange(24, 1, 1, wide).setFormulas([gHead]).setBackground('#333333').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A25').setValue('男').setFontWeight('bold').setBackground('#fafafa');
  sh.getRange('A26').setValue('女').setFontWeight('bold').setBackground('#fafafa');
  var genders = ['男', '女'];
  for (var gi = 0; gi < 2; gi++) {
    for (var gk = 0; gk < months.length; gk++) {
      var gr = 25 + gi;
      var gCount = columnLetter_(countCols[gk]) + gr;
      var gSum = columnLetter_(countCols[gk]) + '$25:' + columnLetter_(countCols[gk]) + '$26';
      sh.getRange(gr, countCols[gk]).setFormula(memberSliceFormula_(months[gk], genders[gi], null, null));
      sh.getRange(gr, pctCols[gk]).setFormula('=IF(SUM(' + gSum + ')=0,"",' + gCount + '/SUM(' + gSum + '))');
    }
  }
  sh.getRange(25, 2, 2, wide - 1).setNumberFormat('#,##0');
  for (var pg = 0; pg < pctCols.length; pg++) sh.getRange(25, pctCols[pg], 2, 1).setNumberFormat('0%');
  sh.getRange(25, 1, 2, wide).setBorder(true, true, true, true, true, true, '#eeeeee', SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');
  sh.getRange('A25:A26').setHorizontalAlignment('left');

  var side = 15;
  sh.getRange(13, side).setValue('申込').setBackground('#111111').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(13, side, 1, 7).merge();
  sh.getRange(14, side).setValue('').setBackground('#333333');
  for (var sm = 0; sm < months.length; sm++) {
    sh.getRange(14, side + 1 + sm).setFormula(monthHead(months[sm])).setBackground('#333333').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }
  var kinds = ['見学', '体験', '両方', '学割'];
  function sideCount(row, kind) {
    sh.getRange(row, side).setValue(kind).setFontWeight('bold').setBackground('#fafafa');
    for (var sk = 0; sk < months.length; sk++) {
      var cell = months[sk];
      var formula;
      if (kind === '学割') {
        formula = '=IF(' + cell + '="","",SUMPRODUCT((\'学割\'!A$2:A$400<>"")*(TEXT(\'学割\'!A$2:A$400,"yyyy-m")=TEXT(' + cell + ',"yyyy-m"))))';
      } else {
        formula = '=IF(' + cell + '="","",SUMPRODUCT((\'見学体験申請\'!K$2:K$400="' + kind + '")*(TEXT(\'見学体験申請\'!H$2:H$400,"yyyy-m")=TEXT(' + cell + ',"yyyy-m"))))';
      }
      sh.getRange(row, side + 1 + sk).setFormula(formula).setNumberFormat('#,##0').setHorizontalAlignment('center');
    }
  }
  for (var si = 0; si < kinds.length; si++) sideCount(15 + si, kinds[si]);
  sh.getRange(15, side, 4, 7).setBorder(true, true, true, true, true, true, '#eeeeee', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(20, side).setValue('移籍').setFontWeight('bold');
  sh.getRange(20, side + 1).setFormula('=IFERROR(' + nip('D14') + ',"")').setNumberFormat('#,##0');
  sh.getRange(20, side + 2).setValue('今月').setFontColor(grey).setFontSize(9);
  sh.getRange(22, side).setValue('特典').setBackground('#111111').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(22, side, 1, 2).merge();
  sh.getRange(23, side).setFormula('=IFERROR(IMPORTRANGE("' + RECEPTION_SOURCE_ID_ + '","日報!B21:C36"),"")');

  var attr = side + 8;
  sh.getRange(13, attr).setFormula('=IF($AA$1="","申込",TEXT($AA$1,"m月")&"の申込")').setBackground('#111111').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(13, attr, 1, 2).merge();
  sh.getRange(14, attr, 1, 2).setValues([['性別', '件数']]).setBackground('#333333').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(15, attr).setValue('男').setFontWeight('bold').setBackground('#fafafa');
  sh.getRange(16, attr).setValue('女').setFontWeight('bold').setBackground('#fafafa');
  function sexFormula(label) {
    return '=IF($AA$1="","",SUMPRODUCT((\'見学体験申請\'!B$2:B$400<>"")*(TEXT(\'見学体験申請\'!H$2:H$400,"yyyy-m")=TEXT($AA$1,"yyyy-m"))*((\'見学体験申請\'!F$2:F$400="' + label + '")+(\'見学体験申請\'!F$2:F$400="' + label + '性"))))';
  }
  sh.getRange(15, attr + 1).setFormula(sexFormula('男')).setNumberFormat('#,##0');
  sh.getRange(16, attr + 1).setFormula(sexFormula('女')).setNumberFormat('#,##0');
  var ageLabels = ['10代', '20代', '30代', '40代', '50代', '60代', '70代以上'];
  sh.getRange(18, attr, 1, 2).setValues([['年代', '件数']]).setBackground('#333333').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  for (var ai = 0; ai < ageLabels.length; ai++) {
    sh.getRange(19 + ai, attr).setValue(ageLabels[ai]).setFontWeight('bold').setBackground('#fafafa');
    var ageMatch = ageLabels[ai] === '70代以上'
      ? '((\'見学体験申請\'!G$2:G$400="70代")+(\'見学体験申請\'!G$2:G$400="70代以上"))'
      : '(\'見学体験申請\'!G$2:G$400="' + ageLabels[ai] + '")';
    sh.getRange(19 + ai, attr + 1).setFormula(
      '=IF($AA$1="","",SUMPRODUCT((\'見学体験申請\'!B$2:B$400<>"")*(TEXT(\'見学体験申請\'!H$2:H$400,"yyyy-m")=TEXT($AA$1,"yyyy-m"))*' + ageMatch + '))'
    ).setNumberFormat('#,##0');
  }
  sh.getRange(15, attr, 2, 2).setBorder(true, true, true, true, true, true, '#eeeeee', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(19, attr, ageLabels.length, 2).setBorder(true, true, true, true, true, true, '#eeeeee', SpreadsheetApp.BorderStyle.SOLID);

  if (sh.getMaxColumns() < 48) sh.insertColumnsAfter(sh.getMaxColumns(), 48 - sh.getMaxColumns());
  for (var p = 0; p < bands.length; p++) {
    var srcRow = 15 + p;
    sh.getRange(10 + p, 40).setFormula('=IF(A' + srcRow + '="","",A' + srcRow + '&" "&TEXT(C' + srcRow + ',"0%"))');
    sh.getRange(10 + p, 41).setFormula('=IF(B' + srcRow + '="","",B' + srcRow + ')');
    sh.getRange(10 + p, 43).setFormula('=IF(OR($AB$1="",A' + srcRow + '=""),"",A' + srcRow + '&" "&TEXT(E' + srcRow + ',"0%"))');
    sh.getRange(10 + p, 44).setFormula('=IF($AB$1="","",D' + srcRow + ')');
    sh.getRange(10 + p, 46).setFormula('=IF(OR($AC$1="",A' + srcRow + '=""),"",A' + srcRow + '&" "&TEXT(G' + srcRow + ',"0%"))');
    sh.getRange(10 + p, 47).setFormula('=IF($AC$1="","",F' + srcRow + ')');
  }
  try { sh.hideColumns(27, 6); } catch (eDate) {}
  try { sh.hideColumns(40, 8); } catch (eH) {}

  var pctRanges = [];
  for (var pr = 0; pr < pctCols.length; pr++) {
    pctRanges.push(sh.getRange(15, pctCols[pr], 7, 1));
    pctRanges.push(sh.getRange(25, pctCols[pr], 2, 1));
  }
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.3).setBackground('#F5C518').setFontColor('#111111').setBold(true).setRanges(pctRanges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.2, 0.2999).setBackground('#7EC8E3').setFontColor('#111111').setRanges(pctRanges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.1, 0.1999).setBackground('#C5DDB3').setFontColor('#111111').setRanges(pctRanges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.0001, 0.0999).setBackground('#F3F3F3').setFontColor('#666666').setRanges(pctRanges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberNotEqualTo(0).setFontColor('#b00020').setBold(true).setRanges([sh.getRange('D9:D11')]).build()
  ];
  sh.setConditionalFormatRules(rules);

  sh.getRange(1, 1, 28, wide).setFontFamily('Meiryo');
  sh.getRange(13, side, 16, 12).setFontFamily('Meiryo');
  sh.setColumnWidth(1, 72);
  sh.setColumnWidths(2, wide - 1, 48);
  sh.setColumnWidth(side, 64);
  sh.setColumnWidths(side + 1, 6, 48);
  sh.setColumnWidth(attr, 72);
  sh.setColumnWidth(attr + 1, 48);
  sh.setRowHeight(1, 32);
  sh.setRowHeight(4, 36);
  sh.setFrozenRows(5);
  memberAnalysisNote_(ss, 'tables');
  return { ok: true };
}

function memberAnalysisMonthTitles_(sh) {
  var nip = sh.getParent().getSheetByName('日報');
  var raw = nip ? String(nip.getRange('B1').getDisplayValue() || nip.getRange('B1').getValue() || '') : '';
  var n = parseInt(raw, 10);
  var pop = String(sh.getRange('B1').getValue() || '月末会員');
  var period = String(sh.getRange('E1').getValue() || '当月');
  var word = pop === '入会者' ? '入会' : pop === '退会' ? '退会' : pop === '月初会員' ? '月初' : '月末';
  if (!n || n < 1000) return [word];
  var y = 2000 + Math.floor(n / 100);
  var m = n % 100;
  var base = pop === '月末会員' ? -1 : 0;
  var shifts = [base];
  if (period === '半年') shifts = [base, base - 1, base - 2];
  else if (period === '昨年比') shifts = [base, base - 12];
  return shifts.map(function (s) {
    var dt = new Date(y, m - 1 + s, 1);
    return (dt.getMonth() + 1) + '月';
  });
}

function memberAnalysisCharts_(sh) {
  if (!sh) return;
  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  var colors = ['#2F6FED', '#14B8A6', '#3FA34D', '#E2B007', '#E07A2F', '#D94B4B', '#7C5CBF'];
  var titles = memberAnalysisMonthTitles_(sh);
  var sources = [40, 43, 46];
  var anchors = [1, 5, 9];
  for (var i = 0; i < titles.length; i++) {
    var chart = sh.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sh.getRange(10, sources[i], 7, 2))
      .setNumHeaders(0)
      .setOption('title', titles[i])
      .setOption('pieHole', 0.42)
      .setOption('pieSliceText', 'percentage')
      .setOption('sliceVisibilityThreshold', 0)
      .setOption('pieSliceTextStyle', { fontName: 'Meiryo', fontSize: 12 })
      .setOption('legend', { position: 'right', textStyle: { fontName: 'Meiryo', fontSize: 12, color: '#111111' } })
      .setOption('titleTextStyle', { fontName: 'Meiryo', fontSize: 15, bold: true, color: '#111111' })
      .setOption('colors', colors)
      .setOption('backgroundColor', '#ffffff')
      .setOption('pieSliceBorderColor', '#ffffff')
      .setOption('width', 400)
      .setOption('height', 360)
      .setPosition(28, anchors[i], 0, 0)
      .build();
    sh.insertChart(chart);
  }
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

/**
 * 規約退会リスト：未納管理ドライブの月シートで会員名（D列）の背景が赤の人＝2ヶ月未納で強制退会。
 * 元のシートは読むだけ。強制退会日は各シート備考の「M/D … 強制退会」、無ければ翌月1日。
 * その月のシートで赤 → 翌月頭に強制退会 → 会員動向ではそのシートの月の退会。
 */
var KIYAKU_SHEET_ = '規約退会リスト';
var KIYAKU_RED_ = '#ff0000';
var KIYAKU_LAYOUT_V_ = 'v1';

function syncKiyakuList_(ss, force) {
  var props = PropertiesService.getDocumentProperties();
  var now = Date.now();
  var last = Number(props.getProperty('KIYAKU_SYNC_AT') || 0);
  if (!force && ss.getSheetByName(KIYAKU_SHEET_) && now - last < 55 * 60 * 1000) return { ok: true, skipped: true };

  var src = SpreadsheetApp.openById(UNPAID_SOURCE_ID_);
  var rows = [];
  src.getSheets().forEach(function (s) {
    var m = s.getName().match(/^(\d{2})年(\d{1,2})月$/);
    if (!m) return;
    var y = 2000 + Number(m[1]);
    var mon = Number(m[2]);
    var lr = Math.min(s.getLastRow(), 200);
    if (lr < 3) return;
    var vals = s.getRange(1, 1, lr, 9).getDisplayValues();
    var bgs = s.getRange(1, 4, lr, 1).getBackgrounds();
    var forced = null;
    for (var i = 0; i < vals.length; i++) {
      var hit = String(vals[i][0] || '').match(/(\d{1,2})\/(\d{1,2})[^\n]*強制退会/);
      if (hit) {
        var fm = Number(hit[1]);
        forced = new Date(fm < mon ? y + 1 : y, fm - 1, Number(hit[2]));
      }
    }
    if (!forced) forced = new Date(y, mon, 1);
    for (var r = 0; r < vals.length; r++) {
      if (String(bgs[r][0]).toLowerCase() !== KIYAKU_RED_) continue;
      var no = String(vals[r][2] || '').replace(/\s/g, '');
      if (!no) continue;
      rows.push([
        forced,
        y + '年' + mon + '月',
        /^\d+$/.test(no) ? Number(no) : no,
        vals[r][3],
        vals[r][4] ? "'" + vals[r][4] : '',
        vals[r][5],
        vals[r][6],
        vals[r][7],
        s.getName(),
        '',
        new Date(y, mon, 1)
      ]);
    }
  });
  rows.sort(function (a, b) { return b[0] - a[0]; });

  var sh = ss.getSheetByName(KIYAKU_SHEET_);
  if (!sh) {
    var after = ss.getSheetByName(UNPAID_TREND_SHEET_) || ss.getSheetByName(UNPAID_SHEET_);
    sh = ss.insertSheet(KIYAKU_SHEET_, after ? after.getIndex() : ss.getSheets().length);
  }
  if (props.getProperty('KIYAKU_LAYOUT') !== KIYAKU_LAYOUT_V_) {
    kiyakuLayout_(sh);
    props.setProperty('KIYAKU_LAYOUT', KIYAKU_LAYOUT_V_);
  }
  var maxRow = Math.max(sh.getLastRow(), 4);
  sh.getRange(4, 1, maxRow - 3, 11).clearContent();
  if (rows.length) {
    if (sh.getMaxRows() < rows.length + 3) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 3 - sh.getMaxRows());
    sh.getRange(4, 1, rows.length, 11).setValues(rows);
    var f = [];
    for (var k = 0; k < rows.length; k++) {
      var c = 'C' + (k + 4);
      f.push(['=IF(' + c + '="","",IF(ISNUMBER(MATCH(' + c + '&"",\'累計退会データ\'!T:T,0)),"反映済み","未反映（会員数から外す）"))']);
    }
    sh.getRange(4, 10, rows.length, 1).setFormulas(f);
  }
  sh.getRange('B1').setValue(rows.length + '名');
  sh.getRange('D1').setValue('更新 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm'));
  props.setProperty('KIYAKU_SYNC_AT', String(now));
  try { ensureMemberCalcColumns_(ss); } catch (eCalc) { console.error(eCalc); }
  return { ok: true, count: rows.length };
}

/**
 * 未納管理：元の未納管理ドライブの月タブ（B1）を、色・取り消し線・チェックボックスごとそのまま写す。
 * 元シートは読むだけ。1〜4行目（集計）はこちらで計算、5行目以降が元シートの1行目以降。
 */
var UNPAID_MIRROR_TOP_ = 5;

function unpaidMonthSheets_(src) {
  return src.getSheets().map(function (s) { return s.getName(); })
    .filter(function (n) { return /^\d{2}年\d{1,2}月$/.test(n); })
    .sort(function (a, b) {
      var pa = a.match(/(\d+)年(\d+)月/), pb = b.match(/(\d+)年(\d+)月/);
      return (Number(pb[1]) * 12 + Number(pb[2])) - (Number(pa[1]) * 12 + Number(pa[2]));
    });
}

function syncUnpaidMirror_(ss, force) {
  var props = PropertiesService.getDocumentProperties();
  var now = Date.now();
  if (!force && now - Number(props.getProperty('UNPAID_MIRROR_AT') || 0) < 9 * 60 * 1000) return { ok: true, skipped: true };
  var sh = ss.getSheetByName(UNPAID_SHEET_);
  if (!sh) return { ok: false };
  var src = SpreadsheetApp.openById(UNPAID_SOURCE_ID_);
  var months = unpaidMonthSheets_(src);
  if (!months.length) return { ok: false };

  var b1 = String(sh.getRange('B1').getDisplayValue() || '').trim();
  var prevLatest = props.getProperty('UNPAID_LATEST') || '';
  if (months.indexOf(b1) < 0 || (prevLatest && b1 === prevLatest && months[0] !== prevLatest)) b1 = months[0];
  props.setProperty('UNPAID_LATEST', months[0]);
  sh.getRange('B1').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(months, true).setAllowInvalid(false).build());
  if (String(sh.getRange('B1').getDisplayValue()) !== b1) sh.getRange('B1').setNumberFormat('@').setValue(b1);

  var tab = src.getSheetByName(b1);
  var lr = Math.max(tab.getLastRow(), 3);
  var lc = Math.max(tab.getLastColumn(), 37);
  var rg = tab.getRange(1, 1, lr, lc);
  var values = rg.getValues();
  var shown = rg.getDisplayValues();
  for (var vr = 0; vr < values.length; vr++) {
    for (var vc = 0; vc < values[vr].length; vc++) {
      var v0 = values[vr][vc];
      if ((v0 === '' && shown[vr][vc] !== '') || (typeof v0 === 'string' && /^[0-9+\-=]/.test(v0))) values[vr][vc] = "'" + (v0 === '' ? shown[vr][vc] : v0);
    }
  }
  var top = UNPAID_MIRROR_TOP_;
  var needRows = top + lr - 1;
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());
  if (sh.getMaxColumns() < lc) sh.insertColumnsAfter(sh.getMaxColumns(), lc - sh.getMaxColumns());

  sh.setFrozenColumns(0);
  var body = sh.getRange(top, 1, sh.getMaxRows() - top + 1, sh.getMaxColumns());
  try { body.breakApart(); } catch (eBa) {}
  body.clear();
  body.clearDataValidations();
  sh.setConditionalFormatRules([]);

  var dest = sh.getRange(top, 1, lr, lc);
  var formats = rg.getNumberFormats();
  for (var fr = 0; fr < values.length; fr++) {
    for (var fc = 0; fc < values[fr].length; fc++) {
      if (typeof values[fr][fc] === 'string' && values[fr][fc].charAt(0) === "'") {
        formats[fr][fc] = '@';
        values[fr][fc] = values[fr][fc].slice(1);
      }
    }
  }
  dest.setNumberFormats(formats);
  dest.setValues(values);
  dest.setBackgrounds(rg.getBackgrounds());
  dest.setFontColors(rg.getFontColors());
  dest.setFontLines(rg.getFontLines());
  dest.setFontWeights(rg.getFontWeights());
  dest.setFontStyles(rg.getFontStyles());
  dest.setFontSizes(rg.getFontSizes());
  dest.setHorizontalAlignments(rg.getHorizontalAlignments());
  dest.setVerticalAlignments(rg.getVerticalAlignments());
  dest.setWrapStrategies(rg.getWrapStrategies());
  dest.setDataValidations(rg.getDataValidations());
  rg.getMergedRanges().forEach(function (m) {
    try { sh.getRange(m.getRow() + top - 1, m.getColumn(), m.getNumRows(), m.getNumColumns()).merge(); } catch (eM) {}
  });
  for (var c = 1; c <= lc; c++) sh.setColumnWidth(c, tab.getColumnWidth(c));
  sh.setFrozenRows(top + 1);

  unpaidMirrorStats_(sh, values);
  sh.getRange('B4').setValue('更新 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm')).setFontSize(8).setFontColor('#7A7A7A');
  props.setProperty('UNPAID_MIRROR_AT', String(now));
  return { ok: true, month: b1, rows: lr };
}

/** 集計（D1:M3）。会員番号あり・支払額>0 の行。回収金額が空でも右3列に「入金」があれば回収済みとみなす */
function unpaidMirrorStats_(sh, d) {
  var head = [];
  for (var c = 0; c < d[0].length; c++) head.push([0, 1, 2].map(function (r) { return d[r] ? String(d[r][c]) : ''; }).join(''));
  var fc = function (re) { for (var i = 0; i < head.length; i++) if (re.test(head[i])) return i; return -1; };
  var num = function (v) { var n = Number(String(v).replace(/[¥,\s]/g, '')); return isNaN(n) ? 0 : n; };
  var cMem = fc(/会員番号/), cPay = fc(/支払額/), cTot = fc(/総額/), cRec = fc(/回収金額|入金金額|レジ打ち金額/);
  if (cMem < 0 || cPay < 0) return;
  var cat = '';
  var s = { n: 0, sp: 0, st: 0, sr: 0, nr: 0 };
  var g = { one: { n: 0, p: 0, r: 0 }, two: { n: 0, p: 0, r: 0 }, bad: { n: 0, p: 0, r: 0 }, jac: { n: 0, p: 0, r: 0 } };
  d.forEach(function (row) {
    if (String(row[0] || '') !== '') cat = String(row[0]);
    var mem = String(row[cMem] || '').trim();
    var pay = num(row[cPay]);
    if (!mem || mem === '会員番号' || mem === '合計' || pay <= 0) return;
    var rec = cRec >= 0 ? num(row[cRec]) : 0;
    if (!rec && cRec >= 0 && /入金/.test(String(row[cRec + 1]) + String(row[cRec + 2]) + String(row[cRec + 3]))) rec = pay;
    s.n++; s.sp += pay; s.st += cTot >= 0 ? num(row[cTot]) : 0; s.sr += rec; if (rec > 0) s.nr++;
    var add = function (k) { g[k].n++; g[k].p += pay; g[k].r += rec; };
    if (/1[ヶヵカか]月/.test(cat)) add('one');
    if (/2[ヶヵカか]月/.test(cat)) add('two');
    if (/貸倒|貸し倒/.test(cat)) add('bad');
    if (/JACCS/.test(cat)) add('jac');
  });
  var yen = function (x) { return '¥' + Math.round(x).toLocaleString('ja-JP'); };
  var rt = function (k) { return g[k].n && g[k].p ? (g[k].r / g[k].p * 100).toFixed(1) + '%' : '対象なし'; };
  var sub = function (k) { return g[k].n ? yen(g[k].r) + ' / ' + yen(g[k].p) : '対象なし'; };
  sh.getRange('D1:M3').setValues([
    ['未納件数', '未納総額', '回収額', '回収率', '未回収額', '回収済み', '1ヶ月未納 回収率', '2ヶ月未納 回収率', '貸倒候補 回収率', 'JACCS 回収率'],
    [s.n + '件', yen(s.sp), yen(s.sr), s.sp ? (s.sr / s.sp * 100).toFixed(1) + '%' : '-', yen(s.sp - s.sr), s.nr + '件', rt('one'), rt('two'), rt('bad'), rt('jac')],
    ['支払額ベース', '手数料込 ' + yen(s.st), '回収金額の合計', '回収額÷未納総額', '残り ' + (s.n - s.nr) + '件', s.n ? Math.round(s.nr / s.n * 100) + '%（人数）' : '', sub('one'), sub('two'), sub('bad'), sub('jac')]
  ]);
}

function kiyakuLayout_(sh) {
  var INK = '#111111', MUTE = '#7A7A7A', RED = '#B91C1C';
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.setTabColor(RED);
  sh.getRange('A1').setValue('規約退会リスト').setFontSize(16).setFontWeight('bold');
  sh.getRange('B1').setFontSize(14).setFontWeight('bold').setFontColor(RED);
  sh.getRange('D1').setFontSize(9).setFontColor(MUTE);
  sh.getRange('A2').setValue('未納管理ドライブで会員名が赤の人（2ヶ月未納で強制退会）。1時間ごとに自動更新。元のシートは読むだけで変更しません。「退会の月」は会員動向で退会に数える月です。')
    .setFontSize(9).setFontColor(MUTE);
  sh.getRange(3, 1, 1, 11).setValues([['強制退会日', '退会の月', '会員番号', '会員名', '電話番号', '未納対象月', '未納理由', '未納額', '未納管理のシート', '累計退会データ', '退会年月（自動）']])
    .setBackground(INK).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(3);
  sh.setRowHeight(1, 34);
  sh.setRowHeight(3, 28);
  var widths = [100, 90, 110, 150, 120, 140, 100, 90, 120, 170, 100];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A4:A500').setNumberFormat('yyyy/m/d');
  sh.getRange('C4:C500').setNumberFormat('0');
  sh.getRange('K4:K500').setNumberFormat('yyyy/m');
  sh.getRange('A4:K500').setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange('D4:D500').setFontWeight('bold').setFontColor(RED);
  sh.getRange('A4:K500').setBorder(null, null, null, null, null, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
  var rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('未反映').setFontColor(RED).setBold(true)
    .setRanges([sh.getRange('J4:J500')]).build());
  sh.setConditionalFormatRules(rules);
  try { sh.hideColumns(11); } catch (eH) {}
}

/** 会員動向（本部シート）9月の解除 49 → 50（9月末の規約退会 落合 悠野を含める）。1回だけ */
var HQ_TREND_ID_ = '1LOOUG97wuiKbhzl0BjJstXgLaaSCZAKNFdD8P3I5x_o';

function fixHqSeptLeave_() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('HQ_SEPT_LEAVE_FIX') === 'v1') return { ok: true, skipped: true };
  var sh = SpreadsheetApp.openById(HQ_TREND_ID_).getSheetByName('経堂');
  if (!sh) return { ok: false };
  var cell = sh.getRange('H9');
  if (Number(cell.getValue()) === 49) cell.setValue(50);
  props.setProperty('HQ_SEPT_LEAVE_FIX', 'v1');
  return { ok: true, value: cell.getValue() };
}

/**
 * 会員動向 31〜46行（オプション月初）4〜9月を各月OP表の「前月末 − 当月1日解約」に統一（10月の 211 と同じ数え方）。
 * 6〜9月は当月開始の無料オプション込みの数字が入っていた。元の数字はセルのメモに残す。1回だけ
 */
var HQ_OPT_HISTORY_ = {
  4: [24, 180, 49, 16, 45, 2, 3, 207, 102, 49, 43, 86, 28, 114, 80, 17],
  5: [24, 209, 73, 26, 61, 2, 3, 230, 132, 73, 43, 114, 25, 113, 146, 20],
  6: [24, 204, 65, 26, 64, 2, 3, 218, 130, 62, 43, 103, 26, 111, 139, 18],
  7: [23, 199, 59, 26, 55, 2, 3, 215, 154, 59, 41, 98, 27, 113, 97, 18],
  8: [23, 193, 55, 24, 54, 2, 3, 207, 134, 55, 40, 95, 25, 124, 91, 18],
  9: [23, 196, 55, 24, 56, 2, 3, 213, 137, 57, 39, 94, 24, 125, 96, 19],
  10: [23, 211, 72, 27, 66, 2, 3, 228, 153, 74, 39, 111, 26, 132, 112, 21]
};
function fixHqOptionHistory_() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('HQ_OPT_HISTORY_FIX') === 'v2') return { ok: true, skipped: true };
  var sh = SpreadsheetApp.openById(HQ_TREND_ID_).getSheetByName('経堂');
  if (!sh) return { ok: false };
  var changed = 0;
  Object.keys(HQ_OPT_HISTORY_).forEach(function (k) {
    var month = Number(k);
    var col = ((month + 8) % 12) + 3;
    if (Number(sh.getRange(1, col).getValue()) !== month) return;
    var rg = sh.getRange(31, col, 16, 1);
    var cur = rg.getValues();
    var notes = rg.getNotes();
    var next = HQ_OPT_HISTORY_[k];
    for (var i = 0; i < 16; i++) {
      if (Number(cur[i][0]) === next[i]) continue;
      notes[i][0] = (notes[i][0] ? notes[i][0] + '\n' : '') + '10/3 ' + (month === 10 ? '9月OP表で修正（当月末−次月解約）' : month + '月OP表で修正（前月末−当月1日解約）') + '。修正前 ' + cur[i][0];
      cur[i][0] = next[i];
      changed++;
    }
    rg.setValues(cur);
    rg.setNotes(notes);
  });
  props.setProperty('HQ_OPT_HISTORY_FIX', 'v2');
  return { ok: true, changed: changed };
}

/**
 * 有料オプション検討：会員動向の「有料ＯＰ計画」（72〜88行）と月初実績（31〜46行）を並べ、
 * 各月OP表の「当月契約開始」（入会時に付く特典）と「翌月1日解約」を横に置く。
 */
var OPT_PLAN_SHEET_ = '有料オプション検討';
var OPT_DATA_SHEET_ = 'OP表データ';
var OPT_PLAN_V_ = 'v1';
var OPT_PLAN_ITEMS_ = [
  ['安心サポート', 73, 31, '単独'], ['安心サポートVIP', 74, 32, '入会パック'], ['水素水', 75, 44, '単独'],
  ['オンラインレッスン', 76, 33, '入会パック'], ['グループリフォーマー', 77, 46, '単独'], ['セルフエステ', 78, 34, '入会パック'],
  ['タンニング', 79, 35, '入会パック'], ['プロテイン12杯', 80, 36, '単独'], ['プロテイン飲み放題', 81, 37, '単独'],
  ['ホットスタジオ', 82, 38, '入会パック'], ['ボディプランナー', 83, 39, '入会パック'], ['マットレンタル', 84, 40, '入会パック'],
  ['ヨガマット契約ロッカー', 85, 41, '単独'], ['レンタルタオル', 86, 42, '入会パック'], ['契約ロッカー', 87, 43, '単独'],
  ['水素水+プロテイン(6)', 88, 45, '入会パック']
];
var OPT_FLOW_ = {
  4: [[0, 127, 9, 126, 4, 39, 89, 0, 0, 128, 127, 127, 0, 127, 6, 162], [0, 98, 10, 102, 1, 29, 73, 0, 0, 105, 97, 103, 0, 99, 9, 96]],
  5: [[0, 73, 4, 70, 0, 24, 52, 0, 0, 73, 73, 71, 0, 70, 2, 70], [0, 78, 6, 78, 2, 24, 49, 0, 0, 85, 75, 82, 0, 81, 1, 78]],
  6: [[0, 58, 5, 58, 1, 22, 36, 0, 0, 58, 96, 57, 0, 58, 2, 59], [1, 63, 3, 64, 1, 22, 45, 0, 0, 61, 72, 60, 2, 63, 1, 100]],
  7: [[0, 68, 16, 65, 2, 21, 49, 0, 0, 65, 67, 65, 1, 65, 1, 65], [0, 74, 5, 69, 2, 23, 50, 0, 0, 73, 87, 69, 2, 68, 3, 71]],
  8: [[0, 92, 6, 91, 4, 38, 58, 0, 0, 95, 97, 92, 0, 92, 2, 93], [0, 89, 5, 91, 3, 38, 56, 0, 0, 89, 94, 90, 1, 93, 3, 88]],
  9: [[0, 87, 11, 86, 3, 25, 63, 0, 0, 87, 90, 87, 1, 86, 2, 89], [0, 72, 4, 69, 1, 22, 53, 0, 0, 72, 75, 70, 1, 69, 0, 73]],
  10: [[0, 15, 1, 14, 5, 4, 10, 0, 0, 15, 14, 14, 0, 14, 1, 15], null]
};

function ensureOptionPlanSheet_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('OPT_PLAN_V') === OPT_PLAN_V_ && ss.getSheetByName(OPT_PLAN_SHEET_)) return { ok: true, skipped: true };
  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#D9D9D9', RED = '#B91C1C', HEAD = '#F3F3F3';

  var data = ss.getSheetByName(OPT_DATA_SHEET_) || ss.insertSheet(OPT_DATA_SHEET_);
  data.clear();
  var rows = [['月', '項目', '当月契約開始', '翌月1日解約']];
  Object.keys(OPT_FLOW_).forEach(function (m) {
    OPT_PLAN_ITEMS_.forEach(function (it, i) {
      var f = OPT_FLOW_[m];
      rows.push([Number(m), it[0], f[0][i], f[1] ? f[1][i] : '']);
    });
  });
  data.getRange(1, 1, rows.length, 4).setValues(rows);
  data.hideSheet();

  var sh = ss.getSheetByName(OPT_PLAN_SHEET_) || ss.insertSheet(OPT_PLAN_SHEET_);
  sh.clear();
  sh.setConditionalFormatRules([]);
  var hq = "'【経堂】会員動向'!";
  sh.getRange('A1').setValue('有料オプション検討').setFontSize(16).setFontWeight('bold').setFontColor(INK);
  sh.getRange('A2').setValue('月').setFontColor(MUTE);
  sh.getRange('B2').setNumberFormat('@').setValue('10').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['4', '5', '6', '7', '8', '9', '10', '11', '12', '1', '2', '3'], true).build());
  sh.getRange('C2').setValue('← 月を選ぶと切り替わります（目標・月初は会員動向、新規開始・解約は各月のOP表）').setFontColor(MUTE).setFontSize(9);

  var head = ['項目', '区分', '目標（有料OP計画）', '月初実績', '差', '達成率', '当月 新規開始', '翌月1日 解約', '解約 ÷ 新規開始'];
  sh.getRange(4, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground(HEAD).setFontColor(INK)
    .setHorizontalAlignment('center').setWrap(true);
  var mcol = 'MATCH(VALUE($B$2),' + hq + '$C$1:$N$1,0)';
  var body = OPT_PLAN_ITEMS_.map(function (it, i) {
    var r = 5 + i;
    return [
      it[0], it[3],
      '=IFERROR(VALUE(TRIM(INDEX(' + hq + '$C$' + it[1] + ':$N$' + it[1] + ',1,' + mcol + '))),"")',
      '=IFERROR(INDEX(' + hq + '$C$' + it[2] + ':$N$' + it[2] + ',1,' + mcol + '),"")',
      '=IF(OR(C' + r + '="",D' + r + '=""),"",D' + r + '-C' + r + ')',
      '=IFERROR(D' + r + '/C' + r + ',"")',
      '=IFERROR(INDEX(FILTER(\'' + OPT_DATA_SHEET_ + '\'!C:C,\'' + OPT_DATA_SHEET_ + '\'!A:A=VALUE($B$2),\'' + OPT_DATA_SHEET_ + '\'!B:B=A' + r + '),1),"")',
      '=IFERROR(INDEX(FILTER(\'' + OPT_DATA_SHEET_ + '\'!D:D,\'' + OPT_DATA_SHEET_ + '\'!A:A=VALUE($B$2),\'' + OPT_DATA_SHEET_ + '\'!B:B=A' + r + '),1),"")',
      '=IFERROR(IF(OR(G' + r + '="",H' + r + '="",G' + r + '=0),"",H' + r + '/G' + r + '),"")'
    ];
  });
  var last = 4 + body.length;
  sh.getRange(5, 1, body.length, head.length).setFormulas(body.map(function (r) { return r.map(String); }));
  sh.getRange(5, 1, body.length, 2).setValues(OPT_PLAN_ITEMS_.map(function (it) { return [it[0], it[3]]; }));
  sh.getRange(5, 2, body.length, 1).setFontColor(MUTE).setHorizontalAlignment('center');
  sh.getRange(5, 3, body.length, 3).setNumberFormat('#,##0;[Red]-#,##0').setHorizontalAlignment('right');
  sh.getRange(5, 4, body.length, 1).setFontWeight('bold');
  sh.getRange(5, 6, body.length, 1).setNumberFormat('0%').setHorizontalAlignment('right');
  sh.getRange(5, 7, body.length, 2).setNumberFormat('#,##0').setHorizontalAlignment('right');
  sh.getRange(5, 9, body.length, 1).setNumberFormat('0%').setHorizontalAlignment('right');
  var tr = last + 1;
  sh.getRange(tr, 1, 1, head.length).setValues([['合計', '', '=SUM(C5:C' + last + ')', '=SUM(D5:D' + last + ')', '=D' + tr + '-C' + tr, '=IFERROR(D' + tr + '/C' + tr + ',"")', '=SUM(G5:G' + last + ')', '=SUM(H5:H' + last + ')', '=IFERROR(H' + tr + '/G' + tr + ',"")']])
    .setFontWeight('bold').setBackground(HEAD);
  sh.getRange(tr, 3, 1, 3).setNumberFormat('#,##0;[Red]-#,##0');
  sh.getRange(tr, 6).setNumberFormat('0%');
  sh.getRange(tr, 9).setNumberFormat('0%');
  sh.getRange(4, 1, tr - 3, head.length).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Arial');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(RED).setRanges([sh.getRange(5, 5, body.length + 1, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.8).setFontColor(RED).setRanges([sh.getRange(5, 9, body.length + 1, 1)]).build()
  ]);

  var notes = [
    ['見かた'],
    ['・月初実績＝前月末の契約数 − 当月1日の解約数（その月に入会した人の特典分は入っていない）'],
    ['・当月 新規開始＝その月に始まった契約。「入会パック」の項目は入会者ほぼ全員に付いていて、入会特典（無料）とみられる'],
    ['・翌月1日 解約＝その月の翌月1日に解約になった数。新規開始とほぼ同じ数なら、特典のまま外れている'],
    [''],
    ['一緒に決めたいこと'],
    ['① 入会特典の無料期間は「入会月だけ」か「翌月まで」か（翌月までなら、月初実績にもまだ無料の人が混ざる）'],
    ['② 有料OP計画の数字は「月初実績」と比べてよいか（＝月初時点で有料の契約数として作られた目標か）'],
    ['③ 1人ずつ有料／無料を分けたい場合、CASIOで出せる明細（契約開始日・金額つきのオプション契約一覧）があるか']
  ];
  sh.getRange(tr + 2, 1, notes.length, 1).setValues(notes).setFontColor(MUTE).setFontSize(9);
  sh.getRange(tr + 2, 1).setFontColor(INK).setFontWeight('bold').setFontSize(10);
  sh.getRange(tr + 7, 1).setFontColor(INK).setFontWeight('bold').setFontSize(10);

  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 90);
  for (var c = 3; c <= head.length; c++) sh.setColumnWidth(c, 105);
  sh.setRowHeight(4, 36);
  sh.setFrozenRows(4);
  sh.setHiddenGridlines(true);
  props.setProperty('OPT_PLAN_V', OPT_PLAN_V_);
  return { ok: true };
}

/**
 * 口コミ管理：未対応の口コミを一覧にして、ここでエンジョイ付与をチェックする。
 * ① EAST付与は口コミ_経堂の「ポイント付与済」（付与アプリが書く）をそのまま表示。
 * ② エンジョイ付与・メモはこのシートで入力し、隠しシート「口コミ付与記録」に残す。両方済で一覧から消える。
 * 10/3 より前に EAST付与済みの口コミは完了扱い。
 */
var REVIEW_TODO_SHEET_ = '口コミ管理';
var REVIEW_LOG_SHEET_ = '口コミ付与記録';
var REVIEW_TODO_ROW_ = 17;
var REVIEW_GBP_URL_ = 'https://business.google.com/reviews';
var REVIEW_ENJOY_SINCE_ = new Date(2026, 9, 3);

function reviewKey_(row) {
  if (row[15]) return String(row[15]);
  var t = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyyMMddHHmm') : String(row[0]);
  return t + '_' + String(row[5]);
}

function reviewTruthy_(v) {
  return v === true || /^(true|☑|済)$/i.test(String(v).trim());
}

function reviewLog_(ss) {
  var log = ss.getSheetByName(REVIEW_LOG_SHEET_);
  if (!log) {
    log = ss.insertSheet(REVIEW_LOG_SHEET_);
    log.getRange(1, 1, 1, 5).setValues([['key', 'エンジョイ付与', '付与日', 'メモ', '氏名']]);
    log.hideSheet();
  }
  var map = {};
  if (log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 5).getValues().forEach(function (r, i) {
      if (r[0]) map[String(r[0])] = { row: i + 2, enjoy: r[1] === true, date: r[2], memo: r[3], name: r[4] };
    });
  }
  return { sheet: log, map: map };
}

function reviewLogPut_(log, key, enjoy, date, memo, name) {
  var cur = log.map[key];
  var vals = [[key, enjoy === true, date || '', memo || '', name || '']];
  if (cur) log.sheet.getRange(cur.row, 1, 1, 5).setValues(vals);
  else {
    log.sheet.appendRow(vals[0]);
    log.map[key] = { row: log.sheet.getLastRow() };
  }
  log.map[key].enjoy = enjoy === true;
  log.map[key].date = date;
  log.map[key].memo = memo;
}

function syncReviewTodo_(ss) {
  var sh = ss.getSheetByName(REVIEW_TODO_SHEET_);
  var src = ss.getSheetByName('口コミ_経堂');
  if (!sh || !src) return { ok: false };
  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#D9D9D9', RED = '#B91C1C', HEAD = '#111111';
  var log = reviewLog_(ss);

  var maxR = Math.max(sh.getMaxRows(), REVIEW_TODO_ROW_);
  if (maxR >= REVIEW_TODO_ROW_) {
    var cur = sh.getRange(REVIEW_TODO_ROW_, 1, maxR - REVIEW_TODO_ROW_ + 1, 12).getValues();
    cur.forEach(function (r) {
      var key = String(r[11] || '');
      if (!key) return;
      var old = log.map[key] || {};
      if (old.enjoy !== (r[6] === true) || String(old.memo || '') !== String(r[8] || '')) {
        reviewLogPut_(log, key, r[6] === true, r[6] === true ? (r[7] || new Date()) : '', r[8], r[1]);
      }
    });
  }

  var last = src.getLastRow();
  var rows = last >= 3 ? src.getRange(3, 1, last - 2, 23).getValues() : [];
  var now = new Date();
  var ym = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');
  var list = [];
  rows.forEach(function (r) {
    if (!r[4] || !(r[0] instanceof Date)) return;
    if (Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyyMM') !== ym) return;
    var key = reviewKey_(r);
    var lg = log.map[key] || {};
    list.push([r[0], r[4], String(r[5] || ''), r[3], r[9], '', lg.enjoy === true, lg.date || '', lg.memo || '', '', '', key]);
  });
  list.sort(function (a, b) { return b[0].getTime() - a[0].getTime(); });

  var props = PropertiesService.getDocumentProperties();
  var sig = 'v2|' + ym + '|' + list.map(function (t) { return t[11]; }).join(',');
  if (props.getProperty('REVIEW_TODO_SIG') === sig && String(sh.getRange(16, 1).getValue()) === '口コミ日') return { ok: true, skipped: true };

  try { src.hideSheet(); } catch (eH) {}
  var b9 = sh.getRange('B9');
  if (/#gid=/.test(String(b9.getFormula() || ''))) b9.setValue('下の一覧 ↓');

  var area = sh.getRange(13, 1, Math.max(sh.getMaxRows() - 12, 1), 12);
  area.clearContent();
  area.clearFormat();
  area.clearDataValidations();
  sh.setConditionalFormatRules([]);
  var month = Number(ym.slice(4));
  sh.getRange('A12:C12').copyTo(sh.getRange('A13:C13'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.getRange('A13:C13').setValues([['5. Googleの口コミに返信', '', 'Googleビジネスプロフィールで口コミに返信する']]);
  sh.getRange('B13').setFormula('=HYPERLINK("' + REVIEW_GBP_URL_ + '","Google口コミに返信 ↗")');
  sh.getRange('A14').setValue('流れ：① 口コミ付与アプリでEAST付与（元の回答シートでチェックされると自動で ✓ 済）→ ② エンジョイ付与をしたら G にチェック → 両方済で灰色になります').setFontColor(MUTE).setFontSize(9);
  var first = REVIEW_TODO_ROW_, lastRow = REVIEW_TODO_ROW_ + Math.max(list.length, 1) - 1;
  sh.getRange('A15').setFormula('="今月の口コミ（' + month + '月）　' + list.length + '件　未対応 "&SUMPRODUCT((L' + first + ':L' + lastRow + '<>"")*(((F' + first + ':F' + lastRow + '<>"✓ 済")+(G' + first + ':G' + lastRow + '<>TRUE))>0))&"件"')
    .setFontWeight('bold').setFontSize(12).setFontColor(INK);
  var head = ['口コミ日', '氏名', '会員番号', '評価', '来店日', '① EAST付与', '② エンジョイ付与', '付与日', 'メモ'];
  sh.getRange(16, 1, 1, head.length).setValues([head]).setBackground(HEAD).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  if (list.length) {
    var need = REVIEW_TODO_ROW_ + list.length - 1;
    if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
    sh.getRange(REVIEW_TODO_ROW_, 1, list.length, 12).setValues(list);
    sh.getRange(REVIEW_TODO_ROW_, 6, list.length, 1).setFormulas(list.map(function (t, i) {
      var r = REVIEW_TODO_ROW_ + i;
      return ['=IF(SUMPRODUCT((\'口コミ_経堂\'!$P$3:$P="' + t[11] + '")*((\'口コミ_経堂\'!$V$3:$V=TRUE)+(\'口コミ_経堂\'!$V$3:$V="TRUE")))+IFERROR(COUNTIFS(\'口コミ_経堂\'!$A$3:$A,A' + r + ',\'口コミ_経堂\'!$F$3:$F,C' + r + ',\'口コミ_経堂\'!$V$3:$V,TRUE),0)>0,"✓ 済","未")'];
    }));
    sh.getRange(REVIEW_TODO_ROW_, 1, list.length, 1).setNumberFormat('M/d');
    sh.getRange(REVIEW_TODO_ROW_, 5, list.length, 1).setNumberFormat('M/d');
    sh.getRange(REVIEW_TODO_ROW_, 8, list.length, 1).setNumberFormat('M/d');
    sh.getRange(REVIEW_TODO_ROW_, 3, list.length, 1).setNumberFormat('@');
    sh.getRange(REVIEW_TODO_ROW_, 7, list.length, 1).insertCheckboxes();
    sh.getRange(REVIEW_TODO_ROW_, 1, list.length, 8).setHorizontalAlignment('center');
    sh.getRange(REVIEW_TODO_ROW_, 2, list.length, 1).setHorizontalAlignment('left').setFontWeight('bold');
    sh.getRange(REVIEW_TODO_ROW_, 1, list.length, head.length).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID).setFontColor(INK);
    sh.getRange(REVIEW_TODO_ROW_, 12, list.length, 1).setFontColor('#FFFFFF').setFontSize(6);
    var rowsRg = sh.getRange(REVIEW_TODO_ROW_, 1, list.length, head.length);
    var fRg = sh.getRange(REVIEW_TODO_ROW_, 6, list.length, 1);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($F' + REVIEW_TODO_ROW_ + '="✓ 済",$G' + REVIEW_TODO_ROW_ + '=TRUE)')
        .setFontColor('#B0B0B0').setStrikethrough(true).setBackground('#F5F5F5').setRanges([rowsRg]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('未').setFontColor(RED).setBold(true).setRanges([fRg]).build()
    ]);
  } else {
    sh.getRange(REVIEW_TODO_ROW_, 1).setValue('今月の口コミはまだありません').setFontColor(MUTE);
  }
  props.setProperty('REVIEW_TODO_SIG', sig);
  return { ok: true, rows: list.length };
}

function reviewTodoOnEdit_(e) {
  var sh = e.range.getSheet();
  var r0 = e.range.getRow(), c0 = e.range.getColumn();
  if (r0 < REVIEW_TODO_ROW_ || (c0 !== 7 && c0 !== 9)) return false;
  var ss = sh.getParent();
  var log = reviewLog_(ss);
  var n = e.range.getNumRows();
  var vals = sh.getRange(r0, 1, n, 12).getValues();
  vals.forEach(function (r, i) {
    var key = String(r[11] || '');
    if (!key) return;
    var enjoy = r[6] === true;
    var date = enjoy ? (r[7] || new Date()) : '';
    if (enjoy && !r[7]) sh.getRange(r0 + i, 8).setValue(date).setNumberFormat('M/d');
    if (!enjoy && r[7]) sh.getRange(r0 + i, 8).clearContent();
    reviewLogPut_(log, key, enjoy, date, r[8], r[1]);
  });
  return true;
}

/**
 * 入会・退会分析（会議用）：累計入会データ・累計退会データから数式で集計。B2 の基準月で全部切り替わる。
 * 退会の月＝その月末で辞めた人（CASIOの退会年月が翌月）。会員動向の「解除」と同じ数え方。
 */
var JL_SHEET_ = '入会・退会分析';
var JL_V_ = 'v4';

function ensureJoinLeaveAnalysis_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (ss.getSheetByName(JL_SHEET_) && props.getProperty('JL_V')) return migrateJoinLeaveSheet_(ss);
  var join = ss.getSheetByName('累計入会データ');
  var leave = ss.getSheetByName('累計退会データ');
  if (!join || !leave) return { ok: false };
  join.getRange('Y1').setValue('プラン（自動）');
  join.getRange('Y2').setFormula('=SCAN("",C2:C8000,LAMBDA(a,c,IF(c="",a,c)))');
  join.getRange('Z1').setValue('在籍月数（自動）');
  join.getRange('Z2').setFormula('=ARRAYFORMULA(IF(V2:V8000=0,"",IF(W2:W8000="",999,(YEAR(W2:W8000)*12+MONTH(W2:W8000))-(YEAR(V2:V8000)*12+MONTH(V2:V8000)))))');
  leave.getRange('V1').setValue('プラン（自動）');
  leave.getRange('V2').setFormula('=SCAN("",C2:C8000,LAMBDA(a,c,IF(c="",a,c)))');

  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#D9D9D9', HEAD = '#F3F3F3', RED = '#B91C1C';
  var sh = ss.getSheetByName(JL_SHEET_) || ss.insertSheet(JL_SHEET_);
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.clear();
  sh.setConditionalFormatRules([]);
  if (sh.getMaxColumns() < 16) sh.insertColumnsAfter(sh.getMaxColumns(), 16 - sh.getMaxColumns());
  if (sh.getMaxRows() < 90) sh.insertRowsAfter(sh.getMaxRows(), 90 - sh.getMaxRows());

  var J = "'累計入会データ'!", L = "'累計退会データ'!";
  var jV = J + '$V$2:$V$8000', jW = J + '$W$2:$W$8000', jX = J + '$X$2:$X$8000', jI = J + '$I$2:$I$8000', jY = J + '$Y$2:$Y$8000';
  var lU = L + '$U$2:$U$8000', lP = L + '$P$2:$P$8000', lJ = L + '$J$2:$J$8000', lI = L + '$I$2:$I$8000', lV = L + '$V$2:$V$8000';

  var section = function (row, text) {
    sh.getRange(row, 1).setValue(text).setFontSize(13).setFontWeight('bold').setFontColor(INK);
  };
  var header = function (row, labels) {
    sh.getRange(row, 1, 1, labels.length).setValues([labels]).setBackground(INK).setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  };
  var box = function (row, n, cols) {
    sh.getRange(row, 1, n, cols).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  };

  sh.getRange('A1').setValue('入会・退会分析').setFontSize(18).setFontWeight('bold').setFontColor(INK);
  sh.getRange('A2').setValue('基準月').setFontColor(MUTE).setHorizontalAlignment('right');
  sh.getRange('B2').setFormula('=DATE(YEAR(TODAY()),MONTH(TODAY()),1)').setNumberFormat('yyyy"年"m"月"')
    .setFontWeight('bold').setFontSize(12).setBackground('#FFF7E6').setHorizontalAlignment('center');
  sh.getRange('C2').setValue('← 例「2026/9/1」と入れるとその月で全部切り替わります（空にすると今月に戻ります）').setFontColor(MUTE).setFontSize(9);
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
  sh.getRange('P1').setFormula('=IF(ISNUMBER($B$2),DATE(YEAR($B$2),MONTH($B$2),1),DATE(YEAR(TODAY()),MONTH(TODAY()),1))');
  var BM = '$P$1';

  // ① 12ヶ月の推移
  section(4, '① 12ヶ月の推移');
  sh.getRange('C4').setValue('退会＝その月末で辞めた人（会員動向の「解除」と同じ）。日報の月の入会・退会は日報（当月入会・当月末退会）、それ以外は累計データから').setFontColor(MUTE).setFontSize(9);
  header(5, ['月', '月初会員', '入会', '退会', '純増', '退会率', '規約退会（参考）']);
  var t = [];
  for (var i = 0; i < 12; i++) {
    var r = 6 + i;
    t.push([
      '=EDATE(' + BM + ',' + (i - 11) + ')',
      '=COUNTIFS(' + jV + ',">0",' + jV + ',"<"&A' + r + ')-COUNTIFS(' + lU + ',"<="&A' + r + ')',
      '=IF(TEXT(A' + r + ',"yymm")=TEXT(\'日報\'!$B$1,"0"),N(\'日報\'!$C$13),COUNTIF(' + jV + ',A' + r + '))',
      '=IF(TEXT(A' + r + ',"yymm")=TEXT(\'日報\'!$B$1,"0"),N(\'日報\'!$C$15),COUNTIF(' + lU + ',EDATE(A' + r + ',1)))',
      '=C' + r + '-D' + r,
      '=IFERROR(D' + r + '/B' + r + ',"")',
      '=COUNTIF(\'規約退会リスト\'!$K$4:$K$500,EDATE(A' + r + ',1))'
    ]);
  }
  sh.getRange(6, 1, 12, 7).setFormulas(t);
  sh.getRange(18, 1, 1, 7).setFormulas([['="12ヶ月 計"', '=AVERAGE(B6:B17)', '=SUM(C6:C17)', '=SUM(D6:D17)', '=C18-D18', '=IFERROR(AVERAGE(F6:F17),"")', '=SUM(G6:G17)']])
    .setFontWeight('bold').setBackground(HEAD);
  sh.getRange('A6:A17').setNumberFormat('yyyy/m');
  sh.getRange('B6:E18').setNumberFormat('#,##0;[Red]-#,##0');
  sh.getRange('B18').setNumberFormat('#,##0" (平均)"');
  sh.getRange('F6:F18').setNumberFormat('0.0%');
  sh.getRange('F18').setNumberFormat('0.0%" (平均)"');
  sh.getRange('G6:G18').setNumberFormat('0');
  sh.getRange('A17:G17').setFontWeight('bold');
  box(5, 14, 7);

  // ② 退会した人の内訳
  section(21, '② 退会した人の内訳');
  sh.getRange('C21').setValue('直近12ヶ月（①の期間）と基準月だけの人数').setFontColor(MUTE).setFontSize(9);
  header(22, ['区分', '項目', '直近12ヶ月', '割合', '基準月']);
  var isD = 'ISNUMBER(' + lU + ')';
  var win12 = '(' + isD + '*(IFERROR(' + lU + '*1,0)>=EDATE($A$6,1))*(IFERROR(' + lU + '*1,0)<=EDATE($A$17,1)))';
  var winM = '(' + isD + '*(IFERROR(' + lU + '*1,0)=EDATE($A$17,1)))';
  var P = 'IFERROR(VALUE(' + lP + '),-1)', A = 'IFERROR(VALUE(' + lJ + '),-1)';
  var leaveRows = [
    ['在籍期間', '3ヶ月以内', '(' + P + '>=0)*(' + P + '<=3)'],
    ['', '4〜6ヶ月', '(' + P + '>=4)*(' + P + '<=6)'],
    ['', '7〜12ヶ月', '(' + P + '>=7)*(' + P + '<=12)'],
    ['', '1〜2年', '(' + P + '>=13)*(' + P + '<=24)'],
    ['', '2年超', '(' + P + '>=25)'],
    ['年代', '10代', '(' + A + '>=0)*(' + A + '<20)'],
    ['', '20代', '(' + A + '>=20)*(' + A + '<30)'],
    ['', '30代', '(' + A + '>=30)*(' + A + '<40)'],
    ['', '40代', '(' + A + '>=40)*(' + A + '<50)'],
    ['', '50代', '(' + A + '>=50)*(' + A + '<60)'],
    ['', '60代以上', '(' + A + '>=60)'],
    ['男女', '男', '(' + lI + '="男")'],
    ['', '女', '(' + lI + '="女")'],
    ['プラン', 'ナショナル会員U', '(' + lV + '="ナショナル会員U")'],
    ['', '法人個人月払B', '(' + lV + '="法人個人月払B")'],
    ['', '法人個人月払A', '(' + lV + '="法人個人月払A")'],
    ['', '閉店移籍会員3', '(' + lV + '="閉店移籍会員3")']
  ];
  var lr = leaveRows.map(function (x, k) {
    var r = 23 + k;
    return [x[0], x[1], '=SUMPRODUCT(' + win12 + '*' + x[2] + ')', '=IFERROR(C' + r + '/$P$2,"")', '=SUMPRODUCT(' + winM + '*' + x[2] + ')'];
  });
  sh.getRange('P2').setFormula('=SUMPRODUCT(' + win12 + ')');
  sh.getRange('P3').setFormula('=SUMPRODUCT(' + winM + ')');
  var otherRow = 23 + lr.length;
  lr.push(['', 'その他', '=$P$2-SUM(C' + (otherRow - 4) + ':C' + (otherRow - 1) + ')', '=IFERROR(C' + otherRow + '/$P$2,"")', '=$P$3-SUM(E' + (otherRow - 4) + ':E' + (otherRow - 1) + ')']);
  sh.getRange(23, 3, lr.length, 3).setFormulas(lr.map(function (x) { return [x[2], x[3], x[4]]; }));
  sh.getRange(23, 1, lr.length, 2).setValues(lr.map(function (x) { return [x[0], x[1]]; }));
  sh.getRange(23, 1, lr.length, 1).setFontWeight('bold').setFontColor(INK);
  sh.getRange(23, 4, lr.length, 1).setNumberFormat('0%');
  box(22, lr.length + 1, 5);
  [27, 33, 35].forEach(function (r) { sh.getRange(r, 1, 1, 5).setBorder(null, null, true, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID); });
  var leaveEnd = 22 + lr.length;

  // ③ 入会した月ごとの継続率
  var c0 = leaveEnd + 3;
  section(c0, '③ 入会した月ごとの継続率');
  sh.getRange(c0, 3).setValue('入会した人のうち、何ヶ月後もまだ在籍しているか（まだその月が来ていないところは空欄）').setFontColor(MUTE).setFontSize(9);
  header(c0 + 1, ['入会月', '入会数', '1ヶ月後', '3ヶ月後', '6ヶ月後', '12ヶ月後']);
  var ks = [1, 3, 6, 12];
  var co = [];
  for (var m = 0; m < 12; m++) {
    var r2 = c0 + 2 + m;
    var row = ['=EDATE(' + BM + ',' + (m - 12) + ')', '=COUNTIF(' + jV + ',A' + r2 + ')'];
    ks.forEach(function (k) {
      row.push('=IF(OR(B' + r2 + '=0,EDATE(A' + r2 + ',' + k + ')>' + BM + '),"",(COUNTIFS(' + jV + ',A' + r2 + ',' + jW + ',"")+COUNTIFS(' + jV + ',A' + r2 + ',' + jW + ',">"&EDATE(A' + r2 + ',' + k + ')))/B' + r2 + ')');
    });
    co.push(row);
  }
  sh.getRange(c0 + 2, 1, 12, 6).setFormulas(co);
  var avgRow = c0 + 14;
  var avg = ['="平均"', '=SUM(B' + (c0 + 2) + ':B' + (c0 + 13) + ')'];
  ['C', 'D', 'E', 'F'].forEach(function (col) {
    avg.push('=IFERROR(SUMPRODUCT(N(' + col + (c0 + 2) + ':' + col + (c0 + 13) + '),$B' + (c0 + 2) + ':$B' + (c0 + 13) + ')/SUMPRODUCT((' + col + (c0 + 2) + ':' + col + (c0 + 13) + '<>"")*$B' + (c0 + 2) + ':$B' + (c0 + 13) + '),"")');
  });
  sh.getRange(avgRow, 1, 1, 6).setFormulas([avg]).setFontWeight('bold').setBackground(HEAD);
  sh.getRange(c0 + 2, 1, 12, 1).setNumberFormat('yyyy/m');
  sh.getRange(c0 + 2, 3, 13, 4).setNumberFormat('0.0%');
  box(c0 + 1, 14, 6);
  var coRange = sh.getRange(c0 + 2, 3, 12, 4);

  // ④ 続きやすさの違い（6ヶ月継続率）
  var d0 = avgRow + 3;
  section(d0, '④ 続きやすさの違い（6ヶ月後も在籍している割合）');
  sh.getRange(d0, 3).setValue('6ヶ月たった入会者（基準月の18〜7ヶ月前に入会した人）で比べています').setFontColor(MUTE).setFontSize(9);
  header(d0 + 1, ['区分', '項目', '入会数', '6ヶ月後も在籍', '6ヶ月継続率']);
  var win = '(' + jV + '>=EDATE(' + BM + ',-18))*(' + jV + '<=EDATE(' + BM + ',-7))';
  var stay = '(' + J + '$Z$2:$Z$8000>6)';
  var segs = [
    ['全体', '全員', '1'],
    ['年代', '10代', '(' + jX + '<20)'],
    ['', '20代', '(' + jX + '>=20)*(' + jX + '<30)'],
    ['', '30代', '(' + jX + '>=30)*(' + jX + '<40)'],
    ['', '40代', '(' + jX + '>=40)*(' + jX + '<50)'],
    ['', '50代', '(' + jX + '>=50)*(' + jX + '<60)'],
    ['', '60代以上', '(' + jX + '>=60)'],
    ['男女', '男', '(' + jI + '="男")'],
    ['', '女', '(' + jI + '="女")'],
    ['プラン', 'ナショナル会員U', '(' + jY + '="ナショナル会員U")'],
    ['', '法人個人月払B', '(' + jY + '="法人個人月払B")'],
    ['', '法人個人月払A', '(' + jY + '="法人個人月払A")'],
    ['', '閉店移籍会員3', '(' + jY + '="閉店移籍会員3")']
  ];
  var sg = segs.map(function (x, k) {
    var r3 = d0 + 2 + k;
    return [x[0], x[1], '=SUMPRODUCT(' + win + '*' + x[2] + ')', '=SUMPRODUCT(' + win + '*' + x[2] + '*' + stay + ')', '=IFERROR(D' + r3 + '/C' + r3 + ',"")'];
  });
  sh.getRange(d0 + 2, 3, sg.length, 3).setFormulas(sg.map(function (x) { return [x[2], x[3], x[4]]; }));
  sh.getRange(d0 + 2, 1, sg.length, 2).setValues(sg.map(function (x) { return [x[0], x[1]]; }));
  sh.getRange(d0 + 2, 1, sg.length, 1).setFontWeight('bold');
  sh.getRange(d0 + 2, 1, 1, 5).setBackground(HEAD).setFontWeight('bold');
  sh.getRange(d0 + 2, 5, sg.length, 1).setNumberFormat('0.0%').setFontWeight('bold');
  box(d0 + 1, sg.length + 1, 5);
  var segRange = sh.getRange(d0 + 3, 5, sg.length - 1, 1);

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(C' + (c0 + 2) + '),C' + (c0 + 2) + '<C$' + avgRow + '-0.03)')
      .setFontColor(RED).setBold(true).setRanges([coRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(E' + (d0 + 3) + '),E' + (d0 + 3) + '<$E$' + (d0 + 2) + '-0.03)')
      .setFontColor(RED).setBold(true).setRanges([segRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(RED).setRanges([sh.getRange('G6:G18')]).build()
  ]);

  var notes = d0 + 2 + sg.length + 1;
  sh.getRange(notes, 1, 3, 1).setValues([
    ['赤字＝平均より3ポイント以上低いところ（③は各列の平均、④は全体と比べて）'],
    ['規約退会（参考）＝規約退会リストの人数。CASIOの退会データにまだ載っていない場合があるので「退会」とは別に表示'],
    ['プランは累計データの契約名称から判定。データは「累計入会データ」「累計退会データ」から自動で集計されます']
  ]).setFontColor(MUTE).setFontSize(9);

  sh.getRange(1, 1, notes + 3, 7).setFontFamily('Arial');
  sh.getRange(5, 2, notes, 6).setHorizontalAlignment('right');
  sh.getRange(22, 2, lr.length + 1, 1).setHorizontalAlignment('left');
  sh.getRange(d0 + 1, 2, sg.length + 1, 1).setHorizontalAlignment('left');
  sh.getRange('A5:G5').setHorizontalAlignment('center');
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 130);
  for (var cw = 3; cw <= 7; cw++) sh.setColumnWidth(cw, 100);
  sh.setColumnWidth(8, 24);
  sh.hideColumns(16);
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(2);

  var ch1 = sh.newChart().asComboChart()
    .addRange(sh.getRange('A5:A17')).addRange(sh.getRange('C5:D17')).addRange(sh.getRange('F5:F17'))
    .setOption('title', '入会・退会と退会率（12ヶ月）')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('series', { 0: { type: 'bars', color: '#111111' }, 1: { type: 'bars', color: '#B0B0B0' }, 2: { type: 'line', color: '#B91C1C', targetAxisIndex: 1 } })
    .setOption('vAxes', { 0: { title: '人数' }, 1: { title: '退会率', format: 'percent' } })
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 620).setOption('height', 300)
    .setPosition(4, 9, 0, 0).build();
  sh.insertChart(ch1);
  var ch2 = sh.newChart().asBarChart()
    .addRange(sh.getRange('B23:C27'))
    .setOption('title', '退会した人の在籍期間（直近12ヶ月）')
    .setOption('colors', ['#111111']).setOption('legend', { position: 'none' })
    .setOption('width', 620).setOption('height', 260)
    .setPosition(21, 9, 0, 0).build();
  sh.insertChart(ch2);
  var ch3 = sh.newChart().asLineChart()
    .addRange(sh.getRange(c0 + 1, 1, 13, 1)).addRange(sh.getRange(c0 + 1, 4, 13, 2))
    .setOption('title', '入会月ごとの3ヶ月・6ヶ月継続率')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('colors', ['#111111', '#B91C1C']).setOption('vAxis', { format: 'percent' })
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 620).setOption('height', 300)
    .setPosition(c0, 9, 0, 0).build();
  sh.insertChart(ch3);

  props.setProperty('JL_V', JL_V_);
  props.deleteProperty('JL_MIG');
  return migrateJoinLeaveSheet_(ss);
}

/**
 * 入会・退会分析の手直し（1回だけ）。グラフ・ユーザーが編集したところには触らない。
 * 灰色の説明文を消す、基準月をプルダウン（「今月（自動）」＋データのある月）に、文字をメイリオに。
 */
function migrateJoinLeaveSheet_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('JL_MIG') === 'v1') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(JL_SHEET_);
  if (!sh) return { ok: false };
  var rows = Math.min(sh.getMaxRows(), 120);
  var rg = sh.getRange(1, 1, rows, 8);
  var vals = rg.getValues();
  var colors = rg.getFontColors();
  var cleared = 0;
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < 8; c++) {
      if (vals[r][c] === '' || String(colors[r][c]).toLowerCase() !== '#7a7a7a') continue;
      if (r === 1 && c === 0) continue;
      sh.getRange(r + 1, c + 1).clearContent().setFontColor(null).setFontSize(10);
      cleared++;
    }
  }

  var AUTO = '今月（自動）';
  var J = "'累計入会データ'!$V$2:$V$8000";
  sh.getRange('Q1').setValue(AUTO);
  sh.getRange('Q2').setFormula('=ARRAYFORMULA(TEXT(EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1-SEQUENCE(DATEDIF(MINIFS(' + J + ',' + J + ',">0"),DATE(YEAR(TODAY()),MONTH(TODAY()),1),"M")+1)),"yyyy年m月"))');
  sh.hideColumns(17);
  var b2 = sh.getRange('B2');
  b2.clearDataValidations();
  b2.setNumberFormat('@').setValue(AUTO);
  b2.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange('Q1:Q200'), true).setAllowInvalid(false).build());
  b2.setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12).setBackground('#FFF7E6')
    .setBorder(true, true, true, true, null, null, '#111111', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('P1').setFormula('=IF(OR(B2="",B2="' + AUTO + '"),DATE(YEAR(TODAY()),MONTH(TODAY()),1),IFERROR(DATE(VALUE(LEFT(B2,4)),VALUE(REGEXEXTRACT(B2,"年(\\d+)月")),1),DATE(YEAR(TODAY()),MONTH(TODAY()),1)))');
  sh.getRange('A2').setValue('基準月 ▶').setFontWeight('bold').setFontColor('#111111').setHorizontalAlignment('right');

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily('Meiryo');
  props.setProperty('JL_MIG', 'v1');
  return { ok: true, cleared: cleared };
}

/** 会員動向（本部）の月の値。今年度は実績/見込の行、前年度は前年の行。見つからなければ 0 */
function jlHqExpr_(monExpr, rowCur, rowPrev, key) {
  var H = "'【経堂】会員動向'!";
  return 'LET(mon,' + monExpr + ',fyr,YEAR(mon)-(MONTH(mon)<4),cur,YEAR(TODAY())-(MONTH(TODAY())<4),' +
    'colx,IFERROR(MATCH(MONTH(mon)&"",ARRAYFORMULA(' + H + '$C$1:$N$1&""),0),0),' +
    'rowx,IF(fyr=cur,' + rowCur + ',IF(fyr=cur-1,' + rowPrev + ',0)),' +
    'val,IF(OR(colx=0,rowx=0),"",IF(ISNUMBER(SEARCH("' + key + '",INDEX(' + H + '$B$1:$B$20,rowx))),INDEX(' + H + '$C$1:$N$20,rowx,colx),"")),' +
    'IFERROR(VALUE(SUBSTITUTE(val&"",",","")),0))';
}

/**
 * 入会・退会分析 ①：月初・入会・退会は本部の会員動向の数字を優先（今月の入会・退会は日報）。
 * 会員動向に無い月だけ累計データから数える。③ の平均は対象月が3つ未満なら出さない。1回だけ。
 */
function migrateJoinLeaveHq_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('JL_MIG2') === 'v1') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(JL_SHEET_);
  if (!sh || !ss.getSheetByName('【経堂】会員動向')) return { ok: false };
  var J = "'累計入会データ'!", L = "'累計退会データ'!";
  var jV = J + '$V$2:$V$8000', lU = L + '$U$2:$U$8000';
  var thisMonth = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  var out = [];
  for (var r = 6; r <= 17; r++) {
    var A = '$A' + r;
    var isNow = 'TEXT(' + A + ',"yymm")=TEXT(\'日報\'!$B$1,"0")';
    var calcB = 'COUNTIFS(' + jV + ',">0",' + jV + ',"<"&' + A + ')-COUNTIFS(' + lU + ',"<="&' + A + ')';
    out.push([
      '=LET(h,' + jlHqExpr_(A, 3, 4, '月初') + ',IF(h>0,h,' + calcB + '))',
      '=IF(' + isNow + ',N(\'日報\'!$C$13),LET(h,' + jlHqExpr_(A, 6, 7, '入') + ',IF(AND(h>0,' + A + '<' + thisMonth + '),h,COUNTIF(' + jV + ',' + A + '))))',
      '=IF(' + isNow + ',N(\'日報\'!$C$15),LET(h,' + jlHqExpr_(A, 9, 10, '解') + ',IF(AND(h>0,' + A + '<' + thisMonth + '),h,COUNTIF(' + lU + ',EDATE(' + A + ',1)))))'
    ]);
  }
  sh.getRange(6, 2, 12, 3).setFormulas(out);
  var avg = ['C', 'D', 'E', 'F'].map(function (c) {
    var rg = c + '45:' + c + '56';
    return '=IF(COUNT(' + rg + ')<3,"—",IFERROR(SUMPRODUCT(N(' + rg + '),$B45:$B56)/SUMPRODUCT((' + rg + '<>"")*$B45:$B56),""))';
  });
  sh.getRange('C57:F57').setFormulas([avg]);
  protectInputsOnly_(sh, ['B2']);
  props.setProperty('JL_MIG2', 'v1');
  return { ok: true };
}

/**
 * 入会・退会分析の見た目（1回だけ）。グラフには触らない。
 * 上に基準月の数字カードを足し（行を4つ挿入）、継続率は色の濃さで、6ヶ月継続率は横棒で見せる。
 */
function migrateJoinLeaveLook_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('JL_LOOK') === 'v1') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(JL_SHEET_);
  if (!sh) return { ok: false };
  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#E3E3E3', SOFT = '#F7F7F7', RED = '#B91C1C';

  if (String(sh.getRange('A4').getValue()).indexOf('①') === 0) {
    var keep = sh.getConditionalFormatRules().filter(function (rule) {
      var rg = rule.getRanges()[0];
      return !(rg && rg.getRow() === 45 && rg.getColumn() === 3);
    });
    keep.push(SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#E6A1A1', SpreadsheetApp.InterpolationType.NUMBER, '0.5')
      .setGradientMidpointWithValue('#F7DCDC', SpreadsheetApp.InterpolationType.NUMBER, '0.85')
      .setGradientMaxpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setRanges([sh.getRange('C45:F56')]).build());
    sh.setConditionalFormatRules(keep);
    sh.insertRowsAfter(2, 4);
  }
  if (String(sh.getRange('A8').getValue()).indexOf('①') !== 0) return { ok: false };
  var base = 21, prev = 20;
  sh.getRange('A3:G6').clearDataValidations();
  sh.getRange('A3:G6').clearFormat().setFontFamily('Meiryo').setVerticalAlignment('middle');
  sh.setRowHeight(3, 22);
  sh.setRowHeight(4, 44);
  sh.setRowHeight(5, 22);
  sh.setRowHeight(6, 14);
  sh.getRange('A3').setValue('基準月').setFontSize(9).setFontColor(MUTE).setFontWeight('bold');
  sh.getRange('A4').setFormula('=TEXT($P$1,"yyyy年m月")').setFontSize(12).setFontWeight('bold');
  sh.getRange('A5').setValue('前月').setFontSize(9).setFontColor(MUTE);
  var cards = [
    ['B', '月初会員', '#,##0', '+#,##0;-#,##0;0'],
    ['C', '入会', '#,##0', '+#,##0;-#,##0;0'],
    ['D', '退会', '#,##0', '+#,##0;-#,##0;0'],
    ['E', '純増', '+#,##0;-#,##0;0', '+#,##0;-#,##0;0'],
    ['F', '退会率', '0.0%', '+0.0%;-0.0%;0.0%'],
    ['G', '規約退会', '#,##0', '+#,##0;-#,##0;0']
  ];
  cards.forEach(function (c) {
    sh.getRange(c[0] + '3').setValue(c[1]).setFontSize(9).setFontColor(MUTE).setFontWeight('bold');
    sh.getRange(c[0] + '4').setFormula('=' + c[0] + base).setNumberFormat(c[2]).setFontSize(22).setFontWeight('bold');
    sh.getRange(c[0] + '5').setFormula('=IFERROR(TEXT(' + c[0] + prev + ',"' + c[2] + '"),"")').setFontSize(9).setFontColor(MUTE);
  });
  sh.getRange('A3:G5').setBackground(SOFT).setHorizontalAlignment('left')
    .setBorder(true, null, null, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  var rules = sh.getConditionalFormatRules();

  sh.getRange('A10:G22').setBorder(null, null, null, null, null, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('A21:G21').setBackground('#FDF1F1').setFontWeight('bold');
  for (var r = 10; r <= 22; r++) sh.setRowHeight(r, 22);

  var seg0 = 66, seg1 = 78;
  sh.getRange('F65').setValue('').setBackground(INK);
  jlSegBars_(sh);
  sh.setColumnWidth(6, Math.max(sh.getColumnWidth(6), 110));
  sh.setConditionalFormatRules(rules);
  sh.getRange(1, 1, 90, 7).setFontFamily('Meiryo');
  props.setProperty('JL_LOOK', 'v1');
  return { ok: true };
}

/** ④ の横棒。50%〜100% を棒の長さにして差を見やすくする。全体より3pt以上低いと赤 */
function jlSegBars_(sh) {
  var seg0 = 66, seg1 = 78;
  for (var s = seg0; s <= seg1; s++) {
    var bar = function (color) {
      return 'SPARKLINE(MAX(0,E' + s + '-0.5),{"charttype","bar";"max",0.5;"color1","' + color + '"})';
    };
    sh.getRange('F' + s).setFormula('=IF(E' + s + '="","",IF(E' + s + '<$E$' + seg0 + '-0.03,' + bar('#B91C1C') + ',' + bar('#111111') + '))');
  }
}

function migrateJoinLeaveBars_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('JL_LOOK') !== 'v1' || props.getProperty('JL_BARS') === 'v1') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(JL_SHEET_);
  if (!sh || String(sh.getRange('A64').getValue()).indexOf('④') !== 0) return { ok: false, a64: String(sh && sh.getRange('A64').getValue()) };
  jlSegBars_(sh);
  props.setProperty('JL_BARS', 'v1');
  return { ok: true };
}

/**
 * 入会・退会分析を3ブロック横並びに作り直す（1回だけ）。グラフは消さずに範囲・位置・大きさだけ変える。
 * 左＝① 3年間の推移（直近12ヶ月と1〜3年前の同じ月）、中＝グラフ、右＝② 退会の内訳・③ 継続率・④ 続きやすさ。
 * 上のカードで基準月の前年比・過去3年平均比・直近3ヶ月平均比を見る。月の数字は AL:AP（非表示）の48ヶ月分から引く。
 */
var JL_3Y_V_ = 'v1';
function rebuildJoinLeave3y_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('JL_3Y') === JL_3Y_V_) return { ok: true, skipped: true };
  var sh = ss.getSheetByName(JL_SHEET_);
  if (!sh || !ss.getSheetByName('【経堂】会員動向')) return { ok: false };
  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#E3E3E3', SOFT = '#F7F7F7', HEAD = '#F3F3F3', RED = '#B91C1C', PINK = '#FDF1F1';
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;
  var AUTO = '今月（自動）';
  var keepB2 = String(sh.getRange('B2').getDisplayValue() || '') || AUTO;

  var NC = 48, NR = 70;
  if (sh.getMaxColumns() < NC) sh.insertColumnsAfter(sh.getMaxColumns(), NC - sh.getMaxColumns());
  if (sh.getMaxRows() < NR) sh.insertRowsAfter(sh.getMaxRows(), NR - sh.getMaxRows());
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  sh.setFrozenRows(0);
  var all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart();
  all.clearDataValidations();
  sh.setConditionalFormatRules([]);
  sh.clear();
  sh.showColumns(1, sh.getMaxColumns());
  sh.setRowHeights(1, sh.getMaxRows(), 21);

  var J = "'累計入会データ'!", L = "'累計退会データ'!";
  var jV = J + '$V$2:$V$8000', jW = J + '$W$2:$W$8000', jX = J + '$X$2:$X$8000', jI = J + '$I$2:$I$8000', jY = J + '$Y$2:$Y$8000';
  var lU = L + '$U$2:$U$8000', lP = L + '$P$2:$P$8000', lJ = L + '$J$2:$J$8000', lI = L + '$I$2:$I$8000', lV = L + '$V$2:$V$8000';
  var BM = '$AJ$1';
  var thisMonth = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  var hdr = function (rg) {
    return rg.setBackground(INK).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  };
  var section = function (a1, text) { sh.getRange(a1).setValue(text).setFontSize(13).setFontWeight('bold').setFontColor(INK); };

  // 非表示の作業列：AJ1 基準月 / AJ2・AJ3 ②の合計 / AK 基準月の選択肢 / AL:AP 48ヶ月分 / AR:AU グラフ用36ヶ月
  sh.getRange('AJ1').setFormula('=IF(OR(B2="",B2="' + AUTO + '"),' + thisMonth + ',IFERROR(DATE(VALUE(LEFT(B2,4)),VALUE(REGEXEXTRACT(B2,"年(\\d+)月")),1),' + thisMonth + '))');
  sh.getRange('AK1').setValue(AUTO);
  sh.getRange('AK2').setFormula('=ARRAYFORMULA(TEXT(EDATE(' + thisMonth + ',1-SEQUENCE(DATEDIF(MINIFS(' + jV + ',' + jV + ',">0"),' + thisMonth + ',"M")+1)),"yyyy年m月"))');
  sh.getRange('AL1:AP1').setValues([['月', '月初会員', '入会', '退会', '退会率']]);
  var hf = [];
  for (var k = 0; k < 48; k++) {
    var r = k + 2, A = '$AL' + r;
    var isNow = 'TEXT(' + A + ',"yymm")=TEXT(\'日報\'!$B$1,"0")';
    var calcB = 'COUNTIFS(' + jV + ',">0",' + jV + ',"<"&' + A + ')-COUNTIFS(' + lU + ',"<="&' + A + ')';
    hf.push([
      '=EDATE(' + BM + ',' + (k - 47) + ')',
      '=LET(h,' + jlHqExpr_(A, 3, 4, '月初') + ',IF(h>0,h,' + calcB + '))',
      '=IF(' + isNow + ',N(\'日報\'!$C$13),LET(h,' + jlHqExpr_(A, 6, 7, '入') + ',IF(AND(h>0,' + A + '<' + thisMonth + '),h,COUNTIF(' + jV + ',' + A + '))))',
      '=IF(' + isNow + ',N(\'日報\'!$C$15),LET(h,' + jlHqExpr_(A, 9, 10, '解') + ',IF(AND(h>0,' + A + '<' + thisMonth + '),h,COUNTIF(' + lU + ',EDATE(' + A + ',1)))))',
      '=IFERROR(AO' + r + '/AM' + r + ',"")'
    ]);
  }
  sh.getRange(2, 38, 48, 5).setFormulas(hf);
  sh.getRange('AL2:AL49').setNumberFormat('yyyy/m');
  sh.getRange('AP2:AP49').setNumberFormat('0.0%');
  sh.getRange('AR1:AU1').setValues([['月', '入会', '退会', '退会率']]);
  var cd = [];
  for (var c = 0; c < 36; c++) { var hr = 14 + c; cd.push(['=AL' + hr, '=AN' + hr, '=AO' + hr, '=AP' + hr]); }
  sh.getRange(2, 44, 36, 4).setFormulas(cd);
  sh.getRange('AR2:AR37').setNumberFormat('yyyy/m');
  sh.getRange('AU2:AU37').setNumberFormat('0.0%');

  // タイトル・基準月
  sh.getRange('A1').setValue('入会・退会分析').setFontSize(18).setFontWeight('bold').setFontColor(INK);
  sh.getRange('A2').setValue('基準月 ▶').setFontWeight('bold').setFontColor(INK).setHorizontalAlignment('right');
  var b2 = sh.getRange('B2:C2').merge();
  sh.getRange('B2').setNumberFormat('@').setValue(keepB2);
  b2.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange('AK1:AK200'), true).setAllowInvalid(false).build());
  b2.setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12).setBackground('#FFF7E6')
    .setBorder(true, true, true, true, null, null, INK, SOLID);

  // カード（3〜5行）。基準月＝AL:AP の49行目、前月48、前年37、2年前25、3年前13、直近3ヶ月46〜48
  sh.setRowHeight(3, 22); sh.setRowHeight(4, 44); sh.setRowHeight(5, 22); sh.setRowHeight(6, 14);
  var card = function (c1, label, valF, fmt, subF, size) {
    [3, 4, 5].forEach(function (rr) { sh.getRange(rr, c1, 1, 2).merge(); });
    sh.getRange(3, c1).setValue(label).setFontSize(9).setFontColor(MUTE).setFontWeight('bold');
    sh.getRange(4, c1).setFormula(valF).setNumberFormat(fmt).setFontSize(size || 22).setFontWeight('bold').setFontColor(INK);
    sh.getRange(5, c1).setFormula(subF).setFontSize(9).setFontColor(MUTE);
    sh.getRange(3, c1, 3, 2).setBackground(SOFT).setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setBorder(true, null, null, true, null, null, INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.getRange(3, c1 + 1, 3, 1).setBorder(null, null, null, true, null, null, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
  };
  var pm = function (v) { return 'TEXT(' + v + ',"+#,##0;-#,##0;±0")'; };
  card(1, '基準月', '=TEXT(' + BM + ',"yyyy年m月")', 'General', '=IF(' + BM + '=' + thisMonth + ',"今月は途中経過","")', 14);
  card(3, '月初会員', '=AM49', '#,##0', '="前年 "&TEXT(AM37,"#,##0")&"（"&' + pm('AM49-AM37') + '&"）"');
  card(5, '入会', '=AN49', '#,##0', '="前月 "&TEXT(AN48,"#,##0")');
  card(7, '退会', '=AO49', '#,##0', '="前月 "&TEXT(AO48,"#,##0")');
  card(9, '純増', '=AN49-AO49', '+#,##0;-#,##0;0', '="前月 "&' + pm('AN48-AO48'));
  card(11, '退会率', '=AP49', '0.0%', '="前月 "&TEXT(AP48,"0.0%")');
  [['AN', '入会', 15], ['AO', '退会', 21]].forEach(function (g) {
    var x = g[0], avg3y = 'AVERAGE(' + x + '37,' + x + '25,' + x + '13)', avg3m = 'AVERAGE(' + x + '46:' + x + '48)';
    card(g[2], g[1] + '｜前年同月比', '=IFERROR(' + x + '49/' + x + '37,"")', '0%', '="前年 "&' + x + '37&"（"&' + pm(x + '49-' + x + '37') + '&"）"');
    card(g[2] + 2, g[1] + '｜過去3年平均比', '=IFERROR(' + x + '49/' + avg3y + ',"")', '0%', '="3年平均 "&TEXT(' + avg3y + ',"0.0")');
    card(g[2] + 4, g[1] + '｜直近3ヶ月比', '=IFERROR(' + x + '49/' + avg3m + ',"")', '0%', '="3ヶ月平均 "&TEXT(' + avg3m + ',"0.0")');
  });

  // ① 3年間の推移
  section('A7', '① 3年間の推移（直近12ヶ月と、1〜3年前の同じ月）');
  sh.getRange('A8:A9').merge().setValue('月');
  sh.getRange('B8:B9').merge().setValue('月初会員');
  sh.getRange('C8:G8').merge().setValue('入会');
  sh.getRange('H8:L8').merge().setValue('退会');
  sh.getRange('M8:M9').merge().setValue('退会率');
  sh.getRange('C9:L9').setValues([['今回', '1年前', '2年前', '3年前', '3年平均', '今回', '1年前', '2年前', '3年前', '3年平均']]);
  hdr(sh.getRange('A8:M9')).setFontSize(9);
  sh.getRange('C9:L9').setBackground('#3A3A3A');
  var mx = [];
  for (var i = 0; i < 12; i++) {
    var rr = 10 + i, h = 38 + i;
    mx.push(['=AL' + h, '=AM' + h,
      '=AN' + h, '=AN' + (h - 12), '=AN' + (h - 24), '=AN' + (h - 36), '=IFERROR(AVERAGE(D' + rr + ':F' + rr + '),"")',
      '=AO' + h, '=AO' + (h - 12), '=AO' + (h - 24), '=AO' + (h - 36), '=IFERROR(AVERAGE(I' + rr + ':K' + rr + '),"")',
      '=AP' + h]);
  }
  sh.getRange(10, 1, 12, 13).setFormulas(mx);
  sh.getRange(22, 1, 1, 13).setFormulas([['="12ヶ月計"', '=AVERAGE(B10:B21)',
    '=SUM(C10:C21)', '=SUM(D10:D21)', '=SUM(E10:E21)', '=SUM(F10:F21)', '=IFERROR(AVERAGE(D22:F22),"")',
    '=SUM(H10:H21)', '=SUM(I10:I21)', '=SUM(J10:J21)', '=SUM(K10:K21)', '=IFERROR(AVERAGE(I22:K22),"")',
    '=IFERROR(SUM(H10:H21)/SUM(B10:B21),"")']]);
  sh.getRange('A10:A21').setNumberFormat('yyyy/m');
  sh.getRange('B10:L22').setNumberFormat('#,##0');
  sh.getRange('G10:G21').setNumberFormat('0.0');
  sh.getRange('L10:L21').setNumberFormat('0.0');
  sh.getRange('M10:M22').setNumberFormat('0.0%');
  sh.getRange('A10:M22').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange('B10:M22').setHorizontalAlignment('right');
  sh.getRange('A10:A22').setHorizontalAlignment('center');
  sh.getRange('G10:G22').setBackground(SOFT);
  sh.getRange('L10:L22').setBackground(SOFT);
  sh.getRange('C10:C22').setFontWeight('bold');
  sh.getRange('H10:H22').setFontWeight('bold');
  sh.getRange('A10:M22').setBorder(null, null, true, null, null, true, LINE, SOLID);
  ['C', 'H', 'M'].forEach(function (col) { sh.getRange(col + '8:' + col + '22').setBorder(null, true, null, null, null, null, '#BDBDBD', SOLID); });
  sh.getRange('A21:M21').setBackground(PINK).setFontWeight('bold');
  sh.getRange('A22:M22').setBackground(HEAD).setFontWeight('bold').setBorder(true, null, true, null, null, null, INK, SOLID);
  for (var rh = 8; rh <= 22; rh++) sh.setRowHeight(rh, 22);

  // ② 退会した人の内訳（AB〜AF）
  section('AB7', '② 退会した人の内訳（直近12ヶ月）');
  sh.getRange('AB8:AF8').setValues([['区分', '項目', '直近12ヶ月', '割合', '基準月']]);
  hdr(sh.getRange('AB8:AF8')).setFontSize(9);
  var isD = 'ISNUMBER(' + lU + ')';
  var win12 = '(' + isD + '*(IFERROR(' + lU + '*1,0)>=EDATE(' + BM + ',-10))*(IFERROR(' + lU + '*1,0)<=EDATE(' + BM + ',1)))';
  var winM = '(' + isD + '*(IFERROR(' + lU + '*1,0)=EDATE(' + BM + ',1)))';
  var P = 'IFERROR(VALUE(' + lP + '),-1)', AG = 'IFERROR(VALUE(' + lJ + '),-1)';
  var leaveRows = [
    ['在籍期間', '3ヶ月以内', '(' + P + '>=0)*(' + P + '<=3)'],
    ['', '4〜6ヶ月', '(' + P + '>=4)*(' + P + '<=6)'],
    ['', '7〜12ヶ月', '(' + P + '>=7)*(' + P + '<=12)'],
    ['', '1〜2年', '(' + P + '>=13)*(' + P + '<=24)'],
    ['', '2年超', '(' + P + '>=25)'],
    ['年代', '10代', '(' + AG + '>=0)*(' + AG + '<20)'],
    ['', '20代', '(' + AG + '>=20)*(' + AG + '<30)'],
    ['', '30代', '(' + AG + '>=30)*(' + AG + '<40)'],
    ['', '40代', '(' + AG + '>=40)*(' + AG + '<50)'],
    ['', '50代', '(' + AG + '>=50)*(' + AG + '<60)'],
    ['', '60代以上', '(' + AG + '>=60)'],
    ['男女', '男', '(' + lI + '="男")'],
    ['', '女', '(' + lI + '="女")'],
    ['プラン', 'ナショナル会員U', '(' + lV + '="ナショナル会員U")'],
    ['', '法人個人月払B', '(' + lV + '="法人個人月払B")'],
    ['', '法人個人月払A', '(' + lV + '="法人個人月払A")'],
    ['', '閉店移籍会員3', '(' + lV + '="閉店移籍会員3")']
  ];
  sh.getRange('AJ2').setFormula('=SUMPRODUCT(' + win12 + ')');
  sh.getRange('AJ3').setFormula('=SUMPRODUCT(' + winM + ')');
  var lr = leaveRows.map(function (x, n) {
    var r2 = 9 + n;
    return [x[0], x[1], '=SUMPRODUCT(' + win12 + '*' + x[2] + ')', '=IFERROR(AD' + r2 + '/$AJ$2,"")', '=SUMPRODUCT(' + winM + '*' + x[2] + ')'];
  });
  var oth = 9 + lr.length;
  lr.push(['', 'その他', '=$AJ$2-SUM(AD' + (oth - 4) + ':AD' + (oth - 1) + ')', '=IFERROR(AD' + oth + '/$AJ$2,"")', '=$AJ$3-SUM(AF' + (oth - 4) + ':AF' + (oth - 1) + ')']);
  sh.getRange(9, 28, lr.length, 2).setValues(lr.map(function (x) { return [x[0], x[1]]; }));
  sh.getRange(9, 30, lr.length, 3).setFormulas(lr.map(function (x) { return [x[2], x[3], x[4]]; }));
  var lEnd = 8 + lr.length;
  sh.getRange('AB9:AB' + lEnd).setFontWeight('bold');
  sh.getRange('AE9:AE' + lEnd).setNumberFormat('0%');
  sh.getRange('AD9:AF' + lEnd).setHorizontalAlignment('right');
  sh.getRange('AB9:AF' + lEnd).setBorder(null, null, true, null, null, true, LINE, SOLID);
  [13, 19, 21].forEach(function (r3) { sh.getRange('AB' + r3 + ':AF' + r3).setBorder(null, null, true, null, null, null, INK, SOLID); });

  // ③ 入会した月ごとの継続率（AB〜AG）
  var c0 = lEnd + 3, avgRow = c0 + 14;
  section('AB' + c0, '③ 入会した月ごとの継続率');
  sh.getRange(c0 + 1, 28, 1, 6).setValues([['入会月', '入会数', '1ヶ月後', '3ヶ月後', '6ヶ月後', '12ヶ月後']]);
  hdr(sh.getRange(c0 + 1, 28, 1, 6)).setFontSize(9);
  var co = [];
  for (var m = 0; m < 12; m++) {
    var r4 = c0 + 2 + m, AB = 'AB' + r4, AC = 'AC' + r4;
    var row = ['=EDATE(' + BM + ',' + (m - 12) + ')', '=COUNTIF(' + jV + ',' + AB + ')'];
    [1, 3, 6, 12].forEach(function (kk) {
      row.push('=IF(OR(' + AC + '=0,EDATE(' + AB + ',' + kk + ')>' + BM + '),"",(COUNTIFS(' + jV + ',' + AB + ',' + jW + ',"")+COUNTIFS(' + jV + ',' + AB + ',' + jW + ',">"&EDATE(' + AB + ',' + kk + ')))/' + AC + ')');
    });
    co.push(row);
  }
  sh.getRange(c0 + 2, 28, 12, 6).setFormulas(co);
  var cA = c0 + 2, cZ = c0 + 13;
  var avg = ['="平均"', '=SUM(AC' + cA + ':AC' + cZ + ')'].concat(['AD', 'AE', 'AF', 'AG'].map(function (col) {
    var rg = col + cA + ':' + col + cZ;
    return '=IF(COUNT(' + rg + ')<3,"—",IFERROR(SUMPRODUCT(N(' + rg + '),$AC' + cA + ':$AC' + cZ + ')/SUMPRODUCT((' + rg + '<>"")*$AC' + cA + ':$AC' + cZ + '),""))';
  }));
  sh.getRange(avgRow, 28, 1, 6).setFormulas([avg]);
  sh.getRange('AB' + cA + ':AB' + cZ).setNumberFormat('yyyy/m').setHorizontalAlignment('center');
  sh.getRange(cA, 30, 13, 4).setNumberFormat('0.0%');
  sh.getRange(cA, 29, 13, 5).setHorizontalAlignment('right');
  sh.getRange(cA, 28, 12, 6).setBorder(null, null, true, null, null, true, LINE, SOLID);
  sh.getRange(avgRow, 28, 1, 6).setBackground(HEAD).setFontWeight('bold').setBorder(true, null, true, null, null, null, INK, SOLID);

  // ④ 続きやすさの違い（AB〜AG）
  var d0 = avgRow + 3;
  section('AB' + d0, '④ 続きやすさの違い（6ヶ月後も在籍している割合）');
  sh.getRange(d0 + 1, 28, 1, 6).setValues([['区分', '項目', '入会数', '6ヶ月後も在籍', '6ヶ月継続率', '']]);
  hdr(sh.getRange(d0 + 1, 28, 1, 6)).setFontSize(9);
  var win = '(' + jV + '>=EDATE(' + BM + ',-18))*(' + jV + '<=EDATE(' + BM + ',-7))';
  var stay = '(' + J + '$Z$2:$Z$8000>6)';
  var segs = [
    ['全体', '全員', '1'],
    ['年代', '10代', '(' + jX + '<20)'],
    ['', '20代', '(' + jX + '>=20)*(' + jX + '<30)'],
    ['', '30代', '(' + jX + '>=30)*(' + jX + '<40)'],
    ['', '40代', '(' + jX + '>=40)*(' + jX + '<50)'],
    ['', '50代', '(' + jX + '>=50)*(' + jX + '<60)'],
    ['', '60代以上', '(' + jX + '>=60)'],
    ['男女', '男', '(' + jI + '="男")'],
    ['', '女', '(' + jI + '="女")'],
    ['プラン', 'ナショナル会員U', '(' + jY + '="ナショナル会員U")'],
    ['', '法人個人月払B', '(' + jY + '="法人個人月払B")'],
    ['', '法人個人月払A', '(' + jY + '="法人個人月払A")'],
    ['', '閉店移籍会員3', '(' + jY + '="閉店移籍会員3")']
  ];
  var s0 = d0 + 2, s1 = d0 + 1 + segs.length;
  var sg = segs.map(function (x, n) {
    var r5 = s0 + n;
    var bar = function (color) { return 'SPARKLINE(MAX(0,AF' + r5 + '-0.5),{"charttype","bar";"max",0.5;"color1","' + color + '"})'; };
    return [x[0], x[1], '=SUMPRODUCT(' + win + '*' + x[2] + ')', '=SUMPRODUCT(' + win + '*' + x[2] + '*' + stay + ')', '=IFERROR(AE' + r5 + '/AD' + r5 + ',"")',
      '=IF(AF' + r5 + '="","",IF(AF' + r5 + '<$AF$' + s0 + '-0.03,' + bar(RED) + ',' + bar(INK) + '))'];
  });
  sh.getRange(s0, 28, sg.length, 2).setValues(sg.map(function (x) { return [x[0], x[1]]; }));
  sh.getRange(s0, 30, sg.length, 4).setFormulas(sg.map(function (x) { return [x[2], x[3], x[4], x[5]]; }));
  sh.getRange('AB' + s0 + ':AB' + s1).setFontWeight('bold');
  sh.getRange('AF' + s0 + ':AF' + s1).setNumberFormat('0.0%').setFontWeight('bold');
  sh.getRange(s0, 30, sg.length, 3).setHorizontalAlignment('right');
  sh.getRange(s0, 28, sg.length, 6).setBorder(null, null, true, null, null, true, LINE, SOLID);
  sh.getRange(s0, 28, 1, 6).setBackground(HEAD);

  // 条件付き書式
  var errs = [];
  var cf = function () { return SpreadsheetApp.newConditionalFormatRule(); };
  try { sh.setConditionalFormatRules([
    cf().whenFormulaSatisfied('=AND(ISNUMBER($G10),C10<$G10)').setFontColor(RED).setRanges([sh.getRange('C10:C22')]).build(),
    cf().whenFormulaSatisfied('=AND(ISNUMBER($L10),H10>$L10)').setFontColor(RED).setRanges([sh.getRange('H10:H22')]).build(),
    cf().whenNumberLessThan(1).setFontColor(RED).setRanges([sh.getRange('O4'), sh.getRange('Q4'), sh.getRange('S4')]).build(),
    cf().whenNumberGreaterThan(1).setFontColor(RED).setRanges([sh.getRange('U4'), sh.getRange('W4'), sh.getRange('Y4')]).build(),
    cf().setGradientMinpointWithValue('#E6A1A1', SpreadsheetApp.InterpolationType.NUMBER, '0.5')
      .setGradientMidpointWithValue('#F7DCDC', SpreadsheetApp.InterpolationType.NUMBER, '0.85')
      .setGradientMaxpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setRanges([sh.getRange(cA, 30, 12, 4)]).build(),
    cf().whenFormulaSatisfied('=AND(ISNUMBER(AF' + (s0 + 1) + '),AF' + (s0 + 1) + '<$AF$' + s0 + '-0.03)').setFontColor(RED)
      .setRanges([sh.getRange('AF' + (s0 + 1) + ':AF' + s1)]).build()
  ]); } catch (e) { errs.push('cf:' + e.message); }

  // 幅・フォント・非表示
  sh.setColumnWidth(1, 70); sh.setColumnWidth(2, 66);
  for (var w = 3; w <= 12; w++) sh.setColumnWidth(w, 52);
  sh.setColumnWidth(13, 60); sh.setColumnWidth(14, 24);
  for (var w2 = 15; w2 <= 26; w2++) sh.setColumnWidth(w2, 62);
  sh.setColumnWidth(27, 24);
  [72, 120, 76, 76, 84, 110, 24].forEach(function (px, n) { sh.setColumnWidth(28 + n, px); });
  sh.getRange(1, 1, NR, 34).setFontFamily('Meiryo');
  sh.hideColumns(36, sh.getMaxColumns() - 35);
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(2);

  // グラフは消さずに範囲・位置・大きさだけ。上から順に ①推移 / 在籍期間 / 継続率
  var cs = sh.getCharts().slice().sort(function (a, b) {
    var x = a.getContainerInfo(), y = b.getContainerInfo();
    return (x.getAnchorRow() - y.getAnchorRow()) || (x.getAnchorColumn() - y.getAnchorColumn());
  });
  var place = function (ch, ranges, row, hpx, title, hidden) {
    try {
      var bd = ch.modify().clearRanges();
      ranges.forEach(function (a) { bd.addRange(sh.getRange(a)); });
      bd.setPosition(row, 15, 0, 0).setOption('width', 744).setOption('height', hpx);
      if (hidden) bd.setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH);
      if (title) bd.setOption('title', title);
      sh.updateChart(bd.build());
    } catch (e) { errs.push(row + ':' + e.message); }
  };
  if (cs[0]) place(cs[0], ['AR1:AR37', 'AS1:AT37', 'AU1:AU37'], 7, 352, '入会・退会と退会率（3年）', true);
  if (cs[1]) place(cs[1], ['AC9:AD13'], 24, 260);
  if (cs[2]) place(cs[2], ['AB' + (c0 + 1) + ':AB' + cZ, 'AE' + (c0 + 1) + ':AF' + cZ], 38, 300);

  protectInputsOnly_(sh, ['B2']);
  props.setProperty('JL_3Y', JL_3Y_V_);
  return { ok: !errs.length, charts: cs.length, errs: errs };
}

function fixJoinLeave3yChart_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('JL_3Y') !== JL_3Y_V_ || props.getProperty('JL_3Y_CH') === 'v1') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(JL_SHEET_);
  var target = function () {
    return sh.getCharts().filter(function (c) {
      var ci = c.getContainerInfo();
      return ci.getAnchorColumn() !== 15 || ci.getAnchorRow() === 7;
    });
  };
  var cs = target();
  if (cs.length !== 1) { props.setProperty('JL_3Y_CH', 'v1'); return { ok: false, left: cs.length }; }
  var ch = cs[0], log = [];
  var step = function (name, fn) {
    try { sh.updateChart(fn(ch.modify()).build()); ch = target()[0] || ch; log.push(name + ':ok'); return true; }
    catch (e) { log.push(name + ':' + e.message.slice(0, 60)); return false; }
  };
  log.push('ranges:' + ch.getRanges().map(function (r) { return r.getA1Notation(); }).join(','));
  step('pos', function (b) { return b.setPosition(7, 15, 0, 0).setOption('width', 744).setOption('height', 352); });
  step('title', function (b) { return b.setOption('title', '入会・退会と退会率（3年）'); });
  var okR = step('ranges', function (b) { return b.clearRanges().addRange(sh.getRange('AR1:AR37')).addRange(sh.getRange('AS1:AT37')).addRange(sh.getRange('AU1:AU37')); });
  if (!okR) okR = step('range1', function (b) { return b.clearRanges().addRange(sh.getRange('AR1:AU37')); });
  if (okR) step('hidden', function (b) { return b.setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH); });
  props.setProperty('JL_3Y_CH', 'v1');
  return { ok: true, log: log };
}

/** ① の推移グラフを Sheets API で付け替える（デザインはそのまま、データ範囲・タイトル・位置・大きさだけ）。1回だけ */
function moveJoinLeaveTrendChartApi_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('JL_3Y') !== JL_3Y_V_ || props.getProperty('JL_3Y_CH2') === 'v1') return { ok: true, skipped: true };
  var got;
  try { got = Sheets.Spreadsheets.get(ss.getId(), { fields: 'sheets(properties(sheetId,title),charts(chartId,spec,position))' }); }
  catch (e) { return { ok: false, get: String(e.message).slice(0, 300) }; }
  var sheet = (got.sheets || []).filter(function (s) { return s.properties.title === JL_SHEET_; })[0];
  if (!sheet) return { ok: false, sheet: 'none' };
  var sid = sheet.properties.sheetId;
  var ch = (sheet.charts || []).filter(function (c) {
    var a = c.position && c.position.overlayPosition && c.position.overlayPosition.anchorCell;
    return !a || a.columnIndex !== 14 || a.rowIndex === 6;
  });
  if (ch.length !== 1 || !ch[0].spec.basicChart) { props.setProperty('JL_3Y_CH2', 'v1'); return { ok: false, found: ch.length }; }
  var c = ch[0], spec = c.spec, bc = spec.basicChart;
  var src = function (col) { return { sourceRange: { sources: [{ sheetId: sid, startRowIndex: 0, endRowIndex: 37, startColumnIndex: col, endColumnIndex: col + 1 }] } }; };
  bc.domains = [{ domain: src(43) }];
  (bc.series || []).forEach(function (s, i) { if (i < 3) s.series = src(44 + i); });
  bc.series = (bc.series || []).slice(0, 3);
  bc.headerCount = 1;
  spec.title = '入会・退会と退会率（3年）';
  spec.hiddenDimensionStrategy = 'SHOW_ALL';
  var body = {
    requests: [
      { updateChartSpec: { chartId: c.chartId, spec: spec } },
      { updateEmbeddedObjectPosition: { objectId: c.chartId, fields: '*', newPosition: { overlayPosition: { anchorCell: { sheetId: sid, rowIndex: 6, columnIndex: 14 }, offsetXPixels: 0, offsetYPixels: 0, widthPixels: 744, heightPixels: 352 } } } }
    ]
  };
  var err = '';
  try { Sheets.Spreadsheets.batchUpdate(body, ss.getId()); } catch (e) { err = String(e.message).slice(0, 300); }
  props.setProperty('JL_3Y_CH2', 'v1');
  return { ok: !err, err: err };
}

/** 入力欄以外を「編集前に警告」にする（店舗の人の誤入力よけ） */
function protectInputsOnly_(sh, inputs) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  var p = sh.protect().setDescription('入力はプルダウンの欄だけ');
  p.setWarningOnly(true);
  p.setUnprotectedRanges(inputs.map(function (a) { return sh.getRange(a); }));
}

/**
 * 会議用：期間はプルダウンで選ぶ（M1 かんたん期間・B2 対象月・D2 比較）。選ぶと B4:C5 の日付を自動で入れる。
 * 退会は「実際に辞めた月（最終在籍月）」で数える＝会員動向の「解除」と同じ。1回だけ。
 */
var KAIGI_SHEET_ = '会議用';
var KAIGI_MODES_ = ['月を選ぶ', '開始月〜対象月', '直近2ヶ月', '直近3ヶ月', '直近6ヶ月', '今年度（4月〜）', '手入力'];
var KAIGI_CMP_ = ['前年同期', '前の期間', '月ごとの増減', '比較なし'];

function migrateKaigiSheet_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('KAIGI_MIG') === 'v1') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(KAIGI_SHEET_);
  var mst = ss.getSheetByName('分析用_会員マスタ');
  var agg = ss.getSheetByName('分析用_期間集計');
  if (!sh || !mst || !agg) return { ok: false };

  mst.getRange('W1').setValue('退会基準日（最終在籍月末）');
  mst.getRange('W2').setFormula('=ARRAYFORMULA(IF(S2:S8000="","",IF(P2:P8000="","",IFERROR(EOMONTH(P2:P8000,-1),""))))');
  var rows = Math.min(agg.getMaxRows(), 120);
  var fr = agg.getRange(1, 1, rows, 5).getFormulas();
  var changed = 0;
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < 5; c++) {
      var f = fr[r][c];
      if (!f || f.indexOf("'分析用_会員マスタ'!O$2:O$8000") < 0) continue;
      agg.getRange(r + 1, c + 1).setFormula(f.split("'分析用_会員マスタ'!O$2:O$8000").join("'分析用_会員マスタ'!W$2:W$8000"));
      changed++;
    }
  }

  var J = "'累計入会データ'!$V$2:$V$8000";
  sh.getRange('AH1').setFormula('=ARRAYFORMULA(TEXT(EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1-SEQUENCE(DATEDIF(MINIFS(' + J + ',' + J + ',">0"),DATE(YEAR(TODAY()),MONTH(TODAY()),1),"M")+1)),"yyyy年m月"))');
  sh.hideColumns(34);
  var drop = function (a, list, value) {
    var cell = sh.getRange(a);
    cell.clearDataValidations();
    cell.setNumberFormat('@').setValue(value);
    var rule = Array.isArray(list)
      ? SpreadsheetApp.newDataValidation().requireValueInList(list, true)
      : SpreadsheetApp.newDataValidation().requireValueInRange(list, true);
    cell.setDataValidation(rule.setAllowInvalid(false).build());
    cell.setHorizontalAlignment('center').setFontWeight('bold').setBackground('#FFF7E6')
      .setBorder(true, true, true, true, null, null, '#111111', SpreadsheetApp.BorderStyle.SOLID);
  };
  var b4 = sh.getRange('B4').getValue();
  var month = b4 instanceof Date ? Utilities.formatDate(b4, 'Asia/Tokyo', 'yyyy年M月') : '';
  if (!month) {
    var prev = new Date(); prev.setDate(1); prev.setMonth(prev.getMonth() - 1);
    month = Utilities.formatDate(prev, 'Asia/Tokyo', 'yyyy年M月');
  }
  drop('M1', KAIGI_MODES_, KAIGI_MODES_[0]);
  drop('B2', sh.getRange('AH1:AH200'), month);
  drop('D2', KAIGI_CMP_, KAIGI_CMP_[0]);
  sh.getRange('A2').setValue('対象月 ▶').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('C2').setValue('比較 ▶').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('B4:E5').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  kaigiApplyPeriod_(sh);
  protectInputsOnly_(sh, ['D1', 'F1', 'H1', 'J1', 'M1', 'B2', 'D2', 'B4:E5']);
  props.setProperty('KAIGI_MIG', 'v1');
  return { ok: true, changed: changed };
}

/**
 * 会議用の期間開始時会員・入会・退会：期間が月単位なら本部の会員動向（今月の入会・退会は日報）の合計を使う。
 * 会員動向に無い期間は従来どおり累計データから。1回だけ。
 */
function migrateKaigiHq_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('KAIGI_MIG2') === 'v1') return { ok: true, skipped: true };
  var agg = ss.getSheetByName('分析用_期間集計');
  if (!agg || !ss.getSheetByName('【経堂】会員動向')) return { ok: false };
  var thisMonth = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  var sumExpr = function (X, rowCur, rowPrev, key, nippoCell) {
    return 'LET(s,' + X + '$6,e,' + X + '$7,IF(NOT(IFERROR(AND(ISNUMBER(s),ISNUMBER(e),DAY(s)=1,e=EOMONTH(e,0),e>=s),FALSE)),"",' +
      'LET(ms,ARRAYFORMULA(EDATE(s,SEQUENCE(DATEDIF(s,e+1,"M"),1,0))),' +
      'vs,MAP(ms,LAMBDA(mm,IF(mm=' + thisMonth + ',N(\'日報\'!' + nippoCell + '),IF(mm>' + thisMonth + ',0,' + jlHqExpr_('mm', rowCur, rowPrev, key) + ')))),' +
      'IF(COUNTIF(vs,">0")=ROWS(vs),SUM(vs),""))))';
  };
  var startExpr = function (X) {
    return 'IFERROR(IF(DAY(' + X + '$6)=1,LET(v,' + jlHqExpr_(X + '$6', 3, 4, '月初') + ',IF(v>0,v,"")),""),"")';
  };
  var done = 0;
  ['B', 'C', 'D', 'E'].forEach(function (X) {
    [[18, startExpr(X)], [20, sumExpr(X, 6, 7, '入', '$C$13')], [21, sumExpr(X, 9, 10, '解', '$C$15')]].forEach(function (p) {
      var cell = agg.getRange(X + p[0]);
      var f = cell.getFormula();
      if (!f || f.indexOf('hqv') >= 0) return;
      cell.setFormula('=LET(hqv,' + p[1] + ',IF(hqv<>"",hqv,' + f.slice(1) + '))');
      done++;
    });
  });
  var mst = ss.getSheetByName('分析用_会員マスタ');
  if (mst) mst.getRange('W2:W8000').setNumberFormat('yyyy/m/d');
  props.setProperty('KAIGI_MIG2', 'v1');
  return { ok: true, done: done };
}

/**
 * 会議用を見やすく作り直す（v2）。上にプルダウン、まん中に大きい数字のカード、下に表。
 * 数字は 分析用_期間集計 から読むだけ。期間の日付は B6:C6（今回）・H6:I6（比較）。
 */
var KAIGI_V_ = 'v4';
var KAIGI_CELLS_ = { mode: 'B4', month: 'E4', cmp: 'H4', cur: 'B6:C6', prev: 'H6:I6', from: 'K4' };

function rebuildKaigiSheet_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('KAIGI_V') === KAIGI_V_) return { ok: true, skipped: true };
  var sh = ss.getSheetByName(KAIGI_SHEET_);
  var agg = ss.getSheetByName('分析用_期間集計');
  if (!sh || !agg) return { ok: false };
  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#E3E3E3', SOFT = '#F7F7F7', RED = '#B91C1C', INPUT = '#FFF7E6';
  var G = "'分析用_期間集計'!";

  var v2 = /^\d{4}年\d{1,2}月$/.test(String(sh.getRange('E4').getDisplayValue()));
  var oldMonth = String(sh.getRange(v2 ? 'E4' : 'B2').getDisplayValue());
  var oldCmp = String(sh.getRange(v2 ? 'H4' : 'D2').getDisplayValue());
  if (!/^\d{4}年\d{1,2}月$/.test(oldMonth)) {
    var pm = new Date(); pm.setDate(1); pm.setMonth(pm.getMonth() - 1);
    oldMonth = Utilities.formatDate(pm, 'Asia/Tokyo', 'yyyy年M月');
  }
  if (KAIGI_CMP_.indexOf(oldCmp) < 0) oldCmp = KAIGI_CMP_[0];

  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  var full = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  try { full.breakApart(); } catch (eB) {}
  full.clearDataValidations();
  sh.clear();
  sh.setConditionalFormatRules([]);
  try { sh.showColumns(1, sh.getMaxColumns()); } catch (eS) {}
  try { sh.showRows(1, sh.getMaxRows()); } catch (eR) {}
  if (sh.getMaxColumns() < 34) sh.insertColumnsAfter(sh.getMaxColumns(), 34 - sh.getMaxColumns());
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.setHiddenGridlines(true);

  sh.setColumnWidth(1, 18);
  sh.setColumnWidth(2, 150);
  for (var c = 3; c <= 13; c++) sh.setColumnWidth(c, 82);
  sh.setColumnWidth(14, 18);
  var area = sh.getRange(1, 1, 60, 14);
  area.setFontFamily('Meiryo').setFontSize(10).setFontColor(INK).setVerticalAlignment('middle');
  for (var r = 1; r <= 60; r++) sh.setRowHeight(r, 24);

  // 見出し
  sh.setRowHeight(1, 46);
  sh.getRange('B1:F1').merge().setValue('会議用レポート').setFontSize(18).setFontWeight('bold');
  sh.getRange('H1:M1').merge().setFormula('=' + KAIGI_CELLS_.month + '&IF(' + KAIGI_CELLS_.mode + '="月を選ぶ","",IF(' + KAIGI_CELLS_.mode + '="手入力","（手入力）","　"&' + KAIGI_CELLS_.mode + '))&"　｜　"&IF(' + KAIGI_CELLS_.cmp + '="比較なし","比較なし",' + KAIGI_CELLS_.cmp + '&"と比較")')
    .setHorizontalAlignment('right').setFontColor(MUTE).setFontSize(10);
  sh.getRange('B1:M1').setBorder(null, null, true, null, null, null, RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(2, 12);

  // 操作（プルダウン）
  sh.setRowHeight(3, 20);
  [['B3', '期間の選び方'], ['E3', '対象月'], ['H3', '比較']].forEach(function (p) {
    sh.getRange(p[0]).setValue(p[1]).setFontSize(9).setFontColor(MUTE);
  });
  sh.setRowHeight(4, 32);
  var dropdown = function (a1, list, value) {
    var rg = sh.getRange(a1).merge();
    var cell = rg.getCell(1, 1);
    cell.setNumberFormat('@').setValue(value);
    var rule = Array.isArray(list)
      ? SpreadsheetApp.newDataValidation().requireValueInList(list, true)
      : SpreadsheetApp.newDataValidation().requireValueInRange(list, true);
    cell.setDataValidation(rule.setAllowInvalid(false).build());
    rg.setBackground(INPUT).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, INK, SpreadsheetApp.BorderStyle.SOLID);
  };
  var J = "'累計入会データ'!$V$2:$V$8000";
  sh.getRange('AH1').setFormula('=ARRAYFORMULA(TEXT(EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1-SEQUENCE(DATEDIF(MINIFS(' + J + ',' + J + ',">0"),DATE(YEAR(TODAY()),MONTH(TODAY()),1),"M")+1)),"yyyy年m月"))');
  dropdown('B4:C4', KAIGI_MODES_, KAIGI_MODES_[0]);
  dropdown('E4:F4', sh.getRange('AH1:AH200'), oldMonth);
  dropdown('H4:I4', KAIGI_CMP_, oldCmp);

  // 期間（手入力のときだけ書き換える）
  sh.setRowHeight(5, 20);
  sh.getRange('B5').setValue('今回の期間').setFontSize(9).setFontColor(MUTE);
  sh.getRange('H5').setValue('比較の期間').setFontSize(9).setFontColor(MUTE);
  sh.setRowHeight(6, 26);
  ['B6', 'C6', 'H6', 'I6'].forEach(function (a) {
    sh.getRange(a).setNumberFormat('yyyy/m/d').setHorizontalAlignment('center')
      .setBorder(null, null, true, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange('B6:C6').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange('H6:I6').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange('D6').setFormula('=IF(OR(B6="",C6=""),"",C6-B6+1&"日間")').setFontColor(MUTE).setFontSize(9);
  sh.getRange('J6').setFormula('=IF(OR(H6="",I6=""),"",I6-H6+1&"日間")').setFontColor(MUTE).setFontSize(9);
  sh.setRowHeight(7, 16);

  // 主要指標の表（カードもここを読む）
  var cmpOff = 'IF(' + KAIGI_CELLS_.cmp + '="比較なし","",';
  var metrics = [
    ['期間開始時の会員', 18, '#,##0', '+#,##0;-#,##0;0', 0],
    ['入会', 20, '#,##0', '+#,##0;-#,##0;0', 1],
    ['退会', 21, '#,##0', '+#,##0;-#,##0;0', -1],
    ['純増', 23, '+#,##0;-#,##0;0', '+#,##0;-#,##0;0', 1],
    ['退会率', 24, '0.0%', '+0.0%;-0.0%;0.0%', -1],
    ['入会（1日あたり）', 25, '0.00', '+0.00;-0.00;0.00', 1],
    ['退会（1日あたり）', 26, '0.00', '+0.00;-0.00;0.00', -1],
    ['移籍', 22, '#,##0', '+#,##0;-#,##0;0', 0],
    ['再入会', 27, '#,##0', '+#,##0;-#,##0;0', 0],
    ['短期退会（6ヶ月以内）', 29, '#,##0', '+#,##0;-#,##0;0', -1]
  ];
  var T0 = 15;
  var tableHead = function (row, col, labels) {
    sh.getRange(row, col, 1, labels.length).setValues([labels]).setBackground(INK).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
    sh.getRange(row, col).setHorizontalAlignment('left');
  };
  var section = function (a1, text) {
    sh.getRange(a1).setValue(text).setFontSize(12).setFontWeight('bold');
  };
  var cmpHead = '=IF(' + KAIGI_CELLS_.cmp + '="比較なし","比較",' + KAIGI_CELLS_.cmp + ')';
  section('B13', '主要指標');
  tableHead(14, 2, ['指標', '今回', '比較', '差']);
  sh.getRange(14, 4).setFormula(cmpHead);
  var rules = [];
  metrics.forEach(function (m, i) {
    var row = T0 + i;
    sh.getRange(row, 2).setValue(m[0]);
    sh.getRange(row, 3).setFormula('=' + G + 'B' + m[1]).setNumberFormat(m[2]).setFontWeight('bold');
    sh.getRange(row, 4).setFormula('=' + cmpOff + G + 'C' + m[1] + ')').setNumberFormat(m[2]).setFontColor(MUTE);
    sh.getRange(row, 5).setFormula('=IF(OR(C' + row + '="",D' + row + '=""),"",C' + row + '-D' + row + ')').setNumberFormat(m[3]);
    if (m[4] === -1) rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(RED).setBold(true).setRanges([sh.getRange(row, 5)]).build());
    if (m[4] === 1) rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(RED).setBold(true).setRanges([sh.getRange(row, 5)]).build());
  });
  sh.getRange(T0, 3, metrics.length, 3).setHorizontalAlignment('right');
  sh.getRange(T0, 2, metrics.length, 4).setBorder(null, null, true, null, null, true, LINE, SpreadsheetApp.BorderStyle.SOLID);

  // 大きい数字のカード
  var rowOf = function (name) { for (var i = 0; i < metrics.length; i++) if (metrics[i][0] === name) return T0 + i; return 0; };
  var cards = [
    ['B', 'C', '入会', '入会'],
    ['D', 'E', '退会', '退会'],
    ['F', 'G', '純増', '純増'],
    ['H', 'I', '退会率', '退会率'],
    ['J', 'K', '期間開始時の会員', '期間開始時の会員'],
    ['L', 'M', '短期退会（6ヶ月以内）', '短期退会（6ヶ月以内）']
  ];
  sh.setRowHeight(8, 22);
  sh.setRowHeight(9, 48);
  sh.setRowHeight(10, 22);
  sh.setRowHeight(11, 16);
  cards.forEach(function (cd) {
    var row = rowOf(cd[3]);
    var m = metrics[row - T0];
    var span = cd[0] + '8:' + cd[1] + '8';
    var lab = sh.getRange(span); if (cd[0] !== cd[1]) lab.merge();
    lab.setValue(cd[2]).setFontSize(9).setFontColor(MUTE).setFontWeight('bold').setHorizontalAlignment('left');
    var val = sh.getRange(cd[0] + '9:' + cd[1] + '9'); if (cd[0] !== cd[1]) val.merge();
    val.setFormula('=C' + row).setNumberFormat(m[2]).setFontSize(26).setFontWeight('bold').setHorizontalAlignment('left');
    var sub = sh.getRange(cd[0] + '10:' + cd[1] + '10'); if (cd[0] !== cd[1]) sub.merge();
    var diffText = m[1] === 24
      ? 'TEXT(E' + row + '*100,"+0.0;-0.0;0.0")&"pt"'
      : 'TEXT(E' + row + ',"' + m[3] + '")';
    sub.setFormula('=IF(D' + row + '="","",D14&" "&TEXT(D' + row + ',"' + m[2] + '")&"（"&' + diffText + '&"）")')
      .setFontSize(9).setFontColor(MUTE).setHorizontalAlignment('left');
    var box = sh.getRange(cd[0] + '8:' + cd[1] + '10').setBackground(SOFT);
    box.setBorder(true, null, null, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    box.setBorder(null, null, null, true, null, null, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.getRange(cd[0] + '8:' + cd[0] + '10').setBorder(null, true, null, null, null, null, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
    var subCell = sh.getRange(cd[0] + '10');
    if (m[4] === -1) rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=N($E$' + row + ')>0').setFontColor(RED).setBold(true).setRanges([subCell]).build());
    if (m[4] === 1) rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=N($E$' + row + ')<0').setFontColor(RED).setBold(true).setRanges([subCell]).build());
  });

  // 内訳の表（構成比はバー）
  var breakdown = function (top, col, title, rowsSpec, withBar) {
    section(sh.getRange(top, col).getA1Notation(), title);
    var head = withBar ? ['区分', '今回', '比較', '差', '構成比', ''] : ['区分', '今回', '比較', '差'];
    tableHead(top + 1, col, head);
    sh.getRange(top + 1, col + 2).setFormula(cmpHead);
    var first = top + 2, n = rowsSpec.length, totalRow = first + n;
    var L = function (k) { return sh.getRange(1, col + k).getA1Notation().replace(/\d+/, ''); };
    rowsSpec.forEach(function (s, i) {
      var row = first + i;
      if (typeof s[0] === 'number') sh.getRange(row, col).setFormula('=' + G + 'A' + s[0]);
      else sh.getRange(row, col).setValue(s[0]);
      sh.getRange(row, col + 1).setFormula('=' + G + 'B' + s[1]).setNumberFormat('#,##0').setFontWeight('bold');
      sh.getRange(row, col + 2).setFormula('=' + cmpOff + G + 'C' + s[1] + ')').setNumberFormat('#,##0').setFontColor(MUTE);
      sh.getRange(row, col + 3).setFormula('=IF(OR(' + L(1) + row + '="",' + L(2) + row + '=""),"",' + L(1) + row + '-' + L(2) + row + ')').setNumberFormat('+#,##0;-#,##0;0');
      if (withBar) {
        sh.getRange(row, col + 4).setFormula('=IFERROR(' + L(1) + row + '/' + L(1) + '$' + totalRow + ',"")').setNumberFormat('0%').setFontColor(MUTE);
        sh.getRange(row, col + 5).setFormula('=IF(' + L(4) + row + '="","",SPARKLINE(' + L(4) + row + ',{"charttype","bar";"max",1;"color1","#111111"}))');
      }
    });
    sh.getRange(totalRow, col).setValue('合計').setFontWeight('bold');
    sh.getRange(totalRow, col + 1).setFormula('=SUM(' + L(1) + first + ':' + L(1) + (totalRow - 1) + ')').setNumberFormat('#,##0').setFontWeight('bold');
    sh.getRange(totalRow, col + 2).setFormula('=IF(COUNT(' + L(2) + first + ':' + L(2) + (totalRow - 1) + ')=0,"",SUM(' + L(2) + first + ':' + L(2) + (totalRow - 1) + '))').setNumberFormat('#,##0').setFontColor(MUTE);
    sh.getRange(totalRow, col + 3).setFormula('=IF(OR(' + L(1) + totalRow + '="",' + L(2) + totalRow + '=""),"",' + L(1) + totalRow + '-' + L(2) + totalRow + ')').setNumberFormat('+#,##0;-#,##0;0');
    var w = head.length;
    sh.getRange(first, col + 1, n + 1, 3).setHorizontalAlignment('right');
    sh.getRange(first, col, n, w).setBorder(null, null, true, null, null, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(totalRow, col, 1, w).setBackground(SOFT).setBorder(true, null, true, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(N($' + L(1) + first + ')=0,N($' + L(2) + first + ')=0)')
      .setFontColor('#BDBDBD').setRanges([sh.getRange(first, col, n, w)]).build());
    return totalRow;
  };
  breakdown(13, 8, '入会した人の年代', [['10代', 34], ['20代', 35], ['30代', 36], ['40代', 37], ['50代', 38], ['60代', 39], ['70代以上', 40]], true);
  breakdown(26, 2, '入会した人の男女', [['男', 44], ['女', 45]], false);
  breakdown(26, 8, '辞めた人の在籍期間', [['3ヶ月以内', 70], ['4〜6ヶ月', 71], ['7〜12ヶ月', 72], ['1〜2年', 73], ['2〜3年', 74], ['3年超', 75]], true);
  section('B32', '退会理由（人数がある理由だけ）');
  tableHead(33, 2, ['理由', '今回', '比較', '差']);
  sh.getRange('D33').setFormula(cmpHead);
  sh.getRange('B34').setValue('合計').setFontWeight('bold');
  sh.getRange('C34').setFormula('=SUM(' + G + 'B50:B65)').setNumberFormat('#,##0').setFontWeight('bold');
  sh.getRange('D34').setFormula('=' + cmpOff + 'SUM(' + G + 'C50:C65))').setNumberFormat('#,##0').setFontColor(MUTE);
  sh.getRange('E34').setFormula('=IF(OR(C34="",D34=""),"",C34-D34)').setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange('B34:E34').setBackground(SOFT).setBorder(null, null, true, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B35').setFormula('=ARRAYFORMULA(LET(lab,' + G + 'A50:A65,cur,' + G + 'B50:B65,prv,IF(' + KAIGI_CELLS_.cmp + '="比較なし",IF(SEQUENCE(16),""),' + G + 'C50:C65),' +
    'k,IFERROR(cur*1,0)+IFERROR(prv*1,0)>0,IFERROR(FILTER(HSTACK(lab,cur,prv,IF(prv="","",cur-prv)),k),"")))');
  sh.getRange('C35:C50').setNumberFormat('#,##0').setFontWeight('bold');
  sh.getRange('D35:D50').setNumberFormat('#,##0').setFontColor(MUTE);
  sh.getRange('E35:E50').setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange('C34:E50').setHorizontalAlignment('right');
  sh.getRange('B35:E50').setBorder(null, null, null, null, null, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.setConditionalFormatRules(rules);
  sh.hideColumns(34);

  agg.getRange('B6:E7').setFormulas([
    ["='会議用'!B6", "='会議用'!H6", '', ''],
    ["='会議用'!C6", "='会議用'!I6", '', '']
  ]);
  kaigiApplyPeriod_(sh);
  protectInputsOnly_(sh, [KAIGI_CELLS_.mode, KAIGI_CELLS_.month, KAIGI_CELLS_.cmp, KAIGI_CELLS_.cur, KAIGI_CELLS_.prev]);
  props.deleteProperty('KAIGI_MON');
  props.deleteProperty('KAIGI_FROM');
  props.deleteProperty('KAIGI_SURVEY');
  props.deleteProperty('KAIGI_BYMON');
  props.setProperty('KAIGI_V', KAIGI_V_);
  return { ok: true };
}

/**
 * 退会アンケート（管理画面のCSV）。1人1行に直して「退会アンケート」シートへ。会員番号で上書き、無い人は残す。
 * rows: [会員番号, 氏名, 利用頻度(q7), 回数(q8), 時間帯(q11), 解除理由(q12)]
 * 理由・頻度は表記ゆれや昔の選択肢をまとめた「まとめ」列も作る。退会月は累計退会データから（最終在籍月）。
 */
var LEAVE_SURVEY_SHEET_ = '退会アンケート';
var LEAVE_CLUBS_ = ['エニタイム', 'ルネサンス', 'ファストジム', 'アーバンクラシック', 'パーソナルジム', 'ピラティスミラー', 'LAVA', 'その他'];
var LEAVE_REASONS_ = ['引越し・転勤', '仕事に専念・仕事多忙', '金銭的理由（会費が高い）', '一時的に来られなくなる', '体調不良・入院・ケガ・病気',
  '飽きた', '学業専念', '交通不便', '看病・家事・育児', '妊娠', '設備に不満', '混雑（FWエリア）', 'スタッフが不満', 'キャンペーン終了']
  .concat(LEAVE_CLUBS_.map(function (c) { return '移籍：' + c; }));
var LEAVE_FREQS_ = ['週0回', '週1回未満', '週1回', '週2回', '週3回', '週4回', '週5回以上', '不明'];

function leaveReasonGroup_(a) {
  var s = String(a || '').replace(/\s/g, '');
  if (!s) return ['', ''];
  var m = s.match(/^他クラブ移籍(?:[（(](.+)[)）])?$/);
  if (m) {
    var club = LEAVE_CLUBS_.filter(function (c) { return m[1] && m[1].toUpperCase().indexOf(c.toUpperCase()) >= 0; })[0] || 'その他';
    return ['移籍：' + club, m[1] || 'その他'];
  }
  var map = [[/引越|転勤/, 0], [/仕事/, 1], [/金銭|会費/, 2], [/一時的/, 3], [/体調|入院|ケガ|病気/, 4], [/飽き/, 5], [/学業/, 6],
    [/交通/, 7], [/看病|家事|育児/, 8], [/妊娠/, 9], [/設備/, 10], [/混/, 11], [/スタッフ/, 12], [/キャンペーン/, 13]];
  for (var i = 0; i < map.length; i++) if (map[i][0].test(s)) return [LEAVE_REASONS_[map[i][1]], ''];
  return ['その他', ''];
}

function leaveFreqGroup_(q7, q8) {
  var a = String(q7 || '').trim();
  var m = a.match(/^週(\d)回/);
  if (m) return 'w' + m[1] === 'w5' ? '週5回以上' : '週' + m[1] + '回';
  if (/週5回以上/.test(a)) return '週5回以上';
  var n = Number(String(q8 || '').replace(/[０-９]/g, function (d) { return String.fromCharCode(d.charCodeAt(0) - 0xFEE0); }).trim());
  if (!isFinite(n) || String(q8 || '').trim() === '') return '不明';
  if (a === '週') return n <= 0 ? '週0回' : n < 1 ? '週1回未満' : n >= 5 ? '週5回以上' : '週' + Math.floor(n) + '回';
  if (a === '月') {
    if (n <= 0) return '週0回';
    if (n < 4) return '週1回未満';
    var w = Math.floor(n / 4);
    return w >= 5 ? '週5回以上' : '週' + w + '回';
  }
  return '不明';
}

function importLeaveSurvey_(rows) {
  var ss = SpreadsheetApp.openById(WS_CONFIG.SPREADSHEET_ID);
  var sh = ss.getSheetByName(LEAVE_SURVEY_SHEET_) || ss.insertSheet(LEAVE_SURVEY_SHEET_);
  var head = ['会員番号', '氏名', '退会理由（まとめ）', '移籍先', '退会理由（回答そのまま）', '利用頻度（まとめ）', '利用頻度（回答そのまま）', '利用時間帯', '退会月（最終在籍月）'];
  var keep = {};
  var last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 8).getValues().forEach(function (r) { if (r[0] !== '') keep[String(r[0])] = r; });
  var added = 0, updated = 0;
  rows.forEach(function (r) {
    var id = String(r[0] || '').replace(/^0+/, '');
    if (!id) return;
    var g = leaveReasonGroup_(r[5]);
    var q7 = String(r[2] || ''), q8 = String(r[3] || '');
    var rawF = /^週\d|週5回以上/.test(q7) ? q7 : (q7 + (q8 ? ' ' + q8 : '')).trim();
    if (keep[id]) updated++; else added++;
    keep[id] = [id, String(r[1] || ''), g[0], g[1], String(r[5] || ''), leaveFreqGroup_(q7, q8), rawF, String(r[4] || '').split('\t').join('、')];
  });
  var out = Object.keys(keep).map(function (k) { return keep[k]; });
  out.sort(function (a, b) { return String(a[0]) < String(b[0]) ? -1 : 1; });
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setBackground('#111111').setFontColor('#FFFFFF').setFontWeight('bold');
  if (out.length) {
    sh.getRange(2, 1, out.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, out.length, 8).setValues(out);
  }
  sh.getRange('I2').setFormula('=MAP(A2:A,LAMBDA(a,IF(a="","",LET(m,MAXIFS(\'累計退会データ\'!$U$2:$U$8000,\'累計退会データ\'!$T$2:$T$8000,a&""),IF(m=0,"",EDATE(m,-1))))))');
  sh.getRange('I2:I').setNumberFormat('yyyy/m');
  sh.getRange(1, 1, Math.max(out.length + 1, 2), head.length).setFontFamily('Meiryo');
  sh.setFrozenRows(1);
  [90, 110, 170, 110, 200, 110, 140, 200, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  PropertiesService.getDocumentProperties().setProperty('LEAVE_SURVEY_AT', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
  return { ok: true, total: out.length, added: added, updated: updated };
}

/** メニュー「ファイルUP」→ 数値更新ファイル。ブラウザでCSVを読んで uploadNumbersFile に渡す */
function openNumbersUpload() {
  var html = HtmlService.createHtmlOutputFromFile('numbersui').setWidth(560).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, '数値更新ファイルUP');
}

/**
 * 数値更新ファイル1つ分。{name, grid}（CSVはブラウザで表にして送る）または {name, base64}（Excel）。
 * 中身で判定：退会アンケート（uid・q_id・answer）／入会手続き一覧表（利用開始年月）／退会手続き一覧表（退会年月）。
 */
function uploadNumbersFile(file) {
  try {
    var grid = file && file.grid ? file.grid : gessho3ReadExcel_(file || {});
    var res = numbersImportGrid_(grid, String(file && file.name || ''));
    var props = PropertiesService.getDocumentProperties();
    if (res.ok && res.kind !== 'leaveSurvey') props.setProperty('CUMULATIVE_AT', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
    try { syncTopRefreshStatus_(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) {}
    return res;
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function numbersNorm_(v) { return String(v == null ? '' : v).replace(/[\s\u3000]/g, ''); }

function numbersImportGrid_(grid, name) {
  var hr = -1;
  for (var r = 0; r < Math.min(grid.length, 15) && hr < 0; r++) {
    var t = grid[r].map(numbersNorm_);
    if (t.indexOf('uid') >= 0 && t.indexOf('q_id') >= 0) {
      var iu = t.indexOf('uid'), iname = t.indexOf('name'), iq = t.indexOf('q_id'), ia = t.indexOf('answer');
      var people = {}, order = [];
      grid.slice(r + 1).forEach(function (x) {
        if (!x || !x[iu]) return;
        if (!people[x[iu]]) { people[x[iu]] = { name: x[iname] || '' }; order.push(x[iu]); }
        people[x[iu]][String(x[iq])] = x[ia];
      });
      var rows = order.map(function (u) { var d = people[u]; return [u, d.name, d['7'] || '', d['8'] || '', d['11'] || '', d['12'] || '']; })
        .filter(function (x) { return x[5] || x[2]; });
      var res = importLeaveSurvey_(rows);
      res.kind = 'leaveSurvey'; res.label = '退会アンケート';
      return res;
    }
    if (t.indexOf('会員番号') >= 0) hr = r;
  }
  if (hr < 0) throw new Error('会員番号の見出しが見つからないため、種類を判定できません: ' + name);
  var head = grid[hr].map(numbersNorm_);
  var isLeave = head.some(function (h) { return h.indexOf('退会年月') >= 0 || h.indexOf('退会届出日') >= 0; });
  var isJoin = head.some(function (h) { return h.indexOf('利用開始') >= 0; });
  if (!isLeave && !isJoin) throw new Error('入会・退会どちらの手続き一覧表か判定できません（利用開始年月・退会年月の列がありません）: ' + name);
  return appendCumulative_(isLeave ? '累計退会データ' : '累計入会データ', grid, hr, name);
}

/** 手続き一覧表を累計シートの末尾へ足す。同じ人・同じ手続き（会員番号＋届出日／退会年月）はとばす。T列より右の自動列は触らない */
function appendCumulative_(sheetName, grid, hr, fileName) {
  var ss = SpreadsheetApp.openById(WS_CONFIG.SPREADSHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(sheetName + 'がありません');
  var leave = sheetName === '累計退会データ';
  var width = leave ? 19 : 15;
  var target = sh.getRange(1, 1, 1, width).getValues()[0].map(numbersNorm_);
  var src = grid[hr].map(numbersNorm_);
  var map = target.map(function (h) { return src.indexOf(h); });
  var matched = map.filter(function (i) { return i >= 0; }).length;
  var colNo = target.indexOf('会員番号');
  if (map[colNo] < 0 || matched < 8) throw new Error(fileName + '：列の見出しが' + sheetName + 'と合いません（一致 ' + matched + '列）。');
  var dateCols = {}, ymCols = {};
  target.forEach(function (h, i) {
    if (/年月$/.test(h)) ymCols[i] = true;
    else if (/日$/.test(h)) dateCols[i] = true;
  });
  var keyCol = target.indexOf(leave ? '退会年月' : '届出日');
  var fmt = function (v, i) {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', ymCols[i] ? 'yyyy/MM' : 'yyyy/MM/dd');
    if (typeof v === 'number' && (dateCols[i] || ymCols[i]) && v > 20000 && v < 80000) {
      var d = new Date(Math.round((v - 25569) * 86400000));
      return Utilities.formatDate(d, 'UTC', ymCols[i] ? 'yyyy/MM' : 'yyyy/MM/dd');
    }
    return v == null ? '' : v;
  };
  var keyOf = function (no, k) {
    var n = String(no).replace(/\D/g, '').replace(/^0+/, '');
    return n + '|' + String(k).replace(/\D/g, '').slice(0, ymCols[keyCol] ? 6 : 8);
  };
  var lastRow = 1;
  var colF = sh.getRange(1, colNo + 1, Math.max(sh.getMaxRows(), 2), 1).getValues();
  for (var i = colF.length - 1; i >= 1; i--) if (String(colF[i][0]) !== '') { lastRow = i + 1; break; }
  var have = {};
  if (lastRow >= 2) {
    var ex = sh.getRange(2, 1, lastRow - 1, width).getValues();
    ex.forEach(function (row) { have[keyOf(row[colNo], fmt(row[keyCol], keyCol))] = true; });
  }
  var carry = {};
  var out = [], skipped = 0;
  grid.slice(hr + 1).forEach(function (row) {
    var vals = map.map(function (si, i) { return si < 0 ? '' : fmt(row[si], i); });
    [0, 1, 2].forEach(function (i) { if (vals[i] !== '') carry[i] = vals[i]; else if (carry[i]) vals[i] = carry[i]; });
    var no = String(vals[colNo]).replace(/\s/g, '');
    if (!/\d{3,}/.test(no)) return;
    var k = keyOf(no, vals[keyCol]);
    if (have[k]) { skipped++; return; }
    have[k] = true;
    out.push(vals);
  });
  if (out.length) {
    if (sh.getMaxRows() < lastRow + out.length) sh.insertRowsAfter(sh.getMaxRows(), lastRow + out.length - sh.getMaxRows());
    sh.getRange(lastRow + 1, colNo + 1, out.length, 1).setNumberFormat('@');
    sh.getRange(lastRow + 1, 1, out.length, width).setValues(out);
  }
  return { ok: true, kind: leave ? 'leave' : 'join', label: leave ? '退会手続き一覧表' : '入会手続き一覧表', added: out.length, skipped: skipped, total: lastRow - 1 + out.length };
}

/**
 * 会議用の退会理由を退会アンケートの文字に置き換え、利用頻度と、月ごとの件数（O30〜）を足す。
 * 退会月は最終在籍月（会議用の退会と同じ数え方）。
 */
function addKaigiSurvey_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('KAIGI_V') !== KAIGI_V_ || props.getProperty('KAIGI_MON') !== 'v1' || props.getProperty('KAIGI_SURVEY') === 'v4') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(KAIGI_SHEET_);
  if (!sh || !ss.getSheetByName(LEAVE_SURVEY_SHEET_)) return { ok: false, reason: 'no survey sheet' };
  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#E3E3E3', SOFT = '#F7F7F7';
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;
  var S = "'" + LEAVE_SURVEY_SHEET_ + "'!";
  var mon = S + '$I$2:$I', rea = S + '$C$2:$C', frq = S + '$F$2:$F';
  var curS = 'DATE(YEAR($B$6),MONTH($B$6),1)', curE = '$C$6';
  var cmpS = 'DATE(YEAR($H$6),MONTH($H$6),1)', cmpE = '$I$6';
  var off = KAIGI_CELLS_.cmp + '="比較なし"';
  var cnt = function (s, e, col, lab) { return 'COUNTIFS(' + mon + ',">="&' + s + ',' + mon + ',"<="&' + e + ',' + col + ',' + lab + ')'; };
  var cntAll = function (s, e, col) { return 'COUNTIFS(' + mon + ',">="&' + s + ',' + mon + ',"<="&' + e + ',' + col + ',"<>",' + col + ',"<>その他")'; };
  var head = function (rg) { return rg.setBackground(INK).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center'); };
  var cmpHead = '=IF(' + off + ',"比較",' + KAIGI_CELLS_.cmp + ')';

  var table = function (top, col, title, labels, colRef, withBar) {
    var L = function (k) { return sh.getRange(1, col + k).getA1Notation().replace(/\d+/, ''); };
    var w = withBar ? 6 : 4;
    sh.getRange(top, col, labels.length + 4, w).clear();
    sh.getRange(top, col).setValue(title).setFontSize(12).setFontWeight('bold');
    var h = withBar ? ['区分', '今回', '比較', '差', '構成比', ''] : ['理由', '今回', '比較', '差'];
    head(sh.getRange(top + 1, col, 1, w).setValues([h]));
    sh.getRange(top + 1, col).setHorizontalAlignment('left');
    sh.getRange(top + 1, col + 2).setFormula(cmpHead);
    var tot = top + 2, first = top + 3, lastR = first + labels.length - 1;
    sh.getRange(tot, col).setValue('合計（回答した人）');
    sh.getRange(tot, col + 1).setFormula('=' + cntAll(curS, curE, colRef));
    sh.getRange(tot, col + 2).setFormula('=IF(OR(' + off + ',$H$6=""),"",' + cntAll(cmpS, cmpE, colRef) + ')');
    sh.getRange(tot, col + 3).setFormula('=IF(' + L(2) + tot + '="","",' + L(1) + tot + '-' + L(2) + tot + ')');
    var f = labels.map(function (lab, i) {
      var r = first + i;
      var row = [lab, '=' + cnt(curS, curE, colRef, '$' + L(0) + r), '=IF(OR(' + off + ',$H$6=""),"",' + cnt(cmpS, cmpE, colRef, '$' + L(0) + r) + ')',
        '=IF(' + L(2) + r + '="","",' + L(1) + r + '-' + L(2) + r + ')'];
      if (withBar) row.push('=IFERROR(' + L(1) + r + '/' + L(1) + '$' + tot + ',"")', '=IF(N(' + L(4) + r + ')=0,"",SPARKLINE(' + L(4) + r + ',{"charttype","bar";"max",1;"color1","#111111"}))');
      return row;
    });
    sh.getRange(first, col, labels.length, w).setFormulas(f);
    sh.getRange(first, col, labels.length, 1).setValues(labels.map(function (x) { return [x]; }));
    var all = sh.getRange(tot, col, labels.length + 1, w).setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(tot, col + 1, labels.length + 1, 1).setNumberFormat('#,##0').setFontWeight('bold');
    sh.getRange(tot, col + 2, labels.length + 1, 1).setNumberFormat('#,##0').setFontColor(MUTE);
    sh.getRange(tot, col + 3, labels.length + 1, 1).setNumberFormat('+#,##0;-#,##0;0');
    if (withBar) sh.getRange(tot, col + 4, labels.length + 1, 1).setNumberFormat('0%').setFontColor(MUTE);
    sh.getRange(tot, col + 1, labels.length + 1, 4).setHorizontalAlignment('right');
    sh.getRange(first, col, labels.length, w).setBorder(null, null, true, null, null, true, LINE, SOLID);
    sh.getRange(tot, col, 1, w).setBackground(SOFT).setFontWeight('bold').setBorder(null, null, true, null, null, null, INK, SOLID);
    return { first: first, last: lastR, w: w, tot: tot };
  };

  // 退会理由：回答率が期間で違うので、件数より「割合」で比べる。割合の差が大きい理由を下に要約
  var reasonTable = function (top) {
    var labels = LEAVE_REASONS_, n = labels.length;
    sh.getRange(top, 2).setValue('退会理由の比較（退会アンケート）').setFontSize(12).setFontWeight('bold');
    head(sh.getRange(top + 1, 2, 1, 6).setValues([['理由', '今回', '割合', '比較', '割合', '割合の差']]));
    sh.getRange(top + 1, 2).setHorizontalAlignment('left');
    sh.getRange(top + 1, 5).setFormula(cmpHead);
    var tot = top + 2, first = top + 3, last = first + n - 1;
    var cmpOn = 'OR(' + off + ',$H$6="")';
    sh.getRange(tot, 2, 1, 6).setFormulas([[
      '="合計（回答した人）"', '=' + cntAll(curS, curE, rea), '=IF(C' + tot + '=0,"",1)',
      '=IF(' + cmpOn + ',"",' + cntAll(cmpS, cmpE, rea) + ')', '=IF(N(E' + tot + ')=0,"",1)', '=""'
    ]]);
    var f = labels.map(function (lab, i) {
      var r = first + i;
      return [lab, '=' + cnt(curS, curE, rea, '$B' + r), '=IFERROR(C' + r + '/C$' + tot + ',"")',
        '=IF(' + cmpOn + ',"",' + cnt(cmpS, cmpE, rea, '$B' + r) + ')', '=IFERROR(E' + r + '/E$' + tot + ',"")',
        '=IF(OR(D' + r + '="",F' + r + '=""),"",D' + r + '-F' + r + ')'];
    });
    sh.getRange(first, 2, n, 6).setFormulas(f);
    sh.getRange(first, 2, n, 1).setValues(labels.map(function (x) { return [x]; }));
    sh.getRange(tot, 2, 1, 1).setValue('合計（回答した人）');
    var all = sh.getRange(tot, 2, n + 1, 6).setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(tot, 3, n + 1, 1).setNumberFormat('#,##0').setFontWeight('bold');
    sh.getRange(tot, 4, n + 1, 1).setNumberFormat('0%').setFontWeight('bold');
    sh.getRange(tot, 5, n + 1, 1).setNumberFormat('#,##0').setFontColor(MUTE);
    sh.getRange(tot, 6, n + 1, 1).setNumberFormat('0%').setFontColor(MUTE);
    sh.getRange(tot, 7, n + 1, 1).setNumberFormat('+0.0%;-0.0%;0.0%').setBackground(SOFT);
    sh.getRange(tot, 3, n + 1, 5).setHorizontalAlignment('right');
    sh.getRange(first, 2, n, 6).setBorder(null, null, true, null, null, true, LINE, SOLID);
    sh.getRange(tot, 2, 1, 6).setBackground(SOFT).setFontWeight('bold').setBorder(null, null, true, null, null, null, INK, SOLID);
    [first + 13].forEach(function (r) { sh.getRange(r, 2, 1, 6).setBorder(null, null, true, null, null, null, '#BDBDBD', SOLID); });

    var note = last + 1;
    sh.getRange('B' + note).setFormula('=IFERROR("回答 "&C' + tot + '&"人／退会 "&C17&"人（回答率 "&TEXT(C' + tot + '/C17,"0%")&"）"&IF(N(E' + tot + ')>0,"　比較 "&E' + tot + '&"人／"&D17&"人（"&TEXT(E' + tot + '/D17,"0%")&"）",""),"")')
      .setFontSize(9).setFontColor(MUTE);
    var G = 'G' + first + ':G' + last, B = 'B' + first + ':B' + last;
    var top3 = function (asc, cond) {
      var k = '(' + G + '<>"")*(' + G + cond + ')';
      return 'IFERROR(TEXTJOIN("、",TRUE,MAP(SORTN(FILTER(' + B + ',' + k + '),3,0,FILTER(' + G + ',' + k + '),' + asc + '),' +
        'LAMBDA(x,x&" "&TEXT(XLOOKUP(x,' + B + ',' + G + ')*100,"+0.0;-0.0")&"pt"))),"特になし")';
    };
    sh.getRange('B' + (note + 1)).setValue('増えた理由').setFontWeight('bold').setFontColor('#B91C1C').setFontSize(10);
    sh.getRange('C' + (note + 1)).setFormula('=IF(N(E' + tot + ')=0,"比較なし",' + top3('FALSE', '>=0.02') + ')').setFontSize(10);
    sh.getRange('B' + (note + 2)).setValue('減った理由').setFontWeight('bold').setFontSize(10);
    sh.getRange('C' + (note + 2)).setFormula('=IF(N(E' + tot + ')=0,"比較なし",' + top3('TRUE', '<=-0.02') + ')').setFontSize(10);
    sh.getRange('B' + (note + 1) + ':B' + (note + 2)).setFontFamily('Meiryo');
    sh.getRange('C' + (note + 1) + ':C' + (note + 2)).setFontFamily('Meiryo');
    var rr = sh.getConditionalFormatRules().filter(function (ru) {
      var g0 = ru.getRanges()[0];
      return !(g0 && g0.getRow() >= top && g0.getColumn() >= 2 && g0.getColumn() <= 7);
    });
    rr.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.02).setFontColor('#B91C1C').setBold(true).setRanges([sh.getRange(G)]).build());
    rr.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($D' + first + '),$D' + first + '>=0.15)').setBackground('#FDF1F1').setRanges([sh.getRange('B' + first + ':D' + last)]).build());
    sh.setConditionalFormatRules(rr);
    return { first: first, last: last, tot: tot };
  };
  sh.getRange('B32:G64').clear();
  sh.getRange('O30:AB75').clear();
  var t1 = reasonTable(32);
  var t2 = table(37, 8, '辞めた人の利用頻度（退会アンケート）', LEAVE_FREQS_, frq, true);

  var matrix = function (top, title, labels, colRef) {
    sh.getRange(top, 15, labels.length + 3, 14).clear();
    sh.getRange(top, 15).setValue(title).setFontSize(12).setFontWeight('bold');
    var hr = top + 1, first = top + 2, tot = first + labels.length;
    var hdr = ['=IF(TRUE,"' + (colRef === rea ? '理由' : '利用頻度') + '")'];
    for (var i = 0; i < 12; i++) hdr.push('=IF($AJ$' + (2 + i) + '="","",$AJ$' + (2 + i) + ')');
    hdr.push('="合計"');
    sh.getRange(hr, 15, 1, 14).setFormulas([hdr]);
    head(sh.getRange(hr, 15, 1, 14)).setNumberFormat('yyyy/m');
    sh.getRange(hr, 15).setHorizontalAlignment('left');
    var cols = 'PQRSTUVWXYZ'.split('').concat(['AA']);
    var f = labels.map(function (lab, k) {
      var r = first + k;
      var row = [lab];
      cols.forEach(function (c) { row.push('=IF(' + c + '$' + hr + '="","",COUNTIFS(' + mon + ',' + c + '$' + hr + ',' + colRef + ',$O' + r + '))'); });
      row.push('=SUM(P' + r + ':AA' + r + ')');
      return row;
    });
    sh.getRange(first, 15, labels.length, 14).setFormulas(f);
    var tr = ['合計'];
    cols.concat(['AB']).forEach(function (c) { tr.push('=IF(' + c + '$' + hr + '="","",SUM(' + c + first + ':' + c + (tot - 1) + '))'); });
    sh.getRange(tot, 15, 1, 14).setFormulas([tr]);
    sh.getRange(first, 15, labels.length, 1).setValues(labels.map(function (x) { return [x]; }));
    sh.getRange(tot, 15).setValue('合計');
    sh.getRange(first, 15, labels.length + 1, 14).setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(first, 16, labels.length + 1, 13).setNumberFormat('0').setHorizontalAlignment('right');
    sh.getRange(first, 28, labels.length + 1, 1).setFontWeight('bold').setBackground(SOFT);
    sh.getRange(first, 15, labels.length, 14).setBorder(null, null, true, null, null, true, LINE, SOLID);
    sh.getRange(tot, 15, 1, 14).setBackground(SOFT).setFontWeight('bold').setBorder(true, null, true, null, null, null, INK, SOLID);
    return sh.getRange(first, 16, labels.length + 1, 13);
  };
  var m1 = matrix(30, '退会理由 月ごとの件数（退会アンケート・今回の期間）', LEAVE_REASONS_, rea);
  var m2Top = 30 + LEAVE_REASONS_.length + 5;
  var m2 = matrix(m2Top, '利用頻度 月ごとの件数（退会アンケート・今回の期間）', LEAVE_FREQS_, frq);
  sh.setColumnWidth(28, 64);
  sh.setColumnWidth(15, 160);

  var rules = sh.getConditionalFormatRules().filter(function (ru) {
    var bc = ru.getBooleanCondition();
    return !(bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.NUMBER_EQUAL_TO);
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setFontColor('#C8C8C8').setRanges([m1, m2,
    sh.getRange(t1.first, 3, LEAVE_REASONS_.length, 2), sh.getRange(t2.first, 9, LEAVE_FREQS_.length, 2)]).build());
  sh.setConditionalFormatRules(rules);
  var at = props.getProperty('LEAVE_SURVEY_AT') || '';
  sh.getRange(m2Top + LEAVE_FREQS_.length + 4, 15).setValue('回答した人だけの数字です。「その他」は分析に向かないので数えていません').setFontSize(9).setFontColor(MUTE);
  props.setProperty('KAIGI_SURVEY', 'v4');
  return { ok: true };
}

/**
 * 会議用の12行目から下を「月ごとに横へ並べる」形に作り直す。期間の合計は上のカードだけ。
 * どの表も 1ヶ月＝2列（今回｜比較の同じ順番の月）。表は縦に積む。上のプルダウンと列を共有するので、使わない月の列は隠さず空欄。
 * 月は AJ2:AJ13（今回）・AN2:AN13（比較）、月初・入会・退会は AK:AM／AO:AQ（addKaigiMonthly_ の非表示列）。
 */
var KAIGI_SLOTS_ = 12;
function rebuildKaigiByMonth_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('KAIGI_V') !== KAIGI_V_ || props.getProperty('KAIGI_MON') !== 'v1' || props.getProperty('KAIGI_BYMON') === 'v4') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(KAIGI_SHEET_);
  if (!sh) return { ok: false };
  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#E3E3E3', SOFT = '#F7F7F7', RED = '#B91C1C';
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;
  var C = KAIGI_CELLS_, G = "'分析用_期間集計'!";
  var off = C.cmp + '="比較なし"';
  var LAST_COL = 2 + KAIGI_SLOTS_ * 2;
  var colL = function (c) { return sh.getRange(1, c).getA1Notation().replace(/\d+/, ''); };

  var maxR = sh.getMaxRows();
  if (maxR < 220) sh.insertRowsAfter(maxR, 220 - maxR);
  var body = sh.getRange(12, 1, sh.getMaxRows() - 11, 33);
  try { body.breakApart(); } catch (eB) {}
  body.clearDataValidations();
  body.clear();
  sh.showColumns(3, 31);
  sh.setColumnWidth(2, 180);
  for (var c = 3; c <= LAST_COL; c++) sh.setColumnWidth(c, 58);
  for (var c2 = LAST_COL + 1; c2 <= 33; c2++) sh.setColumnWidth(c2, 18);
  for (var r0 = 12; r0 <= 220; r0++) sh.setRowHeight(r0, 22);

  // カード：期間の合計だけ（分析用_期間集計を直接読む）。差は AS 列（非表示）に置いて色分けに使う
  var cards = [['B', 20, '#,##0', '+#,##0;-#,##0;0', 1], ['D', 21, '#,##0', '+#,##0;-#,##0;0', -1], ['F', 23, '+#,##0;-#,##0;0', '+#,##0;-#,##0;0', 1],
    ['H', 24, '0.0%', '', -1], ['J', 18, '#,##0', '+#,##0;-#,##0;0', 0], ['L', 29, '#,##0', '+#,##0;-#,##0;0', -1]];
  var rules = [];
  cards.forEach(function (cd, i) {
    var dr = 'AS' + (2 + i);
    sh.getRange(dr).setFormula('=IF(OR(' + off + ',' + G + 'C' + cd[1] + '="",' + G + 'B' + cd[1] + '=""),"",' + G + 'B' + cd[1] + '-' + G + 'C' + cd[1] + ')');
    sh.getRange(cd[0] + '9').setFormula('=' + G + 'B' + cd[1]).setNumberFormat(cd[2]);
    var diff = cd[1] === 24 ? 'TEXT(' + dr + '*100,"+0.0;-0.0;0.0")&"pt"' : 'TEXT(' + dr + ',"' + cd[3] + '")';
    var sub = sh.getRange(cd[0] + '10');
    sub.setFormula('=IF(' + dr + '="","",IF(' + C.cmp + '="前年同期","前年","前期間")&" "&TEXT(' + G + 'C' + cd[1] + ',"' + cd[2] + '")&CHAR(10)&"差 "&' + diff + ')')
      .setWrap(true).setFontSize(9).setVerticalAlignment('top');
    if (cd[4] === -1) rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=N($' + dr.replace(/(\d+)/, '$$$1') + ')>0').setFontColor(RED).setBold(true).setRanges([sub]).build());
    if (cd[4] === 1) rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=N($' + dr.replace(/(\d+)/, '$$$1') + ')<0').setFontColor(RED).setBold(true).setRanges([sub]).build());
  });
  sh.setRowHeight(10, 36);
  sh.getRange('L8').setValue('6ヶ月内の退会');
  sh.getRange('B8:M8').setWrap(false).setFontSize(9);
  ['B6', 'C6', 'H6', 'I6'].forEach(function (a) { sh.getRange(a).setNumberFormat('yy/m/d').setFontSize(9).setWrap(false); });
  sh.getRange('B3:M3').setFontSize(8);
  sh.getRange('K3').setValue('開始月（開始月〜対象月）');
  sh.getRange('E3').setValue('対象月（最後の月）');
  sh.getRange('K4:L4').getMergedRanges().length || sh.getRange('K4:L4').merge();
  sh.getRange(C.cmp).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(KAIGI_CMP_, true).setAllowInvalid(false).build());
  sh.getRange('H1').setFormula('=IF(' + C.mode + '="開始月〜対象月",' + C.from + '&"〜"&' + C.month + ',' + C.month + '&IF(' + C.mode + '="月を選ぶ","",IF(' + C.mode + '="手入力","（手入力）","　"&' + C.mode + ')))&"　｜　"&IF(' + C.cmp + '="比較なし","比較なし",IF(' + C.cmp + '="月ごとの増減","前の月からの増減",' + C.cmp + '&"と比較"))');
  var MOM = C.cmp + '="月ごとの増減"';
  var momRows = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B$4<>"開始月〜対象月"')
    .setFontColor('#BDBDBD').setBackground('#F3F3F3').setRanges([sh.getRange('K4:L4')]).build());

  var curM = function (k) { return '$AJ$' + (2 + k); };
  var cmpM = function (k) { return '$AN$' + (2 + k); };
  var slotHead = function (row) {
    var top = [], sub = [];
    for (var k = 0; k < KAIGI_SLOTS_; k++) {
      top.push('=IF(' + curM(k) + '="","",TEXT(' + curM(k) + ',"yyyy年m月"))', '');
      sub.push('=IF(' + curM(k) + '="","","今回")', '=IF(' + curM(k) + '="","",IF(' + MOM + ',' + (k ? '"前月比"' : '""') + ',IF(' + cmpM(k) + '="","",TEXT(' + cmpM(k) + ',"yy年m月"))))');
    }
    sh.getRange(row, LAST_COL + 1, 2, 1).merge().setValue('最初→最後').setBackground(INK).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 3, 1, LAST_COL - 2).setFormulas([top]);
    sh.getRange(row + 1, 3, 1, LAST_COL - 2).setFormulas([sub]);
    for (var k2 = 0; k2 < KAIGI_SLOTS_; k2++) sh.getRange(row, 3 + k2 * 2, 1, 2).merge();
    sh.getRange(row, 2, 2, LAST_COL - 1).setBackground(INK).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
    sh.getRange(row + 1, 4, 1, LAST_COL - 3).setFontColor('#BDBDBD');
    for (var k3 = 1; k3 < KAIGI_SLOTS_; k3++) sh.getRange(row, 3 + k3 * 2, 2, 1).setBorder(null, true, null, null, null, null, '#FFFFFF', SOLID);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=INDEX($C$' + row + ':$' + colL(LAST_COL) + '$' + row + ',1,FLOOR((COLUMN()-3)/2)*2+1)=""')
      .setBackground('#FFFFFF').setFontColor('#FFFFFF').setRanges([sh.getRange(row, 3, 2, LAST_COL - 2)]).build());
  };
  var slotBorders = function (r1, n) {
    for (var k = 1; k < KAIGI_SLOTS_; k++) sh.getRange(r1, 3 + k * 2, n, 1).setBorder(null, true, null, null, null, null, '#BDBDBD', SOLID);
  };
  var title = function (row, text, note) {
    sh.getRange(row, 2).setValue(text).setFontSize(12).setFontWeight('bold');
    if (note) sh.getRange(row, 7).setValue(note).setFontSize(9).setFontColor(MUTE);
  };
  var zeroRanges = [];

  var row = 12;
  // 主要指標
  title(row, '主要指標（月ごと）', '右の列＝比較の月（「月ごとの増減」なら前の月からの増減）。右端＝期間の最初の月→最後の月。赤＝悪くなった月');
  slotHead(row + 1);
  sh.getRange(row + 1, 2, 2, 1).merge().setValue('指標').setHorizontalAlignment('left').setVerticalAlignment('middle');
  var mets = [['月初の会員', 'AM', 'AQ', '#,##0', 0], ['入会', 'AK', 'AO', '#,##0', 1], ['退会', 'AL', 'AP', '#,##0', -1],
    ['純増', 'AK-AL', 'AO-AP', '+#,##0;-#,##0;0', 1], ['退会率', 'AL/AM', 'AP/AQ', '0.0%', -1]];
  var first = row + 3;
  var mf = mets.map(function (m, i) {
    var out = [m[0]];
    for (var k = 0; k < KAIGI_SLOTS_; k++) {
      var h = 2 + k;
      var ex = function (e) { return e.replace(/A[A-Q]/g, function (x) { return x + h; }); };
      out.push('=IF(' + curM(k) + '="","",IFERROR(' + ex(m[1]) + ',""))', '=IF(OR(' + curM(k) + '="",' + cmpM(k) + '=""),"",IFERROR(' + ex(m[2]) + ',""))');
    }
    return out;
  });
  sh.getRange(first, 2, mets.length, LAST_COL - 1).setFormulas(mf);
  sh.getRange(first, 2, mets.length, 1).setValues(mets.map(function (m) { return [m[0]]; }));
  mets.forEach(function (m, i) {
    var r = first + i;
    momRows.push({ r: r, kind: i === 4 ? 'pt' : i === 3 ? 'diff' : 'pct', dir: m[4] });
    sh.getRange(r, 3, 1, LAST_COL - 2).setNumberFormat(m[3]);
    if (m[4] !== 0) {
      for (var k = 0; k < KAIGI_SLOTS_; k++) {
        var a = colL(3 + k * 2) + r, b = colL(4 + k * 2) + r;
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + a + '),ISNUMBER(' + b + '),' + a + (m[4] === 1 ? '<' : '>') + b + ')')
          .setFontColor(RED).setRanges([sh.getRange(a)]).build());
      }
    }
  });
  sh.getRange(first, 3, mets.length, LAST_COL - 2).setHorizontalAlignment('right');
  for (var k4 = 0; k4 < KAIGI_SLOTS_; k4++) {
    sh.getRange(first, 3 + k4 * 2, mets.length, 1).setFontWeight('bold');
    sh.getRange(first, 4 + k4 * 2, mets.length, 1).setFontColor(MUTE);
  }
  sh.getRange(first, 2, mets.length, LAST_COL - 1).setBorder(null, null, true, null, null, true, LINE, SOLID);
  slotBorders(first, mets.length);
  row = first + mets.length + 2;

  // 内訳：件数の段と割合の段を縦に並べる
  var J = "'累計入会データ'!", S = "'" + LEAVE_SURVEY_SHEET_ + "'!";
  var breakdown = function (text, note, labels, countF, opt) {
    opt = opt || {};
    title(row, text, note);
    slotHead(row + 1);
    sh.getRange(row + 1, 2, 2, 1).merge().setValue('区分').setHorizontalAlignment('left').setVerticalAlignment('middle');
    var tot = row + 3, f1 = tot + 1, n = labels.length;
    var extra = opt.extra || [];
    var lab0 = row + 3 + 1 + n + extra.length;
    var f2 = lab0 + 1;
    var lines = [];
    var totRow = [opt.totalLabel || '合計'];
    for (var k = 0; k < KAIGI_SLOTS_; k++) {
      var a = colL(3 + k * 2), b = colL(4 + k * 2);
      totRow.push('=IF(' + curM(k) + '="","",SUM(' + a + f1 + ':' + a + (f1 + n - 1) + '))', '=IF(OR(' + curM(k) + '="",' + cmpM(k) + '="",' + off + '),"",SUM(' + b + f1 + ':' + b + (f1 + n - 1) + '))');
    }
    lines.push(totRow);
    labels.forEach(function (lab, i) {
      var r = f1 + i, out = [lab];
      for (var k = 0; k < KAIGI_SLOTS_; k++) {
        out.push('=IF(' + curM(k) + '="","",' + countF(curM(k), '$B' + r, i) + ')', '=IF(OR(' + curM(k) + '="",' + cmpM(k) + '="",' + off + '),"",' + countF(cmpM(k), '$B' + r, i) + ')');
      }
      lines.push(out);
    });
    extra.forEach(function (ex) {
      var out = [ex[0]];
      for (var k = 0; k < KAIGI_SLOTS_; k++) {
        var a = colL(3 + k * 2), b = colL(4 + k * 2);
        out.push('=IF(' + curM(k) + '="","",' + ex[1](2 + k, 'AJ', a, tot) + ')', '=IF(OR(' + curM(k) + '="",' + cmpM(k) + '="",' + off + '),"",' + ex[1](2 + k, 'AN', b, tot) + ')');
      }
      lines.push(out);
    });
    var bandRow = ['割合（その月の合計に対して）'];
    for (var k5 = 0; k5 < KAIGI_SLOTS_ * 2; k5++) bandRow.push('=""');
    lines.push(bandRow);
    labels.forEach(function (lab, i) {
      var src = f1 + i, out = [lab];
      for (var k = 0; k < KAIGI_SLOTS_; k++) {
        var a = colL(3 + k * 2), b = colL(4 + k * 2);
        out.push('=IF(N(' + a + '$' + tot + ')=0,"",' + a + src + '/' + a + '$' + tot + ')', '=IF(N(' + b + '$' + tot + ')=0,"",' + b + src + '/' + b + '$' + tot + ')');
      }
      lines.push(out);
    });
    var height = lines.length;
    momRows.push({ r: tot, kind: 'pct', dir: 0 });
    for (var m1 = 0; m1 < n; m1++) { momRows.push({ r: f1 + m1, kind: 'pct', dir: 0 }); momRows.push({ r: f2 + m1, kind: 'pt', dir: 0 }); }
    extra.forEach(function (ex, j) { momRows.push({ r: f1 + n + j, kind: ex[2] === '0%' ? 'pt' : 'pct', dir: 0 }); });
    sh.getRange(tot, 2, height, LAST_COL - 1).setFormulas(lines);
    sh.getRange(tot, 2, height, 1).setValues(lines.map(function (l) { return [l[0]]; }));
    var cntBlock = sh.getRange(tot, 3, 1 + n, LAST_COL - 2).setNumberFormat('#,##0');
    sh.getRange(f2, 3, n, LAST_COL - 2).setNumberFormat('0%');
    extra.forEach(function (ex, j) { sh.getRange(f1 + n + j, 3, 1, LAST_COL - 2).setNumberFormat(ex[2]); });
    sh.getRange(tot, 3, height, LAST_COL - 2).setHorizontalAlignment('right');
    for (var k6 = 0; k6 < KAIGI_SLOTS_; k6++) {
      sh.getRange(tot, 3 + k6 * 2, height, 1).setFontWeight('bold');
      sh.getRange(tot, 4 + k6 * 2, height, 1).setFontColor(MUTE);
    }
    sh.getRange(tot, 2, height, LAST_COL - 1).setBorder(null, null, true, null, null, true, LINE, SOLID);
    sh.getRange(tot, 2, 1, LAST_COL - 1).setBackground(SOFT).setFontWeight('bold').setBorder(null, null, true, null, null, null, INK, SOLID);
    if (extra.length) sh.getRange(f1 + n, 2, extra.length, LAST_COL - 1).setFontColor(MUTE).setFontSize(9);
    sh.getRange(lab0, 2, 1, LAST_COL - 1).setBackground(SOFT).setFontWeight('bold').setFontSize(9).setFontColor(MUTE)
      .setBorder(true, null, true, null, null, null, INK, SOLID);
    slotBorders(tot, height);
    zeroRanges.push(sh.getRange(f1, 3, n, LAST_COL - 2), sh.getRange(f2, 3, n, LAST_COL - 2));
    if (opt.redUp) {
      for (var k7 = 0; k7 < KAIGI_SLOTS_; k7++) {
        var a7 = colL(3 + k7 * 2), b7 = colL(4 + k7 * 2);
        rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + a7 + f2 + '),ISNUMBER(' + b7 + f2 + '),' + a7 + f2 + '-' + b7 + f2 + '>=' + opt.redUp + ')')
          .setFontColor(RED).setBackground('#FDF1F1').setRanges([sh.getRange(a7 + f2 + ':' + a7 + (f2 + n - 1))]).build());
      }
    }
    if (opt.sep) opt.sep.forEach(function (i) {
      sh.getRange(f1 + i - 1, 2, 1, LAST_COL - 1).setBorder(null, null, true, null, null, null, '#BDBDBD', SOLID);
      sh.getRange(f2 + i - 1, 2, 1, LAST_COL - 1).setBorder(null, null, true, null, null, null, '#BDBDBD', SOLID);
    });
    row = f2 + n + 2;
  };

  var ages = [['10代', 1, 19], ['20代', 20, 29], ['30代', 30, 39], ['40代', 40, 49], ['50代', 50, 59], ['60代', 60, 69], ['70代以上', 70, 200]];
  breakdown('入会した人の年代（月ごと）', '累計入会データから', ages.map(function (a) { return a[0]; }), function (m, lab, i) {
    return 'COUNTIFS(' + J + '$V$2:$V$8000,' + m + ',' + J + '$X$2:$X$8000,">=' + ages[i][1] + '",' + J + '$X$2:$X$8000,"<=' + ages[i][2] + '")';
  });
  breakdown('入会した人の男女（月ごと）', '累計入会データから', ['男', '女'], function (m, lab) {
    return 'COUNTIFS(' + J + '$V$2:$V$8000,' + m + ',' + J + '$I$2:$I$8000,' + lab + ')';
  });
  var ten = [['3ヶ月以内', 0, 3], ['4〜6ヶ月', 4, 6], ['7〜12ヶ月', 7, 12], ['1〜2年', 13, 24], ['2〜3年', 25, 36], ['3年超', 37, 9999]];
  breakdown('辞めた人の在籍期間（月ごと）', '辞めた月＝最終在籍月。累計入会データから', ten.map(function (a) { return a[0]; }), function (m, lab, i) {
    return 'COUNTIFS(' + J + '$W$2:$W$8000,EDATE(' + m + ',1),' + J + '$Z$2:$Z$8000,">=' + ten[i][1] + '",' + J + '$Z$2:$Z$8000,"<=' + ten[i][2] + '")';
  });
  breakdown('退会理由（月ごと・退会アンケート）', '回答した人だけ。「その他」は数えていません。赤＝比較より2pt以上増えた理由', LEAVE_REASONS_, function (m, lab) {
    return 'COUNTIFS(' + S + '$I$2:$I,' + m + ',' + S + '$C$2:$C,' + lab + ')';
  }, {
    totalLabel: '合計（回答した人）', redUp: 0.02, sep: [14],
    extra: [
      ['退会した人（会員動向）', function (h, base) { return (base === 'AJ' ? 'AL' : 'AP') + h; }, '#,##0'],
      ['回答率', function (h, base, col, tot) { return 'IFERROR(' + col + tot + '/' + (base === 'AJ' ? 'AL' : 'AP') + h + ',"")'; }, '0%']
    ]
  });
  breakdown('辞めた人の利用頻度（月ごと・退会アンケート）', '回答した人だけ', LEAVE_FREQS_, function (m, lab) {
    return 'COUNTIFS(' + S + '$I$2:$I,' + m + ',' + S + '$F$2:$F,' + lab + ')';
  }, { totalLabel: '合計（回答した人）' });

  // 「月ごとの増減」：右の列を前の月からの増減に。右端（AA）は期間の最初の月→最後の月
  var odd = [];
  for (var o = 0; o < KAIGI_SLOTS_; o++) odd.push(1 + o * 2);
  var chg = function (kind, a, p) {
    if (kind === 'pt') return 'TEXT((' + a + '-' + p + ')*100,"+0.0;-0.0;0.0")&"pt"';
    if (kind === 'diff') return 'TEXT(' + a + '-' + p + ',"+#,##0;-#,##0;0")';
    return 'IF(' + p + '=0,"",TEXT(' + a + '/' + p + '-1,"+0%;-0%;0%"))';
  };
  momRows.forEach(function (m) {
    var rg = sh.getRange(m.r, 3, 1, LAST_COL - 2);
    var fs = rg.getFormulas()[0];
    for (var k = 1; k < KAIGI_SLOTS_; k++) {
      var a = colL(3 + k * 2) + m.r, p = colL(1 + k * 2) + m.r;
      var ex = 'IF(OR(NOT(ISNUMBER(' + a + ')),NOT(ISNUMBER(' + p + '))),"",' + chg(m.kind, a, p) + ')';
      fs[k * 2 + 1] = '=IF(' + MOM + ',' + ex + ',' + String(fs[k * 2 + 1]).replace(/^=/, '') + ')';
      if (m.dir) rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(' + MOM.replace(/([A-Z]+)(\d+)/, '$$$1$$$2') + ',ISNUMBER(' + a + '),ISNUMBER(' + p + '),' + a + (m.dir === 1 ? '<' : '>') + p + ')')
        .setFontColor(RED).setBold(true).setRanges([sh.getRange(colL(4 + k * 2) + m.r)]).build());
    }
    fs[1] = '=IF(' + MOM + ',"",' + String(fs[1]).replace(/^=/, '') + ')';
    rg.setFormulas([fs]);
    var rowR = 'C' + m.r + ':' + colL(LAST_COL) + m.r;
    sh.getRange(m.r, LAST_COL + 1).setBackground(SOFT).setHorizontalAlignment('right').setFontWeight('bold').setFontFamily('Meiryo')
      .setBorder(null, null, true, null, null, null, LINE, SOLID).setFormula('=IFERROR(LET(v,CHOOSECOLS(' + rowR + ',' + odd.join(',') + '),n,FILTER(v,ISNUMBER(v)),c,COLUMNS(n),' +
      'IF(c<2,"",' + chg(m.kind, 'INDEX(n,1,c)', 'INDEX(n,1,1)') + ')),"")');
    if (m.dir) rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH(' + colL(LAST_COL + 1) + m.r + '&"","^' + (m.dir === 1 ? '-' : '\\+') + '")')
      .setFontColor(RED).setBold(true).setRanges([sh.getRange(m.r, LAST_COL + 1)]).build());
  });
  sh.setColumnWidth(LAST_COL + 1, 76);

  sh.getRange(12, 2, row - 12, LAST_COL - 1).setFontFamily('Meiryo').setVerticalAlignment('middle');
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setFontColor('#C8C8C8').setRanges(zeroRanges).build());
  sh.setConditionalFormatRules(rules);
  sh.hideColumns(34, sh.getMaxColumns() - 33);
  kaigiApplyPeriod_(sh);
  props.setProperty('KAIGI_BYMON', 'v4');
  return { ok: true, rows: row };
}

/** 会議用に「開始月」プルダウン（K4:L4）を足す。期間の選び方に「開始月〜対象月」「直近2ヶ月」を追加 */
function addKaigiFromMonth_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('KAIGI_V') !== KAIGI_V_ || props.getProperty('KAIGI_FROM') === 'v1') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(KAIGI_SHEET_);
  if (!sh) return { ok: false };
  var C = KAIGI_CELLS_;
  sh.getRange(C.mode).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(KAIGI_MODES_, true).setAllowInvalid(false).build());
  sh.getRange('E3').setValue('対象月（期間の最後の月）');
  sh.getRange('K3').setValue('開始月（「開始月〜対象月」のとき）').setFontSize(9).setFontColor('#7A7A7A');
  var rg = sh.getRange('K4:L4');
  rg.breakApart(); rg.merge();
  var cur = String(sh.getRange(C.month).getDisplayValue());
  var cell = sh.getRange(C.from);
  cell.setNumberFormat('@').setValue(cur);
  cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange('AH1:AH200'), true).setAllowInvalid(false).build());
  rg.setBackground('#FFF7E6').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center').setFontFamily('Meiryo')
    .setBorder(true, true, true, true, null, null, '#111111', SpreadsheetApp.BorderStyle.SOLID);
  var rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B$4<>"開始月〜対象月"')
    .setFontColor('#BDBDBD').setBackground('#F3F3F3').setRanges([rg]).build());
  sh.setConditionalFormatRules(rules);
  sh.getRange('H1').setFormula('=IF(' + C.mode + '="開始月〜対象月",' + C.from + '&"〜"&' + C.month + ',' + C.month + '&IF(' + C.mode + '="月を選ぶ","",IF(' + C.mode + '="手入力","（手入力）","　"&' + C.mode + ')))&"　｜　"&IF(' + C.cmp + '="比較なし","比較なし",' + C.cmp + '&"と比較")');
  protectInputsOnly_(sh, [C.mode, C.month, C.cmp, C.cur, C.prev, C.from]);
  props.setProperty('KAIGI_FROM', 'v1');
  return { ok: true };
}

/** 月の 月初会員・入会・退会（本部の会員動向を優先、今月の入会・退会は日報、無い月は累計データ）。A は月初日のセル */
function jlMonthF_(A) {
  var jV = "'累計入会データ'!$V$2:$V$8000", lU = "'累計退会データ'!$U$2:$U$8000";
  var thisMonth = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  var isNow = 'TEXT(' + A + ',"yymm")=TEXT(\'日報\'!$B$1,"0")';
  var calcB = 'COUNTIFS(' + jV + ',">0",' + jV + ',"<"&' + A + ')-COUNTIFS(' + lU + ',"<="&' + A + ')';
  return {
    start: 'LET(h,' + jlHqExpr_(A, 3, 4, '月初') + ',IF(h>0,h,' + calcB + '))',
    join: 'IF(' + isNow + ',N(\'日報\'!$C$13),LET(h,' + jlHqExpr_(A, 6, 7, '入') + ',IF(AND(h>0,' + A + '<' + thisMonth + '),h,COUNTIF(' + jV + ',' + A + '))))',
    leave: 'IF(' + isNow + ',N(\'日報\'!$C$15),LET(h,' + jlHqExpr_(A, 9, 10, '解') + ',IF(AND(h>0,' + A + '<' + thisMonth + '),h,COUNTIF(' + lU + ',EDATE(' + A + ',1)))))'
  };
}

/**
 * 会議用の右側に「月ごとの比較」（O〜AA）。期間の各月と、比較期間の同じ順番の月を並べて比率を出す。
 * 月の数字は AJ:AQ（非表示）で計算。
 */
function addKaigiMonthly_(ss) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('KAIGI_V') !== KAIGI_V_ || props.getProperty('KAIGI_MON') === 'v1') return { ok: true, skipped: true };
  var sh = ss.getSheetByName(KAIGI_SHEET_);
  if (!sh) return { ok: false };
  var INK = '#111111', MUTE = '#7A7A7A', LINE = '#E3E3E3', SOFT = '#F7F7F7', RED = '#B91C1C';
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;
  if (sh.getMaxColumns() < 44) sh.insertColumnsAfter(sh.getMaxColumns(), 44 - sh.getMaxColumns());
  sh.getRange('O13:AQ40').clear();
  sh.showColumns(15, 30);

  var hf = [];
  for (var i = 0; i < 12; i++) {
    var r = 2 + i;
    var cur = 'AJ' + r, cmp = 'AN' + r;
    var fc = jlMonthF_(cur), fp = jlMonthF_(cmp);
    var blank = function (ref, f) { return '=IF(' + ref + '="","",' + f + ')'; };
    hf.push([
      '=IF(OR($B$6="",$C$6=""),"",IF(EDATE(DATE(YEAR($B$6),MONTH($B$6),1),' + i + ')>$C$6,"",EDATE(DATE(YEAR($B$6),MONTH($B$6),1),' + i + ')))',
      blank(cur, fc.join), blank(cur, fc.leave), blank(cur, fc.start),
      '=IF(OR(' + cur + '="",$H$6="",' + KAIGI_CELLS_.cmp + '="比較なし"),"",EDATE(DATE(YEAR($H$6),MONTH($H$6),1),' + i + '))',
      blank(cmp, fp.join), blank(cmp, fp.leave), blank(cmp, fp.start)
    ]);
  }
  sh.getRange('AJ1:AQ1').setValues([['今回の月', '入会', '退会', '月初', '比較の月', '入会', '退会', '月初']]);
  sh.getRange(2, 36, 12, 8).setFormulas(hf);
  sh.getRange('AJ2:AJ13').setNumberFormat('yyyy/m');
  sh.getRange('AN2:AN13').setNumberFormat('yyyy/m');

  sh.getRange('O13').setValue('月ごとの比較').setFontSize(12).setFontWeight('bold').setFontFamily('Meiryo');
  sh.getRange('R13').setFormula('=IF(' + KAIGI_CELLS_.cmp + '="比較なし","",' + KAIGI_CELLS_.cmp + '&"の同じ順番の月と比べています（比率＝今回÷比較）")')
    .setFontSize(9).setFontColor(MUTE);
  var head = ['月', '比較の月', '入会', '比較', '比率', '退会', '比較', '比率', '純増', '比較', '退会率', '比較', '差'];
  sh.getRange(14, 15, 1, 13).setValues([head]).setBackground(INK).setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(9).setHorizontalAlignment('center');
  [17, 20, 23, 25].forEach(function (c) { sh.getRange(14, c).setBorder(null, true, null, null, null, null, '#FFFFFF', SOLID); });

  var rows = [];
  for (var k = 0; k < 12; k++) {
    var h = 2 + k, rr = 15 + k;
    var on = function (f) { return '=IF($AJ' + h + '="","",' + f + ')'; };
    var oc = function (f) { return '=IF($AN' + h + '="","",' + f + ')'; };
    rows.push([
      on('$AJ' + h), oc('$AN' + h),
      on('AK' + h), oc('AO' + h), '=IFERROR(IF(OR(Q' + rr + '="",R' + rr + '=""),"",Q' + rr + '/R' + rr + '),"")',
      on('AL' + h), oc('AP' + h), '=IFERROR(IF(OR(T' + rr + '="",U' + rr + '=""),"",T' + rr + '/U' + rr + '),"")',
      on('AK' + h + '-AL' + h), oc('AO' + h + '-AP' + h),
      on('IFERROR(AL' + h + '/AM' + h + ',"")'), oc('IFERROR(AP' + h + '/AQ' + h + ',"")'),
      '=IF(OR(Y' + rr + '="",Z' + rr + '=""),"",Y' + rr + '-Z' + rr + ')'
    ]);
  }
  sh.getRange(15, 15, 12, 13).setFormulas(rows);
  var sumIf = function (col) { return '=IF(COUNT(' + col + '15:' + col + '26)=0,"",SUM(' + col + '15:' + col + '26))'; };
  sh.getRange(27, 15, 1, 13).setFormulas([[
    '="合計"', '=""',
    sumIf('Q'), sumIf('R'), '=IFERROR(IF(OR(Q27="",R27=""),"",Q27/R27),"")',
    sumIf('T'), sumIf('U'), '=IFERROR(IF(OR(T27="",U27=""),"",T27/U27),"")',
    sumIf('W'), sumIf('X'),
    '=IFERROR(SUM(AL2:AL13)/SUM(AM2:AM13),"")', '=IF(R27="","",IFERROR(SUM(AP2:AP13)/SUM(AQ2:AQ13),""))',
    '=IF(OR(Y27="",Z27=""),"",Y27-Z27)'
  ]]);
  var body = sh.getRange(15, 15, 13, 13);
  body.setFontFamily('Meiryo').setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('right');
  sh.getRange('O15:P27').setHorizontalAlignment('center').setNumberFormat('yyyy/m');
  sh.getRange('P15:P27').setFontColor(MUTE);
  ['Q', 'T'].forEach(function (c) { sh.getRange(c + '15:' + c + '27').setNumberFormat('#,##0').setFontWeight('bold'); });
  ['R', 'U', 'X'].forEach(function (c) { sh.getRange(c + '15:' + c + '27').setNumberFormat('#,##0').setFontColor(MUTE); });
  ['S', 'V'].forEach(function (c) { sh.getRange(c + '15:' + c + '27').setNumberFormat('0%').setFontWeight('bold').setBackground(SOFT); });
  sh.getRange('W15:W27').setNumberFormat('+#,##0;-#,##0;0').setFontWeight('bold');
  sh.getRange('X15:X27').setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange('Y15:Y27').setNumberFormat('0.0%').setFontWeight('bold');
  sh.getRange('Z15:Z27').setNumberFormat('0.0%').setFontColor(MUTE);
  sh.getRange('AA15:AA27').setNumberFormat('+0.0%;-0.0%;0.0%').setBackground(SOFT);
  sh.getRange(15, 15, 12, 13).setBorder(null, null, true, null, null, true, LINE, SOLID);
  sh.getRange(27, 15, 1, 13).setFontWeight('bold').setBackground(SOFT).setBorder(true, null, true, null, null, null, INK, SOLID);
  [17, 20, 23, 25].forEach(function (c) { sh.getRange(15, c, 13, 1).setBorder(null, true, null, null, null, null, '#BDBDBD', SOLID); });
  sh.getRange('O27').setHorizontalAlignment('center');
  for (var c2 = 15; c2 <= 27; c2++) sh.setColumnWidth(c2, c2 <= 16 ? 76 : 64);
  sh.setColumnWidth(28, 18);

  var rules = sh.getConditionalFormatRules().filter(function (ru) { return ru.getRanges()[0].getColumn() < 15; });
  var cf = function () { return SpreadsheetApp.newConditionalFormatRule(); };
  rules.push(cf().whenNumberLessThan(1).setFontColor(RED).setRanges([sh.getRange('S15:S27')]).build());
  rules.push(cf().whenNumberGreaterThan(1).setFontColor(RED).setRanges([sh.getRange('V15:V27')]).build());
  rules.push(cf().whenFormulaSatisfied('=AND(ISNUMBER(W15),ISNUMBER(X15),W15<X15)').setFontColor(RED).setRanges([sh.getRange('W15:W27')]).build());
  rules.push(cf().whenNumberGreaterThan(0).setFontColor(RED).setRanges([sh.getRange('AA15:AA27')]).build());
  sh.setConditionalFormatRules(rules);

  sh.hideColumns(34, sh.getMaxColumns() - 33);
  var editors = [];
  try { editors = ss.getEditors().map(function (u) { return u.getEmail(); }); } catch (e) { editors = ['?' + e.message.slice(0, 40)]; }
  props.setProperty('KAIGI_MON', 'v1');
  return { ok: true, editors: editors };
}

function kaigiApplyPeriod_(sh) {
  var mode = String(sh.getRange(KAIGI_CELLS_.mode).getValue());
  if (mode === '手入力') return;
  var m = String(sh.getRange(KAIGI_CELLS_.month).getDisplayValue()).match(/^(\d{4})年(\d{1,2})月$/);
  if (!m) return;
  var f = String(sh.getRange(KAIGI_CELLS_.from).getDisplayValue()).match(/^(\d{4})年(\d{1,2})月$/);
  var from = f ? new Date(Number(f[1]), Number(f[2]) - 1, 1) : null;
  var p = kaigiPeriod_(mode, Number(m[1]), Number(m[2]), String(sh.getRange(KAIGI_CELLS_.cmp).getValue()), from);
  sh.getRange(KAIGI_CELLS_.cur).setValues([[p.start, p.end]]);
  sh.getRange(KAIGI_CELLS_.prev).setValues([[p.cs, p.ce]]);
}

function kaigiPeriod_(mode, y, mo, cmp, from) {
  var start = new Date(y, mo - 1, 1);
  if (mode === '開始月〜対象月' && from) {
    if (from > start) { var t = from; from = start; start = t; y = start.getFullYear(); mo = start.getMonth() + 1; }
    start = from;
  }
  if (mode === '直近2ヶ月') start = new Date(y, mo - 2, 1);
  if (mode === '直近3ヶ月') start = new Date(y, mo - 3, 1);
  if (mode === '直近6ヶ月') start = new Date(y, mo - 6, 1);
  if (mode.indexOf('今年度') === 0) start = new Date(mo >= 4 ? y : y - 1, 3, 1);
  var end = new Date(y, mo, 0);
  var cs = '', ce = '';
  if (cmp === '前年同期') {
    cs = new Date(start.getFullYear() - 1, start.getMonth(), 1);
    ce = new Date(end.getFullYear() - 1, end.getMonth() + 1, 0);
  } else if (cmp === '前の期間') {
    var n = (end.getFullYear() - start.getFullYear()) * 12 + end.getMonth() - start.getMonth() + 1;
    cs = new Date(start.getFullYear(), start.getMonth() - n, 1);
    ce = new Date(start.getFullYear(), start.getMonth(), 0);
  }
  return { start: start, end: end, cs: cs, ce: ce };
}

function kaigiOnEdit_(e) {
  var sh = e.range.getSheet();
  var a = e.range.getA1Notation().split(':')[0];
  if (a === KAIGI_CELLS_.from) {
    if (String(sh.getRange(KAIGI_CELLS_.mode).getValue()) !== '開始月〜対象月') sh.getRange(KAIGI_CELLS_.mode).setValue('開始月〜対象月');
    kaigiApplyPeriod_(sh);
    return;
  }
  if (a === KAIGI_CELLS_.mode || a === KAIGI_CELLS_.month || a === KAIGI_CELLS_.cmp) {
    kaigiApplyPeriod_(sh);
    return;
  }
  var r = e.range.getRow(), c = e.range.getColumn();
  if (r === 6 && (c === 2 || c === 3 || c === 8 || c === 9) && String(sh.getRange(KAIGI_CELLS_.mode).getValue()) !== '手入力') {
    sh.getRange(KAIGI_CELLS_.mode).setValue('手入力');
  }
}

/** トップ A2 に日報の自動更新の状態を出す。新しい更新があれば IMPORTRANGE を読み直させる */
function syncTopRefreshStatus_(ss) {
  var sh = ss.getSheetByName(HUB_HOME_SHEET_);
  if (!sh) return;
  var h = callReceptionRefreshApi_('health');
  var parse = function (s) {
    try { return s ? Utilities.parseDate(String(s), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') : null; } catch (e) { return null; }
  };
  var fmt = function (d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'M/d HH:mm'); };
  var checked = parse(h && h.checked), last = parse(h && h.lastUpdate);
  var status, bad = true;
  if (!h || !h.ok) status = '⚠ 自動更新の状態を確認できませんでした';
  else if (!checked || Date.now() - checked.getTime() > 10 * 60 * 1000) status = '⚠ 数字の自動更新が止まっています' + (checked ? '（最終チェック ' + fmt(checked) + '）' : '');
  else if (h.lastError) status = '⚠ 前回の自動更新でエラー：' + String(h.lastError).slice(0, 40);
  else {
    status = '● 最終更新　日報 ' + (last ? fmt(last) : '—');
    bad = false;
  }
  var gessho = '—';
  try {
    var idx = ss.getSheetByName('月初３ファイル');
    if (idx && idx.getLastRow() > 1) {
      var ts = idx.getRange(2, 2, idx.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }).filter(function (d) { return d instanceof Date; });
      if (ts.length) gessho = Utilities.formatDate(new Date(Math.max.apply(null, ts)), 'Asia/Tokyo', 'M/d');
    }
  } catch (eG) {}
  var survey = String(PropertiesService.getDocumentProperties().getProperty('LEAVE_SURVEY_AT') || '').replace(/^\d{4}\//, '').replace(/^0/, '').replace(/\/0/, '/').replace(/ .*/, '');
  var cum = String(PropertiesService.getDocumentProperties().getProperty('CUMULATIVE_AT') || '').replace(/^\d{4}\//, '').replace(/^0/, '').replace(/\/0/, '/').replace(/ .*/, '');
  var guide = '　｜　月初ファイル ' + gessho + '　｜　入会・退会一覧 ' + (cum || '—') + '　｜　退会アンケート ' + (survey || '—');
  var cell = sh.getRange('A2');
  if (String(cell.getDisplayValue()) !== status + guide) {
    var rich = SpreadsheetApp.newRichTextValue().setText(status + guide)
      .setTextStyle(0, status.length, SpreadsheetApp.newTextStyle().setForegroundColor(bad ? '#B91C1C' : '#111111').setBold(true).setFontSize(10).build())
      .setTextStyle(status.length, status.length + guide.length, SpreadsheetApp.newTextStyle().setForegroundColor('#7A7A7A').setBold(false).setFontSize(9).build())
      .build();
    cell.setRichTextValue(rich);
  }
  var props = PropertiesService.getDocumentProperties();
  if (h && h.lastUpdate && props.getProperty('NUMBERS_LAST_SEEN') !== h.lastUpdate) {
    try { reloadReceptionImports_(ss); } catch (e1) {}
    try { reloadNippoMirror_(ss); } catch (e2) {}
    props.setProperty('NUMBERS_LAST_SEEN', h.lastUpdate);
  }
  return { status: status };
}

/** 日報の写し（IMPORTRANGE）を式の書き換えで読み直させる */
function reloadNippoMirror_(ss) {
  var sh = ss.getSheetByName('日報');
  if (!sh) return;
  var cell = sh.getRange('A1');
  var f = cell.getFormula();
  if (!/IMPORTRANGE/i.test(f)) return;
  f = f.indexOf('"&"",') >= 0 ? f.replace('"&"",', '",') : f.replace(/IMPORTRANGE\("([^"]+)",/i, 'IMPORTRANGE("$1"&"",');
  cell.setFormula(f);
}

function clearHqOptionNotes_() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('HQ_OPT_NOTES_CLEAR') === 'v1') return { ok: true, skipped: true };
  var sh = SpreadsheetApp.openById(HQ_TREND_ID_).getSheetByName('経堂');
  if (!sh) return { ok: false };
  var rg = sh.getRange(31, 3, 16, 7);
  var notes = rg.getNotes().map(function (row) {
    return row.map(function (n) {
      return String(n || '').split('\n').filter(function (line) { return !/OP表で修正/.test(line); }).join('\n');
    });
  });
  rg.setNotes(notes);
  var left = 0;
  rg.getNotes().forEach(function (row) { row.forEach(function (n) { if (/OP表で修正/.test(n)) left++; }); });
  try { openWorkspaceSpreadsheet_().getSheetByName('WorkspaceSync').getRange('H3').setValue('HQメモ残り ' + left + ' ' + new Date()); } catch (eL) {}
  props.setProperty('HQ_OPT_NOTES_CLEAR', 'v1');
  return { ok: true };
}

/** 日報 10月の月初：落合 悠野（10/1 規約退会・男）を外す 1516→1515 / 男 1075→1074。1回だけ */
function fixNippoOctKiyaku_() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('NIPPO_OCT_KIYAKU_FIX') === 'v1') return { ok: true, skipped: true };
  var nippo = SpreadsheetApp.openById(RECEPTION_SOURCE_ID_).getSheetByName('日報');
  if (!nippo || String(nippo.getRange('B1').getDisplayValue()) !== '2610') return { ok: false };
  if (Number(nippo.getRange('C12').getValue()) === 1516 && Number(nippo.getRange('F12').getValue()) === 1075) {
    nippo.getRange('C12').setValue(1515).setNote(
      '月初会員 = 当月末在籍 1535 − 法人都度6 − OGF3 − ゴールド3 − 当月開始7 − 規約退会1（落合 悠野 10/1）= 1515\n' +
      '会員動向：9月初 1471 ＋ 入会94 − 退会50 = 1515'
    );
    nippo.getRange('F12').setValue(1074);
  }
  props.setProperty('NIPPO_OCT_KIYAKU_FIX', 'v1');
  return { ok: true };
}


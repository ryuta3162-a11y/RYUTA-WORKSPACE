/**
 * JOYFIT24経堂 — 入会・退会管理（サイドバー）
 *
 * 入会・退会の表示更新入口。メニューには出さない（アンケート以外の管理UIは非表示）。
 *
 * 残メニュー: アンケート
 */

// ── 共通スタイル ──

/** JOYFIT調 … 赤(#C21632)・白・黒・グレーのみ（立体感あり） */
function joyfitManagementCss_() {
  return [
    ":root { --red:#C21632; --red-dark:#9A1128; --ink:#111111; --muted:#5f6368; --line:#dadce0; --bg:#eceff1; --white:#ffffff; }",
    "* { box-sizing: border-box; }",
    "body { font-family: Meiryo, 'Segoe UI', sans-serif; margin:0; color:var(--ink); background:",
    "  radial-gradient(ellipse at top, #fafafa 0%, var(--bg) 55%); min-height:100%; }",
    ".hero { position:relative; padding:24px 18px 20px; color:#fff;",
    "  background: linear-gradient(160deg, #1a1a1a 0%, #0a0a0a 55%, #1c0a0e 100%);",
    "  box-shadow: 0 10px 28px rgba(0,0,0,.35), inset 0 -1px 0 rgba(255,255,255,.06); overflow:hidden; }",
    ".hero:before { content:''; position:absolute; left:0; right:0; top:0; height:3px;",
    "  background:linear-gradient(90deg, var(--red-dark), var(--red), #ff4d6a); }",
    ".hero:after { content:''; position:absolute; right:-30px; top:-40px; width:120px; height:120px;",
    "  border-radius:50%; background:radial-gradient(circle, rgba(194,22,50,.35), transparent 70%); }",
    ".brand { position:relative; color:var(--red); font-size:10px; font-weight:800; letter-spacing:2.4px; text-shadow:0 1px 0 rgba(0,0,0,.4); }",
    "h1 { position:relative; font-size:21px; margin:8px 0 12px; letter-spacing:.4px; font-weight:800;",
    "  text-shadow:0 2px 8px rgba(0,0,0,.45); }",
    ".state { position:relative; display:flex; align-items:center; gap:8px; color:#e8eaed; font-size:12px; line-height:1.5; }",
    ".dot { width:8px; height:8px; border-radius:50%; background:var(--red);",
    "  box-shadow:0 0 0 4px rgba(194,22,50,.22), 0 0 12px rgba(194,22,50,.55); flex:none; }",
    "main { padding:14px 12px 16px; }",
    ".card { background:linear-gradient(180deg, #ffffff 0%, #fbfbfb 100%); border:1px solid rgba(0,0,0,.06);",
    "  border-radius:16px; padding:14px; margin-bottom:12px;",
    "  box-shadow: 0 1px 0 rgba(255,255,255,.9) inset, 0 8px 22px rgba(0,0,0,.08), 0 2px 4px rgba(0,0,0,.04); }",
    ".sec { font-size:11px; font-weight:800; color:var(--ink); margin:0 0 10px; letter-spacing:.4px;",
    "  padding-bottom:8px; border-bottom:1px solid #eee; }",
    ".grid { display:grid; grid-template-columns:repeat(4,1fr); gap:7px; }",
    ".grid3 { display:grid; grid-template-columns:repeat(3,1fr); gap:7px; }",
    ".metric { min-width:0; text-align:center; padding:10px 2px; border-radius:12px;",
    "  background:linear-gradient(180deg, #f7f8f9 0%, #eef0f2 100%);",
    "  box-shadow: inset 0 1px 0 #fff, 0 1px 2px rgba(0,0,0,.05); border:1px solid rgba(0,0,0,.04); }",
    ".metric b { display:block; font-size:18px; letter-spacing:-.5px; color:var(--ink); font-weight:800; }",
    ".metric span { color:var(--muted); font-size:9px; white-space:nowrap; }",
    ".meta { color:var(--muted); font-size:10px; margin:10px 2px 0; line-height:1.5; }",
    ".hint { color:var(--muted); font-size:10px; line-height:1.55; margin:8px 2px 0; padding:8px 10px;",
    "  background:#f4f5f6; border-radius:8px; border-left:3px solid var(--red); }",
    "button.action { width:100%; border:0; border-radius:12px; padding:12px 10px; font-family:inherit;",
    "  font-size:12px; font-weight:800; cursor:pointer; transition:.18s ease; margin-top:8px; letter-spacing:.2px; }",
    "button.action:first-of-type { margin-top:0; }",
    "button.action:hover { transform:translateY(-1px); }",
    "button.action:active { transform:translateY(1px); }",
    "button.action:disabled { opacity:.55; cursor:wait; transform:none; }",
    ".sync { color:#fff; background:linear-gradient(180deg, #d41c3a 0%, var(--red) 45%, var(--red-dark) 100%);",
    "  box-shadow: 0 1px 0 rgba(255,255,255,.25) inset, 0 8px 18px rgba(194,22,50,.32), 0 2px 0 #7a0d20; }",
    ".sync:hover { filter:brightness(1.05); }",
    ".test { color:var(--ink); background:linear-gradient(180deg, #ffffff, #e8eaed);",
    "  box-shadow: 0 1px 0 #fff inset, 0 4px 10px rgba(0,0,0,.08); border:1px solid #d0d3d6; }",
    ".send { color:#fff; background:linear-gradient(180deg, #2a2a2a, #111111);",
    "  box-shadow: 0 1px 0 rgba(255,255,255,.12) inset, 0 6px 14px rgba(0,0,0,.28); }",
    ".warn { color:var(--ink); background:linear-gradient(180deg, #fafafa, #eceff1);",
    "  border:1px solid #d0d3d6; box-shadow:0 1px 0 #fff inset, 0 3px 8px rgba(0,0,0,.06); }",
    ".progress { display:none; margin-top:8px; color:var(--red); font-size:11px; text-align:center; font-weight:800; }",
    "#error { display:none; margin:0 0 10px; padding:11px 12px; color:#fff;",
    "  background:linear-gradient(180deg, #d41c3a, var(--red)); border-radius:12px; font-size:11px; white-space:pre-wrap;",
    "  box-shadow:0 6px 16px rgba(194,22,50,.28); }",
    ".foot { color:var(--muted); font-size:10px; text-align:center; line-height:1.55; padding:4px 8px 10px; }",
    /* マニュアル */
    ".manual { margin-top:12px; border-radius:12px; overflow:hidden;",
    "  border:1px solid #e0e0e0; background:#fafafa;",
    "  box-shadow: inset 0 1px 0 #fff, 0 2px 6px rgba(0,0,0,.04); }",
    ".manual-toggle { width:100%; border:0; background:linear-gradient(180deg,#fff,#f3f4f5);",
    "  padding:11px 12px; display:flex; align-items:center; justify-content:space-between;",
    "  font-family:inherit; font-size:12px; font-weight:800; color:var(--ink); cursor:pointer; }",
    ".manual-toggle .label { display:flex; align-items:center; gap:8px; }",
    ".manual-toggle .badge { display:inline-block; font-size:9px; color:#fff; background:var(--red);",
    "  padding:2px 7px; border-radius:999px; letter-spacing:.3px; box-shadow:0 2px 6px rgba(194,22,50,.3); }",
    ".manual-toggle .chevron { color:var(--muted); font-size:10px; transition:transform .2s; }",
    ".manual-toggle.open .chevron { transform:rotate(180deg); }",
    ".manual-body { display:none; padding:0 12px 12px; border-top:1px solid #eee; background:#fff; }",
    ".manual-body.open { display:block; }",
    ".steps { list-style:none; margin:10px 0 0; padding:0; }",
    ".steps li { position:relative; padding:10px 0 10px 28px; font-size:11px; line-height:1.55; color:#333;",
    "  border-bottom:1px dashed #ececec; }",
    ".steps li:last-child { border-bottom:0; padding-bottom:2px; }",
    ".steps .n { position:absolute; left:0; top:10px; width:20px; height:20px; border-radius:50%;",
    "  background:linear-gradient(180deg, #2a2a2a, #111); color:#fff; font-size:10px; font-weight:800;",
    "  display:flex; align-items:center; justify-content:center;",
    "  box-shadow:0 2px 6px rgba(0,0,0,.2); }",
    ".steps .sub { display:block; margin-top:3px; color:var(--muted); font-size:10px; }",
    "a.sheet-link { color:var(--red); font-weight:800; text-decoration:none; border-bottom:1px solid rgba(194,22,50,.35);",
    "  cursor:pointer; }",
    "a.sheet-link:hover { color:var(--red-dark); border-bottom-color:var(--red-dark); }",
    "a.sheet-link:active { opacity:.7; }"
  ].join("\n");
}

function joyfitManagementCommonJs_() {
  return [
    "function showError(err) {",
    "  var box = document.getElementById('error');",
    "  box.textContent = 'エラー: ' + (err && err.message ? err.message : err);",
    "  box.style.display = 'block';",
    "}",
    "function clearError() { document.getElementById('error').style.display = 'none'; }",
    "function setBusy(busy) {",
    "  var buttons = document.querySelectorAll('button.action');",
    "  for (var i = 0; i < buttons.length; i++) buttons[i].disabled = !!busy;",
    "}",
    "function setProgress(text) {",
    "  var el = document.getElementById('progress');",
    "  if (!el) return;",
    "  el.textContent = text || '';",
    "  el.style.display = text ? 'block' : 'none';",
    "}",
    "function runAction(fnName, progressText, confirmText) {",
    "  if (confirmText && !confirm(confirmText)) return;",
    "  clearError(); setBusy(true); setProgress(progressText || '処理中...');",
    "  var runner = google.script.run",
    "    .withSuccessHandler(function(data){ setBusy(false); setProgress(''); if (data) render(data); })",
    "    .withFailureHandler(function(err){ setBusy(false); setProgress(''); showError(err); });",
    "  runner[fnName]();",
    "}",
    "function toggleManual(btn) {",
    "  var body = btn.parentNode.querySelector('.manual-body');",
    "  var open = !body.classList.contains('open');",
    "  body.classList.toggle('open', open);",
    "  btn.classList.toggle('open', open);",
    "}",
    "function openSheet(name) {",
    "  clearError();",
    "  google.script.run.withFailureHandler(showError).joyfitOpenSheetByName(name);",
    "}"
  ].join("\n");
}

/** マニュアル内リンクからシートを開く */
function joyfitOpenSheetByName(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = String(sheetName || "").trim();
  if (!name) throw new Error("シート名が空です。");

  if (name === "アンケート_取込" || name === "アンケート" || name === "アンケート_退会") {
    ensureAnketoSheets_(ss);
  }
  if (name === "入会・退会") {
    setupDisplaySheet_(ss);
  }
  if (name === "6ヶ月割管理" || name === "在籍期間") {
    if (typeof kyodoGetTenureSheet_ === "function") {
      var tenure = kyodoGetTenureSheet_(ss, false);
      if (!tenure) throw new Error("シート「6ヶ月割管理」が見つかりません。先に最新化を実行してください。");
      ss.setActiveSheet(tenure);
      return true;
    }
  }

  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error("シート「" + name + "」が見つかりません。");
  ss.setActiveSheet(sh);
  return true;
}

function joyfitManualHtml_(stepsHtml) {
  return [
    '<div class="manual">',
    '<button type="button" class="manual-toggle" onclick="toggleManual(this)">',
    '<span class="label"><span class="badge">GUIDE</span>マニュアルはこちら</span>',
    '<span class="chevron">▼</span>',
    "</button>",
    '<div class="manual-body">',
    '<ol class="steps">',
    stepsHtml,
    "</ol></div></div>"
  ].join("");
}

function joyfitManualStep_(num, html, sub) {
  return (
    "<li><span class=\"n\">" + num + "</span>" + html +
    (sub ? "<span class=\"sub\">" + sub + "</span>" : "") +
    "</li>"
  );
}

function joyfitSheetLink_(sheetName, label) {
  var text = label || sheetName;
  return (
    "<a class=\"sheet-link\" href=\"#\" onclick=\"openSheet('" + sheetName + "');return false;\">" +
    text + "</a>"
  );
}

/** マニュアルからボタン処理を直接呼ぶリンク */
function joyfitActionLink_(fnName, label, progressText) {
  var prog = progressText || "処理中...";
  return (
    "<a class=\"sheet-link\" href=\"#\" onclick=\"runAction('" + fnName + "','" + prog +
    "');return false;\">" + label + "</a>"
  );
}

// ── 入口 ──

/** JOYFITメニュー: 入会・退会管理サイドバーを開く */
function openMembershipMonthManagement() {
  var html = HtmlService.createHtmlOutput(buildMembershipMonthManagementHtml_())
    .setTitle("入会・退会管理")
    .setWidth(390);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** 旧メニュー名の互換（残っていても開ける） */
function openMembershipManagement() {
  openMembershipMonthManagement();
}

function openMonthManagement() {
  openMembershipMonthManagement();
}

/** JOYFITメニュー: 年月を選んで入退会だけ更新 */
function openMembershipMonthUpdateDialog() {
  var now = new Date();
  var options = [];
  for (var delta = -12; delta <= 1; delta++) {
    var d = new Date(now.getFullYear(), now.getMonth() + delta, 1);
    var value = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
    var label = d.getFullYear() + "年" + (d.getMonth() + 1) + "月";
    var selected = delta === 0 ? " selected" : "";
    options.push("<option value=\"" + value + "\"" + selected + ">" + label + "</option>");
  }

  var html = HtmlService.createHtmlOutput([
    "<!DOCTYPE html><html><head><base target=\"_top\"><style>",
    "body{font-family:Arial,'Noto Sans JP',sans-serif;padding:20px;color:#222}",
    "h2{font-size:18px;margin:0 0 8px}.note{font-size:12px;color:#666;line-height:1.6;margin-bottom:14px}",
    "select,button{width:100%;box-sizing:border-box;font-size:15px;padding:10px;margin-top:8px}",
    "button{background:#c21632;color:white;border:0;border-radius:6px;font-weight:bold;cursor:pointer}",
    "button:disabled{opacity:.5}.status{white-space:pre-wrap;font-size:12px;margin-top:12px;line-height:1.6}",
    "</style></head><body>",
    "<h2>入退会更新（年月選択）</h2>",
    "<div class=\"note\">対象シートがない場合は「原本」から自動作成します。<br>E=一般入会 / F=法人入会 / 退会E=一般 / F71=法人前受</div>",
    "<select id=\"ym\">", options.join(""), "</select>",
    "<button id=\"run\" onclick=\"runUpdate()\">この年月を更新</button>",
    "<div id=\"status\" class=\"status\"></div>",
    "<script>",
    "function runUpdate(){",
    " var value=document.getElementById('ym').value.split('-');",
    " var button=document.getElementById('run'); var status=document.getElementById('status');",
    " if(!confirm(value[0]+'年'+Number(value[1])+'月の入退会を更新しますか？'))return;",
    " button.disabled=true; status.textContent='Gmailを確認して更新中です…';",
    " google.script.run.withSuccessHandler(function(result){",
    "   status.textContent=result.message; button.disabled=false;",
    " }).withFailureHandler(function(err){",
    "   status.textContent='エラー: '+(err&&err.message?err.message:err); button.disabled=false;",
    " }).runMembershipMonthUpdate(Number(value[0]),Number(value[1]));",
    "}",
    "</script></body></html>"
  ].join(""))
    .setWidth(420)
    .setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, "入退会更新");
}

function runMembershipMonthUpdate(year, monthNumber) {
  year = Number(year);
  monthNumber = Number(monthNumber);
  if (!Number.isInteger(year) || !Number.isInteger(monthNumber) || monthNumber < 1 || monthNumber > 12) {
    throw new Error("年月の指定が正しくありません。");
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error("別の更新処理が実行中です。少し待ってから再実行してください。");
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var month = monthNumber - 1;
    var prepared = ensureMembershipMonthlySheet_(ss, year, month);
    var result = rebuildMembershipDailyCountsForMonth_(ss, year, month);
    var next = shiftYearMonth_(year, month, 1);
    var dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
    var withdrawRows = readWithdrawData_(dataSheet);
    var nextCorporate = countCorporateWithdrawAdvance_(withdrawRows, next.year, next.month);
    var nextPrepared = null;
    if (nextCorporate > 0) {
      nextPrepared = ensureMembershipMonthlySheet_(ss, next.year, next.month);
      syncMembershipDailyCountsForMonth_(ss, next.year, next.month);
    }
    ss.setActiveSheet(prepared.sheet);
    var audit = result.audit;
    var currentCorporate = countCorporateWithdrawAdvance_(withdrawRows, year, month);
    return {
      message:
        prepared.sheetName + (prepared.created ? " を原本から作成し、更新しました。" : " を更新しました。") +
        "\n入会 " + audit.enrollTotal + "（一般 " + audit.enrollGeneral + " / 法人 " + audit.enrollCorporate + "）" +
        "\n一般退会 " + audit.withdrawGeneral + " / 当月法人退会(F71) " + currentCorporate +
        (nextCorporate > 0
          ? "\n翌月法人退会 " + nextCorporate + " → " + nextPrepared.sheetName + "!F71" +
            (nextPrepared.created ? "（原本から作成）" : "")
          : "") +
        "\nシートを開いて数値をご確認ください。"
    };
  } finally {
    lock.releaseLock();
  }
}

function openOperatingMonthSwitchDialog() {
  throw new Error("運用月切替はメニューから外しています。日報B1とOP集計B1は手で合わせてください。");
}

function switchOperatingMonth(year, monthNumber) {
  year = Number(year);
  monthNumber = Number(monthNumber);
  if (!Number.isInteger(year) || !Number.isInteger(monthNumber) || monthNumber < 1 || monthNumber > 12) {
    throw new Error("年月の指定が正しくありません。");
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var month = monthNumber - 1;
  var sheetName = formatMonthlySheetName_(year, month);
  if (!ss.getSheetByName(sheetName)) {
    throw new Error("シート「" + sheetName + "」がありません。先に「入退会更新（年月選択）」で作成してください。");
  }

  var opSheet = getOpSheet_(ss);
  if (!opSheet) throw new Error("「OP集計」シートが見つかりません。");
  var oldLabel = String(opSheet.getRange("B1").getDisplayValue() || opSheet.getRange("B1").getValue() || "").trim();
  archiveManualOpeningForLabel_(ss, opSheet, oldLabel);

  var newLabel = year + "年" + monthNumber + "月";
  opSheet.getRange("B1").setNumberFormat("@").setValue(newLabel);
  setNippoMonthSheetName_(ss, sheetName);
  syncSummarySheetFromOpData_(ss, year, month, opSheet);
  loadManualOpeningForMonth_(ss, opSheet, year, month);
  ss.setActiveSheet(ss.getSheetByName(sheetName));

  return {
    message: "運用月を" + newLabel + "へ切り替えました。\n日報B1: " + sheetName + "\nOP集計B1: " + newLabel
  };
}

// ── サマリー（開いた瞬間用・シートを何度も読み直さない） ──

function getMembershipMonthManagementSummary() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var member = getMembershipManagementSummary_();
    var month = getMonthManagementSummary_();
    var anketo = getAnketoManagementSummary_(ss);

    return {
      message: member.message,
      enrollTotal: member.enrollTotal,
      withdrawTotal: member.withdrawTotal,
      enrollView: member.enrollView,
      withdrawView: member.withdrawView,
      enrollLabel: member.enrollLabel,
      withdrawLabel: member.withdrawLabel,
      cancelTotal: member.cancelTotal,
      sheetName: month.sheetName,
      monthlyExists: month.monthlyExists,
      nippoB1: month.nippoB1,
      memberOpening: month.memberOpening,
      monthLabel: month.monthLabel,
      archiveRows: month.archiveRows,
      anketoRawRows: anketo.rawRows,
      anketoReady: anketo.ready,
      anketoMonth: anketo.monthLabel
    };
  } catch (err) {
    Logger.log("getMembershipMonthManagementSummary: " + err.message);
    return {
      message: "読み込みエラー: " + err.message,
      enrollTotal: 0,
      withdrawTotal: 0,
      enrollView: 0,
      withdrawView: 0,
      enrollLabel: "—",
      withdrawLabel: "—",
      cancelTotal: 0,
      sheetName: "—",
      monthlyExists: false,
      nippoB1: "—",
      memberOpening: "—",
      monthLabel: "—",
      archiveRows: 0,
      anketoRawRows: 0,
      anketoReady: false,
      anketoMonth: "—"
    };
  }
}

/**
 * ダッシュボード用の軽い集計
 * ※ setupDisplaySheet_ や isWithdrawalCancelled_ の都度シート読込はしない
 *   （開くだけで「確認中」のまま固まる原因だった）
 */
function getMembershipManagementSummary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  var displaySheet = ss.getSheetByName(SHEET_NAME_MEMBERS);

  // 表示中の件数は、画面の入会・退会シートをそのまま数える（スプレッドシートと一致）
  var enrollLabel = "—";
  var withdrawLabel = "—";
  var enrollView = 0;
  var withdrawView = 0;
  if (displaySheet) {
    enrollLabel = normalizeYearMonthLabel_(
      displaySheet.getRange("B1").getDisplayValue() || displaySheet.getRange("B1").getValue()
    ) || "—";
    withdrawLabel = normalizeYearMonthLabel_(
      displaySheet.getRange("G1").getDisplayValue() || displaySheet.getRange("G1").getValue()
    ) || "—";
    enrollView = countNonEmptyInColumn_(displaySheet, 2, DISPLAY_START_ROW);
    withdrawView = countNonEmptyInColumn_(displaySheet, 7, DISPLAY_START_ROW);
  }

  // 合計はデータシートの氏名列だけ数える（全列読みしない）
  var enrollTotal = 0;
  var withdrawTotal = 0;
  if (dataSheet && dataSheet.getLastRow() >= DATA_START_ROW) {
    enrollTotal = countNonEmptyInColumn_(dataSheet, DATA_ENROLL_COL + 1, DATA_START_ROW);
    withdrawTotal = countNonEmptyInColumn_(dataSheet, DATA_WITHDRAW_COL + 1, DATA_START_ROW);
  }

  // 退会キャンセル件数はマップを1回だけ読む
  var cancelTotal = 0;
  if (typeof kyodoLoadWithdrawCancelMap_ === "function") {
    try {
      var map = kyodoLoadWithdrawCancelMap_(ss);
      cancelTotal = Object.keys(map.byEmail || {}).length;
    } catch (e) {
      cancelTotal = 0;
    }
  }

  return {
    enrollTotal: enrollTotal,
    withdrawTotal: withdrawTotal,
    cancelTotal: cancelTotal,
    enrollView: enrollView,
    withdrawView: withdrawView,
    enrollLabel: enrollLabel,
    withdrawLabel: withdrawLabel,
    message: enrollTotal || withdrawTotal || enrollView || withdrawView
      ? "シートの内容を表示しています。"
      : "データがありません。メニュー「更新」または「入退会更新（年月選択）」を実行してください。"
  };
}

/** 指定列の非空白セル数（ヘッダー行は除く） */
function countNonEmptyInColumn_(sheet, col, startRow) {
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return 0;
  var numRows = lastRow - startRow + 1;
  var values = sheet.getRange(startRow, col, numRows, 1).getValues();
  var n = 0;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) n++;
  }
  return n;
}

function getMonthManagementSummary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth();
  var sheetName = formatMonthlySheetName_(year, month);
  var monthlyExists = !!ss.getSheetByName(sheetName);

  var nippo = ss.getSheetByName("日報");
  var nippoB1 = nippo
    ? String(nippo.getRange("B1").getDisplayValue() || nippo.getRange("B1").getValue() || "").trim()
    : "";
  var memberOpening = nippo ? nippo.getRange("C12").getDisplayValue() || nippo.getRange("C12").getValue() : "";

  var archive = ss.getSheetByName(SHEET_NAME_OPENING_ARCHIVE);
  var archiveRows = archive && archive.getLastRow() > 1 ? archive.getLastRow() - 1 : 0;

  return {
    sheetName: sheetName,
    monthlyExists: monthlyExists,
    nippoB1: nippoB1 || "—",
    memberOpening: memberOpening === "" || memberOpening === null ? "—" : String(memberOpening),
    archiveRows: archiveRows,
    monthLabel: year + "年" + (month + 1) + "月"
  };
}

function getAnketoManagementSummary_(ss) {
  var raw = ss.getSheetByName(SHEET_NAME_ANKETO_RAW);
  var rawRows = raw && raw.getLastRow() > 1 ? raw.getLastRow() - 1 : 0;
  var sheet = ss.getSheetByName(SHEET_NAME_ANKETO);
  var monthLabel = "—";
  if (sheet) {
    monthLabel = String(sheet.getRange("B1").getDisplayValue() || sheet.getRange("B1").getValue() || "").trim() || "—";
  }
  return {
    ready: rawRows > 0,
    rawRows: rawRows,
    monthLabel: monthLabel
  };
}

// ── 入会・退会アクション ──

/**
 * 入会・退会の表示更新（Gmailは触らない）
 * 新着メールはメニュー「更新」または19:30自動
 */
function membershipManagementUpdate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDataSheet_(ss);
  setupDisplaySheet_(ss);
  var dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  var enrollCount = readEnrollData_(dataSheet).length;
  var withdrawCount = readWithdrawData_(dataSheet).length;
  if (enrollCount === 0 && withdrawCount === 0) {
    throw new Error(
      "データがまだありません。\nメニュー「更新」または「入退会更新（年月選択）」を先に実行してください。"
    );
  }
  refreshMembershipDisplay_(ss);
  try { syncMembershipDailyCountsSilent_(); } catch (err) { Logger.log(err); }
  var summary = getMembershipMonthManagementSummary();
  summary.message = "表示を更新しました。新着メールはメニュー「更新」または毎日19:30の自動取込です。";
  return summary;
}

/**
 * 任意: 直近3日のメールだけ軽く取り込む（Gmailあり・重い場合あり）
 */
function membershipManagementFetchMail_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) {
    throw new Error("別の処理が実行中です。1分ほど待ってから再度お試しください。");
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    setupDataSheet_(ss);
    var daysBack = typeof MEMBERSHIP_UPDATE_DAYS_BACK !== "undefined"
      ? MEMBERSHIP_UPDATE_DAYS_BACK : 3;
    var enrollAdded = fetchMembershipEmailsIncrementalKind_(ss, "enroll", daysBack);
    var withdrawAdded = fetchMembershipEmailsIncrementalKind_(ss, "withdraw", daysBack);
    refreshMembershipDisplay_(ss);
    try { syncMembershipDailyCountsSilent_(); } catch (err) { Logger.log(err); }
    var summary = getMembershipMonthManagementSummary();
    summary.message = "メール取込完了 … 入会 +" + enrollAdded +
      " / 退会 +" + withdrawAdded + "（直近" + daysBack + "日）";
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function membershipManagementRefetchAll() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error("別の更新処理が実行中です。少し待ってから再実行してください。");
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    setupDataSheet_(ss);
    migrateLegacyLayouts_(ss);
    setupDisplaySheet_(ss);
    installMembershipLabelsSilent_();

    var full = fetchAllMembershipEmailsCore_(ss, { skipLabels: true });
    repairMonthValuesInData_(ss);
    dedupeStoredMemberData_(ss);
    refreshMembershipDisplay_(ss);
    try { syncMembershipDailyCountsSilent_(); } catch (err) { Logger.log(err); }

    var summary = getMembershipMonthManagementSummary();
    summary.message = "全件再取得完了 … 入会 " + full.enrollTotal + " / 退会 " + full.withdrawTotal;
    return summary;
  } finally {
    lock.releaseLock();
  }
}

/** 入会・退会シートを前面に出す（一覧の確認用・色もJOYFIT調に整える） */
function membershipManagementOpenSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    refreshMembershipDisplay_(ss);
  } catch (e) {
    setupDisplaySheet_(ss);
  }
  var sh = ss.getSheetByName(SHEET_NAME_MEMBERS);
  if (!sh) throw new Error("「入会・退会」シートが見つかりません。");
  ss.setActiveSheet(sh);
  var summary = getMembershipMonthManagementSummary();
  summary.message = "入会・退会シートを開きました。B1=入会月 / G1=退会月。";
  return summary;
}

// ── 月初アクション ──

function monthManagementSetup() {
  throw new Error("月初更新作業はメニューから外しています。月初は手動で行ってください。");
}

function monthManagementRestore() {
  throw new Error("月初復元はメニューから外しています。月初は手動で行ってください。");
}

function monthManagementImportAll() {
  throw new Error("月初一括取込はメニューから外しています。月初は手動で行ってください。");
}

function monthManagementOpenNippo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("日報");
  if (!sh) throw new Error("「日報」シートが見つかりません。");
  ss.setActiveSheet(sh);
  return getMembershipMonthManagementSummary();
}

// ── アンケート管理（独立） ──

function openAnketoManagement() {
  var html = HtmlService.createHtmlOutput(buildAnketoManagementHtml_())
    .setTitle("アンケート管理")
    .setWidth(390);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAnketoManagementDashboardSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ensureAnketoSheets_(ss);
    applyAnketoJoyfitLayout_(ss);
  } catch (e) {
    Logger.log(e);
  }
  var anketo = getAnketoManagementSummary_(ss);
  return {
    anketoRawRows: anketo.rawRows,
    anketoReady: anketo.ready,
    anketoMonth: anketo.monthLabel,
    message: anketo.ready
      ? "取込済 " + anketo.rawRows + " 行 / 表示月 " + anketo.monthLabel
      : "取込シートにCSVを貼り、「更新」を押してください。"
  };
}

/**
 * 更新1回: レイアウト整え → 取込シート反映 → 集計更新
 */
function anketoManagementUpdate() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error("別の更新処理が実行中です。少し待ってから再実行してください。");
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAnketoSheets_(ss);
    applyAnketoJoyfitLayout_(ss);

    var parts = [];
    try {
      var pack = importAnketoCsvFromPasteSheetCore_(ss);
      importAnketoCsvFinishImport_(pack);
      parts.push("取込反映OK");
    } catch (e) {
      parts.push("取込シートに新規なし");
    }

    try {
      var result = syncAnketoDisplay_(ss);
      parts.push("集計 " + result.surveyCount + "人 / 突合 " + result.linkedCount);
    } catch (e2) {
      var summaryErr = getAnketoManagementDashboardSummary();
      summaryErr.message = parts.join(" … ") + " / " + e2.message;
      return summaryErr;
    }

    applyAnketoJoyfitLayout_(ss);
    var summary = getAnketoManagementDashboardSummary();
    summary.message = "更新完了 … " + parts.join(" … ");
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function anketoManagementOpenPasteSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAnketoSheets_(ss);
  applyAnketoJoyfitLayout_(ss);
  var sh = ss.getSheetByName(SHEET_NAME_ANKETO_PASTE);
  if (!sh) throw new Error("「アンケート_取込」シートを作成できませんでした。");
  ss.setActiveSheet(sh);
  var summary = getAnketoManagementDashboardSummary();
  summary.message = "取込シートを開きました。CSVを貼って「更新」を押してください。";
  return summary;
}

function anketoManagementOpenSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAnketoSheets_(ss);
  applyAnketoJoyfitLayout_(ss);
  var sh = ss.getSheetByName(SHEET_NAME_ANKETO);
  if (!sh) throw new Error("「アンケート」シートが見つかりません。");
  ss.setActiveSheet(sh);
  var summary = getAnketoManagementDashboardSummary();
  summary.message = "アンケートシートを開きました。B1で退会月を選べます。";
  return summary;
}

function buildAnketoManagementHtml_() {
  var manual = joyfitManualHtml_([
    joyfitManualStep_(1,
      joyfitActionLink_("anketoManagementOpenPasteSheet", "取込シート", "") +
      " にCSVを貼る"),
    joyfitManualStep_(2,
      joyfitActionLink_("anketoManagementUpdate", "更新", "更新中...") +
      " … レイアウト＋取込＋集計まで一括"),
    joyfitManualStep_(3,
      joyfitActionLink_("anketoManagementOpenSheet", "アンケート", "") +
      " で結果を確認（B1=退会月）")
  ].join(""));

  return [
    "<!DOCTYPE html><html><head><base target=\"_top\"><style>",
    joyfitManagementCss_(),
    "</style></head><body>",
    "<header class=\"hero\">",
    "<div class=\"brand\">JOYFIT24 KYODO</div>",
    "<h1>アンケート管理</h1>",
    "<div class=\"state\"><span class=\"dot\"></span><span id=\"message\">確認中...</span></div>",
    "</header>",
    "<main>",
    "<div id=\"error\"></div>",
    "<div id=\"progress\" class=\"progress\"></div>",

    "<section class=\"card\">",
    "<div class=\"sec\">■ マニュアル</div>",
    manual,
    "</section>",

    "<section class=\"card\">",
    "<div class=\"sec\">■ 状態</div>",
    "<div class=\"grid3\">",
    "<div class=\"metric\"><b id=\"anketoRaw\">-</b><span>取込行</span></div>",
    "<div class=\"metric\"><b id=\"anketoMonth\">-</b><span>表示月</span></div>",
    "<div class=\"metric\"><b id=\"anketoReady\">-</b><span>状態</span></div>",
    "</div>",
    "<div style=\"height:10px\"></div>",
    "<button class=\"action sync\" onclick=\"runAction('anketoManagementUpdate','更新中...')\">↻ 更新</button>",
    "<button class=\"action test\" onclick=\"runAction('anketoManagementOpenPasteSheet','')\">▤ 取込シート</button>",
    "<button class=\"action send\" onclick=\"runAction('anketoManagementOpenSheet','')\">▤ アンケート</button>",
    "</section>",

    "<div class=\"foot\">更新1回で色分け・取込・集計まで完了</div>",
    "</main>",
    "<script>",
    joyfitManagementCommonJs_(),
    "function render(data) {",
    "  document.getElementById('message').textContent = data.message;",
    "  document.getElementById('anketoRaw').textContent = data.anketoRawRows;",
    "  document.getElementById('anketoMonth').textContent = data.anketoMonth;",
    "  document.getElementById('anketoReady').textContent = data.anketoReady ? '取込済' : '未取込';",
    "}",
    "function refresh() {",
    "  google.script.run.withSuccessHandler(render).withFailureHandler(showError).getAnketoManagementDashboardSummary();",
    "}",
    "refresh();",
    "</script></body></html>"
  ].join("");
}

// ── HTML ──

function buildMembershipMonthManagementHtml_() {
  var manual = joyfitManualHtml_([
    joyfitManualStep_(1,
      joyfitActionLink_("membershipManagementUpdate", "表示を更新", "表示を更新中...") +
      " … 既存データから数字をシートに反映"),
    joyfitManualStep_(2,
      joyfitActionLink_("membershipManagementOpenSheet", "入会・退会シート", "") +
      " … 一覧（B1=入会月 / G1=退会月）"),
    joyfitManualStep_(3,
      "漏れがあるとき … JOYFIT「入退会更新（年月選択）」")
  ].join(""));

  return [
    "<!DOCTYPE html><html><head><base target=\"_top\"><style>",
    joyfitManagementCss_(),
    "</style></head><body>",
    "<header class=\"hero\">",
    "<div class=\"brand\">JOYFIT24 KYODO</div>",
    "<h1>入会・退会管理</h1>",
    "<div class=\"state\"><span class=\"dot\"></span><span id=\"message\">確認中...</span></div>",
    "</header>",
    "<main>",
    "<div id=\"error\"></div>",
    "<div id=\"progress\" class=\"progress\"></div>",

    "<section class=\"card\">",
    "<div class=\"sec\">■ マニュアル</div>",
    manual,
    "</section>",

    "<section class=\"card\">",
    "<div class=\"sec\">■ 入会・退会</div>",
    "<div class=\"grid\">",
    "<div class=\"metric\"><b id=\"enrollTotal\">-</b><span>入会合計</span></div>",
    "<div class=\"metric\"><b id=\"withdrawTotal\">-</b><span>退会合計</span></div>",
    "<div class=\"metric\"><b id=\"enrollView\">-</b><span>表示入会</span></div>",
    "<div class=\"metric\"><b id=\"withdrawView\">-</b><span>表示退会</span></div>",
    "</div>",
    "<div id=\"memberMeta\" class=\"meta\"></div>",
    "<div style=\"height:10px\"></div>",
    "<button class=\"action sync\" onclick=\"runAction('membershipManagementUpdate','表示を更新中...')\">↻ 表示を更新</button>",
    "<button class=\"action send\" onclick=\"runAction('membershipManagementOpenSheet','')\">▤ 入会・退会シート</button>",
    "</section>",

    "<div class=\"foot\">新着メールは毎日19:30に自動取込<br>月初作業は手動（メニューからは省いています）</div>",
    "</main>",
    "<script>",
    joyfitManagementCommonJs_(),
    "function render(data) {",
    "  document.getElementById('message').textContent = data.message;",
    "  document.getElementById('enrollTotal').textContent = data.enrollTotal;",
    "  document.getElementById('withdrawTotal').textContent = data.withdrawTotal;",
    "  document.getElementById('enrollView').textContent = data.enrollView;",
    "  document.getElementById('withdrawView').textContent = data.withdrawView;",
    "  document.getElementById('memberMeta').textContent =",
    "    '表示月 入会 ' + data.enrollLabel + ' / 退会 ' + data.withdrawLabel +",
    "    (data.cancelTotal ? '　キャンセル ' + data.cancelTotal + '件' : '');",
    "}",
    "function refresh() {",
    "  clearError();",
    "  document.getElementById('message').textContent = '確認中...';",
    "  var finished = false;",
    "  var timer = setTimeout(function(){",
    "    if (finished) return;",
    "    finished = true;",
    "    showError({ message: '読み込みがタイムアウトしました。管理画面.gs が最新か確認し、サイドバーを開き直してください。' });",
    "    document.getElementById('message').textContent = '読み込み失敗';",
    "  }, 15000);",
    "  google.script.run",
    "    .withSuccessHandler(function(data){",
    "      if (finished) return;",
    "      finished = true;",
    "      clearTimeout(timer);",
    "      render(data);",
    "    })",
    "    .withFailureHandler(function(err){",
    "      if (finished) return;",
    "      finished = true;",
    "      clearTimeout(timer);",
    "      showError(err);",
    "      document.getElementById('message').textContent = '読み込み失敗';",
    "    })",
    "    .getMembershipMonthManagementSummary();",
    "}",
    "refresh();",
    "</script></body></html>"
  ].join("");
}

// ── オプション管理 ──

/** JOYFITメニュー: オプション管理サイドバー */
function openOptionManagement() {
  var html = HtmlService.createHtmlOutput(buildOptionManagementHtml_())
    .setTitle("オプション管理")
    .setWidth(390);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getOptionManagementSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nippo = ss.getSheetByName("日報");
  var opSheet = typeof getOpSheet_ === "function" ? getOpSheet_(ss) : ss.getSheetByName(SHEET_NAME_OP);

  var ym = typeof resolveNippoTargetYearMonth_ === "function"
    ? resolveNippoTargetYearMonth_(ss)
    : { year: new Date().getFullYear(), month: new Date().getMonth() };

  var monthLabel = ym.year + "年" + (ym.month + 1) + "月";
  var nippoB1 = nippo
    ? String(nippo.getRange("B1").getDisplayValue() || nippo.getRange("B1").getValue() || "").trim()
    : "—";
  var memberOpening = nippo
    ? (nippo.getRange("C12").getDisplayValue() || nippo.getRange("C12").getValue() || "—")
    : "—";

  var startTotal = 0;
  var stopTotal = 0;
  if (nippo && typeof OPTION_LIST !== "undefined") {
    var n = OPTION_LIST.length;
    var cVals = nippo.getRange(NIPPO_OP_START_ROW, NIPPO_COL_CURRENT, n, 1).getValues();
    var fVals = nippo.getRange(NIPPO_OP_START_ROW, NIPPO_COL_STOP, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      startTotal += Number(cVals[i][0]) || 0;
      stopTotal += Number(fVals[i][0]) || 0;
    }
  }

  var opRows = 0;
  var summaryB1 = "—";
  if (opSheet) {
    summaryB1 = String(opSheet.getRange("B1").getDisplayValue() || opSheet.getRange("B1").getValue() || "").trim() || "—";
    if (typeof readOpLogValues_ === "function") {
      opRows = Math.max(0, readOpLogValues_(opSheet).length - 1);
    }
  }

  return {
    monthLabel: monthLabel,
    nippoB1: nippoB1 || "—",
    memberOpening: String(memberOpening),
    startTotal: startTotal,
    stopTotal: stopTotal,
    netTotal: startTotal - stopTotal,
    optionCount: typeof OPTION_LIST !== "undefined" ? OPTION_LIST.length : 0,
    opRows: opRows,
    summaryB1: summaryB1,
    message: "利用開始 " + startTotal + " / 解約 " + stopTotal + "（純増減 " + (startTotal - stopTotal) + "）",
    guide: "数値の最新化はメニュー「更新」。ここは確認・シートを開く用。"
  };
}

/** 互換: メニュー「更新」と同じ処理へ寄せる */
function optionManagementUpdate() {
  var result = runSimpleDailyUpdateCore_(true);
  var summary = getOptionManagementSummary();
  summary.message = "更新完了 … " + (result && result.label ? result.label + " / " : "") +
    "利用開始 " + summary.startTotal + " / 解約 " + summary.stopTotal;
  return summary;
}

/** @deprecated 月初は手動運用。エディタからのみ */
function optionManagementRestoreOpening() {
  throw new Error("月初反映はメニューから外しています。月初は手入力で管理してください。");
}

function optionManagementOpenSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSpreadsheet();
  var sh = getOpSheet_(ss);
  ss.setActiveSheet(sh);
  var summary = getOptionManagementSummary();
  summary.message = "OP集計を開きました。左=集計(A〜H) / 右=契約明細(I列〜・新しい順)。";
  return summary;
}

function optionManagementOpenNippo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("日報");
  if (!sh) throw new Error("「日報」シートが見つかりません。");
  ss.setActiveSheet(sh);
  return getOptionManagementSummary();
}

function buildOptionManagementHtml_() {
  var manual = joyfitManualHtml_([
    joyfitManualStep_(1,
      "数値の最新化はメニュー「更新」（オプション＋入会退会）"),
    joyfitManualStep_(2,
      joyfitActionLink_("optionManagementOpenSheet", "OP集計", "") +
      " … 左が集計、右(I列〜・新しい順)が契約明細"),
    joyfitManualStep_(3,
      "月初(D列)は手動で入力・管理します"),
    joyfitManualStep_(4,
      joyfitActionLink_("installOptionMailLabels", "OP：追加・停止ラベル", "ラベル作成中...") +
      " … 入会時OPは入会メール側")
  ].join(""));

  return [
    "<!DOCTYPE html><html><head><base target=\"_top\"><style>",
    joyfitManagementCss_(),
    "</style></head><body>",
    "<header class=\"hero\">",
    "<div class=\"brand\">JOYFIT24 KYODO</div>",
    "<h1>オプション管理</h1>",
    "<div class=\"state\"><span class=\"dot\"></span><span id=\"message\">確認中...</span></div>",
    "</header>",
    "<main>",
    "<div id=\"error\"></div>",
    "<div id=\"progress\" class=\"progress\"></div>",

    "<section class=\"card\">",
    "<div class=\"sec\">■ マニュアル</div>",
    manual,
    "</section>",

    "<section class=\"card\">",
    "<div class=\"sec\">■ 当月</div>",
    "<div class=\"grid3\">",
    "<div class=\"metric\"><b id=\"startTotal\">-</b><span>利用開始</span></div>",
    "<div class=\"metric\"><b id=\"stopTotal\">-</b><span>解約</span></div>",
    "<div class=\"metric\"><b id=\"netTotal\">-</b><span>純増減</span></div>",
    "</div>",
    "<div id=\"meta\" class=\"meta\"></div>",
    "<div style=\"height:10px\"></div>",
    "<button class=\"action send\" onclick=\"runAction('optionManagementOpenSheet','')\">▤ OP集計</button>",
    "<button class=\"action test\" onclick=\"runAction('optionManagementOpenNippo','')\">▤ 日報</button>",
    "<button class=\"action\" onclick=\"runAction('installOptionMailLabels','ラベル作成中...')\">🏷 OP追加・停止ラベル</button>",
    "</section>",

    "<div class=\"foot\">メニュー「更新」= 直近1週間ラベル＋既読＋数値更新</div>",
    "</main>",
    "<script>",
    joyfitManagementCommonJs_(),
    "function render(data) {",
    "  document.getElementById('message').textContent = data.message;",
    "  document.getElementById('startTotal').textContent = data.startTotal;",
    "  document.getElementById('stopTotal').textContent = data.stopTotal;",
    "  document.getElementById('netTotal').textContent = data.netTotal;",
    "  document.getElementById('meta').textContent =",
    "    data.monthLabel + '　OP集計B1 ' + data.summaryB1 + '　明細 ' + data.opRows + '行';",
    "}",
    "function refresh() {",
    "  google.script.run.withSuccessHandler(render).withFailureHandler(showError).getOptionManagementSummary();",
    "}",
    "refresh();",
    "</script></body></html>"
  ].join("");
}

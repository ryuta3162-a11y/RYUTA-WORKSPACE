/**
 * JOYFIT24経堂 — 6ヶ月割管理画面
 *
 * 既存の取込・判定・送信処理は変更せず、スタッフ向けの入口だけを提供する。
 * このファイルを削除しても既存メニューの機能には影響しない。
 * 実処理は 退会キャンセル.gs 側（取込・在籍期間・送信）。
 */

var KYODO_MANAGEMENT_TITLE = "6カ月継続管理";

/** JOYFITメニューから開くスタッフ向けサイドバー */
function kyodoOpenCampaignManagement() {
  var html = HtmlService.createHtmlOutput(kyodoBuildCampaignManagementHtml_())
    .setTitle(KYODO_MANAGEMENT_TITLE)
    .setWidth(390);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** 管理画面に表示する件数だけを返す（氏名・メールアドレスは返さない） */
function kyodoGetCampaignManagementSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tenureSh = typeof kyodoGetTenureSheet_ === "function"
    ? kyodoGetTenureSheet_(ss, false)
    : ss.getSheetByName(KYODO_TENURE_SHEET_NAME);
  var memberSh = ss.getSheetByName("入会者");
  var result = {
    ready: false,
    total: 0,
    selected: 0,
    pending: 0,
    sent: 0,
    failed: 0,
    memberNumberMissing: 0,
    cancelled: 0,
    memberMasterRows: memberSh ? Math.max(memberSh.getLastRow() - 1, 0) : 0,
    memberMasterAvailable: !!memberSh,
    message: "「6カ月継続管理」シートはまだ作成されていません。「更新」を押してください。",
    guide: ""
  };

  if (!tenureSh || tenureSh.getLastRow() < 2) return result;

  var data = tenureSh.getDataRange().getValues();
  var header = data[0];
  var memberNoCol = kyodoFindHeaderColIndex_(header, ["会員番号"]);
  var nameCol = kyodoFindHeaderColIndex_(header, ["名前", "氏名", "会員氏名"]);
  var emailCol = kyodoFindHeaderColIndex_(header, ["メールアドレス", "メール"]);
  var targetCol = kyodoFindHeaderColIndex_(header, [KYODO_TENURE_SEND_TARGET_HEADER]);
  var sentCol = kyodoFindHeaderColIndex_(header, [KYODO_CAMPAIGN_SEND_LOG_HEADER, "送信済"]);
  var cancelCol = kyodoFindHeaderColIndex_(header, [KYODO_TENURE_CANCEL_HEADER]);

  for (var r = 1; r < data.length; r++) {
    var hasPerson = (nameCol >= 0 && data[r][nameCol]) || (emailCol >= 0 && data[r][emailCol]);
    if (!hasPerson) continue;

    result.total++;
    if (memberNoCol < 0 || !String(data[r][memberNoCol] || "").trim()) {
      result.memberNumberMissing++;
    }

    var cancelled = cancelCol >= 0 && (
      data[r][cancelCol] === true || String(data[r][cancelCol]).toUpperCase() === "TRUE"
    );
    if (cancelled) {
      result.cancelled = (result.cancelled || 0) + 1;
      continue;
    }

    var selected = targetCol >= 0 && data[r][targetCol] === true;
    var sendLog = sentCol >= 0 ? String(data[r][sentCol] || "").trim() : "";
    var failed = sendLog.indexOf("失敗") === 0;
    var sent = !!sendLog && !failed;

    if (selected) result.selected++;
    if (sent) result.sent++;
    if (failed) result.failed++;
    if (selected && !sent) result.pending++;
  }

  result.ready = true;
  result.cancelled = result.cancelled || 0;
  if (result.pending > 0) {
    result.message = "チェック済みで、まだメールしていない人が " + result.pending + " 名います。";
  } else if (result.total > 0) {
    result.message = "リストは " + result.total + " 名です（キャンセル " + result.cancelled + " 名除くと有効 " +
      (result.total - result.cancelled) + " 名）。";
  } else {
    result.message = "リストが空です。「更新」を押してください。";
  }
  result.guide =
    "リスト＝全員 / 未送信＝送信対象チェック済（メール待ち） / 表示切替は見え方だけ変更 / 非表示でもチェックONなら送信対象";
  return result;
}

/** 既存の取込処理をそのまま呼ぶ */
function kyodoManagementImportEnrollment() {
  kyodoImportEnrollmentSmart_();
  return kyodoGetCampaignManagementSummary();
}

/** 既存の在籍期間判定をそのまま呼ぶ */
function kyodoManagementPrepareTargets() {
  kyodoPrepareCampaignTargets_();
  return kyodoGetCampaignManagementSummary();
}

/** 管理画面用: ダイアログなしで最新の入会メールを反映する */
function kyodoManagementSyncEnrollment() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error("別の更新処理が実行中です。少し待ってから再実行してください。");
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = kyodoGetOutputSheet_(ss);
    var isEmpty = !sh || sh.getLastRow() < 2;
    var now = new Date();
    var rangeStart = isEmpty ? KYODO_ENROLL_RANGE_START : new Date(now.getFullYear(), now.getMonth(), 1);
    var result = kyodoImportEnrollmentEmailsCore_(rangeStart, kyodoEndOfToday_());

    var matchSh = ss.getSheetByName(KYODO_MATCH_SHEET_NAME);
    if (matchSh && matchSh.getLastRow() >= 2) {
      try { kyodoHookupEmailsCore_(); } catch (hookErr) { Logger.log(hookErr); }
    }
    return { imported: result.count, messages: result.messageCount };
  } finally {
    lock.releaseLock();
  }
}

/** 管理画面用: リストを更新するだけ（シートは開かない） */
function kyodoManagementPrepareAndOpen() {
  return kyodoManagementRefreshList_();
}

/** 「更新」ボタン … Gmailを見て6ヶ月割管理を作り直す */
function kyodoManagementRefreshList_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error("別の更新処理が実行中です。少し待ってから再実行してください。");
  try {
    kyodoCheckWithdrawalTenureRecentCore_();
    var summary = kyodoGetCampaignManagementSummary();
    summary.message = "更新しました。リスト " + summary.total + " 名 / 未送信 " + summary.pending + " 名";
    return summary;
  } finally {
    lock.releaseLock();
  }
}

/** サイドバーから呼ぶ公開入口（末尾 _ の関数は google.script.run から呼べない） */
function kyodoManagementRefreshList() {
  return kyodoManagementRefreshList_();
}

/** 6ヶ月割管理シートへ移動 */
function kyodoManagementOpenTenureSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = kyodoGetTenureSheet_(ss, false);
  if (!sh) throw new Error("先に『最新化・一覧を開く』を実行してください。");
  ss.setActiveSheet(sh);
  return true;
}

/** サイドバー用: 6カ月継続管理シートの表示順・絞り込みを切り替える */
function kyodoManagementApplyTenureView(mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = kyodoGetTenureSheet_(ss, false);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error("先に「更新」を押して、6カ月継続管理のリストを作成してください。");
  }

  var viewMode = String(mode || "tenure_desc");
  kyodoManagementShowAllTenureRows_(sh);
  kyodoManagementSortTenureSheet_(sh, viewMode);

  var monthMatch = viewMode.match(/^month_([1-5])$/);
  if (monthMatch) {
    kyodoManagementHideTenureRowsExceptMonth_(sh, Number(monthMatch[1]));
  }

  ss.setActiveSheet(sh);
  var summary = kyodoGetCampaignManagementSummary();
  summary.message = kyodoManagementViewMessage_(viewMode);
  return summary;
}

function kyodoManagementShowAllTenureRows_(sh) {
  var maxRows = sh.getMaxRows();
  if (maxRows > 1) sh.showRows(2, maxRows - 1);
}

function kyodoManagementSortTenureSheet_(sh, mode) {
  var lastRow = sh.getLastRow();
  if (lastRow < 3) return;

  var sortSpecs;
  if (mode === "enroll_desc") {
    sortSpecs = [
      { column: 9, ascending: true },
      { column: 4, ascending: false }
    ];
  } else if (mode === "withdraw_desc") {
    sortSpecs = [
      { column: 9, ascending: true },
      { column: 5, ascending: false }
    ];
  } else {
    sortSpecs = [
      { column: 9, ascending: true },
      { column: 6, ascending: false },
      { column: 4, ascending: false }
    ];
  }
  sh.getRange(2, 1, lastRow - 1, 9).sort(sortSpecs);
}

function kyodoManagementHideTenureRowsExceptMonth_(sh, month) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var values = sh.getRange(2, 6, lastRow - 1, 1).getValues();
  var start = null;
  var length = 0;

  for (var i = 0; i < values.length; i++) {
    var rowNumber = i + 2;
    var shouldHide = Number(values[i][0]) !== month;
    if (shouldHide) {
      if (start === null) {
        start = rowNumber;
        length = 1;
      } else {
        length++;
      }
    } else if (start !== null) {
      sh.hideRows(start, length);
      start = null;
      length = 0;
    }
  }
  if (start !== null) sh.hideRows(start, length);
}

function kyodoManagementViewMessage_(mode) {
  var labels = {
    tenure_desc: "表示を「在籍期間が長い順」にしました。",
    enroll_desc: "表示を「入会日が新しい順」にしました。",
    withdraw_desc: "表示を「退会手続き日が新しい順」にしました。"
  };
  var monthMatch = String(mode || "").match(/^month_([1-5])$/);
  if (monthMatch) return monthMatch[1] + "ヶ月の人だけ表示しました。";
  return labels[mode] || "全員を表示しました。";
}

/** 既存のテスト下書き画面を開く */
function kyodoManagementOpenTestDraft() {
  kyodoCreateDraftTest();
}

/** 既存の本番送信画面を開く（既存の最終確認も維持） */
function kyodoManagementOpenProductionSend() {
  var preview = kyodoGetTenureSendPreview_();
  if (!preview.pending) {
    throw new Error("チェックONかつ未送信の対象者がいません。");
  }
  kyodoSendCampaignMailToRed();
}

function kyodoBuildCampaignManagementHtml_() {
  var css = typeof joyfitManagementCss_ === "function" ? joyfitManagementCss_() : "";
  var commonJs = typeof joyfitManagementCommonJs_ === "function" ? joyfitManagementCommonJs_() : "";
  var tenureLink = typeof joyfitSheetLink_ === "function"
    ? joyfitSheetLink_("6カ月継続管理")
    : "6カ月継続管理";
  var manual = typeof joyfitManualHtml_ === "function"
    ? joyfitManualHtml_([
      joyfitManualStep_(1,
        "「更新」を押します。",
        "2026/4/1以降入会・満6ヶ月未満が " + tenureLink + " に入ります（在籍期間が長い順）。"),
      joyfitManualStep_(2,
        tenureLink + " で「送信対象」にチェックを入れます。",
        "チェックした人が「未送信」に入ります。退会を取り消した人は「退会キャンセル」にチェック（退会していなかった扱い）。"),
      joyfitManualStep_(3,
        "「テスト下書き」で文面を確認します。",
        "CASE1〜3 に件名・本文を保存できます。自分宛の下書きだけ作られ、会員には送られません。"),
      joyfitManualStep_(4,
        "「チェックした人へメール」を押します。",
        "「未送信」の人だけに送ります。送ると「送信済」に変わります。")
    ].join(""))
    : "";

  return `<!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <style>
      ${css}
      .view-panel{display:flex;flex-direction:column;gap:10px}
      .view-title{font-size:12px;font-weight:900;letter-spacing:.08em;color:#5f6368;text-transform:uppercase}
      .view-row{display:grid;grid-template-columns:1fr 1fr;gap:8px}
      .view-row.months{grid-template-columns:repeat(5,1fr)}
      .view-btn{border:1px solid #dfe3e8;background:#fff;color:#111;border-radius:10px;padding:10px 8px;font-weight:800;font-size:12px;cursor:pointer;box-shadow:0 2px 0 #d7dce2}
      .view-btn:hover{border-color:#c21632;color:#c21632}
      .view-btn.primary{background:#111;color:#fff;border-color:#111}
      .view-btn.danger{background:#c21632;color:#fff;border-color:#c21632}
    </style>
  </head>
  <body>
    <header class="hero">
      <div class="brand">JOYFIT24 KYODO</div>
      <h1>6カ月継続管理</h1>
      <div class="state"><span class="dot"></span><span id="message">確認中...</span></div>
    </header>
    <main>
      <div id="error"></div>
      <section class="card">
        <div class="grid">
          <div class="metric"><b id="total">-</b><span>リスト</span></div>
          <div class="metric"><b id="pending">-</b><span>未送信</span></div>
          <div class="metric"><b id="sent">-</b><span>送信済</span></div>
          <div class="metric"><b id="missing">-</b><span>番号なし</span></div>
        </div>
        <div id="guide" class="meta"></div>
        <div id="memberInfo" class="meta"></div>
        ${manual}
      </section>
      <section class="card">
        <button class="action sync" onclick="syncAll(this)">↻ 更新</button>
        <div id="progress" class="progress"></div>
        <div class="view-panel">
          <div class="view-title">表示切替</div>
          <div class="view-row">
            <button class="view-btn danger" onclick="changeView(this,'tenure_desc')">期間順 5→1</button>
            <button class="view-btn" onclick="changeView(this,'enroll_desc')">入会順</button>
            <button class="view-btn" onclick="changeView(this,'withdraw_desc')">退会順</button>
            <button class="view-btn primary" onclick="changeView(this,'all')">全表示</button>
          </div>
          <div class="view-row months">
            <button class="view-btn" onclick="changeView(this,'month_1')">1</button>
            <button class="view-btn" onclick="changeView(this,'month_2')">2</button>
            <button class="view-btn" onclick="changeView(this,'month_3')">3</button>
            <button class="view-btn" onclick="changeView(this,'month_4')">4</button>
            <button class="view-btn" onclick="changeView(this,'month_5')">5</button>
          </div>
        </div>
        <button class="action test" onclick="openTest(this)">✉ テスト下書き</button>
        <button class="action send" onclick="openProduction(this)">✓ チェックした人へメール</button>
      </section>
      <div class="foot">表示切替は見え方だけです<br>送信前はチェックONの人を確認してください</div>
    </main>

    <script>
      ${commonJs}
      var busyLock = false;
      function setBusyBtn(btn, busy) {
        busyLock = !!busy;
        setBusy(busy);
        if (btn) btn.disabled = !!busy;
      }
      function render(data) {
        document.getElementById('message').textContent = data.message;
        document.getElementById('total').textContent = data.total;
        document.getElementById('pending').textContent = data.pending;
        document.getElementById('sent').textContent = data.sent;
        document.getElementById('missing').textContent = data.memberNumberMissing;
        document.getElementById('guide').textContent = data.guide || '';
        document.getElementById('memberInfo').textContent = data.memberMasterAvailable
          ? '入会者マスタ ' + data.memberMasterRows + '件（番号の自動補完用）'
          : '入会者マスタなし（番号は手入力してください）';
      }
      function refresh() {
        if (busyLock) return;
        google.script.run.withSuccessHandler(render).withFailureHandler(showError).kyodoGetCampaignManagementSummary();
      }
      function syncAll(btn) {
        clearError(); setBusyBtn(btn, true); setProgress('更新中...');
        var finished = false;
        var watchdog = setTimeout(function() {
          if (finished) return;
          setBusyBtn(btn, false);
          setProgress('');
          showError({message:'更新に3分以上かかっています。処理が継続している場合があります。少し待ってから管理画面を開き直してください。'});
        }, 180000);
        google.script.run
          .withSuccessHandler(function(data){ finished=true; clearTimeout(watchdog); setBusyBtn(btn,false); setProgress(''); render(data); })
          .withFailureHandler(function(err){ finished=true; clearTimeout(watchdog); setBusyBtn(btn,false); setProgress(''); showError(err); })
          .kyodoManagementRefreshList();
      }
      function changeView(btn, mode) {
        clearError(); setBusyBtn(btn, true); setProgress('表示を切替中...');
        google.script.run
          .withSuccessHandler(function(data){ setBusyBtn(btn,false); setProgress(''); render(data); })
          .withFailureHandler(function(e){ setBusyBtn(btn,false); setProgress(''); showError(e); })
          .kyodoManagementApplyTenureView(mode);
      }
      function openTest(btn) {
        clearError(); setBusyBtn(btn, true);
        google.script.run.withSuccessHandler(function(){setBusyBtn(btn,false);}).withFailureHandler(function(e){setBusyBtn(btn,false);showError(e);}).kyodoManagementOpenTestDraft();
      }
      function openProduction(btn) {
        clearError(); setBusyBtn(btn, true);
        google.script.run.withSuccessHandler(function(){setBusyBtn(btn,false); refresh();}).withFailureHandler(function(e){setBusyBtn(btn,false);showError(e);}).kyodoManagementOpenProductionSend();
      }
      refresh();
      // 開きっぱなしでも負荷を増やさない。非表示中は呼び出さず、30秒ごとに更新する。
      setInterval(function() {
        if (!document.hidden) refresh();
      }, 30000);
    </script>
  </body>
  </html>`;
}

/**
 * JOYFIT24経堂 — オプション契約メール集計（メイン）
 *
 * スプレッドシート:
 * https://docs.google.com/spreadsheets/d/14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w/
 *
 * 【自動】平日21:00 / 土日祝20:00 → OP数値更新 → 日報送信（jf-kyoudou）
 * カスタムメニューは出さない。数値はトリガーが書く。
 */

const OPTION_SPREADSHEET_ID = "14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w";

const SEARCH_QUERY =
  'from:info@joyfit-service.jp subject:(オプションご契約につきまして OR ご入会ありがとうございます)';

/** 追加・停止メールのみ（入会時OPは入会メールから取る） */
const OPTION_CHANGE_SEARCH_QUERY =
  'from:info@joyfit-service.jp subject:オプションご契約につきまして';

/** @deprecated 入会時OPラベルは使わない（入会メール側で持つ） */
const LABEL_OPTION_SIGNUP = "オプションメール/入会時";
const LABEL_OPTION_CHANGE = "オプションメール/追加・停止";

/** 「更新」時にラベル付与・既読にする直近日数 */
const DAILY_MAIL_HYGIENE_DAYS = 7;

/** 集計(A〜H) ＋ 契約明細(I列〜) を1枚にまとめたシート */
const SHEET_NAME_OP = "OP集計";
const SHEET_NAME_LOG_LEGACY = "OPデータ";
const SHEET_NAME_SUMMARY_LEGACY = "集計";
const OP_LOG_START_COL = 9;           // I列 … 契約明細の開始（新しい契約が上）
const OP_LOG_OLD_START_COL = 20;      // 旧T列（一度だけIへ移行）
const SHEET_NAME_OPENING_ARCHIVE = "月初アーカイブ";
const SUMMARY_COL_OPENING_MAIL = 7;   // G … 月初(日報メールより)
const SUMMARY_COL_NEXT_OPENING = 8;   // H … G(月初)+F(純増減)

// 旧名互換（内部では getOpSheet_ を使う）
const SHEET_NAME_LOG = SHEET_NAME_OP;
const SHEET_NAME_SUMMARY = SHEET_NAME_OP;

const NIPPO_MAIL_FROM = "jf-kyoudou@okamoto-group.co.jp";
const NIPPO_MAIL_SUBJECT_TAG = "報連相";

const OP_LOG_HEADERS = [
  "受信日時", "氏名", "区分",
  "オプション名(メール記載)", "オプション名(集計用)", "メールID"
];

const OPTION_LIST = [
  "安心サポート", "安心サポートVIP", "水素水", "オンラインレッスン",
  "体組成計", "契約ロッカー1,500", "レンタルマット", "プロテイン12杯",
  "プロテイン無制限", "プロテイン＋水素水", "レンタルタオル", "タンニング",
  "セルフエステ", "ホットスタジオ", "ヨガロッカー", "ピラティスリフォーマー"
];

// ── メニュー・イベント ──

function getBoundSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) { /* clasp / 時間トリガー */ }
  return SpreadsheetApp.openById(OPTION_SPREADSHEET_ID);
}

function onOpen() {
  installAnketoMenu_();
  try {
    const sh = getBoundSpreadsheet_().getSheetByName(SHEET_NAME_OP);
    if (sh) migrateOpLogTtoI_(sh);
  } catch (e) {
    Logger.log("OP明細配置: " + e);
  }
}

/** 残すメニューはアンケートだけ。入会・退会／OP／6ヶ月はシートは残し、メニューには出さない。 */
function installAnketoMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("アンケート")
    .addItem("管理を開く", "openAnketoManagement")
    .addToUi();
}

/** @deprecated 互換。アンケート以外のカスタムメニューは出さない */
function installSimpleUpdateMenu_() {
  installAnketoMenu_();
}

/** @deprecated 互換。アンケート以外のカスタムメニューは出さない */
function installJoyfitMenu_() {
  installAnketoMenu_();
}

/** 廃止した競合分析の週次トリガーを外す（関数本体は削除済み） */
function removeCompetitorTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === "runCompetitorAnalysisWeekly_" || fn === "runCompetitorAnalysisAuto") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * 【①ラベルを修正】
 * 過去分も含め、入会・退会・OP追加停止のラベルを正しい状態に付け直す。
 * ※件数が多いとGASの6分制限で途中停止することがある → 完了するまで①を繰り返し実行。
 */
function fixJoyfitMailLabels() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    ui.alert("ラベル修正", "別の処理が実行中です。少し待ってから再度お試しください。", ui.ButtonSet.OK);
    return;
  }
  try {
    const result = fixJoyfitMailLabelsCore_();
    if (result.done) {
      ui.alert(
        "① ラベルを修正",
        "すべて完了しました。\n\n" +
          "・旧「" + LABEL_OPTION_SIGNUP + "」解除 … " + result.removedSignupLabel + " スレッド\n" +
          "・入会ラベル付け直し … " + result.enrollThreads + " スレッド\n" +
          "・退会ラベル付け直し … " + result.withdrawThreads + " スレッド\n" +
          "・OP追加・停止ラベル … " + result.optionChangeThreads + " スレッド\n\n" +
          "次はメニュー「更新 → ② 数値を更新」を実行してください。",
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        "① ラベルを修正（続きあり）",
        "時間制限のため途中まで完了しました。\n\n" +
          result.progressText + "\n\n" +
          "もう一度「① ラベルを修正」を押して続きを実行してください。\n" +
          "（過去メール全部を正しく付け直す作業です）",
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    ui.alert("ラベル修正エラー", String(err && err.message ? err.message : err), ui.ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

const JOYFIT_LABEL_FIX_PROP_ = "joyfitLabelFixState_v1";
const JOYFIT_LABEL_FIX_TIME_MS_ = 4.5 * 60 * 1000;

function fixJoyfitMailLabelsCore_() {
  const props = PropertiesService.getScriptProperties();
  let state = {};
  try {
    state = JSON.parse(props.getProperty(JOYFIT_LABEL_FIX_PROP_) || "{}");
  } catch (e) {
    state = {};
  }
  if (!state.step) {
    state = {
      step: "strip",
      removedSignupLabel: 0,
      enrollThreads: 0,
      withdrawThreads: 0,
      optionChangeThreads: 0,
      enrollStart: 0,
      withdrawStart: 0,
      optionStart: 0
    };
  }

  const started = Date.now();
  function timeUp_() {
    return Date.now() - started > JOYFIT_LABEL_FIX_TIME_MS_;
  }
  function save_() {
    props.setProperty(JOYFIT_LABEL_FIX_PROP_, JSON.stringify(state));
  }
  function progressText_() {
    return (
      "進捗 … " + state.step + "\n" +
      "旧入会時ラベル解除 " + state.removedSignupLabel +
      " / 入会 " + state.enrollThreads +
      " / 退会 " + state.withdrawThreads +
      " / OP追加停止 " + state.optionChangeThreads
    );
  }

  if (typeof installMembershipLabelsSilent_ === "function") {
    installMembershipLabelsSilent_();
  }
  installOptionMailLabelsSilent_();

  // ── 1) 旧「オプションメール/入会時」を全解除 ──
  if (state.step === "strip") {
    while (!timeUp_()) {
      const n = stripDeprecatedOptionSignupLabelBatch_(100);
      state.removedSignupLabel += n;
      if (n === 0) {
        state.step = "enroll";
        state.enrollStart = 0;
        break;
      }
      save_();
    }
    save_();
    if (state.step === "strip") {
      return { done: false, progressText: progressText_(), removedSignupLabel: state.removedSignupLabel, enrollThreads: state.enrollThreads, withdrawThreads: state.withdrawThreads, optionChangeThreads: state.optionChangeThreads };
    }
  }

  const enrollQuery = typeof NYUKAI_SEARCH_QUERY !== "undefined"
    ? NYUKAI_SEARCH_QUERY
    : 'from:info@joyfit-service.jp subject:ご入会ありがとうございます';
  const withdrawQuery = typeof TAIKAI_SEARCH_QUERY !== "undefined"
    ? TAIKAI_SEARCH_QUERY
    : 'from:info@joyfit-service.jp subject:ご退会のお手続きについて';

  // ── 2) 入会ラベル（過去すべて・100件ずつ） ──
  if (state.step === "enroll") {
    while (!timeUp_()) {
      const batch = GmailApp.search(enrollQuery, state.enrollStart, 100);
      if (!batch.length) {
        state.step = "withdraw";
        state.withdrawStart = 0;
        break;
      }
      if (typeof applyNyukaiLabelsToThreads_ === "function") {
        applyNyukaiLabelsToThreads_(batch);
      }
      state.enrollThreads += batch.length;
      state.enrollStart += batch.length;
      save_();
      if (batch.length < 100) {
        state.step = "withdraw";
        state.withdrawStart = 0;
        break;
      }
    }
    save_();
    if (state.step === "enroll") {
      return { done: false, progressText: progressText_(), removedSignupLabel: state.removedSignupLabel, enrollThreads: state.enrollThreads, withdrawThreads: state.withdrawThreads, optionChangeThreads: state.optionChangeThreads };
    }
  }

  // ── 3) 退会ラベル（過去すべて） ──
  if (state.step === "withdraw") {
    while (!timeUp_()) {
      const batch = GmailApp.search(withdrawQuery, state.withdrawStart, 100);
      if (!batch.length) {
        state.step = "option";
        state.optionStart = 0;
        break;
      }
      if (typeof applyTaikaiLabelsToThreads_ === "function") {
        applyTaikaiLabelsToThreads_(batch);
      }
      state.withdrawThreads += batch.length;
      state.withdrawStart += batch.length;
      save_();
      if (batch.length < 100) {
        state.step = "option";
        state.optionStart = 0;
        break;
      }
    }
    save_();
    if (state.step === "withdraw") {
      return { done: false, progressText: progressText_(), removedSignupLabel: state.removedSignupLabel, enrollThreads: state.enrollThreads, withdrawThreads: state.withdrawThreads, optionChangeThreads: state.optionChangeThreads };
    }
  }

  // ── 4) OP追加・停止のみ（過去すべて） ──
  if (state.step === "option") {
    while (!timeUp_()) {
      const batch = GmailApp.search(OPTION_CHANGE_SEARCH_QUERY, state.optionStart, 100);
      if (!batch.length) {
        state.step = "done";
        break;
      }
      applyOptionChangeLabelsToThreads_(batch);
      state.optionChangeThreads += batch.length;
      state.optionStart += batch.length;
      save_();
      if (batch.length < 100) {
        state.step = "done";
        break;
      }
    }
    save_();
    if (state.step === "option") {
      return { done: false, progressText: progressText_(), removedSignupLabel: state.removedSignupLabel, enrollThreads: state.enrollThreads, withdrawThreads: state.withdrawThreads, optionChangeThreads: state.optionChangeThreads };
    }
  }

  props.deleteProperty(JOYFIT_LABEL_FIX_PROP_);
  return {
    done: true,
    removedSignupLabel: state.removedSignupLabel,
    enrollThreads: state.enrollThreads,
    withdrawThreads: state.withdrawThreads,
    optionChangeThreads: state.optionChangeThreads,
    membershipThreads: state.enrollThreads + state.withdrawThreads
  };
}

/** 旧「オプションメール/入会時」を最大 batchSize 件だけ外す */
function stripDeprecatedOptionSignupLabelBatch_(batchSize) {
  const label = GmailApp.getUserLabelByName(LABEL_OPTION_SIGNUP);
  if (!label) return 0;
  const n = Math.max(1, Math.min(100, batchSize || 100));
  const threads = label.getThreads(0, n);
  for (let i = 0; i < threads.length; i++) {
    threads[i].removeLabel(label);
  }
  return threads.length;
}

/** @deprecated 互換 … バッチ解除を繰り返す */
function stripDeprecatedOptionSignupLabelAll_() {
  let total = 0;
  for (let i = 0; i < 50; i++) {
    const n = stripDeprecatedOptionSignupLabelBatch_(100);
    total += n;
    if (n === 0) break;
  }
  return total;
}

/**
 * 【②数値を更新】日報B1の月を正として、当月のOP・入会退会を最新化
 * （ラベル作業は①で行う。ここでは数値のみ）
 */
function runSimpleDailyUpdate() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    ui.alert(
      "数値更新",
      "別の処理が実行中です。\n上の「スクリプトを実行しています」が消えてから、もう一度「② 数値を更新」を押してください。",
      ui.ButtonSet.OK
    );
    return;
  }
  try {
    const result = runSimpleDailyUpdateCore_(false);
    ui.alert("② 数値を更新", result.message, ui.ButtonSet.OK);
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    ui.alert(
      "数値更新エラー",
      msg +
        (msg.indexOf("タイムアウト") !== -1
          ? "\n\n処理を軽くしました。上の実行バーが消えてから、もう一度実行してください。"
          : ""),
      ui.ButtonSet.OK
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Gmailを再検索せず、当月OPログの同一人物×同一オプション重複だけ畳む。
 * clasp run / 手動修復用。
 */
function repairCurrentMonthOpDuplicates_() {
  const ss = getBoundSpreadsheet_();
  const ym = resolveNippoTargetYearMonth_(ss);
  ensureSummaryMonthB1_(ss, ym.year, ym.month);
  const opSheet = getOpSheet_(ss);
  const logData = readOpLogValues_(opSheet);
  const header = logData.length ? logData[0] : OP_LOG_HEADERS.slice();
  const keptOther = [];
  const keptThis = [];
  logData.slice(1).forEach(function (row) {
    const d = parseOpLogDate_(row[0]);
    if (d && d.getFullYear() === ym.year && d.getMonth() === ym.month) {
      keptThis.push(row);
    } else {
      keptOther.push(row);
    }
  });
  const uniqueThis = compactOpLogRows_(keptThis);
  writeOpLogValues_(opSheet, [header].concat(keptOther).concat(uniqueThis));
  syncSummarySheetFromOpData_(ss, ym.year, ym.month, opSheet);
  updateNippoSheetForMonth(ss, ym.year, ym.month, opSheet);
  try {
    if (typeof refreshMembershipDisplay_ === "function") refreshMembershipDisplay_(ss, ym);
  } catch (e) {
    Logger.log("入会・退会表示: " + e);
  }
  Logger.log(
    "OP重複修復 " + ym.label +
      " 当月 " + keptThis.length + " → " + uniqueThis.length
  );
  return {
    ok: true,
    label: ym.label,
    before: keptThis.length,
    after: uniqueThis.length
  };
}

/** 自動トリガーからも使える本体（ダイアログなし） */
function runSimpleDailyUpdateCore_(silent) {
  // setupSpreadsheet() は重いので数値更新では呼ばない（タイムアウト原因）
  const ss = getBoundSpreadsheet_();
  const nippoYm = resolveNippoTargetYearMonth_(ss);
  ensureSummaryMonthB1_(ss, nippoYm.year, nippoYm.month);

  // アンケートは別メニュー想定。数値更新ではスキップして時間を稼ぐ
  const opSheet = getOpSheet_(ss);
  executeFetchMonthForYm_(ss, opSheet, nippoYm.year, nippoYm.month, true, { light: true });

  // executeFetchMonth 内で日報反映済み。ここでは集計結果だけ読む
  const opCounts = countOpDataForMonth_(opSheet, nippoYm.year, nippoYm.month).counts;
  const contractTotal = sumOpCountField_(opCounts, "contract");
  const stopTotal = sumOpCountField_(opCounts, "stop");
  const newSignupTotal = sumOpCountField_(opCounts, "newSignup");
  const opAddTotal = sumOpCountField_(opCounts, "opAdd");
  const nippo = ss.getSheetByName("日報");
  const anchor = nippo ? resolveNippoOpAnchor_(nippo) : { headerRow: NIPPO_OP_HEADER_ROW, startRow: NIPPO_OP_START_ROW };

  try {
    if (typeof syncMembershipDailyCountsForMonth_ === "function") {
      syncMembershipDailyCountsForMonth_(ss, nippoYm.year, nippoYm.month);
    }
  } catch (err) {
    Logger.log("日別入会・退会: " + (err && err.message ? err.message : err));
  }

  // 表示シートの全面再描画は重いので、失敗しても数値更新は成功扱いにする
  try {
    if (typeof refreshMembershipDisplay_ === "function") {
      refreshMembershipDisplay_(ss, nippoYm);
    }
  } catch (e) {
    Logger.log("入会・退会表示: " + e);
  }

  const diag = LAST_ENROLL_OP_DIAG_ || {};
  const diagLine = diag.targets
    ? ("・入会メール診断 … 対象" + diag.targets + " / OPあり" + diag.withOp +
      " / OPなし" + diag.withoutOp + "（シート「OP取込診断」）\n")
    : "";

  return {
    year: nippoYm.year,
    month: nippoYm.month,
    label: nippoYm.label,
    sheetName: formatMonthlySheetName_(nippoYm.year, nippoYm.month),
    message:
      "完了しました。\n\n" +
      "対象月 … " + nippoYm.label + "（日報B1＝当月キャッチ）\n" +
      "・OP契約合計 " + contractTotal + "（新規入会内訳 " + newSignupTotal +
      " + 追加 " + opAddTotal + "） / 解約 " + stopTotal + "\n" +
      diagLine +
      "・日報反映 … " + anchor.startRow + "行目〜\n\n" +
      "※OP契約 = 入会メール内訳 + 追加(利用開始)\n" +
      "※OP解約 = 追加・停止メールの利用停止\n" +
      "※ラベルがおかしいときは先に「① ラベルを修正」"
  };
}

/**
 * 直近N日の入会・退会・OP追加停止にラベルを付け、未読を既読にする
 */
function prepareRecentMailHygiene_(daysBack) {
  const days = daysBack == null ? DAILY_MAIL_HYGIENE_DAYS : daysBack;
  if (typeof installMembershipLabelsSilent_ === "function") {
    installMembershipLabelsSilent_();
  }
  installOptionMailLabelsSilent_();

  let membershipThreads = 0;
  let optionChangeThreads = 0;
  let markedRead = 0;

  if (typeof applyMembershipLabelsRecent_ === "function") {
    membershipThreads = applyMembershipLabelsRecent_(days) || 0;
  }
  optionChangeThreads = applyOptionChangeLabelsRecent_(days);

  if (typeof markRecentJoyfitMailsRead_ === "function") {
    markedRead = markRecentJoyfitMailsRead_(days);
  } else {
    markedRead = markRecentJoyfitMailsReadFallback_(days);
  }

  // 旧「オプションメール/入会時」は入会メールから外す
  stripDeprecatedOptionSignupLabelRecent_(days);

  return {
    membershipThreads: membershipThreads,
    optionChangeThreads: optionChangeThreads,
    markedRead: markedRead
  };
}

function applyOptionChangeLabelsRecent_(daysBack) {
  const days = daysBack == null ? DAILY_MAIL_HYGIENE_DAYS : daysBack;
  const afterQuery = " after:" + (
    typeof gmailAfterDays_ === "function"
      ? gmailAfterDays_(days)
      : Utilities.formatDate(
          new Date(Date.now() - days * 86400000),
          Session.getScriptTimeZone(),
          "yyyy/MM/dd"
        )
  );
  const searchFn = typeof searchGmailRecentThreads_ === "function"
    ? searchGmailRecentThreads_
    : function (q, n) { return GmailApp.search(q, 0, Math.min(100, n || 100)); };
  const threads = searchFn(OPTION_CHANGE_SEARCH_QUERY + afterQuery, 100);
  applyOptionChangeLabelsToThreads_(threads);
  return threads.length;
}

function stripDeprecatedOptionSignupLabelRecent_(daysBack) {
  const label = GmailApp.getUserLabelByName(LABEL_OPTION_SIGNUP);
  if (!label) return;
  const days = daysBack == null ? DAILY_MAIL_HYGIENE_DAYS : daysBack;
  const after = typeof gmailAfterDays_ === "function"
    ? gmailAfterDays_(days)
    : Utilities.formatDate(new Date(Date.now() - days * 86400000), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const searchFn = typeof searchGmailRecentThreads_ === "function"
    ? searchGmailRecentThreads_
    : function (q, n) { return GmailApp.search(q, 0, Math.min(100, n || 100)); };
  const q = 'from:info@joyfit-service.jp subject:ご入会ありがとうございます after:' + after;
  searchFn(q, 100).forEach(function (thread) {
    thread.removeLabel(label);
  });
}

function markRecentJoyfitMailsReadFallback_(daysBack) {
  const days = daysBack == null ? DAILY_MAIL_HYGIENE_DAYS : daysBack;
  const after = typeof gmailAfterDays_ === "function"
    ? gmailAfterDays_(days)
    : Utilities.formatDate(new Date(Date.now() - days * 86400000), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const searchFn = typeof searchGmailRecentThreads_ === "function"
    ? searchGmailRecentThreads_
    : function (q, n) { return GmailApp.search(q, 0, Math.min(100, n || 100)); };
  const queries = [
    'from:info@joyfit-service.jp subject:ご入会ありがとうございます after:' + after,
    'from:info@joyfit-service.jp subject:ご退会のお手続きについて after:' + after,
    OPTION_CHANGE_SEARCH_QUERY + " after:" + after
  ];
  let count = 0;
  queries.forEach(function (q) {
    searchFn(q, 100).forEach(function (thread) {
      if (thread.isUnread()) {
        thread.markRead();
        count++;
      }
    });
  });
  return count;
}

/** Apps Script エディタから実行可 … 読み込みエラーが無いか確認 */
function diagnoseJoyfitScript_() {
  const checks = [
    ["OPTION_LIST", typeof OPTION_LIST !== "undefined" && OPTION_LIST.length],
    ["NIPPO_MAIL_FROM", typeof NIPPO_MAIL_FROM !== "undefined"],
    ["updateMembershipSheet", typeof updateMembershipSheet === "function"],
    ["kyodoImportEnrollmentSmart_", typeof kyodoImportEnrollmentSmart_ === "function"],
    ["kyodoOpenCampaignManagement", typeof kyodoOpenCampaignManagement === "function"],
    ["openMembershipMonthManagement", typeof openMembershipMonthManagement === "function"],
    ["openOptionManagement", typeof openOptionManagement === "function"],
    ["openAnketoManagement", typeof openAnketoManagement === "function"],
    ["runSimpleDailyUpdate", typeof runSimpleDailyUpdate === "function"],
    ["sendShopDailyReportSilent_", typeof sendShopDailyReportSilent_ === "function"],
    ["formatMonthlySheetName_", typeof formatMonthlySheetName_ === "function"],
    ["syncAnketoDisplay_", typeof syncAnketoDisplay_ === "function"]
  ];
  const failed = checks.filter(function (c) { return !c[1]; }).map(function (c) { return c[0]; });
  if (failed.length) {
    throw new Error("未読み込み: " + failed.join(", ") + " … gasフォルダの全ファイル貼り直しを確認");
  }
  Logger.log("JOYFIT スクリプト読み込み OK");
  return "OK。スプレッドシートのタブを閉じて開き直すと JOYFIT メニューが出ます。";
}

function onEdit(e) {
  if (!e || !e.range) return;
  if (handleAnketoSheetEdit_(e)) return;
  if (handleTaikaiSheetEdit_(e)) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() === SHEET_NAME_OP && e.range.getA1Notation() === "B1") {
    const ss = e.source;
    const opSheet = getOpSheet_(ss);
    if (!opSheet) return;
    archiveManualOpeningForLabel_(ss, opSheet, e.oldValue);
    const ym = resolveSummaryTargetYearMonth_(ss);
    syncSummarySheetFromOpData_(ss, ym.year, ym.month, opSheet);
    loadManualOpeningForMonth_(ss, opSheet, ym.year, ym.month);
    return;
  }

  if (sheet.getName() === SHEET_NAME_OP &&
      e.range.getColumn() <= SUMMARY_COL_OPENING_MAIL &&
      e.range.getLastColumn() >= SUMMARY_COL_OPENING_MAIL &&
      e.range.getRow() <= 2 + OPTION_LIST.length &&
      e.range.getLastRow() >= 3) {
    const ss = e.source;
    const opSheet = getOpSheet_(ss);
    const ym = resolveSummaryTargetYearMonth_(ss);
    archiveManualOpeningForLabel_(ss, opSheet, ym.label);
  }
}

/** G3:G18を指定月の手入力月初として保存 */
function archiveManualOpeningForLabel_(ss, opSheet, label) {
  const match = String(label || "").trim().match(/^(\d{4})年(\d{1,2})月$/);
  if (!match || !opSheet) return false;
  const raw = opSheet.getRange(3, SUMMARY_COL_OPENING_MAIL, OPTION_LIST.length, 1).getValues();
  const hasInput = raw.some(function (row) { return row[0] !== "" && row[0] !== null; });
  if (!hasInput) return false;
  const values = raw.map(function (row) {
    const n = parseMonthlyNumeric_(row[0]);
    return n === null ? 0 : n;
  });
  saveOpeningToArchive_(ss, parseInt(match[1], 10), parseInt(match[2], 10) - 1, values, "OP集計G列（手入力）");
  return true;
}

/** 月を切り替えたとき、保存済み月初をG列へ戻す。未保存なら空欄。 */
function loadManualOpeningForMonth_(ss, opSheet, year, month) {
  const archived = loadOpeningFromArchive_(ss, year, month);
  if (archived && archived.values) {
    writeSummaryOpeningColumn_(opSheet, archived.values);
  }
  // 保存値がない場合もG列は消さない。手入力値の保護を最優先する。
  ensureSummaryFormulas_(opSheet);
}

/** JOYFITメニュー: 現在表示中のG3:G18を明示的に保存 */
function saveCurrentManualOpening() {
  throw new Error("オプション月初の保存はメニューから外しています。月初は手入力で管理してください。");
}

// ── トリガー（初回のみ installOptionDailyTrigger を実行）──

// ── トリガー（毎日：数値更新してから日報送信。それ以外は置かない）──

function installOptionDailyTrigger() {
  const result = resetAllTriggersInstallDaily_();
  try {
    SpreadsheetApp.getUi().alert(
      "トリガー設定",
      "既存トリガーは全部消して、日報用だけ入れ直しました。\n\n" +
        "平日 21:00 … 数値更新 → 日報メール送信\n" +
        "土日祝 20:00 … 数値更新 → 日報メール送信\n" +
        "送信元は jf-kyoudou のままです。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log("installOptionDailyTrigger: " + JSON.stringify(result));
  }
  return result;
}

/** 全トリガー削除 → 日報前の数値更新＋送信だけ残す */
function resetAllTriggersInstallDaily_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("runDailyUpdateAndSendAt20_")
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .nearMinute(0)
    .create();
  ScriptApp.newTrigger("runDailyUpdateAndSendAt21_")
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .nearMinute(0)
    .create();
  return ScriptApp.getProjectTriggers().map(function (t) {
    return {
      fn: t.getHandlerFunction(),
      source: String(t.getTriggerSource()),
      event: String(t.getEventType())
    };
  });
}

function removeOptionDailyTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
}

function deleteCompetitorSheets_() {
  const ss = getBoundSpreadsheet_();
  const deleted = [];
  ss.getSheets().slice().forEach(function (sh) {
    if (ss.getSheets().length <= 1) return;
    const name = sh.getName();
    if (name === "競合分析" || name.indexOf("競合_") === 0) {
      ss.deleteSheet(sh);
      deleted.push(name);
    }
  });
  return deleted;
}

/**
 * clasp から実行: 不要トリガー全削除・日報用再設定・競合タブ削除・当月の数値更新
 * （日報メールは送らない）
 */
function bootstrapDailyOpAuto_() {
  const triggers = resetAllTriggersInstallDaily_();
  const competitorSheets = deleteCompetitorSheets_();
  const update = runSimpleDailyUpdateCore_(true);
  return {
    triggers: triggers,
    competitorSheets: competitorSheets,
    updateMessage: update && update.message,
    year: update && update.year,
    month: update && update.month,
    label: update && update.label
  };
}

/**
 * 土日祝 20:00 用 … 該当日だけ更新→送信
 */
function runDailyUpdateAndSendAt20_() {
  if (typeof isWeekendOrJapaneseHoliday_ !== "function" ||
      !isWeekendOrJapaneseHoliday_(new Date())) {
    Logger.log("runDailyUpdateAndSendAt20_: 平日のためスキップ");
    return;
  }
  runDailyUpdateAndSendSilent_();
}

/**
 * 平日 21:00 用 … 該当日だけ更新→送信
 */
function runDailyUpdateAndSendAt21_() {
  if (typeof isWeekendOrJapaneseHoliday_ === "function" &&
      isWeekendOrJapaneseHoliday_(new Date())) {
    Logger.log("runDailyUpdateAndSendAt21_: 土日祝のためスキップ");
    return;
  }
  runDailyUpdateAndSendSilent_();
}

/**
 * 数値更新 → 日報送信（曜日ラッパーから呼ぶ本体）
 * （未更新のまま送信されるのを防ぐ／同日二重送信防止は送信側）
 */
function runDailyUpdateAndSendSilent_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 60 * 1000)) {
    Logger.log("runDailyUpdateAndSendSilent_: 別処理実行中のためスキップ");
    return;
  }
  try {
    // 夜間自動: ラベル軽整備（直近）→ 当月数値 → 送信
    try { prepareRecentMailHygiene_(DAILY_MAIL_HYGIENE_DAYS); } catch (e) { Logger.log(e); }
    runSimpleDailyUpdateCore_(true);
    if (typeof sendShopDailyReportSilent_ === "function") {
      sendShopDailyReportSilent_();
    } else {
      Logger.log("sendShopDailyReportSilent_ が見つかりません");
    }
  } catch (err) {
    Logger.log("runDailyUpdateAndSendSilent_ エラー: " + (err && err.message ? err.message : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** 時間トリガーから呼ばれる（ダイアログなし）※一体型からも呼ぶ */
function smartUpdateDataSilent() {
  try {
    syncAnketoSilent_();
  } catch (err) {
    Logger.log("アンケート集計: " + err.message);
  }
  smartUpdateDataCore_(true);
}

/** 手動実行用（ダイアログあり） */
function smartUpdateData() {
  smartUpdateDataCore_(false);
}

function smartUpdateDataCore_(silent) {
  setupSpreadsheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const opSheet = getOpSheet_(ss);
  const logRows = readOpLogValues_(opSheet);

  // 日報の月を正として OP集計B1 を合わせる（日報オプション未反映の主因だった）
  const nippoYm = resolveNippoTargetYearMonth_(ss);
  ensureSummaryMonthB1_(ss, nippoYm.year, nippoYm.month);

  if (logRows.length <= 1) {
    if (silent) {
      executeFetchAll_(ss, opSheet, true);
    } else {
      const ui = SpreadsheetApp.getUi();
      const res = ui.alert(
        "初回セットアップ",
        "OP集計の明細が空です。過去すべてのメールを一括取得しますか？\n（はい推奨・数分かかる場合あり）",
        ui.ButtonSet.YES_NO
      );
      if (res === ui.Button.YES) {
        executeFetchAll_(ss, opSheet, false);
      } else {
        executeFetchMonthForYm_(ss, opSheet, nippoYm.year, nippoYm.month, false);
      }
    }
    return;
  }

  executeFetchMonthForYm_(ss, opSheet, nippoYm.year, nippoYm.month, silent);
}

// ── OP集計シート（A〜H=集計 / I列〜=契約明細・新しい順）──

function getOpSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_NAME_OP);
  const legacySummary = ss.getSheetByName(SHEET_NAME_SUMMARY_LEGACY);
  const legacyLog = ss.getSheetByName(SHEET_NAME_LOG_LEGACY);

  if (!sh && legacySummary) {
    legacySummary.setName(SHEET_NAME_OP);
    sh = legacySummary;
  }
  if (!sh) sh = ss.insertSheet(SHEET_NAME_OP);

  // 旧OPデータの明細をI列へ一度だけ移行
  if (legacyLog && legacyLog.getLastRow() > 0) {
    const hasLog = String(sh.getRange(1, OP_LOG_START_COL).getValue() || "").trim();
    if (!hasLog) {
      const data = legacyLog.getDataRange().getValues();
      if (data.length && data[0].length) {
        sh.getRange(1, OP_LOG_START_COL, data.length, data[0].length).setValues(data);
      }
    }
  }
  migrateOpLogTtoI_(sh);
  return sh;
}

/** 旧T列の契約明細をI列へ移す（1回だけ） */
function migrateOpLogTtoI_(opSheet) {
  if (!opSheet) return;
  const oldHeader = String(opSheet.getRange(1, OP_LOG_OLD_START_COL).getValue() || "").trim();
  const newHeader = String(opSheet.getRange(1, OP_LOG_START_COL).getValue() || "").trim();
  const oldLooksLikeLog = oldHeader.indexOf("受信") !== -1;
  const newLooksLikeLog = newHeader.indexOf("受信") !== -1;
  if (!oldLooksLikeLog || newLooksLikeLog) return;

  const last = opSheet.getLastRow();
  if (last < 1) return;
  const numCols = OP_LOG_HEADERS.length;
  const oldCol = opSheet.getRange(1, OP_LOG_OLD_START_COL, last, 1).getValues();
  let oldLast = 1;
  for (let i = oldCol.length - 1; i >= 0; i--) {
    if (oldCol[i][0] !== "" && oldCol[i][0] != null) {
      oldLast = i + 1;
      break;
    }
  }
  const data = opSheet.getRange(1, OP_LOG_OLD_START_COL, oldLast, numCols).getValues();
  const sorted = sortOpLogRowsNewestFirst_(data);
  opSheet.getRange(1, OP_LOG_START_COL, sorted.length, numCols).setValues(sorted);
  if (sorted.length > 1) {
    opSheet.getRange(2, OP_LOG_START_COL, sorted.length - 1, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  }
  opSheet.getRange(1, OP_LOG_OLD_START_COL, oldLast, numCols).clearContent();
  // 旧「▶ 契約明細」（S列）を消す
  try {
    opSheet.getRange(1, OP_LOG_OLD_START_COL - 1).clearContent();
  } catch (e) { /* ignore */ }
  opSheet.setFrozenColumns(SUMMARY_COL_NEXT_OPENING);
}

function readOpLogValues_(opSheet) {
  if (!opSheet) return [OP_LOG_HEADERS.slice()];
  const lastRow = opSheet.getLastRow();
  if (lastRow < 1) return [OP_LOG_HEADERS.slice()];
  const numCols = OP_LOG_HEADERS.length;
  const values = opSheet.getRange(1, OP_LOG_START_COL, lastRow, numCols).getValues();
  const out = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0 || values[i][0] || values[i][1] || values[i][5]) out.push(values[i]);
  }
  if (!out.length) return [OP_LOG_HEADERS.slice()];
  if (String(out[0][0] || "").indexOf("受信") === -1) {
    out.unshift(OP_LOG_HEADERS.slice());
  }
  return out;
}

function writeOpLogValues_(opSheet, rows) {
  if (!opSheet) return;
  const numCols = OP_LOG_HEADERS.length;
  const sorted = sortOpLogRowsNewestFirst_(rows);
  const writeLen = Math.max((sorted && sorted.length) || 0, 1);
  // シート全体の getLastRow() だと左の集計以外で膨らむことがある → I列側だけ見る
  const prevLast = findOpLogLastRow_(opSheet);
  const clearTo = Math.max(prevLast, writeLen, 1);
  opSheet.getRange(1, OP_LOG_START_COL, clearTo, numCols).clearContent();
  if (!sorted || !sorted.length) {
    opSheet.getRange(1, OP_LOG_START_COL, 1, numCols).setValues([OP_LOG_HEADERS]);
    return;
  }
  opSheet.getRange(1, OP_LOG_START_COL, sorted.length, sorted[0].length).setValues(sorted);
  const header = opSheet.getRange(1, OP_LOG_START_COL, 1, OP_LOG_HEADERS.length);
  header.setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");
  if (sorted.length > 1) {
    opSheet.getRange(2, OP_LOG_START_COL, sorted.length - 1, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  }
}

/** 契約明細を受信日時の新しい順にする（当月が上に来る） */
function sortOpLogRowsNewestFirst_(rows) {
  if (!rows || rows.length <= 1) return rows || [];
  const header = rows[0];
  const body = rows.slice(1).slice();
  body.sort(function (a, b) {
    const da = parseOpLogDate_(a[0]);
    const db = parseOpLogDate_(b[0]);
    const ta = da ? da.getTime() : 0;
    const tb = db ? db.getTime() : 0;
    if (tb !== ta) return tb - ta;
    // 同時刻なら氏名で安定ソート
    return String(a[1] || "").localeCompare(String(b[1] || ""), "ja");
  });
  return [header].concat(body);
}

/** OP明細（I列）の最終データ行 */
function findOpLogLastRow_(opSheet) {
  const last = opSheet.getLastRow();
  if (last < 1) return 1;
  const col = opSheet.getRange(1, OP_LOG_START_COL, last, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) {
    if (col[i][0] !== "" && col[i][0] != null) return i + 1;
  }
  return 1;
}

function applyOpLogHeaderStyle_(opSheet) {
  const header = opSheet.getRange(1, OP_LOG_START_COL, 1, OP_LOG_HEADERS.length);
  header.setValues([OP_LOG_HEADERS]);
  header.setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff")
    .setBorder(null, null, true, null, null, null, "#C21632", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const opSheet = getOpSheet_(ss);
  setupOpSheetLayout_(opSheet, ss);
}

function setupOpSheetLayout_(opSheet, ss) {
  // 左: 集計 A〜G
  opSheet.getRange("A1")
    .setValue("対象月を選択 ➡")
    .setFontWeight("bold")
    .setFontColor("#5f6368")
    .setHorizontalAlignment("right");

  const monthCell = opSheet.getRange("B1");
  monthCell.clearDataValidations();
  monthCell.clearFormat();
  monthCell.setNumberFormat("@");

  const monthList = buildOpMonthList_(ss);
  if (monthList.length > 0) {
    monthCell.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(monthList, true).build()
    );
    const val = String(monthCell.getDisplayValue() || monthCell.getValue() || "").trim();
    if (!monthList.includes(val)) monthCell.setValue(monthList[0]);
  } else {
    const today = new Date();
    monthCell.setValue(today.getFullYear() + "年" + (today.getMonth() + 1) + "月");
  }
  monthCell.setBackground("#f1f3f4").setFontWeight("bold").setFontColor("#111111")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null, "#C21632", SpreadsheetApp.BorderStyle.SOLID);

  opSheet.getRange(2, 1, 1, 8).setValues([[
    "オプション名", "利用開始(新規入会)", "利用開始(OP追加)",
    "利用開始合計", "利用停止数", "翌月の±", "月初(日報)", "翌月月初"
  ]]);
  opSheet.getRange("A2:H2").setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setBorder(null, null, true, null, null, null, "#C21632", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  opSheet.setFrozenRows(2);

  const existingNames = opSheet.getRange(3, 1, OPTION_LIST.length, 1).getValues();
  const needNames = existingNames.some(function (r, i) { return String(r[0] || "") !== OPTION_LIST[i]; });
  if (needNames) {
    // 名称列だけ直す。G列の手入力月初値は消さない。
    opSheet.getRange(3, 1, OPTION_LIST.length, 1)
      .setValues(OPTION_LIST.map(function (n) { return [n]; }));
  }
  ensureSummaryFormulas_(opSheet);
  opSheet.getRange(3, 1, OPTION_LIST.length, 8).setFontColor("#111111");
  for (let i = 0; i < OPTION_LIST.length; i++) {
    opSheet.getRange(3 + i, 1, 1, 8).setBackground(i % 2 === 0 ? "#ffffff" : "#f1f3f4");
  }
  // Gは確認済み月初の入力欄、Hは翌月月初の自動計算欄。
  opSheet.getRange(2, SUMMARY_COL_OPENING_MAIL).setBackground("#f6c344").setFontColor("#111111");
  opSheet.getRange(3, SUMMARY_COL_OPENING_MAIL, OPTION_LIST.length, 1)
    .setBackground("#fff2cc").setHorizontalAlignment("center").setNumberFormat("0")
    .setBorder(null, true, null, true, null, null, "#d6b656", SpreadsheetApp.BorderStyle.SOLID);
  opSheet.getRange(2, SUMMARY_COL_NEXT_OPENING).setBackground("#b7e1cd").setFontColor("#111111");
  opSheet.getRange(3, SUMMARY_COL_NEXT_OPENING, OPTION_LIST.length, 1)
    .setBackground("#e6f4ea").setHorizontalAlignment("center").setNumberFormat("0")
    .setBorder(null, true, null, true, null, null, "#81c995", SpreadsheetApp.BorderStyle.SOLID);
  opSheet.setColumnWidth(SUMMARY_COL_OPENING_MAIL, 110);
  opSheet.setColumnWidth(SUMMARY_COL_NEXT_OPENING, 110);

  // 右隣: 契約明細 I列〜（新しい契約が上）。Hは翌月月初なのでラベルは書かない
  const logHeader = String(opSheet.getRange(1, OP_LOG_START_COL).getValue() || "").trim();
  if (!logHeader || logHeader.indexOf("受信") === -1) {
    applyOpLogHeaderStyle_(opSheet);
  } else {
    applyOpLogHeaderStyle_(opSheet);
  }
  // 集計と明細の境が分かるよう、I1見出しに「新しい順」を補足しない（ヘッダは固定名）
  opSheet.setColumnWidth(OP_LOG_START_COL, 140);
  opSheet.setFrozenColumns(SUMMARY_COL_NEXT_OPENING);
}

/** 集計B1: OP集計の明細にある月 ＋ 当月 */
function buildOpMonthList_(ss) {
  const keys = {};
  const opSheet = getOpSheet_(ss);
  const data = readOpLogValues_(opSheet);
  for (let i = 1; i < data.length; i++) {
    const d = parseOpLogDate_(data[i][0]);
    if (!d) continue;
    const label = d.getFullYear() + "年" + (d.getMonth() + 1) + "月";
    keys[label] = d.getFullYear() * 100 + (d.getMonth() + 1);
  }
  const today = new Date();
  const cur = today.getFullYear() + "年" + (today.getMonth() + 1) + "月";
  keys[cur] = today.getFullYear() * 100 + (today.getMonth() + 1);

  return Object.keys(keys).sort(function (a, b) { return keys[b] - keys[a]; });
}

// ── OPメール取得 ──

function executeFetchMonth_(ss, logSheet, silent) {
  const opSheet = logSheet || getOpSheet_(ss);
  let targetStr = String(opSheet.getRange("B1").getValue() || "").trim();
  let match = targetStr.match(/^(\d{4})年(\d{1,2})月$/);
  if (!match) {
    const today = new Date();
    targetStr = today.getFullYear() + "年" + (today.getMonth() + 1) + "月";
    opSheet.getRange("B1").setValue(targetStr);
    match = [null, today.getFullYear(), today.getMonth() + 1];
  }
  const targetYear = parseInt(match[1], 10);
  const targetMonth = parseInt(match[2], 10) - 1;
  executeFetchMonthForYm_(ss, opSheet, targetYear, targetMonth, silent);
}

/** 指定年月のオプションメールを取り、OP集計・日報・入会退会まで反映
 * opt.light … 数値更新向け高速パス（ラベル付与・入会全検索・余分な書式を省略）
 */
function executeFetchMonthForYm_(ss, logSheet, targetYear, targetMonth, silent, opt) {
  opt = opt || {};
  const light = !!opt.light;
  const opSheet = logSheet || getOpSheet_(ss);
  ensureSummaryMonthB1_(ss, targetYear, targetMonth);

  // 入会データの新着だけ軽く追記（ラベル付与は①に任せる）
  try {
    if (typeof fetchMembershipEmailsIncremental_ === "function") {
      fetchMembershipEmailsIncremental_(ss, light ? 10 : 14, { skipLabels: true });
    } else if (!light) {
      fetchMembershipEmailsSilent();
    }
  } catch (err) {
    Logger.log("入会・退会取得エラー: " + err.message);
  }

  const logData = readOpLogValues_(opSheet);
  const header = logData.length ? logData[0] : OP_LOG_HEADERS.slice();
  const keptOtherMonths = [];
  const keptThisMonth = [];
  logData.slice(1).forEach(function (row) {
    const d = parseOpLogDate_(row[0]);
    if (d && d.getFullYear() === targetYear && d.getMonth() === targetMonth) {
      keptThisMonth.push(row);
    } else {
      keptOtherMonths.push(row);
    }
  });

  // 入会は「ご利用開始日」基準のため、メール受信が前後月にズレる分も拾う
  const prev = new Date(targetYear, targetMonth - 1, 1);
  const afterNext = new Date(targetYear, targetMonth + 2, 1);
  const afterStr = prev.getFullYear() + "/" + String(prev.getMonth() + 1).padStart(2, "0") + "/01";
  const beforeStr = afterNext.getFullYear() + "/" + String(afterNext.getMonth() + 1).padStart(2, "0") + "/01";

  // 入会メール + 追加停止メール（light も同じ検索。件数だけ上限）
  const gmailQuery = SEARCH_QUERY + " after:" + afterStr + " before:" + beforeStr;
  const threads = light
    ? searchGmailRecentThreadsSafe_(gmailQuery, 200)
    : searchGmailAllThreads_(gmailQuery);
  if (!light) applyOptionMailLabelsToThreads_(threads);
  const fetched = extractDataFromThreads(threads, targetYear, targetMonth);

  // 当月ログは消さずマージ（light の不完全検索で B/C/D が消えるのを防ぐ）
  // ただし同一人物×同一OP×開始/停止はメールIDが違っても1件（二重配信・入会ID補完の積み増し防止）
  const uniqueThisMonth = compactOpLogRows_(keptThisMonth);
  const seen = {};
  uniqueThisMonth.forEach(function (row) {
    seen[opLogDedupeKey_(row)] = true;
  });
  const added = [];
  function absorb_(rows) {
    for (let i = 0; i < (rows || []).length; i++) {
      const key = opLogDedupeKey_(rows[i]);
      if (seen[key]) continue;
      seen[key] = true;
      added.push(rows[i]);
    }
  }
  absorb_(fetched);
  const fromStored = extractOpFromStoredEnrollments_(ss, targetYear, targetMonth, seen);
  absorb_(fromStored);
  Logger.log(
    "OP取込 " + targetYear + "/" + (targetMonth + 1) +
      (light ? " [light]" : "") +
      " 当月既存=" + uniqueThisMonth.length +
      "（重複畳み前=" + keptThisMonth.length + "）" +
      " Gmail新規=" + fetched.length +
      " 入会ID補完=+" + fromStored.length +
      " 追記=" + added.length
  );

  writeOpLogValues_(opSheet, [header].concat(keptOtherMonths).concat(uniqueThisMonth).concat(added));
  try {
    if (typeof backfillEnrollmentsFromOpLog_ === "function") {
      backfillEnrollmentsFromOpLog_(ss, targetYear, targetMonth, opSheet);
    }
  } catch (enrollErr) {
    Logger.log("入会名簿をOPから補完: " + enrollErr);
  }
  syncSummarySheetFromOpData_(ss, targetYear, targetMonth, opSheet);
  updateNippoSheetForMonth(ss, targetYear, targetMonth, opSheet);

  if (!light) {
    try {
      if (typeof syncMembershipDailyCountsForMonth_ === "function") {
        syncMembershipDailyCountsForMonth_(ss, targetYear, targetMonth);
      } else {
        syncMembershipDailyCountsSilent_();
      }
      const monthlySheet = ss.getSheetByName(formatMonthlySheetName_(targetYear, targetMonth));
      if (monthlySheet) formatMonthlyDailyArea_(monthlySheet);
    } catch (err) {
      Logger.log("日別入会・退会の反映エラー: " + err.message);
    }
  }

  if (!silent) {
    const signupStats = countSignupOpStats_(uniqueThisMonth.concat(added));
    SpreadsheetApp.getUi().alert(
      "OP更新",
      (targetMonth + 1) + "月分を取り込みました。\n\n" +
        "・当月既存 " + uniqueThisMonth.length + " 行\n" +
        "・Gmail新規 " + fetched.length + " 行\n" +
        "・入会メールID補完 +" + fromStored.length + " 行\n" +
        "・追記 " + added.length + " 行\n" +
        "・うち「利用開始(新規入会)」 " + signupStats.rows + " 行 / " + signupStats.people + " 人\n\n" +
        "OP集計・日報も更新済みです。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function opLogKindBucket_(cat) {
  return String(cat || "").indexOf("利用停止") !== -1 ? "stop" : "start";
}

function opLogMonthStamp_(row) {
  const d = parseOpLogDate_(row && row[0]);
  if (!d) return "";
  return d.getFullYear() + "-" + String(d.getMonth() + 1);
}

/**
 * 同一契約の重複キー。メールIDは使わない。
 * （入会メールが2通＋入会ID補完で、6ヶ月割0円パックが2〜4倍になるため）
 */
function opLogDedupeKey_(row) {
  const name = String((row && row[1]) || "").replace(/\s+/g, "");
  const opt = String((row && row[4]) || "").trim();
  const bucket = opLogKindBucket_(row && row[2]);
  const ym = opLogMonthStamp_(row);
  const fallback = name ? "" : String((row && row[5]) || "");
  return ym + "|" + name + "|" + fallback + "|" + bucket + "|" + opt;
}

function compactOpLogRows_(rows) {
  const seen = {};
  const out = [];
  for (let i = 0; i < (rows || []).length; i++) {
    const key = opLogDedupeKey_(rows[i]);
    if (!key || seen[key]) continue;
    seen[key] = true;
    out.push(rows[i]);
  }
  return out;
}

/** light用: 件数上限つきGmail検索（全件ページングしない） */
function searchGmailRecentThreadsSafe_(query, maxThreads) {
  const limit = Math.max(1, Math.min(500, maxThreads || 200));
  try {
    if (typeof searchGmailRecentThreads_ === "function") {
      return searchGmailRecentThreads_(query, limit);
    }
  } catch (e) { /* fallthrough */ }
  const batch = GmailApp.search(query, 0, Math.min(100, limit));
  return batch || [];
}

function countSignupOpStats_(rows) {
  const people = {};
  let n = 0;
  for (let i = 0; i < (rows || []).length; i++) {
    const status = String(rows[i][2] || "");
    if (status.indexOf("新規入会") === -1) continue;
    n++;
    const key = String(rows[i][5] || "") || String(rows[i][1] || "");
    if (key) people[key] = true;
  }
  return { rows: n, people: Object.keys(people).length };
}

/**
 * 入会・退会_データにある当月入会メールIDを直接開き、OP内訳を取る
 * （Gmail検索だけでは入会35に対してOPが足りない問題の本命対策）
 */
var LAST_ENROLL_OP_DIAG_ = null;

function extractOpFromStoredEnrollments_(ss, targetYear, targetMonth, seen) {
  const out = [];
  seen = seen || {};
  LAST_ENROLL_OP_DIAG_ = {
    opened: 0, withOp: 0, withoutOp: 0, noMsgId: 0, enroll: 0, targets: 0, missingNames: []
  };
  if (typeof readEnrollData_ !== "function") return out;

  const dataSheet = ss.getSheetByName(
    typeof SHEET_NAME_DATA !== "undefined" ? SHEET_NAME_DATA : "入会・退会_データ"
  );
  if (!dataSheet) return out;

  const monthLabel = targetYear + "年" + (targetMonth + 1) + "月";
  const rows = readEnrollData_(dataSheet);
  const diagRows = [["氏名", "区分", "入会月", "OP件数", "検出OP", "本文に月分", "本文長"]];

  for (let i = 0; i < rows.length; i++) {
    const ymRaw = rows[i][2];
    const ym = typeof normalizeYearMonthLabel_ === "function"
      ? normalizeYearMonthLabel_(ymRaw)
      : String(ymRaw || "").trim();
    if (ym !== monthLabel) continue;
    LAST_ENROLL_OP_DIAG_.targets++;

    const personName = String(rows[i][1] || "").trim();
    const category = String(rows[i][3] || "").trim();
    const msgId = String(rows[i][4] || "").trim();
    if (!msgId) {
      LAST_ENROLL_OP_DIAG_.noMsgId++;
      diagRows.push([personName, category, ym, 0, "(メールIDなし)", "", ""]);
      continue;
    }

    try {
      const message = GmailApp.getMessageById(msgId);
      if (!message) {
        LAST_ENROLL_OP_DIAG_.noMsgId++;
        diagRows.push([personName, category, ym, 0, "(メール取得失敗)", "", ""]);
        continue;
      }
      LAST_ENROLL_OP_DIAG_.opened++;
      const emailDate = message.getDate();
      const body = getMessageBodyText_(message);
      const hasMonthShare = /月分/.test(body) ? "あり" : "なし";
      const parsed = parseSignupEmail(emailDate, body, msgId);
      const opts = [];
      let added = 0;
      for (let j = 0; j < parsed.length; j++) {
        const row = padLogDataRow_(parsed[j], msgId);
        // 入会月とズレたら、入会データの月に寄せる
        let d = parseOpLogDate_(row[0]);
        if (!d || d.getFullYear() !== targetYear || d.getMonth() !== targetMonth) {
          row[0] = new Date(targetYear, targetMonth, 1);
          d = row[0];
        }
        const key = opLogDedupeKey_(row);
        if (!OPTION_LIST.includes(String(row[4] || ""))) continue;
        opts.push(String(row[4] || ""));
        if (seen[key]) continue;
        seen[key] = true;
        out.push(row);
        added++;
      }
      const uniqOpts = opts.filter(function (v, idx, arr) { return arr.indexOf(v) === idx; });
      diagRows.push([
        personName || extractSignupName_(body),
        category,
        ym,
        uniqOpts.length,
        uniqOpts.join(" / "),
        hasMonthShare,
        String(body || "").length
      ]);
      if (uniqOpts.length > 0) {
        LAST_ENROLL_OP_DIAG_.withOp++;
      } else {
        LAST_ENROLL_OP_DIAG_.withoutOp++;
        if (LAST_ENROLL_OP_DIAG_.missingNames.length < 12) {
          LAST_ENROLL_OP_DIAG_.missingNames.push(personName || "(無名)");
        }
      }
    } catch (err) {
      Logger.log("入会メールID OP取得失敗: " + msgId + " / " + (err && err.message ? err.message : err));
      diagRows.push([personName, category, ym, 0, "(エラー)", "", ""]);
    }
  }

  try {
    writeOpEnrollDiagSheet_(ss, diagRows, monthLabel);
  } catch (e) {
    Logger.log("OP取込診断シート: " + (e && e.message ? e.message : e));
  }

  Logger.log(
    "入会ID補完: 対象=" + LAST_ENROLL_OP_DIAG_.targets +
      " 開封=" + LAST_ENROLL_OP_DIAG_.opened +
      " OPあり=" + LAST_ENROLL_OP_DIAG_.withOp +
      " OPなし=" + LAST_ENROLL_OP_DIAG_.withoutOp +
      " 行追加=" + out.length
  );
  return out;
}

function writeOpEnrollDiagSheet_(ss, diagRows, monthLabel) {
  const name = "OP取込診断";
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const prev = Math.max(sh.getLastRow(), 2);
  const cols = Math.max((diagRows[0] && diagRows[0].length) || 7, 7);
  sh.getRange(1, 1, prev, cols).clearContent();
  sh.getRange(1, 1).setValue("対象月: " + monthLabel + " / 更新: " + new Date());
  sh.getRange(2, 1, diagRows.length, diagRows[0].length).setValues(diagRows);
  sh.setFrozenRows(2);
}

function executeFetchAll_(ss, logSheet, silent) {
  const opSheet = logSheet || getOpSheet_(ss);
  const threads = searchGmailAllThreads_(SEARCH_QUERY);
  applyOptionMailLabelsToThreads_(threads);
  const newData = extractDataFromThreads(threads, null, null);
  if (newData.length > 0) {
    newData.sort(function (a, b) { return a[0].getTime() - b[0].getTime(); });
  }
  writeOpLogValues_(opSheet, [OP_LOG_HEADERS.slice()].concat(newData));

  setupSpreadsheet();

  const summaryYm = resolveSummaryTargetYearMonth_(ss);
  syncSummarySheetFromOpData_(ss, summaryYm.year, summaryYm.month, opSheet);

  const nippoYm = resolveNippoTargetYearMonth_(ss);
  updateNippoSheetForMonth(ss, nippoYm.year, nippoYm.month, opSheet);
  formatNippoOpArea_(ss);

  try {
    fetchMembershipEmailsSilent();
  } catch (err) {
    Logger.log("入会・退会取得エラー: " + err.message);
  }

  try {
    syncMembershipDailyCountsSilent_();
    const monthlySheet = ss.getSheetByName(formatMonthlySheetName_(nippoYm.year, nippoYm.month));
    if (monthlySheet) formatMonthlyDailyArea_(monthlySheet);
  } catch (err) {
    Logger.log("日別入会・退会の反映エラー: " + err.message);
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "OP一括取得",
      "合計 " + newData.length + " 件を取得しました。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/** 日報ボタン用: Gmail取得なしで日報C列だけ更新 */
function refreshNippoCurrentMonthFromLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const opSheet = getOpSheet_(ss);
  if (readOpLogValues_(opSheet).length < 2) return;

  const ym = resolveNippoTargetYearMonth_(ss);
  const summaryYm = resolveSummaryTargetYearMonth_(ss);
  if (summaryYm.year === ym.year && summaryYm.month === ym.month) {
    syncSummarySheetFromOpData_(ss, summaryYm.year, summaryYm.month, opSheet);
  }
  updateNippoSheetForMonth(ss, ym.year, ym.month, opSheet);
}

// ── Gmail共通 ──

function searchGmailAllThreads_(query) {
  const out = [];
  let start = 0;
  const pageSize = 100;
  const maxThreads = 10000;
  while (start < maxThreads) {
    const batch = GmailApp.search(query, start, pageSize);
    if (!batch.length) break;
    for (let i = 0; i < batch.length; i++) out.push(batch[i]);
    start += batch.length;
    if (batch.length < pageSize) break;
  }
  return out;
}

/** オプション「追加・停止」メールにだけラベルを付ける（入会時ラベルは使わない） */
function applyOptionChangeLabelsToThreads_(threads) {
  const changeLabel = getOrCreateOptionMailLabel_(LABEL_OPTION_CHANGE);
  const deprecatedSignup = GmailApp.getUserLabelByName(LABEL_OPTION_SIGNUP);
  (threads || []).forEach(function (thread) {
    let hasChange = false;
    thread.getMessages().forEach(function (message) {
      const subject = String(message.getSubject() || "");
      if (subject.indexOf("オプションご契約につきまして") !== -1) hasChange = true;
    });
    if (deprecatedSignup) thread.removeLabel(deprecatedSignup);
    thread.removeLabel(changeLabel);
    if (hasChange) thread.addLabel(changeLabel);
  });
}

/** @deprecated 互換名 … 中身は追加・停止のみ */
function applyOptionMailLabelsToThreads_(threads) {
  applyOptionChangeLabelsToThreads_(threads);
}

function getOrCreateOptionMailLabel_(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

/** 追加・停止ラベル作成だけ（入会時は作らない） */
function installOptionMailLabelsSilent_() {
  getOrCreateOptionMailLabel_(LABEL_OPTION_CHANGE);
}

/**
 * メニュー用: OP追加・停止ラベルを作成し、対象メールへ付与
 */
function installOptionMailLabelsCore_() {
  // 互換: フルのラベル修正へ委譲
  return fixJoyfitMailLabelsCore_();
}

function installOptionMailLabels() {
  fixJoyfitMailLabels();
}

function getMessageBodyText_(message) {
  const plain = String(message.getPlainBody() || "").trim();
  // 料金行が十分ある plain なら HTML を取りに行かない（タイムアウト対策）
  const monthHits = plain ? (plain.match(/月分/g) || []).length : 0;
  if (plain && monthHits >= 2 && /円|税込/.test(plain)) return plain;
  if (plain && /お支払い内容|月会費/.test(plain) && monthHits >= 1 && plain.length > 400) {
    return plain;
  }
  const fromHtml = htmlToPlainTextForOp_(message.getBody() || "");
  return pickRicherMailBody_(plain, fromHtml);
}

function htmlToPlainTextForOp_(html) {
  if (!html) return "";
  return String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|li|h\d)>/gi, "\n")
    .replace(/<\/td>/gi, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function pickRicherMailBody_(a, b) {
  function score_(t) {
    const s = String(t || "");
    let n = Math.min(s.length, 20000) / 200;
    if (/月会費|お支払い内容|ご契約中のオプション/.test(s)) n += 50;
    if (/月分/.test(s)) n += 20;
    if (/円|税込|消費税/.test(s)) n += 15;
    if (/VIP|プロテイン|タンニング|ホットスタジオ|レンタル/.test(s)) n += 10;
    return n;
  }
  if (!a) return b || "";
  if (!b) return a;
  return score_(a) >= score_(b) ? a : b;
}

// ── OPメール解析 ──

function extractDataFromThreads(threads, targetYear, targetMonth) {
  const newData = [];
  const seen = {};

  for (let t = 0; t < threads.length; t++) {
    const messages = threads[t].getMessages();
    for (let m = 0; m < messages.length; m++) {
      const message = messages[m];
      const emailDate = message.getDate();
      const subject = String(message.getSubject() || "");
      const body = getMessageBodyText_(message);
      const msgId = message.getId();

      try {
        let parsed = [];
        if (subject.indexOf("オプションご契約につきまして") !== -1) {
          parsed = parseOptionContractEmail(emailDate, body, msgId);
        } else if (subject.indexOf("ご入会ありがとうございます") !== -1) {
          parsed = parseSignupEmail(emailDate, body, msgId);
        }
        for (let i = 0; i < parsed.length; i++) {
          const row = padLogDataRow_(parsed[i], msgId);
          if (!OPTION_LIST.includes(String(row[4] || ""))) continue;
          const d = parseOpLogDate_(row[0]);
          if (targetYear !== null && targetMonth !== null) {
            if (!d || d.getFullYear() !== targetYear || d.getMonth() !== targetMonth) continue;
          }
          const key = opLogDedupeKey_(row);
          if (seen[key]) continue;
          seen[key] = true;
          newData.push(row);
        }
      } catch (e) {
        Logger.log("OP解析エラー: " + subject + " / " + e.message);
      }
    }
  }
  return newData;
}

function padLogDataRow_(row, msgId) {
  const r = row.slice();
  while (r.length < 5) r.push("");
  return [r[0], r[1], r[2], r[3], normalizeOptionName(String(r[4] || r[3] || "")), msgId || r[5] || ""];
}

function parseSignupEmail(date, body, msgId) {
  const name = extractSignupName_(body);
  const businessDate = resolveSignupBusinessDate_(body, date);
  // 1) 「○○(N月分)...円」正規表現（HTMLが1行化しても拾える・本命）
  // 2) 行分割後の内訳
  // 3) キーワード近傍
  const got = {};
  const a = parseSignupFeeItemsRegex_(businessDate, body, name, msgId, got);
  for (let i = 0; i < a.length; i++) got[a[i][4]] = true;
  const b = parseSignupBreakdown_(businessDate, body, name, msgId);
  for (let i = 0; i < b.length; i++) {
    if (got[b[i][4]]) continue;
    got[b[i][4]] = true;
    a.push(b[i]);
  }
  const c = parseSignupLooseLines_(businessDate, body, name, msgId, got);
  for (let j = 0; j < c.length; j++) {
    if (got[c[j][4]]) continue;
    got[c[j][4]] = true;
    a.push(c[j]);
  }
  const d = parseSignupByOptionKeywords_(businessDate, body, name, msgId, got);
  for (let k = 0; k < d.length; k++) {
    if (got[d[k][4]]) continue;
    got[d[k][4]] = true;
    a.push(d[k]);
  }
  return a;
}

/** 料金内訳セクションだけ切り出す（説明文の誤検知を防ぐ） */
function extractFeeSection_(body) {
  const text = String(body || "");
  if (!text) return "";
  const start = text.search(/月会費.*内訳|お支払い内容|ご契約中のオプション|お支払い内容は以下/);
  if (start < 0) {
    // 見出しが無くても「(N月分)」があればその前後を料金帯とみなす
    const m = text.search(/[（(]\s*\d{1,2}\s*月分\s*[）)]/);
    if (m < 0) return text;
    return text.substring(Math.max(0, m - 80), Math.min(text.length, m + 1200));
  }
  const rest = text.substring(start);
  const endRel = rest.search(/APP登録方法|お支払方法|注意事項|ご案内|ご利用規約|クレジットカード|JOYFIT App/);
  return endRel > 0 ? rest.substring(0, endRel) : rest;
}

/**
 * HTML化で料金が1行に潰れていても、
 * 「VIPあんしんサポート(7月分)...750円」形式を直接拾う
 */
function parseSignupFeeItemsRegex_(date, body, name, msgId, alreadyGot) {
  const results = [];
  const got = alreadyGot || {};
  const text = explodePaymentText_(extractFeeSection_(body));
  if (!text) return results;

  // 例: VIPあんしんサポート(7月分)(消費税 10%) 750円(税込825円)
  const re = /([^\n\r]{2,40}?)[（(]\s*(\d{1,2})\s*月分\s*[）)][^\n\r]{0,60}?(\d[\d,]*)\s*円/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const raw = String(m[1] || "").replace(/^[\s\-・･>]+/, "").trim();
    if (!raw || isMembershipFeeRaw_(raw)) continue;
    const norm = normalizeOptionName(raw);
    if (!OPTION_LIST.includes(norm) || got[norm]) continue;
    got[norm] = true;
    results.push([
      date, name, "利用開始(新規入会)",
      (raw + "(" + m[2] + "月分) " + m[3] + "円").replace(/\s+/g, " "),
      norm, msgId
    ]);
  }

  // 月分が無い「オプション名 ... 0円」も料金帯にあれば拾う
  if (results.length === 0 || Object.keys(got).length < 2) {
    const lines = splitPaymentLines_(text);
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (!/円|税込|¥|￥/.test(line)) continue;
      if (isMembershipFeeRaw_(line) || isBreakdownNoiseLine_(line)) continue;
      const norm = normalizeOptionName(line);
      if (!OPTION_LIST.includes(norm) || got[norm]) continue;
      got[norm] = true;
      results.push([date, name, "利用開始(新規入会)", line, norm, msgId]);
    }
  }
  return results;
}

function isMembershipFeeRaw_(raw) {
  const s = String(raw || "");
  if (/法人個人月払|法人会員|ナショナル会員|入会金|事務手数料|月会費|小計|合計/.test(s)) return true;
  if (/法人.*月払/.test(s)) return true;
  return false;
}

/** 1行化した料金表を、項目の頭で改行する */
function explodePaymentText_(text) {
  let s = String(text || "");
  s = s.replace(/[ \t　]+/g, " ");
  s = s.replace(/(-{5,}|={5,}|_{5,})/g, "\n$1\n");
  // 税込○円の直後で改行（次の項目が続くケース）
  s = s.replace(/円(?=\()/g, "円\n");

  const heads = [
    "VIPあんしんサポート", "VIP安心サポート", "あんしんサポートVIP", "安心サポートVIP",
    "水素水＋プロテイン", "プロテイン＋水素水", "水素水+プロテイン", "プロテイン+水素水",
    "オンライン・レッスン", "オンラインレッスン",
    "ボディープランナー", "ボディプランナー", "体組成計",
    "契約ロッカー", "レンタルマット", "マットレンタル",
    "プロテイン無制限", "プロテイン12杯", "プロテイン１２杯",
    "レンタルタオル", "タオルレンタル",
    "タンニング", "セルフエステ", "ホットスタジオ", "HOTスタジオ",
    "ヨガロッカー", "ピラティス",
    "あんしんサポート", "安心サポート",
    "法人個人月払", "ナショナル会員", "入会金", "事務手数料", "水素水"
  ];
  // 長い名称をプレースホルダ化してから改行（短い名称に食い破られない）
  const saved = [];
  heads.sort(function (a, b) { return b.length - a.length; });
  for (let i = 0; i < heads.length; i++) {
    const h = heads[i];
    if (s.indexOf(h) === -1) continue;
    const token = "@@OPHEAD" + saved.length + "@@";
    saved.push(h);
    s = s.split(h).join("\n" + token);
  }
  for (let j = 0; j < saved.length; j++) {
    s = s.split("@@OPHEAD" + j + "@@").join(saved[j]);
  }
  return s;
}

function splitPaymentLines_(text) {
  return explodePaymentText_(text).split(/\r?\n/).map(function (l) {
    return String(l || "").trim();
  }).filter(Boolean);
}

/**
 * 料金表の行が崩れていても、料金セクション内のオプション名＋月分/円 で拾う
 */
function parseSignupByOptionKeywords_(date, body, name, msgId, alreadyGot) {
  const results = [];
  const got = alreadyGot || {};
  const text = explodePaymentText_(extractFeeSection_(body));
  if (!text) return results;

  const aliases = [
    { key: "プロテイン＋水素水", re: /水素水\s*[＋+]\s*プロテイン|プロテイン\s*[＋+]\s*水素水/g },
    { key: "安心サポートVIP", re: /VIP\s*(あんしん|安心)\s*サポート|(あんしん|安心)\s*サポート\s*VIP|VIPあんしんサポート|VIP安心サポート/g },
    { key: "体組成計", re: /ボディ\s*[プ]?ランナー|体組成計/g },
    { key: "レンタルマット", re: /マット\s*レンタル|レンタル\s*マット/g },
    { key: "ピラティスリフォーマー", re: /ピラティス/g },
    { key: "契約ロッカー1,500", re: /契約ロッカー|ロッカー\s*1[,，]?500|ロッカー\s*1500/g },
    { key: "オンラインレッスン", re: /オンライン\s*・?\s*レッスン/g },
    { key: "レンタルタオル", re: /レンタル\s*タオル|タオル\s*レンタル/g },
    { key: "ホットスタジオ", re: /ホット\s*スタジオ|HOT\s*スタジオ/gi },
    { key: "セルフエステ", re: /セルフ\s*エステ/g },
    { key: "タンニング", re: /タンニング/g },
    { key: "水素水", re: /水素水/g },
    { key: "プロテイン無制限", re: /プロテイン\s*(無制限|制限なし)/g },
    { key: "プロテイン12杯", re: /プロテイン\s*(12|１２)\s*杯?/g },
    { key: "ヨガロッカー", re: /ヨガ\s*ロッカー/g },
    { key: "安心サポート", re: /(あんしん|安心)\s*サポート/g }
  ];

  for (let i = 0; i < aliases.length; i++) {
    const item = aliases[i];
    if (got[item.key]) continue;
    if (!OPTION_LIST.includes(item.key)) continue;
    item.re.lastIndex = 0;
    const m = item.re.exec(text);
    if (!m) continue;
    const around = text.substring(Math.max(0, m.index - 40), Math.min(text.length, m.index + 120));
    if (!/月分|円|税込|消費税|¥|￥/.test(around)) continue;
    if (item.key === "水素水" && /プロテイン/.test(around)) continue;
    if (item.key === "安心サポート" && /VIP/i.test(around)) continue;
    got[item.key] = true;
    results.push([date, name, "利用開始(新規入会)", around.replace(/\s+/g, " ").trim(), item.key, msgId]);
  }
  return results;
}

/** 入会・退会と同じく、ご利用開始日があればその日を集計日にする */
function resolveSignupBusinessDate_(body, emailDate) {
  const startMatch = String(body || "").match(/ご利用開始日[の]?\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (startMatch) {
    return new Date(
      parseInt(startMatch[1], 10),
      parseInt(startMatch[2], 10) - 1,
      parseInt(startMatch[3], 10)
    );
  }
  return emailDate;
}

function extractSignupName_(body) {
  const m1 = String(body || "").match(/お名前\s*[:：]\s*(.+?)\s*様/);
  if (m1) return String(m1[1]).replace(/^[>\s]+/, "").trim();
  const lines = String(body || "").split(/\r?\n/);
  for (let i = 0; i < Math.min(lines.length, 20); i++) {
    const line = String(lines[i] || "").trim();
    if (!line || /JOYFIT|受付番号|この度は|お支払い|ご利用開始|ご入会内容|会員情報/.test(line)) continue;
    const m2 = line.match(/^(.{1,30}?)\s*様\s*$/);
    if (m2) return String(m2[1]).replace(/^[>\s]+/, "").trim();
  }
  return "";
}

function isBreakdownNoiseLine_(line) {
  const s = String(line || "").trim();
  if (!s || s.indexOf("----") !== -1) return true;
  if (s.indexOf("小計") !== -1 || s.indexOf("合計") !== -1) return true;
  if (s.indexOf("初期費用") !== -1 || s.indexOf("入会金") !== -1 || s.indexOf("事務手数料") !== -1) return true;
  if (s.indexOf("ナショナル会員") !== -1) return true;
  if (s.indexOf("法人個人月払") !== -1 || s.indexOf("法人会員") !== -1) return true;
  if (/法人.*月払/.test(s)) return true;
  if (s.indexOf("月会費") !== -1 && s.indexOf("内訳") !== -1) return true;
  if (s.indexOf("お支払い内容") !== -1) return true;
  return false;
}

function parseSignupBreakdown_(date, body, name, msgId) {
  const results = [];
  const lines = splitPaymentLines_(extractFeeSection_(body) || String(body || ""));
  let inBreakdown = false;
  // 見出しが無くても月分があれば料金帯として扱う
  if (/月分/.test(lines.join("\n"))) inBreakdown = true;
  const got = {};
  for (let i = 0; i < lines.length; i++) {
    const line = String(lines[i] || "");
    if (/月会費.*内訳/.test(line) || /ご契約中のオプション/.test(line) || /お支払い内容/.test(line)) {
      inBreakdown = true;
      continue;
    }
    if (inBreakdown && /APP登録方法|お支払方法|注意事項|ご案内|ご利用規約|クレジットカード|JOYFIT App/.test(line)) break;
    if (!inBreakdown || isBreakdownNoiseLine_(line) || isMembershipFeeRaw_(line)) continue;

    const norm = normalizeOptionName(line);
    if (!OPTION_LIST.includes(norm) || got[norm]) continue;
    if (line.indexOf("円") === -1 && line.indexOf("¥") === -1 && line.indexOf("￥") === -1 &&
        line.indexOf("月分") === -1) {
      continue;
    }
    got[norm] = true;
    results.push([
      date, name, "利用開始(新規入会)",
      line.replace(/\(\d+月分\).*/, "").replace(/（\d+月分）.*/, "").trim(),
      norm, msgId
    ]);
  }
  return results;
}

function parseSignupLooseLines_(date, body, name, msgId, alreadyGot) {
  const results = [];
  const got = alreadyGot || {};
  const lines = splitPaymentLines_(extractFeeSection_(body) || String(body || ""));
  for (let i = 0; i < lines.length; i++) {
    let line = String(lines[i] || "").trim();
    if (!line || isBreakdownNoiseLine_(line) || isMembershipFeeRaw_(line)) continue;
    const next = String(lines[i + 1] || "").trim();
    const joined = line + " " + next;
    const candidate = (/月分|円|税込|消費税|¥|￥/.test(line) ? line : joined);
    if (candidate.indexOf("月分") === -1 && candidate.indexOf("円") === -1 &&
        candidate.indexOf("税込") === -1 && candidate.indexOf("消費税") === -1 &&
        candidate.indexOf("¥") === -1 && candidate.indexOf("￥") === -1) {
      continue;
    }
    const norm = normalizeOptionName(candidate);
    if (!OPTION_LIST.includes(norm) || got[norm]) continue;
    got[norm] = true;
    results.push([date, name, "利用開始(新規入会)", candidate, norm, msgId]);
  }
  return results;
}

function parseOptionContractEmail(date, body, msgId) {
  const results = [];
  const nameMatch = String(body || "").match(/([^\r\n]{1,40}?)\s*様/);
  const name = nameMatch ? String(nameMatch[1]).replace(/^[>\s]+/, "").trim() : "";
  const re = /[（(【\[](利用開始|利用停止)[）)】\]]\s*([^\r\n]+)/g;
  let m;
  while ((m = re.exec(String(body || ""))) !== null) {
    const status = m[1] === "利用開始" ? "利用開始(OP追加)" : "利用停止";
    const raw = String(m[2]).trim();
    if (!raw || isJunkOptionRaw_(raw)) continue;
    const norm = normalizeOptionName(raw);
    // OPTION_LIST に無い行は捨てる（「について」「の場合でも合計」などの誤検知防止）
    if (!OPTION_LIST.includes(norm)) continue;
    const businessDate = resolveOptionLineBusinessDate_(raw, date);
    results.push([businessDate, name, status, raw, norm, msgId]);
  }
  return results;
}

function isJunkOptionRaw_(raw) {
  const s = String(raw || "").trim();
  if (!s || s.length < 2) return true;
  // オプション名ではありえない文言だけ弾く（税込○円の正規行は残す）
  if (/について|場合でも|^合計|^小計|^の場合|^円/.test(s)) return true;
  return false;
}

/** 行に「8月分」などがあればその月を集計月にする */
function resolveOptionLineBusinessDate_(raw, emailDate) {
  const m = String(raw || "").match(/[（(]\s*(\d{1,2})\s*月分\s*[）)]/);
  if (!m) return emailDate;
  let month = parseInt(m[1], 10) - 1;
  if (month < 0 || month > 11) return emailDate;
  let year = emailDate.getFullYear();
  const em = emailDate.getMonth();
  if (em <= 1 && month >= 10) year -= 1;
  if (em >= 10 && month <= 1) year += 1;
  return new Date(year, month, Math.min(emailDate.getDate(), 28));
}

function normalizeOptionName(rawName) {
  const s = String(rawName || "");
  if (s.includes("水素水") && s.includes("プロテイン")) return "プロテイン＋水素水";
  if (s.includes("VIP") && (s.includes("あんしん") || s.includes("安心"))) return "安心サポートVIP";
  if (s.includes("ボディプランナー") || s.includes("ボディープランナー") || s.includes("体組成")) {
    return "体組成計";
  }
  if (s.includes("マットレンタル") || s.includes("レンタルマット")) return "レンタルマット";
  if (s.includes("ピラティス")) return "ピラティスリフォーマー";
  if (s.includes("ロッカー") && /(1[,，]?500|１[,，]?５００|1500)/.test(s)) return "契約ロッカー1,500";
  if (s.includes("ヨガ") && s.includes("ロッカー")) return "ヨガロッカー";
  if (s.includes("レンタルタオル") || (s.includes("タオル") && s.includes("レンタル"))) return "レンタルタオル";
  if (s.includes("ホットスタジオ") || s.includes("HOTスタジオ")) return "ホットスタジオ";
  if (s.includes("オンラインレッスン") || s.includes("オンライン・レッスン")) return "オンラインレッスン";
  if (s.includes("セルフエステ")) return "セルフエステ";
  if (s.includes("タンニング")) return "タンニング";
  if (s.includes("プロテイン") && (s.includes("無制限") || s.includes("制限なし"))) return "プロテイン無制限";
  if (s.includes("プロテイン") && (s.includes("12") || s.includes("１２"))) return "プロテイン12杯";
  if (s.includes("安心サポート") || s.includes("あんしんサポート")) return "安心サポート";
  if (s.includes("水素水")) return "水素水";

  const byLen = OPTION_LIST.slice().sort(function (a, b) { return b.length - a.length; });
  for (let i = 0; i < byLen.length; i++) {
    if (s.indexOf(byLen[i]) !== -1) return byLen[i];
  }
  return s;
}

// ── 集計・日報 ──

function parseOpLogDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number" && v > 30000) {
    const base = new Date(1899, 11, 30);
    const d = new Date(base.getTime() + Math.floor(v) * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(v || "").trim();
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function resolveNippoTargetYearMonth_(ss) {
  const nippo = ss.getSheetByName("日報");
  if (nippo) {
    const b1 = nippo.getRange("B1").getDisplayValue() || nippo.getRange("B1").getValue();
    const parsed = parseYearMonthFlexible_(b1);
    if (parsed) return parsed;
  }
  return resolveSummaryTargetYearMonth_(ss);
}

/** 日報B1 / 表示文字列から年月を取る（2608 / 2026年8月 / 2026/08 など） */
function parseYearMonthFlexible_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return {
      year: raw.getFullYear(),
      month: raw.getMonth(),
      label: raw.getFullYear() + "年" + (raw.getMonth() + 1) + "月"
    };
  }
  const s = String(raw || "").trim();
  if (!s) return null;

  let m = s.match(/^(\d{2})(\d{2})$/);
  if (m) {
    const y = 2000 + parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    if (mo >= 0 && mo <= 11) return { year: y, month: mo, label: y + "年" + (mo + 1) + "月" };
  }
  m = s.match(/^(\d{4})年(\d{1,2})月$/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    if (mo >= 0 && mo <= 11) return { year: y, month: mo, label: s };
  }
  m = s.match(/^(\d{4})[\/\-.](\d{1,2})$/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    if (mo >= 0 && mo <= 11) return { year: y, month: mo, label: y + "年" + (mo + 1) + "月" };
  }
  return null;
}

function resolveSummaryTargetYearMonth_(ss) {
  const summary = getOpSheet_(ss);
  if (summary) {
    const s = summary.getRange("B1").getDisplayValue() || summary.getRange("B1").getValue();
    const parsed = parseYearMonthFlexible_(s);
    if (parsed) return parsed;
    const text = String(s || "").trim();
    const ym = text.match(/^(\d{4})年(\d{1,2})月$/);
    if (ym) {
      return { year: parseInt(ym[1], 10), month: parseInt(ym[2], 10) - 1, label: text };
    }
  }
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth(), label: now.getFullYear() + "年" + (now.getMonth() + 1) + "月" };
}

function countOpDataForMonth_(logSheet, targetYear, targetMonth) {
  const data = readOpLogValues_(logSheet);
  const counts = OPTION_LIST.map(function () { return { newSignup: 0, opAdd: 0, stop: 0 }; });
  const idxByOpt = {};
  for (let o = 0; o < OPTION_LIST.length; o++) idxByOpt[OPTION_LIST[o]] = o;

  const seen = {};
  for (let i = 1; i < data.length; i++) {
    const d = parseOpLogDate_(data[i][0]);
    if (!d || d.getFullYear() !== targetYear || d.getMonth() !== targetMonth) continue;
    const key = opLogDedupeKey_(data[i]);
    if (seen[key]) continue;
    seen[key] = true;
    const cat = String(data[i][2] || "").trim();
    const opt = String(data[i][4] || "").trim();
    const idx = idxByOpt[opt];
    if (idx === undefined) continue;
    if (cat.indexOf("利用停止") !== -1) {
      counts[idx].stop++;
    } else if (cat.indexOf("新規入会") !== -1) {
      counts[idx].newSignup++;
    } else if (cat.indexOf("利用開始") !== -1 || cat.indexOf("OP追加") !== -1) {
      counts[idx].opAdd++;
    }
  }
  return { counts: counts };
}

function syncSummarySheetFromOpData_(ss, targetYear, targetMonth, logSheet) {
  const summarySheet = logSheet || getOpSheet_(ss);
  if (!summarySheet) return;
  const result = countOpDataForMonth_(summarySheet, targetYear, targetMonth);
  const bCol = [], cCol = [], dCol = [], eCol = [];
  for (let i = 0; i < OPTION_LIST.length; i++) {
    const start = result.counts[i].newSignup + result.counts[i].opAdd;
    bCol.push([result.counts[i].newSignup]);
    cCol.push([result.counts[i].opAdd]);
    dCol.push([start]);
    eCol.push([result.counts[i].stop]);
  }
  summarySheet.getRange(3, 2, OPTION_LIST.length, 1).setValues(bCol);
  summarySheet.getRange(3, 3, OPTION_LIST.length, 1).setValues(cCol);
  summarySheet.getRange(3, 4, OPTION_LIST.length, 1).setValues(dCol);
  summarySheet.getRange(3, 5, OPTION_LIST.length, 1).setValues(eCol);
  ensureSummaryFormulas_(summarySheet);
}

function ensureSummaryOpeningColumn_(summarySheet) {
  summarySheet.getRange(2, SUMMARY_COL_OPENING_MAIL).setValue("月初(日報)");
  summarySheet.getRange(2, SUMMARY_COL_OPENING_MAIL).setFontWeight("bold");
  summarySheet.getRange(2, SUMMARY_COL_OPENING_MAIL).setBackground("#fff2cc");
}

function writeSummaryOpeningColumn_(summarySheet, values) {
  ensureSummaryOpeningColumn_(summarySheet);
  const gCol = values.map(function (v) { return [v]; });
  summarySheet.getRange(3, SUMMARY_COL_OPENING_MAIL, OPTION_LIST.length, 1).setValues(gCol);
  summarySheet.getRange(3, SUMMARY_COL_OPENING_MAIL, OPTION_LIST.length, 1)
    .setHorizontalAlignment("center").setNumberFormat("0");
}

function ensureSummaryFormulas_(summarySheet) {
  const monthText = String(summarySheet.getRange("B1").getDisplayValue() || summarySheet.getRange("B1").getValue() || "").trim();
  const match = monthText.match(/^(\d{4})年(\d{1,2})月$/);
  let currentLabel = "月初";
  let nextLabel = "翌月月初";
  if (match) {
    currentLabel = parseInt(match[2], 10) + "月月初";
    const next = new Date(parseInt(match[1], 10), parseInt(match[2], 10), 1);
    nextLabel = (next.getMonth() + 1) + "月月初";
  }
  summarySheet.getRange(2, SUMMARY_COL_OPENING_MAIL).setValue(currentLabel);
  summarySheet.getRange(2, SUMMARY_COL_NEXT_OPENING).setValue(nextLabel);
  for (let i = 0; i < OPTION_LIST.length; i++) {
    const row = i + 3;
    summarySheet.getRange(row, 6).setFormula("=D" + row + "-E" + row);
    summarySheet.getRange(row, SUMMARY_COL_NEXT_OPENING)
      .setFormula("=IF(G" + row + "=\"\",\"\",G" + row + "+F" + row + ")");
  }
}

// ── 日報オプション数値（B=名, C=月初, D=契約, E=解約, F=増減）──
// 見た目・結合・見出しは手作業。GASの更新は D/E の数値と F の式だけ触る。

const NIPPO_OP_HEADER_ROW = 20;
const NIPPO_OP_START_ROW = 21;
const NIPPO_COL_LABEL = 2;       // B … オプション名
const NIPPO_COL_OPENING = 3;     // C … 月初（手入力・更新では触らない）
const NIPPO_COL_CURRENT = 4;     // D … 契約（GASが書く）
const NIPPO_COL_STOP = 5;        // E … 解約（GASが書く）
const NIPPO_COL_NET = 6;         // F … 増減（式 =D-E）
const NIPPO_MEMBER_OPENING_CELL = "C12"; // 月初会員数
const NIPPO_OP_FONT = "Meiryo";
const NIPPO_COLOR_INK = "#111111";
const NIPPO_COLOR_WHITE = "#FFFFFF";
const NIPPO_COLOR_GRAY = "#F2F2F2";
/** 互換定数（旧呼び出し向け・見た目更新では使わない） */
const NIPPO_OP_OPENING_BG = NIPPO_COLOR_GRAY;
const NIPPO_OP_STOP_BG = NIPPO_COLOR_WHITE;
const NIPPO_OP_NET_BG = NIPPO_COLOR_WHITE;
const NIPPO_OP_HEADER_BG = NIPPO_COLOR_INK;
const NIPPO_OP_HEADER_FG = NIPPO_COLOR_WHITE;
const NIPPO_OP_ACCENT = NIPPO_COLOR_INK;
const NIPPO_STYLE_AREA_ROW = 5;
const NIPPO_STYLE_AREA_ROWS = 32;
const NIPPO_STYLE_AREA_COL = 1;
const NIPPO_STYLE_AREA_COLS = 9;
const NIPPO_OP_STYLE_PROP = "nippoSheetStyleVersion";
const NIPPO_OP_STYLE_VERSION = "v5-numbers-only";

/** 月次シート（2606など）のオプション欄 … 日報とは別 */
const MONTHLY_SHEET_OP_START = 21;
const MONTHLY_SHEET_COL_CURRENT = 3;   // C 当月
const MONTHLY_SHEET_COL_TOTAL = 5;     // E 合計
const MONTHLY_SHEET_COL_OPENING = 7;   // G 月初
const MONTHLY_SHEET_COL_NET = 8;       // H 純増減

function updateNippoSheetForMonth(ss, targetYear, targetMonth, logSheet) {
  const nippoSheet = ss.getSheetByName("日報");
  if (!nippoSheet) return;

  const anchor = resolveNippoOpAnchor_(nippoSheet);
  ensureNippoOpFormulasLight_(nippoSheet, anchor);
  ensureNippoIriaiBreakdownFormulas_(nippoSheet);

  const result = countOpDataForMonth_(logSheet, targetYear, targetMonth);
  const n = OPTION_LIST.length;
  const labels = nippoSheet
    .getRange(anchor.startRow, NIPPO_COL_LABEL, n, 1)
    .getDisplayValues();

  const byOpt = {};
  for (let i = 0; i < n; i++) {
    byOpt[OPTION_LIST[i]] = result.counts[i];
  }

  const currentCol = [];
  const stopCol = [];
  for (let i = 0; i < n; i++) {
    const counts = resolveOpCountsForNippoLabel_(labels[i][0], byOpt, result.counts[i]);
    currentCol.push([counts.newSignup + counts.opAdd]);
    stopCol.push([counts.stop]);
  }

  nippoSheet.getRange(anchor.startRow, NIPPO_COL_CURRENT, n, 1).setValues(currentCol);
  nippoSheet.getRange(anchor.startRow, NIPPO_COL_STOP, n, 1).setValues(stopCol);
  Logger.log(
    "日報OP反映: " + targetYear + "/" + (targetMonth + 1) +
      " headerRow=" + anchor.headerRow + " startRow=" + anchor.startRow +
      " 契約合計=" + sumOpCountField_(result.counts, "contract") +
      " 解約合計=" + sumOpCountField_(result.counts, "stop")
  );
}

/** 日報の「月初/契約/解約」見出し行を探して、データ開始行を決める */
function resolveNippoOpAnchor_(nippoSheet) {
  const fallback = { headerRow: NIPPO_OP_HEADER_ROW, startRow: NIPPO_OP_START_ROW };
  if (!nippoSheet) return fallback;

  const maxScan = Math.min(100, nippoSheet.getMaxRows());
  // B〜F を走査（numRows, numColumns）
  const block = nippoSheet.getRange(1, NIPPO_COL_LABEL, maxScan, 5).getDisplayValues();
  for (let r = 0; r < block.length; r++) {
    const c = String(block[r][1] || "").replace(/\s/g, "");
    const d = String(block[r][2] || "").replace(/\s/g, "");
    const e = String(block[r][3] || "").replace(/\s/g, "");
    const cOk = c.indexOf("月初") !== -1;
    const dOk = d.indexOf("契約") !== -1;
    const eOk = e.indexOf("解約") !== -1;
    if (cOk && dOk && eOk) {
      return { headerRow: r + 1, startRow: r + 2 };
    }
  }
  return fallback;
}

function resolveOpCountsForNippoLabel_(label, byOpt, fallbackCounts) {
  const empty = { newSignup: 0, opAdd: 0, stop: 0 };
  const raw = String(label || "").trim();
  if (!raw) return fallbackCounts || empty;
  if (byOpt[raw]) return byOpt[raw];
  if (typeof normalizeOptionName === "function") {
    const norm = normalizeOptionName(raw);
    if (byOpt[norm]) return byOpt[norm];
  }
  for (let i = 0; i < OPTION_LIST.length; i++) {
    if (raw.indexOf(OPTION_LIST[i]) !== -1 || OPTION_LIST[i].indexOf(raw) !== -1) {
      return byOpt[OPTION_LIST[i]] || empty;
    }
  }
  return fallbackCounts || empty;
}

function sumOpCountField_(counts, field) {
  let sum = 0;
  for (let i = 0; i < counts.length; i++) {
    if (field === "contract") sum += counts[i].newSignup + counts[i].opAdd;
    else if (field === "newSignup") sum += counts[i].newSignup;
    else if (field === "opAdd") sum += counts[i].opAdd;
    else if (field === "stop") sum += counts[i].stop;
  }
  return sum;
}

/**
 * 日報の見た目は手作業。更新処理からは呼ばない。
 * （互換のため残置・中身は何もしない）
 */
function ensureNippoOpLayoutOnce_(nippoSheet) {
  // no-op … デザインはシート側で維持
}

/** @deprecated 互換 … 見た目は触らない */
function setupNippoOpLayout_(nippoSheet) {
  // no-op
}

function formatNippoOpArea_(ss) {
  // no-op … 更新は数値のみ（updateNippoSheetForMonth）
}

/** 手動メンテ用: オプション表を作り直す（普段の更新では使わない） */
function rebuildNippoOpTable_(nippoSheet) {
  if (!nippoSheet) return;
  try {
    nippoSheet.getRange(NIPPO_OP_HEADER_ROW, NIPPO_COL_LABEL).breakApart();
  } catch (e) { /* 未結合なら無視 */ }
  const lastRow = NIPPO_OP_START_ROW + OPTION_LIST.length - 1;
  const wipe = nippoSheet.getRange(
    NIPPO_OP_HEADER_ROW, NIPPO_COL_LABEL, lastRow, NIPPO_COL_NET
  );
  wipe.clearContent();
  writeNippoOptionLabels_(nippoSheet);
  nippoSheet.getRange(NIPPO_OP_HEADER_ROW, NIPPO_COL_LABEL, 1, 5)
    .setValues([["オプション", "月初", "契約", "解約", "増減"]]);
  ensureNippoOpFormulasLight_(nippoSheet);
}

function isNippoOpNewHeader_(nippoSheet) {
  const hr = NIPPO_OP_HEADER_ROW;
  const c = String(nippoSheet.getRange(hr, NIPPO_COL_OPENING).getDisplayValue() || "").trim();
  const d = String(nippoSheet.getRange(hr, NIPPO_COL_CURRENT).getDisplayValue() || "").trim();
  const e = String(nippoSheet.getRange(hr, NIPPO_COL_STOP).getDisplayValue() || "").trim();
  const f = String(nippoSheet.getRange(hr, NIPPO_COL_NET).getDisplayValue() || "").trim();
  return c === "月初" && (d === "契約" || d === "契約数") &&
    (e === "解約" || e === "解約数") && (f === "増減" || f === "純増減");
}

function writeNippoOptionLabels_(nippoSheet) {
  const n = OPTION_LIST.length;
  const labels = OPTION_LIST.map(function (name) { return [name]; });
  const labelRange = nippoSheet.getRange(NIPPO_OP_START_ROW, NIPPO_COL_LABEL, n, 1);
  labelRange.setValues(labels);
}

/**
 * 日報 A5:I36 全体の見た目（初回のみ）
 * 色は黒 / 白 / 薄グレーのみ。枠線に色は付けない（黒のみ）。
 */
function applyNippoSheetStyleOnce_(nippoSheet) {
  if (!nippoSheet) return;
  const n = OPTION_LIST.length;
  const hr = NIPPO_OP_HEADER_ROW;
  const lastOpRow = NIPPO_OP_START_ROW + n - 1;

  const area = nippoSheet.getRange(
    NIPPO_STYLE_AREA_ROW, NIPPO_STYLE_AREA_COL,
    NIPPO_STYLE_AREA_ROWS, NIPPO_STYLE_AREA_COLS
  );

  // 全体の文字・余白をそろえる
  area.setFontFamily(NIPPO_OP_FONT);
  area.setFontSize(10);
  area.setFontColor(NIPPO_COLOR_INK);
  area.setFontWeight("normal");
  area.setVerticalAlignment("middle");
  area.setBackground(NIPPO_COLOR_WHITE);
  // 枠線の色付けなし（クリア）… メール貼付時に色枠が邪魔にならないように
  area.setBorder(false, false, false, false, false, false);

  for (let r = NIPPO_STYLE_AREA_ROW; r < NIPPO_STYLE_AREA_ROW + NIPPO_STYLE_AREA_ROWS; r++) {
    nippoSheet.setRowHeight(r, 22);
  }

  // 冒頭メッセージ（5〜7行）は太字
  nippoSheet.getRange(5, 2, 3, 8).setFontWeight("bold");
  nippoSheet.getRange(5, 2, 3, 8).setHorizontalAlignment("left");

  // 会員サマリー（9〜18行）ラベル左寄せ
  nippoSheet.getRange(9, 2, 10, 1).setHorizontalAlignment("left");
  nippoSheet.getRange(9, 2, 10, 1).setFontWeight("bold");

  // オプション見出し
  const headers = [["オプション", "月初", "契約", "解約", "増減"]];
  nippoSheet.getRange(hr, NIPPO_COL_LABEL, 1, 5).setValues(headers);
  const headerRange = nippoSheet.getRange(hr, NIPPO_COL_LABEL, 1, 5);
  headerRange.setFontFamily(NIPPO_OP_FONT);
  headerRange.setFontSize(10);
  headerRange.setFontWeight("bold");
  headerRange.setFontColor(NIPPO_COLOR_WHITE);
  headerRange.setBackground(NIPPO_COLOR_INK);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");

  // オプション表 … 黒枠のみ（色付き枠は使わない）
  const opTable = nippoSheet.getRange(hr, NIPPO_COL_LABEL, lastOpRow - hr + 1, 5);
  opTable.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // 名称列
  const labelRange = nippoSheet.getRange(NIPPO_OP_START_ROW, NIPPO_COL_LABEL, n, 1);
  labelRange.setFontWeight("bold");
  labelRange.setHorizontalAlignment("left");
  labelRange.setWrap(false);
  labelRange.setBackground(NIPPO_COLOR_WHITE);

  // 数値列
  const numRange = nippoSheet.getRange(NIPPO_OP_START_ROW, NIPPO_COL_OPENING, n, 4);
  numRange.setHorizontalAlignment("center");
  numRange.setNumberFormat("0;-0;0");
  numRange.setBackground(NIPPO_COLOR_WHITE);

  // 月初列だけ薄いグレー（手入力の目印・3色以内）
  nippoSheet.getRange(NIPPO_OP_START_ROW, NIPPO_COL_OPENING, n, 1).setBackground(NIPPO_COLOR_GRAY);

  nippoSheet.setRowHeight(hr, 24);
}

/** 互換名 */
function applyNippoOpStyleOnce_(nippoSheet) {
  applyNippoSheetStyleOnce_(nippoSheet);
}

/** 互換名（中身は初回スタイルへ） */
function ensureNippoOpHeaders_(nippoSheet) {
  applyNippoSheetStyleOnce_(nippoSheet);
}

function ensureNippoOpFormulas_(nippoSheet) {
  ensureNippoOpFormulasLight_(nippoSheet);
}

/** 増減式だけ軽く確認・セット（既にあれば触らない） */
/** 日報の移籍・復会・紹介は 2609 の H/I/J/K を参照する（手入力0のままにしない） */
function ensureNippoIriaiBreakdownFormulas_(nippoSheet) {
  if (!nippoSheet) return;
  const dayI = '=IFERROR(N(INDEX(INDIRECT("\'"&$B$1&"\'!$I$5:$I$35"),DAY(TODAY())))+N(INDEX(INDIRECT("\'"&$B$1&"\'!$H$5:$H$35"),DAY(TODAY()))),0)';
  const dayJ = '=IFERROR(N(INDEX(INDIRECT("\'"&$B$1&"\'!$J$5:$J$35"),DAY(TODAY()))),0)';
  const dayK = '=IFERROR(N(INDEX(INDIRECT("\'"&$B$1&"\'!$K$5:$K$35"),DAY(TODAY()))),0)';
  const monthHI = '=IFERROR(N(INDIRECT($B$1&"!H36"))+N(INDIRECT($B$1&"!I36")),0)';
  const monthJ = '=IFERROR(N(INDIRECT($B$1&"!J36")),0)';
  const monthK = '=IFERROR(N(INDIRECT($B$1&"!K36")),0)';
  nippoSheet.getRange("D10").setFormula(dayI);
  nippoSheet.getRange("F10").setFormula(dayJ);
  nippoSheet.getRange("H10").setFormula(dayK);
  nippoSheet.getRange("D14").setFormula(monthHI);
  nippoSheet.getRange("F14").setFormula(monthJ);
  nippoSheet.getRange("H14").setFormula(monthK);
}

function ensureNippoOpFormulasLight_(nippoSheet, anchorOpt) {
  const anchor = anchorOpt || resolveNippoOpAnchor_(nippoSheet);
  const sample = String(nippoSheet.getRange(anchor.startRow, NIPPO_COL_NET).getFormula() || "");
  const startRow = anchor.startRow;
  if (new RegExp("^=D" + startRow + "\\s*-\\s*E" + startRow + "$", "i").test(sample)) return;
  const n = OPTION_LIST.length;
  const formulas = [];
  for (let i = 0; i < n; i++) {
    const r = startRow + i;
    formulas.push(["=D" + r + "-E" + r]);
  }
  nippoSheet.getRange(startRow, NIPPO_COL_NET, n, 1).setFormulas(formulas);
}

/**
 * 旧並び（C当月/D月初/E純増減/F解約）→ 新並び（C月初/D契約/E解約/F増減）へ1回だけ移す
 */
function migrateNippoOpColumnOrderIfNeeded_(nippoSheet) {
  if (!nippoSheet) return;
  const cHeader = String(nippoSheet.getRange(NIPPO_OP_HEADER_ROW, 3).getDisplayValue() ||
    nippoSheet.getRange(NIPPO_OP_HEADER_ROW, 3).getValue() || "").trim();
  if (cHeader !== "当月") return;

  const n = OPTION_LIST.length;
  const oldCurrent = nippoSheet.getRange(NIPPO_OP_START_ROW, 3, n, 1).getValues();
  const oldOpening = nippoSheet.getRange(NIPPO_OP_START_ROW, 4, n, 1).getValues();
  const oldStop = nippoSheet.getRange(NIPPO_OP_START_ROW, 6, n, 1).getValues();
  nippoSheet.getRange(NIPPO_OP_START_ROW, 3, n, 1).setValues(oldOpening);
  nippoSheet.getRange(NIPPO_OP_START_ROW, 4, n, 1).setValues(oldCurrent);
  nippoSheet.getRange(NIPPO_OP_START_ROW, 5, n, 1).setValues(oldStop);
}

/** 過去に非表示にした G・H 列を再表示 */
function ensureNippoOpColumnsVisible_(nippoSheet) {
  try {
    nippoSheet.showColumns(7, 2);
  } catch (e) {
    Logger.log("日報 G・H 列の表示: " + e.message);
  }
}

function formatNippoOpDataRange_(range) {
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
  range.setFontFamily(NIPPO_OP_FONT);
  range.setFontSize(10);
  range.setNumberFormat("0;-0;0");
}

function formatNippoOpFormulaRange_(range) {
  formatNippoOpDataRange_(range);
  range.setFontWeight("normal");
}

/** @deprecated 互換 … 初回整備のみ */
function applyNippoOpFormatting_(nippoSheet) {
  applyNippoSheetStyleOnce_(nippoSheet);
}

function resetNippoOpProgressColumns_(nippoSheet) {
  if (!nippoSheet) return;
  const n = OPTION_LIST.length;
  const zeros = [];
  for (let i = 0; i < n; i++) zeros.push([0]);
  const cRange = nippoSheet.getRange(NIPPO_OP_START_ROW, NIPPO_COL_CURRENT, n, 1);
  const stopRange = nippoSheet.getRange(NIPPO_OP_START_ROW, NIPPO_COL_STOP, n, 1);
  cRange.setValues(zeros);
  stopRange.setValues(zeros);
  formatNippoOpDataRange_(cRange);
  formatNippoOpDataRange_(stopRange);
  ensureNippoOpFormulas_(nippoSheet);
}

// ── 店舗日報メールから月初を取得 ──

function formatNippoMailDate_(year, month, day) {
  return year + "." + String(month + 1).padStart(2, "0") + "." + String(day).padStart(2, "0");
}

function buildNippoMailSearchQuery_(dateStr) {
  return "from:" + NIPPO_MAIL_FROM +
    ' subject:"JF24経堂" subject:"' + NIPPO_MAIL_SUBJECT_TAG + '" subject:"' + dateStr + '"';
}

function findDailyReportMailForDate_(year, month, day) {
  const dateStr = formatNippoMailDate_(year, month, day);
  const threads = GmailApp.search(buildNippoMailSearchQuery_(dateStr), 0, 3);
  if (!threads.length) return null;
  let newest = null;
  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      if (String(message.getSubject() || "").indexOf(dateStr) < 0) return;
      if (!newest || message.getDate().getTime() > newest.getDate().getTime()) {
        newest = message;
      }
    });
  });
  return newest;
}

/** 日報メール本文から月初会員数を抽出（翌月月初会員数とは別） */
function parseNippoMailMemberOpening_(body, mode) {
  const text = stripHtmlToText_(body);
  if (mode === "nextMonth") {
    const m = text.match(/翌月月初会員数[^\d\-]*(\d+)/);
    return m ? parseInt(m[1], 10) : null;
  }
  const lines = text.split(/\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.indexOf("翌月月初会員数") >= 0) continue;
    const m = line.match(/月初会員数[^\d\-]*(\d+)/);
    if (m) return parseInt(m[1], 10);
  }
  return null;
}

/**
 * 当月の月初会員数 … 原則その月1日の日報メール「月初会員数」
 * 1日メールが無い場合は前月末メールの「翌月月初会員数」
 */
function fetchMemberOpeningFromDailyReportMail_(year, month, options) {
  options = options || {};
  const monthLabel = year + "年" + (month + 1) + "月";

  const msg1 = findDailyReportMailForDate_(year, month, 1);
  if (msg1) {
    const value = parseNippoMailMemberOpening_(msg1.getBody(), "opening");
    if (value !== null && (value > 0 || options.allowZero)) {
      return {
        value: value,
        source: "日報メール " + msg1.getSubject() + "（月初会員数）",
        monthLabel: monthLabel,
      };
    }
  }

  const prev = shiftYearMonth_(year, month, -1);
  const lastDay = new Date(year, month, 0).getDate();
  const msgEnd = findDailyReportMailForDate_(prev.year, prev.month, lastDay);
  if (msgEnd) {
    const value = parseNippoMailMemberOpening_(msgEnd.getBody(), "nextMonth");
    if (value !== null && (value > 0 || options.allowZero)) {
      return {
        value: value,
        source: "日報メール " + msgEnd.getSubject() + "（翌月月初会員数）",
        monthLabel: monthLabel,
      };
    }
  }

  return {
    value: null,
    source: "日報メール未検出（月初会員数 " + formatNippoMailDate_(year, month, 1) + "）",
    monthLabel: monthLabel,
  };
}

function writeNippoMemberOpening_(nippo, value) {
  if (value === null || value === undefined) return;
  const cell = nippo.getRange(NIPPO_MEMBER_OPENING_CELL);
  cell.setValue(value);
  cell.setNumberFormat("0");
  cell.setHorizontalAlignment("center");
}

function stripHtmlToText_(html) {
  return String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<\/td>/gi, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/g, "&")
    .replace(/\u3000/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .trim();
}

function normalizeNippoOptionLabel_(text) {
  return String(text || "").replace(/^[\s■]+/, "").replace(/\s+/g, "").trim();
}

function matchOptionNameFromMailLabel_(label) {
  const norm = normalizeNippoOptionLabel_(label);
  if (!norm) return null;
  if (norm.includes("水素水") && norm.includes("プロテイン")) return "プロテイン＋水素水";
  if (norm.includes("VIP") && (norm.includes("安心") || norm.includes("あんしん"))) return "安心サポートVIP";
  if ((norm.includes("安心サポート") || norm.includes("あんしんサポート")) && !norm.includes("VIP")) {
    return "安心サポート";
  }
  const ordered = OPTION_LIST.slice().sort(function (a, b) {
    return b.replace(/\s+/g, "").length - a.replace(/\s+/g, "").length;
  });
  for (let i = 0; i < ordered.length; i++) {
    const opt = ordered[i];
    const optNorm = opt.replace(/\s+/g, "");
    // メール側の名称が長い／完全一致のみ（短い名称が長いOP名に吸われない）
    if (norm === optNorm || norm.indexOf(optNorm) >= 0) {
      return opt;
    }
  }
  return null;
}

/**
 * 日報メールHTMLからオプション表を解析
 * 戻り値: { オプション名: { current, total, opening } }
 */
function parseNippoMailOptionsTable_(html) {
  const parsed = {};
  OPTION_LIST.forEach(function (name) { parsed[name] = null; });

  const rows = String(html || "").match(/<tr[\s\S]*?<\/tr>/gi) || [];
  rows.forEach(function (tr) {
    const cells = [];
    const tdRe = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    let m;
    while ((m = tdRe.exec(tr)) !== null) {
      cells.push(stripHtmlToText_(m[1]).replace(/\s+/g, " ").trim());
    }
    if (cells.length < 2) return;

    let optName = null;
    let numStart = -1;
    for (let i = 0; i < cells.length; i++) {
      optName = matchOptionNameFromMailLabel_(cells[i]);
      if (optName) {
        numStart = i + 1;
        break;
      }
    }
    if (!optName || numStart < 0) return;

    const nums = [];
    for (let j = numStart; j < cells.length && nums.length < 4; j++) {
      const n = parseMonthlyNumeric_(cells[j]);
      if (n !== null) nums.push(n);
    }
    assignParsedNippoOptionRow_(parsed, optName, nums);
  });

  if (!OPTION_LIST.some(function (n) { return parsed[n] !== null; })) {
    return parseNippoMailOptionsFromPlain_(stripHtmlToText_(html));
  }
  return parsed;
}

/** 日報メール表 … オプション名の次は「当月・月初・純増減・解約」の順 */
function assignParsedNippoOptionRow_(parsed, optName, nums) {
  if (!nums.length) return;
  if (nums.length >= 4) {
    parsed[optName] = {
      current: nums[0],
      opening: nums[1],
      net: nums[2],
      stop: nums[3],
      total: nums[1] + nums[2],
    };
  } else if (nums.length === 3) {
    parsed[optName] = {
      current: nums[0],
      opening: nums[1],
      net: nums[2],
      total: nums[1] + nums[2],
    };
  } else if (nums.length === 2) {
    parsed[optName] = { current: nums[0], opening: nums[1], total: nums[1] };
  } else {
    parsed[optName] = { current: nums[0], opening: nums[0], total: nums[0] };
  }
}

function parseNippoMailOptionsFromPlain_(text) {
  const parsed = {};
  OPTION_LIST.forEach(function (name) { parsed[name] = null; });
  const lines = String(text || "").split(/\n/);
  lines.forEach(function (line) {
    const optName = matchOptionNameFromMailLabel_(line);
    if (!optName) return;
    const nums = line.match(/-?\d+/g);
    if (!nums || nums.length < 1) return;
    const n = nums.map(function (s) { return parseInt(s, 10); });
    assignParsedNippoOptionRow_(parsed, optName, n.length >= 4 ? n.slice(-4) : n);
  });
  return parsed;
}

function parsedMailToOpeningValues_(parsed, useColumn) {
  return OPTION_LIST.map(function (name) {
    const row = parsed[name];
    if (!row) return 0;
    if (useColumn === "total") return row.total;
    return row.opening;
  });
}

/** 月初メールの表が十分に読めたかを確認（解析失敗による全0上書きを防ぐ） */
function validateOpeningMailParse_(parsed, useColumn) {
  let matched = 0;
  let numeric = 0;
  let positive = 0;
  OPTION_LIST.forEach(function (name) {
    const row = parsed[name];
    if (!row) return;
    matched++;
    const value = useColumn === "total" ? row.total : row.opening;
    if (typeof value === "number" && !isNaN(value)) {
      numeric++;
      if (value > 0) positive++;
    }
  });
  const minimumRows = Math.ceil(OPTION_LIST.length * 0.75);
  return {
    valid: matched >= minimumRows && numeric >= minimumRows && positive > 0,
    matched: matched,
    expected: OPTION_LIST.length,
  };
}

/**
 * 当月の月初 … 原則その月1日の店舗日報メールの「月初」列
 * 1日メールが無い場合は前月末メールの「合計」列
 */
function fetchOpeningFromDailyReportMail_(year, month, options) {
  options = options || {};
  const monthLabel = year + "年" + (month + 1) + "月";

  const msg1 = findDailyReportMailForDate_(year, month, 1);
  if (msg1) {
    const parsed = parseNippoMailOptionsTable_(msg1.getBody());
    const check = validateOpeningMailParse_(parsed, "opening");
    const values = parsedMailToOpeningValues_(parsed, "opening");
    if (check.valid) {
      return {
        values: values,
        source: "日報メール " + msg1.getSubject() + "（月初列）",
        subject: msg1.getSubject(),
        monthLabel: monthLabel,
      };
    }
  }

  const prev = shiftYearMonth_(year, month, -1);
  const lastDay = new Date(year, month, 0).getDate();
  const msgEnd = findDailyReportMailForDate_(prev.year, prev.month, lastDay);
  if (msgEnd) {
    const parsed = parseNippoMailOptionsTable_(msgEnd.getBody());
    const check = validateOpeningMailParse_(parsed, "total");
    const values = parsedMailToOpeningValues_(parsed, "total");
    if (check.valid) {
      return {
        values: values,
        source: "日報メール " + msgEnd.getSubject() + "（合計列→翌月月初）",
        subject: msgEnd.getSubject(),
        monthLabel: monthLabel,
      };
    }
  }

  if (!options.skipArchive) {
    const archiveHit = loadOpeningFromArchive_(SpreadsheetApp.getActiveSpreadsheet(), year, month);
    if (archiveHit) return archiveHit;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const legacy = readPrevMonthEndOpeningCounts_(ss, prev.year, prev.month);
  if (legacy.values) {
    return {
      values: legacy.values,
      source: legacy.source + "（フォールバック）",
      subject: "",
      monthLabel: monthLabel,
    };
  }

  return {
    values: null,
    source: "日報メール未検出（" + formatNippoMailDate_(year, month, 1) + "）",
    monthLabel: monthLabel,
  };
}

function ensureOpeningArchiveSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME_OPENING_ARCHIVE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_OPENING_ARCHIVE);
    sheet.getRange(1, 1, 1, 4).setValues([["年月", "オプション名", "月初", "取得元"]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveOpeningToArchive_(ss, year, month, values, source) {
  const sheet = ensureOpeningArchiveSheet_(ss);
  const label = year + "年" + (month + 1) + "月";
  const lastRow = sheet.getLastRow();
  const numRows = Math.max(0, lastRow - 1);
  const data = numRows > 0 ? sheet.getRange(2, 1, numRows, 4).getValues() : [];
  const kept = data.filter(function (row) { return String(row[0]) !== label; });
  const rows = [];
  for (let i = 0; i < OPTION_LIST.length; i++) {
    rows.push([label, OPTION_LIST[i], values[i], source]);
  }
  const all = kept.concat(rows);
  if (all.length === 0) return;
  sheet.getRange(2, 1, all.length, 4).setValues(all);
  const extra = sheet.getLastRow() - 1 - all.length;
  if (extra > 0) {
    sheet.getRange(2 + all.length, 1, extra, 4).clearContent();
  }
}

function loadOpeningFromArchive_(ss, year, month) {
  const sheet = ss.getSheetByName(SHEET_NAME_OPENING_ARCHIVE);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const label = year + "年" + (month + 1) + "月";
  const numRows = sheet.getLastRow() - 1;
  const data = sheet.getRange(2, 1, numRows, 4).getValues();
  const values = OPTION_LIST.map(function () { return null; });
  let found = false;
  let source = "月初アーカイブ";
  data.forEach(function (row) {
    if (String(row[0]) !== label) return;
    const idx = OPTION_LIST.indexOf(String(row[1]));
    if (idx < 0) return;
    values[idx] = parseMonthlyNumeric_(row[2]) || 0;
    if (row[3]) source = String(row[3]);
    found = true;
  });
  if (!found) return null;
  return {
    values: values.map(function (v) { return v === null ? 0 : v; }),
    source: source,
    monthLabel: label,
  };
}

function writeNippoOpeningColumn_(nippo, values) {
  const n = OPTION_LIST.length;
  const dRange = nippo.getRange(NIPPO_OP_START_ROW, NIPPO_COL_OPENING, n, 1);
  // 数値だけ。見た目はシート側のデザインを維持
  dRange.setValues(values.map(function (v) { return [v]; }));
}

/**
 * 日報メールから当月のオプション月初を取得し、日報D列・集計G列・アーカイブへ反映
 * ※ C12（月初会員数）は手入力運用のため自動上書きしない
 */
function applyOpeningFromDailyReportMail_(ss, year, month) {
  const result = fetchOpeningFromDailyReportMail_(year, month, { skipArchive: true, allowZero: true });
  const nippo = ss.getSheetByName("日報");

  if (!result.values) {
    Logger.log("オプション月初取得失敗: " + result.source);
    return { restored: false, source: result.source };
  }

  if (nippo) {
    writeNippoOpeningColumn_(nippo, result.values);
  }

  saveOpeningToArchive_(ss, year, month, result.values, result.source);
  Logger.log("オプション月初反映: " + result.source);
  return { restored: true, source: result.source };
}

function syncSummaryOpeningFromMail_(ss, year, month, writeNippoToo) {
  const result = fetchOpeningFromDailyReportMail_(year, month);
  if (!result.values) return false;
  if (writeNippoToo) {
    const nippo = ss.getSheetByName("日報");
    if (nippo) writeNippoOpeningColumn_(nippo, result.values);
  }
  saveOpeningToArchive_(ss, year, month, result.values, result.source);
  return true;
}

/** 月初(D) … 当月1日の店舗日報メールを正とする */
function carryForwardNippoOpeningBalances_(ss, year, month) {
  return applyOpeningFromDailyReportMail_(ss, year, month);
}

/** フォールバック: yyMMシート・日報から推定 */
function readPrevMonthEndOpeningCounts_(ss, prevYear, prevMonth) {
  const n = OPTION_LIST.length;
  const prevName = formatMonthlySheetName_(prevYear, prevMonth);
  const prevSheet = ss.getSheetByName(prevName);

  if (prevSheet) {
    const totals = readNumericColumn_(prevSheet, MONTHLY_SHEET_OP_START, MONTHLY_SHEET_COL_TOTAL, n);
    if (totals && totals.some(function (v) { return v > 0; })) {
      return { values: totals, source: "シート「" + prevName + "」合計(E列)" };
    }

    const openings = readNumericColumn_(prevSheet, MONTHLY_SHEET_OP_START, MONTHLY_SHEET_COL_OPENING, n);
    if (openings && openings.some(function (v) { return v > 0; })) {
      const nets = readMonthlySheetNetColumn_(ss, prevSheet, prevYear, prevMonth, n);
      const values = openings.map(function (g, i) { return g + (nets[i] || 0); });
      return { values: values, source: "シート「" + prevName + "」月初(G)+純増減(H)" };
    }

    const currents = readNumericColumn_(prevSheet, MONTHLY_SHEET_OP_START, MONTHLY_SHEET_COL_CURRENT, n);
    if (currents && openings) {
      const nets = readMonthlySheetNetColumn_(ss, prevSheet, prevYear, prevMonth, n);
      const values = currents.map(function (c, i) {
        const g = openings[i] || 0;
        const net = nets[i] || 0;
        return g + net;
      });
      if (values.some(function (v) { return v > 0; })) {
        return { values: values, source: "シート「" + prevName + "」G+C/Hから推定" };
      }
    }
  }

  const nippo = ss.getSheetByName("日報");
  if (nippo) {
    const dVals = readNumericColumn_(nippo, NIPPO_OP_START_ROW, NIPPO_COL_OPENING, n);
    const eVals = readNumericColumnFromNetFormulas_(nippo, n);
    if (dVals && (dVals.some(function (v) { return v > 0; }) || eVals.some(function (v) { return v !== 0; }))) {
      const values = dVals.map(function (d, i) { return d + eVals[i]; });
      return { values: values, source: "日報 D列+E列" };
    }
  }

  const opOnly = resolvePrevMonthNetFromOp_(ss, prevYear, prevMonth);
  if (opOnly) {
    return { values: null, source: "先月シート・日報に月初なし（OPのみ:" + prevName + "）" };
  }

  return { values: null, source: "参照先なし（" + prevName + "）" };
}

function readNumericColumn_(sheet, startRow, col, n) {
  const raw = sheet.getRange(startRow, col, n, 1).getValues();
  const hasAny = raw.some(function (row) { return parseMonthlyNumeric_(row[0]) !== null; });
  if (!hasAny) return null;
  return raw.map(function (row) {
    const v = parseMonthlyNumeric_(row[0]);
    return v === null ? 0 : v;
  });
}

function readMonthlySheetNetColumn_(ss, prevSheet, prevYear, prevMonth, n) {
  const hRaw = prevSheet.getRange(MONTHLY_SHEET_OP_START, MONTHLY_SHEET_COL_NET, n, 1).getValues();
  const fromOp = resolvePrevMonthNetFromOp_(ss, prevYear, prevMonth) || [];
  return hRaw.map(function (row, i) {
    const h = parseMonthlyNumeric_(row[0]);
    if (h !== null) return h;
    return fromOp[i] || 0;
  });
}

function readNumericColumnFromNetFormulas_(nippo, n) {
  const formulas = nippo.getRange(NIPPO_OP_START_ROW, NIPPO_COL_NET, n, 1).getFormulas();
  const hasFormula = formulas.some(function (row) { return String(row[0] || "").startsWith("="); });
  if (hasFormula) {
    return nippo.getRange(NIPPO_OP_START_ROW, NIPPO_COL_NET, n, 1).getValues().map(function (row) {
      return parseMonthlyNumeric_(row[0]) || 0;
    });
  }
  return readNumericColumn_(nippo, NIPPO_OP_START_ROW, NIPPO_COL_NET, n) || [];
}

function resolvePrevMonthNetFromOp_(ss, prevYear, prevMonth) {
  const opSheet = getOpSheet_(ss);
  if (readOpLogValues_(opSheet).length < 2) return null;
  const opPrev = countOpDataForMonth_(opSheet, prevYear, prevMonth);
  return opPrev.counts.map(function (c) {
    return c.newSignup + c.opAdd - c.stop;
  });
}

/** 月初更新の日報まわり（表の再構築・繰越・当月リセット） */
function runMonthlyNippoPhase_(ss, year, month, rebuildTable) {
  const nippo = ss.getSheetByName("日報");
  if (!nippo) throw new Error("「日報」シートが見つかりません。");

  if (rebuildTable) {
    rebuildNippoOpTable_(nippo);
  } else {
    setupNippoOpLayout_(nippo);
  }

  const carry = carryForwardNippoOpeningBalances_(ss, year, month);
  resetNippoOpProgressColumns_(nippo);
  setNippoMonthSheetName_(ss, formatMonthlySheetName_(year, month));
  ensureSummaryMonthB1_(ss, year, month);
  return carry;
}

/** 旧メニュー入口（無効）… 月初は手動運用 */
function restoreNippoOpeningMenu_() {
  SpreadsheetApp.getUi().alert(
    "月初復元",
    "この機能はメニューから外しています。月初は手入力で管理してください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** 旧メニュー入口（無効）… 月初は手動運用 */
function importAllOpeningsFromNippoMailsMenu_() {
  SpreadsheetApp.getUi().alert(
    "月初一括取込",
    "この機能はメニューから外しています。月初は手入力で管理してください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ── 月次シート（原本→2607など）・月初更新 ──

const MONTHLY_TEMPLATE_SHEET = "原本";

/** 旧メニュー入口（無効）… 月初は手動運用 */
function runMonthlySheetSetupMenu_() {
  SpreadsheetApp.getUi().alert(
    "月初更新作業",
    "この機能はメニューから外しています。月初は手入力で管理してください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** 毎月1日 19:30 … 当月シート作成 or 月初(D)の取得 */
function maybeAutoMonthlySheetSetup_() {
  const now = new Date();
  if (now.getDate() !== 1) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = now.getFullYear();
  const month = now.getMonth();
  const sheetName = formatMonthlySheetName_(year, month);

  if (!ss.getSheetByName(sheetName)) {
    runMonthlySheetSetupCore_(true);
    return;
  }

  ensureSummaryMonthB1_(ss, year, month);
  setNippoMonthSheetName_(ss, sheetName);
  applyOpeningFromDailyReportMail_(ss, year, month);
}

/** 毎月2〜3日 19:30 … 1日送信の日報メールが届いたあと月初(D)を再取得 */
function maybeRefreshOpeningFromDayOneMail_(ss) {
  const now = new Date();
  const day = now.getDate();
  if (day < 2 || day > 3) return;
  applyOpeningFromDailyReportMail_(ss, now.getFullYear(), now.getMonth());
}

function runMonthlySheetSetupCore_(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const sheetName = formatMonthlySheetName_(year, month);
  const prev = shiftYearMonth_(year, month, -1);
  const prevSheetName = formatMonthlySheetName_(prev.year, prev.month);
  const existing = ss.getSheetByName(sheetName);

  if (existing) {
    if (!silent) {
      const res = ui.alert(
        "月初更新作業",
        "シート「" + sheetName + "」は既にあります。\n\n" +
          "次をまとめて実行します:\n" +
          "・日報オプション欄の再構築\n" +
          "・店舗日報メールから月初(D)・集計(G)を復元\n" +
          "・Gmail取得・集計・当月数値の反映\n" +
          "・入会・退会・日別欄更新\n\n実行しますか？",
        ui.ButtonSet.YES_NO
      );
      if (res !== ui.Button.YES) return;
    }
    const carry = runMonthlyNippoPhase_(ss, year, month, true);
    runMonthlyFullDataSync_(ss, year, month, silent);
    return { sheetName: sheetName, created: false, opening: carry };
  }

  const template = ss.getSheetByName(MONTHLY_TEMPLATE_SHEET);
  if (!template) {
    throw new Error("シート「" + MONTHLY_TEMPLATE_SHEET + "」が見つかりません。");
  }

  if (!silent) {
    const res = ui.alert(
      "月初更新作業",
      "原本から「" + sheetName + "」を作成し、\n" +
        "メール取得・集計・日報反映まで一括で実行します。\n\n" +
        "・B2 … " + year + " / C2 … " + (month + 1) + "\n" +
        "・日報 B1 … " + sheetName + "\n" +
        "・日報 D列(月初) … 店舗日報メール " + formatNippoMailDate_(year, month, 1) + " から取得\n" +
        "・" + sheetName + " … 日別欄のみ空にして運用開始\n\n実行しますか？",
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
  }

  const newSheet = template.copyTo(ss);
  newSheet.setName(sheetName);
  ss.setActiveSheet(newSheet);

  newSheet.getRange("B2").setValue(year);
  newSheet.getRange("C2").setValue(month + 1);
  newSheet.getRange("C2").setNumberFormat("0");

  clearMonthlyDailyCountColumns_(newSheet);
  formatMonthlyDailyArea_(newSheet);

  const carry = runMonthlyNippoPhase_(ss, year, month, true);
  runMonthlyFullDataSync_(ss, year, month, silent);
  return { sheetName: sheetName, created: true, opening: carry };
}

/** Gmail取得→集計→日報→入会退会→日別欄まで一括更新 */
function runMonthlyFullDataSync_(ss, year, month, silent) {
  setupSpreadsheet();
  ensureSummaryMonthB1_(ss, year, month);
  setNippoMonthSheetName_(ss, formatMonthlySheetName_(year, month));

  const opSheet = getOpSheet_(ss);
  if (readOpLogValues_(opSheet).length <= 1) {
    executeFetchAll_(ss, opSheet, true);
  } else {
    executeFetchMonth_(ss, opSheet, true);
  }

  formatNippoOpArea_(ss);
  const monthlySheet = ss.getSheetByName(formatMonthlySheetName_(year, month));
  if (monthlySheet) formatMonthlyDailyArea_(monthlySheet);

  const doneMsg =
    "月初更新が完了しました。\n\n" +
    "・OPメール取得 → OP集計\n" +
    "・店舗日報メール → 日報D列・OP集計G列\n" +
    "・入会・退会・日別欄\n\n" +
    "日報シートで数値をご確認ください。";

  if (silent) {
    Logger.log(doneMsg);
  } else {
    SpreadsheetApp.getUi().alert("月初更新作業", doneMsg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ensureSummaryMonthB1_(ss, year, month) {
  const summary = getOpSheet_(ss);
  if (!summary) return;
  summary.getRange("B1").setValue(year + "年" + (month + 1) + "月");
}

function formatMonthlySheetName_(year, month) {
  const yy = String(year % 100).padStart(2, "0");
  const mm = String(month + 1).padStart(2, "0");
  return yy + mm;
}

function shiftYearMonth_(year, month, deltaMonths) {
  const d = new Date(year, month + deltaMonths, 1);
  return { year: d.getFullYear(), month: d.getMonth() };
}

function setNippoMonthSheetName_(ss, sheetName) {
  const nippoSheet = ss.getSheetByName("日報");
  if (!nippoSheet) return;
  nippoSheet.getRange("B1").setNumberFormat("@");
  nippoSheet.getRange("B1").setValue(sheetName);
}

/** 前月の純増減（日報F列またはOP） */
function resolvePrevMonthNetValues_(ss, prevYear, prevMonth) {
  const n = OPTION_LIST.length;
  const nippo = ss.getSheetByName("日報");
  const prevName = formatMonthlySheetName_(prevYear, prevMonth);

  if (nippo) {
    const b1 = String(nippo.getRange("B1").getDisplayValue() || nippo.getRange("B1").getValue() || "").trim();
    if (b1 === prevName) {
      const eVals = readNumericColumnFromNetFormulas_(nippo, n);
      if (eVals.some(function (v) { return v !== 0; })) return eVals;
    }
  }

  return resolvePrevMonthNetFromOp_(ss, prevYear, prevMonth);
}

function clearMonthlyDailyCountColumns_(sheet) {
  const empty = [];
  for (let d = 0; d < 31; d++) empty.push([""]);
  sheet.getRange(5, 5, 31, 1).setValues(empty);
  sheet.getRange(5, 6, 31, 1).setValues(empty);
  sheet.getRange(40, 5, 31, 1).setValues(empty);
  sheet.getRange(71, 6).clearContent();
}

function formatMonthlyDailyArea_(sheet) {
  const areas = [
    sheet.getRange(5, 5, 31, 1),
    sheet.getRange(5, 7, 31, 1),
    sheet.getRange(40, 5, 31, 1),
  ];
  areas.forEach(function (range) {
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
    range.setFontFamily(NIPPO_OP_FONT);
    range.setFontSize(10);
  });
}

function parseMonthlyNumeric_(value) {
  if (typeof value === "number" && !isNaN(value)) return value;
  const s = String(value || "").trim();
  if (!s) return null;
  const n = parseFloat(s.replace(/,/g, ""));
  return isNaN(n) ? null : n;
}

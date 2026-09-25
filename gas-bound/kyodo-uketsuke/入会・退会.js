/**
 * JOYFIT24経堂 — 入会・退会メール集計
 *
 * 【入会・退会】     見る用（A〜D=入会, F〜J=退会, 3行目から）
 * 【入会・退会_データ】非表示（全メール保管）
 *
 * 入会区分: 6ヶ月割 / 法人会員
 * 退会の J列「退会キャンセル」ON → 退会数から除外（K列にメールID・非表示）
 */

const SHEET_NAME_MEMBERS = "入会・退会";
const SHEET_NAME_DATA = "入会・退会_データ";

const DISPLAY_HEADER_ROW = 2;
const DISPLAY_START_ROW = 3;
const DISPLAY_COLS = 11;

const DATA_ENROLL_COLS = 5;
const DATA_WITHDRAW_COLS = 6;
const DISPLAY_MAX_SCAN = 300;

const DATA_HEADER_ROW = 1;
const DATA_START_ROW = 2;
const DATA_ENROLL_COL = 1;
const DATA_WITHDRAW_COL = 7;

const NYUKAI_SEARCH_QUERY =
  'from:info@joyfit-service.jp subject:(ご入会ありがとうございます)';
const TAIKAI_SEARCH_QUERY =
  'from:info@joyfit-service.jp subject:("ご退会のお手続きについて" OR "【JOYFIT24経堂】ご退会のお手続きについて")';

const LABEL_NYUKAI_GENERAL = "入会メール/一般会員";
const LABEL_NYUKAI_CORPORATE = "入会メール/法人会員";
const LABEL_TAIKAI_GENERAL = "退会メール/一般会員";
const LABEL_TAIKAI_CORPORATE = "退会メール/法人会員";

const ENROLL_HEADERS = ["タイムスタンプ", "氏名", "入会月", "区分", "メールID"];
const WITHDRAW_HEADERS = ["タイムスタンプ", "氏名", "退会月", "区分", "メールID", "退会キャンセル"];
/** 入会・退会シート退会側の表示ヘッダー（F〜J列。メールIDはK列非表示） */
const DISPLAY_WITHDRAW_HEADERS = ["タイムスタンプ", "氏名", "退会月", "区分", "退会キャンセル"];

const CATEGORY_SIX_MONTH = "6ヶ月割";
const CATEGORY_CORPORATE = "法人会員";

/**
 * 月次日報シート（例: 2607）の日別入力列
 *
 * 入会ブロック（5〜35行）:
 *   D = E〜K の合計（数式）← 日報C13もここをSUM
 *   E = メール由来の一般/6ヶ月割（GASが書く）
 *   F = 法人（GASが書く）
 *   G〜K = 移籍・復会・紹介など（手入力・GASは触らない）
 *   D36 = D5:D35 の合計（数式）
 * 退会ブロック:
 *   E40〜 = 退会（GASが書く）
 */
const DAILY_ENROLL_SUM_COL = 4;       // D列 … SUM(E:K) 数式のみ
const DAILY_ENROLL_GENERAL_COL = 5;   // E列 … 一般/6ヶ月割（メール）
const DAILY_ENROLL_CORPORATE_COL = 6; // F列 … 法人（メール）
const DAILY_ENROLL_SUM_FROM_COL = 5;  // E
const DAILY_ENROLL_SUM_TO_COL = 11;   // K
const DAILY_WITHDRAW_COL = 5;         // E列 … 退会ブロック側
const DAILY_WITHDRAW_CORPORATE_ADVANCE_COL = 6; // F列 … 法人前受
const DAILY_WITHDRAW_CORPORATE_ADVANCE_ROW = 71;
const DAILY_ENROLL_ROW_START = 5;     // 1日 → 5行目、31日 → 35行目
const DAILY_WITHDRAW_ROW_START = 40;  // 1日 → 40行目、31日 → 70行目
const DAILY_COUNT_DAYS = 31;
const DAILY_ENROLL_TOTAL_ROW = 36;    // D36 = 当月入会合計

/**
 * 19:30自動・管理画面「メール取込」… 直近N日だけGmailを確認
 * （3日だと取りこぼしやすいので広めに取る。毎日走る前提）
 */
const SILENT_FETCH_DAYS_BACK = 14;
const MEMBERSHIP_UPDATE_DAYS_BACK = 14;
/** 1回の検索で見るスレッド上限（新しい順。GAS上限100） */
const MEMBERSHIP_UPDATE_MAX_THREADS = 100;
/** 全件再取得後にラベルを付け直す日数（全期間ラベル付けはタイムアウトしやすい） */
const MEMBERSHIP_LABEL_CATCHUP_DAYS = 45;

// ── メニュー ──

/** JOYFITメニュー: Gmailの入会・退会メールをすべて取り直す（_データを置換） */
function refetchAllMembershipEmailsMenu_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "入会・退会メールを全件再取得",
    "Gmailに残っている入会・退会メールをすべて取り直します。\n\n" +
      "・「入会・退会_データ」は置き換わります\n" +
      "・Gmailに無い古い月は復元できません\n" +
      "・数分かかることがあります（目安3〜5分）\n" +
      "・6分を超えると途中で止まることがあります\n\n実行しますか？",
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  setupDataSheet_(ss);
  migrateLegacyLayouts_(ss);
  setupDisplaySheet_(ss);
  installMembershipLabelsSilent_();

  // データは全件。ラベルは直近だけ（全期間はタイムアウトしやすい）
  const full = fetchAllMembershipEmailsCore_(ss, { skipLabels: true });
  applyMembershipLabelsRecent_(MEMBERSHIP_LABEL_CATCHUP_DAYS);
  repairMonthValuesInData_(ss);
  dedupeStoredMemberData_(ss);
  const view = refreshMembershipDisplay_(ss);

  ui.alert(
    "入会・退会メールを全件再取得",
    "完了しました。\n\n" +
      "取得 … 入会 " + full.enrollTotal + " 件 / 退会 " + full.withdrawTotal + " 件\n" +
      "表示 … 入会 " + view.enrollLabel + "：" + view.enrollCount + " 件 / " +
      "退会 " + view.withdrawLabel + "：" + view.withdrawCount + " 件\n\n" +
      "アンケートを見る場合は JOYFIT → アンケートを更新 を実行してください。",
    ui.ButtonSet.OK
  );
}

function updateMembershipSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  setupDataSheet_(ss);
  migrateLegacyLayouts_(ss);
  setupDisplaySheet_(ss);
  installMembershipLabelsSilent_();
  repairMonthValuesInData_(ss);
  dedupeStoredMemberData_(ss);

  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  const enrollCount = readEnrollData_(dataSheet).length;
  const withdrawCount = readWithdrawData_(dataSheet).length;
  const isEmpty = enrollCount === 0 && withdrawCount === 0;

  let enrollAdded = 0;
  let withdrawAdded = 0;
  let enrollTotal = enrollCount;
  let withdrawTotal = withdrawCount;

  if (isEmpty) {
    const res = ui.alert(
      "初回取得",
      "まだデータがありません。\n過去のメールをすべて取得します（数分かかる場合があります）。\n\n実行しますか？",
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) {
      refreshMembershipDisplay_(ss);
      return;
    }
    const full = fetchAllMembershipEmailsCore_(ss);
    enrollTotal = full.enrollTotal;
    withdrawTotal = full.withdrawTotal;
  } else {
    const inc = fetchMembershipEmailsIncremental_(ss, MEMBERSHIP_UPDATE_DAYS_BACK);
    enrollAdded = inc.enrollAdded;
    withdrawAdded = inc.withdrawAdded;
    enrollTotal = readEnrollData_(dataSheet).length;
    withdrawTotal = readWithdrawData_(dataSheet).length;
  }

  // 再送・スレッド返信で溜まった重複をここで潰す
  dedupeStoredMemberData_(ss);
  enrollTotal = readEnrollData_(dataSheet).length;
  withdrawTotal = readWithdrawData_(dataSheet).length;

  const view = refreshMembershipDisplay_(ss);

  try {
    syncMembershipDailyCountsSilent_();
  } catch (err) {
    Logger.log("日別入会・退会の反映エラー: " + err.message);
  }

  ui.alert(
    "入会・退会を更新",
    "完了しました。\n\n" +
      (isEmpty
        ? "取得 … 入会 " + enrollTotal + " 件 / 退会 " + withdrawTotal + " 件（初回・全件）\n"
        : "新規 … 入会 +" + enrollAdded + " 件 / 退会 +" + withdrawAdded + " 件（直近" + MEMBERSHIP_UPDATE_DAYS_BACK + "日）\n" +
          "合計 … 入会 " + enrollTotal + " 件 / 退会 " + withdrawTotal + " 件\n") +
      "表示 … 入会 " + view.enrollLabel + "：" + view.enrollCount + " 件 / " +
      "退会 " + view.withdrawLabel + "：" + view.withdrawCount + " 件" +
      (view.withdrawCancelCount ? "（キャンセル " + view.withdrawCancelCount + " 件除く）" : "") + "\n\n" +
      "※OP・日報は毎日19:30に自動更新されます。\n" +
      "※当月シート（例: 6月なら2606）の日別欄にも反映されます。\n" +
      "※過去をまとめて取り直す … JOYFIT「入会・退会メールを全件再取得」",
    ui.ButtonSet.OK
  );
}

/** 過去メールをすべて取得（初回・全件再取得） */
function fetchAllMembershipEmailsCore_(ss, opt) {
  opt = opt || {};
  // Gmail取得が途中で失敗しても既存データを失わないよう、
  // 全件の取得完了後にシートを置き換える。
  const enrollRows = fetchMembershipMailRows_(NYUKAI_SEARCH_QUERY, extractNyukaiRowsFromThreads_);
  const withdrawRows = fetchMembershipMailRows_(TAIKAI_SEARCH_QUERY, extractTaikaiRowsFromThreads_);

  writeEnrollData_(ss, enrollRows);
  writeWithdrawData_(ss, withdrawRows);

  if (!opt.skipLabels) {
    const enrollThreads = searchGmailAllThreads_(NYUKAI_SEARCH_QUERY);
    const withdrawThreads = searchGmailAllThreads_(TAIKAI_SEARCH_QUERY);
    applyNyukaiLabelsToThreads_(enrollThreads);
    applyTaikaiLabelsToThreads_(withdrawThreads);
  }

  return { enrollTotal: enrollRows.length, withdrawTotal: withdrawRows.length };
}

/** 年ごとにGmail検索して結合（1回の検索が巨大になりすぎないよう分割） */
function fetchMembershipMailRows_(baseQuery, extractFn) {
  const startYear = 2022;
  const endYear = new Date().getFullYear() + 1;
  const threads = [];
  const seen = {};

  for (let y = startYear; y < endYear; y++) {
    const q = baseQuery + " after:" + y + "/1/1 before:" + (y + 1) + "/1/1";
    searchGmailAllThreads_(q).forEach(function (thread) {
      const id = thread.getId();
      if (seen[id]) return;
      seen[id] = true;
      threads.push(thread);
    });
  }

  return extractFn(threads);
}

/**
 * 直近N日・新しいスレッド最大N件だけ確認し、新着のみ追加（手動「最新化」用）
 * 全件検索はしない（終わらない原因だった）
 */
function fetchMembershipEmailsIncremental_(ss, daysBack, opt) {
  opt = opt || {};
  const enrollAdded = fetchMembershipEmailsIncrementalKind_(ss, "enroll", daysBack, opt);
  const withdrawAdded = fetchMembershipEmailsIncrementalKind_(ss, "withdraw", daysBack, opt);
  return { enrollAdded: enrollAdded, withdrawAdded: withdrawAdded };
}

/** kind: "enroll" | "withdraw" */
function fetchMembershipEmailsIncrementalKind_(ss, kind, daysBack, opt) {
  opt = opt || {};
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  const isEnroll = kind !== "withdraw";
  const startCol = isEnroll ? DATA_ENROLL_COL : DATA_WITHDRAW_COL;
  const colCount = isEnroll ? DATA_ENROLL_COLS : DATA_WITHDRAW_COLS;
  // メールID列だけ読む（全列読みは重い＝やらかしポイントだった）
  const existingIds = loadMessageIdsFromSheet_(dataSheet, startCol + 4);
  const seen = { msgId: Object.assign({}, existingIds), composite: {} };
  const lookbackDays = daysBack == null ? MEMBERSHIP_UPDATE_DAYS_BACK : daysBack;
  const afterQuery = " after:" + gmailAfterDays_(lookbackDays);
  const query = (isEnroll ? NYUKAI_SEARCH_QUERY : TAIKAI_SEARCH_QUERY) + afterQuery;
  const parseFn = isEnroll ? parseNyukaiMessage_ : parseTaikaiMessage_;

  const newRows = [];
  const threads = searchGmailRecentThreads_(query, MEMBERSHIP_UPDATE_MAX_THREADS);
  for (let t = 0; t < threads.length; t++) {
    collectNewMembershipFromThread_(
      threads[t], existingIds, seen, newRows, colCount, parseFn
    );
  }

  // 新着だけ末尾に追記（全件書き直ししない＝もう一つのやらかしポイントだった）
  if (newRows.length > 0) {
    appendMembershipRows_(dataSheet, startCol, colCount, newRows);
  }
  // 手動更新時だけラベルを整える。自動更新は既存ラベルを触らず高速化。
  if (!opt.skipLabels) {
    if (isEnroll) applyNyukaiLabelsToThreads_(threads);
    else applyTaikaiLabelsToThreads_(threads);
  }
  return newRows.length;
}

/** メールID列だけ読んで重複判定用マップを作る（高速） */
function loadMessageIdsFromSheet_(dataSheet, idCol) {
  const ids = {};
  if (!dataSheet) return ids;
  const lastRow = dataSheet.getLastRow();
  if (lastRow < DATA_START_ROW) return ids;
  const numRows = lastRow - DATA_START_ROW + 1;
  const values = dataSheet.getRange(DATA_START_ROW, idCol, numRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const id = String(values[i][0] || "").trim();
    if (id) ids[id] = true;
  }
  return ids;
}

/** 新着行だけ末尾に追記（既存データを消して書き直さない） */
function appendMembershipRows_(dataSheet, startCol, colCount, newRows) {
  if (!dataSheet || !newRows.length) return;
  const normalized = newRows.map(function (r) { return normalizeMemberRow_(r, colCount); });
  const lastRow = dataSheet.getLastRow();
  const startRow = lastRow >= DATA_START_ROW ? lastRow + 1 : DATA_START_ROW;
  const monthCol = startCol + 2;
  dataSheet.getRange(startRow, monthCol, normalized.length, 1).setNumberFormat("@");
  dataSheet.getRange(startRow, startCol, normalized.length, colCount).setValues(normalized);
  dataSheet.getRange(startRow, startCol, normalized.length, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  applyMonthTextToColumn_(
    dataSheet, monthCol, startRow,
    normalized.map(function (r) { return r[2]; })
  );
}

/** 新しいスレッドだけ（最大 maxThreads 件・1ページ） */
function searchGmailRecentThreads_(query, maxThreads) {
  const limit = Math.max(1, Math.min(100, maxThreads || MEMBERSHIP_UPDATE_MAX_THREADS));
  return GmailApp.search(query, 0, limit);
}

/** スレッド内の未登録メールを取り込む
 * 同一スレッドは成立する最新1通だけ（返信・再送の二重登録を防ぐ）
 */
function collectNewMembershipFromThread_(thread, existingIds, seen, outRows, colCount, parseFn) {
  const messages = getThreadMessagesNewestFirst_(thread);
  for (let i = 0; i < messages.length; i++) {
    const message = messages[i];
    const msgId = message.getId();
    if (existingIds[msgId]) continue;
    const row = parseFn(message);
    if (!row) continue;
    if (!registerMembershipRow_(row, seen, outRows, colCount)) {
      // 氏名+月が既存なら、このメールIDも消化済み扱いにする
      existingIds[msgId] = true;
      continue;
    }
    existingIds[msgId] = true;
    // スレッド内の残りも同一人物の再送・返信とみなして消化
    for (let j = i + 1; j < messages.length; j++) {
      existingIds[messages[j].getId()] = true;
    }
    break;
  }
}

/** @deprecated collectNewMembershipFromThread_ を使う */
function collectNewestMembershipFromThread_(thread, existingIds, seen, outRows, colCount, parseFn) {
  collectNewMembershipFromThread_(thread, existingIds, seen, outRows, colCount, parseFn);
}

function gmailAfterDate_(monthsBack) {
  const d = new Date();
  d.setDate(1);
  d.setMonth(d.getMonth() - (monthsBack || 0));
  return d.getFullYear() + "/" + String(d.getMonth() + 1).padStart(2, "0") + "/01";
}

function gmailAfterDays_(daysBack) {
  const d = new Date();
  d.setDate(d.getDate() - (daysBack || 0));
  return d.getFullYear() + "/" + String(d.getMonth() + 1).padStart(2, "0") + "/" +
    String(d.getDate()).padStart(2, "0");
}

/** 毎日の自動更新／メニュー更新から呼ばれる（直近N日の新着追記＋ラベル） */
function fetchMembershipEmailsSilent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDataSheet_(ss);
  // 数値とラベルを両方更新（以前 skipLabels:true だとラベルが止まった）
  return fetchMembershipEmailsIncremental_(
    ss,
    SILENT_FETCH_DAYS_BACK,
    { skipLabels: false }
  );
}

/** 直近N日の入会・退会メールにラベルを付ける。付与したスレッド数を返す */
function applyMembershipLabelsRecent_(daysBack) {
  const days = daysBack == null ? MEMBERSHIP_LABEL_CATCHUP_DAYS : daysBack;
  const afterQuery = " after:" + gmailAfterDays_(days);
  const limit = Math.min(100, MEMBERSHIP_UPDATE_MAX_THREADS);
  const enrollThreads = searchGmailRecentThreads_(NYUKAI_SEARCH_QUERY + afterQuery, limit);
  const withdrawThreads = searchGmailRecentThreads_(TAIKAI_SEARCH_QUERY + afterQuery, limit);
  applyNyukaiLabelsToThreads_(enrollThreads);
  applyTaikaiLabelsToThreads_(withdrawThreads);
  return enrollThreads.length + withdrawThreads.length;
}

/** 直近N日の入会・退会・OP追加停止メールを既読にする */
function markRecentJoyfitMailsRead_(daysBack) {
  const days = daysBack == null ? 7 : daysBack;
  const afterQuery = " after:" + gmailAfterDays_(days);
  const limit = Math.min(100, MEMBERSHIP_UPDATE_MAX_THREADS);
  const queries = [
    NYUKAI_SEARCH_QUERY + afterQuery,
    TAIKAI_SEARCH_QUERY + afterQuery,
    'from:info@joyfit-service.jp subject:オプションご契約につきまして' + afterQuery
  ];
  let count = 0;
  queries.forEach(function (q) {
    searchGmailRecentThreads_(q, limit).forEach(function (thread) {
      if (thread.isUnread()) {
        thread.markRead();
        count++;
      }
    });
  });
  return count;
}

function handleTaikaiSheetEdit_(e) {
  if (!e || !e.range) return false;
  if (e.range.getSheet().getName() !== SHEET_NAME_MEMBERS) return false;
  const cell = e.range.getA1Notation();
  if (cell === "B1" || cell === "G1") {
    refreshMembershipDisplay_(e.source);
    return true;
  }
  if (e.range.getColumn() === 10 && e.range.getRow() >= DISPLAY_START_ROW) {
    persistWithdrawCancelFromDisplay_(e);
    try {
      syncMembershipDailyCountsSilent_();
    } catch (err) {
      Logger.log("日別退会の再反映エラー: " + err.message);
    }
    return true;
  }
  return false;
}

function persistWithdrawCancelFromDisplay_(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const mailId = String(sheet.getRange(row, 11).getValue() || "").trim();
  const checked = e.range.getValue() === true;
  const ss = e.source;
  const rows = readWithdrawData_(ss.getSheetByName(SHEET_NAME_DATA));
  let updated = false;

  for (let i = 0; i < rows.length; i++) {
    if (mailId && String(rows[i][4] || "") === mailId) {
      rows[i][5] = checked;
      updated = true;
      break;
    }
  }
  if (!updated && !mailId) {
    const name = String(sheet.getRange(row, 7).getValue() || "").trim();
    const month = normalizeYearMonthLabel_(sheet.getRange(row, 8).getDisplayValue() || sheet.getRange(row, 8).getValue());
    for (let j = 0; j < rows.length; j++) {
      if (String(rows[j][1] || "").trim() === name && normalizeYearMonthLabel_(rows[j][2]) === month) {
        rows[j][5] = checked;
        updated = true;
        break;
      }
    }
  }
  if (updated) writeWithdrawData_(ss, rows);
}

// ── シート構造 ──

function setupDataSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME_DATA);
  sheet.hideSheet();
  sheet.getRange(DATA_HEADER_ROW, DATA_ENROLL_COL, 1, DATA_ENROLL_COLS).setValues([ENROLL_HEADERS]);
  sheet.getRange(DATA_HEADER_ROW, DATA_WITHDRAW_COL, 1, DATA_WITHDRAW_COLS).setValues([WITHDRAW_HEADERS]);
  sheet.getRange(DATA_HEADER_ROW, DATA_ENROLL_COL, 1, DATA_ENROLL_COLS).setFontWeight("bold").setBackground("#e2efda");
  sheet.getRange(DATA_HEADER_ROW, DATA_WITHDRAW_COL, 1, DATA_WITHDRAW_COLS).setFontWeight("bold").setBackground("#fce4d6");
  sheet.setFrozenRows(1);
}

function setupDisplaySheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME_MEMBERS);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME_MEMBERS);

  removeSheetFilter_(sheet);

  // JOYFIT調（赤・白・黒・グレー）… 6ヶ月割管理と統一
  sheet.getRange("A1")
    .setValue("入会年月を選択 ➡")
    .setFontWeight("bold")
    .setFontColor("#5f6368")
    .setHorizontalAlignment("right");
  sheet.getRange("F1")
    .setValue("退会年月を選択 ➡")
    .setFontWeight("bold")
    .setFontColor("#5f6368")
    .setHorizontalAlignment("right");
  sheet.getRange(DISPLAY_HEADER_ROW, 1, 1, 4).setValues([ENROLL_HEADERS.slice(0, 4)]);
  sheet.getRange(DISPLAY_HEADER_ROW, 6, 1, DISPLAY_WITHDRAW_HEADERS.length)
    .setValues([DISPLAY_WITHDRAW_HEADERS]);
  const enrollHeader = sheet.getRange(DISPLAY_HEADER_ROW, 1, 1, 4);
  const withdrawHeader = sheet.getRange(DISPLAY_HEADER_ROW, 6, 1, DISPLAY_WITHDRAW_HEADERS.length);
  enrollHeader.setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  withdrawHeader.setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  // 入会側と退会側の区切り（赤ライン）
  sheet.getRange(DISPLAY_HEADER_ROW, 1, 1, 4).setBorder(null, null, true, null, null, null, "#C21632", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(DISPLAY_HEADER_ROW, 6, 1, DISPLAY_WITHDRAW_HEADERS.length)
    .setBorder(null, null, true, null, null, null, "#C21632", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.setFrozenRows(DISPLAY_START_ROW - 1);
  sheet.setColumnWidth(1, 155);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 95);
  sheet.setColumnWidth(4, 85);
  sheet.setColumnWidth(6, 155);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 95);
  sheet.setColumnWidth(9, 85);
  sheet.setColumnWidth(10, 95);
  sheet.hideColumns(5);
  sheet.hideColumns(11);
  sheet.getRange("B1").setNumberFormat("@");
  sheet.getRange("G1").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  sheet.getRange("H:H").setNumberFormat("@");
}

function refreshMembershipDisplay_(ss, ymOpt) {
  setupDisplaySheet_(ss);
  const displaySheet = ss.getSheetByName(SHEET_NAME_MEMBERS);
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  try {
    if (displaySheet && displaySheet.isSheetHidden()) displaySheet.showSheet();
  } catch (e) { /* ignore */ }
  const allEnroll = readEnrollData_(dataSheet);
  const allWithdraw = readWithdrawData_(dataSheet);

  if (ymOpt && ymOpt.year != null && ymOpt.month != null) {
    const label = ymOpt.year + "年" + (ymOpt.month + 1) + "月";
    setMonthTextCell_(displaySheet.getRange("B1"), label);
    setMonthTextCell_(displaySheet.getRange("G1"), label);
  }

  setupMonthSelectorCell_(displaySheet.getRange("B1"), extractMonthListFromRows_(allEnroll), true);
  setupMonthSelectorCell_(displaySheet.getRange("G1"), extractMonthListFromRows_(allWithdraw), true);

  const enrollLabel = normalizeYearMonthLabel_(
    displaySheet.getRange("B1").getDisplayValue() || displaySheet.getRange("B1").getValue()
  );
  const withdrawLabel = normalizeYearMonthLabel_(
    displaySheet.getRange("G1").getDisplayValue() || displaySheet.getRange("G1").getValue()
  );

  const filteredEnroll = allEnroll.filter(function (r) {
    return normalizeYearMonthLabel_(r[2]) === enrollLabel;
  });
  const filteredWithdraw = allWithdraw.filter(function (r) {
    return normalizeYearMonthLabel_(r[2]) === withdrawLabel;
  });

  sortRowsNewestFirst_(filteredEnroll);
  sortRowsNewestFirst_(filteredWithdraw);
  const cancelMap = loadWithdrawCancelMapIfAvailable_();
  writeDisplayRows_(displaySheet, filteredEnroll, filteredWithdraw, cancelMap);

  return {
    enrollCount: filteredEnroll.length,
    withdrawCount: filteredWithdraw.filter(function (r) { return !isWithdrawalCancelled_(r, cancelMap); }).length,
    withdrawCancelCount: filteredWithdraw.filter(function (r) { return isWithdrawalCancelled_(r, cancelMap); }).length,
    enrollLabel: enrollLabel,
    withdrawLabel: withdrawLabel
  };
}

function writeDisplayRows_(sheet, enrollRows, withdrawRows, cancelMap) {
  removeSheetFilter_(sheet);

  const rowCount = Math.max(enrollRows.length, withdrawRows.length);
  const prevRows = countDisplayDataRows_(sheet);
  const clearRows = Math.max(rowCount, prevRows);

  if (clearRows > 0) {
    sheet.getRange(DISPLAY_START_ROW, 1, clearRows, DISPLAY_COLS).clearContent();
    sheet.getRange(DISPLAY_START_ROW, 10, clearRows, 1).removeCheckboxes();
  }
  if (rowCount === 0) return;

  const withdrawCount = withdrawRows.length;
  const displayRows = [];
  for (let i = 0; i < rowCount; i++) {
    const row = new Array(DISPLAY_COLS).fill("");
    if (i < enrollRows.length) {
      row[0] = enrollRows[i][0];
      row[1] = enrollRows[i][1];
      row[2] = normalizeYearMonthLabel_(enrollRows[i][2]);
      row[3] = enrollRows[i][3];
    }
    if (i < withdrawCount) {
      row[5] = withdrawRows[i][0];
      row[6] = withdrawRows[i][1];
      row[7] = normalizeYearMonthLabel_(withdrawRows[i][2]);
      row[8] = withdrawRows[i][3];
      // 表示は6ヶ月割管理の退会キャンセルも反映（編集の正は6ヶ月割管理側）
      row[9] = isWithdrawalCancelled_(withdrawRows[i], cancelMap);
      row[10] = withdrawRows[i][4];
    }
    displayRows.push(row);
  }

  sheet.getRange(DISPLAY_START_ROW, 1, displayRows.length, DISPLAY_COLS).setValues(displayRows);
  sheet.getRange(DISPLAY_START_ROW, 1, displayRows.length, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  sheet.getRange(DISPLAY_START_ROW, 6, displayRows.length, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  // データ行もJOYFIT調（白地・黒文字・薄いグレー交互）
  const enrollDataRange = sheet.getRange(DISPLAY_START_ROW, 1, displayRows.length, 4);
  const withdrawDataRange = sheet.getRange(DISPLAY_START_ROW, 6, displayRows.length, 5);
  enrollDataRange.setFontColor("#111111").setVerticalAlignment("middle");
  withdrawDataRange.setFontColor("#111111").setVerticalAlignment("middle");
  const enrollBackgrounds = [];
  const withdrawBackgrounds = [];
  for (let i = 0; i < displayRows.length; i++) {
    const bg = i % 2 === 0 ? "#ffffff" : "#f1f3f4";
    enrollBackgrounds.push([bg, bg, bg, bg]);
    withdrawBackgrounds.push([bg, bg, bg, bg, bg]);
  }
  enrollDataRange.setBackgrounds(enrollBackgrounds);
  withdrawDataRange.setBackgrounds(withdrawBackgrounds);
  if (withdrawCount > 0) {
    sheet.getRange(DISPLAY_START_ROW, 10, withdrawCount, 1).insertCheckboxes();
  }

  const enrollMonths = [];
  const withdrawMonths = [];
  for (let i = 0; i < rowCount; i++) {
    enrollMonths.push(i < enrollRows.length ? normalizeYearMonthLabel_(enrollRows[i][2]) : "");
    withdrawMonths.push(i < withdrawRows.length ? normalizeYearMonthLabel_(withdrawRows[i][2]) : "");
  }
  applyMonthTextToColumn_(sheet, 3, DISPLAY_START_ROW, enrollMonths);
  applyMonthTextToColumn_(sheet, 8, DISPLAY_START_ROW, withdrawMonths);
}

/** 3行目から連続する表示行のみ数える（900行目以降の残骸は無視） */
function countDisplayDataRows_(sheet) {
  const maxRow = Math.min(sheet.getLastRow(), DISPLAY_START_ROW + DISPLAY_MAX_SCAN - 1);
  if (maxRow < DISPLAY_START_ROW) return 0;

  const numRows = maxRow - DISPLAY_START_ROW + 1;
  const values = sheet.getRange(DISPLAY_START_ROW, 1, numRows, DISPLAY_COLS).getValues();
  let count = 0;
  let emptyStreak = 0;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const hasData = row[0] || row[1] || row[2] || row[5] || row[6] || row[7];
    if (hasData) {
      count = i + 1;
      emptyStreak = 0;
    } else {
      emptyStreak++;
      if (emptyStreak >= 10 && count > 0) break;
    }
  }
  return count;
}

function setupMonthSelectorCell_(cell, monthList, keepCurrent) {
  cell.clearDataValidations();
  cell.clearFormat();
  cell.setBackground("#f1f3f4")
    .setFontWeight("bold")
    .setFontColor("#111111")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null, "#C21632", SpreadsheetApp.BorderStyle.SOLID);

  if (!monthList.length) {
    setMonthTextCell_(cell, "（データなし）");
    return;
  }

  const val = normalizeYearMonthLabel_(cell.getDisplayValue() || cell.getValue());
  const pick = (keepCurrent && monthList.indexOf(val) >= 0) ? val : monthList[0];
  setMonthTextCell_(cell, pick);

  cell.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(monthList, true).build()
  );
}

function extractMonthListFromRows_(rows) {
  const keys = {};
  rows.forEach(function (r) {
    const label = normalizeYearMonthLabel_(r[2]);
    if (label) keys[label] = yearMonthSortKey_(label);
  });
  return Object.keys(keys).sort(function (a, b) { return keys[b] - keys[a]; });
}

function yearMonthSortKey_(label) {
  const m = String(label).match(/^(\d{4})年(\d{1,2})月$/);
  return m ? parseInt(m[1], 10) * 100 + parseInt(m[2], 10) : 0;
}

function removeSheetFilter_(sheet) {
  if (!sheet) return;
  const f = sheet.getFilter();
  if (f) f.remove();
}

// ── データ読み書き ──

function readEnrollData_(dataSheet) {
  return readDataBlock_(dataSheet, DATA_ENROLL_COL, DATA_ENROLL_COLS);
}

function readWithdrawData_(dataSheet) {
  return readDataBlock_(dataSheet, DATA_WITHDRAW_COL, DATA_WITHDRAW_COLS);
}

function readDataBlock_(dataSheet, startCol, colCount) {
  if (!dataSheet) return [];
  const lastRow = dataSheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];
  const numRows = lastRow - DATA_START_ROW + 1;
  return dataSheet.getRange(DATA_START_ROW, startCol, numRows, colCount)
    .getValues()
    .filter(function (r) { return r[0] || r[1]; })
    .map(function (r) { return normalizeMemberRow_(r, colCount); });
}

function writeEnrollData_(ss, rows) {
  writeDataBlock_(ss, DATA_ENROLL_COL, dedupeMemberRows_(rows, DATA_ENROLL_COLS), DATA_ENROLL_COLS);
}

function writeWithdrawData_(ss, rows) {
  writeDataBlock_(ss, DATA_WITHDRAW_COL, dedupeMemberRows_(rows, DATA_WITHDRAW_COLS), DATA_WITHDRAW_COLS);
}

function writeDataBlock_(ss, startCol, rows, colCount) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return;

  const normalized = rows.map(function (r) { return normalizeMemberRow_(r, colCount); });
  sortRowsNewestFirst_(normalized);

  const lastRow = dataSheet.getLastRow();
  if (lastRow >= DATA_START_ROW) {
    dataSheet.getRange(DATA_START_ROW, startCol, lastRow - DATA_START_ROW + 1, colCount).clearContent();
  }
  if (!normalized.length) return;

  const monthCol = startCol + 2;
  dataSheet.getRange(DATA_START_ROW, monthCol, normalized.length, 1).setNumberFormat("@");
  dataSheet.getRange(DATA_START_ROW, startCol, normalized.length, colCount).setValues(normalized);
  dataSheet.getRange(DATA_START_ROW, startCol, normalized.length, 1).setNumberFormat("yyyy/mm/dd hh:mm");

  const monthLabels = normalized.map(function (r) { return r[2]; });
  applyMonthTextToColumn_(dataSheet, monthCol, DATA_START_ROW, monthLabels);
}

/** 裏方データの「入会月」「退会月」を文字列に直す */
function repairMonthValuesInData_(ss) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return;

  const enroll = readEnrollData_(dataSheet);
  const withdraw = readWithdrawData_(dataSheet);
  let changed = false;

  enroll.forEach(function (r) {
    const fixed = normalizeYearMonthLabel_(r[2]);
    if (r[2] instanceof Date || /Mon |GMT/.test(String(r[2])) || fixed !== r[2]) {
      r[2] = fixed;
      changed = true;
    }
  });
  withdraw.forEach(function (r) {
    const fixed = normalizeYearMonthLabel_(r[2]);
    if (r[2] instanceof Date || /Mon |GMT/.test(String(r[2])) || fixed !== r[2]) {
      r[2] = fixed;
      changed = true;
    }
  });

  if (changed) {
    writeEnrollData_(ss, enroll);
    writeWithdrawData_(ss, withdraw);
  }
}

function clearAllMemberData_(ss) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (dataSheet && dataSheet.getLastRow() >= DATA_START_ROW) {
    dataSheet.getRange(DATA_START_ROW, 1, dataSheet.getLastRow() - DATA_START_ROW + 1, 12).clearContent();
  }
}

function loadMessageIds_(rows) {
  const ids = {};
  rows.forEach(function (r) {
    const id = String(r[4] || "").trim();
    if (id) ids[id] = true;
  });
  return ids;
}

function normalizeMemberRow_(row, colCount) {
  const n = colCount || DATA_ENROLL_COLS;
  const out = row.slice(0, n);
  while (out.length < n) out.push("");
  out[2] = normalizeYearMonthLabel_(out[2]);
  if (n >= DATA_WITHDRAW_COLS && out[3] === "一般会員") out[3] = CATEGORY_SIX_MONTH;
  if (n >= DATA_WITHDRAW_COLS) out[5] = isWithdrawalCancelled_(out);
  if (n === DATA_ENROLL_COLS && out[3] === "一般会員") out[3] = CATEGORY_SIX_MONTH;
  return out;
}

/**
 * 退会キャンセル判定
 * 1) 入会・退会シートの旧チェック列
 * 2) 6ヶ月割管理の「退会キャンセル」（cancelMap を渡したときだけ）
 * ※ cancelMap 未指定で毎回シートを読むとサイドバーが固まるので禁止
 */
function isWithdrawalCancelled_(row, cancelMap) {
  if (!row) return false;
  if (row.length >= 6) {
    const v = row[5];
    if (v === true || String(v).toUpperCase() === "TRUE") return true;
  }
  if (cancelMap && typeof kyodoIsWithdrawCancelledByMember_ === "function") {
    return kyodoIsWithdrawCancelledByMember_(row[1], null, cancelMap, row[2]);
  }
  return false;
}

function loadWithdrawCancelMapIfAvailable_() {
  if (typeof kyodoLoadWithdrawCancelMap_ !== "function") return null;
  try {
    return kyodoLoadWithdrawCancelMap_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    Logger.log("退会キャンセルマップ: " + err.message);
    return null;
  }
}

function normalizeYearMonthLabel_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return formatYearMonth_(value.getFullYear(), value.getMonth() + 1);
  }
  const s = String(value || "").trim();
  if (!s || s === "（データなし）") return s;

  let m = s.match(/^(\d{4})年(\d{1,2})月$/);
  if (m) return formatYearMonth_(parseInt(m[1], 10), parseInt(m[2], 10));

  m = s.match(/^(\d{4})\/(\d{1,2})$/);
  if (m) return formatYearMonth_(parseInt(m[1], 10), parseInt(m[2], 10));

  if (/Mon |Tue |Wed |Thu |Fri |Sat |Sun |GMT/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return formatYearMonth_(d.getFullYear(), d.getMonth() + 1);
  }

  return s;
}

function formatYearMonth_(year, month) {
  return year + "年" + month + "月";
}

/** 「2026年6月」を日付にされないようテキストとして書き込む */
function setMonthTextCell_(cell, text) {
  const label = normalizeYearMonthLabel_(text);
  cell.setNumberFormat("@");
  if (!label || label === "（データなし）") {
    cell.setValue(label || "");
    return;
  }
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(label).build());
}

function applyMonthTextToColumn_(sheet, col, startRow, labels) {
  if (!labels || !labels.length) return;
  const values = labels.map(function (label) {
    return [label ? normalizeYearMonthLabel_(label) : ""];
  });
  const range = sheet.getRange(startRow, col, values.length, 1);
  range.setNumberFormat("@");
  range.setValues(values);
}

function sortRowsNewestFirst_(rows) {
  rows.sort(function (a, b) {
    const da = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
    const db = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
    return db - da;
  });
}

function dedupeStoredMemberData_(ss) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return;

  const enroll = dedupeMemberRows_(readEnrollData_(dataSheet), DATA_ENROLL_COLS);
  const withdraw = dedupeMemberRows_(readWithdrawData_(dataSheet), DATA_WITHDRAW_COLS);
  const rawEnroll = readEnrollData_(dataSheet);
  const rawWithdraw = readWithdrawData_(dataSheet);
  if (enroll.length !== rawEnroll.length || withdraw.length !== rawWithdraw.length) {
    writeEnrollData_(ss, enroll);
    writeWithdrawData_(ss, withdraw);
  }
}

function createMemberDedupeState_() {
  return { msgId: {}, composite: {} };
}

function loadMemberDedupeState_(rows) {
  const seen = createMemberDedupeState_();
  rows.forEach(function (r) {
    rememberMemberRowKeys_(r, seen);
  });
  return seen;
}

function rememberMemberRowKeys_(row, seen) {
  const colCount = row.length >= DATA_WITHDRAW_COLS ? DATA_WITHDRAW_COLS : DATA_ENROLL_COLS;
  const normalized = normalizeMemberRow_(row, colCount);
  const msgId = String(normalized[4] || "").trim();
  if (msgId) seen.msgId[msgId] = true;
  seen.composite[buildMemberRowDedupeKey_(normalized)] = true;
}

function canRegisterMembershipRow_(row, seen, colCount) {
  const normalized = normalizeMemberRow_(row, colCount);
  const msgId = String(normalized[4] || "").trim();
  if (msgId && seen.msgId[msgId]) return false;
  const composite = buildMemberRowDedupeKey_(normalized);
  if (seen.composite[composite]) return false;
  return true;
}

function registerMembershipRow_(row, seen, rows, colCount) {
  if (!canRegisterMembershipRow_(row, seen, colCount)) return false;
  const normalized = normalizeMemberRow_(row, colCount);
  rememberMemberRowKeys_(normalized, seen);
  rows.push(normalized);
  return true;
}

/**
 * 重複キー … 同一人物×同一入会/退会月×区分は1件
 * （以前は分単位時刻も含めていたため、再送メールで倍増していた）
 */
function buildMemberRowDedupeKey_(row) {
  const name = normalizeMemberName_(row[1]);
  const month = normalizeYearMonthLabel_(row[2]);
  const cat = String(row[3] || "").trim();
  return name + "\x1f" + month + "\x1f" + cat;
}

function normalizeMemberName_(name) {
  return String(name || "").replace(/\s+/g, "").trim();
}

function truncateToMinute_(date) {
  const d = date instanceof Date ? new Date(date.getTime()) : new Date(date);
  d.setSeconds(0, 0);
  return d;
}

function dedupeMemberRows_(rows, colCount) {
  const seen = createMemberDedupeState_();
  const out = [];
  const n = colCount || DATA_ENROLL_COLS;
  rows.forEach(function (r) {
    if (!canRegisterMembershipRow_(r, seen, n)) return;
    registerMembershipRow_(r, seen, out, n);
  });
  return out;
}

function dedupeRows_(rows, colCount) {
  return dedupeMemberRows_(rows, colCount);
}

// ── 指定月の入会・退会をGmailから再構築 ──

/**
 * Apps Scriptエディタから実行する修復用関数。
 * 日報B1（例: 2607）が示す月だけをGmailから取り直し、月次シートへ再反映する。
 *
 * 入会   … 対象月に受信した入会メール
 * 一般退会 … 対象月に受信（当月末退会）
 * 法人退会 … 前月に受信（翌月末退会）
 */
function rebuildCurrentMembershipDailyCountsFromGmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ym = resolveCurrentMonthYm_();
  const result = rebuildMembershipDailyCountsForMonth_(ss, ym.year, ym.month);
  refreshMembershipDisplay_(ss);
  return result.message;
}

function rebuildMembershipDailyCountsForMonth_(ss, year, month) {
  setupDataSheet_(ss);
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  const targetLabel = formatYearMonth_(year, month + 1);
  // Gmailの日付検索は境界のタイムゾーン差を吸収するため前後1日広く取得し、
  // 最終判定は message.getDate() と算出済み退会月で厳密に行う。
  const enrollSearchStart = new Date(year, month, 0);
  const enrollSearchEnd = new Date(year, month + 1, 2);
  const withdrawSearchStart = new Date(year, month - 1, 0);
  const withdrawSearchEnd = new Date(year, month + 1, 2);

  // 入会は対象月受信分だけを正として取り直す。
  const enrollQuery = NYUKAI_SEARCH_QUERY +
    " after:" + formatMembershipGmailDate_(enrollSearchStart) +
    " before:" + formatMembershipGmailDate_(enrollSearchEnd);
  const enrollThreads = searchGmailAllThreads_(enrollQuery);
  const fetchedEnroll = extractNyukaiRowsFromThreads_(enrollThreads).filter(function (row) {
    return isMemberTimestampInMonth_(row[0], year, month);
  });

  // 退会は、対象月の一般＋前月受信の法人が必要なので2カ月分を見る。
  const withdrawQuery = TAIKAI_SEARCH_QUERY +
    " after:" + formatMembershipGmailDate_(withdrawSearchStart) +
    " before:" + formatMembershipGmailDate_(withdrawSearchEnd);
  const withdrawThreads = searchGmailAllThreads_(withdrawQuery);
  const fetchedWithdraw = extractTaikaiRowsFromThreads_(withdrawThreads).filter(function (row) {
    return isMemberTimestampInMonth_(row[0], year, month);
  });

  // 選択月のラベルを一般・法人で付け直す（反対ラベルも除去）
  applyNyukaiLabelsToThreads_(enrollThreads);
  applyTaikaiLabelsToThreads_(withdrawThreads);

  // 対象月だけをGmailの結果で置換し、他の月の保存済みデータは残す。
  const retainedEnroll = readEnrollData_(dataSheet).filter(function (row) {
    return !isMemberTimestampInMonth_(row[0], year, month);
  });
  const retainedWithdraw = readWithdrawData_(dataSheet).filter(function (row) {
    return !isMemberTimestampInMonth_(row[0], year, month);
  });

  writeEnrollData_(ss, retainedEnroll.concat(fetchedEnroll));
  writeWithdrawData_(ss, retainedWithdraw.concat(fetchedWithdraw));
  syncMembershipDailyCountsForMonth_(ss, year, month);

  const audit = buildMembershipMonthAudit_(
    readEnrollData_(dataSheet),
    readWithdrawData_(dataSheet),
    year,
    month
  );
  const sheetName = formatMonthlySheetName_(year, month);
  const message =
    sheetName + " を再集計しました。" +
    "\n入会 " + audit.enrollTotal +
    "（一般/6ヶ月 " + audit.enrollGeneral +
    "・法人 " + audit.enrollCorporate + "）" +
    "\n退会 " + audit.withdrawTotal +
    "（一般 " + audit.withdrawGeneral +
    "・法人 " + audit.withdrawCorporate + "）" +
    "\n退会キャンセル除外 " + audit.withdrawCancelled;
  Logger.log(message);

  return {
    sheetName: sheetName,
    enrollFetched: fetchedEnroll.length,
    withdrawFetched: fetchedWithdraw.length,
    audit: audit,
    message: message
  };
}

/** 年月選択更新用: 原本から月次シートを作成（既存なら何もしない） */
function ensureMembershipMonthlySheet_(ss, year, month) {
  const sheetName = formatMonthlySheetName_(year, month);
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return { sheet: sheet, sheetName: sheetName, created: false };

  const template = ss.getSheetByName("原本");
  if (!template) throw new Error("シート「原本」が見つかりません。");
  sheet = template.copyTo(ss).setName(sheetName);
  sheet.getRange("B2").setValue(year);
  sheet.getRange("C2").setValue(month + 1).setNumberFormat("0");

  // 自動入力欄だけ初期化。G〜Kなどの手入力欄には触れない。
  sheet.getRange(DAILY_ENROLL_ROW_START, DAILY_ENROLL_GENERAL_COL, DAILY_COUNT_DAYS, 1).clearContent();
  sheet.getRange(DAILY_ENROLL_ROW_START, DAILY_ENROLL_CORPORATE_COL, DAILY_COUNT_DAYS, 1).clearContent();
  sheet.getRange(DAILY_WITHDRAW_ROW_START, DAILY_WITHDRAW_COL, DAILY_COUNT_DAYS, 1).clearContent();
  sheet.getRange(DAILY_WITHDRAW_CORPORATE_ADVANCE_ROW, DAILY_WITHDRAW_CORPORATE_ADVANCE_COL).clearContent();
  ensureMonthlyEnrollSumFormulas_(sheet);
  return { sheet: sheet, sheetName: sheetName, created: true };
}

function formatMembershipGmailDate_(date) {
  return date.getFullYear() + "/" +
    String(date.getMonth() + 1).padStart(2, "0") + "/" +
    String(date.getDate()).padStart(2, "0");
}

function isMemberTimestampInMonth_(value, year, month) {
  const date = parseMemberTimestamp_(value);
  return !!date && date.getFullYear() === year && date.getMonth() === month;
}

function buildMembershipMonthAudit_(enrollRows, withdrawRows, year, month) {
  const targetLabel = formatYearMonth_(year, month + 1);
  const audit = {
    enrollGeneral: 0,
    enrollCorporate: 0,
    enrollTotal: 0,
    withdrawGeneral: 0,
    withdrawCorporate: 0,
    withdrawCancelled: 0,
    withdrawTotal: 0
  };

  enrollRows.forEach(function (row) {
    if (!isMemberTimestampInMonth_(row[0], year, month)) return;
    if (String(row[3] || "").trim() === CATEGORY_CORPORATE) {
      audit.enrollCorporate++;
    } else {
      audit.enrollGeneral++;
    }
  });

  const cancelMap = loadWithdrawCancelMapIfAvailable_();
  withdrawRows.forEach(function (row) {
    if (normalizeYearMonthLabel_(row[2]) !== targetLabel) return;
    if (isWithdrawalCancelled_(row, cancelMap)) {
      audit.withdrawCancelled++;
      return;
    }
    if (String(row[3] || "").trim() === CATEGORY_CORPORATE) {
      audit.withdrawCorporate++;
    } else {
      audit.withdrawGeneral++;
    }
  });

  audit.enrollTotal = audit.enrollGeneral + audit.enrollCorporate;
  audit.withdrawTotal = audit.withdrawGeneral + audit.withdrawCorporate;
  return audit;
}

// ── 月次日報シート（2606など）への日別反映 ──

/**
 * 当月の月次シートへ、タイムスタンプ日付ごとの入会・退会件数を書き込む。
 * E=一般/6ヶ月割 / F=法人 / 退会E40〜（キャンセル除く）
 * D列は E〜K の合計数式（上書きしない）
 */
function syncMembershipDailyCountsSilent_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ym = resolveCurrentMonthYm_();
  syncMembershipDailyCountsForMonth_(ss, ym.year, ym.month);
}

function syncMembershipDailyCountsForMonth_(ss, year, month) {
  const sheetName = formatMonthlySheetName_(year, month);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("月次シートが見つかりません: " + sheetName);
    return { sheetName: sheetName, updated: false };
  }

  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return { sheetName: sheetName, updated: false };

  // 集計前に必ず重複を潰す（倍増の主因）
  try {
    dedupeStoredMemberData_(ss);
  } catch (err) {
    Logger.log("入会・退会重複整理: " + (err && err.message ? err.message : err));
  }

  const enrollRows = readEnrollData_(dataSheet);
  const withdrawRows = readWithdrawData_(dataSheet);

  const enrollGeneralByDay = countMembersByDay_(enrollRows, year, month, CATEGORY_SIX_MONTH);
  const enrollCorporateByDay = countMembersByDay_(enrollRows, year, month, CATEGORY_CORPORATE);
  const withdrawByDay = countWithdrawalsByDay_(withdrawRows, year, month);
  const corporateAdvance = countCorporateWithdrawAdvance_(withdrawRows, year, month);

  writeDailyCountColumn_(sheet, DAILY_ENROLL_ROW_START, DAILY_ENROLL_GENERAL_COL, enrollGeneralByDay, "");
  writeDailyCountColumn_(sheet, DAILY_ENROLL_ROW_START, DAILY_ENROLL_CORPORATE_COL, enrollCorporateByDay, "");
  writeDailyCountColumn_(sheet, DAILY_WITHDRAW_ROW_START, DAILY_WITHDRAW_COL, withdrawByDay, "");
  sheet.getRange(DAILY_WITHDRAW_CORPORATE_ADVANCE_ROW, DAILY_WITHDRAW_CORPORATE_ADVANCE_COL)
    .setValue(corporateAdvance || "")
    .setNumberFormat("0");
  ensureMonthlyEnrollSumFormulas_(sheet);

  return { sheetName: sheetName, updated: true };
}

/**
 * D5:D35 = SUM(E:K) / D36 = SUM(D5:D35)
 * 直値で埋まっていても数式に差し替える
 */
function ensureMonthlyEnrollSumFormulas_(sheet) {
  if (!sheet) return;
  const start = DAILY_ENROLL_ROW_START;
  const days = DAILY_COUNT_DAYS;
  const fromLetter = columnToLetter_(DAILY_ENROLL_SUM_FROM_COL);
  const toLetter = columnToLetter_(DAILY_ENROLL_SUM_TO_COL);
  const formulas = [];
  for (let i = 0; i < days; i++) {
    const row = start + i;
    formulas.push(["=SUM(" + fromLetter + row + ":" + toLetter + row + ")"]);
  }
  sheet.getRange(start, DAILY_ENROLL_SUM_COL, days, 1).setFormulas(formulas);
  sheet.getRange(DAILY_ENROLL_TOTAL_ROW, DAILY_ENROLL_SUM_COL)
    .setFormula("=SUM(D" + start + ":D" + (start + days - 1) + ")");
}

function columnToLetter_(col) {
  let n = col;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** エディタから実行可 … 当月シートの D列数式だけ直す */
function repairMonthlyEnrollSumFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ym = resolveCurrentMonthYm_();
  const sheetName = formatMonthlySheetName_(ym.year, ym.month);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("シート「" + sheetName + "」が見つかりません。日報B1を確認してください。");
  ensureMonthlyEnrollSumFormulas_(sheet);
  return sheetName + " の D5:D35 と D36 に合計数式を入れました。";
}

/** 指定列の日別件数だけを書き込む（書式・装飾は触らない） */
function writeDailyCountColumn_(sheet, startRow, col, countsByDay, zeroValue) {
  const values = [];
  for (let day = 1; day <= DAILY_COUNT_DAYS; day++) {
    const n = countsByDay[day] || 0;
    values.push([n > 0 ? n : zeroValue]);
  }
  sheet.getRange(startRow, col, DAILY_COUNT_DAYS, 1).setValues(values);
}

function resolveCurrentMonthYm_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (typeof resolveNippoTargetYearMonth_ === "function") {
    const ym = resolveNippoTargetYearMonth_(ss);
    if (ym && ym.year) return { year: ym.year, month: ym.month };
  }
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() };
}

function formatMonthlySheetName_(year, month) {
  const yy = String(year % 100).padStart(2, "0");
  const mm = String(month + 1).padStart(2, "0");
  return yy + mm;
}

function countMembersByDay_(rows, year, month, category) {
  const counts = {};
  const seen = {};
  for (let d = 1; d <= DAILY_COUNT_DAYS; d++) counts[d] = 0;
  rows.forEach(function (r) {
    if (category && !matchesEnrollCategory_(r[3], category)) return;
    // 入会月（ご利用開始日ベース）で当月判定。受信日ズレでも月を誤らない
    const ym = parseYearMonthFromLabel_(r[2]);
    if (!ym || ym.year !== year || ym.month !== month) return;
    const name = normalizeMemberName_(r[1]);
    const uniq = name + "\x1f" + String(r[3] || "").trim() + "\x1f" + normalizeYearMonthLabel_(r[2]);
    if (!name || seen[uniq]) return;
    seen[uniq] = true;
    const ts = parseMemberTimestamp_(r[0]);
    let day = 1;
    if (ts && ts.getFullYear() === year && ts.getMonth() === month) {
      day = ts.getDate();
    }
    if (day >= 1 && day <= DAILY_COUNT_DAYS) counts[day]++;
  });
  return counts;
}

function matchesEnrollCategory_(value, category) {
  const cat = String(value || "").trim();
  if (cat === category) return true;
  if (category === CATEGORY_SIX_MONTH && cat === "一般会員") return true;
  if (category === CATEGORY_CORPORATE && cat === "法人会員") return true;
  return false;
}

/** 一般退会は退会月＋受信日で当月の日別欄へ計上。法人退会はF71へ分離する。 */
function countWithdrawalsByDay_(rows, year, month) {
  const counts = {};
  const seen = {};
  for (let d = 1; d <= DAILY_COUNT_DAYS; d++) counts[d] = 0;
  const cancelMap = loadWithdrawCancelMapIfAvailable_();
  rows.forEach(function (r) {
    if (isWithdrawalCancelled_(r, cancelMap)) return;
    if (String(r[3] || "").trim() === CATEGORY_CORPORATE) return;
    const ym = parseYearMonthFromLabel_(r[2]);
    if (!ym || ym.year !== year || ym.month !== month) return;
    const name = normalizeMemberName_(r[1]);
    const uniq = name + "\x1f" + normalizeYearMonthLabel_(r[2]);
    if (!name || seen[uniq]) return;
    seen[uniq] = true;
    const ts = parseMemberTimestamp_(r[0]);
    let day = 1;
    if (ts && ts.getFullYear() === year && ts.getMonth() === month) {
      day = ts.getDate();
    }
    if (day >= 1 && day <= DAILY_COUNT_DAYS) counts[day]++;
  });
  return counts;
}

/** 法人退会（翌月末）は、メールから算出した退会月の F71へ計上 */
function countCorporateWithdrawAdvance_(rows, year, month) {
  const cancelMap = loadWithdrawCancelMapIfAvailable_();
  const seen = {};
  let count = 0;
  rows.forEach(function (r) {
    if (isWithdrawalCancelled_(r, cancelMap)) return;
    if (String(r[3] || "").trim() !== CATEGORY_CORPORATE) return;
    const ym = parseYearMonthFromLabel_(r[2]);
    if (!ym) return;
    if (ym.year !== year || ym.month !== month) return;
    const name = normalizeMemberName_(r[1]);
    const uniq = name + "\x1f" + normalizeYearMonthLabel_(r[2]);
    if (!name || seen[uniq]) return;
    seen[uniq] = true;
    count++;
  });
  return count;
}

function parseYearMonthFromLabel_(label) {
  const normalized = normalizeYearMonthLabel_(label);
  const m = String(normalized).match(/^(\d{4})年(\d{1,2})月$/);
  if (!m) return null;
  const month = parseInt(m[2], 10) - 1;
  if (month < 0 || month > 11) return null;
  return { year: parseInt(m[1], 10), month: month };
}

function parseMemberTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 30000) {
    const base = new Date(1899, 11, 30);
    const d = new Date(base.getTime() + Math.floor(value) * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(value || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ── 旧レイアウト移行 ──

function migrateLegacyLayouts_(ss) {
  setupDataSheet_(ss);
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (readEnrollData_(dataSheet).length > 0 || readWithdrawData_(dataSheet).length > 0) return;

  const display = ss.getSheetByName(SHEET_NAME_MEMBERS);
  if (!display) return;

  const scanEnd = Math.min(display.getLastRow(), DISPLAY_START_ROW + DISPLAY_MAX_SCAN - 1);
  if (scanEnd < DISPLAY_START_ROW) return;

  const numRows = scanEnd - DISPLAY_START_ROW + 1;
  const read5 = function (col) {
    return display.getRange(DISPLAY_START_ROW, col, numRows, 5).getValues()
      .filter(function (r) { return r[0] || r[1]; })
      .map(function (r) { return normalizeMemberRow_(r, DATA_ENROLL_COLS); });
  };

  const enroll = read5(11).length > 0 ? read5(11) : read5(1);
  let withdraw = read5(16).length > 0 ? read5(16) : read5(6);
  if (!withdraw.length && enroll.length && !read5(16).length && !read5(6).length) {
    withdraw = enroll;
    enroll.length = 0;
  }

  if (enroll.length) writeEnrollData_(ss, dedupeRows_(enroll, DATA_ENROLL_COLS));
  if (withdraw.length) writeWithdrawData_(ss, dedupeRows_(withdraw, DATA_WITHDRAW_COLS));

  removeSheetFilter_(display);
  display.getRange(DISPLAY_START_ROW, 1, numRows, 20).clearContent();

  ["退会", "退会データ"].forEach(function (name) {
    const legacy = ss.getSheetByName(name);
    if (legacy) ss.deleteSheet(legacy);
  });
}

// ── 入会メール解析 ──

function parseNyukaiMessage_(message) {
  const subject = String(message.getSubject() || "");
  if (subject.indexOf("ご入会") === -1) return null;

  const body = getMessageBodyText_(message);
  const date = message.getDate();
  const msgId = message.getId();
  const name = extractNyukaiName_(body);
  const ym = calcNyukaiYm_(body, date);
  if (!name || !ym) return null;

  return [date, name, ym.label, detectNyukaiCategory_(subject, body), msgId];
}

function extractNyukaiName_(body) {
  const text = String(body || "");
  const m1 = text.match(/お名前\s*[:：]\s*(.+?)\s*様/);
  if (m1) return String(m1[1]).replace(/^[>\s]+/, "").trim();

  const lines = text.split(/\r?\n/);
  for (let i = 0; i < Math.min(lines.length, 20); i++) {
    const line = String(lines[i] || "").trim();
    if (!line) continue;
    if (/JOYFIT|受付番号|この度は|お支払い|ご利用開始|ご入会内容|会員情報/.test(line)) continue;
    const m = line.match(/^(.{1,30}?)\s*様\s*$/);
    if (m) return String(m[1]).replace(/^[>\s]+/, "").trim();
  }
  return "";
}

function calcNyukaiYm_(body, emailDate) {
  const startMatch = String(body || "").match(/ご利用開始日[の]?\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (startMatch) {
    return { label: normalizeYearMonthLabel_(startMatch[1] + "年" + parseInt(startMatch[2], 10) + "月") };
  }
  const monthMatch = String(body || "").match(/[（(](\d{1,2})月分[）)]/);
  if (monthMatch) {
    return { label: normalizeYearMonthLabel_(emailDate.getFullYear() + "年" + monthMatch[1] + "月") };
  }
  return { label: normalizeYearMonthLabel_(emailDate.getFullYear() + "年" + (emailDate.getMonth() + 1) + "月") };
}

function detectNyukaiCategory_(subject, body) {
  const subjectText = String(subject || "");
  const bodyText = String(body || "");

  // 1. お支払い／内訳の会費行に「法人」→ 法人（いちばん確実）
  if (hasCorporateFeeLine_(bodyText)) return CATEGORY_CORPORATE;

  // 2. 件名に店舗タグ → 一般（6ヶ月割含む）
  if (subjectText.indexOf("【JOYFIT24経堂】") !== -1) return CATEGORY_SIX_MONTH;

  // 3. ナショナル会員など → 一般
  if (/ナショナル会員/.test(bodyText)) return CATEGORY_SIX_MONTH;

  // 4. 件名が「ご入会ありがとうございます！」系で店舗タグなし → 法人寄り
  if (/ご入会ありがとうございます/.test(subjectText) && subjectText.indexOf("【JOYFIT24経堂】") === -1) {
    return CATEGORY_CORPORATE;
  }

  if (/法人個人月払|法人会員/.test(bodyText)) return CATEGORY_CORPORATE;
  return "不明";
}

/** 会費行に法人プランがあるか（法人個人月払B など） */
function hasCorporateFeeLine_(body) {
  const lines = String(body || "").split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = String(lines[i] || "");
    if (line.indexOf("法人") === -1) continue;
    if (/法人個人月払|法人会員|法人.*月払/.test(line)) return true;
  }
  return false;
}

function getThreadMessagesNewestFirst_(thread) {
  const messages = thread.getMessages();
  const out = [];
  for (let i = messages.length - 1; i >= 0; i--) out.push(messages[i]);
  return out;
}

function extractNyukaiRowsFromThreads_(threads) {
  const rows = [];
  const seen = createMemberDedupeState_();
  const existingIds = {};
  (threads || []).forEach(function (thread) {
    collectNewMembershipFromThread_(
      thread, existingIds, seen, rows, DATA_ENROLL_COLS, parseNyukaiMessage_
    );
  });
  return rows;
}

// ── 退会メール解析 ──

function parseTaikaiMessage_(message) {
  const subject = String(message.getSubject() || "");
  if (subject.indexOf("ご退会") === -1) return null;

  const body = getMessageBodyText_(message);
  const date = message.getDate();
  const msgId = message.getId();
  const name = extractTaikaiName_(body) || extractTaikaiNameFromSubject_(subject);
  const finalText = resolveTaikaiFinalDateText_(subject, body);
  const ym = calcTaikaiWithdrawalYm_(date, finalText);
  if (!ym || !name) return null;

  return [date, name, ym.label, detectTaikaiCategory_(subject, body), msgId];
}

function extractTaikaiNameFromSubject_(subject) {
  const m = String(subject || "").match(/[-－—]\s*(.+?)\s*様/);
  if (!m) return "";
  return String(m[1]).replace(/^[>\s]+/, "").trim();
}

function resolveTaikaiFinalDateText_(subject, body) {
  const text = String(body || "");
  const finalMatch = text.match(/【最終ご利用日】[：:\s]*([^\r\n]+)/);
  if (finalMatch) return String(finalMatch[1]).trim();
  if (String(subject || "").indexOf("【JOYFIT24経堂】") !== -1) {
    if (text.indexOf("当月末") !== -1) return "当月末";
    if (text.indexOf("翌月末") !== -1) return "翌月末";
  }
  return "";
}

function extractTaikaiName_(body) {
  const text = String(body || "");
  const m1 = text.match(/お名前\s*[:：]\s*(.+?)\s*様/);
  if (m1) return String(m1[1]).replace(/^[>\s]+/, "").trim();

  const lines = text.split(/\r?\n/);
  for (let i = 0; i < Math.min(lines.length, 15); i++) {
    const line = String(lines[i] || "").trim();
    if (!line || /JOYFIT|ご退会|日頃より|以下情報|会員情報/.test(line)) continue;
    const m = line.match(/^(.{1,30}?)\s*様\s*$/);
    if (m) return String(m[1]).replace(/^[>\s]+/, "").trim();
  }
  return "";
}

function calcTaikaiWithdrawalYm_(emailDate, finalDateText) {
  let y = emailDate.getFullYear();
  let m = emailDate.getMonth();
  if (finalDateText.indexOf("翌月末") !== -1) {
    m += 1;
    if (m > 11) { m = 0; y += 1; }
  } else if (finalDateText.indexOf("当月末") !== -1) {
    // 受信月が退会月
  } else {
    return null;
  }
  return { label: normalizeYearMonthLabel_(y + "年" + (m + 1) + "月") };
}

function detectTaikaiCategory_(subject, body) {
  if (subject.indexOf("【JOYFIT24経堂】") !== -1) return "一般会員";
  if (body.indexOf("ご退会申請の受付") !== -1) return "一般会員";
  if (body.indexOf("ご退会のお手続きが完了") !== -1) return "法人会員";
  return "不明";
}

function extractTaikaiRowsFromThreads_(threads) {
  const rows = [];
  const seen = createMemberDedupeState_();
  const existingIds = {};
  (threads || []).forEach(function (thread) {
    collectNewMembershipFromThread_(
      thread, existingIds, seen, rows, DATA_WITHDRAW_COLS, parseTaikaiMessage_
    );
  });
  return rows;
}

// ── Gmailラベル ──

function installMembershipLabelsSilent_() {
  [LABEL_NYUKAI_GENERAL, LABEL_NYUKAI_CORPORATE, LABEL_TAIKAI_GENERAL, LABEL_TAIKAI_CORPORATE]
    .forEach(function (name) { getOrCreateGmailLabel_(name); });
}

function applyNyukaiLabelsToThreads_(threads) {
  const labelGeneral = getOrCreateGmailLabel_(LABEL_NYUKAI_GENERAL);
  const labelCorporate = getOrCreateGmailLabel_(LABEL_NYUKAI_CORPORATE);
  threads.forEach(function (thread) {
    const categories = {};
    thread.getMessages().forEach(function (message) {
      if (String(message.getSubject() || "").indexOf("ご入会") === -1) return;
      const cat = detectNyukaiCategory_(message.getSubject(), getMessageBodyText_(message));
      categories[cat] = true;
    });
    thread.removeLabel(labelGeneral);
    thread.removeLabel(labelCorporate);
    if (categories[CATEGORY_SIX_MONTH]) thread.addLabel(labelGeneral);
    if (categories[CATEGORY_CORPORATE]) thread.addLabel(labelCorporate);
  });
}

function applyTaikaiLabelsToThreads_(threads) {
  const labelGeneral = getOrCreateGmailLabel_(LABEL_TAIKAI_GENERAL);
  const labelCorporate = getOrCreateGmailLabel_(LABEL_TAIKAI_CORPORATE);
  threads.forEach(function (thread) {
    const categories = {};
    thread.getMessages().forEach(function (message) {
      if (String(message.getSubject() || "").indexOf("ご退会") === -1) return;
      const cat = detectTaikaiCategory_(message.getSubject(), getMessageBodyText_(message));
      categories[cat] = true;
    });
    thread.removeLabel(labelGeneral);
    thread.removeLabel(labelCorporate);
    if (categories["一般会員"]) thread.addLabel(labelGeneral);
    if (categories["法人会員"]) thread.addLabel(labelCorporate);
  });
}

function getOrCreateGmailLabel_(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

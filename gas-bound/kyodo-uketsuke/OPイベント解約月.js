/**
 * 【一時】OPイベント者シートの解約月(J)埋め
 *
 * シート「OPイベント者」:
 *   A 氏名 / C〜F オプション(チェック) / H 適用月(触らない) / J 解約月 ← ここを埋める
 *
 * 判定:
 *   OP集計の契約明細で「利用停止」＝オプション解約メールが来ている
 *   → その受信月を J に「4月」形式で書く（2026年対象）
 *
 * 使い方:
 *   1. このファイルを GAS に追加
 *   2. スプレッドシートを開き直す → メニュー「OPイベント（一時）」
 *   3. 「①2026年OP明細を取ってから解約月を埋める」を実行（推奨）
 * 終わったらこのファイルとメニュー呼び出しを削除してOK
 */

const OP_EVENT_SHEET_NAME = "OPイベント者";
const OP_EVENT_TARGET_YEAR = 2026;
const OP_EVENT_NAME_COL = 1;       // A
const OP_EVENT_OPT_START_COL = 3;  // C
const OP_EVENT_OPT_END_COL = 6;    // F
const OP_EVENT_CANCEL_FLAG_COL = 9; // I 解約
const OP_EVENT_CANCEL_MONTH_COL = 10; // J 解約月

function installOpEventTempMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("OPイベント（一時）")
    .addItem("①2026年OP明細を取ってから解約月を埋める", "fillOpEventCancelMonthsWithFetch")
    .addItem("②既存OP明細だけで解約月を埋める", "fillOpEventCancelMonthsFromLogOnly")
    .addToUi();
}

/** 推奨: 2026年4月〜当月のOPメールを取り直してから J を埋める */
function fillOpEventCancelMonthsWithFetch() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    ui.alert("OPイベント", "別の処理が実行中です。少し待ってから再度お試しください。", ui.ButtonSet.OK);
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSpreadsheet();
    const opSheet = getOpSheet_(ss);
    const endMonth = resolveOpEventFetchEndMonth_(); // 0-based
    let fetched = 0;
    for (let m = 3; m <= endMonth; m++) { // 4月=3 …
      executeFetchMonthForYm_(ss, opSheet, OP_EVENT_TARGET_YEAR, m, true);
      fetched++;
    }
    const result = fillOpEventCancelMonthsCore_(ss);
    ui.alert(
      "OPイベント",
      "完了しました。\n\n" +
        "OP明細再取得 … " + OP_EVENT_TARGET_YEAR + "年4月〜" + (endMonth + 1) + "月（" + fetched + "ヶ月）\n" +
        result.message,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("OPイベントエラー", String(err && err.message ? err.message : err), ui.ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

/** すでにOP明細がある前提で、Jだけ埋める */
function fillOpEventCancelMonthsFromLogOnly() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    ui.alert("OPイベント", "別の処理が実行中です。少し待ってから再度お試しください。", ui.ButtonSet.OK);
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = fillOpEventCancelMonthsCore_(ss);
    ui.alert("OPイベント", result.message, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert("OPイベントエラー", String(err && err.message ? err.message : err), ui.ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

function resolveOpEventFetchEndMonth_() {
  const now = new Date();
  if (now.getFullYear() > OP_EVENT_TARGET_YEAR) return 11;
  if (now.getFullYear() < OP_EVENT_TARGET_YEAR) return 3;
  return Math.max(3, now.getMonth());
}

/**
 * OPイベント者の各行について、
 * 氏名(A)＋チェックされたオプション(C〜F) に一致する「利用停止」をOP明細から探し、
 * 解約月(J)へ「n月」を書く。
 */
function fillOpEventCancelMonthsCore_(ss) {
  const sheet = ss.getSheetByName(OP_EVENT_SHEET_NAME);
  if (!sheet) {
    throw new Error("シート「" + OP_EVENT_SHEET_NAME + "」が見つかりません");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { filled: 0, checked: 0, message: "データ行がありません。" };
  }

  const headerOpts = sheet
    .getRange(1, OP_EVENT_OPT_START_COL, 1, OP_EVENT_OPT_END_COL - OP_EVENT_OPT_START_COL + 1)
    .getDisplayValues()[0];
  const optKeys = headerOpts.map(function (h) {
    return mapOpEventHeaderToOptionKey_(h);
  });

  const numRows = lastRow - 1;
  const numOptCols = OP_EVENT_OPT_END_COL - OP_EVENT_OPT_START_COL + 1;
  // getRange(row, column, numRows, numColumns) ※第3引数は「行数」であって最終行番号ではない
  const names = sheet.getRange(2, OP_EVENT_NAME_COL, numRows, 1).getDisplayValues();
  const checks = sheet
    .getRange(2, OP_EVENT_OPT_START_COL, numRows, numOptCols)
    .getValues();

  const stopIndex = buildOpEventStopIndex_(ss);
  // 既存の I / J はヒットした行だけ更新（未ヒットで手入力を消さない）
  const existingJ = sheet.getRange(2, OP_EVENT_CANCEL_MONTH_COL, numRows, 1).getValues();
  const existingI = sheet.getRange(2, OP_EVENT_CANCEL_FLAG_COL, numRows, 1).getValues();
  const outMonths = [];
  const outFlags = [];
  let filled = 0;
  let withOpt = 0;
  let noHit = 0;

  for (let i = 0; i < numRows; i++) {
    const nameKey = normalizeOpEventPersonName_(names[i][0]);
    const checkedOpts = [];
    for (let c = 0; c < optKeys.length; c++) {
      if (isOpEventChecked_(checks[i][c]) && optKeys[c]) {
        checkedOpts.push(optKeys[c]);
      }
    }

    if (!nameKey || !checkedOpts.length) {
      outMonths.push([existingJ[i][0]]);
      outFlags.push([existingI[i][0]]);
      continue;
    }
    withOpt++;

    const months = [];
    const seenM = {};
    for (let o = 0; o < checkedOpts.length; o++) {
      const stops = stopIndex[nameKey + "\t" + checkedOpts[o]] || [];
      for (let s = 0; s < stops.length; s++) {
        const label = stops[s].monthLabel;
        if (!seenM[label]) {
          seenM[label] = true;
          months.push({ label: label, time: stops[s].time });
        }
      }
    }

    if (!months.length) {
      outMonths.push([existingJ[i][0]]);
      outFlags.push([existingI[i][0]]);
      noHit++;
      continue;
    }

    // 複数の利用停止があっても、いちばん新しい月だけを書く（例: 1月と6月 → 6月）
    months.sort(function (a, b) { return b.time - a.time; });
    outMonths.push([months[0].label]);
    outFlags.push([true]);
    filled++;
  }

  sheet.getRange(2, OP_EVENT_CANCEL_MONTH_COL, numRows, 1).setValues(outMonths);
  sheet.getRange(2, OP_EVENT_CANCEL_FLAG_COL, numRows, 1).setValues(outFlags);

  return {
    filled: filled,
    checked: withOpt,
    message:
      "対象シート … " + OP_EVENT_SHEET_NAME + "\n" +
      "年 … " + OP_EVENT_TARGET_YEAR + "\n" +
      "オプション指定あり … " + withOpt + " 行\n" +
      "解約月を記入 … " + filled + " 行\n" +
      "明細に利用停止なし … " + noHit + " 行\n\n" +
      "※ H（適用月）は変更していません\n" +
      "※ J は最新の解約月のみ（例: 1月と6月なら 6月）"
  };
}

/** ヘッダー表記 → OP集計用のオプション名 */
function mapOpEventHeaderToOptionKey_(header) {
  const h = String(header || "").trim();
  if (!h) return "";
  // シート表記のゆれを先に吸収
  if (h.indexOf("プロテイン") !== -1 && h.indexOf("水素") !== -1) return "プロテイン＋水素水";
  if (h.indexOf("タンニング") !== -1) return "タンニング";
  if (h.indexOf("セルフエステ") !== -1) return "セルフエステ";
  if (h.indexOf("ボディプランナー") !== -1 || h.indexOf("ボディープランナー") !== -1) {
    return "体組成計"; // normalizeOptionName と同じ対応
  }
  if (typeof normalizeOptionName === "function") {
    return normalizeOptionName(h);
  }
  return h;
}

function isOpEventChecked_(v) {
  if (v === true || v === "TRUE" || v === "true") return true;
  if (v === false || v === "" || v === null) return false;
  if (typeof v === "string") {
    const s = v.trim().toLowerCase();
    return s === "✓" || s === "✔" || s === "1" || s === "on" || s === "yes";
  }
  return !!v;
}

function normalizeOpEventPersonName_(name) {
  return String(name || "")
    .replace(/[ 　\t\r\n]/g, "")
    .replace(/様$/g, "")
    .trim();
}

/**
 * 2026年の「利用停止」を 氏名×オプション → [{monthLabel, time}] で索引化
 */
function buildOpEventStopIndex_(ss) {
  const opSheet = getOpSheet_(ss);
  const rows = readOpLogValues_(opSheet);
  const index = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const cat = String(row[2] || "").trim();
    if (cat.indexOf("利用停止") === -1) continue;

    const d = parseOpLogDate_(row[0]);
    if (!d || d.getFullYear() !== OP_EVENT_TARGET_YEAR) continue;

    const nameKey = normalizeOpEventPersonName_(row[1]);
    const optKey = String(row[4] || "").trim() ||
      (typeof normalizeOptionName === "function" ? normalizeOptionName(row[3]) : String(row[3] || "").trim());
    if (!nameKey || !optKey) continue;

    const key = nameKey + "\t" + optKey;
    if (!index[key]) index[key] = [];
    index[key].push({
      monthLabel: (d.getMonth() + 1) + "月",
      time: d.getTime()
    });
  }
  return index;
}

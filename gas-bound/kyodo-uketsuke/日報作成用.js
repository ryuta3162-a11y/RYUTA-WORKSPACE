/**
 * JOYFIT24経堂 — 日報メール
 *
 * createShopDailyReportDraft … 手動：下書き作成（通常は使わない）
 * sendShopDailyReportSilent_ … 送信本体
 * runDailyUpdateAndSendAt20_ / At21_ … 土日祝20時 / 平日21時（更新→送信）
 */

const NIPPO_MAIL_RECIPIENTS = [
  "m-harada@okamoto-group.co.jp",
  "mito-sato@okamoto-group.co.jp",
  "r-kusaka@okamoto-group.co.jp",
  "h-nakata@okamoto-group.co.jp",
  "s-kurokawa@okamoto-group.co.jp",
  "ka-yoshida@okamoto-group.co.jp",
  "m-tokushige@okamoto-group.co.jp",
  "k-moriyasu@okamoto-group.co.jp",
  "k-ishibashi@okamoto-group.co.jp",
  "m-osari@okamoto-group.co.jp",
  "yuka-hachiya@okamoto-group.co.jp"
];

// NIPPO_MAIL_FROM … オプション出力.gs で定義（重複 const だと全スクリプトが動かない）
const NIPPO_SHEET_NAME = "日報";
const NIPPO_RANGE_START_ROW = 8;
const NIPPO_RANGE_START_COL = 2;
const NIPPO_RANGE_ROWS = 33;
const NIPPO_RANGE_COLS = 8;

const JAPAN_HOLIDAY_CALENDAR_ID = "ja.japanese#holiday@group.v8.calendar.google.com";
const SENT_REPORT_PROP_PREFIX = "dailyReportSent_";

/** 手動：下書き作成 */
function createShopDailyReportDraft() {
  const mail = buildShopDailyReportMail_();
  if (!mail) return;
  GmailApp.createDraft(mail.recipient, mail.subject, "", mail.options);
  Logger.log("店舗日報の下書きを作成しました: " + mail.subject);
}

/** @deprecated 旧名。土日祝20時ラッパーと同じ */
function sendDailyReportAt20_() {
  if (typeof runDailyUpdateAndSendAt20_ === "function") {
    runDailyUpdateAndSendAt20_();
  }
}

/** @deprecated 旧名。平日21時ラッパーと同じ */
function sendDailyReportAt21_() {
  if (typeof runDailyUpdateAndSendAt21_ === "function") {
    runDailyUpdateAndSendAt21_();
  } else if (typeof runDailyUpdateAndSendSilent_ === "function") {
    runDailyUpdateAndSendSilent_();
  } else {
    sendShopDailyReportSilent_();
  }
}

/** 自動送信（ダイアログなし・同日二重送信防止） */
function sendShopDailyReportSilent_() {
  const tz = Session.getScriptTimeZone();
  const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  if (wasDailyReportSentToday_(todayKey)) {
    Logger.log("日報メールは本日送信済みのためスキップ: " + todayKey);
    return;
  }

  const mail = buildShopDailyReportMail_();
  if (!mail) return;

  GmailApp.sendEmail(mail.recipient, mail.subject, "", mail.options);
  markDailyReportSentToday_(todayKey);
  Logger.log("店舗日報を送信しました: " + mail.subject);
}

function buildShopDailyReportMail_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log("スプレッドシートを取得できませんでした");
    return null;
  }

  const sheet = ss.getSheetByName(NIPPO_SHEET_NAME);
  if (!sheet) {
    Logger.log("「日報」シートが見つかりません");
    return null;
  }

  const startRow = NIPPO_RANGE_START_ROW;
  const startCol = NIPPO_RANGE_START_COL;
  const numRows = NIPPO_RANGE_ROWS;
  const numCols = NIPPO_RANGE_COLS;

  const formattedDate = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), "yyyy.MM.dd"
  );
  const subject = "JF24経堂　報連相 " + formattedDate;

  const range = sheet.getRange(startRow, startCol, numRows, numCols);
  const values = range.getDisplayValues();
  const fontColors = range.getFontColors();
  const backgrounds = range.getBackgrounds();
  const hAligns = range.getHorizontalAlignments();
  const vAligns = range.getVerticalAlignments();

  const mergeMap = buildMergeMap_(range, startRow, startCol, numRows, numCols);
  const htmlBody = buildNippoHtmlBody_(values, fontColors, backgrounds, hAligns, vAligns, mergeMap, numRows, numCols);

  const options = { htmlBody: htmlBody, cc: "" };
  if (NIPPO_MAIL_FROM) options.from = NIPPO_MAIL_FROM;

  return {
    recipient: NIPPO_MAIL_RECIPIENTS.join(","),
    subject: subject,
    options: options
  };
}

function buildMergeMap_(range, startRow, startCol, numRows, numCols) {
  const mergeMap = [];
  for (let r = 0; r < numRows; r++) {
    mergeMap[r] = [];
    for (let c = 0; c < numCols; c++) {
      mergeMap[r][c] = { skip: false, rowSpan: 1, colSpan: 1 };
    }
  }

  const merges = range.getMergedRanges();
  for (let i = 0; i < merges.length; i++) {
    const rng = merges[i];
    const relativeRow = rng.getRow() - startRow;
    const relativeCol = rng.getColumn() - startCol;
    const mNumRows = rng.getNumRows();
    const mNumCols = rng.getNumColumns();

    if (relativeRow < 0 || relativeRow >= numRows || relativeCol < 0 || relativeCol >= numCols) continue;

    mergeMap[relativeRow][relativeCol].rowSpan = mNumRows;
    mergeMap[relativeRow][relativeCol].colSpan = mNumCols;
    for (let r = 0; r < mNumRows; r++) {
      for (let c = 0; c < mNumCols; c++) {
        if (r === 0 && c === 0) continue;
        if (relativeRow + r < numRows && relativeCol + c < numCols) {
          mergeMap[relativeRow + r][relativeCol + c].skip = true;
        }
      }
    }
  }
  return mergeMap;
}

function buildNippoHtmlBody_(values, fontColors, backgrounds, hAligns, vAligns, mergeMap, numRows, numCols) {
  let htmlBody = [
    '<div style="font-family: sans-serif; font-size: 10pt; color: #333;">',
    "<p>お元気様です。<br>本日の業務日報です。<br>よろしくお願いします。</p>",
    '<table border="0" cellspacing="0" cellpadding="0" style="',
    "border-collapse: collapse; border: none; font-size: 10pt;",
    "line-height: 1.3; font-weight: bold; margin-bottom: 20px;\">"
  ].join("");

  let started = false;
  for (let i = 0; i < numRows; i++) {
    // 先頭だけの完全空白行は出さない（旧仕様のグレー帯の原因）
    if (!started && isNippoMailRowEmpty_(values[i])) continue;
    started = true;

    htmlBody += "<tr>";
    for (let j = 0; j < numCols; j++) {
      if (mergeMap[i][j].skip) continue;

      let val = values[i][j];
      if (val === "") val = "&nbsp;";

      const bg = normalizeNippoMailBg_(backgrounds[i][j]);
      const color = normalizeNippoMailFg_(fontColors[i][j]);
      let hAlign = hAligns[i][j];
      const vAlign = vAligns[i][j];
      if (j === 0) hAlign = "left";

      let spanAttr = "";
      if (mergeMap[i][j].rowSpan > 1) spanAttr += ' rowspan="' + mergeMap[i][j].rowSpan + '"';
      if (mergeMap[i][j].colSpan > 1) spanAttr += ' colspan="' + mergeMap[i][j].colSpan + '"';

      htmlBody += "<td" + spanAttr + ' style="border: none; padding: 4px 6px; background-color: ' + bg +
        "; color: " + color + "; font-weight: bold; text-align: " + hAlign +
        "; vertical-align: " + vAlign + '; font-size: 10pt; white-space: nowrap;">' + val + "</td>";
    }
    htmlBody += "</tr>";
  }

  htmlBody += [
    "</table><br><div>",
    "☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆<br>",
    "株式会社ヤマウチ<br>",
    "スポーツクラブJOYFIT24経堂<br>",
    "〒156-0052<br>",
    "東京都世田谷区経堂5-23-13<br>",
    "TEL：03-6804-4100<br>",
    "スタッフ常駐時間：平日　10：00～21：00<br>",
    "　　　　　　　　　土日祝　12：00～20：00<br>",
    "※毎週月曜日・木曜日は終日スタッフ不在でございます。<br>",
    "☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆<br>",
    "</div></div>"
  ].join("");

  return htmlBody;
}

/** 旧仕様で先頭行を #444444 に塗り替えていたのが、上のグレー帯の原因 */
function isNippoMailRowEmpty_(rowValues) {
  if (!rowValues || !rowValues.length) return true;
  for (let i = 0; i < rowValues.length; i++) {
    if (String(rowValues[i] || "").trim() !== "") return false;
  }
  return true;
}

function normalizeNippoMailBg_(bg) {
  const s = String(bg || "").trim().toLowerCase();
  if (!s || s === "#ffffff" || s === "#fff" || s === "white" || s === "transparent") {
    return "#ffffff";
  }
  return s;
}

function normalizeNippoMailFg_(fg) {
  const s = String(fg || "").trim().toLowerCase();
  if (!s || s === "#000000" || s === "#000" || s === "black") return "#111111";
  return s;
}

function isWeekendOrJapaneseHoliday_(date) {
  const day = date.getDay();
  if (day === 0 || day === 6) return true;
  return isJapanesePublicHoliday_(date);
}

function isJapanesePublicHoliday_(date) {
  try {
    const cal = CalendarApp.getCalendarById(JAPAN_HOLIDAY_CALENDAR_ID);
    if (!cal) return false;
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const end = new Date(start.getTime() + 86400000);
    return cal.getEvents(start, end).length > 0;
  } catch (err) {
    Logger.log("祝日カレンダー参照エラー（土日のみ判定）: " + err.message);
    return false;
  }
}

function wasDailyReportSentToday_(todayKey) {
  return PropertiesService.getScriptProperties().getProperty(SENT_REPORT_PROP_PREFIX + todayKey) === "1";
}

function markDailyReportSentToday_(todayKey) {
  PropertiesService.getScriptProperties().setProperty(SENT_REPORT_PROP_PREFIX + todayKey, "1");
}

/** @deprecated 送信専用トリガーは廃止。installOptionDailyTrigger が一体型を設定する */
function installDailyReportSendTriggers_() {
  removeDailyReportSendTriggers_();
}

function removeDailyReportSendTriggers_() {
  // 日報送信は runDailyUpdateAndSendAt20_/21_ に一本化。誤って消さない。
}

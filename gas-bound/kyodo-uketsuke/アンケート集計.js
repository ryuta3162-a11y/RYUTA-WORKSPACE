/**
 * JOYFIT24経堂 — 退会アンケート集計
 *
 * 【アンケート_退会】 CSV取込のマスタ（uid×q_id で重複なし・追記マージ）
 * 【アンケート】       月別サマリー・明細（退会メールと氏名で突合）
 *
 * 同姓同名の再入会・再退会は uid / 退会メールを別件として1:1で順番突合
 */

const SHEET_NAME_ANKETO = "アンケート管理";
const SHEET_NAME_ANKETO_LEGACY = "アンケート";
const SHEET_NAME_ANKETO_RAW = "アンケート_退会";
const SHEET_NAME_ANKETO_PASTE = "アンケート_取込";
const SHEET_NAME_ANKETO_UPLOAD = "アンケート_送信";
const ANKETO_UPLOAD_BATCH_ROWS = 150;

/** サマリー配色（JOYFIT調・赤・黒・グレー・白） */
const ANKETO_THEME = {
  black: "#111111",
  white: "#ffffff",
  grayDark: "#5f6368",
  grayMid: "#9aa0a6",
  grayLight: "#f1f3f4",
  grayPale: "#f5f5f5",
  redPale: "#fde8ec",
  redSoft: "#f8cfd8",
  red: "#C21632",
  redDark: "#8f0f25",
  headerBg: "#111111",
};

const ANKETO_RAW_HEADERS = ["shop_id", "uid", "name", "q_id", "answer"];
const ANKETO_Q_SKIP = { "a_back-page": true, "6": true };

const ANKETO_DETAIL_HEADERS = [
  "退会月", "受信日時", "氏名", "uid", "利用頻度", "利用日数", "利用時間帯", "退会理由", "区分", "確認"
];

// ── メニュー ──

/** JOYFITメニュー: CSV取込ダイアログ（ブラウザでCSV解析→行バッチ送信） */
function importAnketoCsvMenu_() {
  ensureAnketoSheets_(SpreadsheetApp.getActiveSpreadsheet());
  showAnketoCsvImportDialog_();
}

/** JOYFITメニュー: アンケート_取込 シートに貼ったCSVをマージ（確実・推奨） */
function importAnketoCsvFromPasteSheetMenu_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    const pack = importAnketoCsvFromPasteSheetCore_(ss);
    const res = importAnketoCsvFinishImport_(pack);
    ui.alert("退会アンケートを取込", res.message, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert("退会アンケートを取込", "エラー: " + err.message, ui.ButtonSet.OK);
  }
}

/** JOYFITメニュー: マスタ＋退会メールからアンケートシートを再集計 */
function updateAnketoSheetMenu_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    const result = syncAnketoDisplay_(ss);
    ui.alert(
      "アンケートを更新",
      "完了しました。\n\n" +
        "CSV行: " + result.rawRows + "\n" +
        "アンケート回答(uid): " + result.surveyCount + "\n" +
        "退会メール: " + result.withdrawCount + "\n" +
        "回答あり: " + result.linkedCount + "\n\n" +
        "表示月: " + result.monthLabel,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("アンケートを更新", "エラー: " + err.message, ui.ButtonSet.OK);
  }
}

/** 入会・退会更新・19:30自動処理から呼ぶ（ダイアログなし） */
function syncAnketoSilent_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const raw = ss.getSheetByName(SHEET_NAME_ANKETO_RAW);
    if (!raw || raw.getLastRow() < 2) return;
    syncAnketoDisplay_(ss);
  } catch (err) {
    Logger.log("アンケート集計: " + err.message);
  }
}

function handleAnketoSheetEdit_(e) {
  if (!e || !e.range) return false;
  if (e.range.getSheet().getName() !== SHEET_NAME_ANKETO) return false;
  if (e.range.getRow() !== 1 || e.range.getColumn() < 2 || e.range.getColumn() > 4) return false;
  syncAnketoDisplay_(e.source);
  return true;
}

// ── メイン ──

function syncAnketoDisplay_(ss) {
  ensureAnketoSheets_(ss);
  const rawRows = readAnketoRawRows_(ss);
  if (!rawRows.length) {
    throw new Error(
      "退会アンケートのデータがありません。\n" +
        "「アンケート_取込」にCSVを貼り、アンケート管理の「更新」を押してください。"
    );
  }

  const surveys = buildSurveyRecordsByUid_(rawRows);
  const withdraws = readActiveWithdrawRows_(ss);

  const monthList = buildAnketoMonthList_(withdraws);
  const sheet = ss.getSheetByName(SHEET_NAME_ANKETO);
  setupAnketoMonthSelector_(sheet, monthList);

  const selectedMonths = readAnketoSelectedMonths_(sheet, monthList);
  if (!selectedMonths.length) {
    throw new Error("比較する退会月を B1〜D1 から選んでください。");
  }

  const monthData = selectedMonths.map(function (label) {
    const monthRows = linkSurveyToWithdrawalsForMonth_(surveys, withdraws, label);
    const matched = monthRows.filter(function (r) { return r.survey; });
    return {
      label: label,
      monthRows: monthRows,
      matched: matched,
      meta: {
        withdrawMonthTotal: countWithdrawForMonth_(withdraws, label),
        linkedTotal: matched.length,
      },
      stats: {
        q12: collectAnketoAnswers_(matched, "q12"),
        q7: collectAnketoAnswers_(matched, "q7"),
        q11: expandAnketoSlotCounts_(collectAnketoAnswers_(matched, "q11")),
      },
    };
  });

  const primary = monthData[0];
  const linkedCount = primary.monthRows.filter(function (r) { return r.matchStatus === "回答あり"; }).length;

  renderAnketoCompareSheet_(sheet, monthData);
  applyAnketoJoyfitLayout_(ss);

  return {
    rawRows: rawRows.length,
    surveyCount: surveys.length,
    withdrawCount: withdraws.length,
    linkedCount: linkedCount,
    monthLabel: primary.label,
    monthDetail: primary.monthRows.length,
  };
}

// ── シート準備 ──

function ensureAnketoSheets_(ss) {
  let raw = ss.getSheetByName(SHEET_NAME_ANKETO_RAW);
  if (!raw) {
    raw = ss.insertSheet(SHEET_NAME_ANKETO_RAW);
    raw.getRange(1, 1, 1, ANKETO_RAW_HEADERS.length).setValues([ANKETO_RAW_HEADERS]);
    raw.setFrozenRows(1);
  }
  applyAnketoHeaderStyle_(raw, ANKETO_RAW_HEADERS.length);

  let sheet = ss.getSheetByName(SHEET_NAME_ANKETO);
  if (!sheet) {
    const legacy = ss.getSheetByName(SHEET_NAME_ANKETO_LEGACY);
    if (legacy) {
      legacy.setName(SHEET_NAME_ANKETO);
      sheet = legacy;
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_ANKETO);
  }

  let paste = ss.getSheetByName(SHEET_NAME_ANKETO_PASTE);
  if (!paste) {
    paste = ss.insertSheet(SHEET_NAME_ANKETO_PASTE);
    paste.getRange(1, 1, 1, ANKETO_RAW_HEADERS.length).setValues([ANKETO_RAW_HEADERS]);
    paste.setFrozenRows(1);
    paste.getRange("A1").setNote(
      "CSVをここに貼る（またはファイル→インポート）→ アンケート管理の「更新」"
    );
  }
  applyAnketoHeaderStyle_(paste, ANKETO_RAW_HEADERS.length);

  try {
    raw.hideSheet();
  } catch (e) {
    Logger.log("アンケート_退会 非表示: " + e.message);
  }

  ensureAnketoUploadBufferSheet_(ss);
  applyAnketoJoyfitLayout_(ss);
}

/** JOYFIT調ヘッダー（黒地・白文字・赤下線） */
function applyAnketoHeaderStyle_(sheet, numCols) {
  if (!sheet) return;
  const header = sheet.getRange(1, 1, 1, numCols);
  header.setFontWeight("bold")
    .setBackground(ANKETO_THEME.headerBg)
    .setFontColor(ANKETO_THEME.white)
    .setHorizontalAlignment("center")
    .setBorder(null, null, true, null, null, null, ANKETO_THEME.red, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/** アンケート関連シートの見た目を統一 */
function applyAnketoJoyfitLayout_(ss) {
  const paste = ss.getSheetByName(SHEET_NAME_ANKETO_PASTE);
  if (paste) applyAnketoHeaderStyle_(paste, ANKETO_RAW_HEADERS.length);

  const raw = ss.getSheetByName(SHEET_NAME_ANKETO_RAW);
  if (raw) applyAnketoHeaderStyle_(raw, ANKETO_RAW_HEADERS.length);

  const sheet = ss.getSheetByName(SHEET_NAME_ANKETO);
  if (!sheet) return;
  sheet.getRange("A1")
    .setFontWeight("bold")
    .setFontColor(ANKETO_THEME.grayDark)
    .setHorizontalAlignment("center");
  ["B1", "C1", "D1"].forEach(function (addr) {
    sheet.getRange(addr)
      .setBackground(ANKETO_THEME.grayLight)
      .setFontColor(ANKETO_THEME.black)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, null, null, ANKETO_THEME.red, SpreadsheetApp.BorderStyle.SOLID);
  });
  sheet.getRange("E1").setFontColor(ANKETO_THEME.grayMid).setFontSize(9);
}

function setupAnketoMonthSelector_(sheet, monthList) {
  sheet.getRange("A1")
    .setValue("比較する退会月 ➡")
    .setFontWeight("bold")
    .setFontColor(ANKETO_THEME.grayDark)
    .setHorizontalAlignment("center");
  sheet.getRange("E1")
    .clearContent()
    .setFontSize(9)
    .setFontColor(ANKETO_THEME.grayMid);

  ["B1", "C1", "D1"].forEach(function (addr, i) {
    const cell = sheet.getRange(addr);
    cell.clearDataValidations();
    cell.setNumberFormat("@");
    cell.setBackground(ANKETO_THEME.grayLight)
      .setFontColor(ANKETO_THEME.black)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, null, null, ANKETO_THEME.red, SpreadsheetApp.BorderStyle.SOLID);
    sheet.setColumnWidth(cell.getColumn(), 112);
    if (!monthList.length) return;

    const withBlank = [""].concat(monthList);
    cell.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(withBlank, true).build()
    );

    const cur = String(cell.getValue() || "").trim();
    if (cur && monthList.includes(cur)) return;

    if (addr === "B1") cell.setValue(monthList[0] || "");
    else if (addr === "C1") cell.setValue(monthList.length > 1 ? monthList[1] : "");
    else cell.setValue("");
  });
}

function readAnketoSelectedMonths_(sheet, monthList) {
  const out = [];
  ["B1", "C1", "D1"].forEach(function (addr) {
    const v = String(sheet.getRange(addr).getValue() || "").trim();
    if (!v || !monthList.includes(v)) return;
    if (out.indexOf(v) < 0) out.push(v);
  });
  return out;
}

// ── CSV（縦持ち）読み込み ──

function readAnketoRawRows_(ss) {
  const raw = ss.getSheetByName(SHEET_NAME_ANKETO_RAW);
  if (!raw || raw.getLastRow() < 2) return [];

  const values = raw.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const shop = String(row[0] || "").trim();
    const uid = String(row[1] || "").trim();
    const name = String(row[2] || "").trim();
    const qid = String(row[3] || "").trim();
    const ans = String(row[4] || "").trim();
    if (!uid || !qid) continue;
    if (ANKETO_Q_SKIP[qid]) continue;
    out.push({ shop: shop, uid: uid, name: name, qId: qid, answer: ans });
  }
  return out;
}

function readAnketoRowsFromValues_(values) {
  const out = [];
  if (!values || !values.length) return out;
  let start = 0;
  const h0 = String(values[0][0] || "").trim().toLowerCase();
  if (h0 === "shop_id" || h0.indexOf("shop") >= 0) start = 1;
  for (let i = start; i < values.length; i++) {
    const row = normalizeAnketoCsvRow_(values[i]);
    if (row) out.push(row);
  }
  return out;
}

/** 取込シート … 5列分割済み or A列にCSV1行ずつ の両方に対応 */
function readAnketoRowsFromPasteSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const hasMultiCol = values.some(function (row) {
    for (let c = 1; c < ANKETO_RAW_HEADERS.length; c++) {
      if (String(row[c] || "").trim()) return true;
    }
    return false;
  });

  if (hasMultiCol) {
    return readAnketoRowsFromValues_(values);
  }

  const lines = [];
  values.forEach(function (row) {
    const line = String(row[0] || "").trim();
    if (line) lines.push(line);
  });
  if (!lines.length) return [];
  return parseAnketoCsvText_(lines.join("\n"));
}

function buildSurveyRecordsByUid_(rawRows) {
  const byUid = {};
  rawRows.forEach(function (r) {
    if (!byUid[r.uid]) {
      byUid[r.uid] = {
        uid: r.uid,
        name: r.name,
        nameNorm: normalizeMemberName_(r.name),
        q7: "",
        q8: "",
        q11: "",
        q12: "",
      };
    }
    if (r.name && !byUid[r.uid].name) byUid[r.uid].name = r.name;
    if (r.qId === "7") byUid[r.uid].q7 = r.answer;
    else if (r.qId === "8") byUid[r.uid].q8 = r.answer;
    else if (r.qId === "11") byUid[r.uid].q11 = r.answer;
    else if (r.qId === "12") byUid[r.uid].q12 = r.answer;
  });

  return Object.keys(byUid).map(function (uid) { return byUid[uid]; })
    .filter(function (s) { return s.q7 || s.q11 || s.q12; })
    .sort(function (a, b) { return String(a.uid).localeCompare(String(b.uid)); });
}

// ── 退会メールデータ ──

function readActiveWithdrawRows_(ss) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return [];
  const rows = readWithdrawData_(dataSheet);
  const tz = Session.getScriptTimeZone();
  const out = [];

  const cancelMap = typeof loadWithdrawCancelMapIfAvailable_ === "function"
    ? loadWithdrawCancelMapIfAvailable_()
    : null;
  rows.forEach(function (row) {
    if (isWithdrawalCancelled_(row, cancelMap)) return;
    const ts = parseMemberTimestamp_(row[0]);
    const month = normalizeYearMonthLabel_(row[2]);
    out.push({
      ts: ts,
      tsLabel: ts ? Utilities.formatDate(ts, tz, "yyyy/MM/dd HH:mm") : String(row[0] || ""),
      name: String(row[1] || "").trim(),
      nameNorm: normalizeMemberName_(row[1]),
      month: month,
      category: String(row[3] || "").trim(),
      sortKey: (ts ? ts.getTime() : 0) + "\x1f" + String(row[1] || ""),
    });
  });

  out.sort(function (a, b) {
    if (a.sortKey < b.sortKey) return -1;
    if (a.sortKey > b.sortKey) return 1;
    return 0;
  });
  return out;
}

function buildAnketoMonthList_(withdraws) {
  const keys = {};
  withdraws.forEach(function (w) {
    if (!w.month || w.month === "（データなし）") return;
    const m = w.month.match(/^(\d{4})年(\d{1,2})月$/);
    if (!m) return;
    keys[w.month] = parseInt(m[1], 10) * 100 + parseInt(m[2], 10);
  });
  const today = new Date();
  const cur = formatYearMonth_(today.getFullYear(), today.getMonth() + 1);
  keys[cur] = today.getFullYear() * 100 + (today.getMonth() + 1);

  return Object.keys(keys).sort(function (a, b) { return keys[b] - keys[a]; });
}

function parseAnketoMonthLabel_(label) {
  const m = String(label || "").trim().match(/^(\d{4})年(\d{1,2})月$/);
  if (!m) return null;
  return {
    year: parseInt(m[1], 10),
    month: parseInt(m[2], 10) - 1,
    label: formatYearMonth_(parseInt(m[1], 10), parseInt(m[2], 10)),
  };
}

function countWithdrawForMonth_(withdraws, monthLabel) {
  let n = 0;
  withdraws.forEach(function (w) {
    if (w.month === monthLabel) n++;
  });
  return n;
}

// ── 突合（同姓同名は受信日時順 × uid順で1:1）──

/** 表示月の退会分だけ突合（全期間ペアリングより高速） */
function linkSurveyToWithdrawalsForMonth_(surveys, withdraws, monthLabel) {
  const monthWithdraws = withdraws.filter(function (w) { return w.month === monthLabel; });
  if (!monthWithdraws.length) return [];

  const surveysByName = groupAnketoByKey_(surveys, "nameNorm");
  const usedUid = {};

  monthWithdraws.forEach(function (w) {
    const key = w.nameNorm;
    if (!surveysByName[key]) surveysByName[key] = [];
  });

  Object.keys(surveysByName).forEach(function (nameNorm) {
    surveysByName[nameNorm].sort(function (a, b) {
      return String(a.uid).localeCompare(String(b.uid));
    });
  });

  const pools = {};
  Object.keys(surveysByName).forEach(function (nameNorm) {
    pools[nameNorm] = surveysByName[nameNorm].slice();
  });

  const linked = monthWithdraws.map(function (w) {
    const pool = pools[w.nameNorm] || [];
    const s = pool.length ? pool.shift() : null;
    if (s) usedUid[s.uid] = true;
    return {
      survey: s,
      withdraw: w,
      matchStatus: s ? "回答あり" : "退会のみ",
      month: w.month,
    };
  });

  linked.sort(function (a, b) {
    const ta = a.withdraw.ts ? a.withdraw.ts.getTime() : 0;
    const tb = b.withdraw.ts ? b.withdraw.ts.getTime() : 0;
    return tb - ta;
  });
  return linked;
}

function linkSurveyToWithdrawals_(surveys, withdraws) {
  const surveysByName = groupAnketoByKey_(surveys, "nameNorm");
  const withdrawsByName = groupAnketoByKey_(withdraws, "nameNorm");
  const allNames = {};
  Object.keys(surveysByName).forEach(function (k) { allNames[k] = true; });
  Object.keys(withdrawsByName).forEach(function (k) { allNames[k] = true; });

  const linked = [];
  Object.keys(allNames).forEach(function (nameNorm) {
    const sList = (surveysByName[nameNorm] || []).slice().sort(function (a, b) {
      return String(a.uid).localeCompare(String(b.uid));
    });
    const wList = (withdrawsByName[nameNorm] || []).slice().sort(function (a, b) {
      return (a.ts ? a.ts.getTime() : 0) - (b.ts ? b.ts.getTime() : 0);
    });

    const pairCount = Math.max(sList.length, wList.length);
    for (let i = 0; i < pairCount; i++) {
      const s = sList[i] || null;
      const w = wList[i] || null;
      let status = "回答あり";
      if (s && !w) status = "アンケートのみ";
      else if (!s && w) status = "退会のみ";
      linked.push({
        survey: s,
        withdraw: w,
        matchStatus: status,
        month: w ? w.month : "",
      });
    }
  });

  linked.sort(function (a, b) {
    const ta = a.withdraw && a.withdraw.ts ? a.withdraw.ts.getTime() : 0;
    const tb = b.withdraw && b.withdraw.ts ? b.withdraw.ts.getTime() : 0;
    return tb - ta;
  });
  return linked;
}

function groupAnketoByKey_(list, keyField) {
  const out = {};
  list.forEach(function (item) {
    const k = item[keyField];
    if (!k) return;
    if (!out[k]) out[k] = [];
    out[k].push(item);
  });
  return out;
}

function filterLinkedForMonth_(linked, monthLabel) {
  return linked.filter(function (r) {
    return r.withdraw && r.withdraw.month === monthLabel;
  });
}

// ── 描画 ──

function styleAnketoPctCell_(cell, pct) {
  if (pct >= 0.3) {
    cell.setBackground(ANKETO_THEME.red).setFontColor(ANKETO_THEME.white).setFontWeight("bold");
  } else if (pct >= 0.2) {
    cell.setBackground(ANKETO_THEME.redSoft).setFontColor(ANKETO_THEME.redDark).setFontWeight("bold");
  } else if (pct >= 0.15) {
    cell.setBackground(ANKETO_THEME.redPale).setFontColor(ANKETO_THEME.black);
  } else {
    cell.setBackground(ANKETO_THEME.white).setFontColor(ANKETO_THEME.grayMid);
  }
}

function styleAnketoHeaderRow_(range) {
  range.setFontWeight("bold").setBackground(ANKETO_THEME.headerBg).setFontColor(ANKETO_THEME.white);
}

function styleAnketoSubHeaderRow_(range) {
  range.setFontWeight("bold").setBackground(ANKETO_THEME.grayLight).setFontColor(ANKETO_THEME.black);
}

function clearAnketoReportArea_(sheet, clearRows, maxCol) {
  const safeRows = Math.max(clearRows, Math.max(sheet.getLastRow(), 3) - 2);
  const safeCols = Math.min(sheet.getMaxColumns(), Math.max(maxCol, sheet.getLastColumn(), 20));
  const reportArea = sheet.getRange(3, 1, safeRows, safeCols);

  // 月見出しの結合セルが残ると、次回の setValues と衝突してエラーになるため必ず解除する。
  reportArea.breakApart();

  reportArea
    .clearContent()
    .setBackground(null)
    .setFontWeight("normal")
    .setFontColor(ANKETO_THEME.black);

  const keptRules = sheet.getConditionalFormatRules().filter(function (rule) {
    return rule.getRanges().every(function (r) { return r.getLastRow() <= 2; });
  });
  if (keptRules.length !== sheet.getConditionalFormatRules().length) {
    sheet.setConditionalFormatRules(keptRules);
  }
}

function renderAnketoCompareSheet_(sheet, monthData) {
  const colCount = 1 + monthData.length * 2;
  const clearRows = Math.max(500, sheet.getLastRow() - 2);
  const maxCol = Math.max(colCount, ANKETO_DETAIL_HEADERS.length);
  clearAnketoReportArea_(sheet, clearRows, maxCol);

  let row = 3;
  row = writeAnketoKpiCompare_(sheet, row, monthData);
  row += 1;
  row = writeAnketoCompareBlock_(sheet, row, "■ 退会理由", monthData, "q12");
  row += 1;
  row = writeAnketoCompareBlock_(sheet, row, "■ 利用頻度", monthData, "q7");
  row += 1;
  row = writeAnketoCompareBlock_(sheet, row, "■ 利用時間帯（複数選択を分解）", monthData, "q11");
  row += 1;
  writeAnketoDetailCompact_(sheet, row, monthData);

  sheet.setFrozenRows(2);
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), 1))
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  applyAnketoCompareColWidths_(sheet, monthData);
  applyAnketoDetailColWidths_(sheet);
  if (typeof joyfitCompactDisplaySheet_ === "function") {
    joyfitCompactDisplaySheet_(sheet, 10, 80, 3, []);
  }
}

function applyAnketoCompareColWidths_(sheet, monthData) {
  let maxItemLen = 6;
  ["q12", "q7", "q11"].forEach(function (key) {
    buildAnketoCompareItemList_(monthData, key).forEach(function (item) {
      maxItemLen = Math.max(maxItemLen, String(item).length);
    });
  });
  sheet.setColumnWidth(1, Math.min(420, Math.max(260, maxItemLen * 8 + 24)));

  monthData.forEach(function (m, i) {
    const labelLen = String(m.label).length;
    const pairW = Math.max(148, labelLen * 9 + 48);
    sheet.setColumnWidth(2 + i * 2, Math.max(52, Math.round(pairW * 0.38)));
    sheet.setColumnWidth(3 + i * 2, Math.max(72, Math.round(pairW * 0.62)));
  });
}

function applyAnketoDetailColWidths_(sheet) {
  const widths = [84, 132, 104, 112, 80, 44, 168, 220, 68, 60];
  widths.forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, Math.max(sheet.getColumnWidth(i + 1), w));
  });
}

function writeAnketoKpiCompare_(sheet, startRow, monthData) {
  const headerRow = startRow;
  const header = ["指標"];
  monthData.forEach(function (m) { header.push(m.label); });

  const headerRange = sheet.getRange(headerRow, 1, 1, header.length);
  headerRange.setValues([header]).setHorizontalAlignment("center");
  styleAnketoHeaderRow_(headerRange);

  const kpiDefs = [
    {
      label: "退会数",
      value: function (m) { return m.meta.withdrawMonthTotal; },
      format: "0",
    },
    {
      label: "アンケート総数",
      value: function (m) { return m.meta.linkedTotal; },
      format: "0",
    },
  ];

  const values = kpiDefs.map(function (kpi) {
    const row = [kpi.label];
    monthData.forEach(function (m) { row.push(kpi.value(m)); });
    return row;
  });

  const dataStart = headerRow + 1;
  sheet.getRange(dataStart, 1, values.length, header.length).setValues(values);
  sheet.getRange(dataStart, 1, values.length, header.length)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.getRange(dataStart, 1, values.length, 1).setFontWeight("bold").setFontColor(ANKETO_THEME.black);
  for (let i = 0; i < values.length; i++) {
    if (i % 2 === 1) {
      sheet.getRange(dataStart + i, 1, 1, header.length).setBackground(ANKETO_THEME.grayPale);
    }
  }
  kpiDefs.forEach(function (kpi, i) {
    sheet.getRange(dataStart + i, 2, 1, monthData.length).setNumberFormat(kpi.format).setHorizontalAlignment("center");
  });

  return dataStart + values.length;
}

function buildAnketoCompareItemList_(monthData, fieldKey) {
  const score = {};
  monthData.forEach(function (m, i) {
    const counter = m.stats[fieldKey] || {};
    Object.keys(counter).forEach(function (k) {
      if (!score[k]) score[k] = 0;
      score[k] += counter[k] * (i === 0 ? 1000 : 1);
    });
  });
  return Object.keys(score).sort(function (a, b) { return score[b] - score[a]; });
}

function writeAnketoCompareBlock_(sheet, startRow, title, monthData, fieldKey) {
  const titleRange = sheet.getRange(startRow, 1, 1, 1 + monthData.length * 2);
  titleRange.breakApart();
  titleRange
    .setBackground(fieldKey === "q12" ? ANKETO_THEME.red : ANKETO_THEME.black)
    .setFontColor(ANKETO_THEME.white)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.getRange(startRow, 1).setValue(title).setFontSize(11);

  const colCount = 1 + monthData.length * 2;
  const headerRow = startRow + 1;
  const subHeaderRow = headerRow + 1;

  const header1 = ["項目"];
  monthData.forEach(function (m) { header1.push(m.label, ""); });
  const header1Range = sheet.getRange(headerRow, 1, 1, colCount);
  header1Range.setValues([header1]).setHorizontalAlignment("center");
  styleAnketoHeaderRow_(header1Range);

  monthData.forEach(function (m, i) {
    sheet.getRange(headerRow, 2 + i * 2, 1, 2).merge();
  });

  const header2 = [""];
  monthData.forEach(function () { header2.push("件数", "構成比"); });
  const header2Range = sheet.getRange(subHeaderRow, 1, 1, colCount);
  header2Range.setValues([header2]).setHorizontalAlignment("center");
  styleAnketoSubHeaderRow_(header2Range);

  const items = buildAnketoCompareItemList_(monthData, fieldKey);
  const dataRows = items.map(function (item) {
    const row = [item];
    monthData.forEach(function (m) {
      const counter = m.stats[fieldKey] || {};
      const count = counter[item] || 0;
      const total = Object.keys(counter).reduce(function (s, k) { return s + counter[k]; }, 0);
      row.push(count, total ? count / total : 0);
    });
    return row;
  });

  const dataStart = subHeaderRow + 1;
  if (dataRows.length) {
    sheet.getRange(dataStart, 1, dataRows.length, colCount)
      .setValues(dataRows)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.getRange(dataStart, 1, dataRows.length, 1).setWrap(true);
    monthData.forEach(function (m, i) {
      const countCol = 2 + i * 2;
      const pctCol = countCol + 1;
      sheet.getRange(dataStart, countCol, dataRows.length, 1).setNumberFormat("0").setHorizontalAlignment("center");
      sheet.getRange(dataStart, pctCol, dataRows.length, 1).setNumberFormat("0.0%").setHorizontalAlignment("center");
    });
    for (let i = 0; i < dataRows.length; i++) {
      const rowNum = dataStart + i;
      const baseBg = i % 2 === 1 ? ANKETO_THEME.grayPale : ANKETO_THEME.white;
      const isTopReason = fieldKey === "q12" && i === 0;
      sheet.getRange(rowNum, 1)
        .setBackground(isTopReason ? ANKETO_THEME.redPale : baseBg)
        .setFontColor(isTopReason ? ANKETO_THEME.redDark : ANKETO_THEME.black)
        .setFontWeight(isTopReason ? "bold" : "normal")
        .setHorizontalAlignment("center");
      monthData.forEach(function (m, j) {
        const countCol = 2 + j * 2;
        const pctCol = countCol + 1;
        const pct = dataRows[i][pctCol - 1];
        sheet.getRange(rowNum, countCol)
          .setBackground(isTopReason ? ANKETO_THEME.redPale : baseBg)
          .setFontColor(ANKETO_THEME.black)
          .setFontWeight(isTopReason ? "bold" : "normal")
          .setHorizontalAlignment("center");
        styleAnketoPctCell_(sheet.getRange(rowNum, pctCol), pct);
      });
    }
  } else {
    sheet.getRange(dataStart, 1).setValue("（データなし）").setFontColor(ANKETO_THEME.grayMid);
    return dataStart + 1;
  }

  return dataStart + dataRows.length;
}

function writeAnketoDetailCompact_(sheet, startRow, monthData) {
  sheet.getRange(startRow, 1)
    .setValue("■ 詳細のみ")
    .setFontWeight("bold")
    .setFontSize(10)
    .setFontColor(ANKETO_THEME.grayMid);

  const headerRow = startRow + 1;
  const headerRange = sheet.getRange(headerRow, 1, 1, ANKETO_DETAIL_HEADERS.length);
  headerRange.setValues([ANKETO_DETAIL_HEADERS])
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  styleAnketoSubHeaderRow_(headerRange);

  const monthOrder = {};
  monthData.forEach(function (m, i) { monthOrder[m.label] = i; });

  const allRows = [];
  monthData.forEach(function (m) {
    m.monthRows.forEach(function (r) { allRows.push(r); });
  });
  allRows.sort(function (a, b) {
    const ma = a.withdraw && a.withdraw.month ? monthOrder[a.withdraw.month] : 99;
    const mb = b.withdraw && b.withdraw.month ? monthOrder[b.withdraw.month] : 99;
    if (ma !== mb) return ma - mb;
    const ta = a.withdraw && a.withdraw.ts ? a.withdraw.ts.getTime() : 0;
    const tb = b.withdraw && b.withdraw.ts ? b.withdraw.ts.getTime() : 0;
    return tb - ta;
  });

  const detailValues = allRows.map(function (r) {
    const w = r.withdraw;
    const s = r.survey;
    return [
      w ? w.month : "",
      w ? w.tsLabel : "",
      w ? w.name : (s ? s.name : ""),
      s ? s.uid : "",
      s ? s.q7 : "",
      s ? s.q8 : "",
      s ? s.q11 : "",
      s ? s.q12 : "",
      w ? w.category : "",
      r.matchStatus,
    ];
  });

  if (!detailValues.length) return headerRow + 1;

  const dataStart = headerRow + 1;
  sheet.getRange(dataStart, 1, detailValues.length, ANKETO_DETAIL_HEADERS.length)
    .setValues(detailValues)
    .setFontSize(9)
    .setFontColor(ANKETO_THEME.black)
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.getRange(dataStart, 2, detailValues.length, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  sheet.getRange(dataStart, 1, detailValues.length, ANKETO_DETAIL_HEADERS.length).setNumberFormat("@");
  for (let i = 0; i < detailValues.length; i++) {
    if (i % 2 === 1) {
      sheet.getRange(dataStart + i, 1, 1, ANKETO_DETAIL_HEADERS.length).setBackground(ANKETO_THEME.grayPale);
    }
  }

  return dataStart + detailValues.length;
}

function collectAnketoAnswers_(rows, field) {
  const c = {};
  rows.forEach(function (r) {
    if (!r.survey) return;
    const v = String(r.survey[field] || "").trim();
    if (!v) return;
    c[v] = (c[v] || 0) + 1;
  });
  return c;
}

function expandAnketoSlotCounts_(counter) {
  const slotCounter = {};
  Object.keys(counter).forEach(function (ans) {
    ans.split("\t").forEach(function (slot) {
      slot = slot.trim();
      if (!slot) return;
      slotCounter[slot] = (slotCounter[slot] || 0) + counter[ans];
    });
  });
  return slotCounter;
}

// ── CSV取込（HTMLダイアログ）──

function showAnketoCsvImportDialog_() {
  const html = HtmlService.createHtmlOutput(buildAnketoCsvImportHtml_())
    .setWidth(540)
    .setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, "退会アンケートCSVを取込");
}

function buildAnketoCsvImportHtml_() {
  return [
    "<!DOCTYPE html><html><head><base target=\"_top\">",
    "<style>",
    "body{font-family:'Meiryo',sans-serif;padding:16px;color:#333;margin:0}",
    ".note{font-size:12px;color:#555;background:#e8f0fe;padding:10px;border-radius:4px;",
    "border-left:4px solid #1a73e8;margin-bottom:14px;line-height:1.5}",
    "#dropZone{border:2px dashed #9aa0a6;border-radius:8px;padding:28px 16px;text-align:center;",
    "background:#fafafa;cursor:pointer;transition:.15s}",
    "#dropZone.hover{border-color:#1a73e8;background:#eef4ff}",
    "#dropZone .main{font-size:15px;font-weight:bold;margin-bottom:6px}",
    "#dropZone .sub{font-size:12px;color:#666}",
    "#fileName{font-size:12px;color:#1a73e8;margin:10px 0 0;min-height:18px}",
    ".actions{text-align:right;margin-top:16px}",
    "button{padding:10px 18px;border:none;border-radius:4px;cursor:pointer;font-size:14px;font-weight:bold}",
    ".btn-go{background:#1a73e8;color:#fff;margin-left:8px}",
    ".btn-cancel{background:#f1f3f4;color:#333}",
    "button:disabled{background:#ccc;cursor:not-allowed}",
    "</style></head><body>",
    "<div class=\"note\">",
    "Web入会システムの退会アンケートCSVをここにドロップするか、ファイルを選択してください。<br>",
    "・既存データは保持したまま <b>uid × 設問</b> 単位でマージします（重複しません）<br>",
    "・取込後に「アンケート」シートを自動更新します<br>",
    "・ダイアログが進まない場合は <b>取込シートからマージ（推奨）</b> を使ってください",
    "</div>",
    "<div id=\"dropZone\">",
    "<div class=\"main\">CSVをドロップ</div>",
    "<div class=\"sub\">またはクリックしてファイルを選択（.csv）</div>",
    "<div id=\"fileName\"></div>",
    "</div>",
    "<input type=\"file\" id=\"fileInput\" accept=\".csv,text/csv\" style=\"display:none\">",
    "<div class=\"actions\">",
    "<button class=\"btn-cancel\" onclick=\"google.script.host.close()\">キャンセル</button>",
    "<button id=\"goBtn\" class=\"btn-go\" disabled onclick=\"upload()\">取込して更新</button>",
    "</div>",
    "<script>",
    "var pickedRows=null,pickedName='';",
    "var zone=document.getElementById('dropZone');",
    "var input=document.getElementById('fileInput');",
    "zone.onclick=function(){input.click();};",
    "zone.ondragover=function(e){e.preventDefault();zone.classList.add('hover');};",
    "zone.ondragleave=function(){zone.classList.remove('hover');};",
    "zone.ondrop=function(e){e.preventDefault();zone.classList.remove('hover');",
    "if(e.dataTransfer.files.length) readFile(e.dataTransfer.files[0]);};",
    "input.onchange=function(){if(input.files.length) readFile(input.files[0]);};",
    "function decodeCsvBytes(buffer){",
    "try{return new TextDecoder('shift-jis').decode(buffer);}",
    "catch(e1){try{return new TextDecoder('windows-31j').decode(buffer);}",
    "catch(e2){return new TextDecoder('utf-8').decode(buffer);}}}",
    "function parseCsvText(text){",
    "var rows=[],row=[],field='',inQ=false,i,c,n;",
    "text=String(text||'').replace(/^\\uFEFF/,'');",
    "for(i=0;i<text.length;i++){",
    "c=text.charAt(i);n=text.charAt(i+1);",
    "if(inQ){",
    "if(c=='\"'&&n=='\"'){field+='\"';i++;}",
    "else if(c=='\"')inQ=false;",
    "else field+=c;",
    "}else{",
    "if(c=='\"')inQ=true;",
    "else if(c==','){row.push(field);field='';}",
    "else if(c=='\\n'||(c=='\\r'&&n=='\\n')){",
    "if(c=='\\r')i++;row.push(field);field='';",
    "if(row.length>1||String(row[0]||'').trim())rows.push(row);row=[];",
    "}else field+=c;}}",
    "if(field||row.length){row.push(field);rows.push(row);}",
    "return rows;}",
    "function readFile(file){",
    "if(!/\\.csv$/i.test(file.name)){alert('CSVファイルを選んでください');return;}",
    "pickedName=file.name;",
    "document.getElementById('fileName').textContent=file.name+' ('+Math.round(file.size/1024)+' KB)';",
    "document.getElementById('goBtn').disabled=true;",
    "var reader=new FileReader();",
    "reader.onload=function(ev){",
    "try{",
    "var rows=parseCsvText(decodeCsvBytes(ev.target.result));",
    "if(rows.length&&String(rows[0][0]||'').toLowerCase().indexOf('shop')>=0)rows.shift();",
    "pickedRows=rows.filter(function(r){return String(r[1]||'').trim();});",
    "if(!pickedRows.length)throw new Error('有効な行がありません');",
    "document.getElementById('goBtn').disabled=false;",
    "}catch(err){pickedRows=null;",
    "alert('CSV読込エラー:\\n'+(err&&err.message?err.message:err));",
    "document.getElementById('goBtn').disabled=true;}};",
    "reader.readAsArrayBuffer(file);}",
    "function upload(){",
    "if(!pickedRows||!pickedRows.length)return;",
    "var btn=document.getElementById('goBtn');",
    "btn.disabled=true;",
    "uploadRowBatches(pickedRows,pickedName,btn);}",
    "function uploadRowBatches(rows,fileName,btn){",
    "var BATCH=" + ANKETO_UPLOAD_BATCH_ROWS + ";",
    "var total=Math.ceil(rows.length/BATCH);",
    "var session='',idx=0,activeTimer=null;",
    "function resetBtn(){btn.disabled=false;btn.textContent='取込して更新';}",
    "function fail(msg){if(activeTimer)clearTimeout(activeTimer);alert(msg);resetBtn();}",
    "function sendNext(){",
    "if(idx>=total){",
    "if(activeTimer)clearTimeout(activeTimer);",
    "btn.textContent='集計中...';",
    "google.script.run",
    ".withSuccessHandler(function(pack){",
    "google.script.run",
    ".withSuccessHandler(function(res){alert(res.message);google.script.host.close();})",
    ".withFailureHandler(function(err){",
    "fail('集計エラー:\\n'+(err&&err.message?err.message:err));})",
    ".importAnketoCsvFinishImport_(pack);",
    "})",
    ".withFailureHandler(function(err){",
    "fail('取込エラー:\\n'+(err&&err.message?err.message:err));})",
    ".importAnketoCsvFinishUpload_(session,fileName);",
    "return;}",
    "var part=rows.slice(idx*BATCH,(idx+1)*BATCH);",
    "var cur=idx;",
    "if(activeTimer)clearTimeout(activeTimer);",
    "activeTimer=setTimeout(function(){",
    "fail('サーバー応答がありません（'+(cur+1)+'/'+total+'）。\\n\\n'",
    "+'「取込シートからマージ（推奨）」をお試しください。');",
    "},60000);",
    "google.script.run",
    ".withSuccessHandler(function(s){",
    "if(activeTimer)clearTimeout(activeTimer);",
    "session=s;idx++;",
    "btn.textContent='CSV送信中 '+idx+'/'+total+'...';sendNext();})",
    ".withFailureHandler(function(err){",
    "fail('送信エラー（'+(cur+1)+'/'+total+'）:\\n'",
    "+(err&&err.message?err.message:err));})",
    ".anketoCsvUploadRowsBatch_(session,cur,part,fileName,cur===0,total);}",
    "btn.textContent='CSV送信中 0/'+total+'...';sendNext();}",
    "google.script.run.withFailureHandler(function(){}).anketoCsvUploadPing_();",
    "</script></body></html>",
  ].join("");
}

/** ダイアログ表示時: 接続確認 */
function anketoCsvUploadPing_() {
  return true;
}

function ensureAnketoUploadBufferSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME_ANKETO_UPLOAD);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_ANKETO_UPLOAD);
    try {
      sheet.hideSheet();
    } catch (e) {
      Logger.log("アンケート_送信 非表示: " + e.message);
    }
  }
  return sheet;
}

/** 分割送信: 行バッチ受信（非表示シートに保存） */
function anketoCsvUploadRowsBatch_(session, batchIndex, rows, fileName, isFirst, totalBatches) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAnketoUploadBufferSheet_(ss);
  const idx = Number(batchIndex);
  const batches = Number(totalBatches);

  if (!session || isFirst) {
    session = Utilities.getUuid();
    sheet.clear();
    sheet.getRange(1, 1, 1, 4).setValues([["meta", session, String(fileName || ""), batches]]);
  }

  if (rows && rows.length) {
    const startRow = 2 + idx * ANKETO_UPLOAD_BATCH_ROWS;
    const values = rows.map(function (r) {
      return [
        String(r[0] || ""),
        String(r[1] || ""),
        String(r[2] || ""),
        String(r[3] || ""),
        String(r[4] || ""),
      ];
    });
    sheet.getRange(startRow, 1, values.length, ANKETO_RAW_HEADERS.length).setValues(values);
  }
  return session;
}

/** 分割送信: 受信行を結合してマージ */
function importAnketoCsvFinishUpload_(session, fileName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_ANKETO_UPLOAD);
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error("アップロードデータが見つかりません。もう一度お試しください。");
  }

  const meta = sheet.getRange(1, 1, 1, 4).getValues()[0];
  const storedSession = String(meta[1] || "");
  if (storedSession && session && storedSession !== session) {
    throw new Error("アップロードセッションが一致しません。もう一度お試しください。");
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, ANKETO_RAW_HEADERS.length).getValues();
  sheet.clear();
  try {
    sheet.hideSheet();
  } catch (e) {
    Logger.log("アンケート_送信 クリア後非表示: " + e.message);
  }

  const incoming = [];
  values.forEach(function (row) {
    const normalized = normalizeAnketoCsvRow_(row);
    if (normalized) incoming.push(normalized);
  });
  if (!incoming.length) {
    throw new Error("CSVに有効な回答がありません。ファイル形式を確認してください。");
  }

  return importAnketoCsvMergeFromIncoming_(incoming, fileName || String(meta[2] || ""));
}

/** アンケート_取込 シートからマージ */
function importAnketoCsvFromPasteSheetCore_(ss) {
  ensureAnketoSheets_(ss);
  const paste = ss.getSheetByName(SHEET_NAME_ANKETO_PASTE);
  if (!paste || paste.getLastRow() < 2) {
    throw new Error(
      "シート「" + SHEET_NAME_ANKETO_PASTE + "」にCSVを貼り付けてください。\n" +
        "（タブが無ければメニュー実行後に作成されます）"
    );
  }

  const incoming = readAnketoRowsFromPasteSheet_(paste);
  if (!incoming.length) {
    throw new Error(
      "「" + SHEET_NAME_ANKETO_PASTE + "」に有効なデータがありません。\n" +
        "文字化けしている場合は「ファイル→インポート」でCSVを直接読み込んでください。"
    );
  }

  const existing = readAnketoRawRows_(ss);
  const merged = mergeAnketoRawRows_(existing, incoming);
  writeAnketoRawSheet_(ss, merged.rows);

  if (paste.getLastRow() > 1) {
    paste.getRange(2, 1, paste.getLastRow() - 1, ANKETO_RAW_HEADERS.length).clearContent();
  }

  PropertiesService.getScriptProperties().setProperty(
    "anketoLastImport",
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") +
      "\t取込シート"
  );

  return { merged: merged, incomingCount: incoming.length, fileName: "取込シート" };
}

/** ダイアログから一括実行（互換用・小さいCSVのみ） */
function importAnketoCsvFromUpload_(base64, fileName) {
  const pack = importAnketoCsvMergeOnly_(base64, fileName);
  return importAnketoCsvFinishImport_(pack);
}

/** CSVマージのみ（高速） */
function importAnketoCsvMergeOnly_(base64, fileName) {
  const incoming = parseAnketoCsvText_(decodeAnketoCsvBytes_(base64));
  return importAnketoCsvMergeFromIncoming_(incoming, fileName);
}

function importAnketoCsvMergeFromIncoming_(incoming, fileName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAnketoSheets_(ss);

  if (!incoming.length) {
    throw new Error("CSVに有効な回答がありません。ファイル形式を確認してください。");
  }

  const existing = readAnketoRawRows_(ss);
  const merged = mergeAnketoRawRows_(existing, incoming);
  writeAnketoRawSheet_(ss, merged.rows);

  PropertiesService.getScriptProperties().setProperty(
    "anketoLastImport",
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") +
      "\t" + String(fileName || "")
  );

  return { merged: merged, incomingCount: incoming.length, fileName: fileName || "" };
}

/** ダイアログ2段階目: マージ済みデータで集計のみ */
function importAnketoCsvFinishImport_(pack) {
  const display = syncAnketoDisplay_(SpreadsheetApp.getActiveSpreadsheet());
  return buildAnketoImportResultMessage_(pack, display, pack.fileName || "");
}

function buildAnketoImportResultMessage_(pack, display, fileName) {
  const merged = pack.merged;
  return {
    message:
      "取込が完了しました。\n\n" +
      "ファイル: " + (fileName || "（不明）") + "\n" +
      "今回のCSV … 有効 " + pack.incomingCount + " 行\n" +
      "マージ … 新規 " + merged.added + " / 更新 " + merged.updated +
      " / 変更なし " + merged.unchanged + "\n" +
      "マスタ合計 … " + merged.rows.length + " 行（uid " + merged.uidCount + " 人）\n\n" +
      "アンケート表示 … " + display.monthLabel +
      "（退会 " + display.monthDetail + " 件）",
  };
}

function decodeAnketoCsvBytes_(base64) {
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes);
  try {
    return blob.getDataAsString("Shift_JIS");
  } catch (e1) {
    try {
      return blob.getDataAsString("Windows-31J");
    } catch (e2) {
      return blob.getDataAsString("UTF-8");
    }
  }
}

function parseAnketoCsvText_(text) {
  const parsed = Utilities.parseCsv(String(text || "").replace(/^\uFEFF/, ""));
  if (!parsed.length) return [];

  let start = 0;
  const h0 = String(parsed[0][0] || "").trim().toLowerCase();
  if (h0 === "shop_id" || h0.indexOf("shop") >= 0) start = 1;

  const out = [];
  for (let i = start; i < parsed.length; i++) {
    const row = normalizeAnketoCsvRow_(parsed[i]);
    if (row) out.push(row);
  }
  return out;
}

function normalizeAnketoCsvRow_(row) {
  if (!row || !row.length) return null;
  const shop = String(row[0] || "").trim().replace(/^\"|\"$/g, "");
  const uid = String(row[1] || "").trim().replace(/^\"|\"$/g, "");
  const name = String(row[2] || "").trim().replace(/^\"|\"$/g, "");
  const qid = String(row[3] || "").trim().replace(/^\"|\"$/g, "");
  const ans = String(row[4] || "").trim().replace(/^\"|\"$/g, "");
  if (!uid || !qid) return null;
  if (ANKETO_Q_SKIP[qid]) return null;
  return { shop: shop, uid: uid, name: name, qId: qid, answer: ans };
}

function mergeAnketoRawRows_(existing, incoming) {
  const map = {};
  existing.forEach(function (r) {
    map[anketoRawRowKey_(r)] = {
      shop: r.shop,
      uid: r.uid,
      name: r.name,
      qId: r.qId,
      answer: r.answer,
    };
  });

  let added = 0;
  let updated = 0;
  let unchanged = 0;

  incoming.forEach(function (r) {
    const key = anketoRawRowKey_(r);
    const prev = map[key];
    if (!prev) {
      map[key] = {
        shop: r.shop,
        uid: r.uid,
        name: r.name,
        qId: r.qId,
        answer: r.answer,
      };
      added++;
      return;
    }
    const next = {
      shop: r.shop || prev.shop,
      uid: r.uid,
      name: r.name || prev.name,
      qId: r.qId,
      answer: r.answer,
    };
    if (next.shop === prev.shop && next.name === prev.name && next.answer === prev.answer) {
      unchanged++;
      return;
    }
    map[key] = next;
    updated++;
  });

  const rows = Object.keys(map).map(function (k) { return map[k]; });
  rows.sort(function (a, b) {
    const u = String(a.uid).localeCompare(String(b.uid));
    if (u !== 0) return u;
    const qa = parseInt(a.qId, 10);
    const qb = parseInt(b.qId, 10);
    if (!isNaN(qa) && !isNaN(qb) && qa !== qb) return qa - qb;
    return String(a.qId).localeCompare(String(b.qId));
  });

  const uidSet = {};
  rows.forEach(function (r) { uidSet[r.uid] = true; });

  return {
    rows: rows,
    added: added,
    updated: updated,
    unchanged: unchanged,
    uidCount: Object.keys(uidSet).length,
  };
}

function anketoRawRowKey_(r) {
  return String(r.uid) + "\x1f" + String(r.qId);
}

function writeAnketoRawSheet_(ss, rows) {
  let raw = ss.getSheetByName(SHEET_NAME_ANKETO_RAW);
  if (raw) {
    try {
      ss.deleteSheet(raw);
    } catch (e) {
      Logger.log("アンケート_退会 再作成: " + e.message);
    }
  }
  raw = ss.insertSheet(SHEET_NAME_ANKETO_RAW);
  raw.getRange(1, 1, 1, ANKETO_RAW_HEADERS.length).setValues([ANKETO_RAW_HEADERS]);
  raw.getRange(1, 1, 1, ANKETO_RAW_HEADERS.length).setFontWeight("bold").setBackground("#f3f3f3");
  raw.setFrozenRows(1);

  if (!rows.length) {
    try { raw.hideSheet(); } catch (e) {}
    return;
  }

  const values = rows.map(function (r) {
    return [r.shop, r.uid, r.name, r.qId, r.answer];
  });
  const chunk = 2000;
  for (let i = 0; i < values.length; i += chunk) {
    const part = values.slice(i, i + chunk);
    raw.getRange(2 + i, 1, part.length, ANKETO_RAW_HEADERS.length).setValues(part);
  }
  try {
    raw.hideSheet();
  } catch (e) {
    Logger.log("アンケート_退会 非表示: " + e.message);
  }
}

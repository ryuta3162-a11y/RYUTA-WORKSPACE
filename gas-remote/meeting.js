/**
 * 会議用 会員分析 Dashboard（数式中心）
 * - 一回だけ「作り直す」でシートと数式を配置
 * - 以降の集計・比較はすべてスプレッドシート関数（GAS不要）
 * - 既存の累計入会／累計退会／日報／会員分析は変更しない
 */

var MEETING_SHEET_ = '会議用';
var MEETING_MASTER_ = '分析用_会員マスタ';
var MEETING_CALC_ = '分析用_期間集計';
var MEETING_JOIN_ = '累計入会データ';
var MEETING_LEAVE_ = '累計退会データ';
var MEETING_END_ = 8000;
var MEETING_VER_ = '1';

function rebuildMeetingDashboardFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var r = buildMeetingDashboard_(ss);
    ss.toast('会議用ダッシュボードを作りました（' + (r.sheets || []).join(' / ') + '）', '会議用', 10);
    var sh = ss.getSheetByName(MEETING_SHEET_);
    if (sh) ss.setActiveSheet(sh);
    return r;
  } catch (err) {
    ss.toast('会議用エラー: ' + String(err && err.message ? err.message : err), '会議用', 12);
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function buildMeetingDashboard_(ss) {
  var theme = hubTheme_();
  var master = meetingEnsureSheet_(ss, MEETING_MASTER_, MEETING_SHEET_);
  var calc = meetingEnsureSheet_(ss, MEETING_CALC_, MEETING_MASTER_);
  var dash = meetingEnsureSheet_(ss, MEETING_SHEET_, '会員分析');
  meetingBuildMaster_(master, theme);
  meetingBuildCalc_(calc, theme);
  meetingBuildDash_(dash, theme);
  PropertiesService.getDocumentProperties().setProperty('MEETING_DASH_VER', MEETING_VER_);
  return { ok: true, sheets: [MEETING_MASTER_, MEETING_CALC_, MEETING_SHEET_], ver: MEETING_VER_ };
}

function meetingEnsureSheet_(ss, name, afterName) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    var after = afterName ? ss.getSheetByName(afterName) : null;
    sh = ss.insertSheet(name, after ? after.getIndex() : ss.getSheets().length);
  }
  sh.clear();
  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  sh.clearConditionalFormatRules();
  try { sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); }); } catch (eP) {}
  return sh;
}

/** 1行=1入会履歴。退会日は累計退会を番号キーで照合。日付は届出日／退会届出日（日単位） */
function meetingBuildMaster_(sh, theme) {
  sh.setTabColor('#424242');
  sh.setHiddenGridlines(true);
  var end = MEETING_END_;
  var J = "'" + MEETING_JOIN_ + "'!";
  var L = "'" + MEETING_LEAVE_ + "'!";
  var headers = [
    '元行', '会員番号', '氏名', 'フリガナ', '性別', '生年月日', '年齢表示', 'TEL',
    '契約名称', '会員区分', '会員種類', '手続',
    '届出日', '利用開始年月', '退会届出日', '退会年月', '退会理由', '在籍期間月',
    '番号キー', '再入会フラグ', '参照元', '期間終了時年齢用生年月日'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(theme.ink).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9).setWrap(true);
  sh.setFrozenRows(1);

  // A 元行
  sh.getRange('A2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",ROW(' + J + 'F2:F' + end + ')))');
  // B 会員番号
  sh.getRange('B2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'F2:F' + end + '))');
  // C 氏名
  sh.getRange('C2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'G2:G' + end + '))');
  // D フリガナ
  sh.getRange('D2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'H2:H' + end + '))');
  // E 性別
  sh.getRange('E2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'I2:I' + end + '))');
  // F 生年月日
  sh.getRange('F2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'K2:K' + end + '))');
  // G 年齢表示（マスタ上の年齢列・参考）
  sh.getRange('G2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'J2:J' + end + '))');
  // H TEL
  sh.getRange('H2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'L2:L' + end + '))');
  // I 契約
  sh.getRange('I2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'C2:C' + end + '))');
  // J 区分
  sh.getRange('J2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'A2:A' + end + '))');
  // K 種類
  sh.getRange('K2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'B2:B' + end + '))');
  // L 手続
  sh.getRange('L2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'O2:O' + end + '))');
  // M 届出日（日単位の入会日として使う）
  sh.getRange('M2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'M2:M' + end + '))');
  // N 利用開始年月（月単位。会員動向の入会定義に近い）
  sh.getRange('N2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'N2:N' + end + '))');
  // O 退会届出日（番号キー照合）
  sh.getRange('O2').setFormula(
    '=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",IFERROR(XLOOKUP(TEXT(VALUE(' + J + 'F2:F' + end + '),"0"),' +
    L + 'T2:T' + end + ',' + L + 'Q2:Q' + end + ',IFERROR(XLOOKUP(' + J + 'F2:F' + end + ',' + L + 'F2:F' + end + ',' + L + 'Q2:Q' + end + '),""))))'
  );
  // P 退会年月
  sh.getRange('P2').setFormula(
    '=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",IFERROR(XLOOKUP(TEXT(VALUE(' + J + 'F2:F' + end + '),"0"),' +
    L + 'T2:T' + end + ',' + L + 'R2:R' + end + ',IFERROR(XLOOKUP(' + J + 'F2:F' + end + ',' + L + 'F2:F' + end + ',' + L + 'R2:R' + end + '),""))))'
  );
  // Q 退会理由
  sh.getRange('Q2').setFormula(
    '=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",IFERROR(XLOOKUP(TEXT(VALUE(' + J + 'F2:F' + end + '),"0"),' +
    L + 'T2:T' + end + ',' + L + 'S2:S' + end + ',IFERROR(XLOOKUP(' + J + 'F2:F' + end + ',' + L + 'F2:F' + end + ',' + L + 'S2:S' + end + '),""))))'
  );
  // R 在籍期間月（退会側P列）
  sh.getRange('R2').setFormula(
    '=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",IFERROR(XLOOKUP(TEXT(VALUE(' + J + 'F2:F' + end + '),"0"),' +
    L + 'T2:T' + end + ',' + L + 'P2:P' + end + ',IFERROR(XLOOKUP(' + J + 'F2:F' + end + ',' + L + 'F2:F' + end + ',' + L + 'P2:P' + end + '),""))))'
  );
  // S 番号キー
  sh.getRange('S2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",IFERROR(TEXT(VALUE(' + J + 'F2:F' + end + '),"0"),' + J + 'F2:F' + end + ')))');
  // T 再入会フラグ（手続「復」＝確定。候補判定は会議用の検索・定義欄を参照）
  sh.getRange('T2').setFormula(
    '=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",IF(' + J + 'O2:O' + end + '="復","再入会","")))'
  );
  // U 参照元
  sh.getRange('U2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","","累計入会データ"))');
  // V 生年月日（年齢計算用コピー）
  sh.getRange('V2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' + J + 'K2:K' + end + '))');

  sh.getRange(2, 13, 1, 4).setNumberFormat('yyyy/m/d');
  sh.setColumnWidths(1, headers.length, 90);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(8, 70);
  sh.getRange('A1').setNote(
    '日単位分析の入会日＝届出日(M)、退会日＝退会届出日(O)。\n' +
    '利用開始年月(N)・退会年月(P)は月単位。会員動向の入会定義に近いのは利用開始年月。\n' +
    '架空の日付は作らない。'
  );
  try { sh.hideSheet(); } catch (eH) {}
}

/**
 * 期間A〜Dの集計結果をここに置く。会議用シートはここを参照するだけ。
 * 行: 指標、列: A〜D
 */
function meetingBuildCalc_(sh, theme) {
  sh.setTabColor('#6B6B6B');
  sh.setHiddenGridlines(true);
  sh.getRange('A1').setValue('分析用_期間集計').setFontWeight('bold').setFontSize(12);
  sh.getRange('A2').setValue('会議用シートの入力（開始日・終了日・フィルタ）を参照して集計。手入力しないこと。')
    .setFontColor(theme.mute).setFontSize(9);
  sh.getRange('A2:F2').merge();

  // 期間パラメータを会議用から参照
  sh.getRange('A4').setValue('期間パラメータ（会議用から）').setFontWeight('bold');
  sh.getRange(5, 1, 1, 5).setValues([['', 'A', 'B', 'C', 'D']])
    .setBackground(theme.ink).setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange('A6').setValue('開始');
  sh.getRange('A7').setValue('終了');
  sh.getRange('A8').setValue('日数');
  sh.getRange('B6').setFormula('=\'' + MEETING_SHEET_ + '\'!B5');
  sh.getRange('C6').setFormula('=\'' + MEETING_SHEET_ + '\'!D5');
  sh.getRange('D6').setFormula('=\'' + MEETING_SHEET_ + '\'!F5');
  sh.getRange('E6').setFormula('=\'' + MEETING_SHEET_ + '\'!H5');
  sh.getRange('B7').setFormula('=\'' + MEETING_SHEET_ + '\'!B6');
  sh.getRange('C7').setFormula('=\'' + MEETING_SHEET_ + '\'!D6');
  sh.getRange('D7').setFormula('=\'' + MEETING_SHEET_ + '\'!F6');
  sh.getRange('E7').setFormula('=\'' + MEETING_SHEET_ + '\'!H6');
  for (var c = 2; c <= 5; c++) {
    var col = columnLetter_(c);
    sh.getRange(8, c).setFormula('=IF(OR(' + col + '6="",' + col + '7=""),"",' + col + '7-' + col + '6+1)');
  }
  sh.getRange('B6:E7').setNumberFormat('yyyy/m/d');

  // フィルタ参照
  sh.getRange('A10').setValue('フィルタ（会議用）').setFontWeight('bold');
  sh.getRange('A11').setValue('性別');
  sh.getRange('B11').setFormula('=\'' + MEETING_SHEET_ + '\'!L4');
  sh.getRange('A12').setValue('年代');
  sh.getRange('B12').setFormula('=\'' + MEETING_SHEET_ + '\'!L5');
  sh.getRange('A13').setValue('契約');
  sh.getRange('B13').setFormula('=\'' + MEETING_SHEET_ + '\'!L6');
  sh.getRange('A14').setValue('手続');
  sh.getRange('B14').setFormula('=\'' + MEETING_SHEET_ + '\'!L7');

  // KPI表
  var metrics = [
    ['期間開始時会員', 'start'],
    ['期間終了時会員', 'end'],
    ['期間内入会', 'join'],
    ['期間内退会', 'leave'],
    ['純増', 'net'],
    ['退会率', 'leaveRate'],
    ['1日あたり入会', 'joinPerDay'],
    ['1日あたり退会', 'leavePerDay'],
    ['再入会（手続=復）', 'rejoin'],
    ['再入会候補', 'rejoinCand'],
    ['短期退会≦6ヶ月', 'short6'],
    ['短期退会≦6ヶ月率', 'short6Rate']
  ];
  sh.getRange('A16').setValue('主要KPI（届出日／退会届出日ベース・日単位）').setFontWeight('bold');
  sh.getRange(17, 1, 1, 5).setValues([['指標', 'A', 'B', 'C', 'D']])
    .setBackground(theme.ink).setFontColor('#ffffff').setFontWeight('bold');
  for (var i = 0; i < metrics.length; i++) {
    var row = 18 + i;
    sh.getRange(row, 1).setValue(metrics[i][0]).setFontWeight('bold').setBackground(theme.soft);
    for (var p = 0; p < 4; p++) {
      sh.getRange(row, 2 + p).setFormula(meetingMetricFormula_(metrics[i][1], p));
    }
  }
  sh.getRange(18, 2, 5, 4).setNumberFormat('#,##0');
  sh.getRange(23, 2, 1, 4).setNumberFormat('0.00%');
  sh.getRange(24, 2, 2, 4).setNumberFormat('0.00');
  sh.getRange(26, 2, 2, 4).setNumberFormat('#,##0');
  sh.getRange(28, 2, 1, 4).setNumberFormat('#,##0');
  sh.getRange(29, 2, 1, 4).setNumberFormat('0.00%');

  // 年代 × 期間
  sh.getRange('A31').setValue('入会の年代（期間内入会・届出日）').setFontWeight('bold');
  var ages = ['10代', '20代', '30代', '40代', '50代', '60代', '70代以上'];
  var ageMin = [0, 20, 30, 40, 50, 60, 70];
  var ageMax = [20, 30, 40, 50, 60, 70, 200];
  sh.getRange(32, 1, 1, 5).setValues([['年代', 'A', 'B', 'C', 'D']]).setBackground(theme.mid).setFontColor('#ffffff');
  for (var a = 0; a < ages.length; a++) {
    sh.getRange(33 + a, 1).setValue(ages[a]).setBackground(theme.soft);
    for (var p2 = 0; p2 < 4; p2++) {
      sh.getRange(33 + a, 2 + p2).setFormula(meetingAgeJoinFormula_(p2, ageMin[a], ageMax[a]));
    }
  }
  sh.getRange(33, 2, 7, 4).setNumberFormat('#,##0');

  // 性別
  sh.getRange('A42').setValue('入会の性別（期間内入会）').setFontWeight('bold');
  sh.getRange(43, 1, 1, 5).setValues([['性別', 'A', 'B', 'C', 'D']]).setBackground(theme.mid).setFontColor('#ffffff');
  ['男', '女'].forEach(function (g, gi) {
    sh.getRange(44 + gi, 1).setValue(g).setBackground(theme.soft);
    for (var p3 = 0; p3 < 4; p3++) {
      sh.getRange(44 + gi, 2 + p3).setFormula(meetingGenderJoinFormula_(p3, g));
    }
  });

  // 退会理由
  sh.getRange('A48').setValue('退会理由（期間内退会・コード）').setFontWeight('bold');
  sh.getRange('A49').setValue('※累計退会の理由列は記号（M/A/N…）。意味一覧は会議用の定義欄を参照。')
    .setFontSize(8).setFontColor(theme.mute);
  sh.getRange(50, 1, 1, 5).setValues([['理由', 'A', 'B', 'C', 'D']]).setBackground(theme.mid).setFontColor('#ffffff');
  var reasons = ['M', 'A', 'N', 'U', 'B', 'D', 'W', 'S', 'X', 'R', 'V', 'I', 'T', 'その他'];
  for (var ri = 0; ri < reasons.length; ri++) {
    sh.getRange(51 + ri, 1).setValue(reasons[ri]).setBackground(theme.soft);
    for (var p4 = 0; p4 < 4; p4++) {
      sh.getRange(51 + ri, 2 + p4).setFormula(meetingReasonFormula_(p4, reasons[ri]));
    }
  }

  // 在籍期間区分（退会者の在籍期間月）
  sh.getRange('A67').setValue('退会者の在籍期間（期間内退会）').setFontWeight('bold');
  var tens = [
    ['0〜3ヶ月', 0, 3],
    ['4〜6ヶ月', 4, 6],
    ['7〜12ヶ月', 7, 12],
    ['13〜24ヶ月', 13, 24],
    ['25〜36ヶ月', 25, 36],
    ['37ヶ月以上', 37, 9999]
  ];
  sh.getRange(68, 1, 1, 5).setValues([['在籍期間', 'A', 'B', 'C', 'D']]).setBackground(theme.mid).setFontColor('#ffffff');
  for (var ti = 0; ti < tens.length; ti++) {
    sh.getRange(69 + ti, 1).setValue(tens[ti][0]).setBackground(theme.soft);
    for (var p5 = 0; p5 < 4; p5++) {
      sh.getRange(69 + ti, 2 + p5).setFormula(meetingTenureFormula_(p5, tens[ti][1], tens[ti][2]));
    }
  }

  sh.getRange('A1').setNote('このシートは会議用の計算置き場。値を直接書き換えない。');
  try { sh.hideSheet(); } catch (eH2) {}
}

/** フィルタ付き期間条件の共通部品（SUMPRODUCT用） */
function meetingFilterParts_() {
  var M = "'" + MEETING_MASTER_ + "'!";
  return {
    gender: '(($B$11="すべて")+(' + M + 'E$2:E$' + MEETING_END_ + '=$B$11))',
    contract: '(($B$13="すべて")+(' + M + 'I$2:I$' + MEETING_END_ + '=$B$13))',
    proc: '(($B$14="すべて")+(' + M + 'L$2:L$' + MEETING_END_ + '=$B$14))'
  };
}

function meetingPeriodCol_(p) {
  return columnLetter_(2 + p); // B=期間A ... E=期間D
}

function meetingMetricFormula_(kind, p) {
  var M = "'" + MEETING_MASTER_ + "'!";
  var col = meetingPeriodCol_(p);
  var start = col + '6';
  var end = col + '7';
  var days = col + '8';
  var f = meetingFilterParts_();
  var base = f.gender + '*' + f.contract + '*' + f.proc;
  var joinDate = M + 'M$2:M$' + MEETING_END_;
  var leaveDate = M + 'O$2:O$' + MEETING_END_;
  var rejoin = M + 'T$2:T$' + MEETING_END_;
  var tenure = M + 'R$2:R$' + MEETING_END_;

  function sp(expr) {
    return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + expr + ')*(' + base + ')),0))';
  }

  if (kind === 'join') {
    return sp('(' + joinDate + '<>"")*(' + joinDate + '>=' + start + ')*(' + joinDate + '<=' + end + ')');
  }
  if (kind === 'leave') {
    return sp('(' + leaveDate + '<>"")*(' + leaveDate + '>=' + start + ')*(' + leaveDate + '<=' + end + ')');
  }
  if (kind === 'start') {
    return sp('(' + joinDate + '<>"")*(' + joinDate + '<' + start + ')*((' + leaveDate + '="")+(' + leaveDate + '>=' + start + '))');
  }
  if (kind === 'end') {
    return sp('(' + joinDate + '<>"")*(' + joinDate + '<=' + end + ')*((' + leaveDate + '="")+(' + leaveDate + '>' + end + '))');
  }
  if (kind === 'net') {
    return '=IF(OR(' + start + '="",' + end + '=""),"",' + col + '20-' + col + '21)';
  }
  if (kind === 'leaveRate') {
    return '=IF(OR(' + start + '="",' + end + '="",' + col + '18=0),"",' + col + '21/' + col + '18)';
  }
  if (kind === 'joinPerDay') {
    return '=IF(OR(' + start + '="",' + end + '="",' + days + '=0),"",' + col + '20/' + days + ')';
  }
  if (kind === 'leavePerDay') {
    return '=IF(OR(' + start + '="",' + end + '="",' + days + '=0),"",' + col + '21/' + days + ')';
  }
  if (kind === 'rejoin') {
    return sp('(' + joinDate + '<>"")*(' + joinDate + '>=' + start + ')*(' + joinDate + '<=' + end + ')*(' + rejoin + '="再入会")');
  }
  if (kind === 'rejoinCand') {
    return sp('(' + joinDate + '<>"")*(' + joinDate + '>=' + start + ')*(' + joinDate + '<=' + end + ')*(' + rejoin + '="再入会候補")');
  }
  if (kind === 'short6') {
    return sp('(' + leaveDate + '<>"")*(' + leaveDate + '>=' + start + ')*(' + leaveDate + '<=' + end + ')*(IFERROR(VALUE(' + tenure + '),999)<=6)');
  }
  if (kind === 'short6Rate') {
    return '=IF(OR(' + start + '="",' + end + '="",' + col + '21=0),"",' + col + '28/' + col + '21)';
  }
  return '=""';
}

function meetingAgeJoinFormula_(p, amin, amax) {
  var M = "'" + MEETING_MASTER_ + "'!";
  var col = meetingPeriodCol_(p);
  var start = col + '6';
  var end = col + '7';
  var f = meetingFilterParts_();
  var joinDate = M + 'M$2:M$' + MEETING_END_;
  var birth = M + 'V$2:V$' + MEETING_END_;
  var ageExpr = '(YEAR(' + end + ')-YEAR(' + birth + '))';
  var ageCond = amax >= 200
    ? '(' + birth + '<>"")*(' + ageExpr + '>=' + amin + ')'
    : '(' + birth + '<>"")*(' + ageExpr + '>=' + amin + ')*(' + ageExpr + '<' + amax + ')';
  return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + joinDate + '<>"")*(' + joinDate + '>=' + start + ')*(' + joinDate + '<=' + end + ')*(' + ageCond + ')*(' +
    f.gender + ')*(' + f.contract + ')*(' + f.proc + ')),0))';
}

function meetingGenderJoinFormula_(p, gender) {
  var M = "'" + MEETING_MASTER_ + "'!";
  var col = meetingPeriodCol_(p);
  var start = col + '6';
  var end = col + '7';
  var f = meetingFilterParts_();
  var joinDate = M + 'M$2:M$' + MEETING_END_;
  return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + joinDate + '<>"")*(' + joinDate + '>=' + start + ')*(' + joinDate + '<=' + end + ')*(' +
    M + 'E$2:E$' + MEETING_END_ + '="' + gender + '")*(' + f.contract + ')*(' + f.proc + ')),0))';
}

function meetingReasonFormula_(p, code) {
  var M = "'" + MEETING_MASTER_ + "'!";
  var col = meetingPeriodCol_(p);
  var start = col + '6';
  var end = col + '7';
  var f = meetingFilterParts_();
  var leaveDate = M + 'O$2:O$' + MEETING_END_;
  var reason = M + 'Q$2:Q$' + MEETING_END_;
  if (code === 'その他') {
    var known = '("M","A","N","U","B","D","W","S","X","R","V","I","T")';
    return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + leaveDate + '<>"")*(' + leaveDate + '>=' + start + ')*(' + leaveDate + '<=' + end + ')*(' +
      reason + '<>"")*(ISNA(MATCH(' + reason + ',{' + known.slice(1, -1) + '},0)))*(' + f.gender + ')*(' + f.contract + ')*(' + f.proc + ')),0))';
  }
  return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + leaveDate + '<>"")*(' + leaveDate + '>=' + start + ')*(' + leaveDate + '<=' + end + ')*(' +
    reason + '="' + code + '")*(' + f.gender + ')*(' + f.contract + ')*(' + f.proc + ')),0))';
}

function meetingTenureFormula_(p, tmin, tmax) {
  var M = "'" + MEETING_MASTER_ + "'!";
  var col = meetingPeriodCol_(p);
  var start = col + '6';
  var end = col + '7';
  var f = meetingFilterParts_();
  var leaveDate = M + 'O$2:O$' + MEETING_END_;
  var tenure = M + 'R$2:R$' + MEETING_END_;
  return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + leaveDate + '<>"")*(' + leaveDate + '>=' + start + ')*(' + leaveDate + '<=' + end + ')*' +
    '(IFERROR(VALUE(' + tenure + '),-1)>=' + tmin + ')*(IFERROR(VALUE(' + tenure + '),-1)<=' + tmax + ')*(' +
    f.gender + ')*(' + f.contract + ')*(' + f.proc + ')),0))';
}

/** 会議用ダッシュボード本体（入力セルだけ色付き。あとは参照表示） */
function meetingBuildDash_(sh, theme) {
  sh.setTabColor('#B91C1C');
  sh.setHiddenGridlines(true);
  ensureSheetColumns_(sh, 20);
  var inputBg = '#FFF3CD';
  var C = "'" + MEETING_CALC_ + "'!";

  // タイトル
  sh.getRange('A1:M1').merge().setValue('JOYFIT24経堂　会議用 会員分析 Dashboard')
    .setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground(theme.ink)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(1, 36);

  sh.getRange('A2').setValue('① 期間A〜Dの開始日・終了日を入力　② 必要なら右の条件を指定　③ 下の数字を確認（触るのは黄色セルだけ）')
    .setFontSize(10).setFontColor(theme.mute);
  sh.getRange('A2:M2').merge();

  // 期間設定
  sh.getRange('A3').setValue('比較期間設定').setFontWeight('bold').setFontSize(12).setFontColor(theme.ink);
  sh.getRange(4, 1, 1, 9).setValues([['項目', '期間A', '', '期間B', '', '期間C', '', '期間D', '']]);
  sh.getRange('A4').setBackground(theme.ink).setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange('B4').setValue('今回').setBackground(inputBg);
  sh.getRange('D4').setValue('前年同期間').setBackground(inputBg);
  sh.getRange('F4').setValue('施策前').setBackground(inputBg);
  sh.getRange('H4').setValue('参考').setBackground(inputBg);
  ['B4', 'D4', 'F4', 'H4'].forEach(function (a1) {
    sh.getRange(a1).setFontWeight('bold').setNote('期間の呼び名（自由入力）');
  });

  sh.getRange('A5').setValue('開始日').setFontWeight('bold').setBackground(theme.soft);
  sh.getRange('A6').setValue('終了日').setFontWeight('bold').setBackground(theme.soft);
  sh.getRange('A7').setValue('日数').setFontWeight('bold').setBackground(theme.soft);
  // 初期値: 今年10/13-12/1 と 前年同期間
  sh.getRange('B5').setValue(new Date(2026, 9, 13)).setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('B6').setValue(new Date(2026, 11, 1)).setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('D5').setValue(new Date(2025, 9, 13)).setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('D6').setValue(new Date(2025, 11, 1)).setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('F5').setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('F6').setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('H5').setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('H6').setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('B7').setFormula('=IF(OR(B5="",B6=""),"",B6-B5+1)');
  sh.getRange('D7').setFormula('=IF(OR(D5="",D6=""),"",D6-D5+1)');
  sh.getRange('F7').setFormula('=IF(OR(F5="",F6=""),"",F6-F5+1)');
  sh.getRange('H7').setFormula('=IF(OR(H5="",H6=""),"",H6-H5+1)');
  sh.getRange('B5:H6').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build()
  );

  sh.getRange('A8').setValue('比較パターン').setFontWeight('bold').setBackground(theme.soft);
  sh.getRange('B8').setValue('自由設定').setBackground(inputBg);
  sh.getRange('B8').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['自由設定', '前月比較', '前年同期間', '過去30日', '過去90日'], true).build());
  sh.getRange('C8').setValue('※パターン変更時は開始・終了を手で合わせてください（自動上書きはしません）')
    .setFontSize(8).setFontColor(theme.mute);
  sh.getRange('C8:I8').merge();

  // 分析条件
  sh.getRange('K3').setValue('分析条件').setFontWeight('bold').setFontSize(12);
  sh.getRange('K3:M3').merge().setBackground(theme.ink).setFontColor('#ffffff');
  var filters = [
    ['性別', ['すべて', '男', '女']],
    ['年代', ['すべて', '10代', '20代', '30代', '40代', '50代', '60代', '70代以上']],
    ['契約名称', ['すべて']],
    ['手続', ['すべて', '新', '復']]
  ];
  for (var fi = 0; fi < filters.length; fi++) {
    sh.getRange(4 + fi, 11).setValue(filters[fi][0]).setFontWeight('bold').setBackground(theme.soft);
    sh.getRange(4 + fi, 12).setValue(filters[fi][1][0]).setBackground(inputBg);
    sh.getRange(4 + fi, 12).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(filters[fi][1], true).build()
    );
  }
  // 契約名称はマスタからUNIQUEで補完リスト（データ検証は固定「すべて」＋手入力可）
  sh.getRange('L6').clearDataValidations();
  sh.getRange('L6').setValue('すべて').setBackground(inputBg)
    .setNote('契約名をそのまま入力。空欄や「すべて」で絞り込みなし。');
  sh.getRange('K8').setValue('日付の定義').setFontWeight('bold').setBackground(theme.soft);
  sh.getRange('L8').setValue('届出日／退会届出日（日単位）').setFontSize(9);
  sh.getRange('K9').setValue('制約').setFontWeight('bold').setBackground(theme.soft);
  sh.getRange('L9').setValue('利用開始年月・退会年月は月単位のため日次分析には使わない').setFontSize(8).setFontColor(theme.red);
  sh.getRange('K9:M9').merge();

  // KPI
  sh.getRange('A10').setValue('重要KPI比較（届出日・退会届出日ベース）').setFontWeight('bold').setFontSize(12);
  sh.getRange(11, 1, 1, 9).setValues([[
    '指標', '期間A', '期間B', 'A−B', 'A/B', '期間C', '期間D', 'C−D', 'C/D'
  ]]).setBackground(theme.ink).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');

  var kpiLabels = [
    '期間開始時会員', '期間終了時会員', '期間内入会', '期間内退会', '純増',
    '退会率', '1日あたり入会', '1日あたり退会',
    '再入会（手続=復）', '再入会候補', '短期退会≦6ヶ月', '短期退会≦6ヶ月率'
  ];
  // calc rows 18..29
  for (var ki = 0; ki < kpiLabels.length; ki++) {
    var r = 12 + ki;
    var src = 18 + ki;
    sh.getRange(r, 1).setValue(kpiLabels[ki]).setFontWeight('bold').setBackground(theme.soft).setFontSize(9);
    sh.getRange(r, 2).setFormula('=' + C + 'B' + src);
    sh.getRange(r, 3).setFormula('=' + C + 'C' + src);
    sh.getRange(r, 4).setFormula('=IF(OR(B' + r + '="",C' + r + '=""),"",B' + r + '-C' + r + ')');
    sh.getRange(r, 5).setFormula('=IF(OR(B' + r + '="",C' + r + '="",C' + r + '=0),"",B' + r + '/C' + r + '-1)');
    sh.getRange(r, 6).setFormula('=' + C + 'D' + src);
    sh.getRange(r, 7).setFormula('=' + C + 'E' + src);
    sh.getRange(r, 8).setFormula('=IF(OR(F' + r + '="",G' + r + '=""),"",F' + r + '-G' + r + ')');
    sh.getRange(r, 9).setFormula('=IF(OR(F' + r + '="",G' + r + '="",G' + r + '=0),"",F' + r + '/G' + r + '-1)');
  }
  sh.getRange(12, 2, 5, 8).setNumberFormat('#,##0');
  sh.getRange(16, 2, 1, 8).setNumberFormat('+#,##0;-#,##0;0'); // 純増 row 16 = 12+4
  sh.getRange(17, 2, 1, 3).setNumberFormat('0.00%');
  sh.getRange(17, 4, 1, 1).setNumberFormat('+0.00%;-0.00%;0%');
  sh.getRange(17, 5, 1, 1).setNumberFormat('+0.0%;-0.0%;0%');
  sh.getRange(17, 6, 1, 2).setNumberFormat('0.00%');
  sh.getRange(17, 8, 1, 1).setNumberFormat('+0.00%;-0.00%;0%');
  sh.getRange(17, 9, 1, 1).setNumberFormat('+0.0%;-0.0%;0%');
  sh.getRange(18, 2, 2, 8).setNumberFormat('0.00');
  sh.getRange(20, 2, 2, 8).setNumberFormat('#,##0');
  sh.getRange(22, 2, 1, 8).setNumberFormat('#,##0');
  sh.getRange(23, 2, 1, 3).setNumberFormat('0.00%');
  sh.getRange(23, 5, 1, 1).setNumberFormat('+0.0%;-0.0%;0%');
  sh.getRange(23, 6, 1, 2).setNumberFormat('0.00%');
  sh.getRange(23, 9, 1, 1).setNumberFormat('+0.0%;-0.0%;0%');
  // ratio columns for count rows
  sh.getRange(12, 5, 5, 1).setNumberFormat('+0.0%;-0.0%;0%');
  sh.getRange(12, 9, 5, 1).setNumberFormat('+0.0%;-0.0%;0%');
  sh.getRange(12, 4, 5, 1).setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange(12, 8, 5, 1).setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange(22, 1, 1, 9).setBackground('#FCE8E8'); // 短期退会率を目立たせ
  sh.getRange(23, 1, 1, 9).setBackground('#FCE8E8');

  // 注目ポイント
  sh.getRange('A25').setValue('今回の注目ポイント（期間A vs B・数式判定）').setFontWeight('bold').setFontSize(12);
  sh.getRange('A26').setFormula(
    '=IF(OR(B14="",C14=""),"期間A・Bの開始日と終了日を入れてください",' +
    'IF(AND(ISNUMBER(B14),ISNUMBER(C14),C14<>0,B14/C14-1<=-0.2),"⚠ 入会が比較期間より20%以上減少","")&' +
    'IF(AND(ISNUMBER(B17),ISNUMBER(C17),B17-C17>=0.01)," ⚠ 退会率が+1pt以上悪化","")&' +
    'IF(AND(ISNUMBER(B23),ISNUMBER(C23),B23-C23>=0.01)," ⚠ 6ヶ月以内退会率が悪化","")&' +
    'IF(AND(ISNUMBER(B20),ISNUMBER(C20),C20<>0,B20/C20-1>=0.2)," ✓ 再入会（復）が増加","")&' +
    'IF(AND(ISNUMBER(B16),B16>0)," ✓ 純増がプラス","")&' +
    'IF(AND(ISNUMBER(B16),B16<0)," ⚠ 純増がマイナス",""))'
  ).setFontSize(11).setWrap(true);
  sh.getRange('A26:M26').merge();
  sh.setRowHeight(26, 40);

  // 年代
  sh.getRange('A28').setValue('入会の年代比較').setFontWeight('bold').setFontSize(12);
  sh.getRange(29, 1, 1, 9).setValues([['年代', 'A人数', 'A構成比', 'B人数', 'B構成比', '人数差', '構成比差', 'C人数', 'D人数']])
    .setBackground(theme.mid).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
  var ages = ['10代', '20代', '30代', '40代', '50代', '60代', '70代以上'];
  for (var ai = 0; ai < ages.length; ai++) {
    var ar = 30 + ai;
    var srcA = 33 + ai;
    sh.getRange(ar, 1).setValue(ages[ai]).setBackground(theme.soft).setFontWeight('bold');
    sh.getRange(ar, 2).setFormula('=' + C + 'B' + srcA);
    sh.getRange(ar, 3).setFormula('=IF(SUM($B$30:$B$36)=0,"",B' + ar + '/SUM($B$30:$B$36))');
    sh.getRange(ar, 4).setFormula('=' + C + 'C' + srcA);
    sh.getRange(ar, 5).setFormula('=IF(SUM($D$30:$D$36)=0,"",D' + ar + '/SUM($D$30:$D$36))');
    sh.getRange(ar, 6).setFormula('=IF(OR(B' + ar + '="",D' + ar + '=""),"",B' + ar + '-D' + ar + ')');
    sh.getRange(ar, 7).setFormula('=IF(OR(C' + ar + '="",E' + ar + '=""),"",C' + ar + '-E' + ar + ')');
    sh.getRange(ar, 8).setFormula('=' + C + 'D' + srcA);
    sh.getRange(ar, 9).setFormula('=' + C + 'E' + srcA);
  }
  sh.getRange(30, 2, 7, 1).setNumberFormat('#,##0');
  sh.getRange(30, 3, 7, 1).setNumberFormat('0.0%');
  sh.getRange(30, 4, 7, 1).setNumberFormat('#,##0');
  sh.getRange(30, 5, 7, 1).setNumberFormat('0.0%');
  sh.getRange(30, 6, 7, 1).setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange(30, 7, 7, 1).setNumberFormat('+0.0%;-0.0%;0%');
  sh.getRange(30, 8, 7, 2).setNumberFormat('#,##0');

  // 性別
  sh.getRange('A38').setValue('入会の性別比較').setFontWeight('bold').setFontSize(12);
  sh.getRange(39, 1, 1, 7).setValues([['性別', 'A', 'A構成比', 'B', 'B構成比', '差', '構成比差']])
    .setBackground(theme.mid).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
  ['男', '女'].forEach(function (g, gi) {
    var gr = 40 + gi;
    sh.getRange(gr, 1).setValue(g).setBackground(theme.soft).setFontWeight('bold');
    sh.getRange(gr, 2).setFormula('=' + C + 'B' + (44 + gi));
    sh.getRange(gr, 3).setFormula('=IF(SUM($B$40:$B$41)=0,"",B' + gr + '/SUM($B$40:$B$41))');
    sh.getRange(gr, 4).setFormula('=' + C + 'C' + (44 + gi));
    sh.getRange(gr, 5).setFormula('=IF(SUM($D$40:$D$41)=0,"",D' + gr + '/SUM($D$40:$D$41))');
    sh.getRange(gr, 6).setFormula('=IF(OR(B' + gr + '="",D' + gr + '=""),"",B' + gr + '-D' + gr + ')');
    sh.getRange(gr, 7).setFormula('=IF(OR(C' + gr + '="",E' + gr + '=""),"",C' + gr + '-E' + gr + ')');
  });
  sh.getRange(40, 2, 2, 1).setNumberFormat('#,##0');
  sh.getRange(40, 3, 2, 1).setNumberFormat('0.0%');
  sh.getRange(40, 4, 2, 1).setNumberFormat('#,##0');
  sh.getRange(40, 5, 2, 1).setNumberFormat('0.0%');
  sh.getRange(40, 6, 2, 1).setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange(40, 7, 2, 1).setNumberFormat('+0.0%;-0.0%;0%');

  // 退会理由
  sh.getRange('A43').setValue('退会理由（期間内退会・コード別）').setFontWeight('bold').setFontSize(12);
  sh.getRange(44, 1, 1, 5).setValues([['理由コード', 'A', 'B', '差', '備考']])
    .setBackground(theme.mid).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
  var reasons = ['M', 'A', 'N', 'U', 'B', 'D', 'W', 'S', 'X', 'R', 'V', 'I', 'T', 'その他'];
  for (var ri = 0; ri < reasons.length; ri++) {
    var rr = 45 + ri;
    sh.getRange(rr, 1).setValue(reasons[ri]).setBackground(theme.soft);
    sh.getRange(rr, 2).setFormula('=' + C + 'B' + (51 + ri));
    sh.getRange(rr, 3).setFormula('=' + C + 'C' + (51 + ri));
    sh.getRange(rr, 4).setFormula('=IF(OR(B' + rr + '="",C' + rr + '=""),"",B' + rr + '-C' + rr + ')');
  }
  sh.getRange(45, 2, 14, 3).setNumberFormat('#,##0');
  sh.getRange(45, 4, 14, 1).setNumberFormat('+#,##0;-#,##0;0');

  // 在籍期間
  sh.getRange('A60').setValue('退会者の在籍期間').setFontWeight('bold').setFontSize(12);
  sh.getRange(61, 1, 1, 5).setValues([['区分', 'A', 'B', '差', 'A構成比']])
    .setBackground(theme.mid).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
  var tens = ['0〜3ヶ月', '4〜6ヶ月', '7〜12ヶ月', '13〜24ヶ月', '25〜36ヶ月', '37ヶ月以上'];
  for (var ti = 0; ti < tens.length; ti++) {
    var tr = 62 + ti;
    sh.getRange(tr, 1).setValue(tens[ti]).setBackground(theme.soft);
    sh.getRange(tr, 2).setFormula('=' + C + 'B' + (69 + ti));
    sh.getRange(tr, 3).setFormula('=' + C + 'C' + (69 + ti));
    sh.getRange(tr, 4).setFormula('=IF(OR(B' + tr + '="",C' + tr + '=""),"",B' + tr + '-C' + tr + ')');
    sh.getRange(tr, 5).setFormula('=IF(SUM($B$62:$B$67)=0,"",B' + tr + '/SUM($B$62:$B$67))');
  }
  sh.getRange(62, 2, 6, 3).setNumberFormat('#,##0');
  sh.getRange(62, 4, 6, 1).setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange(62, 5, 6, 1).setNumberFormat('0.0%');

  // 増加・減少 TOP（年代から簡易）
  sh.getRange('A69').setValue('期間A−Bで増えた年代 TOP／減った年代 TOP').setFontWeight('bold').setFontSize(12);
  sh.getRange('A70').setFormula(
    '="増加1位: "&INDEX(SORT(A30:F36,6,FALSE),1,1)&" ("&TEXT(INDEX(SORT(A30:F36,6,FALSE),1,6),"+#,##0;-#,##0;0")&")　"' +
    '&"減少1位: "&INDEX(SORT(A30:F36,6,TRUE),1,1)&" ("&TEXT(INDEX(SORT(A30:F36,6,TRUE),1,6),"+#,##0;-#,##0;0")&")"'
  ).setFontSize(11);
  sh.getRange('A70:M70').merge();

  // 会員検索
  sh.getRange('A72').setValue('会員検索（個人情報はここだけ表示）').setFontWeight('bold').setFontSize(12);
  sh.getRange('A73').setValue('会員番号または氏名').setBackground(theme.soft).setFontWeight('bold');
  sh.getRange('B73').setBackground(inputBg).setNote('番号または氏名を入力');
  sh.getRange('A74').setValue('結果');
  sh.getRange('B74').setFormula(
    '=IF($B$73="","",IFERROR("番号 "&INDEX(\'' + MEETING_MASTER_ + '\'!B:B,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!B:B,0))&' +
    '" / "&INDEX(\'' + MEETING_MASTER_ + '\'!C:C,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!B:B,0))&' +
    '" / "&INDEX(\'' + MEETING_MASTER_ + '\'!E:E,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!B:B,0))&' +
    '" / 届出 "&TEXT(INDEX(\'' + MEETING_MASTER_ + '\'!M:M,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!B:B,0)),"yyyy/m/d")&' +
    '" / 退会届出 "&IFERROR(TEXT(INDEX(\'' + MEETING_MASTER_ + '\'!O:O,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!B:B,0)),"yyyy/m/d"),"在籍")&' +
    '" / 手続 "&INDEX(\'' + MEETING_MASTER_ + '\'!L:L,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!B:B,0))&' +
    '" / TEL "&INDEX(\'' + MEETING_MASTER_ + '\'!H:H,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!B:B,0)),' +
    'IFERROR("番号 "&INDEX(\'' + MEETING_MASTER_ + '\'!B:B,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!C:C,0))&' +
    '" / "&INDEX(\'' + MEETING_MASTER_ + '\'!C:C,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!C:C,0))&' +
    '" / "&INDEX(\'' + MEETING_MASTER_ + '\'!E:E,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!C:C,0))&' +
    '" / 届出 "&TEXT(INDEX(\'' + MEETING_MASTER_ + '\'!M:M,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!C:C,0)),"yyyy/m/d")&' +
    '" / 退会届出 "&IFERROR(TEXT(INDEX(\'' + MEETING_MASTER_ + '\'!O:O,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!C:C,0)),"yyyy/m/d"),"在籍")&' +
    '" / 手続 "&INDEX(\'' + MEETING_MASTER_ + '\'!L:L,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!C:C,0))&' +
    '" / TEL "&INDEX(\'' + MEETING_MASTER_ + '\'!H:H,MATCH($B$73,\'' + MEETING_MASTER_ + '\'!C:C,0)),"見つかりません")))'
  ).setFontSize(10).setWrap(true);
  sh.getRange('B74:M75').merge();

  // グラフ用（入会・退会・純増・退会率）
  sh.getRange('A77').setValue('グラフ用データ（非表示可）').setFontWeight('bold');
  sh.getRange(78, 1, 1, 5).setValues([['指標', 'A', 'B', 'C', 'D']]);
  sh.getRange(79, 1, 4, 1).setValues([['入会'], ['退会'], ['純増'], ['退会率']]);
  sh.getRange('B79').setFormula('=B14');
  sh.getRange('C79').setFormula('=C14');
  sh.getRange('D79').setFormula('=F14');
  sh.getRange('E79').setFormula('=G14');
  sh.getRange('B80').setFormula('=B15');
  sh.getRange('C80').setFormula('=C15');
  sh.getRange('D80').setFormula('=F15');
  sh.getRange('E80').setFormula('=G15');
  sh.getRange('B81').setFormula('=B16');
  sh.getRange('C81').setFormula('=C16');
  sh.getRange('D81').setFormula('=F16');
  sh.getRange('E81').setFormula('=G16');
  sh.getRange('B82').setFormula('=B17');
  sh.getRange('C82').setFormula('=C17');
  sh.getRange('D82').setFormula('=F17');
  sh.getRange('E82').setFormula('=G17');
  try {
    var chart = sh.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(78, 1, 4, 5))
      .setNumHeaders(1)
      .setOption('title', '入会・退会・純増（4期間）')
      .setOption('legend', { position: 'top' })
      .setOption('colors', ['#111111', '#6B6B6B', '#9E9E9E', '#B91C1C'])
      .setPosition(77, 7, 0, 0)
      .build();
    sh.insertChart(chart);
  } catch (eCh) {}

  // 定義・参照元
  sh.getRange('A85').setValue('指標定義・参照元（追跡用）').setFontWeight('bold').setFontSize(12)
    .setBackground(theme.ink).setFontColor('#ffffff');
  sh.getRange('A85:M85').merge();
  var defs = [
    ['指標', '参照シート', '参照列', '条件／定義', '備考'],
    ['期間内入会', '分析用_会員マスタ ← 累計入会データ', '届出日(M)', '開始日≦届出日≦終了日', '日単位。利用開始年月は使わない'],
    ['期間内退会', '分析用_会員マスタ ← 累計退会データ', '退会届出日(O)', '開始日≦退会届出日≦終了日', '日単位。退会年月は月単位'],
    ['期間開始時会員', '同上', '届出日・退会届出日', '届出日<開始 かつ（退会空白 or 退会≧開始）', ''],
    ['期間終了時会員', '同上', '届出日・退会届出日', '届出日≦終了 かつ（退会空白 or 退会＞終了）', ''],
    ['純増', '計算', '', '期間内入会−期間内退会', ''],
    ['退会率', '計算', '', '期間内退会÷期間開始時会員', '開始時0は空欄'],
    ['再入会', '累計入会 手続列', '手続=復', '期間内入会かつ手続「復」', '確定のみ'],
    ['再入会候補', '分析用_会員マスタ', '同一番号キーで過去届出あり', '氏名電話の曖昧一致は未実装', '精度注意'],
    ['短期退会≦6ヶ月', '累計退会 在籍期間', '在籍期間月≦6', '期間内退会者のうち', ''],
    ['年代', '生年月日', '期間終了年−生年', '期間内入会者の年齢帯', '完全な満年齢ではない簡易'],
    ['退会理由', '累計退会 理由', '記号コード', 'M/A/N… 意味表は店舗側定義に従う', ''],
    ['休会者数', '—', '—', '任意期間の休会履歴が元データに無い', '日報の当月休会のみ別途参照可'],
    ['会員動向の入会', '累計入会 利用開始年月', '月単位', '本ダッシュボードの届出日集計とは一致しない場合あり', '重要']
  ];
  sh.getRange(86, 1, defs.length, 5).setValues(defs);
  sh.getRange(86, 1, 1, 5).setBackground(theme.mid).setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(87, 1, defs.length - 1, 5).setFontSize(9).setWrap(true);
  sh.setColumnWidth(1, 150);
  sh.setColumnWidths(2, 8, 72);
  sh.setColumnWidth(4, 70);
  sh.setColumnWidth(5, 70);
  sh.setColumnWidth(11, 90);
  sh.setColumnWidth(12, 160);
  sh.setFrozenRows(2);

  // 品質チェック簡易
  sh.getRange('A102').setValue('データ品質（簡易）').setFontWeight('bold').setFontSize(12);
  sh.getRange('A103').setFormula('="入会マスタ行数: "&COUNTA(\'' + MEETING_MASTER_ + '\'!B2:B' + MEETING_END_ + ')');
  sh.getRange('A104').setFormula('="届出日なし: "&COUNTBLANK(FILTER(\'' + MEETING_MASTER_ + '\'!M2:M' + MEETING_END_ + ',\'' + MEETING_MASTER_ + '\'!B2:B' + MEETING_END_ + '<>""))');
  sh.getRange('A105').setFormula('="退会届出が入会届出より前（疑い）: "&SUMPRODUCT((\'' + MEETING_MASTER_ + '\'!O2:O' + MEETING_END_ + '<>"")*(\'' + MEETING_MASTER_ + '\'!M2:M' + MEETING_END_ + '<>"")*(\'' + MEETING_MASTER_ + '\'!O2:O' + MEETING_END_ + '<\'' + MEETING_MASTER_ + '\'!M2:M' + MEETING_END_ + '))');
  sh.getRange('A106').setValue('※コホート・契約TOP・クロス分析の詳細は、必要なら次版で分析用シートに追加。')
    .setFontSize(9).setFontColor(theme.mute);
}

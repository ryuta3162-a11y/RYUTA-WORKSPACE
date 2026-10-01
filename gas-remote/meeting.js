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
var MEETING_VER_ = '4';

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
  var after = afterName ? ss.getSheetByName(afterName) : null;
  var idx = after ? after.getIndex() : ss.getSheets().length;
  var old = ss.getSheetByName(name);
  if (old) {
    // 古い入力規則・結合が残ると途中で落ちるので削除して作り直す
    try { ss.deleteSheet(old); } catch (eDel) {
      old.clear();
      try { old.clearDataValidations(); } catch (eV) {}
      try { old.getCharts().forEach(function (ch) { old.removeChart(ch); }); } catch (eC) {}
      try { old.clearConditionalFormatRules(); } catch (eR) {}
      return old;
    }
  }
  var sh = ss.insertSheet(name, Math.min(idx, ss.getSheets().length));
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
  // S 番号キー（先に置く。退会照合で使う）
  sh.getRange('S2').setFormula('=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",IFERROR(TEXT(VALUE(' + J + 'F2:F' + end + '),"0"),' + J + 'F2:F' + end + ')))');
  // M 届出日（日付シリアルに正規化。文字列だと期間比較が全部0になる）
  sh.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' +
    'IF(ISNUMBER(' + J + 'M2:M' + end + '),' + J + 'M2:M' + end + ',' +
    'IFERROR(DATEVALUE(' + J + 'M2:M' + end + '),IFERROR(DATEVALUE(TEXT(' + J + 'M2:M' + end + ',"yyyy/m/d")),""))))'
  );
  // N 利用開始年月（月単位。会員動向の入会定義に近い）
  sh.getRange('N2').setFormula(
    '=ARRAYFORMULA(IF(' + J + 'F2:F' + end + '="","",' +
    'IF(ISNUMBER(' + J + 'N2:N' + end + '),' + J + 'N2:N' + end + ',' +
    'IFERROR(DATEVALUE(' + J + 'N2:N' + end + '),IFERROR(DATEVALUE(TEXT(' + J + 'N2:N' + end + ',"yyyy/m/d")),' + J + 'N2:N' + end + '))))'
  );
  // O〜R 退会側（番号キーSで単段照合。ネストXLOOKUPのARRAYFORMULAは#ERROR!になる）
  sh.getRange('O2').setFormula(meetingLeaveMapFormula_(L, end, 'Q', true));
  sh.getRange('P2').setFormula(meetingLeaveMapFormula_(L, end, 'R', true));
  sh.getRange('Q2').setFormula(meetingLeaveMapFormula_(L, end, 'S', false));
  sh.getRange('R2').setFormula(meetingLeaveMapFormula_(L, end, 'P', false));
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
  sh.getRange('A2').clearContent();

  // 期間パラメータを会議用から参照
  sh.getRange('A4').setValue('期間パラメータ（会議用から）').setFontWeight('bold');
  sh.getRange(5, 1, 1, 5).setValues([['', 'A', 'B', 'C', 'D']])
    .setBackground(theme.ink).setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange('A6').setValue('開始');
  sh.getRange('A7').setValue('終了');
  sh.getRange('A8').setValue('日数');
  sh.getRange('B6').setFormula('=\'' + MEETING_SHEET_ + '\'!B4');
  sh.getRange('C6').setFormula('=\'' + MEETING_SHEET_ + '\'!C4');
  sh.getRange('D6').setFormula('=\'' + MEETING_SHEET_ + '\'!D4');
  sh.getRange('E6').setFormula('=\'' + MEETING_SHEET_ + '\'!E4');
  sh.getRange('B7').setFormula('=\'' + MEETING_SHEET_ + '\'!B5');
  sh.getRange('C7').setFormula('=\'' + MEETING_SHEET_ + '\'!C5');
  sh.getRange('D7').setFormula('=\'' + MEETING_SHEET_ + '\'!D5');
  sh.getRange('E7').setFormula('=\'' + MEETING_SHEET_ + '\'!E5');
  for (var c = 2; c <= 5; c++) {
    var col = columnLetter_(c);
    sh.getRange(8, c).setFormula('=IF(OR(' + col + '6="",' + col + '7=""),"",' + col + '7-' + col + '6+1)');
  }
  sh.getRange('B6:E7').setNumberFormat('yyyy/m/d');

  // フィルタは会議用UIから外したので常に「すべて」
  sh.getRange('A10').setValue('フィルタ（固定）').setFontWeight('bold');
  sh.getRange('A11').setValue('性別');
  sh.getRange('B11').setValue('すべて');
  sh.getRange('A12').setValue('年代');
  sh.getRange('B12').setValue('すべて');
  sh.getRange('A13').setValue('契約');
  sh.getRange('B13').setValue('すべて');
  sh.getRange('A14').setValue('手続');
  sh.getRange('B14').setValue('すべて');

  // KPI表
  var metrics = [
    ['期間開始時会員', 'start'],
    ['期間終了時会員', 'end'],
    ['期間内入会', 'join'],
    ['期間内退会', 'leave'],
    ['期間内移籍', 'transfer'],
    ['純増', 'net'],
    ['退会率', 'leaveRate'],
    ['1日あたり入会', 'joinPerDay'],
    ['1日あたり退会', 'leavePerDay'],
    ['再入会（手続=復）', 'rejoin'],
    ['再入会候補', 'rejoinCand'],
    ['短期退会≦6ヶ月', 'short6'],
    ['短期退会≦6ヶ月率', 'short6Rate']
  ];
  sh.getRange('A16').setValue('主要KPI').setFontWeight('bold');
  sh.getRange(17, 1, 1, 5).setValues([['指標', 'A', 'B', 'C', 'D']])
    .setBackground(theme.ink).setFontColor('#ffffff').setFontWeight('bold');
  for (var i = 0; i < metrics.length; i++) {
    var row = 18 + i;
    sh.getRange(row, 1).setValue(metrics[i][0]).setFontWeight('bold').setBackground(theme.soft);
    for (var p = 0; p < 4; p++) {
      sh.getRange(row, 2 + p).setFormula(meetingMetricFormula_(metrics[i][1], p));
    }
  }
  // rows shifted by +1 due to transfer: join=20, leave=21, transfer=22, net=23...
  sh.getRange(18, 2, 6, 4).setNumberFormat('#,##0');
  sh.getRange(24, 2, 1, 4).setNumberFormat('0.00%');
  sh.getRange(25, 2, 2, 4).setNumberFormat('0.00');
  sh.getRange(27, 2, 2, 4).setNumberFormat('#,##0');
  sh.getRange(29, 2, 1, 4).setNumberFormat('#,##0');
  sh.getRange(30, 2, 1, 4).setNumberFormat('0.00%');

  // 年代 × 期間（行32据え置き）
  sh.getRange('A32').setValue('入会の年代').setFontWeight('bold');
  var ages = ['10代', '20代', '30代', '40代', '50代', '60代', '70代以上'];
  var ageMin = [0, 20, 30, 40, 50, 60, 70];
  var ageMax = [20, 30, 40, 50, 60, 70, 200];
  sh.getRange(33, 1, 1, 5).setValues([['年代', 'A', 'B', 'C', 'D']]).setBackground(theme.mid).setFontColor('#ffffff');
  for (var a = 0; a < ages.length; a++) {
    sh.getRange(34 + a, 1).setValue(ages[a]).setBackground(theme.soft);
    for (var p2 = 0; p2 < 4; p2++) {
      sh.getRange(34 + a, 2 + p2).setFormula(meetingAgeJoinFormula_(p2, ageMin[a], ageMax[a]));
    }
  }
  sh.getRange(34, 2, 7, 4).setNumberFormat('#,##0');

  // 性別
  sh.getRange('A42').setValue('入会の性別').setFontWeight('bold');
  sh.getRange(43, 1, 1, 5).setValues([['性別', 'A', 'B', 'C', 'D']]).setBackground(theme.mid).setFontColor('#ffffff');
  ['男', '女'].forEach(function (g, gi) {
    sh.getRange(44 + gi, 1).setValue(g).setBackground(theme.soft);
    for (var p3 = 0; p3 < 4; p3++) {
      sh.getRange(44 + gi, 2 + p3).setFormula(meetingGenderJoinFormula_(p3, g));
    }
  });

  // 退会理由（行番号を+1: 51→52）
  sh.getRange('A48').setValue('退会理由').setFontWeight('bold');
  sh.getRange(49, 1, 1, 5).setValues([['理由', 'A', 'B', 'C', 'D']]).setBackground(theme.mid).setFontColor('#ffffff');
  var reasons = ['M', 'A', 'N', 'U', 'B', 'D', 'W', 'S', 'X', 'R', 'V', 'I', 'T', 'その他'];
  for (var ri = 0; ri < reasons.length; ri++) {
    sh.getRange(50 + ri, 1).setValue(reasons[ri]).setBackground(theme.soft);
    for (var p4 = 0; p4 < 4; p4++) {
      sh.getRange(50 + ri, 2 + p4).setFormula(meetingReasonFormula_(p4, reasons[ri]));
    }
  }

  // 在籍期間
  sh.getRange('A65').setValue('退会者の在籍期間').setFontWeight('bold');
  var tens = [
    ['0〜3ヶ月', 0, 3],
    ['4〜6ヶ月', 4, 6],
    ['7〜12ヶ月', 7, 12],
    ['13〜24ヶ月', 13, 24],
    ['25〜36ヶ月', 25, 36],
    ['37ヶ月以上', 37, 9999]
  ];
  sh.getRange(66, 1, 1, 5).setValues([['在籍期間', 'A', 'B', 'C', 'D']]).setBackground(theme.mid).setFontColor('#ffffff');
  for (var ti = 0; ti < tens.length; ti++) {
    sh.getRange(67 + ti, 1).setValue(tens[ti][0]).setBackground(theme.soft);
    for (var p5 = 0; p5 < 4; p5++) {
      sh.getRange(67 + ti, 2 + p5).setFormula(meetingTenureFormula_(p5, tens[ti][1], tens[ti][2]));
    }
  }

  sh.getRange('A1').setNote('計算置き場');
  try { sh.hideSheet(); } catch (eH2) {}
}

/**
 * 退会マスタ照合（行ごとMAP）。asDate=true なら日付シリアルへ正規化。
 */
function meetingLeaveMapFormula_(leaveSheetRef, end, resultCol, asDate) {
  var body = asDate
    ? 'IF(d="","" ,IF(ISNUMBER(d),d,IFERROR(DATEVALUE(d),IFERROR(DATEVALUE(TEXT(d,"yyyy/m/d")),""))))'
    : 'd';
  return '=MAP(S2:S' + end + ',LAMBDA(k,IF(k="","",LET(d,IFERROR(XLOOKUP(k,' +
    leaveSheetRef + 'T2:T' + end + ',' + leaveSheetRef + resultCol + '2:' + resultCol + end + ',""),""),' + body + '))))';
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
  // 日付は N() で数値化（空は0）。文字日付の比較ズレを防ぐ
  var jd = 'IFERROR(N(' + joinDate + '),0)';
  var ld = 'IFERROR(N(' + leaveDate + '),0)';

  if (kind === 'join') {
    return sp('(' + joinDate + '<>"")*(' + jd + '>=N(' + start + '))*(' + jd + '<=N(' + end + '))');
  }
  if (kind === 'leave') {
    return sp('(' + leaveDate + '<>"")*(' + ld + '>=N(' + start + '))*(' + ld + '<=N(' + end + '))');
  }
  if (kind === 'transfer') {
    // 移籍は日報の当月値のみ（累計に日次移籍が無い）。期間が今月を含むときだけ表示
    return '=IF(OR(' + start + '="",' + end + '=""),"",' +
      'IF(AND(N(' + start + ')<=EOMONTH(TODAY(),0),N(' + end + ')>=DATE(YEAR(TODAY()),MONTH(TODAY()),1)),' +
      'IFERROR(VALUE(\'日報\'!D14),0),""))';
  }
  if (kind === 'start') {
    return sp('(' + joinDate + '<>"")*(' + jd + '<N(' + start + '))*((' + leaveDate + '="")+(' + ld + '>=N(' + start + ')))');
  }
  if (kind === 'end') {
    return sp('(' + joinDate + '<>"")*(' + jd + '<=N(' + end + '))*((' + leaveDate + '="")+(' + ld + '>N(' + end + ')))');
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
    return sp('(' + joinDate + '<>"")*(' + jd + '>=N(' + start + '))*(' + jd + '<=N(' + end + '))*(' + rejoin + '="再入会")');
  }
  if (kind === 'rejoinCand') {
    return sp('(' + joinDate + '<>"")*(' + jd + '>=N(' + start + '))*(' + jd + '<=N(' + end + '))*(' + rejoin + '="再入会候補")');
  }
  if (kind === 'short6') {
    return sp('(' + leaveDate + '<>"")*(' + ld + '>=N(' + start + '))*(' + ld + '<=N(' + end + '))*(IFERROR(VALUE(' + tenure + '),999)<=6)');
  }
  if (kind === 'short6Rate') {
    return '=IF(OR(' + start + '="",' + end + '="",' + col + '21=0),"",' + col + '29/' + col + '21)';
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
  var jd = 'IFERROR(N(' + joinDate + '),0)';
  var ageExpr = '(YEAR(' + end + ')-YEAR(' + birth + '))';
  var ageCond = amax >= 200
    ? '(' + birth + '<>"")*(' + ageExpr + '>=' + amin + ')'
    : '(' + birth + '<>"")*(' + ageExpr + '>=' + amin + ')*(' + ageExpr + '<' + amax + ')';
  return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + joinDate + '<>"")*(' + jd + '>=N(' + start + '))*(' + jd + '<=N(' + end + '))*(' + ageCond + ')*(' +
    f.gender + ')*(' + f.contract + ')*(' + f.proc + ')),0))';
}

function meetingGenderJoinFormula_(p, gender) {
  var M = "'" + MEETING_MASTER_ + "'!";
  var col = meetingPeriodCol_(p);
  var start = col + '6';
  var end = col + '7';
  var f = meetingFilterParts_();
  var joinDate = M + 'M$2:M$' + MEETING_END_;
  var jd = 'IFERROR(N(' + joinDate + '),0)';
  return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + joinDate + '<>"")*(' + jd + '>=N(' + start + '))*(' + jd + '<=N(' + end + '))*(' +
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
  var ld = 'IFERROR(N(' + leaveDate + '),0)';
  if (code === 'その他') {
    var known = '("M","A","N","U","B","D","W","S","X","R","V","I","T")';
    return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + leaveDate + '<>"")*(' + ld + '>=N(' + start + '))*(' + ld + '<=N(' + end + '))*(' +
      reason + '<>"")*(ISNA(MATCH(' + reason + ',{' + known.slice(1, -1) + '},0)))*(' + f.gender + ')*(' + f.contract + ')*(' + f.proc + ')),0))';
  }
  return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + leaveDate + '<>"")*(' + ld + '>=N(' + start + '))*(' + ld + '<=N(' + end + '))*(' +
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
  var ld = 'IFERROR(N(' + leaveDate + '),0)';
  return '=IF(OR(' + start + '="",' + end + '=""),"",IFERROR(SUMPRODUCT((' + leaveDate + '<>"")*(' + ld + '>=N(' + start + '))*(' + ld + '<=N(' + end + '))*' +
    '(IFERROR(VALUE(' + tenure + '),-1)>=' + tmin + ')*(IFERROR(VALUE(' + tenure + '),-1)<=' + tmax + ')*(' +
    f.gender + ')*(' + f.contract + ')*(' + f.proc + ')),0))';
}

/** 会議用ダッシュボード本体：触るのはチェックと黄色の日付だけ */
function meetingBuildDash_(sh, theme) {
  sh.setTabColor('#9E9E9E');
  sh.setHiddenGridlines(true);
  ensureSheetColumns_(sh, 36);
  try { sh.getRange(1, 1, Math.min(sh.getMaxRows(), 120), 32).breakApart(); } catch (eB) {}
  try { sh.clearDataValidations(); } catch (eV) {}
  var inputBg = '#FFF3CD';
  var soft = '#F5F5F5';
  var mid = '#EEEEEE';
  var line = '#E0E0E0';
  var ink = '#424242';
  var mute = '#9E9E9E';
  var C = "'" + MEETING_CALC_ + "'!";
  // 表示スイッチ（チェック）
  var SHOW_JOIN = 'OR($D$1=TRUE,$J$1=TRUE)';
  var SHOW_LEAVE = 'OR($F$1=TRUE,$J$1=TRUE)';
  var SHOW_MOVE = 'OR($H$1=TRUE,$J$1=TRUE)';

  // ---- 1行目：比較対象チェック ----
  sh.getRange('A1').setValue('会議用').setFontSize(13).setFontWeight('bold').setFontColor(ink);
  sh.getRange('C1').setValue('入会').setFontWeight('bold').setFontSize(10).setBackground(soft)
    .setHorizontalAlignment('right');
  sh.getRange('D1').insertCheckboxes().setValue(true).setBackground(inputBg);
  sh.getRange('E1').setValue('退会').setFontWeight('bold').setFontSize(10).setBackground(soft)
    .setHorizontalAlignment('right');
  sh.getRange('F1').insertCheckboxes().setValue(true).setBackground(inputBg);
  sh.getRange('G1').setValue('移籍').setFontWeight('bold').setFontSize(10).setBackground(soft)
    .setHorizontalAlignment('right');
  sh.getRange('H1').insertCheckboxes().setValue(false).setBackground(inputBg);
  sh.getRange('I1').setValue('全部').setFontWeight('bold').setFontSize(10).setBackground(soft)
    .setHorizontalAlignment('right');
  sh.getRange('J1').insertCheckboxes().setValue(false).setBackground(inputBg);
  sh.getRange('L1').setValue('かんたん期間').setFontWeight('bold').setFontSize(9).setBackground(soft);
  sh.getRange('M1').setValue('手入力').setBackground(inputBg).setFontSize(9);
  sh.getRange('M1').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['手入力', '今月', '先月', '過去30日', '過去90日', '前年同期間'], true).build());
  sh.getRange('AE1').setFormula('=TEXT(NOW(),"M/d HH:mm")').setFontSize(8).setFontColor(mute)
    .setHorizontalAlignment('right');
  sh.setRowHeight(1, 26);

  // ---- 期間：開始・終了だけ×4 ----
  sh.getRange(3, 1, 1, 5).setValues([['', '①', '②', '③', '④']])
    .setBackground(mid).setFontColor(ink).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  sh.getRange('B3').setValue('今回').setBackground(inputBg);
  sh.getRange('C3').setValue('前年').setBackground(inputBg);
  sh.getRange('D3').setValue('③').setBackground(inputBg);
  sh.getRange('E3').setValue('④').setBackground(inputBg);
  sh.getRange('A4').setValue('開始').setFontWeight('bold').setBackground(soft).setFontSize(9);
  sh.getRange('A5').setValue('終了').setFontWeight('bold').setBackground(soft).setFontSize(9);
  sh.getRange('A6').setValue('日数').setFontWeight('bold').setBackground(soft).setFontSize(9);

  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth();
  sh.getRange('B4').setValue(new Date(y, m, 1)).setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('B5').setValue(new Date(y, m + 1, 0)).setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('C4').setValue(new Date(y - 1, m, 1)).setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('C5').setValue(new Date(y - 1, m + 1, 0)).setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('D4').setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('D5').setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('E4').setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('E5').setBackground(inputBg).setNumberFormat('yyyy/m/d');
  sh.getRange('B4:E5').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build()
  );
  ['B', 'C', 'D', 'E'].forEach(function (col) {
    sh.getRange(col + '6').setFormula('=IF(OR(' + col + '4="",' + col + '5=""),"",' + col + '5-' + col + '4+1)')
      .setNumberFormat('0').setFontSize(9).setHorizontalAlignment('center');
  });
  sh.getRange(3, 1, 4, 5).setBorder(true, true, true, true, true, true, line, SpreadsheetApp.BorderStyle.SOLID);
  for (var rh = 3; rh <= 6; rh++) sh.setRowHeight(rh, 20);

  // ---- KPI（チェックに応じて表示） ----
  // calc: 18開始 19終了 20入会 21退会 22移籍 23純増 24退会率 25入会/日 26退会/日 27復 28候補 29短期 30短期率
  sh.getRange(8, 1, 1, 7).setValues([['指標', '①', '②', '①−②', '③', '④', '③−④']])
    .setBackground(mid).setFontColor(ink).setFontWeight('bold').setFontSize(8)
    .setHorizontalAlignment('center');

  function kpiRow(row, label, calcRow, showExpr, fmt) {
    sh.getRange(row, 1).setValue(label).setFontWeight('bold').setBackground(soft).setFontSize(9);
    sh.getRange(row, 2).setFormula('=IF(' + showExpr + ',' + C + 'B' + calcRow + ',"")');
    sh.getRange(row, 3).setFormula('=IF(' + showExpr + ',' + C + 'C' + calcRow + ',"")');
    sh.getRange(row, 4).setFormula('=IF(OR(B' + row + '="",C' + row + '=""),"",B' + row + '-C' + row + ')');
    sh.getRange(row, 5).setFormula('=IF(' + showExpr + ',' + C + 'D' + calcRow + ',"")');
    sh.getRange(row, 6).setFormula('=IF(' + showExpr + ',' + C + 'E' + calcRow + ',"")');
    sh.getRange(row, 7).setFormula('=IF(OR(E' + row + '="",F' + row + '=""),"",E' + row + '-F' + row + ')');
    if (fmt) sh.getRange(row, 2, 1, 6).setNumberFormat(fmt);
  }

  kpiRow(9, '入会', 20, SHOW_JOIN, '#,##0');
  kpiRow(10, '退会', 21, SHOW_LEAVE, '#,##0');
  kpiRow(11, '移籍', 22, SHOW_MOVE, '#,##0');
  // 純増は入会か退会を見ているとき
  kpiRow(12, '純増', 23, 'OR(' + SHOW_JOIN + ',' + SHOW_LEAVE + ')', '+#,##0;-#,##0;0');
  kpiRow(13, '退会率', 24, SHOW_LEAVE, '0.0%');
  kpiRow(14, '入会/日', 25, SHOW_JOIN, '0.00');
  kpiRow(15, '退会/日', 26, SHOW_LEAVE, '0.00');
  kpiRow(16, '再入会', 27, SHOW_JOIN, '#,##0');
  kpiRow(17, '短期≦6ヶ月', 29, SHOW_LEAVE, '#,##0');
  sh.getRange(8, 1, 10, 7).setBorder(true, true, true, true, true, true, line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(9, 2, 9, 6).setFontSize(10).setHorizontalAlignment('center');
  for (var rCompact = 9; rCompact <= 17; rCompact++) sh.setRowHeight(rCompact, 20);

  // ---- 右：年代（入会時） ----
  sh.getRange(3, 9, 1, 5).setValues([['年代', '①', '②', '差', '構成①']])
    .setBackground(mid).setFontColor(ink).setFontWeight('bold').setFontSize(8)
    .setHorizontalAlignment('center');
  var ages = ['10代', '20代', '30代', '40代', '50代', '60代', '70代+'];
  for (var ai = 0; ai < ages.length; ai++) {
    var ar = 4 + ai;
    var src = 34 + ai;
    sh.getRange(ar, 9).setValue(ages[ai]).setBackground(soft).setFontWeight('bold').setFontSize(8);
    sh.getRange(ar, 10).setFormula('=IF(' + SHOW_JOIN + ',' + C + 'B' + src + ',"")');
    sh.getRange(ar, 11).setFormula('=IF(' + SHOW_JOIN + ',' + C + 'C' + src + ',"")');
    sh.getRange(ar, 12).setFormula('=IF(OR(J' + ar + '="",K' + ar + '=""),"",J' + ar + '-K' + ar + ')');
    sh.getRange(ar, 13).setFormula('=IF(OR(J' + ar + '="",SUM($J$4:$J$10)=0),"",J' + ar + '/SUM($J$4:$J$10))');
  }
  sh.getRange(4, 10, 7, 3).setNumberFormat('#,##0').setFontSize(9).setHorizontalAlignment('center');
  sh.getRange(4, 13, 7, 1).setNumberFormat('0%').setFontSize(8).setFontColor(mute);
  sh.getRange(3, 9, 8, 5).setBorder(true, true, true, true, true, true, line, SpreadsheetApp.BorderStyle.SOLID);

  // ---- 右：性別 ----
  sh.getRange(3, 15, 1, 4).setValues([['性別', '①', '②', '差']])
    .setBackground(mid).setFontColor(ink).setFontWeight('bold').setFontSize(8)
    .setHorizontalAlignment('center');
  ['男', '女'].forEach(function (g, gi) {
    var gr = 4 + gi;
    sh.getRange(gr, 15).setValue(g).setBackground(soft).setFontWeight('bold').setFontSize(8);
    sh.getRange(gr, 16).setFormula('=IF(' + SHOW_JOIN + ',' + C + 'B' + (44 + gi) + ',"")');
    sh.getRange(gr, 17).setFormula('=IF(' + SHOW_JOIN + ',' + C + 'C' + (44 + gi) + ',"")');
    sh.getRange(gr, 18).setFormula('=IF(OR(P' + gr + '="",Q' + gr + '=""),"",P' + gr + '-Q' + gr + ')');
  });
  sh.getRange(4, 16, 2, 3).setNumberFormat('#,##0').setFontSize(9).setHorizontalAlignment('center');
  sh.getRange(3, 15, 3, 4).setBorder(true, true, true, true, true, true, line, SpreadsheetApp.BorderStyle.SOLID);

  // ---- 右：退会理由（横） ----
  sh.getRange(12, 9).setValue('退会理由').setFontWeight('bold').setFontSize(8).setBackground(mid);
  sh.getRange(12, 9, 1, 5).setBackground(mid);
  sh.getRange(13, 9, 1, 5).setValues([['理由', '①', '②', '差', '']])
    .setBackground(soft).setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center');
  var reasons = ['M', 'A', 'N', 'U', 'B', 'D', 'W', 'S', 'X', 'R', 'V', 'I', 'T', 'その他'];
  for (var ri = 0; ri < reasons.length; ri++) {
    var rr = 14 + ri;
    var srcR = 50 + ri;
    sh.getRange(rr, 9).setValue(reasons[ri]).setBackground(soft).setFontSize(8);
    sh.getRange(rr, 10).setFormula('=IF(' + SHOW_LEAVE + ',' + C + 'B' + srcR + ',"")');
    sh.getRange(rr, 11).setFormula('=IF(' + SHOW_LEAVE + ',' + C + 'C' + srcR + ',"")');
    sh.getRange(rr, 12).setFormula('=IF(OR(J' + rr + '="",K' + rr + '=""),"",J' + rr + '-K' + rr + ')');
  }
  sh.getRange(14, 10, 14, 3).setNumberFormat('#,##0').setFontSize(8).setHorizontalAlignment('center');
  sh.getRange(12, 9, 16, 4).setBorder(true, true, true, true, true, true, line, SpreadsheetApp.BorderStyle.SOLID);

  // ---- 在籍期間（退会）横並び ----
  sh.getRange(12, 15, 1, 4).setValues([['在籍', '①', '②', '差']])
    .setBackground(mid).setFontColor(ink).setFontWeight('bold').setFontSize(8)
    .setHorizontalAlignment('center');
  var tens = ['0〜3', '4〜6', '7〜12', '13〜24', '25〜36', '37+'];
  for (var ti = 0; ti < tens.length; ti++) {
    var tr = 13 + ti;
    sh.getRange(tr, 15).setValue(tens[ti]).setBackground(soft).setFontSize(8);
    sh.getRange(tr, 16).setFormula('=IF(' + SHOW_LEAVE + ',' + C + 'B' + (67 + ti) + ',"")');
    sh.getRange(tr, 17).setFormula('=IF(' + SHOW_LEAVE + ',' + C + 'C' + (67 + ti) + ',"")');
    sh.getRange(tr, 18).setFormula('=IF(OR(P' + tr + '="",Q' + tr + '=""),"",P' + tr + '-Q' + tr + ')');
  }
  sh.getRange(13, 16, 6, 3).setNumberFormat('#,##0').setFontSize(8).setHorizontalAlignment('center');
  sh.getRange(12, 15, 7, 4).setBorder(true, true, true, true, true, true, line, SpreadsheetApp.BorderStyle.SOLID);

  // 列幅
  sh.setColumnWidth(1, 72);
  sh.setColumnWidths(2, 6, 52);
  sh.setColumnWidth(8, 12);
  sh.setColumnWidth(9, 48);
  sh.setColumnWidths(10, 4, 42);
  sh.setColumnWidth(14, 12);
  sh.setColumnWidth(15, 48);
  sh.setColumnWidths(16, 3, 42);
  sh.setColumnWidth(19, 12);
  sh.setColumnWidths(20, 12, 36);
  sh.getRange(1, 1, 28, 18).setFontFamily('Meiryo');
  sh.setFrozenRows(1);
  meetingEnsureOnEditTrigger_();
}

/** かんたん期間・全部チェックの操作 */
function meetingOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== MEETING_SHEET_) return;
    var r = e.range.getRow();
    var c = e.range.getColumn();
    // 全部チェック
    if (r === 1 && c === 10) {
      var all = sh.getRange('J1').getValue() === true;
      sh.getRange('D1').setValue(all);
      sh.getRange('F1').setValue(all);
      sh.getRange('H1').setValue(all);
      return;
    }
    if (r === 1 && (c === 4 || c === 6 || c === 8)) {
      var join = sh.getRange('D1').getValue() === true;
      var leave = sh.getRange('F1').getValue() === true;
      var move = sh.getRange('H1').getValue() === true;
      if (!(join && leave && move)) sh.getRange('J1').setValue(false);
      return;
    }
    // かんたん期間 → ①②の開始終了を埋める
    if (r === 1 && c === 13) {
      var mode = String(sh.getRange('M1').getValue() || '');
      var today = new Date();
      var y = today.getFullYear();
      var m = today.getMonth();
      var a0, a1, b0, b1;
      if (mode === '今月') {
        a0 = new Date(y, m, 1); a1 = new Date(y, m + 1, 0);
        b0 = new Date(y, m - 1, 1); b1 = new Date(y, m, 0);
      } else if (mode === '先月') {
        a0 = new Date(y, m - 1, 1); a1 = new Date(y, m, 0);
        b0 = new Date(y, m - 2, 1); b1 = new Date(y, m - 1, 0);
      } else if (mode === '過去30日') {
        a1 = today; a0 = new Date(today.getTime() - 29 * 86400000);
        b1 = new Date(a0.getTime() - 86400000); b0 = new Date(b1.getTime() - 29 * 86400000);
      } else if (mode === '過去90日') {
        a1 = today; a0 = new Date(today.getTime() - 89 * 86400000);
        b1 = new Date(a0.getTime() - 86400000); b0 = new Date(b1.getTime() - 89 * 86400000);
      } else if (mode === '前年同期間') {
        a0 = new Date(y, m, 1); a1 = new Date(y, m + 1, 0);
        b0 = new Date(y - 1, m, 1); b1 = new Date(y - 1, m + 1, 0);
      } else {
        return;
      }
      sh.getRange('B4').setValue(a0).setNumberFormat('yyyy/m/d');
      sh.getRange('B5').setValue(a1).setNumberFormat('yyyy/m/d');
      sh.getRange('C4').setValue(b0).setNumberFormat('yyyy/m/d');
      sh.getRange('C5').setValue(b1).setNumberFormat('yyyy/m/d');
      sh.getRange('B3').setValue('①');
      sh.getRange('C3').setValue('②');
    }
  } catch (err) {}
}

function meetingEnsureOnEditTrigger_() {
  try {
    var ss = SpreadsheetApp.getActive();
    if (!ss) return;
    var exists = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getHandlerFunction() === 'meetingOnEdit';
    });
    if (!exists) {
      ScriptApp.newTrigger('meetingOnEdit').forSpreadsheet(ss).onEdit().create();
    }
  } catch (eT) {}
}

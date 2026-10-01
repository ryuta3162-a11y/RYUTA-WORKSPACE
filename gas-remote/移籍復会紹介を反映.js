function 日報オプションをB1連動() {
  var book = SpreadsheetApp.openById('14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w');
  var sh = book.getSheetByName('日報');
  var op = book.getSheetByName('OP集計');
  var diag = book.getSheetByName('OP取込診断');
  var join = book.getSheetByName('入会・退会_データ');
  if (!sh) throw new Error('日報なし');
  if (!op || !diag || !join) throw new Error('OP集計 / OP取込診断 / 入会・退会_データ がありません');
  var b1 = String(sh.getRange('B1').getDisplayValue() || '').replace(/\D/g, '');
  var y = b1.length === 6 ? Number(b1.slice(0, 4)) : 2000 + Number(b1.slice(0, 2));
  var m = Number(b1.slice(-2));
  var label = y + '年' + m + '月';
  var lastLog = Math.max(op.getLastRow(), 2);
  var log = op.getRange(2, 9, lastLog - 1, 6).getDisplayValues();
  var seen = {};
  for (var i = 0; i < log.length; i++) {
    if (!/利用開始/.test(String(log[i][2] || ''))) continue;
    seen[String(log[i][5] || '') + '|' + String(log[i][4] || '')] = 1;
  }
  var jLast = Math.max(join.getLastRow() - 1, 0);
  var byName = {};
  if (jLast) {
    var jv = join.getRange(2, 1, jLast, 5).getDisplayValues();
    for (var j = 0; j < jv.length; j++) {
      if (String(jv[j][2] || '') === label) byName[String(jv[j][1] || '').trim()] = jv[j];
    }
  }
  var dLast = Math.max(diag.getLastRow() - 2, 0);
  var added = 0;
  if (dLast) {
    var dv = diag.getRange(3, 1, dLast, 5).getDisplayValues();
    var rows = [];
    for (var r = 0; r < dv.length; r++) {
      if (String(dv[r][2] || '') !== label) continue;
      var name = String(dv[r][0] || '').trim();
      var parts = String(dv[r][4] || '').split(/\s*\/\s*/);
      var jrow = byName[name];
      var mailId = jrow ? String(jrow[4] || '') : '';
      var stamp = jrow ? String(jrow[0] || '') : '';
      var ym = stamp.match(/(20\d{2})\D+(\d{1,2})/);
      var when = (ym && Number(ym[1]) === y && Number(ym[2]) === m)
        ? stamp
        : new Date(y, m - 1, 1, 0, 1, 0);
      for (var o = 0; o < parts.length; o++) {
        var opt = String(parts[o] || '').trim();
        if (!opt || opt === '(エラー)') continue;
        var key = mailId + '|' + opt;
        if (seen[key]) continue;
        seen[key] = 1;
        rows.push([when, name, '利用開始(新規入会)', opt, opt, mailId]);
      }
    }
    if (rows.length) {
      op.getRange(op.getLastRow() + 1, 9, rows.length, 6).setValues(rows);
      added = rows.length;
    }
  }
  var last = 21;
  var optNames = sh.getRange('B21:B40').getDisplayValues();
  for (var n = 0; n < optNames.length; n++) {
    if (String(optNames[n][0] || '').trim()) last = 21 + n;
  }
  var prefix = 'bcode,REGEXREPLACE($B$1&"","[^0-9]",""),yy,IF(LEN(bcode)=6,VALUE(LEFT(bcode,4)),2000+VALUE(LEFT(bcode,2))),mm,VALUE(RIGHT(bcode,2)),label,yy&"年"&mm&"月",startD,DATE(yy,mm,1),endD,DATE(yy,mm+1,1),';
  var dForms = [];
  var eForms = [];
  var fForms = [];
  for (var row = 21; row <= last; row++) {
    dForms.push(['=IFERROR(LET(' + prefix
      + 'newjoin,SUMPRODUCT((\'OP取込診断\'!C$3:C$200=label)*ISNUMBER(SEARCH(" / "&$B' + row + '&" / "," / "&\'OP取込診断\'!E$3:E$200&" / "))),'
      + 'addon,COUNTIFS(\'OP集計\'!$M:$M,$B' + row + ',\'OP集計\'!$K:$K,"利用開始(OP追加)",\'OP集計\'!$I:$I,">="&startD,\'OP集計\'!$I:$I,"<"&endD),'
      + 'newjoin+addon),0)']);
    eForms.push(['=IFERROR(LET(' + prefix
      + 'joiners,IFERROR(FILTER(\'OP取込診断\'!A$3:A$200,(\'OP取込診断\'!C$3:C$200=label)*(\'OP取込診断\'!A$3:A$200<>"")),{"__none__"}),'
      + 'SUMPRODUCT((\'OP集計\'!$M$2:$M$20000=$B' + row + ')*(\'OP集計\'!$K$2:$K$20000="利用停止")*(\'OP集計\'!$I$2:$I$20000>=startD)*(\'OP集計\'!$I$2:$I$20000<endD)*ISNA(MATCH(\'OP集計\'!$J$2:$J$20000,joiners,0)))),0)']);
    fForms.push(['=IFERROR(D' + row + '-E' + row + ',0)']);
  }
  var num = last - 20;
  sh.getRange(21, 4, num, 1).setFormulas(dForms).clearNote();
  sh.getRange(21, 5, num, 1).setFormulas(eForms).clearNote();
  sh.getRange(21, 6, num, 1).setFormulas(fForms);
  sh.getRange(21, 4, num, 3).setNumberFormat('0');
  SpreadsheetApp.flush();
  var names = sh.getRange(21, 2, num, 1).getDisplayValues();
  var vals = sh.getRange(21, 4, num, 3).getDisplayValues();
  var vip = ['', '', ''];
  var pilates = '';
  for (var p = 0; p < names.length; p++) {
    if (names[p][0] === '安心サポートVIP') vip = vals[p];
    if (/ピラティス/.test(String(names[p][0] || ''))) pilates = vals[p][0];
  }
  try {
    SpreadsheetApp.getUi().alert(
      '日報B1=' + sh.getRange('B1').getDisplayValue() + '（' + label + '）の入会月で契約を数えます。\n'
      + '安心サポートVIP　契約' + vip[0] + '　解約' + vip[1] + '\n'
      + 'ピラティスリフォーマー　契約' + pilates + '\n'
      + 'ログに足した新規OP ' + added + '件'
    );
  } catch (eUi) {}
  return { ok: true, added: added, vipContract: vip[0], pilates: pilates };
}
/**
 * 【ここだけ受付状況表に貼る】ワークスペースのスクリプトではない。
 *
 * プロジェクト: 経堂　受付状況表
 * scriptId:     1JvaBDxH580M-WCRa1bk5QZOAZd8veLJDbQccODMDni5Iv9_DHtZVRw7d
 *
 * 貼り方:
 *  1. 受付状況表を開く → 拡張機能 → Apps Script
 *  2. 既存のコードは消さない。一番下にこの関数を追加して保存
 *  3. 関数 OP月末入会を直す を選んで実行
 *
 * 原因: 10月取込が Gmail after:10/01 なので、9/30夜に届いた10月入会
 * （渡辺・朴）がログに載らず、契約1・ピラティス0 になる。
 * 診断は入会月で見ているので 3人とも取れている。こちらは診断＋入会メールIDでログと日報を直す。
 * 日報 C列（月初）は触らない。D/E/F（契約/解約/増減）だけ書く。
 *
 * 実行後の目安（2026年10月・入会3）:
 *  安心サポートVIP 契約3 / ピラティスリフォーマー 契約2
 *  セット7種は契約3（タンニング2・セルフエステ1 は性別で分かれる）
 */
function OP月末入会を直す() {
  var OPTION_LIST = [
    '安心サポート', '安心サポートVIP', '水素水', 'オンラインレッスン',
    '体組成計', '契約ロッカー1,500', 'レンタルマット', 'プロテイン12杯',
    'プロテイン無制限', 'プロテイン＋水素水', 'レンタルタオル', 'タンニング',
    'セルフエステ', 'ホットスタジオ', 'ヨガロッカー', 'ピラティスリフォーマー'
  ];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('日報') || !ss.getSheetByName('OP集計')) {
    ss = SpreadsheetApp.openById('14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w');
  }
  var nip = ss.getSheetByName('日報');
  var op = ss.getSheetByName('OP集計');
  var diag = ss.getSheetByName('OP取込診断');
  var join = ss.getSheetByName('入会・退会_データ');
  if (!nip || !op || !diag || !join) throw new Error('日報 / OP集計 / OP取込診断 / 入会・退会_データ がありません');

  var b1 = String(nip.getRange('B1').getDisplayValue() || '').replace(/\D/g, '');
  var y = b1.length === 6 ? Number(b1.slice(0, 4)) : 2000 + Number(b1.slice(0, 2));
  var m = Number(b1.slice(-2));
  var label = y + '年' + m + '月';
  var startD = new Date(y, m - 1, 1);
  var endD = new Date(y, m, 1);

  var jLast = Math.max(join.getLastRow() - 1, 0);
  var byName = {};
  if (jLast) {
    var jv = join.getRange(2, 1, jLast, 5).getDisplayValues();
    for (var j = 0; j < jv.length; j++) {
      if (String(jv[j][2] || '') === label) byName[String(jv[j][1] || '').trim()] = jv[j];
    }
  }

  var lastLog = Math.max(op.getLastRow(), 2);
  var log = op.getRange(2, 9, lastLog - 1, 6).getDisplayValues();
  var seenStart = {};
  var stopByOpt = {};
  OPTION_LIST.forEach(function (n) { stopByOpt[n] = 0; });
  var joinerNames = {};
  Object.keys(byName).forEach(function (n) { joinerNames[n] = 1; });

  for (var i = 0; i < log.length; i++) {
    var rec = log[i][0];
    var nm = String(log[i][1] || '').trim();
    var kind = String(log[i][2] || '');
    var opt = String(log[i][4] || '').trim();
    var mid = String(log[i][5] || '');
    if (/利用開始/.test(kind)) seenStart[mid + '|' + opt] = 1;
    var dt = rec instanceof Date ? rec : new Date(rec);
    var inMonth = dt && !isNaN(dt.getTime()) && dt >= startD && dt < endD;
    if (inMonth && kind === '利用停止' && opt && !joinerNames[nm]) {
      if (stopByOpt[opt] == null) stopByOpt[opt] = 0;
      stopByOpt[opt]++;
    }
  }

  var dLast = Math.max(diag.getLastRow() - 2, 0);
  var startByOpt = {};
  OPTION_LIST.forEach(function (n) { startByOpt[n] = 0; });
  var rows = [];
  if (dLast) {
    var dv = diag.getRange(3, 1, dLast, 5).getDisplayValues();
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
        var opt2 = String(parts[o] || '').trim();
        if (!opt2 || opt2 === '(エラー)') continue;
        if (startByOpt[opt2] == null) startByOpt[opt2] = 0;
        startByOpt[opt2]++;
        var key = mailId + '|' + opt2;
        if (seenStart[key]) continue;
        seenStart[key] = 1;
        rows.push([when, name, '利用開始(新規入会)', opt2, opt2, mailId]);
      }
    }
  }
  if (rows.length) op.getRange(op.getLastRow() + 1, 9, rows.length, 6).setValues(rows);

  var bCol = [];
  var dCol = [];
  var eCol = [];
  var fCol = [];
  for (var k = 0; k < OPTION_LIST.length; k++) {
    var nameK = OPTION_LIST[k];
    var start = Number(startByOpt[nameK] || 0);
    var stop = Number(stopByOpt[nameK] || 0);
    bCol.push([start]);
    dCol.push([start]);
    eCol.push([stop]);
    fCol.push([start - stop]);
  }
  op.getRange(3, 2, OPTION_LIST.length, 1).setValues(bCol);
  op.getRange(3, 4, OPTION_LIST.length, 1).setValues(dCol);
  op.getRange(3, 5, OPTION_LIST.length, 1).setValues(eCol);
  op.getRange(3, 6, OPTION_LIST.length, 1).setFormulas(
    dCol.map(function (_, idx) { return ['=D' + (idx + 3) + '-E' + (idx + 3)]; })
  );

  nip.getRange(21, 4, OPTION_LIST.length, 1).setValues(dCol).setNote('受付状況表：入会月（診断）で集計。月末夜の翌月入会も含む');
  nip.getRange(21, 5, OPTION_LIST.length, 1).setValues(eCol);
  nip.getRange(21, 6, OPTION_LIST.length, 1).setValues(fCol);
  nip.getRange(21, 4, OPTION_LIST.length, 3).setNumberFormat('0');

  var vip = startByOpt['安心サポートVIP'] || 0;
  var pira = startByOpt['ピラティスリフォーマー'] || 0;
  try {
    SpreadsheetApp.getUi().alert(
      label + ' の新規OPを診断の入会月で直しました。\n' +
      'ログに足した行 ' + rows.length + ' 件\n' +
      '安心サポートVIP 契約' + vip + '\n' +
      'ピラティスリフォーマー 契約' + pira
    );
  } catch (eUi) {}
  return { ok: true, month: label, added: rows.length, vip: vip, pilates: pira };
}

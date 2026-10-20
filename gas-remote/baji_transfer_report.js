/**
 * JOYFIT24経堂 → 馬事公苑 移籍整理（提出用）
 * メニュー「分析」→「馬事移籍｜提出用シートを生成」で実行。
 * 新規スプレッドシートを作成し、提出用4シートを自動出力する。
 */

var BAJI_REPORT_STORE_LAT_ = 35.647655101495786;
var BAJI_REPORT_STORE_LNG_ = 139.6316987253194;
var BAJI_REPORT_BAJI_LAT_ = 35.63957644630146;
var BAJI_REPORT_BAJI_LNG_ = 139.63186202531915;

/** @returns {Array<Object>} */
function bajiReportDistanceRows_() {
  // 徒歩距離は丁目代表地点→店舗（OSM foot）。距離差 = 馬事 - 経堂（＋＝経堂が近い）
  return [
    { id: '3082', area: '桜丘5丁目', k: 1145, b: 1460, util: 78.4 },
    { id: '3359', area: '桜丘4丁目', k: 1505, b: 1109, util: 100.0 },
    { id: '4805', area: '経堂5丁目', k: 141, b: 1254, util: 71.6 },
    { id: '5595', area: '桜丘1丁目', k: 663, b: 816, util: 86.3 },
    { id: '5711', area: '桜丘2丁目', k: 666, b: 974, util: 96.0 },
    { id: '6022', area: '桜丘3丁目', k: 1015, b: 517, util: 90.7 },
    { id: '6111', area: '桜丘3丁目', k: 1015, b: 517, util: 98.7 },
    { id: '6240', area: '赤堤2丁目', k: 2116, b: 2979, util: 97.5 },
    { id: '6265', area: '桜丘1丁目', k: 663, b: 816, util: 94.9 },
    { id: '6320', area: '船橋3丁目', k: 1178, b: 2004, util: 67.1 },
    { id: '6590', area: '桜丘3-9-31付近(常磐寮)', k: 855, b: 421, util: 100.0 },
    { id: '6620', area: '桜丘1丁目', k: 663, b: 816, util: 32.4 },
    { id: '6805', area: '桜丘3-9-31付近(常磐寮)', k: 855, b: 421, util: 89.5 }
  ];
}

/** @returns {Array<Object>} */
function bajiReportUsageRows_() {
  return [
    { id: '3082', baji: 40, kyodo: 11, other: 0, note: '' },
    { id: '3359', baji: 70, kyodo: 0, other: 0, note: '' },
    { id: '4805', baji: 53, kyodo: 21, other: 0, note: '' },
    { id: '5595', baji: 44, kyodo: 7, other: 0, note: '' },
    { id: '5711', baji: 96, kyodo: 4, other: 0, note: '' },
    { id: '6022', baji: 78, kyodo: 2, other: 6, note: '仙台東口6' },
    { id: '6111', baji: 75, kyodo: 0, other: 1, note: '他店舗1（要再確認）' },
    { id: '6240', baji: 78, kyodo: 2, other: 0, note: '' },
    { id: '6265', baji: 56, kyodo: 3, other: 0, note: '' },
    { id: '6320', baji: 47, kyodo: 21, other: 2, note: '上馬2' },
    { id: '6590', baji: 77, kyodo: 0, other: 0, note: '' },
    { id: '6620', baji: 33, kyodo: 69, other: 0, note: '利用時間異常の可能性→例外' },
    { id: '6805', baji: 68, kyodo: 8, other: 0, note: '50%OFF終了直後' },
    { id: '4090', baji: 8, kyodo: 64, other: 0, note: '住所分析除外' }
  ];
}

/**
 * メニュー実行入口。新規ブックを作成して提出用シートを出力する。
 */
function generateBajiTransferReport() {
  var ui = SpreadsheetApp.getUi();
  var title = '経堂→馬事公苑_移籍整理_提出用_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  var ss = SpreadsheetApp.create(title);
  bajiReportBuildAll_(ss);
  var url = ss.getUrl();
  PropertiesService.getUserProperties().setProperty('BAJI_TRANSFER_REPORT_URL', url);
  try {
    ui.alert(
      '提出用シートを作成しました',
      '手編集なしで出力済みです。\n\n' + url + '\n\nシート:\n・提出サマリー\n・距離差一覧\n・利用実績\n・定義と注意',
      ui.ButtonSet.OK
    );
  } catch (eAlert) {
    Logger.log(url);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('提出用シートを作成しました', '馬事移籍', 8);
  return url;
}

/** @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss */
function bajiReportBuildAll_(ss) {
  bajiReportWriteSummary_(ss);
  bajiReportWriteDistance_(ss);
  bajiReportWriteUsage_(ss);
  bajiReportWriteDefs_(ss);
  var junk = ss.getSheetByName('シート1');
  if (junk && ss.getSheets().length > 1) ss.deleteSheet(junk);
}

/** @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss */
function bajiReportWriteSummary_(ss) {
  var sh = ss.getSheetByName('提出サマリー') || ss.insertSheet('提出サマリー', 0);
  sh.clear();
  sh.getRange('A1').setValue('JOYFIT24経堂 → 馬事公苑｜移籍増加の整理（提出用）').setFontWeight('bold').setFontSize(16);
  sh.getRange('A3').setValue('一言結論').setFontWeight('bold');
  sh.getRange('A4').setValue(
    '「通常会費の差が主因」とは断定できない。対象者の多くはすでに馬事公苑を主利用しており、相互利用による所属変更の寄与が大きい。料金差は事実だが、本人意図までは未確認。'
  );
  sh.getRange('A4:F5').merge().setWrap(true);
  sh.setRowHeight(4, 48);

  sh.getRange('A7').setValue('キー数字').setFontWeight('bold');
  var keys = [
    ['移籍関連対象', '14名'],
    ['住所・距離分析', '13名'],
    ['馬事利用率(回数)', '78.8%'],
    ['馬事回数＞経堂', '12名'],
    ['馬事が近い', '5名'],
    ['経堂が近い', '8名'],
    ['経堂近いが馬事多用', '7名'],
    ['月会費差', '971円']
  ];
  sh.getRange(8, 1, 1, keys.length).setValues([keys.map(function (x) { return x[0]; })]).setBackground('#F3F4F6');
  sh.getRange(9, 1, 1, keys.length).setValues([keys.map(function (x) { return x[1]; })]).setFontWeight('bold').setFontSize(13);

  sh.getRange('A11').setValue('店舗・料金（公式）').setFontWeight('bold');
  sh.getRange('A12:C12').setValues([['項目', '経堂', '馬事公苑']]).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange('A13:C17').setValues([
    ['所在地', '世田谷区経堂5-23-13', '世田谷区上用賀2-4-18 コリーヌ馬事公苑A棟3階'],
    ['緯度', BAJI_REPORT_STORE_LAT_, BAJI_REPORT_BAJI_LAT_],
    ['経度', BAJI_REPORT_STORE_LNG_, BAJI_REPORT_BAJI_LNG_],
    ['通常月会費(税込)', '9,350円', '8,379円'],
    ['参照', 'joyfit.jp/kyodo', 'joyfit.jp/bajikouen']
  ]);

  sh.getRange('A19').setValue('仮説の整理').setFontWeight('bold');
  sh.getRange('A20:C20').setValues([['仮説', '評価', 'ひとこと']]).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange('A21:C25').setValues([
    ['自動移籍（主利用の結果）', '説明力 高', '14名中12名が馬事回数＞経堂'],
    ['立地', '部分的', '馬事近い5名は整合。経堂近いのに馬事多用が7名'],
    ['通常料金差', '補助要因の可能性', '971円差は事実。主因の証明材料は不足'],
    ['キャンペーン終了', '一部のみ', '6805は要確認。6320は終了から約1年'],
    ['判定異常', '個別確認', '6620は別扱い（利用時間異常の可能性）']
  ]);

  sh.getRange('A27').setValue('提出時の言い方').setFontWeight('bold');
  sh.getRange('A28').setValue(
    '×「料金差は関係ない」← 距離だけでは言えない\n○「主因を料金差と断定できない。主利用実態（相互利用）の寄与が大きい」'
  );
  sh.getRange('A28:F29').merge().setWrap(true);
  sh.setRowHeight(28, 48);

  sh.getRange('A31').setValue('次に確認すること').setFontWeight('bold');
  sh.getRange('A32').setValue(
    '1) 11/1付の所属変更確定を社内データ照合\n2) 6620の入退館・時間異常を個別確認\n3) 自動移籍の正式判定期間を規定確認\n4) 可能なら通常料金会員全体の移籍率比較'
  );
  sh.getRange('A32:F34').merge().setWrap(true);

  sh.setColumnWidth(1, 28 * 7);
  sh.setColumnWidth(2, 36 * 7);
  sh.setColumnWidth(3, 42 * 7);
}

/** @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss */
function bajiReportWriteDistance_(ss) {
  var sh = ss.getSheetByName('距離差一覧') || ss.insertSheet('距離差一覧');
  sh.clear();
  sh.getRange('A1').setValue('距離差一覧（プラス／マイナス）').setFontWeight('bold').setFontSize(16);
  sh.getRange('A2').setValue(
    '距離差 = 馬事公苑までの徒歩距離 − 経堂までの徒歩距離　｜　＋＝経堂が近い　／　−＝馬事公苑が近い'
  ).setFontColor('#4B5563');

  var headers = ['ID', '居住エリア', '経堂まで(m)', '馬事まで(m)', '距離差(m)', '近い店', '馬事利用率(%)', '区分'];
  sh.getRange(4, 1, 1, headers.length).setValues([headers]).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');

  var rows = bajiReportDistanceRows_();
  for (var i = 0; i < rows.length; i++) {
    var r = 5 + i;
    var row = rows[i];
    sh.getRange(r, 1, 1, 4).setValues([[row.id, row.area, row.k, row.b]]);
    sh.getRange(r, 5).setFormula('=D' + r + '-C' + r);
    sh.getRange(r, 6).setFormula('=IF(E' + r + '>100,"経堂",IF(E' + r + '<-100,"馬事公苑","ほぼ同距離"))');
    sh.getRange(r, 7).setValue(row.util);
    sh.getRange(r, 8).setFormula(
      '=IF(A' + r + '="6620","例外",IF(AND(E' + r + '>100,G' + r + '>=50),"経堂近いが馬事多用",IF(E' + r + '<-100,"立地整合","—")))'
    );
    var diff = row.b - row.k;
    var bg = row.id === '6620' ? '#FEE2E2' : (diff < -100 ? '#D1FAE5' : '#FEF3C7');
    sh.getRange(r, 1, 1, 8).setBackground(bg);
  }

  var last = 4 + rows.length;
  sh.getRange(last + 2, 1).setValue('店舗座標（指定値）').setFontWeight('bold');
  sh.getRange(last + 3, 1, 1, 3).setValues([['店舗', '緯度', '経度']]).setBackground('#1F2937').setFontColor('#FFFFFF');
  sh.getRange(last + 4, 1, 2, 3).setValues([
    ['経堂', BAJI_REPORT_STORE_LAT_, BAJI_REPORT_STORE_LNG_],
    ['馬事公苑', BAJI_REPORT_BAJI_LAT_, BAJI_REPORT_BAJI_LNG_]
  ]);

  sh.getRange(last + 7, 1).setValue('測定の注意').setFontWeight('bold');
  sh.getRange(last + 8, 1).setValue(
    '・徒歩距離は「丁目代表地点→店舗」のOSM徒歩ルート（自宅ピンからの実測ではない）\n' +
    '・同じ丁目の会員は同一距離になる\n' +
    '・詳細な個人住所は外部地図サービスへ未送信\n' +
    '・色：緑＝馬事が近い／黄＝経堂が近い／赤＝6620例外'
  );
  sh.getRange(last + 8, 1, 1, 8).merge().setWrap(true);
  sh.setRowHeight(last + 8, 72);

  sh.setColumnWidth(1, 70);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(6, 90);
  sh.setColumnWidth(7, 110);
  sh.setColumnWidth(8, 160);
}

/** @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss */
function bajiReportWriteUsage_(ss) {
  var sh = ss.getSheetByName('利用実績') || ss.insertSheet('利用実績');
  sh.clear();
  sh.getRange('A1').setValue('利用実績（2026年4〜9月・回数）').setFontWeight('bold').setFontSize(16);
  var headers = ['ID', '馬事公苑', '経堂', 'その他', '総回数', '馬事利用率', '備考'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');

  var rows = bajiReportUsageRows_();
  for (var i = 0; i < rows.length; i++) {
    var r = 4 + i;
    var row = rows[i];
    sh.getRange(r, 1, 1, 4).setValues([[row.id, row.baji, row.kyodo, row.other]]);
    sh.getRange(r, 5).setFormula('=B' + r + '+C' + r + '+D' + r);
    sh.getRange(r, 6).setFormula('=IF(E' + r + '=0,"",B' + r + '/E' + r + ')').setNumberFormat('0.0%');
    sh.getRange(r, 7).setValue(row.note);
    if (row.id === '6620' || row.id === '4090') {
      sh.getRange(r, 1, 1, 7).setBackground('#FEE2E2');
    }
  }
  var totalRow = 4 + rows.length;
  sh.getRange(totalRow, 1).setValue('合計').setFontWeight('bold');
  sh.getRange(totalRow, 2).setFormula('=SUM(B4:B' + (totalRow - 1) + ')');
  sh.getRange(totalRow, 3).setFormula('=SUM(C4:C' + (totalRow - 1) + ')');
  sh.getRange(totalRow, 4).setFormula('=SUM(D4:D' + (totalRow - 1) + ')');
  sh.getRange(totalRow, 5).setFormula('=SUM(E4:E' + (totalRow - 1) + ')');
  sh.getRange(totalRow, 6).setFormula('=IF(E' + totalRow + '=0,"",B' + totalRow + '/E' + totalRow + ')').setNumberFormat('0.0%');
  sh.getRange(totalRow, 1, 1, 6).setFontWeight('bold');

  sh.getRange(totalRow + 2, 1).setValue('注：利用率は回数割合。利用時間割合ではない。4090は住所分析対象外。').setFontColor('#6B7280');
}

/** @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss */
function bajiReportWriteDefs_(ss) {
  var sh = ss.getSheetByName('定義と注意') || ss.insertSheet('定義と注意');
  sh.clear();
  sh.getRange('A1').setValue('定義・測定方法・注意').setFontWeight('bold').setFontSize(16);
  sh.getRange('A3:B3').setValues([['項目', '内容']]).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange('A4:B13').setValues([
    ['距離差', '馬事公苑までの徒歩距離 − 経堂までの徒歩距離'],
    ['プラス', '経堂店の方が近い'],
    ['マイナス', '馬事公苑店の方が近い'],
    ['近い店の判定', '差が＋100m超→経堂、−100m超→馬事、それ以外→ほぼ同距離'],
    ['測定方法', '丁目代表地点からのOSM徒歩ルート（routing.openstreetmap.de foot）'],
    ['代表地点', '国土地理院住所検索の丁目代表値（重心保証なし）'],
    ['個人情報', '詳細な自宅住所は外部地図へ未送信。丁目単位のみ使用'],
    ['料金断定禁止', '移籍後請求8,379円だけで移籍理由・本人意図を断定しない'],
    ['抽出バイアス', '移籍した可能性のある会員のみのため、全体移籍率は出せない'],
    ['判定期間', '4〜9月集計と正式な自動移籍判定期間は一致しない可能性']
  ]);
  sh.setColumnWidth(1, 140);
  sh.setColumnWidth(2, 520);
  sh.getRange('A4:B13').setWrap(true);
}

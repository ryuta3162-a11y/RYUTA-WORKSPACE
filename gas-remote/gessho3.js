/**
 * Workspace メニュー「月初３ファイル」
 * 会員数・オプション・年齢男女のExcelをまとめて読み、受付状況表の日報月初へ書く。
 * 日報の月初は次月月初を使わない。
 * 会員は当月末在籍から、法人都度・OGF・ゴールドと当月開始を引く。
 * オプションは当月末契約数から当月開始数を引く。契約・解約の列は触らない。
 */

var GESSHO3_RECEPTION_ID_ = '14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w';
var GESSHO3_START_ROW_ = 21;
var GESSHO3_LABEL_COL_ = 2;
var GESSHO3_OPENING_COL_ = 3;
var GESSHO3_OP_LIST_ = [
  '安心サポート', '安心サポートVIP', '水素水', 'オンラインレッスン',
  '体組成計', '契約ロッカー1,500', 'レンタルマット', 'プロテイン12杯',
  'プロテイン無制限', 'プロテイン＋水素水', 'レンタルタオル', 'タンニング',
  'セルフエステ', 'ホットスタジオ', 'ヨガロッカー', 'ピラティスリフォーマー'
];
var GESSHO3_EXCLUDE_ = ['法人会員(都度利用)', 'OGF会員', 'ゴールド会員'];
var GESSHO3_OP_SOURCES_ = {
  '安心サポート': ['JOYFITあんしんサポート', 'あんしんサポート', '安心サポート'],
  '安心サポートVIP': ['VIPあんしんサポート', 'VIP安心サポート'],
  '水素水': ['水素水'],
  'オンラインレッスン': ['オンラインレッスン'],
  '体組成計': ['ボディプランナー会員', 'ボディープランナー会員', '体組成計'],
  '契約ロッカー1,500': ['契約ロッカー1500円', '契約ロッカー1,500'],
  'レンタルマット': ['マットレンタル', 'レンタルマット'],
  'プロテイン12杯': ['プロテイン12杯'],
  'プロテイン無制限': ['プロテイン飲み放題', 'プロテイン無制限'],
  'プロテイン＋水素水': ['水素水+プロテイン(6)', '水素水＋プロテイン(6)'],
  'レンタルタオル': ['レンタルタオル', 'Vitalityレンタルタオル'],
  'タンニング': ['タンニング'],
  'セルフエステ': ['セルフエステ'],
  'ホットスタジオ': ['ホットスタジオ', 'ホットスタジオ(無料)'],
  'ヨガロッカー': ['ヨガマット契約ロッカー', 'ヨガロッカー'],
  'ピラティスリフォーマー': ['グループピラティスリフォーマー', 'ピラティスリフォーマー']
};

function openGessho3Files() {
  var html = HtmlService.createHtmlOutputFromFile('gessho3ui')
    .setWidth(760)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '月初３ファイル');
}

function previewGessho3Files(files) {
  try {
    var parsed = gessho3ParseFiles_(files);
    var nippo = gessho3Nippo_();
    return { ok: true, preview: gessho3BuildPreview_(parsed, nippo) };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function applyGessho3Files(payload) {
  try {
    var preview = payload && payload.preview;
    if (!preview || !preview.member) throw new Error('先にExcelを読み取ってください');
    var optionSum = (preview.options || []).reduce(function (sum, row) { return sum + gessho3Num_(row.next); }, 0);
    if (preview.options && preview.options.length && !optionSum) {
      throw new Error('オプションの月初が読めていません。保存を止めています');
    }
    var nippo = gessho3Nippo_();
    var written = gessho3WriteNippo_(nippo, preview, payload.writeGender === true);
    var saved = gessho3SaveMonth_(preview, written);
    try { written.hq = gessho3WriteHq_(preview); } catch (eHq) { written.hq = { ok: false, message: String(eHq.message || eHq) }; }
    return { ok: true, written: written, monthLabel: saved.monthLabel, sheetName: saved.sheetName };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function gessho3Nippo_() {
  var book = SpreadsheetApp.openById(GESSHO3_RECEPTION_ID_);
  var sh = book.getSheetByName('日報');
  if (!sh) throw new Error('受付状況表に日報がありません');
  return sh;
}

function gessho3ParseFiles_(files) {
  if (!files || !files.length) throw new Error('Excelを選んでください');
  var out = { member: null, option: null, gender: null, files: [] };
  var labels = { member: '会員数', option: 'オプション', gender: '年齢・男女' };
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    if (!file || !file.base64) continue;
    var values = gessho3ReadExcel_(file);
    var kind = gessho3Detect_(values);
    var name = file.name || ('ファイル' + (i + 1));
    if (!kind) throw new Error('中身から種類を判定できないファイルです: ' + name);
    if (out[kind]) {
      throw new Error(labels[kind] + 'のファイルが2つあります: ' + out[kind].fileName + ' と ' + name);
    }
    var parsed = kind === 'gender' ? gessho3ParseGender_(values) : gessho3ParseCounts_(values);
    parsed.fileName = name;
    out[kind] = parsed;
    out.files.push({ name: name, kind: kind, label: labels[kind] });
  }
  if (!out.member) throw new Error('会員数のExcelがありません（在籍者数の表）');
  if (!out.option) throw new Error('オプションのExcelがありません（契約数の表）');
  if (!out.gender) throw new Error('年齢・男女のExcelがありません');
  return out;
}

function gessho3ReadExcel_(file) {
  var raw = Utilities.base64Decode(file.base64);
  var bytes = [];
  for (var i = 0; i < raw.length; i++) bytes.push(raw[i] & 255);
  if (bytes.length < 8) throw new Error('ファイルが空です: ' + (file.name || ''));
  if (bytes[0] === 0xD0 && bytes[1] === 0xCF) return gessho3ReadXls_(bytes);
  if (bytes[0] === 0x50 && bytes[1] === 0x4B) return gessho3ReadXlsx_(raw);
  throw new Error('Excel（xls または xlsx）ではありません: ' + (file.name || ''));
}

function gessho3U16_(bytes, i) {
  return bytes[i] | (bytes[i + 1] << 8);
}

function gessho3U32_(bytes, i) {
  return bytes[i] + (bytes[i + 1] << 8) + (bytes[i + 2] << 16) + (bytes[i + 3] * 16777216);
}

function gessho3ReadXls_(bytes) {
  var workbook = gessho3OleStream_(bytes, 'Workbook') || gessho3OleStream_(bytes, 'Book');
  if (!workbook) throw new Error('xlsの中身を読めませんでした');
  var records = [];
  var pos = 0;
  while (pos + 4 <= workbook.length) {
    var op = gessho3U16_(workbook, pos);
    var len = gessho3U16_(workbook, pos + 2);
    var data = workbook.slice(pos + 4, pos + 4 + len);
    records.push({ op: op, data: data });
    pos += 4 + len;
  }
  var sst = [];
  var sheetStart = -1;
  for (var r = 0; r < records.length; r++) {
    if (records[r].op === 0x00FC) sst = gessho3Sst_(records, r);
    if (records[r].op === 0x0809 && records[r].data.length >= 4 && gessho3U16_(records[r].data, 2) === 0x0010) {
      sheetStart = r;
      break;
    }
  }
  if (sheetStart < 0) throw new Error('xlsの表が見つかりません');
  var cells = {};
  var maxR = 0;
  var maxC = 0;
  for (var i = sheetStart + 1; i < records.length; i++) {
    var rec = records[i];
    if (rec.op === 0x000A) break;
    if (rec.op === 0x00FD && rec.data.length >= 10) {
      var row = gessho3U16_(rec.data, 0);
      var col = gessho3U16_(rec.data, 2);
      var idx = gessho3U32_(rec.data, 6);
      cells[row + ':' + col] = sst[idx] || '';
      if (row > maxR) maxR = row;
      if (col > maxC) maxC = col;
    } else if (rec.op === 0x0203 && rec.data.length >= 14) {
      var nrow = gessho3U16_(rec.data, 0);
      var ncol = gessho3U16_(rec.data, 2);
      cells[nrow + ':' + ncol] = gessho3F64_(rec.data, 6);
      if (nrow > maxR) maxR = nrow;
      if (ncol > maxC) maxC = ncol;
    }
  }
  var grid = [];
  for (var rr = 0; rr <= maxR; rr++) {
    var line = [];
    for (var cc = 0; cc <= maxC; cc++) {
      var key = rr + ':' + cc;
      line.push(Object.prototype.hasOwnProperty.call(cells, key) ? cells[key] : '');
    }
    grid.push(line);
  }
  return grid;
}

function gessho3Sst_(records, start) {
  var parts = [records[start].data];
  for (var j = start + 1; j < records.length && records[j].op === 0x003C; j++) parts.push(records[j].data);
  var ri = 0;
  var p = 0;
  function pull(n) {
    var out = [];
    while (out.length < n) {
      if (ri >= parts.length) throw new Error('xlsの文字一覧が途中で切れています');
      if (p >= parts[ri].length) { ri++; p = 0; continue; }
      var take = Math.min(parts[ri].length - p, n - out.length);
      for (var k = 0; k < take; k++) out.push(parts[ri][p + k]);
      p += take;
    }
    return out;
  }
  function pullText(nchars, wide) {
    var text = '';
    var left = nchars;
    var started = false;
    while (left > 0) {
      if (ri >= parts.length) throw new Error('xlsの文字一覧が途中で切れています');
      if (p >= parts[ri].length) {
        ri++;
        p = 0;
        if (ri >= parts.length) throw new Error('xlsの文字一覧が途中で切れています');
        if (started) {
          wide = (parts[ri][p] & 1) !== 0;
          p++;
        }
        continue;
      }
      var bytesPer = wide ? 2 : 1;
      var avail = Math.floor((parts[ri].length - p) / bytesPer);
      if (!avail) { p = parts[ri].length; continue; }
      var takeChars = Math.min(left, avail);
      for (var c = 0; c < takeChars; c++) {
        var code = wide ? (parts[ri][p] | (parts[ri][p + 1] << 8)) : parts[ri][p];
        text += String.fromCharCode(code);
        p += bytesPer;
      }
      left -= takeChars;
      started = true;
    }
    return text;
  }
  var head = pull(8);
  var uniq = head[4] + (head[5] << 8) + (head[6] << 16) + (head[7] * 16777216);
  var sst = [];
  for (var n = 0; n < uniq; n++) {
    var ch = pull(3);
    var cch = ch[0] | (ch[1] << 8);
    var flags = ch[2];
    var rich = 0;
    var phonetic = 0;
    if (flags & 8) {
      var rb = pull(2);
      rich = rb[0] | (rb[1] << 8);
    }
    if (flags & 4) {
      var pb = pull(4);
      phonetic = pb[0] + (pb[1] << 8) + (pb[2] << 16) + (pb[3] * 16777216);
    }
    sst.push(pullText(cch, (flags & 1) !== 0));
    if (rich) pull(rich * 4);
    if (phonetic) pull(phonetic);
  }
  return sst;
}

function gessho3F64_(bytes, i) {
  var buf = new ArrayBuffer(8);
  var view = new Uint8Array(buf);
  for (var k = 0; k < 8; k++) view[k] = bytes[i + k];
  return new DataView(buf).getFloat64(0, true);
}

function gessho3OleStream_(bytes, streamName) {
  if (bytes.length < 512 || bytes[0] !== 0xD0) return null;
  var sec = 1 << gessho3U16_(bytes, 0x1E);
  var mini = 1 << gessho3U16_(bytes, 0x20);
  var dirStart = gessho3U32_(bytes, 0x30);
  var cutoff = gessho3U32_(bytes, 0x38);
  var minifatStart = gessho3U32_(bytes, 0x3C);
  var numMinifat = gessho3U32_(bytes, 0x40);
  var difatStart = gessho3U32_(bytes, 0x44);
  var numDifat = gessho3U32_(bytes, 0x48);
  var difat = [];
  for (var d = 0; d < 109; d++) difat.push(gessho3U32_(bytes, 0x4C + d * 4));
  var ds = difatStart;
  for (var n = 0; n < numDifat && ds < 0xFFFFFFFA; n++) {
    var base = (ds + 1) * sec;
    var count = Math.floor(sec / 4);
    for (var k = 0; k < count - 1; k++) difat.push(gessho3U32_(bytes, base + k * 4));
    ds = gessho3U32_(bytes, base + (count - 1) * 4);
  }
  var fat = [];
  for (var f = 0; f < difat.length; f++) {
    if (difat[f] >= 0xFFFFFFFA) continue;
    var fatBase = (difat[f] + 1) * sec;
    for (var fk = 0; fk < sec / 4; fk++) fat.push(gessho3U32_(bytes, fatBase + fk * 4));
  }
  function chain(start) {
    var out = [];
    var s = start;
    var guard = 0;
    while (s < 0xFFFFFFFA && guard < 100000) {
      out.push(s);
      s = fat[s];
      guard++;
    }
    return out;
  }
  function readChain(start) {
    var out = [];
    var sectors = chain(start);
    for (var i = 0; i < sectors.length; i++) {
      var off = (sectors[i] + 1) * sec;
      for (var b = 0; b < sec && off + b < bytes.length; b++) out.push(bytes[off + b]);
    }
    return out;
  }
  var dirBlob = readChain(dirStart);
  var entries = [];
  for (var e = 0; e + 128 <= dirBlob.length; e += 128) {
    var nameLen = gessho3U16_(dirBlob, e + 64);
    var chars = [];
    var nchars = nameLen > 2 ? (nameLen / 2) - 1 : 0;
    for (var nc = 0; nc < nchars; nc++) chars.push(String.fromCharCode(gessho3U16_(dirBlob, e + nc * 2)));
    entries.push({
      name: chars.join(''),
      type: dirBlob[e + 66],
      start: gessho3U32_(dirBlob, e + 116),
      size: gessho3U32_(dirBlob, e + 120)
    });
  }
  if (!entries.length) return null;
  var minifat = [];
  if (numMinifat && minifatStart < 0xFFFFFFFA) {
    var miniSectors = chain(minifatStart);
    for (var ms = 0; ms < miniSectors.length; ms++) {
      var moff = (miniSectors[ms] + 1) * sec;
      for (var mk = 0; mk < sec / 4; mk++) minifat.push(gessho3U32_(bytes, moff + mk * 4));
    }
  }
  var root = entries[0];
  var ministream = root.start < 0xFFFFFFFA ? readChain(root.start).slice(0, root.size) : [];
  function readMini(start, size) {
    var out = [];
    var s = start;
    var guard = 0;
    while (s < 0xFFFFFFFA && out.length < size && guard < 100000) {
      var off = s * mini;
      for (var b = 0; b < mini && out.length < size && off + b < ministream.length; b++) out.push(ministream[off + b]);
      s = minifat[s];
      guard++;
    }
    return out;
  }
  for (var ei = 0; ei < entries.length; ei++) {
    if (entries[ei].type === 2 && entries[ei].name === streamName) {
      if (entries[ei].size >= cutoff) return readChain(entries[ei].start).slice(0, entries[ei].size);
      return readMini(entries[ei].start, entries[ei].size);
    }
  }
  return null;
}

function gessho3ReadXlsx_(raw) {
  var blobs = Utilities.unzip(Utilities.newBlob(raw, 'application/zip', 'gessho3.xlsx'));
  var sharedXml = '';
  var sheetXml = '';
  for (var i = 0; i < blobs.length; i++) {
    var name = String(blobs[i].getName() || '').replace(/\\/g, '/');
    if (name.slice(-18) === 'sharedStrings.xml') sharedXml = blobs[i].getDataAsString('UTF-8');
    if (/worksheets\/sheet1\.xml$/.test(name)) sheetXml = blobs[i].getDataAsString('UTF-8');
  }
  if (!sheetXml) throw new Error('xlsxの表が見つかりません');
  var nsMain = XmlService.getNamespace('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  var shared = [];
  if (sharedXml) {
    var sis = XmlService.parse(sharedXml).getRootElement().getChildren('si', nsMain);
    for (var s = 0; s < sis.length; s++) {
      var direct = sis[s].getChild('t', nsMain);
      if (direct) {
        shared.push(direct.getText() || '');
        continue;
      }
      var runs = sis[s].getChildren('r', nsMain);
      var text = '';
      for (var run = 0; run < runs.length; run++) {
        var rt = runs[run].getChild('t', nsMain);
        if (rt) text += rt.getText() || '';
      }
      shared.push(text);
    }
  }
  var sheet = XmlService.parse(sheetXml).getRootElement();
  var sheetData = sheet.getChild('sheetData', nsMain);
  var rows = sheetData ? sheetData.getChildren('row', nsMain) : [];
  var grid = [];
  var maxC = 0;
  for (var ri = 0; ri < rows.length; ri++) {
    var rowAttr = rows[ri].getAttribute('r');
    var rowIndex = rowAttr ? Number(rowAttr.getValue()) - 1 : grid.length;
    while (grid.length <= rowIndex) grid.push([]);
    var cells = rows[ri].getChildren('c', nsMain);
    for (var ci = 0; ci < cells.length; ci++) {
      var ref = cells[ci].getAttribute('r');
      var col = ref ? gessho3ColIndex_(ref.getValue()) : grid[rowIndex].length;
      var typeAttr = cells[ci].getAttribute('t');
      var type = typeAttr ? typeAttr.getValue() : '';
      var valueNode = cells[ci].getChild('v', nsMain);
      var value = '';
      if (type === 's' && valueNode) value = shared[Number(valueNode.getText())] || '';
      else if (type === 'inlineStr') {
        var inline = cells[ci].getChild('is', nsMain);
        var inlineText = inline ? inline.getChild('t', nsMain) : null;
        value = inlineText ? (inlineText.getText() || '') : '';
      } else if (valueNode) {
        var num = Number(valueNode.getText());
        value = isNaN(num) ? valueNode.getText() : num;
      }
      while (grid[rowIndex].length <= col) grid[rowIndex].push('');
      grid[rowIndex][col] = value;
      if (col > maxC) maxC = col;
    }
  }
  for (var r = 0; r < grid.length; r++) {
    while (grid[r].length <= maxC) grid[r].push('');
  }
  return grid;
}

function gessho3ColIndex_(ref) {
  var letters = String(ref || '').replace(/[0-9]/g, '');
  var n = 0;
  for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

function gessho3Detect_(values) {
  var text = '';
  for (var r = 0; r < Math.min(3, values.length); r++) text += values[r].join('\t') + '\n';
  var flat = text.replace(/\s/g, '');
  if (flat.indexOf('男性') !== -1 && (flat.indexOf('年齢') !== -1 || flat.indexOf('00-03') !== -1 || flat.indexOf('00－03') !== -1)) {
    return 'gender';
  }
  if (flat.indexOf('契約数') !== -1) return 'option';
  if (flat.indexOf('在籍者数') !== -1) return 'member';
  return '';
}

function gessho3ParseCounts_(values) {
  var headerRow = -1;
  for (var r = 0; r < Math.min(5, values.length); r++) {
    if (gessho3RowText_(values[r]).indexOf('契約名称') !== -1) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) throw new Error('契約名称の見出しがありません');
  var header = values[headerRow];
  var sub = values[headerRow + 1] || [];
  var colName = gessho3FindCol_(header, '契約名称');
  var colPrev = gessho3FindCol_(header, '前月末');
  var colEnd = gessho3FindCol_(header, '当月末');
  var colStart = gessho3StartCol_(header, sub);
  var colNext = gessho3OpeningCol_(header, sub);
  var colPause = gessho3FindCol_(sub, '休会');
  if (colEnd < 0 || colStart < 0 || colName < 0) throw new Error('当月末または当月開始の列がありません');
  var total = null;
  var contracts = [];
  for (var i = headerRow + 2; i < values.length; i++) {
    var row = values[i];
    var lead = gessho3Norm_(row[0]);
    var name = gessho3Norm_(row[colName]);
    if (lead === '合計') {
      total = {
        prev: gessho3Num_(row[colPrev]),
        start: gessho3Num_(row[colStart]),
        end: gessho3Num_(row[colEnd]),
        next: colNext < 0 ? 0 : gessho3Num_(row[colNext]),
        pause: colPause < 0 ? 0 : gessho3Num_(row[colPause])
      };
      continue;
    }
    if (!name || name === '会員種類計' || name === '会員区分計') continue;
    contracts.push({
      name: name,
      prev: gessho3Num_(row[colPrev]),
      start: gessho3Num_(row[colStart]),
      end: gessho3Num_(row[colEnd]),
      next: colNext < 0 ? 0 : gessho3Num_(row[colNext])
    });
  }
  if (!total) throw new Error('合計行がありません');
  return { total: total, contracts: contracts };
}

function gessho3ParseGender_(values) {
  for (var r = 0; r < values.length; r++) {
    var lead = gessho3Norm_(values[r][0]);
    var sub = gessho3Norm_(values[r][1]);
    if (lead === '合計' && sub.indexOf('在籍') !== -1) {
      var male = gessho3Num_(values[r][2]);
      var female = gessho3Num_(values[r][3]);
      return { male: male, female: female, total: male + female };
    }
  }
  throw new Error('年齢表の合計（在籍者数）がありません');
}

function gessho3BuildPreview_(parsed, nippo) {
  var labels = nippo.getRange(GESSHO3_START_ROW_, GESSHO3_LABEL_COL_, GESSHO3_OP_LIST_.length, 1).getDisplayValues();
  var currentOpening = nippo.getRange(GESSHO3_START_ROW_, GESSHO3_OPENING_COL_, GESSHO3_OP_LIST_.length, 1).getValues();
  var mapped = gessho3MapOptions_(parsed.option.contracts);
  var opening = gessho3MemberOpening_(parsed.member);
  var options = [];
  for (var i = 0; i < GESSHO3_OP_LIST_.length; i++) {
    var label = String(labels[i][0] || '') || GESSHO3_OP_LIST_[i];
    var key = gessho3MatchKey_(label);
    var found = mapped.byName[key];
    options.push({
      name: label,
      current: gessho3Num_(currentOpening[i][0]),
      next: found ? found.next : 0,
      end: found ? found.end : 0,
      missing: !found
    });
  }
  var b1 = String(nippo.getRange('B1').getDisplayValue() || '');
  var minus = gessho3GenderMinus_(b1);
  var kiyaku = gessho3KiyakuMinus_(gessho3MonthLabel_(b1));
  opening -= kiyaku.total;
  minus.kiyakuMale = kiyaku.male;
  var male = parsed.gender.male - minus.exMale - minus.startMale - kiyaku.male;
  return {
    b1: b1,
    monthLabel: gessho3MonthLabel_(b1),
    files: parsed.files || [],
    member: {
      current: gessho3Num_(nippo.getRange('C12').getValue()),
      next: opening,
      prev: parsed.member.total.prev,
      end: parsed.member.total.end,
      pause: parsed.member.total.pause,
      kiyaku: kiyaku.total,
      kiyakuNames: kiyaku.names
    },
    gender: {
      rawMale: parsed.gender.male,
      rawFemale: parsed.gender.female,
      male: male,
      female: opening - male,
      total: opening,
      minus: minus,
      fitsOpening: male > 0 && male < opening
    },
    options: options,
    unused: mapped.unused
  };
}

/**
 * 年齢表（当月末在籍）から、月初会員と同じ人を男女別に引く。
 * 特例（法人都度・OGF・ゴールド）は gessho3TokureiRows_、当月開始は累計入会データの利用開始年月が当月の人。
 * 女は「月初会員 − 男」にするので、合計は必ず月初会員と一致する。
 */
function gessho3GenderMinus_(b1) {
  var out = { exMale: 0, exFemale: 0, startMale: 0, startFemale: 0 };
  gessho3TokureiRows_().forEach(function (row) {
    if (row[4] === '男') out.exMale++;
    else if (row[4] === '女') out.exFemale++;
  });
  var label = gessho3MonthLabel_(b1);
  var m = label.match(/(\d{4})年(\d{1,2})月/);
  if (!m) return out;
  var ym = m[1] + '/' + ('0' + m[2]).slice(-2);
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q');
  var src = ss.getSheetByName('累計入会データ');
  if (!src || src.getLastRow() < 2) return out;
  var rows = src.getRange(2, 9, src.getLastRow() - 1, 6).getDisplayValues();
  rows.forEach(function (r) {
    if (String(r[5]).replace(/\s/g, '').indexOf(ym) !== 0) return;
    if (r[0] === '男') out.startMale++;
    else if (r[0] === '女') out.startFemale++;
  });
  return out;
}

/** 規約退会リストで退会年月が当月の人（月初の強制退会）。会員数の表にはまだ残っているので月初から外す */
function gessho3KiyakuMinus_(monthLabel) {
  var out = { total: 0, male: 0, names: [] };
  var m = String(monthLabel || '').match(/(\d{4})年(\d{1,2})月/);
  if (!m) return out;
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q');
  var list = ss.getSheetByName('規約退会リスト');
  if (!list || list.getLastRow() < 4) return out;
  var rows = list.getRange(4, 1, list.getLastRow() - 3, 11).getValues();
  var join = ss.getSheetByName('累計入会データ');
  var gender = {};
  if (join && join.getLastRow() > 1) {
    join.getRange(2, 6, join.getLastRow() - 1, 4).getDisplayValues().forEach(function (r) { gender[String(r[0]).trim()] = r[3]; });
  }
  rows.forEach(function (r) {
    var ym = r[10];
    if (!(ym instanceof Date) || ym.getFullYear() !== Number(m[1]) || ym.getMonth() + 1 !== Number(m[2])) return;
    out.total++;
    out.names.push(r[3]);
    if (gender[String(r[2]).trim()] === '男') out.male++;
  });
  return out;
}

function gessho3MapOptions_(contracts) {
  var byName = {};
  var used = {};
  var keys = Object.keys(GESSHO3_OP_SOURCES_);
  for (var i = 0; i < contracts.length; i++) {
    var contract = contracts[i];
    var norm = gessho3Loose_(contract.name);
    var hit = '';
    for (var k = 0; k < keys.length && !hit; k++) {
      var aliases = GESSHO3_OP_SOURCES_[keys[k]];
      for (var a = 0; a < aliases.length; a++) {
        if (gessho3Loose_(aliases[a]) === norm) {
          hit = keys[k];
          break;
        }
      }
    }
    if (!hit) continue;
    used[i] = true;
    if (!byName[hit]) byName[hit] = { next: 0, end: 0 };
    byName[hit].next += contract.end - contract.start;
    byName[hit].end += contract.end;
  }
  var unused = [];
  for (var u = 0; u < contracts.length; u++) {
    if (!used[u] && (contracts[u].next || contracts[u].end || contracts[u].prev)) unused.push(contracts[u].name);
  }
  return { byName: byName, unused: unused };
}

function gessho3MatchKey_(label) {
  var norm = gessho3Loose_(label);
  var keys = Object.keys(GESSHO3_OP_SOURCES_);
  keys.sort(function (a, b) { return gessho3Loose_(b).length - gessho3Loose_(a).length; });
  for (var i = 0; i < keys.length; i++) {
    if (gessho3Loose_(keys[i]) === norm) return keys[i];
  }
  return label;
}

function gessho3WriteNippo_(nippo, preview, writeGender) {
  var memberCell = nippo.getRange('C12');
  memberCell.setValue(preview.member.next);
  memberCell.setNumberFormat('0');
  memberCell.setNote(
    '月初会員 = 当月末在籍 − 法人都度 − OGF − ゴールド − 当月開始 − 規約退会' +
      (preview.member.kiyaku ? ' ' + preview.member.kiyaku + '（' + (preview.member.kiyakuNames || []).join('・') + '）' : ' 0') +
      '。前月末 ' + preview.member.prev +
      ' / 当月末 ' + preview.member.end +
      ' / 休会 ' + preview.member.pause
  );
  var values = preview.options.map(function (row) { return [row.next]; });
  nippo.getRange(GESSHO3_START_ROW_, GESSHO3_OPENING_COL_, values.length, 1).setValues(values);
  var genderWritten = false;
  if (preview.gender.fitsOpening) {
    var g = preview.gender;
    nippo.getRange('F12').setValue(g.male).setNumberFormat('0').setNote(
      '年齢表の男 ' + g.rawMale + ' − 特例 ' + g.minus.exMale + ' − 当月開始 ' + g.minus.startMale + ' = ' + g.male
    );
    nippo.getRange('H12').setFormula('=C12-F12').setNumberFormat('0');
    genderWritten = true;
  }
  return { member: preview.member.next, genderWritten: genderWritten, options: preview.options.length };
}

/**
 * 会員動向（本部シート「経堂」31〜46行）の当月の列へ、オプション区分別実績表の当月末契約数を入れる。
 * すでに数字が入っているセルは上書きしない。
 */
var GESSHO3_HQ_ID_ = '1LOOUG97wuiKbhzl0BjJstXgLaaSCZAKNFdD8P3I5x_o';
var GESSHO3_HQ_ROWS_ = [
  [31, '安心サポート'], [32, '安心サポートVIP'], [33, 'オンラインレッスン'], [34, 'セルフエステ'],
  [35, 'タンニング'], [36, 'プロテイン12杯'], [37, 'プロテイン無制限'], [38, 'ホットスタジオ'],
  [39, '体組成計'], [40, 'レンタルマット'], [41, 'ヨガロッカー'], [42, 'レンタルタオル'],
  [43, '契約ロッカー1,500'], [44, '水素水'], [45, 'プロテイン＋水素水'], [46, 'ピラティスリフォーマー']
];

function gessho3WriteHq_(preview) {
  var m = String(preview.monthLabel || '').match(/(\d{4})年(\d{1,2})月/);
  if (!m) return { ok: false, message: '月が分かりません' };
  var month = Number(m[2]);
  var col = ((month + 8) % 12) + 3;
  var byKey = {};
  (preview.options || []).forEach(function (row) { byKey[gessho3MatchKey_(row.name)] = row; });
  var sh = SpreadsheetApp.openById(GESSHO3_HQ_ID_).getSheetByName('経堂');
  if (!sh) return { ok: false, message: '会員動向に経堂シートがありません' };
  if (Number(sh.getRange(1, col).getValue()) !== month) return { ok: false, message: '会員動向の月の列が見つかりません' };
  var cur = sh.getRange(31, col, 16, 1).getValues();
  var wrote = 0;
  GESSHO3_HQ_ROWS_.forEach(function (pair, i) {
    var row = byKey[pair[1]];
    if (!row || !row.end || cur[i][0] !== '') return;
    sh.getRange(pair[0], col).setValue(row.end);
    wrote++;
  });
  return { ok: true, column: col, wrote: wrote };
}

function gessho3MonthLabel_(b1) {
  var digits = String(b1 || '').replace(/\D/g, '');
  if (digits.length >= 4) {
    var y = digits.length >= 6 ? Number(digits.slice(0, 4)) : 2000 + Number(digits.slice(0, 2));
    var m = Number(digits.slice(-2));
    if (y && m) return y + '年' + m + '月';
  }
  return '月不明';
}

function gessho3SaveMonth_(preview, written) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var monthLabel = preview.monthLabel || gessho3MonthLabel_(preview.b1);
  var sheetName = '月初_' + monthLabel;
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    var after = ss.getSheetByName('日報');
    sh = ss.insertSheet(sheetName, after ? after.getIndex() + 1 : ss.getSheets().length);
  }
  sh.clear();
  var fileLine = (preview.files || []).map(function (f) { return f.label + ' ← ' + f.name; }).join(' / ');
  var rows = [
    [monthLabel, '取込', new Date()],
    ['元ファイル', fileLine, ''],
    ['月初会員', preview.member.next, '前月末 ' + preview.member.prev + ' / 当月末 ' + preview.member.end],
    ['男', preview.gender.male, '年齢表（当月末）'],
    ['女', preview.gender.female, '年齢表（当月末）'],
    ['休会', preview.member.pause, written.genderWritten ? '男女は日報へ書いた' : '男女は月初と合計が違うため日報には未記入'],
    [],
    ['オプション', '月初', '取込前の日報', '当月末（会員動向へ）']
  ];
  preview.options.forEach(function (row) {
    rows.push([row.name, row.next, row.current, row.end]);
  });
  sh.getRange(1, 1, rows.length, 4).setValues(rows.map(function (row) {
    return [0, 1, 2, 3].map(function (i) { return row[i] === undefined || row[i] === null ? '' : row[i]; });
  }));
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  sh.getRange(8, 1, 1, 4).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 360);
  sh.setTabColor('#111111');
  gessho3UpdateIndex_(ss, monthLabel, sheetName, preview);
  return { monthLabel: monthLabel, sheetName: sheetName };
}

function gessho3UpdateIndex_(ss, monthLabel, sheetName, preview) {
  var index = ss.getSheetByName('月初３ファイル');
  if (!index) index = ss.insertSheet('月初３ファイル', ss.getSheetByName(sheetName).getIndex());
  if (index.getLastRow() < 1 || String(index.getRange('A1').getValue() || '') !== '月') {
    index.clear();
    index.getRange(1, 1, 1, 6).setValues([['月', '取込日時', '月初会員', '男', '女', '保存シート']]);
    index.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
    index.setColumnWidth(1, 140);
    index.setColumnWidth(2, 160);
    index.setColumnWidth(6, 180);
  }
  var last = Math.max(index.getLastRow(), 1);
  var months = last > 1 ? index.getRange(2, 1, last - 1, 1).getDisplayValues() : [];
  var row = 0;
  for (var i = 0; i < months.length; i++) {
    if (String(months[i][0]) === monthLabel) row = i + 2;
  }
  if (!row) row = last + 1;
  index.getRange(row, 1, 1, 6).setValues([[
    monthLabel,
    new Date(),
    preview.member.next,
    preview.gender.male,
    preview.gender.female,
    sheetName
  ]]);
  index.getRange(row, 6).setFormula('=HYPERLINK("#gid=' + ss.getSheetByName(sheetName).getSheetId() + '","' + sheetName + '")');
}

function gessho3MemberOpening_(member) {
  var excluded = 0;
  (member.contracts || []).forEach(function (row) {
    if (GESSHO3_EXCLUDE_.indexOf(gessho3Norm_(row.name)) !== -1) excluded += row.end;
  });
  return member.total.end - excluded - member.total.start;
}

function gessho3StartCol_(header, sub) {
  for (var i = 0; i < header.length; i++) {
    var head = gessho3Norm_(header[i]);
    var child = gessho3Norm_(sub[i]);
    if ((head.indexOf('当月利用') !== -1 || head.indexOf('当月契約') !== -1) && child.indexOf('開始') !== -1) {
      return i;
    }
  }
  return -1;
}

function gessho3OpeningCol_(header, sub) {
  var labelCol = gessho3FindCol_(header, '次月月初');
  if (labelCol < 0) return -1;
  function isCount(text) {
    var t = gessho3Norm_(text);
    if (!t || t.indexOf('開始') !== -1 || t.indexOf('解約') !== -1 || t.indexOf('退会') !== -1) return false;
    return t.indexOf('在籍') !== -1 || t.indexOf('契約数') !== -1;
  }
  if (isCount(sub[labelCol])) return labelCol;
  for (var d = 1; d <= 3; d++) {
    if (labelCol - d >= 0 && isCount(sub[labelCol - d])) return labelCol - d;
    if (labelCol + d < sub.length && isCount(sub[labelCol + d])) return labelCol + d;
  }
  return labelCol;
}

function gessho3FindCol_(row, label) {
  var target = gessho3Norm_(label);
  for (var i = 0; i < row.length; i++) {
    if (gessho3Norm_(row[i]).indexOf(target) !== -1) return i;
  }
  return -1;
}

function gessho3RowText_(row) {
  return row.map(function (cell) { return gessho3Norm_(cell); }).join('');
}

function gessho3Norm_(value) {
  return String(value || '').replace(/\s/g, '');
}

function gessho3Loose_(value) {
  return gessho3Norm_(value).replace(/＋/g, '+').replace(/，/g, ',');
}

function gessho3Num_(value) {
  if (typeof value === 'number') return value;
  var text = String(value || '').replace(/,/g, '').replace(/[^\d.\-]/g, '');
  if (!text) return 0;
  var n = Number(text);
  return isNaN(n) ? 0 : n;
}

function gessho3TokureiRows_() {
  return [
    ['ゴールド会員', '1304000899', '石綿　昌吾', 'イシワタ ショウゴ', '男', 54, '1972/05/17', ''],
    ['ゴールド会員', '1304000901', '寺山　大樹', 'テラヤマ タイキ', '男', 49, '1977/05/17', '累計入会の契約名称は空。ゴールドとして外す'],
    ['ゴールド会員', '1304000905', '寺山 浩美', 'テラヤマ ヒロミ', '女', 57, '1969/01/27', ''],
    ['OGF会員', '3520002643', '石田　真央', 'イシダ マオ', '女', 21, '2004/12/08', ''],
    ['OGF会員', '3520003148', '佐藤 奈美', 'サトウ ナミ', '女', 64, '1962/03/09', ''],
    ['OGF会員', '3510002287', '田村　重己', 'タムラ シゲミ', '男', 72, '1953/12/18', ''],
    ['法人会員（都度利用）', '1304001188', '石井　嘉則', 'イシイ ヨシノリ', '男', 59, '1967/02/03', '累計入会は法人会員(月払)。都度として外す'],
    ['法人会員（都度利用）', '1304005464', '泉井 杏理', 'イズイ アンリ', '女', 35, '1991/07/03', ''],
    ['法人会員（都度利用）', '1304004917', '大森　勇輝', 'オオモリ ユウキ', '男', 37, '1989/07/07', ''],
    ['法人会員（都度利用）', '1304006835', 'キム スヨン', 'キム スヨン', '女', 28, '1998/03/07', ''],
    ['法人会員（都度利用）', '3520001757', '小西　未来', 'コニシ ミク', '女', 38, '1988/02/17', ''],
    ['法人会員（都度利用）', '3520003109', 'パク ジヨン', 'パク ジヨン', '女', 40, '1986/07/30', '指定の読みはパク ジョン']
  ];
}

function gessho3OctoberJoins_() {
  return [
    ['1304007396', '田村 哲朗', 'タムラ テツロウ', '男', 53, '1973/02/15', '2026/09/25'],
    ['1304007418', '朴 ソラ', 'パク ソラ', '女', 36, '1989/12/22', '2026/09/30'],
    ['1304007419', '渡辺 聖磨', 'ワタナベ ショウマ', '男', 39, '1987/09/24', '2026/09/30'],
    ['1304007420', '田中 玄', 'タナカ ゲン', '男', 29, '1997/02/15', '2026/10/01'],
    ['1304007421', '原 啓一朗', 'ハラ ケイイチロウ', '男', 27, '1999/07/22', '2026/10/01'],
    ['1304007422', '水谷 直之', 'ミズタニ ナオユキ', '男', 41, '1985/06/29', '2026/10/01'],
    ['3810006231', '崔 カクチン', 'サイ カクチン', '男', 28, '1998/05/30', '2026/10/01']
  ];
}

/**
 * 月初の内訳を1枚にまとめ、当月入会7人を累計入会データへ足す。
 * 特例12人と当月入会7人は、月初の会員数・男女に含めない。
 */
function installTokureiKaiinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.openById('1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q');
  var junk = ss.getSheetByName('シート9');
  if (junk && junk.getLastRow() < 1) ss.deleteSheet(junk);
  var added = gessho3AppendOctoberJoins_(ss);
  var name = '月初の内訳';
  var mark = '2026年10月 v2';
  var sh = ss.getSheetByName(name);
  if (sh && String(sh.getRange('B1').getValue() || '') === mark) {
    return { ok: true, skipped: true, sheet: name, added: added };
  }
  if (!sh) {
    var after = ss.getSheetByName('累計入会データ');
    sh = ss.insertSheet(name, after ? after.getIndex() + 1 : ss.getSheets().length);
  }
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.setTabColor('#111111');
  var blocks = [];
  function add(row) { blocks.push(row); }
  add(['月初の内訳', mark, '', '', '', '', '', '']);
  add(['当月末在籍 1535 − 特例 12 − 当月入会 7 = 月初会員 1516', '', '', '', '', '', '', '']);
  add(['年齢表 男1086・女449 − 特例 男5女7 − 当月入会 男6女1 = 男1075・女441', '', '', '', '', '', '', '']);
  add(['', '', '', '', '', '', '', '']);
  add(['会員数から外す特例', '12人', '男5', '女7', '', '', '', '']);
  add(['種別', '会員番号', '氏名', '読み仮名', '性別', '年齢', '生年月日', 'メモ']);
  var specialHeader = blocks.length;
  gessho3TokureiRows_().forEach(add);
  add(['', '', '', '', '', '', '', '']);
  add(['当月入会（月初からは外す）', '7人', '男6', '女1', '利用開始 2026/10', '', '', '']);
  add(['会員番号', '氏名', '読み仮名', '性別', '年齢', '生年月日', '届出日', 'メモ']);
  var joinHeader = blocks.length;
  gessho3OctoberJoins_().forEach(function (row) {
    add(row.concat(['電話は写真の端で切れているので未入力。田村・崔は累計入会に既にあった']));
  });
  sh.getRange(1, 1, blocks.length, 8).setValues(blocks);
  sh.getRange(1, 1, 3, 8).setFontWeight('bold');
  sh.getRange(5, 1, 1, 8).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  sh.getRange(specialHeader, 1, 1, 8).setFontWeight('bold').setBackground('#333333').setFontColor('#ffffff');
  sh.getRange(joinHeader - 1, 1, 1, 8).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  sh.getRange(joinHeader, 1, 1, 8).setFontWeight('bold').setBackground('#333333').setFontColor('#ffffff');
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(8, 420);
  sh.setFrozenRows(3);
  return { ok: true, sheet: name, added: added };
}

function gessho3AppendOctoberJoins_(ss) {
  var src = ss.getSheetByName('累計入会データ');
  if (!src) return 0;
  var last = Math.max(src.getLastRow(), 1);
  var ids = src.getRange(2, 6, Math.max(last - 1, 1), 1).getDisplayValues();
  var have = {};
  for (var i = 0; i < ids.length; i++) have[String(ids[i][0]).replace(/\s/g, '')] = true;
  var fresh = [];
  gessho3OctoberJoins_().forEach(function (row) {
    if (have[row[0]]) return;
    fresh.push([
      '', '', '', '', '月払',
      row[0], row[1], row[2], row[3], row[4], row[5], '',
      row[6], '2026/10', '新'
    ]);
  });
  if (!fresh.length) return 0;
  src.getRange(last + 1, 1, fresh.length, 15).setValues(fresh);
  return fresh.length;
}

/**
 * 累計入会・累計退会を見やすくする。列は消さない。
 * 会員分析と経堂マスタは、入会のF/G/Nと退会のF/P/Rの位置を見ている。
 * 隠すだけにする。年齢・性別・電話の空欄はそのまま。
 */
function formatCumulativeSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.openById('1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q');
  var join = ss.getSheetByName('累計入会データ');
  var leave = ss.getSheetByName('累計退会データ');
  if (join) {
    formatCumulativeSheet_(join, [1, 2, 4], {
      3: 200, 5: 72, 6: 120, 7: 150, 8: 160, 9: 48, 10: 48,
      11: 110, 12: 130, 13: 110, 14: 88, 15: 56
    });
  }
  if (leave) {
    formatCumulativeSheet_(leave, [1, 2, 4, 13, 14, 15], {
      3: 200, 5: 72, 6: 120, 7: 150, 8: 160, 9: 48, 10: 48,
      11: 110, 12: 130, 16: 88, 17: 110, 18: 88, 19: 56
    });
  }
  PropertiesService.getDocumentProperties().setProperty('CUMULATIVE_LOOK_V', 'v1');
  return { ok: true, join: !!join, leave: !!leave };
}

function formatCumulativeSheet_(sh, hideCols, widths) {
  var last = Math.max(sh.getLastRow(), 2);
  var cols = Math.max(sh.getLastColumn(), 1);
  sh.setTabColor('#111111');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 40);
  sh.showColumns(1, cols);
  var bandings = sh.getBandings();
  for (var i = 0; i < bandings.length; i++) bandings[i].remove();
  var header = sh.getRange(1, 1, 1, cols);
  header
    .setBackground('#111111')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontFamily('Meiryo')
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  if (last > 1) {
    sh.getRange(2, 1, last - 1, cols)
      .setFontFamily('Meiryo')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sh.getRange(1, 1, last, cols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  }
  header.setBackground('#111111').setFontColor('#ffffff').setFontWeight('bold');
  for (var h = 0; h < hideCols.length; h++) {
    if (hideCols[h] <= cols) sh.hideColumns(hideCols[h]);
  }
  Object.keys(widths).forEach(function (col) {
    var n = Number(col);
    if (n <= cols) sh.setColumnWidth(n, widths[col]);
  });
  var filter = sh.getFilter();
  if (filter) filter.remove();
  sh.getRange(1, 1, last, cols).createFilter();
  sh.getRange(1, 1).setNote(
    '分析に使わない列は隠しています。列ごと削除すると、右側の数字の位置がずれます。年齢・性別・電話が空の行は、あとから入力してください。'
  );
}

/**
 * 2026年10月の取り込みは次月月初を書いていた。日報の月初だけを基準の数へ直し、保存シートは隠す。
 * 契約・解約・増減は触らない。
 */
function hideGesshoSideSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { ok: false };
  ['月初_2026年10月', '月初３ファイル', '月初の内訳', '経堂_入会'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    try { sh.hideSheet(); } catch (eHide) {}
  });
  return { ok: true };
}

/** 日報 月初の女（H12）は 月初会員 − 男 の式で持つ（合計が必ず合う） */
function ensureNippoOpeningFemale_() {
  var nippo = gessho3Nippo_();
  var h12 = nippo.getRange('H12');
  if (h12.getFormula() === '=C12-F12') return;
  h12.setFormula('=C12-F12').setNumberFormat('0');
}

/**
 * 日報 月末安定・翌月月初の男女。メールに載らない J〜N 列に計算を置く。
 * 当月入会の男 = 累計入会データで性別が分かる当月開始者の男 ＋ まだ載っていない人数 × 月初の男比率
 * 当月末退会の男 = 当月末退会 × 月初の男比率（退会メールに性別がないため）
 */
function ensureNippoGenderFlow_() {
  var nippo = gessho3Nippo_();
  var minus = gessho3GenderMinus_(nippo.getRange('B1').getDisplayValue());
  var known = [[minus.startMale, minus.startMale + minus.startFemale]];
  var cur = nippo.getRange('M13:N13').getValues();
  if (cur[0][0] !== known[0][0] || cur[0][1] !== known[0][1]) nippo.getRange('M13:N13').setValues(known);
  var want = {
    J12: '男の内訳（メール外）', M12: '分かる男', N12: '分かる人数',
    J13: '当月入会の男', J15: '当月末退会の男',
    K13: '=IF(C13<=0,0,IF(N13>=C13,ROUND(C13*M13/MAX(N13,1)),M13+ROUND((C13-N13)*F12/C12)))',
    K15: '=IF(C12<=0,0,ROUND(C15*F12/C12))',
    F16: '=F12+K13', F17: '=F16-K15'
  };
  var changed = false;
  Object.keys(want).forEach(function (a1) {
    var rg = nippo.getRange(a1);
    var v = want[a1];
    if (v.charAt(0) === '=') { if (rg.getFormula() !== v) { rg.setFormula(v); changed = true; } }
    else if (rg.getDisplayValue() !== v) { rg.setValue(v); changed = true; }
  });
  if (changed) nippo.getRange('J12:N15').setFontColor('#888888').setFontSize(9);
}

function fixOctoberNippo_() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('OCTOBER_NIPPO_FIX') === 'v1') return { ok: true, skipped: true };
  var options = [23, 211, 132, 72, 152, 26, 74, 2, 3, 112, 113, 66, 27, 229, 39, 21];
  var nippo = gessho3Nippo_();
  nippo.getRange('C12').setValue(1516).setNumberFormat('0').setNote(
    '月初会員 = 当月末在籍 1535 − 法人都度6 − OGF3 − ゴールド3 − 当月開始7 = 1516'
  );
  nippo.getRange('F12').setValue(1075).setNumberFormat('0');
  nippo.getRange('H12').setValue(441).setNumberFormat('0');
  nippo.getRange(21, 3, options.length, 1).setValues(options.map(function (n) { return [n]; }));

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.openById('1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q');
  var arch = ss.getSheetByName('月初_2026年10月');
  if (arch) {
    arch.getRange('B3').setValue(1516);
    arch.getRange('C3').setValue('当月末 1535 − 特例12 − 当月開始7');
    arch.getRange('B4').setValue(1075);
    arch.getRange('C4').setValue('年齢表1086 − 特例男5 − 当月入会男6');
    arch.getRange('B5').setValue(441);
    arch.getRange('C5').setValue('年齢表449 − 特例女7 − 当月入会女1');
    arch.getRange('C6').setValue('男女は日報へ書いた');
    arch.getRange(9, 2, options.length, 1).setValues(options.map(function (n) { return [n]; }));
    try { arch.hideSheet(); } catch (eArch) {}
  }
  var index = ss.getSheetByName('月初３ファイル');
  if (index) {
    var last = Math.max(index.getLastRow(), 1);
    var months = last > 1 ? index.getRange(2, 1, last - 1, 1).getDisplayValues() : [];
    for (var i = 0; i < months.length; i++) {
      if (String(months[i][0]) === '2026年10月') {
        index.getRange(i + 2, 3, 1, 3).setValues([[1516, 1075, 441]]);
      }
    }
    try { index.hideSheet(); } catch (eIndex) {}
  }
  props.setProperty('OCTOBER_NIPPO_FIX', 'v1');
  return { ok: true };
}

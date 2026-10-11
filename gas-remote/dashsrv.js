/**
 * 分析ダッシュボード（Webアプリ ?page=dash）。元のシートは読むだけ。氏名・メールは返さない。
 */
var DASH_ITEMS_ = ['jl', 'visit', 'op', 'hq', 'master', 'member', 'unpaid'];

function dashPage_() {
  var html = HtmlService.createHtmlOutputFromFile('dash');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  html.setTitle('分析ダッシュボード');
  return html;
}

function dashData(item, force) {
  if (DASH_ITEMS_.indexOf(item) < 0) return { ok: false, message: 'unknown item' };
  var cache = CacheService.getScriptCache();
  var key = 'dash4_' + item;
  if (!force) {
    var hit = dashCacheGet_(cache, key);
    if (hit) return hit;
  }
  var ss = openWorkspaceSpreadsheet_();
  var fn = { jl: dashJl_, visit: dashVisit_, op: dashOp_, hq: dashHq_, master: dashMaster_, member: dashMember_, unpaid: dashUnpaid_ }[item];
  var data = fn(ss);
  data.ok = true;
  data.at = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  data.today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var json = JSON.stringify(data);
  dashCachePut_(cache, key, json);
  return json;
}

function dashCacheGet_(cache, key) {
  var n = Number(cache.get(key + '_n') || 0);
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var parts = cache.getAll(keys);
  var s = '';
  for (var j = 0; j < n; j++) {
    if (parts[key + '_' + j] == null) return null;
    s += parts[key + '_' + j];
  }
  return s;
}

function dashCachePut_(cache, key, json) {
  var size = 90000, obj = {}, n = Math.ceil(json.length / size);
  if (n > 50) return;
  for (var i = 0; i < n; i++) obj[key + '_' + i] = json.slice(i * size, (i + 1) * size);
  obj[key + '_n'] = String(n);
  cache.putAll(obj, 600);
}

/** 'yyyy-MM-dd' と時。Date・シリアル値・'yyyy/m/d h:mm' 文字列に対応 */
function dashDate_(v) {
  if (v instanceof Date) {
    return { d: Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd'), h: Number(Utilities.formatDate(v, 'Asia/Tokyo', 'H')) };
  }
  if (typeof v === 'number' && v > 20000) {
    var t = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
    var p = function (x) { return (x < 10 ? '0' : '') + x; };
    return { d: t.getUTCFullYear() + '-' + p(t.getUTCMonth() + 1) + '-' + p(t.getUTCDate()), h: t.getUTCHours() };
  }
  var m = String(v).match(/^(\d{4})[\/\-年](\d{1,2})(?:[\/\-月](\d{1,2}))?(?:日)?(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  var q = function (x) { x = Number(x); return (x < 10 ? '0' : '') + x; };
  return { d: m[1] + '-' + q(m[2]) + '-' + q(m[3] || 1), h: m[4] ? Number(m[4]) : null };
}

function dashYm_(v) {
  if (v === '' || v === null || v === 0) return '';
  var r = dashDate_(v);
  return r ? r.d.slice(0, 7) : '';
}

function dashIdx_(list, v) {
  v = String(v || '').trim() || '（なし）';
  var i = list.indexOf(v);
  if (i < 0) { list.push(v); i = list.length - 1; }
  return i;
}

/** 入会・退会：受付メールの日時・区分 */
function dashJl_(ss) {
  var kinds = [], out = { kinds: kinds, join: [], leave: [] };
  var js = ss.getSheetByName('経堂_入会');
  if (js && js.getLastRow() > 1) {
    js.getRange(2, 1, js.getLastRow() - 1, 4).getValues().forEach(function (r) {
      var t = dashDate_(r[0]);
      if (t) out.join.push([t.d, t.h, dashIdx_(kinds, r[3]), dashYm_(r[2])]);
    });
  }
  var ls = ss.getSheetByName('経堂_退会');
  if (ls && ls.getLastRow() > 1) {
    ls.getRange(2, 1, ls.getLastRow() - 1, 6).getValues().forEach(function (r) {
      if (r[5] === true || String(r[5]).toUpperCase() === 'TRUE') return;
      var t = dashDate_(r[0]);
      if (t) out.leave.push([t.d, t.h, dashIdx_(kinds, r[3]), dashYm_(r[2])]);
    });
  }
  return out;
}

/** オプション：受付メールから作った明細（自動付与・契約追加・停止） */
function dashOp_(ss) {
  var opts = [], cats = ['自動付与', '契約追加', '停止'], out = { opts: opts, cats: cats, ev: [] };
  var sh = ss.getSheetByName('分析用_OP明細');
  if (!sh || sh.getLastRow() < 2) return out;
  var canon = sh.getRange('I2:I40').getValues().map(function (r) { return String(r[0]).trim(); }).filter(String);
  canon.forEach(function (o) { opts.push(o); });
  sh.getRange(2, 2, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
    var t = dashDate_(r[0]);
    var c = cats.indexOf(String(r[2]));
    if (!t || c < 0) return;
    out.ev.push([t.d, c, dashIdx_(opts, r[3])]);
  });
  return out;
}

/** 会員動向（本部）：A1:N115 の表示値をそのまま */
function dashHq_(ss) {
  var sh = ss.getSheetByName('【経堂】会員動向');
  if (!sh) return { rows: [] };
  var v = sh.getRange('A1:N115').getDisplayValues();
  var now = new Date();
  var fy = now.getFullYear() - (now.getMonth() < 3 ? 1 : 0);
  return { fy: fy, rows: v.map(function (r, i) { return [i + 1].concat(r.slice(1)); }) };
}

/** 経堂マスタ：数字の表と今月の読み */
function dashMaster_(ss) {
  var sh = ss.getSheetByName('経堂マスタ');
  if (!sh) return { rows: [] };
  var reading = sh.getRange('F7:G14').getDisplayValues().filter(function (r) { return r[0] || r[1]; });
  var top = sh.getRange('A1:Q6').getDisplayValues(), head = [];
  for (var i = 0; i < top.length - 1; i++) {
    for (var j = 0; j < top[i].length; j++) {
      var lb = String(top[i][j]).trim();
      if (/^当(日|月)/.test(lb) && lb.length <= 8) head.push([lb, String(top[i + 1][j]).trim()]);
    }
  }
  return { rows: sh.getRange('A16:J60').getDisplayValues().filter(function (r) { return String(r[0]).trim(); }), reading: reading, head: head };
}

/** 会員分析：累計入会データ（入会月・退会月・プラン・年齢・性別）と累計退会データ（退会月・理由・在籍） */
function dashMember_(ss) {
  var plans = [], reasons = [], out = { plans: plans, reasons: reasons, mem: [], leave: [] };
  var j = ss.getSheetByName('累計入会データ');
  if (j && j.getLastRow() > 1) {
    var n = j.getLastRow() - 1;
    var iv = j.getRange(2, 9, n, 1).getValues();
    var tv = j.getRange(2, 22, n, 4).getValues();
    for (var i = 0; i < n; i++) {
      var jm = dashYm_(tv[i][0]);
      if (!jm) continue;
      out.mem.push([jm, dashYm_(tv[i][1]), dashIdx_(plans, tv[i][3]), Number(tv[i][2]) || 0, String(iv[i][0]).trim()]);
    }
  }
  var l = ss.getSheetByName('累計退会データ');
  if (l && l.getLastRow() > 1) {
    var m = l.getLastRow() - 1;
    var a = l.getRange(2, 9, m, 14).getValues();
    for (var k = 0; k < m; k++) {
      var um = dashYm_(a[k][12]);
      if (!um) continue;
      out.leave.push([um, dashIdx_(plans, a[k][13]), dashIdx_(reasons, a[k][10]), Number(a[k][7]) || 0, Number(a[k][1]) || 0, String(a[k][0]).trim()]);
    }
  }
  var freqs = [], times = [];
  out.freqs = freqs; out.times = times; out.survey = []; out.sreasons = [];
  var q = ss.getSheetByName('退会アンケート');
  if (q && q.getLastRow() > 1) {
    q.getRange(2, 1, q.getLastRow() - 1, 9).getValues().forEach(function (r) {
      var ym = dashYm_(r[8]);
      if (!ym || !String(r[0]).trim()) return;
      out.survey.push([ym, dashIdx_(out.sreasons, r[2]), dashIdx_(freqs, r[5]), dashIdx_(times, r[7])]);
    });
  }
  return out;
}

/** 見学・体験：申請日・区分・性別・年代・希望日・時刻・入会したか */
function dashVisit_(ss) {
  var kinds = [], sexes = [], ages = [], out = { kinds: kinds, sexes: sexes, ages: ages, ev: [] };
  var sh = ss.getSheetByName('見学体験申請');
  if (!sh || sh.getLastRow() < 2) return out;
  // 入会の受付メールと、メールアドレスまたは氏名（空白を除く）で照合。申し込みから KENGAKU_JOIN_WINDOW_DAYS_ 日以内の入会を数える
  var norm = function (s) { return String(s || '').replace(/[\s\u3000]/g, '').toLowerCase(); };
  var joins = {};
  var js = ss.getSheetByName('経堂_入会');
  if (js && js.getLastRow() > 1) {
    js.getRange(2, 1, js.getLastRow() - 1, 6).getValues().forEach(function (r) {
      var t = dashDate_(r[0]);
      if (!t) return;
      [norm(r[1]), norm(r[5])].forEach(function (k) { if (k) (joins[k] = joins[k] || []).push(t.d); });
    });
  }
  var dayN = function (d) { return Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10)) / 864e5; };
  sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(function (r) {
    var t = dashDate_(r[0]);
    if (!t || !String(r[1]).trim()) return;
    var w = dashDate_(r[7]);
    var tm = r[8] instanceof Date ? Number(Utilities.formatDate(r[8], 'Asia/Tokyo', 'H')) : (String(r[8]).match(/^(\d{1,2}):/) || [])[1];
    var n0 = dayN(t.d), best = null;
    [norm(r[2]), norm(r[3])].forEach(function (k) {
      (joins[k] || []).forEach(function (d) { var x = dayN(d) - n0; if (x >= -1 && x <= KENGAKU_JOIN_WINDOW_DAYS_ && (best === null || x < best)) best = Math.max(0, x); });
    });
    out.ev.push([t.d, dashIdx_(kinds, r[1]), dashIdx_(sexes, r[5]), dashIdx_(ages, r[6]), w ? w.d : '', tm == null ? null : Number(tm), best === null ? 0 : 1, best]);
  });
  return out;
}

/** 未納：未納管理_推移の月ごとの表（A〜K） */
function dashUnpaid_(ss) {
  var sh = ss.getSheetByName('未納管理_推移');
  if (!sh) return { rows: [] };
  var v = sh.getRange(6, 1, Math.max(1, sh.getLastRow() - 5), 11).getDisplayValues();
  return { head: v[0], rows: v.slice(1).filter(function (r) { return /^\d+年\d+月$/.test(String(r[0]).trim()); }) };
}

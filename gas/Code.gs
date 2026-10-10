/**
 * 規定ナビ データAPI(Google Apps Script)
 *
 * - 「規定データ」スプレッドシートに紐づけて使う(拡張機能 → Apps Script)。
 * - シート: 規定 / 条文 / 表 / 言い換え / よくある質問 / 情報(tools/build_data.py が作る xlsx を取り込む)
 * - 画面(GitHub Pages)からは POST(本文はJSON文字列、Content-Type: text/plain)で呼ぶ。
 * - 合言葉は2種類。スクリプト プロパティに設定する。
 *     PASS_GENERAL … 一般向け(公開範囲が「一般」の規定だけ返す)
 *     PASS_EXEC    … 役員向け(すべての規定を返す)
 *   どちらも未設定のあいだは、安全のため何も返さない。
 */

var MAX_FAIL = 30;          // 10分間にこの回数まちがえると一時ロック(全員共通)
var LOCK_SEC = 600;

// ---------------------------------------------------------------- 入口

function doGet() {
  return json_({ ok: true, app: 'kitei-navi', time: Date.now() });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が正しくありません' });
  }
  try {
    var auth = checkPass_(req.pass);
    if (!auth.ok) return json_(auth);
    if (req.action === 'check') return json_({ ok: true, role: auth.role, version: dataVersion_() });
    if (req.action === 'load') return json_(load_(auth.role));
    return json_({ ok: false, error: '不明な操作です: ' + req.action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// ---------------------------------------------------------------- 合言葉

function checkPass_(pass) {
  var props = PropertiesService.getScriptProperties();
  var general = String(props.getProperty('PASS_GENERAL') || '').trim();
  var exec = String(props.getProperty('PASS_EXEC') || '').trim();
  if (!general && !exec) {
    return { ok: false, error: 'GAS のスクリプト プロパティ PASS_GENERAL / PASS_EXEC が設定されていません', code: 'NO_PASSCODE' };
  }
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fail') || 0);
  if (fails >= MAX_FAIL) {
    return { ok: false, error: '合言葉を何度もまちがえたため、10分ほどロックしています', code: 'LOCKED' };
  }
  var p = String(pass || '').trim();
  if (exec && p === exec) return { ok: true, role: '役員' };
  if (general && p === general) return { ok: true, role: '一般' };
  cache.put('fail', String(fails + 1), LOCK_SEC);
  return { ok: false, error: '合言葉が違います', code: 'PASS' };
}

// ---------------------------------------------------------------- データ

function load_(role) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allowAll = role === '役員';

  var regs = [];
  var allowed = {};
  table_(ss, '規定').forEach(function (r) {
    var scope = r['公開範囲'] || '役員';   // 空欄は安全側(役員のみ)に倒す
    if (!allowAll && scope !== '一般') return;
    allowed[r['id']] = true;
    regs.push({
      id: r['id'], no: r['番号'], name: r['規定名'], category: r['分類'], dept: r['所管'],
      scope: scope, summary: r['要約'], kind: r['種類'] || '規程'
    });
  });

  // 条文: [規定id, 順番, 章, 条, 見出し, 本文]
  var sections = [];
  rawRows_(ss, '条文').forEach(function (v) {
    if (allowed[v[0]]) sections.push([v[0], Number(v[1]), v[2], v[3], v[4], v[5]]);
  });

  // 表: [規定id, シート, 行, セル(JSON文字列)]
  var tables = [];
  rawRows_(ss, '表').forEach(function (v) {
    if (allowed[v[0]]) tables.push([v[0], v[1], Number(v[2]), v[3]]);
  });

  var synonyms = table_(ss, '言い換え').map(function (r) {
    return { word: r['検索語'], alts: r['言い換え(カンマ区切り)'] };
  }).filter(function (s) { return s.word; });

  var faq = table_(ss, 'よくある質問').filter(function (r) {
    return r['ボタン名'] && (allowAll || (r['公開範囲'] || '一般') === '一般');
  }).map(function (r) {
    return { label: r['ボタン名'], query: r['検索語'] };
  });

  return {
    ok: true, role: role, version: dataVersion_(),
    regs: regs, sections: sections, tables: tables, synonyms: synonyms, faq: faq
  };
}

/** 「情報」シートのデータ作成日(画面の「データ作成日」に出す) */
function dataVersion_() {
  var info = table_(SpreadsheetApp.getActiveSpreadsheet(), '情報');
  for (var i = 0; i < info.length; i++) {
    if (info[i]['項目'] === 'データ作成日') return info[i]['値'];
  }
  return '';
}

// ---------------------------------------------------------------- シート操作

/** 見出し行をキーにしたオブジェクトの配列 */
function table_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var v = sh.getDataRange().getDisplayValues();
  var head = v[0];
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i].join('') === '') continue;
    var o = {};
    head.forEach(function (h, j) { o[String(h).trim()] = String(v[i][j]).trim(); });
    out.push(o);
  }
  return out;
}

/** 見出し行を除いた生の行 */
function rawRows_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var v = sh.getDataRange().getDisplayValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === '') continue;
    out.push(v[i]);
  }
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 動作確認用:エディタから実行すると件数をログに出す */
function check() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['規定', '条文', '表', '言い換え', 'よくある質問'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    Logger.log(n + ': ' + (sh ? sh.getLastRow() - 1 + '行' : 'シートがありません'));
  });
  var props = PropertiesService.getScriptProperties();
  Logger.log('PASS_GENERAL: ' + (props.getProperty('PASS_GENERAL') ? '設定済み' : '未設定'));
  Logger.log('PASS_EXEC: ' + (props.getProperty('PASS_EXEC') ? '設定済み' : '未設定'));
}

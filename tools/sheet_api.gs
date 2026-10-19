/**
 * sheet_api.gs — スプレッドシートを「DB API」にする（Google Apps Script）
 *
 * これと GitHub Pages の画面を組み合わせると、サーバー不要のWebアプリになります。
 *
 * 手順
 * 1. スプレッドシート（pikogen_db.xlsx をインポートしたもの）を開く
 * 2. 拡張機能 → Apps Script にこのファイルを貼る
 * 3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *      実行するユーザー：自分
 *      アクセスできるユーザー：全員
 * 4. 出てきた URL（.../exec）を assets/js/config.js の API_URL に貼る
 * 5. config.js の MODE を 'api' にする
 *
 * 画面側は ?action=db で全シートをCSV文字列として受け取り、ブラウザで組み立てます。
 * シートは非公開のままでOK（Web App の権限だけ全員に開ける）。
 */

var SHEETS = [
  '設定','マスタ_なぜ','マスタ_欲求','マスタ_認知','マスタ_対象','マスタ_IP','マスタ_感情','マスタ_動詞','マスタ_型',
  'パターン','パターン_IP','パターン_UI','パターン_素材','パターン_音','パターン_感情','パターン_対象',
  'エンジン','エンジン_手引き','ツイスト','ツイスト_動詞',
  '素材','UI','SE','BGM','Prefab','ジュース','IP事実'
];

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'db';
  try {
    if (action === 'db') {
      var data = {};
      SHEETS.forEach(function (name) {
        data[name] = sheetToCsv(name);
      });
      return json_({ ok: true, generated: new Date().toISOString(), data: data });
    }
    if (action === 'ping') {
      return json_({ ok: true, sheets: SHEETS.length });
    }
    return json_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** 1シートを CSV 文字列にする（画面側のパーサと形式を合わせる） */
function sheetToCsv(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return '';
  var values = sh.getDataRange().getDisplayValues();
  if (!values.length) return '';
  return values.map(function (row) {
    return row.map(csvCell_).join(',');
  }).join('\n');
}

function csvCell_(v) {
  var s = v == null ? '' : String(v);
  if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 補助：DBの健康診断 ---------------- */

/** 参照切れ・空欄を一覧にする。教室で使う前に1回実行する。 */
function validateDb() {
  var ss = SpreadsheetApp.getActive();
  var errs = [];

  function rows(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { errs.push(name + ': シートがありません'); return []; }
    var v = sh.getDataRange().getDisplayValues();
    if (v.length < 2) return [];
    var head = v[0].map(function (h) { return String(h).trim(); });
    return v.slice(1).filter(function (r) { return r.join('').trim() !== ''; }).map(function (r) {
      var o = {}; head.forEach(function (h, i) { if (h) o[h] = String(r[i]).trim(); }); return o;
    });
  }

  var ip = {}; rows('マスタ_IP').forEach(function (r) { ip[r.id] = 1; });
  var emo = {}; rows('マスタ_感情').forEach(function (r) { emo[r.id] = 1; });
  var aud = {}; rows('マスタ_対象').forEach(function (r) { aud[r.id] = 1; });
  var ui = {}; rows('UI').forEach(function (r) { ui[r.id] = 1; });
  var as = {}; rows('素材').forEach(function (r) { as[r.id] = 1; });
  var se = {}; rows('SE').forEach(function (r) { se[r.id] = 1; });
  var eng = {}; rows('エンジン').forEach(function (r) { eng[r.id] = 1; });

  var pats = rows('パターン').filter(function (r) { return r.id && String(r.active).toUpperCase() !== 'FALSE'; });
  if (pats.length < 5) errs.push('パターンが ' + pats.length + ' 件しかありません');
  pats.forEach(function (p) {
    if (!eng[p.engine]) errs.push(p.id + ': エンジン ' + p.engine + ' がありません');
    if (!p.oneline) errs.push(p.id + ': oneline が空です');
    if (!p.action) errs.push(p.id + ': action が空です');
    if (!p.success_tpl) errs.push(p.id + ': success_tpl が空です');
  });

  rows('パターン_IP').forEach(function (r) { if (!ip[r.ip_id]) errs.push(r.pattern_id + ': ip_id ' + r.ip_id + ' がありません'); });
  rows('パターン_UI').forEach(function (r) { if (!ui[r.ui_id]) errs.push(r.pattern_id + ': ui_id ' + r.ui_id + ' がありません'); });
  rows('パターン_素材').forEach(function (r) { if (!as[r.asset_id]) errs.push(r.pattern_id + ': asset_id ' + r.asset_id + ' がありません'); });
  rows('パターン_音').forEach(function (r) { if (!se[r.se_id]) errs.push(r.pattern_id + ': se_id ' + r.se_id + ' がありません'); });
  rows('パターン_感情').forEach(function (r) { if (!emo[r.emo_id]) errs.push(r.pattern_id + ': emo_id ' + r.emo_id + ' がありません'); });
  rows('パターン_対象').forEach(function (r) { if (!aud[r.aud_id]) errs.push(r.pattern_id + ': aud_id ' + r.aud_id + ' がありません'); });
  rows('ツイスト_動詞').forEach(function (r) { if (!r.twist_id) errs.push('ツイスト_動詞: twist_id が空です'); });

  // エンジンの手引きが5件未満なら警告
  var byEng = {};
  rows('エンジン_手引き').forEach(function (r) { byEng[r.engine_id] = (byEng[r.engine_id] || 0) + 1; });
  Object.keys(eng).forEach(function (k) { if ((byEng[k] || 0) < 5) errs.push(k + ': 手引きが ' + (byEng[k] || 0) + ' 件しかありません'); });

  Logger.log(errs.length ? '【要修正】\n' + errs.join('\n') : 'OK：参照切れ・空欄はありません');
  return errs;
}

/** 公開CSVを書き出したいとき（GitHub に置く csv/ フォルダ用） */
function exportCsv() {
  var ss = SpreadsheetApp.getActive();
  var folderName = 'pikogen_csv';
  var it = DriveApp.getFoldersByName(folderName);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  var made = [];
  SHEETS.forEach(function (name) {
    var text = sheetToCsv(name);
    if (!text) return;
    var f = folder.createFile(name + '.csv', text, MimeType.CSV);
    made.push(name + '.csv');
  });
  Logger.log('書き出しました: ' + folder.getUrl() + '\n' + made.join(', '));
  return folder.getUrl();
}

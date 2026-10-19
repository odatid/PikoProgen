/**
 * sheet_export.gs — スプレッドシート → assets/js/data.js
 *
 * 使い方
 * 1. このファイルをスプレッドシートの Apps Script に貼る
 * 2. シート名・列名を README の想定に合わせる（列名は data.js のキーと同じ）
 * 3. exportDataJs() を実行 → 実行ログに data.js の中身が出る
 * 4. そのまま assets/js/data.js に貼り替えてコミット
 *
 * 設計方針
 * - 列は「ヘッダー名」で読む（列を挿しても壊れない）
 * - 子テーブル（パターン_IP など）は ID 参照だけを持つ
 * - 出力は window.PK への代入形式（fetch 不要で file:// でも動く）
 */

var SHEET_PATTERN   = 'パターン';
var SHEET_IP        = 'マスタ_IP';
var SHEET_EMO       = 'マスタ_感情';
var SHEET_AUD       = 'マスタ_対象';
var SHEET_TYPE      = 'マスタ_型';
var SHEET_P_IP      = 'パターン_IP';
var SHEET_P_VERB    = 'パターン_動詞';
var SHEET_P_UI      = 'パターン_UI';
var SHEET_P_ASSET   = 'パターン_素材';
var SHEET_P_SE      = 'パターン_音';

function exportDataJs() {
  var ss = SpreadsheetApp.getActive();
  var out = [];
  out.push('/* 自動生成：' + new Date().toISOString() + ' — 手で編集しない */');
  out.push('window.PK = window.PK || {};');

  // ---- マスタ ----
  out.push('PK.AUDIENCES = ' + JSON.stringify(readRows(ss, SHEET_AUD).map(function (r) {
    return { id: s(r.id), label: s(r.label), note: s(r.note), dev: s(r.dev) || 'both' };
  }), null, 2) + ';');

  out.push('PK.IP_VALUES = ' + JSON.stringify(readRows(ss, SHEET_IP).map(function (r) {
    return { id: s(r.id), g: s(r.group), label: s(r.label), mech: s(r.mech) || '△', note: s(r.note) };
  }), null, 2) + ';');

  out.push('PK.EMOTIONS = ' + JSON.stringify(readRows(ss, SHEET_EMO).map(function (r) {
    return { id: s(r.id), label: s(r.label), ok: s(r.ok), ng: s(r.ng), tone: s(r.tone) };
  }), null, 2) + ';');

  // 動詞は「カテゴリ」列でグループ化
  var verbRows = readRows(ss, 'マスタ_動詞');
  var groups = {}, order = [];
  verbRows.forEach(function (r) {
    var g = s(r.group) || 'その他';
    if (!groups[g]) { groups[g] = []; order.push(g); }
    groups[g].push({ id: s(r.id), label: s(r.label), pat: s(r.pat), d: num(r.difficulty, 1) });
  });
  out.push('PK.VERB_GROUPS = ' + JSON.stringify(order.map(function (g) {
    return { g: g, verbs: groups[g] };
  }), null, 2) + ';');

  out.push('PK.TYPES = ' + JSON.stringify(readRows(ss, SHEET_TYPE).map(function (r) {
    return {
      id: s(r.id), label: s(r.label),
      engines: String(s(r.engines) || '').split(/[,、\s]+/).filter(String),
      level: num(r.level, 1)
    };
  }), null, 2) + ';');

  // ---- 子テーブル ----
  var pIp = {}, pVerb = {}, pUi = {}, pAsset = {}, pSe = {};
  readRows(ss, SHEET_P_IP).forEach(function (r) {
    var p = s(r.pattern_id); if (!p) return;
    pIp[p] = pIp[p] || {}; pIp[p][s(r.ip_id)] = num(r.weight, 0.5);
  });
  readRows(ss, SHEET_P_VERB).forEach(function (r) {
    var p = s(r.pattern_id); if (!p) return;
    pVerb[p] = pVerb[p] || []; pVerb[p].push(s(r.verb));
  });
  readRows(ss, SHEET_P_UI).forEach(function (r) {
    var p = s(r.pattern_id); if (!p) return;
    pUi[p] = pUi[p] || []; pUi[p].push([s(r.ui_id), s(r.req) || '必須']);
  });
  readRows(ss, SHEET_P_ASSET).forEach(function (r) {
    var p = s(r.pattern_id); if (!p) return;
    pAsset[p] = pAsset[p] || []; pAsset[p].push(s(r.asset_id));
  });
  readRows(ss, SHEET_P_SE).forEach(function (r) {
    var p = s(r.pattern_id); if (!p) return;
    pSe[p] = pSe[p] || []; pSe[p].push(s(r.se_id));
  });

  // ---- パターン本体 ----
  var patterns = readRows(ss, SHEET_PATTERN).filter(function (r) {
    return s(r.id) && String(s(r.active)).toUpperCase() !== 'FALSE';
  }).map(function (r) {
    var id = s(r.id);
    return {
      id: id,
      name: s(r.name),
      family: s(r.family),
      engine: s(r.engine_id),
      level: num(r.level, 1),
      input: s(r.input_type),
      duration: num(r.duration_sec, 60),
      difficulty: num(r.difficulty, 1),
      oneline: s(r.oneline),
      action: s(r.action),
      successTpl: s(r.success_tpl),
      failTpl: s(r.fail_tpl),
      loop: s(r.loop),
      retry: s(r.retry_reason),
      ipHook: s(r.ip_hook),
      nums: {
        target: num(r.num_target, 0),
        life: num(r.num_life, 3),
        fall: num(r.num_fall, 0),
        interval: num(r.num_interval, 0),
        speedup: num(r.num_speedup, 0)
      },
      verbs: pVerb[id] || [],
      verbGroup: s(r.verb_group),
      ip: pIp[id] || {},
      emo: String(s(r.emotions) || '').split(/[,、\s]+/).filter(String),
      aud: String(s(r.audiences) || '').split(/[,、\s]+/).filter(String),
      imageComp: s(r.image_comp),
      imageMood: s(r.image_mood),
      codeNotes: s(r.code_notes),
      ui: pUi[id] || [],
      assets: pAsset[id] || [],
      se: pSe[id] || [],
      bgm: s(r.bgm),
      warn: String(s(r.warn_tags) || '').split(/[,、\s]+/).filter(String)
    };
  });
  out.push('PK.PATTERNS = ' + JSON.stringify(patterns, null, 2) + ';');

  var text = out.join('\n');
  Logger.log(text);
  return text;
}

/** ヘッダー名 → 値 のオブジェクト配列にする */
function readRows(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) { Logger.log('シートがありません: ' + name); return []; }
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0].map(function (h) { return String(h).trim(); });
  return values.slice(1).filter(function (row) {
    return row.some(function (c) { return String(c).trim() !== ''; });
  }).map(function (row) {
    var o = {};
    head.forEach(function (h, i) { if (h) o[h] = row[i]; });
    return o;
  });
}

function s(v) { return v == null ? '' : String(v).trim(); }
function num(v, d) {
  var n = Number(v);
  return isNaN(n) ? d : n;
}

/** ある1パターンだけ、生成結果を確認したいとき */
function previewPattern(id) {
  var all = JSON.parse('[' + exportDataJs().match(/PK\.PATTERNS = (\[[\s\S]*?\]);/)[1] + ']');
  Logger.log(JSON.stringify(all[0], null, 2));
}

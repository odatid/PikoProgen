/* data.js / engines.js から DB定義JSONを作る（シート生成用） */
const fs=require('fs'),vm=require('vm'),path=require('path');
const ctx={console};vm.createContext(ctx);ctx.window=ctx;ctx.globalThis=ctx;
['assets/js/data.js','assets/js/engines.js'].forEach(f=>vm.runInContext(fs.readFileSync(f,'utf8'),ctx,{filename:f}));
const PK=ctx.PK;
const out={
  meta:{generated:new Date().toISOString(), patterns:PK.PATTERNS.length, engines:Object.keys(PK.ENGINES).length},
  weights:PK.WEIGHTS,
  audiences:PK.AUDIENCES,
  ips:PK.IP_VALUES,
  emotions:PK.EMOTIONS,
  verbs:PK.VERB_GROUPS,
  types:PK.TYPES,
  wantTo:PK.WANT_TO,
  perception:PK.PERCEPTION,
  whyPico:PK.WHY_PICO,
  ui:PK.UI_MASTER,
  assets:PK.ASSET_MASTER,
  se:PK.SE_MASTER,
  bgm:PK.BGM_MASTER,
  juice:PK.JUICE,
  prefabs:PK.PREFABS,
  patterns:PK.PATTERNS,
  engines:PK.ENGINES,
  guideFirst:PK.GUIDE_FIRST,
  guideLast:PK.GUIDE_LAST,
  firstStep:PK.FIRST_STEP,
  twists:PK.TWISTS,
  subverbTwist:PK.SUBVERB_TWIST,
  subverb:PK.SUBVERB,
  ipPatches:PK.IP_PATCHES,
  facts:PK.DNA_FACTS,
  picoSound:PK.PICO_SOUND,
  picoSe:PK.PICO_SE,
  picoLines:PK.PICO_LINES
};
fs.writeFileSync('tools/db.json', JSON.stringify(out,null,1));
console.log('written tools/db.json',
  (fs.statSync('tools/db.json').size/1024).toFixed(0)+'KB',
  'patterns='+out.patterns.length, 'engines='+Object.keys(out.engines).length);

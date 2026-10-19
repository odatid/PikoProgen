/* offline.html を作る（CSS/JS を1枚に結合） */
const fs = require('fs');
const path = require('path');
const root = __dirname;
const html = fs.readFileSync(path.join(root,'index.html'),'utf8');
const css = fs.readFileSync(path.join(root,'assets/css/style.css'),'utf8');
const js = ['config','data','engines','db','generate','app'].map(f=>fs.readFileSync(path.join(root,'assets/js/'+f+'.js'),'utf8')).join('\n\n');

function build(out, demo){
  let doc = html
    .replace('<link rel="stylesheet" href="assets/css/style.css">', '<style>\n'+css+'\n</style>')
    .replace(/<script src="assets\/js\/[a-z]+\.js"><\/script>\n?/g, '');
  const inline = (demo? '<script>window.PK_DEMO = '+JSON.stringify(demo)+';</script>\n':'')
    + '<script>\n'+js+'\n</script>\n</body>';
  doc = doc.replace('</body>', inline);
  fs.writeFileSync(out, doc);
  console.log(out, (doc.length/1024).toFixed(0)+'KB');
}

build(path.join(root,'offline.html'));
fs.mkdirSync('/tmp/shots',{recursive:true});
build('/tmp/shots/intro.html', null);
build('/tmp/shots/s2.html', {step:'d2'});
build('/tmp/shots/s7.html', {step:'cs'});
build('/tmp/shots/s8.html', {step:'pr'});
build('/tmp/shots/s9.html', {step:'un'});
build('/tmp/shots/s9b.html', {step:'s9'});

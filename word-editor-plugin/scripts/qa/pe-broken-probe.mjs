/**
 * גשש: צילום 1:1 של הפסקה, מוגדל בלי החלקה.
 *
 * צילום ב-`clip.scale > 1` מרנדר את הטקסט מחדש ברזולוציה גבוהה, ולכן ארטיפקט
 * ריסטור בגודל האמיתי **נעלם** ממנו. הדרך היחידה לראות את מה שהמשתמש רואה היא
 * לצלם 1:1 ולהגדיל את הפיקסלים.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { openApp } from './harness.mjs';

const DOCX = process.env.PROBE_DOCX;
const OUT = process.env.PROBE_OUT;
if (!DOCX || !OUT) throw new Error('PROBE_DOCX / PROBE_OUT חסרים');

const app = await openApp({ name: 'pe-broken', port: Number(process.env.QA_PORT ?? 9644) });
const buffer = readFileSync(DOCX);

async function openDocx(name) {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    buffer.toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-${name}',url:${JSON.stringify(dataUrl)},name:${JSON.stringify(name + '.docx')},size:${buffer.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 20000 });
  return app.js("document.querySelector('.doc-title-input')?.value");
}

/** מגדילה קטע מצילום 1:1 בלי החלקה, ומצלמת את התוצאה. */
async function magnify(dataUrl, sx, sy, sw, sh, zoom, out) {
  await app.js(
    `window.__mag = new Promise(function(res){
      var img = new Image();
      img.onload = function(){
        var old = document.getElementById('__magbox'); if (old) old.remove();
        var c = document.createElement('canvas');
        c.id = '__magbox';
        c.width = ${sw} * ${zoom}; c.height = ${sh} * ${zoom};
        c.style.cssText = 'position:fixed;left:0;top:0;z-index:999999;background:#fff';
        var g = c.getContext('2d');
        g.imageSmoothingEnabled = false;
        g.drawImage(img, ${sx}, ${sy}, ${sw}, ${sh}, 0, 0, c.width, c.height);
        document.body.appendChild(c);
        var r = c.getBoundingClientRect();
        res(JSON.stringify({ x: r.x, y: r.y, w: r.width, h: r.height }));
      };
      img.src = ${JSON.stringify(dataUrl)};
    })`,
  );
  const box = JSON.parse(await app.js('window.__mag'));
  await app.sleep(300);
  const shot = await app.cdp.send('Page.captureScreenshot', {
    format: 'png',
    clip: { x: box.x, y: box.y, width: box.w, height: box.h, scale: 1 },
    captureBeyondViewport: true,
  });
  writeFileSync(out, Buffer.from(shot.result.data, 'base64'));
  console.log('נשמר:', out, box.w + 'x' + box.h);
  await app.js("document.getElementById('__magbox')?.remove()");
}

try {
  console.log('נפתח:', await openDocx('pe-broken'));

  const rect = await app.js(`JSON.stringify((function(){
    var nodes = document.querySelectorAll('[data-source-node-id] *');
    for (var i=0;i<nodes.length;i++){
      var t = (nodes[i].textContent||'');
      if (t.indexOf('\\u05d5\\u05d4\\u05d9\\u05d4 \\u05d0\\u05e4\\u05e9\\u05e8') >= 0 && !nodes[i].children.length) {
        nodes[i].scrollIntoView({ block: 'center' });
        var r = nodes[i].getBoundingClientRect();
        return { x: r.x, y: r.y, w: r.width, h: r.height };
      }
    }
    return null;
  })())`);
  const r = JSON.parse(rect);
  console.log('פסקה 1:1:', rect);
  await app.sleep(600);

  // צילום 1:1 של כל החלון — בלי re-render
  const full = await app.cdp.send('Page.captureScreenshot', { format: 'png', clip: undefined });
  const dataUrl = 'data:image/png;base64,' + full.result.data;

  // כל השורה, פי 5
  await magnify(dataUrl, Math.round(r.x), Math.round(r.y), Math.round(r.w), Math.round(r.h), 5, OUT);
  // „אפשר” לבדו — הוא בצד ימין של השורה, פי 14
  await magnify(
    dataUrl,
    Math.round(r.x + r.w - 92),
    Math.round(r.y),
    72,
    Math.round(r.h),
    14,
    OUT.replace(/\.png$/, '-word.png'),
  );
} finally {
  await app.close?.();
}

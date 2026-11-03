/**
 * גשש מכריע: האם השבירה של הפ היא **הגופן** או **הצייר**.
 *
 * אותה פסקה, אותו מיקום, אותו גודל — ורק `font-family` מוחלף. אם השבירה
 * נעלמת עם גופן אחר, הסיבה היא שהמנוע בחר Calibri במקום Carizma. אם היא
 * נשארת, הסיבה היא בציור עצמו.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { openApp } from './harness.mjs';

const DOCX = process.env.PROBE_DOCX;
const OUTDIR = process.env.PROBE_OUTDIR;
if (!DOCX || !OUTDIR) throw new Error('PROBE_DOCX / PROBE_OUTDIR חסרים');

const app = await openApp({ name: 'pe-swap', port: Number(process.env.QA_PORT ?? 9645) });
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
  await app.sleep(250);
  const shot = await app.cdp.send('Page.captureScreenshot', {
    format: 'png',
    clip: { x: box.x, y: box.y, width: box.w, height: box.h, scale: 1 },
    captureBeyondViewport: true,
  });
  writeFileSync(out, Buffer.from(shot.result.data, 'base64'));
  await app.js("document.getElementById('__magbox')?.remove()");
}

const LOCATE = `JSON.stringify((function(){
  var nodes = document.querySelectorAll('[data-source-node-id] *');
  for (var i=0;i<nodes.length;i++){
    var t = (nodes[i].textContent||'');
    if (t.indexOf('\\u05d5\\u05d4\\u05d9\\u05d4 \\u05d0\\u05e4\\u05e9\\u05e8') >= 0 && !nodes[i].children.length) {
      nodes[i].scrollIntoView({ block: 'center' });
      window.__target = nodes[i];
      var r = nodes[i].getBoundingClientRect();
      var cs = getComputedStyle(nodes[i]);
      // שרשרת הטרנספורמים מעל האלמנט — האם יש zoom/scale שמרנדר מחדש
      var chain = [], el = nodes[i];
      while (el && el !== document.documentElement) {
        var s = getComputedStyle(el);
        if (s.transform !== 'none' || s.zoom !== '1' || s.filter !== 'none' || s.willChange !== 'auto') {
          chain.push({ cls: (el.className||'').toString().slice(0,40), transform: s.transform, zoom: s.zoom, filter: s.filter, willChange: s.willChange });
        }
        el = el.parentElement;
      }
      return { x: r.x, y: r.y, w: r.width, h: r.height, font: cs.fontFamily, size: cs.fontSize,
               weight: cs.fontWeight, smooth: cs.webkitFontSmoothing, rendering: cs.textRendering,
               dpr: window.devicePixelRatio, transforms: chain };
    }
  }
  return null;
})())`;

try {
  console.log('נפתח:', await openDocx('pe-swap'));
  const info = JSON.parse(await app.js(LOCATE));
  console.log('מצב הציור:', JSON.stringify(info, null, 2));
  const r = info;

  const FONTS = [
    ['calibri', ''],
    ['frankruehl', 'FrankRuehl'],
    ['david', 'David'],
    ['narkisim', 'Narkisim'],
    ['times', 'Times New Roman'],
  ];

  for (const [label, family] of FONTS) {
    await app.js(
      family === ''
        ? "window.__target.style.removeProperty('font-family')"
        : `window.__target.style.setProperty('font-family', ${JSON.stringify(family)}, 'important')`,
    );
    await app.sleep(500);
    const full = await app.cdp.send('Page.captureScreenshot', { format: 'png' });
    const dataUrl = 'data:image/png;base64,' + full.result.data;
    await magnify(
      dataUrl,
      Math.round(r.x + r.w - 92),
      Math.round(r.y),
      72,
      Math.round(r.h),
      14,
      `${OUTDIR}/swap-${label}.png`,
    );
    console.log('נשמר:', label);
  }
} finally {
  await app.close?.();
}

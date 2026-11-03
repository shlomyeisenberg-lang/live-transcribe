/**
 * גשש: למה אותו מסמך נראה שבור אצל אחד ותקין אצל אחר.
 *
 * החשד: קנה המידה של המסך (`devicePixelRatio`). ב-100% כל פיקסל CSS הוא פיקסל
 * מסך אחד, והקו הדק של הפ ב-Calibri נופל בין פיקסלים ומתפורר. ב-125% ומעלה
 * יש יותר פיקסלים לאותה אות, והקו שורד.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { openApp } from './harness.mjs';

const DOCX = process.env.PROBE_DOCX;
const OUTDIR = process.env.PROBE_OUTDIR;
if (!DOCX || !OUTDIR) throw new Error('PROBE_DOCX / PROBE_OUTDIR חסרים');

const app = await openApp({ name: 'pe-dpr', port: Number(process.env.QA_PORT ?? 9646) });
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
        c.width = Math.round(${sw} * ${zoom}); c.height = Math.round(${sh} * ${zoom});
        c.style.cssText = 'position:fixed;left:0;top:0;z-index:999999;background:#fff;width:' + c.width + 'px;height:' + c.height + 'px';
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
      var r = nodes[i].getBoundingClientRect();
      return { x: r.x, y: r.y, w: r.width, h: r.height, dpr: window.devicePixelRatio };
    }
  }
  return null;
})())`;

try {
  console.log('נפתח:', await openDocx('pe-dpr'));

  for (const dpr of [1, 1.25, 1.5, 2]) {
    await app.cdp.send('Emulation.setDeviceMetricsOverride', {
      width: 0,
      height: 0,
      deviceScaleFactor: dpr,
      mobile: false,
    });
    await app.sleep(900);
    const info = JSON.parse(await app.js(LOCATE));
    if (!info) {
      console.log(`dpr=${dpr}: הפסקה לא נמצאה`);
      continue;
    }
    const full = await app.cdp.send('Page.captureScreenshot', { format: 'png' });
    const dataUrl = 'data:image/png;base64,' + full.result.data;
    // הצילום הוא בפיקסלים פיזיים: מכפילים את קואורדינטות ה-CSS ב-dpr
    await magnify(
      dataUrl,
      Math.round((info.x + info.w - 92) * dpr),
      Math.round(info.y * dpr),
      Math.round(72 * dpr),
      Math.round(info.h * dpr),
      14 / dpr,
      `${OUTDIR}/dpr-${String(dpr).replace('.', '_')}.png`,
    );
    console.log(`dpr=${dpr} (דף מדווח ${info.dpr}) → נשמר`);
  }
} finally {
  await app.cdp.send('Emulation.clearDeviceMetricsOverride').catch(() => {});
  await app.close?.();
}

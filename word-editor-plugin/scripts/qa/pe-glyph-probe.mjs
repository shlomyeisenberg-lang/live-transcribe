/**
 * גשש: איך נראית האות פ בגופנים המועמדים, בהגדלה.
 */
import { writeFileSync } from 'node:fs';
import { openApp } from './harness.mjs';

const OUT = process.env.PROBE_OUT;
if (!OUT) throw new Error('PROBE_OUT חסר');

const app = await openApp({ name: 'pe-glyph', port: Number(process.env.QA_PORT ?? 9643) });

try {
  const box = await app.js(`(function(){
    var host = document.createElement('div');
    host.style.cssText = 'position:fixed;left:0;top:0;z-index:999999;background:#fff;padding:10px;direction:rtl';
    var fonts = ['Calibri','FrankRuehl','David','Narkisim','Times New Roman','Arial','Guttman Haim-Condensed'];
    host.innerHTML = fonts.map(function(f){
      return '<div style="display:flex;align-items:center;gap:18px;padding:4px 0;border-bottom:1px solid #eee">'
        + '<span style="font:12px monospace;color:#666;width:150px;direction:ltr">'+f+'</span>'
        + '<span style="font:64px \\''+f+'\\'">\\u05d0\\u05e4\\u05e9\\u05e8</span>'
        + '<span style="font:64px \\''+f+'\\'">\\u05e4</span>'
        + '<span style="font:64px \\''+f+'\\'">\\u05e3</span>'
        + '<span style="font:64px \\''+f+'\\'">\\u05e0\\u05e4\\u05e1\\u05e7</span>'
        + '</div>';
    }).join('');
    document.body.appendChild(host);
    var r = host.getBoundingClientRect();
    return JSON.stringify({ x: r.x, y: r.y, w: r.width, h: r.height });
  })()`);
  const s = JSON.parse(box);
  await app.sleep(500);
  const shot = await app.cdp.send('Page.captureScreenshot', {
    format: 'png',
    clip: { x: s.x, y: s.y, width: s.w, height: s.h, scale: 2 },
    captureBeyondViewport: true,
  });
  writeFileSync(OUT, Buffer.from(shot.result.data, 'base64'));
  console.log('נשמר:', OUT, s.w + 'x' + s.h);
} finally {
  await app.close?.();
}

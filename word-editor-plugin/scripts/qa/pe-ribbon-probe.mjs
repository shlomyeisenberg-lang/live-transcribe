/**
 * גשש: מה שם הגופן שהרצועה מציגה כשהסמן בטקסט העברי.
 *
 * זו הבדיקה שהמשתמש יכול לעשות בלי זכוכית מגדלת: Word יציג „Carizma”,
 * והתוסף — מה שהמנוע באמת בחר.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { openApp } from './harness.mjs';

const DOCX = process.env.PROBE_DOCX;
const OUT = process.env.PROBE_OUT;
if (!DOCX) throw new Error('PROBE_DOCX חסר');

const app = await openApp({ name: 'pe-ribbon', port: Number(process.env.QA_PORT ?? 9647) });
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

try {
  console.log('נפתח:', await openDocx('pe-ribbon'));

  const spot = JSON.parse(
    await app.js(`JSON.stringify((function(){
      var nodes = document.querySelectorAll('[data-source-node-id] *');
      for (var i=0;i<nodes.length;i++){
        var t = (nodes[i].textContent||'');
        if (t.indexOf('\\u05d5\\u05d4\\u05d9\\u05d4 \\u05d0\\u05e4\\u05e9\\u05e8') >= 0 && !nodes[i].children.length) {
          nodes[i].scrollIntoView({ block: 'center' });
          var r = nodes[i].getBoundingClientRect();
          return { x: r.x + r.width - 40, y: r.y + r.height / 2 };
        }
      }
      return null;
    })())`),
  );
  await app.sleep(500);
  await app.clickAt(spot.x, spot.y);
  await app.sleep(900);
  await app.tab('בית');
  await app.sleep(600);

  const shown = await app.js(`JSON.stringify((function(){
    var out = {};
    var inputs = document.querySelectorAll('input');
    for (var i=0;i<inputs.length;i++){
      var el = inputs[i];
      var cls = (el.className||'').toString();
      if (/font/i.test(cls) || /font/i.test(el.getAttribute('aria-label')||'') || /גופן/.test(el.getAttribute('aria-label')||'')) {
        out[cls + '|' + (el.getAttribute('aria-label')||'')] = el.value;
      }
    }
    return out;
  })())`);
  console.log('מה שהרצועה מציגה:', shown);

  if (OUT) {
    const box = JSON.parse(
      await app.js(`JSON.stringify((function(){
        var r = document.querySelector('.ribbon, [class*=ribbon]').getBoundingClientRect();
        return { x: r.x, y: r.y, w: r.width, h: r.height };
      })())`),
    );
    const shot = await app.cdp.send('Page.captureScreenshot', {
      format: 'png',
      clip: { x: box.x, y: box.y, width: box.w, height: Math.min(box.h, 120), scale: 2 },
    });
    writeFileSync(OUT, Buffer.from(shot.result.data, 'base64'));
    console.log('נשמר:', OUT);
  }
} finally {
  await app.close?.();
}

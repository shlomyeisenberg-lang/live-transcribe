/**
 * גשש: באיזה גופן מצויר טקסט עברי שמוצהר **רק** ב-`w:cs`.
 *
 * המסמך `1789062709486-קושר.docx` מצהיר `w:cs="Carizma"` ב-132 ריצות, בלי
 * `w:ascii` ובלי `w:hAnsi`, ועם `<w:rtl/>` בכל ריצה. `docDefaults` מצהיר
 * `w:asciiTheme="minorHAnsi"`, והערכה פותרת אותו ל-`Calibri`.
 *
 * השאלה שהגשש מודד: מה יוצא ב-`font-family` המחושב על הטקסט המצויר.
 */
import { readFileSync } from 'node:fs';
import { openApp } from './harness.mjs';

const DOCX = process.env.PROBE_DOCX;
if (!DOCX) throw new Error('PROBE_DOCX חסר');

const app = await openApp({ name: 'carizma', port: Number(process.env.QA_PORT ?? 9641) });

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

const SCAN = `(function(){
  var out = { spans: [], aliasStyle: null, pickerStyle: null };
  var alias = document.getElementById('otzaria-doc-font-aliases');
  out.aliasStyle = alias ? (alias.textContent || '').slice(0, 800) : null;
  var picker = document.getElementById('otzaria-picker-font-faces');
  out.pickerStyle = picker ? (picker.textContent || '').slice(0, 400) : null;

  var seen = {};
  var nodes = document.querySelectorAll('[data-source-node-id] *');
  for (var i = 0; i < nodes.length && out.spans.length < 12; i++) {
    var el = nodes[i];
    var text = (el.textContent || '').trim();
    if (!text || el.children.length) continue;
    if (!/[\\u05d0-\\u05ea]/.test(text)) continue;
    var cs = getComputedStyle(el);
    var key = cs.fontFamily + '|' + cs.fontSize;
    if (seen[key]) continue;
    seen[key] = 1;
    out.spans.push({
      text: text.slice(0, 24),
      fontFamily: cs.fontFamily,
      fontSize: cs.fontSize,
      inlineFont: el.style ? el.style.fontFamily : '',
      tag: el.tagName,
    });
  }

  // מה הדפדפן באמת פותר לשם Carizma, ומה ל-Calibri
  var ctx = document.createElement('canvas').getContext('2d');
  var PROBE = '\\u05d0\\u05e4\\u05e9\\u05e8 \\u05dc\\u05ea\\u05e8\\u05e5';
  function w(f){ ctx.font = f; return ctx.measureText(PROBE).width; }
  out.measure = {
    carizma_vs_serif: [w('72px "Carizma", serif'), w('72px serif')],
    carizma_vs_mono: [w('72px "Carizma", monospace'), w('72px monospace')],
    calibri_vs_serif: [w('72px "Calibri", serif'), w('72px serif')],
  };
  return JSON.stringify(out);
})()`;

try {
  const title = await openDocx('carizma');
  console.log('נפתח:', title);
  const raw = await app.js(SCAN);
  console.log(JSON.stringify(JSON.parse(raw), null, 2));

  // מה בורר הגופנים מציג
  const picker = await app.js(`JSON.stringify((function(){
    var out = [];
    var nodes = document.querySelectorAll('[data-value]');
    for (var i=0;i<nodes.length;i++){
      var v = nodes[i].getAttribute('data-value') || '';
      if (/ariz/i.test(v)) out.push({ value: v, avail: nodes[i].getAttribute('data-availability'), cls: nodes[i].className });
    }
    return { carizmaRows: out, total: nodes.length };
  })())`);
  console.log('בורר:', picker);
} finally {
  await app.close?.();
}

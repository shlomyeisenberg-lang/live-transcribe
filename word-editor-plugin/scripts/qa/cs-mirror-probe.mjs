/**
 * גשש: מה המנוע כותב ל-`w:rPr` כשמחילים עיצוב על טקסט עברי.
 *
 * הטענה שנבדקת: „בלחיצה על הדגשה המנוע כותב את הצד הלטיני בלבד (`w:b`) ולא את
 * התאום של הכתב המורכב (`w:bCs`)”. היא נשענה עד כה על שתי רשומות בעץ —
 * `src/engine/font-advanced.ts:38` וטבלת caveat ב-issue 4011 — ולא על ריצה
 * טרייה. הגשש הזה מודד אותה מחדש, כי היא זו שמכריעה גם אם נדרשת מראה משלנו
 * בדרך החוצה וגם מה נכון לדווח למעלה.
 *
 * למה זה חשוב: ריצה שמצהירה `<w:rtl/>` — וזו כל ריצה עברית בקובץ שנוצר
 * ב-Word — נקראת ב-Word מ**מחסנית הכתב המורכב**. נמדד שהצד הלטיני לבדו יוצא
 * שם Arial 12pt לא-מודגש.
 *
 * ## למה הגשש מיירט את ההעלאה ואינו משתמש ב-`app.docx()`
 *
 * `window.__qa.exportBase64` קורא ל-`superdoc.export` ישירות. כאן רוצים את
 * הבייטים שנכתבים **לקובץ**, כלומר את המסלול שהשמירה עוברת בו, ולכן היירוט
 * הוא על `fetch(uploadUrl, {method:'PUT'})` — בדיוק כמו ב-rtl-run-export-qa.mjs.
 *
 * הרצה:  node scripts/qa/cs-mirror-probe.mjs   (QA_PORT דורס 9389)
 *
 * ‏**מ-PowerShell**, לא מ-shell שה-cwd שלו `/c/...`: משם Chrome מדווח
 * `file:///C:/…` בעוד ההרתמה מבקשת `file:///c:/…`, והגשש מת על „דפדפן אחר”.
 */
import { openApp, unzip } from './harness.mjs';
import { buildDocx } from './docx-fixtures.mjs';

const RTL = '<w:bidi/>';

/** ריצה עברית שמצהירה `w:rtl`, כמו בכל קובץ שנוצר ב-Word. */
const DECLARED =
  `<w:p><w:pPr>${RTL}</w:pPr>` +
  `<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/><w:rtl/></w:rPr>` +
  `<w:t xml:space="preserve">כותרת מוצהרת</w:t></w:r></w:p>`;

/** ריצה עברית בלי הצהרה, כמו במסמך שנולד בעורך הזה. */
const PLAIN =
  `<w:p><w:pPr>${RTL}</w:pPr>` +
  `<w:r><w:t xml:space="preserve">כותרת רגילה</w:t></w:r></w:p>`;

/** מחסנית מלאה שנכנסת כמות שהיא — לבדיקת השימור בגלגול. */
const COMPLETE =
  `<w:p><w:pPr>${RTL}</w:pPr>` +
  `<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>` +
  `<w:b/><w:bCs/><w:sz w:val="36"/><w:szCs w:val="36"/><w:rtl/></w:rPr>` +
  `<w:t xml:space="preserve">כותרת שלמה</w:t></w:r></w:p>`;

const app = await openApp({ name: 'cs-mirror', port: Number(process.env.QA_PORT ?? 9389) });

async function captureUpload() {
  await app.js(
    `(function(){if(window.__origFetch)return;window.__origFetch=window.fetch.bind(window);` +
      `window.fetch=function(url,opts){` +
      `if(opts&&opts.method==='PUT'){` +
      `var body=opts.body;` +
      `var read=body&&body.arrayBuffer?body.arrayBuffer():Promise.resolve(null);` +
      `return read.then(function(buf){` +
      `if(buf){var b=new Uint8Array(buf),s='';for(var i=0;i<b.length;i++)s+=String.fromCharCode(b[i]);` +
      `window.__savedDocx=btoa(s);}` +
      `return new Response('',{status:200});});}` +
      `return window.__origFetch(url,opts);};})()`,
  );
  await app.js(
    `window.__qaHost.replies['fs.beginBinaryWrite']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{writeToken:'wt-cs',uploadUrl:'https://qa-upload.local/cs',maxBytes:999999999}})}`,
  );
  await app.js(
    `window.__qaHost.replies['fs.commitUserFileWrite']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{cancelled:false,token:'tok-cs',name:'cs.docx',size:1234}})}`,
  );
}

async function openDocx(buffer, name) {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(buffer).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-${name}',url:${JSON.stringify(dataUrl)},name:${JSON.stringify(name + '.docx')},size:${buffer.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 12000 });
  return app.js("document.querySelector('.doc-title-input')?.value");
}

/** מניחה את הסמן בשורה שהטקסט שלה מכיל `needle`, ומסמנת אותה כולה. */
async function selectLine(needle) {
  const spot = await app.js(
    `(function(){
      var frags=document.querySelectorAll('[data-source-node-id][data-pm-start]');
      for(var i=0;i<frags.length;i++){
        if((frags[i].textContent||'').indexOf(${JSON.stringify(needle)})<0) continue;
        var l=frags[i].children[0]; if(!l) return '';
        var r=l.getBoundingClientRect();
        return JSON.stringify({x:Math.round((r.left+r.right)/2),y:Math.round((r.top+r.bottom)/2)});
      }
      return '';
    })()`,
  );
  if (!spot) return false;
  const { x, y } = JSON.parse(spot);
  await app.clickAt(x, y);
  await app.sleep(400);
  await app.press('Home', 'Home', 36);
  await app.sleep(200);
  // ‏8 = Shift ב-CDP.
  await app.press('End', 'End', 35, 8);
  await app.sleep(600);
  return true;
}

/** ה-`rPr` של הריצה שהטקסט שלה מכיל `needle`, מהבייטים שנשמרו. */
function propsOf(xml, needle) {
  for (const run of xml.match(/<w:r[ >][\s\S]*?<\/w:r>/g) ?? []) {
    if (!run.includes(needle)) continue;
    const rPr = (run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) ?? [])[0];
    return rPr ?? '(אין rPr)';
  }
  return '(הריצה לא נמצאה)';
}

async function saveAndRead() {
  await app.js('window.__savedDocx = ""');
  await app.press('s', 'KeyS', 83, 2, 's');
  await app.sleep(6000);
  const saved = await app.js('window.__savedDocx || ""');
  if (!saved) return null;
  return unzip(Buffer.from(saved, 'base64'))['word/document.xml'] ?? '';
}

try {
  await captureUpload();
  const opened = await openDocx(
    buildDocx({ body: DECLARED + PLAIN + COMPLETE }),
    'cs-mirror',
  );
  console.log('המסמך שנפתח:', opened);

  console.log('\n=== 4. שימור בגלגול: שמירה בלי לגעת בכלום ===');
  const untouched = await saveAndRead();
  if (untouched === null) console.log('  לא נתפסה העלאה — ייתכן שהמסמך לא היה „מלוכלך”, וזה עצמו תשובה');
  else console.log('  „כותרת שלמה”:', propsOf(untouched, 'כותרת שלמה'));

  for (const [label, needle, key, code, vk, text] of [
    ['הדגשה', 'כותרת מוצהרת', 'b', 'KeyB', 66, 'b'],
    ['נטייה', 'כותרת מוצהרת', 'i', 'KeyI', 73, 'i'],
  ]) {
    console.log(`\n=== ${label} על ריצה שמצהירה w:rtl ===`);
    if (!(await selectLine(needle))) {
      console.log('  לא נמצאה השורה');
      continue;
    }
    await app.press(key, code, vk, 2, text);
    await app.sleep(1200);
    const xml = await saveAndRead();
    if (xml === null) console.log('  לא נתפסה העלאה');
    else console.log('  ה-rPr שיצא:', propsOf(xml, needle));
  }

  console.log('\n=== הדגשה על ריצה בלי הצהרה (בקרה) ===');
  if (await selectLine('כותרת רגילה')) {
    await app.press('b', 'KeyB', 66, 2, 'b');
    await app.sleep(1200);
    const xml = await saveAndRead();
    console.log('  ה-rPr שיצא:', xml === null ? 'לא נתפסה העלאה' : propsOf(xml, 'כותרת רגילה'));
  } else {
    console.log('  לא נמצאה השורה');
  }
} finally {
  await app.close();
}

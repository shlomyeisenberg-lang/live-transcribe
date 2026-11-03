/**
 * גשש חד-פעמי: מה קיים בזמן בחירה ב-Shift+חץ.
 *
 * שער החצים החזיר שורות ריקות לכל מדידות ה-Shift — לא היסט ולא סמן מצויר —
 * ולכן הבדיקות שם עברו בחלל. השאלה: האם התצלום הסינכרוני עדיין מוסר את ראש
 * הבחירה, ומה מצויר במקום הסמן. שתיהן מכריעות אם `Shift+חץ` בר-תיקון.
 */
import { openApp } from './harness.mjs';
import { zipStored } from './docx-fixtures.mjs';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const body =
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>אבגדהוז ראשון גדול</w:t></w:r></w:p>`;

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:bidi/></w:sectPr></w:body></w:document>`;

const docx = zipStored({
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/document.xml': documentXml,
});

const app = await openApp({ name: 'shift-arrow', port: Number(process.env.QA_PORT ?? 9716) });

/** הבחירה דרך ה-API הא-סינכרוני, שהוא היחיד שמתאר טווחים. */
const asyncSel = () =>
  app.js(`(async function(){
    try {
      var ed = window.__otzariaEditor.superdoc.activeEditor;
      var info = await ed.doc.selection.current();
      var seg = info && info.target && info.target.segments && info.target.segments[0];
      return JSON.stringify({ empty: info ? info.empty : null, blockId: seg ? seg.blockId : null, range: seg ? seg.range : null });
    } catch (e) { return 'threw: ' + String(e).slice(0, 80); }
  })()`);

/** כל מה שמצויר ושמו רומז על בחירה — לא רק הסמן. */
const painted = () =>
  app.js(`(function(){
    var out = [];
    document.querySelectorAll('*').forEach(function (el) {
      var c = String(el.className || '');
      if (c.indexOf('selection') < 0 && c.indexOf('caret') < 0) return;
      var r = el.getBoundingClientRect();
      if (!r.width && !r.height) return;
      out.push(c.slice(0, 44) + ' @' + Math.round(r.left) + '..' + Math.round(r.right));
    });
    return JSON.stringify(out.slice(0, 8));
  })()`);

const snap = () =>
  app.js(`(function(){
    var ed = window.__otzariaEditor && window.__otzariaEditor.superdoc && window.__otzariaEditor.superdoc.activeEditor;
    var s = ed && ed.host && ed.host.readLiveSelectionSyncSnapshot && ed.host.readLiveSelectionSyncSnapshot();
    var marks = [];
    document.querySelectorAll('[class*="sd-v2-local-selection"]').forEach(function (el) {
      var r = el.getBoundingClientRect();
      marks.push(el.className + ' @' + Math.round(r.left) + '..' + Math.round(r.right) + ' y' + Math.round(r.top));
    });
    return JSON.stringify({ snapshot: s, marks: marks }, function (k, v) {
      return v === undefined ? null : v;
    });
  })()`);

try {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-shift',url:${JSON.stringify(dataUrl)},name:'shift.docx',size:${docx.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.click('עיון בקבצים…', { after: 8000 });
  for (let waited = 0; waited < 40_000; waited += 250) {
    await app.sleep(250);
    if (!(await app.exists('.status-load'))) break;
  }
  await app.sleep(2500);
  await app.tab('בית');
  await app.caretPara('אבגדהוז ראשון גדול');

  console.log('\n-- סמן מכווץ --');
  console.log(await snap());

  /* האם המנוע מקבל טווח הפוך? זו השאלה שמכריעה אם `Shift+חץ` בר-תיקון:
     בעברית הראש נע **אחורה** בהיסטים, ולכן start>end. */
  const write = (a, b) =>
    app.js(`(async function(){
      var ed = window.__otzariaEditor.superdoc.activeEditor;
      var story = { kind: 'story', storyType: 'body' };
      var pt = function (o) { return { kind: 'text', blockId: '00000001', offset: o, story: story }; };
      try {
        await ed.authoring.setSelectionTarget({ target: { kind: 'selection', start: pt(${a}), end: pt(${b}), story: story }, focus: true });
      } catch (e) { return 'threw: ' + String(e).slice(0, 80); }
      return 'ok';
    })()`);

  await app.caretPara('אבגדהוז ראשון גדול');
  console.log('== כתיבה ידנית start=2 end=6 (קדימה) ==');
  console.log(await write(2, 6));
  await app.sleep(500);
  console.log('sync   :', await snap());
  console.log('async  :', await asyncSel());
  console.log('מצויר  :', await painted());

  await app.caretPara('אבגדהוז ראשון גדול');
  console.log('== כתיבה ידנית start=6 end=2 (הפוך) ==');
  console.log(await write(6, 2));
  await app.sleep(500);
  console.log('sync   :', await snap());
  console.log('async  :', await asyncSel());
  console.log('מצויר  :', await painted());

  await app.caretPara('אבגדהוז ראשון גדול');
  for (let i = 1; i <= 3; i += 1) {
    await app.press('ArrowRight', 'ArrowRight', 39, 8);
    await app.sleep(350);
    console.log(`\n-- אחרי ${i} × Shift+ArrowRight --`);
    console.log(await snap());
  }
} catch (error) {
  console.error(error);
} finally {
  app.close();
}

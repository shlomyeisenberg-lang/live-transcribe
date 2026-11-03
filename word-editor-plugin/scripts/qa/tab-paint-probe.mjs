/**
 * גשש חד-פעמי: איך המנוע מצייר `<w:tab/>` בתוך שורה, ומה זה עושה למיפוי
 * אינדקס-תו → היסט שעליו נשען `src/engine/rtl-caret.ts`.
 *
 * שער החצים מצא שבפסקה עם טאב היירוט אינו נכנס כלל. החשד: ספירת התווים
 * המצוירים אינה תואמת את טווח ה-pm של השורה, וזה בדיוק השומר שמחזיר את
 * ההקשה למנוע. הגשש מדפיס את המבנה כדי שההכרעה תהיה על מדידה.
 */
import { openApp } from './harness.mjs';
import { zipStored } from './docx-fixtures.mjs';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const body = `<w:p><w:pPr><w:bidi/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>פריט ברשימה</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:bidi/></w:pPr>` +
  `<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">קודם</w:t></w:r>` +
  `<w:r><w:tab/></w:r>` +
  `<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">אחרון</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>שורה שנייה ארוכה למדי כדי שתגלוש מעבר לרוחב העמוד ותיתן שתי שורות מצוירות באותה פסקה אחת</w:t></w:r></w:p>`;

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:bidi/></w:sectPr></w:body></w:document>`;

const docx = zipStored({
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
  '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
  'word/numbering.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  'word/document.xml': documentXml,
});

const app = await openApp({ name: 'tab-paint', port: Number(process.env.QA_PORT ?? 9715) });

try {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-tab',url:${JSON.stringify(dataUrl)},name:'tab.docx',size:${docx.length},access:'readwrite'}})}`,
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

  const dump = await app.js(`(function(){
    var out = [];
    var frags = document.querySelectorAll('[data-source-node-id][data-pm-start]');
    for (var f = 0; f < frags.length; f++) {
      var frag = frags[f];
      var lines = [];
      for (var i = 0; i < frag.children.length; i++) {
        var line = frag.children[i];
        var kids = [];
        var carriers = line.querySelectorAll('.superdoc-text-run, .superdoc-tab');
        for (var j = 0; j < line.children.length; j++) {
          var el = line.children[j];
          var r = el.getBoundingClientRect();
          kids.push({
            tag: el.tagName,
            cls: String(el.className).slice(0, 44),
            text: JSON.stringify((el.textContent || '').slice(0, 14)),
            len: (el.textContent || '').length,
            w: Math.round(r.width * 10) / 10,
            pm: el.getAttribute('data-pm-start') + '..' + el.getAttribute('data-pm-end')
          });
        }
        var textLen = 0;
        var walker = document.createTreeWalker(line, NodeFilter.SHOW_TEXT, null);
        var n, inRun = 0;
        while ((n = walker.nextNode())) {
          textLen += n.data.length;
          if (n.parentElement && n.parentElement.closest('.superdoc-text-run')) inRun += n.data.length;
        }
        var carrierPm = [];
        for (var c = 0; c < carriers.length; c++) {
          carrierPm.push(carriers[c].className + '=' + carriers[c].getAttribute('data-pm-start') + '..' + carriers[c].getAttribute('data-pm-end') + ' len=' + (carriers[c].firstChild && carriers[c].firstChild.nodeType === 3 ? carriers[c].firstChild.data.length : 'no-text'));
        }
        lines.push({
          carriers: carrierPm,
          pmStart: Number(line.getAttribute('data-pm-start')),
          pmEnd: Number(line.getAttribute('data-pm-end')),
          dir: line.getAttribute('dir'),
          span: Number(line.getAttribute('data-pm-end')) - Number(line.getAttribute('data-pm-start')),
          textLen: textLen,
          inRun: inRun,
          kids: kids
        });
      }
      out.push({ id: frag.getAttribute('data-source-node-id'), base: Number(frag.getAttribute('data-pm-start')), parent: frag.parentElement ? frag.parentElement.tagName + '.' + String(frag.parentElement.className).slice(0,30) : null, lines: lines });
    }
    return JSON.stringify(out);
  })()`);

  for (const frag of JSON.parse(dump)) {
    console.log(`\n=== fragment ${frag.id} base=${frag.base}`);
    for (const line of frag.lines) {
      console.log(
        `  שורה [${line.pmStart}..${line.pmEnd}] span=${line.span} dir=${line.dir ?? '—'} ` +
          `תווי-טקסט=${line.textLen} מתוכם בריצה=${line.inRun}`,
      );
      for (const k of line.kids) {
        console.log(`     <${k.tag}.${k.cls}> pm=${k.pm} len=${k.len} w=${k.w} ${k.text}`);
      }
      console.log(`     נושאי pm: ${line.carriers.join(' | ')}`);
    }
  }
} catch (error) {
  console.error(error);
} finally {
  app.close();
}

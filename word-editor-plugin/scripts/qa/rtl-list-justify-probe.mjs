/**
 * סקר: השורה הראשונה של פריט רשימה עברי נשברת לפני הקצה.
 *
 * מה שדווח: „באפשרות מספור הטקסט לא אחיד, ושורה לא מסתיימת בסוף השורה אלא
 * מעט לפניה”. בפריט רשימה עם `<w:bidi/>` והזחה תלויה (ברירת המחדל של Word
 * וגם של כפתור „מספור” ברצועה: `left=720 hanging=360`) השורה הראשונה נשברת
 * מוקדם, וביישור לשני הצדדים נמתחת לרוחב קצר מזה של שאר השורות. ב-LTR אותה
 * רשימה תקינה, וב-Word תקינה בשני הכיוונים.
 *
 * השורש במנוע, בקוד הפתוח — superdoc/docx-editor#4085, ופירוט ב-
 * docs/engine-gaps.md. אין לזה עקיפה מבחוץ שאינה משנה את המסמך.
 *
 * הסקר פותח מסמך עם פסקאות עבריות ממוספרות/עם תבליט ביישור לשני הצדדים
 * ולימין, ופסקת ביקורת LTR, ובודק ב-DOM שקצה הסוף של השורה הראשונה זהה
 * לזה של השורה השנייה — כלומר שהמנוע תוקן. שער סקר: אינו מפיל את הריצה.
 *
 * שימוש:  node scripts/qa/rtl-list-justify-probe.mjs   (QA_PORT דורס 9614)
 */
import { openApp, createReport } from './harness.mjs';
import { buildDocx } from './docx-fixtures.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const HEB =
  'אמר רבי חנינא לא מצאתי לגוף טוב אלא שתיקה וכל המרבה דברים מביא חטא והנה זה כלל גדול בתורה ואין לך דבר שעומד בפני התשובה ולכן צריך אדם לפשפש במעשיו בכל יום ויום ולתקן את דרכיו';
const ENG =
  'The quick brown fox jumps over the lazy dog and keeps running across the wide field while the farmer watches from far away and wonders';
const numPr = (id) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr>`;
const para = (text, pPr, rtl = true) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r>${rtl ? '<w:rPr><w:rtl/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const lvl = (fmt, text) =>
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;
const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvl('decimal', '%1.')}</w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvl('bullet', '&#8226;')}</w:abstractNum>
<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>${lvl('hebrew1', '%1.')}</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`;

const B = '<w:bidi/>';
const J = '<w:jc w:val="both"/>';
/** [שם, פסקה, האם RTL] — הסדר הוא סדר ה-fragments במסמך. */
const CASES = [
  ['פסקה רגילה, לשני הצדדים (ביקורת)', para(HEB, B + J), true],
  ['מספור 1., לשני הצדדים', para(HEB, numPr(1) + B + J), true],
  ['מספור א., לשני הצדדים', para(HEB, numPr(3) + B + J), true],
  ['תבליט, לשני הצדדים', para(HEB, numPr(2) + B + J), true],
  ['מספור 1., לימין', para(HEB, numPr(1) + B), true],
  ['LTR מספור, לשני הצדדים (ביקורת)', para(ENG, numPr(1) + J, false), false],
];
const docx = buildDocx({ body: CASES.map((c) => c[1]).join(''), numbering });

const report = createReport('רשימה עברית — השורה הראשונה מגיעה לקצה', { strict: false });
const app = await openApp({ name: 'rtl-list-justify', port: Number(process.env.QA_PORT ?? 9614) });

try {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-rlj',url:${JSON.stringify(dataUrl)},name:'rtl-list-justify.docx',size:${docx.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 12000 });

  // לכל שורה: מרחק קצה הטקסט (בלי הסמן) משני צדי העמוד.
  const frags = JSON.parse(
    await app.js(`(function(){
      var page = document.querySelector('[data-page-index]');
      if (!page) return '[]';
      var pr = page.getBoundingClientRect();
      return JSON.stringify(Array.prototype.map.call(document.querySelectorAll('.superdoc-fragment[data-source-node-id]'), function(frag){
        return Array.prototype.map.call(frag.querySelectorAll('.superdoc-line'), function(line){
          var marker = line.querySelector('.superdoc-list-marker');
          var minL = 1e9, maxR = -1e9, w = document.createTreeWalker(line, NodeFilter.SHOW_TEXT), n;
          while ((n = w.nextNode())) {
            if ((marker && marker.contains(n)) || !n.textContent.trim()) continue;
            var rg = document.createRange(); rg.selectNodeContents(n);
            Array.prototype.forEach.call(rg.getClientRects(), function(r){ if (r.width > 0) { minL = Math.min(minL, r.left); maxR = Math.max(maxR, r.right); } });
          }
          return { left: Math.round(minL - pr.left), right: Math.round(pr.right - maxR) };
        });
      }));
    })()`),
  );

  CASES.forEach(([name, , rtl], i) => {
    const lines = frags[i];
    if (!lines || lines.length < 2) {
      report.skip(name, `לא נמצאו שתי שורות מצוירות (${lines ? lines.length : 0})`);
      return;
    }
    // קצה הסוף: שמאל ב-RTL, ימין ב-LTR. השורה השנייה אינה האחרונה רק בחלק
    // מהמקרים, ולכן משווים לשוליים של הפסקה הרגילה (השורה הראשונה שלה).
    const end = (l) => (rtl ? l.left : l.right);
    const margin = end(frags[0][0]);
    const first = end(lines[0]);
    const detail = `קצה הסוף של השורה הראשונה ${first}px מהעמוד, השוליים ${margin}px`;
    if (name.includes('לימין')) {
      // ביישור לימין השורה אינה נמתחת; השאלה היא אם נשברה מוקדם. השורה
      // השנייה מלאה (הפסקה בת שלוש שורות), ולכן היא קנה המידה.
      const second = end(lines[1]);
      Math.abs(first - second) <= 20
        ? report.pass(name, `${detail}; השנייה ${second}px`)
        : report.fail(name, `${detail}; השנייה ${second}px — נשברה ${first - second}px מוקדם`);
      return;
    }
    Math.abs(first - margin) <= 3 ? report.pass(name, detail) : report.fail(name, `${detail} — חסרים ${first - margin}px`);
  });
} finally {
  app.close();
}
report.print();

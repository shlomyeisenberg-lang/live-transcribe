/**
 * גשש: `format.paragraph.setFlowOptions` — patch או replace, ומה חוזר בקריאה.
 *
 * למה זה נשאל: הדיאלוג „פסקה” אמור לקבל את התיבה „אל תוסיף רווח בין פסקאות
 * מאותו סגנון” (‏`w:contextualSpacing`), והיא נכתבת דרך אותה פעולה שכותבת את
 * „מעבר עמוד לפני” (‏`pageBreakBefore`, engine/page-break.ts). שתי שאלות
 * מכריעות אם התיבה בכלל אפשרית, ולאף אחת מהן אין תשובה בתיעוד:
 *
 * 1. **האם כתיבה אחת מוחקת את השנייה?** `setSpacing` הוא replace מוכח (ראו
 *    הערת הפתיחה של engine/paragraph-format.ts). אם `setFlowOptions` נוהג
 *    כמותו, סימון התיבה היה מוחק מעבר עמוד שהמשתמש הגדיר — ובשקט.
 * 2. **האם `doc.get()` מחזיר את הערך?** על `pageBreakBefore` נמדד שלא, ולכן
 *    הכפתור ברצועה מחזיק מעקב מקומי. תיבה בדיאלוג אינה יכולה להחזיק מעקב
 *    כזה: היא נפתחת על תצלום, ותיבה שנפתחת ריקה ומאושרת **מכבה** את מה
 *    שהיה.
 *
 *   npm run build && node scripts/flow-options-probe.mjs
 */
import { openApp, sleep } from './qa/harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9371);

/** הפסקה שנמדדת, לפי הטקסט שלה. */
const TARGET_TEXT = 'זוגי';

const lines = [];
const say = (text) => {
  lines.push(text);
  console.log(text);
};

const app = await openApp({ name: 'flow-options', port: PORT });

try {
  // פסקה אחת מספיקה: שתי השאלות נשאלות על אותה פסקה בדיוק. שורה שנייה הייתה
  // דורשת Enter, ו-`type` של המסגרת ממפה כל תו שאינו אות ל-`Space`.
  await app.caret(0);
  await app.type('זוגי');
  await sleep(1500);
  await app.caret(0);
  await sleep(400);

  /** `nodeId` של הפסקה שהסמן בה. */
  const blockId = await app.js(
    '(async () => { const s = await window.__otzariaEditor.superdoc.activeEditor.doc.selection.current(); ' +
      'return (s.target && s.target.segments && s.target.segments[0] && s.target.segments[0].blockId) || null; })()',
  );
  say(`פסקת היעד: ${blockId}`);

  const flow = (fields) =>
    app.js(
      `(async () => { const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
         try {
           const r = await doc.format.paragraph.setFlowOptions({ target: { kind: 'block', nodeType: 'paragraph', nodeId: ${JSON.stringify(blockId)} }, ...${JSON.stringify(fields)} });
           return JSON.stringify({ success: r && r.success, failure: r && r.failure });
         } catch (e) { return JSON.stringify({ threw: String(e && e.message) }); }
       })()`,
    );

  const pPrOf = (files) => {
    const xml = files['word/document.xml'] ?? '';
    const paragraph = (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).find((p) => p.includes(`>${TARGET_TEXT}<`));
    return paragraph?.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? '';
  };

  /* -- 1. מעבר עמוד לבדו -- */
  say(`\nsetFlowOptions({pageBreakBefore:true}) → ${await flow({ pageBreakBefore: true })}`);
  await sleep(900);
  const afterBreak = pPrOf(await app.docx());
  say(`  pPr: ${afterBreak || '(ריק)'}`);

  /* -- 2. ועכשיו contextualSpacing, על אותה פסקה -- */
  say(`\nsetFlowOptions({contextualSpacing:true}) → ${await flow({ contextualSpacing: true })}`);
  await sleep(900);
  const afterBoth = pPrOf(await app.docx());
  say(`  pPr: ${afterBoth || '(ריק)'}`);

  const keptBreak = /<w:pageBreakBefore/.test(afterBoth);
  const wroteContextual = /<w:contextualSpacing/.test(afterBoth);
  say(
    `\n>> patch או replace: ${
      !wroteContextual
        ? 'contextualSpacing לא נכתב כלל'
        : keptBreak
          ? 'PATCH — שתי התכונות חיות יחד'
          : 'REPLACE — הכתיבה השנייה מחקה את pageBreakBefore'
    }`,
  );

  /* -- 3. מה `doc.get()` מחזיר על אותה פסקה -- */
  const props = await app.js(
    `(async () => { const d = await window.__otzariaEditor.superdoc.activeEditor.doc.get();
       const walk = (nodes) => { for (const n of nodes || []) {
         const id = (n.paragraphIds && n.paragraphIds.paraId) || n.id;
         if (id === ${JSON.stringify(blockId)}) return (n.paragraph && n.paragraph.props) || null;
         const inner = walk(n.body || n.children || (n.table && n.table.rows) || []);
         if (inner) return inner;
       } return null; };
       return JSON.stringify(walk(d.body) || null);
     })()`,
  );
  say(`\ndoc.get() props של אותה פסקה: ${props}`);
  const readable = typeof props === 'string' && props.includes('contextualSpacing');
  say(`>> קריאה חזרה: ${readable ? 'כן — הערך במודל' : 'לא — הערך אינו במודל, בדיוק כמו pageBreakBefore'}`);

  say(
    `\n>> מסקנה לתיבה בדיאלוג: ${
      readable && keptBreak
        ? 'אפשרית — נקראת ונכתבת בלי לפגוע בשכנתה'
        : !keptBreak && wroteContextual
          ? 'חסומה — הכתיבה מוחקת מעבר עמוד שהמשתמש הגדיר'
          : 'חסומה — אי אפשר לדעת מה מצבה, ותיבה שנפתחת ריקה ומאושרת מכבה אותה'
    }`,
  );
} catch (error) {
  say(`\nגשש נכשל: ${error.message}`);
  say(String(error.stack));
} finally {
  app.close();
}

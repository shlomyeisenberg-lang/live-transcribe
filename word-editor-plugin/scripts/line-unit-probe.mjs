/**
 * גשש: באילו יחידות `doc.get()` מחזיר את `spacing.line`, לכל `lineRule`.
 *
 * למה זה נשאל: ב-OOXML ל-`w:line` יש **שתי** משמעויות שונות לפי `w:lineRule`
 * — ב-`exact`/`atLeast` הוא מרחק (twips), וב-`auto` הוא כפולה ב-240ths. הקוד
 * כאן ממיר את מה שהמודל מחזיר ב-`× 20` אחיד („נקודות → twips”), והשער
 * `home-paragraph-qa` תפס פסקה שיצאה עם `w:line="5"` אחרי שתי כתיבות
 * שנועדו לשמר אותה — כלומר הערך התכווץ בכל סבב קריאה-כתיבה.
 *
 * הגשש כותב ערך ידוע, קורא את המודל, וכותב שוב את מה שנקרא — ומדפיס את
 * שלושת המספרים. פער ביניהם הוא הבאג.
 *
 *   npm run build && node scripts/line-unit-probe.mjs
 */
import { openApp, sleep } from './qa/harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9372);
const TARGET_TEXT = 'זוגי';

const app = await openApp({ name: 'line-unit', port: PORT });

try {
  await app.caret(0);
  await app.type('זוגי');
  await sleep(1500);
  await app.caret(0);
  await sleep(400);

  const blockId = await app.js(
    '(async () => { const s = await window.__otzariaEditor.superdoc.activeEditor.doc.selection.current(); ' +
      'return (s.target && s.target.segments && s.target.segments[0] && s.target.segments[0].blockId) || null; })()',
  );

  const target = `{ kind: 'block', nodeType: 'paragraph', nodeId: ${JSON.stringify(blockId)} }`;

  const setSpacing = (line, lineRule) =>
    app.js(
      `(async () => { const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
         try { const r = await doc.format.paragraph.setSpacing({ target: ${target}, before: 0, after: 0, line: ${line}, lineRule: '${lineRule}' });
           return JSON.stringify({ success: r && r.success, failure: r && r.failure }); }
         catch (e) { return JSON.stringify({ threw: String(e && e.message) }); } })()`,
    );

  const readSpacing = () =>
    app.js(
      `(async () => { const d = await window.__otzariaEditor.superdoc.activeEditor.doc.get();
         const walk = (nodes) => { for (const n of nodes || []) {
           const id = (n.paragraphIds && n.paragraphIds.paraId) || n.id;
           if (id === ${JSON.stringify(blockId)}) return ((n.paragraph && n.paragraph.props) || {}).spacing || null;
           const inner = walk(n.body || n.children || []);
           if (inner) return inner;
         } return null; };
         return JSON.stringify(walk(d.body) || null); })()`,
    );

  const xmlLine = async (label) => {
    const files = await app.docx();
    const xml = files['word/document.xml'] ?? '';
    const p = (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).find((x) => x.includes(`>${TARGET_TEXT}<`)) ?? '';
    const m = p.match(/<w:spacing[^>]*\/>/)?.[0] ?? '(אין w:spacing)';
    console.log(`  ${label}: ${m}`);
    return m;
  };

  for (const [line, rule] of [[480, 'auto'], [720, 'auto'], [360, 'exact'], [360, 'atLeast']]) {
    console.log(`\n== נכתב line=${line} lineRule=${rule} ==`);
    console.log(`  קבלה: ${await setSpacing(line, rule)}`);
    await sleep(900);
    await xmlLine('ב-XML');
    const model = await readSpacing();
    console.log(`  המודל מחזיר: ${model}`);
    const reported = JSON.parse(model || 'null')?.line;
    console.log(
      `  >> ${
        reported === line
          ? 'twips גולמיים — אותו מספר'
          : reported === line / 20
            ? 'נקודות (line/20) — ההמרה ×20 נכונה'
            : reported === line / 240
              ? 'כפולת שורות (line/240) — ההמרה ×20 שגויה כאן!'
              : `יחידה אחרת: ${reported} מול ${line}`
      }`,
    );
  }
} catch (error) {
  console.error(`גשש נכשל: ${error.message}`);
  console.error(error.stack);
} finally {
  app.close();
}

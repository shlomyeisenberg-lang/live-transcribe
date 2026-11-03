/**
 * גשש: מה Chrome מחזיר ל-`Range` שחותך **בתוך** אשכול גרפמה.
 *
 * זו המדידה שמחזיקה את האיטרציה ב-`readLineChars` (‏`src/engine/rtl-caret.ts`).
 * היא הייתה קודם ב-`scratchpad/surrogate.html`, שאינו במאגר ולכן לא ניתן היה
 * להריץ אותה שוב — כאן היא רצה משורת הפקודה כמו כל מדידה אחרת במאגר:
 *
 *   node scripts/qa/rtl-caret-cluster-probe.mjs        (CDP_PORT עוקף 9481)
 *
 * הגשש **אינו** פותח את התוסף: הוא שואל שאלה על הדפדפן בלבד — איך `Range`
 * מתנהג על טקסט מנוקד — ולכן דף ריק עם `<span dir="rtl">` הוא כל מה שנדרש.
 * יציאה 0 = הכלל שנמדד עדיין מתקיים.
 *
 * ## מה נבדק
 *
 * 1. חיתוך לפי **יחידת UTF-16** בתוך אשכול מחזיר את תיבת האשכול המלאה —
 *    כלומר כמה חריצים באותו x, שכולם מלבד הראשון הם היסט בתוך אות.
 * 2. חיתוך לפי **אשכול** מחזיר בדיוק את אותה תיבה, ופעם אחת.
 * 3. ‏`Intl.Segmenter` קיים בדפדפן, והגבולות שהוא נותן ל-`he` ולברירת המחדל
 *    זהים (כללי UAX-29 אינם תלויי שפה, וזה מה שמצדיק בנייה בלי locale).
 */
import { writeFileSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, requireChrome, onInterrupt } from '../cdp.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));

/**
 * השורות. כל אחת נבחרה מפני שנמדד עליה משהו שהקוד נשען עליו:
 * אות עם ניקוד **וטעם** באמצע שורה מעורבת, אות לטינית מפורקת, ניקוד בתחילת
 * שורה ואחרי ספרות, וזוג תחליף — המקרה שבגללו נכתבה האיטרציה מלכתחילה.
 */
const LINES = [
  'מילה abcשָׁלום',
  'מילה e\u0301a מילה',
  '12שָׁלום',
  'בְּ1948 שְנָה',
  'שלום 😀 עולם',
];

const PAGE = `<!doctype html><meta charset="utf-8"><body style="margin:0">
<div id="host" dir="rtl" style="font:16px serif;width:700px"></div>
<script>
const LINES = ${JSON.stringify(LINES)};
window.measure = () => {
  const host = document.getElementById('host');
  if (typeof Intl.Segmenter !== 'function') return { segmenter: false, rows: [] };
  const seg = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
  const segHe = new Intl.Segmenter('he', { granularity: 'grapheme' });
  const rows = [];
  for (const text of LINES) {
    host.textContent = text;
    const node = host.firstChild;
    const range = document.createRange();
    const box = (from, to) => {
      range.setStart(node, from);
      range.setEnd(node, to);
      const r = range.getBoundingClientRect();
      return [+r.left.toFixed(1), +r.right.toFixed(1)];
    };
    const clusters = [];
    for (const s of seg.segment(text)) {
      clusters.push({ at: s.index, len: s.segment.length, box: box(s.index, s.index + s.segment.length) });
    }
    const heStarts = [...segHe.segment(text)].map((s) => s.index);
    const units = [];
    for (let i = 0; i < text.length; i += 1) units.push({ at: i, box: box(i, i + 1) });
    rows.push({ text, clusters, units, heStarts });
  }
  return { segmenter: true, rows };
};
</script>`;

requireChrome();
const file = join(HERE, '__cluster-probe.html');
writeFileSync(file, PAGE);
const forget = onInterrupt(() => rmSync(file, { force: true }));

const { cdp, close } = await openPage(`file:///${file.split('\\').join('/')}`, {
  port: Number(process.env.CDP_PORT ?? 9481),
  label: 'cluster',
});

let failures = 0;
const fail = (line) => {
  failures += 1;
  console.log('  ** ' + line);
};

try {
  const data = JSON.parse(await cdp.evaluate('JSON.stringify(window.measure())'));
  if (!data.segmenter) {
    console.log('** אין `Intl.Segmenter` בדפדפן הזה — `readLineChars` נשען עליו');
    process.exit(1);
  }

  for (const row of data.rows) {
    console.log('\n=== ' + row.text);
    const starts = row.clusters.map((c) => c.at);
    if (starts.join(',') !== row.heStarts.join(',')) {
      fail(`גבולות he שונים מברירת המחדל: ${row.heStarts} מול ${starts}`);
    }
    for (const c of row.clusters) {
      const inside = row.units.filter((u) => u.at > c.at && u.at < c.at + c.len);
      const same = inside.filter((u) => u.box[0] === c.box[0] && u.box[1] === c.box[1]);
      const mark = c.len > 1 ? (same.length === inside.length ? 'כל היחידות = התיבה המלאה' : `** ${same.length}/${inside.length}`) : '';
      console.log(`  ${c.at} len=${c.len} [${c.box[0]}, ${c.box[1]}] ${mark}`);
      if (c.len > 1 && same.length !== inside.length) {
        fail(`אשכול ב-${c.at}: לא כל היחידות מחזירות את התיבה המלאה — האיטרציה בקוד נשענת על זה`);
      }
    }
  }

  console.log(
    failures
      ? `\n** ${failures} כשלים`
      : '\nהכלל מתקיים: חיתוך בתוך אשכול מחזיר את תיבתו המלאה, ולכן איטרציה לפי יחידה או לפי נקודת קוד הייתה מייצרת חריץ בתוך אות',
  );
} finally {
  close();
  forget();
  rmSync(file, { force: true });
}

process.exit(failures ? 1 : 0);

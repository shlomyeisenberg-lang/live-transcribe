/**
 * שער הסרגל: האם המספרים על שני הסרגלים **נראים**.
 *
 * למה שער ולא בדיקת יחידה: jsdom אינו מחשב פריסה, ולכן שום בדיקה בסוויטה אינה
 * יכולה לדעת היכן `bottom: 100%` על שנתה או `inset-inline-start` על מספר
 * מציבים את התווית בפועל. שני הסרגלים חתכו מספרים בדיוק שם, ו-2,000 בדיקות
 * ירוקות לא ראו את זה:
 *
 *   - האופקי: `bottom: 100%` על שנתה שראשה ב-4px דחף מספר בן 9px אל ‎-5px,
 *     ו-`overflow: hidden` של הרצועה חתך ממנו חמישה — 44% מהספרה נראו.
 *   - האנכי: המספר הוצמד אל תוך תיבת השנתה שרוחבה 6px, ולכן „10” יצא ל-‎-2.5px
 *     בזמן ש-14 פיקסלים מהעמודה עמדו ריקים.
 *
 * מה שנמדד: כל מספר, מול תיבת החיתוך של הסרגל שהוא יושב בו. `overflow: hidden`
 * הוא מה שחותך, ולכן ההשוואה היא מולו ולא מול המסך.
 *
 * ה-CSS **נקרא מהקומפוננטות עצמן** ולא משוכפל כאן. שער שמחזיק עותק של הסגנון
 * מאמת את העותק.
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { openPage, requireChrome } from './cdp.mjs';

const ROOT = process.cwd();

/** ה-`<style>` של קומפוננטה. `scoped` מתקמפל לתכונות, והסלקטורים כאן זהים. */
function componentStyle(file) {
  const source = readFileSync(join(ROOT, 'src/ui/shell', file), 'utf8');
  const match = source.match(/<style scoped>([\s\S]*?)<\/style>/);
  if (!match) throw new Error(`לא נמצא <style scoped> ב-${file}`);
  return match[1];
}

/** גובה הרצועה ורוחב העמודה נקראים מה-CSS — הם מה שהמספרים צריכים להיכנס אליו. */
const horizontal = componentStyle('DocumentRuler.vue');
const vertical = componentStyle('VerticalRuler.vue');

/**
 * טוקנים מינימליים. הצבעים אינם נמדדים; `--font-main` כן — רוחב הספרה הוא
 * חלק מהשאלה, ולכן הגופן הארוז נטען כאן באמת.
 */
const TOKENS = `
  @font-face {
    font-family: 'Assistant';
    src: url('file://${join(ROOT, 'public/fonts/Assistant-Regular.ttf')}') format('truetype');
  }
  :root {
    --font-main: 'Assistant', 'Rubik', Arial, sans-serif;
    --color-surface: #fff;
    --color-surface-container-high: #eee;
    --color-surface-container-highest: #e5e5e5;
    --color-outline-variant: #ccc;
    --color-on-surface-variant: #444;
    --color-primary: #06c;
    --color-on-primary: #fff;
    --word-blue: #2b579a;
    --radius-xs: 3px;
    --radius-pill: 999px;
  }
  body { margin: 0; direction: rtl; font-family: var(--font-main); }
`;

/**
 * A4 ב-100%: 794 × 1123 פיקסלים. התוויות בסנטימטרים מגיעות ל-19 ול-27.
 *
 * העמודה האנכית מקבלת את גובה העמוד המלא ולא גובה חלון: תווית שנמצאת מתחת
 * לתחתית החלון „נחתכת” בגלילה רגילה, וזה לא מה שנמדד כאן. כך כל התוויות
 * בתוך תיבת החיתוך, וכל חיתוך שנשאר הוא של הציר שהבאג ישב בו.
 */
const PAGE_W = 794;
const PAGE_H = 1123;

const hTicks = Array.from({ length: 20 }, (_, i) => {
  const left = Math.round((i + 1) * (PAGE_W / 21));
  return `<div class="doc-ruler__tick doc-ruler__tick--major" style="left:${left}px"><span class="doc-ruler__number">${i + 1}</span></div>`;
}).join('');

const vTicks = Array.from({ length: 27 }, (_, i) => {
  const top = Math.round((i + 1) * (PAGE_H / 28));
  return `<div class="doc-vruler__tick doc-vruler__tick--major" style="top:${top}px"><span class="doc-vruler__number">${i + 1}</span></div>`;
}).join('');

const html = `<!doctype html>
<html dir="rtl"><head><meta charset="utf-8">
<style>${TOKENS}${horizontal}${vertical}</style></head>
<body>
  <div class="doc-ruler" style="width:900px">
    <div class="doc-ruler__page" style="left:50px;width:${PAGE_W}px">
      <div class="doc-ruler__text-area" style="left:96px;width:602px"></div>
      ${hTicks}
      <div class="doc-ruler__handle doc-ruler__handle--indent-start" style="left:96px"></div>
    </div>
  </div>
  <div class="doc-vruler" style="height:${PAGE_H}px">
    <div class="doc-vruler__page" style="top:0;height:${PAGE_H}px">
      ${vTicks}
    </div>
  </div>
</body></html>`;

const path = join(tmpdir(), 'otzaria-word-ruler-check.html');
writeFileSync(path, html);

requireChrome();
const page = await openPage(`file://${path}`, { label: 'ruler' });

const report = await page.cdp.evaluate(`(() => {
  const measure = (rulerSel, numberSel) => {
    const clip = document.querySelector(rulerSel).getBoundingClientRect();
    return [...document.querySelectorAll(numberSel)].map((el) => {
      const b = el.getBoundingClientRect();
      const visible =
        Math.max(0, Math.min(b.right, clip.right) - Math.max(b.left, clip.left)) *
        Math.max(0, Math.min(b.bottom, clip.bottom) - Math.max(b.top, clip.top));
      return { text: el.textContent, ratio: b.width * b.height ? visible / (b.width * b.height) : 0 };
    });
  };
  return JSON.stringify({
    horizontal: measure('.doc-ruler', '.doc-ruler__number'),
    vertical: measure('.doc-vruler', '.doc-vruler__number'),
  });
})()`);

page.close();

const { horizontal: h, vertical: v } = JSON.parse(report);
const cut = [
  ...h.map((n) => ({ ...n, ruler: 'אופקי' })),
  ...v.map((n) => ({ ...n, ruler: 'אנכי' })),
].filter((n) => n.ratio < 0.999);

const worst = (list) => Math.min(...list.map((n) => n.ratio));
console.log(
  `אופקי: ${h.length} מספרים, הגרוע ${(worst(h) * 100).toFixed(1)}% נראה — ` +
    `אנכי: ${v.length} מספרים, הגרוע ${(worst(v) * 100).toFixed(1)}%`,
);

if (cut.length) {
  for (const n of cut) {
    console.error(`  ✗ הסרגל ה${n.ruler}: „${n.text}” נחתך — ${(n.ratio * 100).toFixed(1)}% נראה`);
  }
  console.error(`שער הסרגל נכשל: ${cut.length} מספרים נחתכים ב-overflow של הרצועה.`);
  process.exit(1);
}

console.log('שער הסרגל עבר: כל המספרים בשני הסרגלים נראים במלואם.');

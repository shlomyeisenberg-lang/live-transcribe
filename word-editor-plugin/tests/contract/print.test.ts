/**
 * ההנחות שגלון ההדפסה נשען עליהן, מקובעות מול החבילה.
 *
 * הגלון מכוון אל **המחלקות של המנוע** (`superdoc-page`, `superdoc-layout`) —
 * וזו הטענה שהכי קל שתתיישן בשקט: שינוי שם מחלקה בגרסת superdoc הבאה משאיר
 * את ה-CSS תקין לחלוטין ואת ההדפסה שבורה, בלי שאף בדיקה תיפול ובלי שאיש
 * ישים לב. הבדיקה הזאת נופלת כשזה קורה.
 *
 * כאן גם מקובעות שתי המדידות שהמימוש נשען עליהן ואינן נראות בקוד שלנו:
 * שהמנוע מגדיר `@media print` משלו (ולכן אין צורך לחזור על מה שהוא כבר עושה),
 * ושהפרויקציה הציבורית של גודל הדף היא twips/1440 — כלומר אינצ'ים.
 *
 * התבנית — קריאת ה-bundle של המנוע והחיפוש בסביבת המופע — היא זו של
 * tests/contract/page-break.test.ts.
 */
import { describe, it, expect } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { ENGINE_LAYOUT_CLASS, ENGINE_PAGE_CLASS } from '../../src/engine/print';

const ENGINE_DIR = join(process.cwd(), 'node_modules/@superdoc/docx-engine/dist');
const ENGINE_BUNDLES = ['document-runtime.js', 'docx-engine.es.js'] as const;
const cache = new Map<string, string>();

function engineSource(name: string): string | null {
  if (cache.has(name)) return cache.get(name)!;
  const path = join(ENGINE_DIR, name);
  if (!existsSync(path)) return null;
  const source = readFileSync(path, 'utf8');
  cache.set(name, source);
  return source;
}

function ensureBundles(): void {
  if (!ENGINE_BUNDLES.some((name) => engineSource(name) !== null)) {
    throw new Error(`לא נמצא bundle של מנוע ה-DOCX ב-${ENGINE_DIR}`);
  }
}

/**
 * מחזירה boolean ולא את ה-bundle: `expect(bundle).toContain(x)` על מחרוזת של
 * חמישה מגה-בייט מדפיס את כל הקובץ בכשל. אותו שיקול כמו ב-page-break.test.ts.
 */
function engineHas(needle: string): boolean {
  ensureBundles();
  return ENGINE_BUNDLES.some((name) => engineSource(name)?.includes(needle) === true);
}

/** האם `near` מופיע בסביבת אחד המופעים של `needle`. */
function engineHasNear(needle: string, near: string, span: number): boolean {
  ensureBundles();
  for (const name of ENGINE_BUNDLES) {
    const source = engineSource(name);
    if (!source) continue;
    for (let at = source.indexOf(needle); at !== -1; at = source.indexOf(needle, at + 1)) {
      if (source.slice(Math.max(0, at - span), at + span).includes(near)) return true;
    }
  }
  return false;
}

describe('המחלקות שהגלון מכוון אליהן', () => {
  it('`superdoc-page` היא תיבת העמוד שהמנוע מצייר', () => {
    expect(ENGINE_PAGE_CLASS).toBe('superdoc-page');
    // בטבלת השמות של השכבה שמציירת (`'PAGE': 'superdoc-page'`).
    expect(engineHas(`'${ENGINE_PAGE_CLASS}'`)).toBe(true);
    expect(engineHasNear(`'${ENGINE_PAGE_CLASS}'`, "'PAGE'", 200)).toBe(true);
  });

  it('`superdoc-layout` הוא מיכל העמודים', () => {
    expect(ENGINE_LAYOUT_CLASS).toBe('superdoc-layout');
    expect(engineHasNear(`'${ENGINE_LAYOUT_CLASS}'`, "'container'", 200)).toBe(true);
  });
});

describe('מה שהמנוע כבר עושה בהדפסה — ומה שהוא משאיר לנו', () => {
  it('למנוע יש `@media print` משלו על תיבת העמוד', () => {
    // ולכן הגלון שלנו אינו חוזר על המרווח, הגבול ומעבר העמוד: הם שם. אם זה
    // ייעלם, `page-break-after` יצטרך לעבור אלינו — ולכן זה נבדק.
    // ה-bundle ממוזער עם escape לרווחים (`@media\\x20print`), ולכן זה הנוסח
    // שמחפשים — חיפוש על „@media print” אינו נמצא בו בכלל.
    const MEDIA_PRINT = '@media\\x20print';
    expect(engineHas(MEDIA_PRINT)).toBe(true);
    // מ-2.23 הפתיח `@media print {` עובר לטבלת המחרוזות של ה-bundle, רחוק
    // מגוף הכלל, ולכן מעגנים על הגוף עצמו: הצל מבוטל ומעבר עמוד אחרי כל
    // תיבה. השילוב הזה קיים רק בכלל ההדפסה של העמוד.
    const ESC = '\\x0a\\x20\\x20\\x20\\x20';
    expect(engineHas(`box-shadow:\\x20none;${ESC}page-break-after:\\x20always`)).toBe(true);
  });

  it('הצל של העמוד הוא סגנון inline — ולכן דורש `!important`', () => {
    // `boxShadow` נכתב על ה-style של האלמנט, וההצהרה של המנוע במדיית print
    // (`box-shadow: none`) אינה `!important` ולכן אינה מנצחת אותו. נמדד בפלט.
    expect(engineHasNear("'boxShadow'", 'sd-layout-page-shadow', 300)).toBe(true);
  });

  it('הזום מיושם כ-transform, ולכן ההדפסה חייבת לבטל אותו', () => {
    // נמדד ב-CDP: `matrix(0.5, 0, 0, 0.5, 0, 0)` על `.superdoc-layout` בזום
    // 50%. בלי הביטול נדפס מסמך מוקטן בפינת הגיליון.
    expect(engineHas('scale(')).toBe(true);
  });
});

describe('היחידות של גודל הדף', () => {
  it('הפרויקציה הציבורית היא twips/1440 — כלומר אינצ׳ים', () => {
    // 0x5a0 = 1440. הסטרים מכפילים באותו מספר (ראו engine/page-setup.ts),
    // וההיפוך הזה הוא מה שהופך את `pageSetup.width` לאינצ'ים.
    expect(engineHas('/0x5a0')).toBe(true);
    expect(engineHasNear('/0x5a0', 'isFinite', 160)).toBe(true);
  });

  it('`pageSetup` הציבורי נבנה מ-widthTwips/heightTwips', () => {
    expect(engineHasNear("'paperSize'", 'widthTwips', 400)).toBe(true);
  });
});

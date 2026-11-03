/**
 * ההכרעה על „מעבר עמוד”, מקובעת מול החבילה.
 *
 * ## מה ההכרעה
 *
 * מה ש-Word קורא לו Insert ▸ Page Break הוא `<w:br w:type="page"/>` בסמן.
 * **אין לזה API ציבורי ב-2.8.0**, ולכן הפקד מממש `w:pageBreakBefore` על הפסקה
 * — התכונה ש-Word מציג כ„מעבר עמוד לפני”. ההבדל אמיתי (הפסקה כולה עוברת, ולא
 * מתפצלת בסמן), ולכן התווית היא „התחל בעמוד חדש”.
 *
 * ## למה זה נבדק ולא רק מתועד
 *
 * ההכרעה נשענת על **היעדר** של משהו, וזו הטענה שהכי קל שתתיישן: ברגע
 * ש-superdoc יוסיף פקודת page-break או פעולת Document API כזאת, ההערה בקוד
 * תישאר נכונה-לכאורה והפקד יישאר על המסלול הפחות טוב. הבדיקה הזאת נופלת
 * כשזה קורה, וזה בדיוק הרגע שבו כדאי לשנות את המימוש.
 *
 * §12 בתכנית: „פקד שאין לו API ציבורי אמין מסומן „לא זמין בגרסה זו”; לא
 * מממשים אותו דרך XML ידני או DOM פנימי.” כאן **יש** API ציבורי אמין — הוא
 * פשוט עושה משהו מדויק יותר ממה שהתווית „מעבר עמוד” הבטיחה.
 */
import { describe, it, expect } from 'vitest';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { PAGE_BREAK_OPERATION } from '../../src/engine/page-break';

const PKG = join(process.cwd(), 'node_modules/superdoc/dist');
const API = join(PKG, 'document-api/src');
const CHUNKS_DIR = join(PKG, 'chunks');
const ENGINE_DIR = join(process.cwd(), 'node_modules/@superdoc/docx-engine/dist');

function read(path: string): string {
  return readFileSync(path, 'utf8');
}

/** שם ה-chunk נושא hash שמשתנה בכל build, ולכן הוא נמצא ולא נכתב. */
function readControllerChunk(): string {
  const file = readdirSync(CHUNKS_DIR).find((name) =>
    /^create-super-doc-ui-.*\.es\.js$/.test(name),
  );
  if (!file) throw new Error('לא נמצא ה-chunk של controller ה-UI ב-superdoc');
  return read(join(CHUNKS_DIR, file));
}

const CHUNK = readControllerChunk();

/**
 * ה-bundles של המנוע. הקטן ראשון, וה-`document-runtime` אינו מכיל את הכול —
 * `insertPageBreakAtSelection` יושב רק ב-`docx-engine.es.js`. הקריאה עצלה
 * ומוקאשת: מדובר בעשרות מגה-בייטים, ואין סיבה לקרוא את הגדול אם הקטן הספיק.
 */
const ENGINE_BUNDLES = ['document-runtime.js', 'docx-engine.es.js'] as const;
const engineCache = new Map<string, string>();

function engineSource(name: string): string | null {
  if (engineCache.has(name)) return engineCache.get(name)!;
  const path = join(ENGINE_DIR, name);
  if (!existsSync(path)) return null;
  const source = read(path);
  engineCache.set(name, source);
  return source;
}

function ensureBundles(): void {
  if (!ENGINE_BUNDLES.some((name) => engineSource(name) !== null)) {
    throw new Error(`לא נמצא bundle של מנוע ה-DOCX ב-${ENGINE_DIR}`);
  }
}

/**
 * האם המנוע מכיל את המחרוזת.
 *
 * מחזירה boolean ולא את ה-bundle: `expect(bundle).toContain(x)` על מחרוזת של
 * חמישה מגה-בייט מדפיס את כל הקובץ בכשל, וזה הופך כשל אמיתי לפלט שאי אפשר
 * לקרוא. נמדד — זה קרה בכתיבת הבדיקה הזאת.
 */
function engineHas(needle: string): boolean {
  ensureBundles();
  return ENGINE_BUNDLES.some((name) => engineSource(name)?.includes(needle) === true);
}

/**
 * האם `near` מופיע בסביבת **אחד** המופעים של `needle`.
 *
 * כל המופעים ולא הראשון: המזהה של פעולה מופיע ב-bundle גם בטבלת המטא-דאטה
 * (תיאור, referenceDocPath), גם במפת הניתוב וגם בוולידטור. המופע הראשון הוא
 * המטא-דאטה, ובדיקה עליו לבדו נכשלת אף שהטענה נכונה — נמדד.
 */
function engineHasNear(needle: string, near: string, span: number): boolean {
  ensureBundles();
  for (const name of ENGINE_BUNDLES) {
    const source = engineSource(name);
    if (!source) continue;
    for (let index = source.indexOf(needle); index !== -1; index = source.indexOf(needle, index + 1)) {
      if (source.slice(index, index + span).includes(near)) return true;
    }
  }
  return false;
}

/** כל מזהי הפקודות ב-`COMMAND_CATALOG` של ה-controller. */
const COMMAND_IDS = [...CHUNK.matchAll(/^\t\tid: "([^"]+)",$/gm)].map((match) => match[1]);

describe('המסלול שנבחר קיים', () => {
  it('נמצאו מזהי פקודות ב-catalog — אחרת הבדיקות למטה חסרות משמעות', () => {
    expect(COMMAND_IDS.length).toBeGreaterThan(20);
  });

  it(`${PAGE_BREAK_OPERATION} הוא OperationId אמיתי`, () => {
    // זה מה שהופך את מפתח היכולת ל-`operations[...]` לתקף. מזהה שאינו בקטלוג
    // היה מחזיר `undefined`, כלומר פקד שמנוטרל לנצח בלי שאיש יבין למה.
    expect(read(join(API, 'contract/command-catalog.d.ts'))).toContain(
      `"${PAGE_BREAK_OPERATION}"`,
    );
  });

  it('`pageBreakBefore` הוא שדה בקלט של הפעולה', () => {
    const types = read(join(API, 'paragraphs/paragraphs.types.d.ts'));
    const input = types.slice(types.indexOf('ParagraphsSetFlowOptionsInput'));

    expect(input.slice(0, input.indexOf('}'))).toContain('pageBreakBefore?: boolean');
  });

  it('הוולידטור של המנוע מקבל `pageBreakBefore` בפעולה הזאת', () => {
    // ההצהרה בטיפוסים אינה מבטיחה שהמימוש קורא את השדה: שלושת הדגלים
    // המזרח-אסייתיים מוצהרים שם ואינם מגיעים ל-XML.
    expect(engineHas(`'${PAGE_BREAK_OPERATION}'`), 'הפעולה לא נמצאה ב-bundle').toBe(true);
    expect(engineHasNear(`'${PAGE_BREAK_OPERATION}'`, "'pageBreakBefore'", 400)).toBe(true);
  });
});

describe('המסלולים שנדחו — והסיבה שהם נדחו', () => {
  it('אין פקודת מעבר עמוד ב-COMMAND_CATALOG של ה-controller', () => {
    // אם זה נשבר, יש עכשיו מסלול טוב יותר והפקד צריך לעבור אליו.
    expect(COMMAND_IDS.filter((id) => /break/i.test(id))).toEqual([]);
  });

  it('אין פעולת Document API שמכניסה מעבר עמוד בסמן', () => {
    const catalog = read(join(API, 'contract/command-catalog.d.ts'));
    const pageBreakOps = [...catalog.matchAll(/"([\w.]*[Pp]ageBreak[\w.]*)"/g)].map((m) => m[1]);

    // `format.paragraph.setFlowOptions` אינו מכיל „pageBreak” בשם שלו, ולכן
    // רשימה ריקה כאן פירושה: אין פעולה ייעודית למעבר עמוד.
    expect(pageBreakOps).toEqual([]);
    expect(catalog).toContain(`"${PAGE_BREAK_OPERATION}"`);
  });

  it('`create.sectionBreak` הוא מעבר מקטע, ולכן נדחה', () => {
    // החוזה עצמו מתעד שהוא פולט פסקת-נשא לפני המקטע האחרון בגוף — כלומר לא
    // בסמן — ומקטע חדש מנתק כותרות והגדרות עמוד. שינוי מבני שלא התבקש.
    const sections = read(join(API, 'sections/sections.types.d.ts'));

    expect(sections).toContain('CreateSectionBreakInput');
    expect(sections).toContain('sectPr');
    expect(sections).toContain('just before the existing final body section');
  });

  it('הצינור הפנימי של המנוע כן יודע מעבר עמוד — ואינו נחשף', () => {
    // `insertPageBreakAtSelection` קשור ל-`Mod-Enter` בטיפול הקלט של המנוע.
    // זה מסביר את ההערה שהייתה בקוד על Ctrl+Enter, אבל הוא אינו על `doc`,
    // אינו ב-`ui.commands` ואינו ב-`OperationId` — ולכן אין ממנו מסלול לפקד
    // ברצועה, והקיצור אינו מוצג על הכפתור (לא אומת מקצה לקצה).
    expect(engineHas('insertPageBreakAtSelection')).toBe(true);
    expect(engineHasNear('insertPageBreakAtSelection', 'Mod-Enter', 600)).toBe(true);

    expect(CHUNK.includes('insertPageBreakAtSelection')).toBe(false);
    expect(
      read(join(API, 'contract/command-catalog.d.ts')).includes('insertPageBreak'),
    ).toBe(false);
  });
});

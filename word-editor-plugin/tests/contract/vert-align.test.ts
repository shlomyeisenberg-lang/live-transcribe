/**
 * ההכרעה על „כתב עליון” ו„כתב תחתי”, מקובעת מול החבילה.
 *
 * ## מה ההכרעה
 *
 * ה-tooltip אמר „אינו נתמך במנוע הנוכחי” וה-`disabled` היה קשיח. **המנוע כן
 * תומך**: `vertAlign` הוא מפתח ב-`InlineRunPatch`, `format.vertAlign` הוא
 * `OperationId` בקטלוג, ו-`FormatApi` חושף alias ציבורי לכל מפתח כזה. מה שחסר
 * היה פקודה ב-`COMMAND_CATALOG` של ה-controller — ולכן המסלול הוא ה-Document
 * API, בדיוק כמו הערות שוליים ומעבר עמוד.
 *
 * ## למה זה נבדק ולא רק מתועד
 *
 * שתי טענות הפוכות נשענות כאן על החבילה, ושתיהן יכולות להתיישן בשקט:
 *
 *   1. **יש** API לכתיבה. אם הוא ייעלם, הפקד יציג „אינו זמין בגרסה זו”
 *      במקום לעבוד, ואף אחד לא יבין למה — הבדיקה הזאת נופלת קודם.
 *   2. **אין** קריאה של המצב באף אחד משלושת המשטחים שהרצועה קוראת מהם, ולכן
 *      אין חיווי „דלוק” על הכפתור (ההסבר המלא ב-engine/vert-align.ts). ברגע
 *      שיתווסף אחד — פקודה ב-registry, `vertAlign` ב-`MatchStyle` — כדאי
 *      לעבור אליו, וזה הרגע שבו הבדיקה הזאת נופלת.
 *
 * התבנית היא זו של tests/contract/page-break.test.ts.
 */
import { describe, it, expect } from 'vitest';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { VERT_ALIGN_OPERATION } from '../../src/engine/vert-align';

const API = join(process.cwd(), 'node_modules/superdoc/dist/document-api/src');
const CHUNKS_DIR = join(process.cwd(), 'node_modules/superdoc/dist/chunks');
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

/** כל מזהי הפקודות ב-`COMMAND_CATALOG` של ה-controller. */
const COMMAND_IDS = [...CHUNK.matchAll(/^\t\tid: "([^"]+)",$/gm)].map((match) => match[1]);

function engineHas(needle: string): boolean {
  for (const name of ['document-runtime.js', 'docx-engine.es.js']) {
    const path = join(ENGINE_DIR, name);
    if (existsSync(path) && read(path).includes(needle)) return true;
  }
  return false;
}

describe('המסלול שנבחר קיים', () => {
  it('נמצאו מזהי פקודות ב-catalog — אחרת הבדיקות למטה חסרות משמעות', () => {
    expect(COMMAND_IDS.length).toBeGreaterThan(20);
  });

  it('`vertAlign` הוא מפתח ב-InlineRunPatch, עם שלושת הערכים', () => {
    const patch = read(join(API, 'format/inline-run-patch.d.ts'));

    expect(patch).toContain(
      "vertAlign?: 'superscript' | 'subscript' | 'baseline' | null;",
    );
  });

  it(`${VERT_ALIGN_OPERATION} הוא OperationId אמיתי`, () => {
    // זה מה שהופך את מפתח היכולת `canSetVertAlign` ל-`operations[...]` תקף.
    // מזהה שאינו בקטלוג היה מחזיר `undefined`, כלומר פקד מנוטרל לנצח.
    expect(read(join(API, 'contract/command-catalog.d.ts'))).toContain(
      `"${VERT_ALIGN_OPERATION}"`,
    );
    expect(read(join(API, 'contract/operation-definitions.d.ts'))).toContain(
      `"${VERT_ALIGN_OPERATION}"`,
    );
  });

  it('`doc.format.vertAlign` הוא alias ציבורי ולא פנימי', () => {
    // `FormatApi extends FormatInlineAliasApi`, ו-`FormatInlineAliasApi` הוא
    // mapped type על **כל** מפתח ב-InlineRunPatch. כלומר המתודה קיימת בחוזה
    // בלי שהיא כתובה בו בשמה.
    const format = read(join(API, 'format/format.d.ts'));

    expect(format).toContain('interface FormatApi extends FormatInlineAliasApi');
    expect(format).toContain('[K in InlineRunPatchKey]');
  });

  it('הרַנטַיים v2 — זה שרץ בדפדפן — מנתב את הפעולה', () => {
    // ההצהרה בטיפוסים אינה מבטיחה מימוש ברַנטַיים שאנחנו רצים בו. הפעולה
    // מופיעה ברשימת הפעולות שה-v2 מנתב לפאסדה.
    // מ-docx-engine 0.16 המזהה נבנה כ-`'format.'+key`, ולכן נמצא רק רישום המפתח.
    const routed = engineHas(`'${VERT_ALIGN_OPERATION}'`) || engineHas("'vertAlign','string','w:vertAlign'");
    expect(routed).toBe(true);
  });

  it('`vertAlign` נשמר כתכונה של mark מסוג textStyle', () => {
    // זה מה שמסביר למה `activeMarks` אינו יכול לדווח את הערך: הוא מדווח שמות
    // marks, וכאן שם ה-mark הוא `textStyle` והכתב הוא תכונה שלו.
    expect(CHUNK).toContain('markTextStyleValue("vertAlign", "string", "w:vertAlign"');
  });
});

describe('המסלולים שאינם קיימים — והם הסיבה לצורה של הפקד', () => {
  it('אין פקודת superscript/subscript ב-COMMAND_CATALOG של ה-controller', () => {
    // אם זה נשבר, יש עכשיו פקודה ברצועה והפקד צריך לעבור אליה: `useCommand`
    // נותן גם `active` וגם `enabled` מהמנוע, כלומר גם חיווי „דלוק”.
    expect(COMMAND_IDS.filter((id) => /script|vert/i.test(id))).toEqual([]);
  });

  it('`MatchStyle` של `query.match` אינו מדווח vertAlign', () => {
    // המשטח היחיד שכן מדווח ערכי עיצוב לריצה. אין בו את הכתב, ולכן אין דרך
    // ציבורית זולה לדעת אם הבחירה כבר בכתב עליון — ומכאן היעדר החיווי.
    const match = read(join(API, 'types/query-match.types.d.ts'));
    const style = match.slice(match.indexOf('interface MatchStyle'));

    expect(style.slice(0, style.indexOf('}\n'))).not.toContain('vertAlign');
  });

  it('`SelectionInfo.activeMarks` הוא שמות marks בלבד', () => {
    const selection = read(join(API, 'selection/selection.types.d.ts'));

    expect(selection).toContain('activeMarks: string[]');
    expect(selection).toContain('ProseMirror mark type names');
  });

  it('הקריאה שכן אפשרית — `SDRunProps.verticalAlign` — קיימת, ונדחתה', () => {
    // היא דורשת למפות היסט סמן לריצה בתוך הבלוק, כלומר לשחזר בקוד שלנו את מה
    // שהמנוע כבר עושה. אם `MatchStyle` יקבל את השדה, הבדיקה שמעל תיפול וזה
    // יהיה הרגע לעבור.
    expect(read(join(API, 'types/sd-props.d.ts'))).toContain(
      "verticalAlign?: 'baseline' | 'superscript' | 'subscript';",
    );
  });
});

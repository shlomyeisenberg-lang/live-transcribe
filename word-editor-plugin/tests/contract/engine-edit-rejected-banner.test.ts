/**
 * חוזה באנר „edit-rejected" של המנוע — שהוא לא חוזר בשקט.
 *
 * superdoc צייר הודעת מצב אנגלית מעל בד המסמך גם תחת `ui: false` (issue
 * #3957 במעלה הזרם; נמדד ב-2.11.0 וב-2.12.0), ו-engine-chrome.css הסתיר אותה
 * בכלל שנשען על מבנה DOM שאינו חוזה מתועד: תכונה ומחלקת עוטף. ב-2.15.0-next.15
 * הבאנר ירד מהאריזה, והדחייה מגיעה רק כ-`exception` עם `code: 'edit-rejected'`
 * — ש-create-editor.ts רושם ללוג. לכן הכלל נמחק, והבדיקה הזאת הפכה כיוון:
 * במקום לוודא שהעוגנים קיימים, היא מוודאת שהם לא חזרו.
 *
 * בלי הבדיקה, שדרוג מנוע שמחזיר את הבאנר היה מחזיר אנגלית למסך בשקט.
 * מה שנמדד:
 *
 *   1. התכונה `data-superdoc-v2-edit-rejected` ומחלקת העוטף
 *      `superdoc__mutation-status` אינן באריזה. אם אחת מהן חוזרת — צריך
 *      להחזיר את ההסתרה (בלי `:has()`, ראו engine-chrome.css).
 *   2. הדחייה עדיין נמסרת כ-`exception` עם הקוד הזה — זה האות היחיד שנשאר
 *      לנו על מחיקה שנדחתה.
 *   3. בגיליון לא נשאר כלל שמסתיר DOM שאינו קיים.
 *
 * הקריאה היא של מחרוזות מהאריזה, לצורך התאמת ממשק בלבד.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

/** הבאנר היה של חבילת `superdoc`, לא של אריזת מנוע ה-DOCX. */
const SHELL = join(process.cwd(), 'node_modules/superdoc/dist/superdoc.es.js');
const STYLE_SHEET = join(process.cwd(), 'src/styles/engine-chrome.css');

const shell = readFileSync(SHELL, 'utf8');
const sheet = readFileSync(STYLE_SHEET, 'utf8');

/** העוגן שהכלל תלה עליו את עצמו. */
export const EDIT_REJECTED_HOOK = 'data-superdoc-v2-edit-rejected';
/** העוטף שנשא את ה-`position: sticky`. */
export const MUTATION_STATUS_CLASS = 'superdoc__mutation-status';

describe('חוזה באנר edit-rejected', () => {
  it('הבאנר אינו באריזת superdoc — ולכן אין מה להסתיר', () => {
    expect(shell).not.toContain(EDIT_REJECTED_HOOK);
    expect(shell).not.toContain(MUTATION_STATUS_CLASS);
  });

  it('הדחייה עדיין נמסרת כ-exception עם הקוד edit-rejected', () => {
    expect(shell).toContain('"edit-rejected"');
    expect(shell).toMatch(/emit\("exception", createV2KeyboardEditRejectionException\(/);
  });

  it('בגיליון לא נשאר כלל שמסתיר את הבאנר', () => {
    // בלי הערות: ההערה שמסבירה למה הכלל נמחק מזכירה את שני העוגנים.
    const rules = sheet.replace(/\/\*[\s\S]*?\*\//g, '');
    expect(rules).not.toContain(EDIT_REJECTED_HOOK);
    expect(rules).not.toContain(MUTATION_STATUS_CLASS);
  });
});

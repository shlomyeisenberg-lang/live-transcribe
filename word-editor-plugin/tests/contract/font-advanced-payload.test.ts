/**
 * חוזה שמות המפתחות ש"גופן מתקדם" שולח ל-`format.apply`, מול ה-`.d.ts`
 * **האמיתי** של המנוע (`InlineRunPatch`).
 *
 * ## למה זה נבדק כך
 *
 * הפאץ' שהדיאלוג בונה (`FontAdvancedPatch`) משתמש בשמות שדה נוחים לממשק —
 * `boldCs`, `italicCs`, `complexScript` — שאינם שמות המפתחות שהמנוע מכיר:
 * המנוע מכיר `bCs`, `iCs`, `cs` (נמדד; ראו הערת הפתיחה של
 * `src/engine/font-advanced.ts`). `buildInlinePatch` העתיק בעבר את שמות
 * ה-patch כמות שהם, ושלח את שלושת המפתחות הלא-מוכרים למנוע.
 *
 * `format.apply` היא **קריאה אחת** על patch מרובה-מפתחות: מפתח לא-מוכר אחד
 * מחזיר `INVALID_INPUT: Unknown inline property` ומפיל את כל הקריאה —
 * כולל שדות תקינים לגמרי שנשלחו יחד איתו (ריווח תווים, מיקום, אפקטים,
 * שם גופן). בדיקת יחידה עם doc מדומה לא הייתה תופסת את זה, כי היא לא בהכרח
 * משווה מול רשימת המפתחות **האמיתית** של המנוע — ולכן הבדיקה כאן קוראת את
 * ה-`.d.ts` הארוז בפועל, לא מניחה אותו.
 *
 * מה שלא נבדק כאן: ההמרה ליחידות OOXML (twips/חצאי-נקודות) והכתיבה בפועל —
 * אלה נבדקים באימות בדפדפן (`scripts/qa/home-font-qa.mjs`) וב-
 * `tests/unit/font-advanced.test.ts`.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildInlinePatch, type FontAdvancedPatch } from '../../src/engine/font-advanced';

const INLINE_RUN_PATCH_DTS = join(
  process.cwd(),
  'node_modules/superdoc/dist/document-api/src/format/inline-run-patch.d.ts',
);

function read(path: string): string {
  return readFileSync(path, 'utf8');
}

/** כל שמות המפתחות של `InlineRunPatch`, נחלצים מגוף ה-interface עצמו. */
function engineInlineKeys(): Set<string> {
  const text = read(INLINE_RUN_PATCH_DTS);
  const start = text.indexOf('export interface InlineRunPatch {');
  const body = text.slice(start, text.indexOf('}\n', start));
  return new Set([...body.matchAll(/^\s{4}(\w+)\??:/gm)].map((m) => m[1]));
}

const FULL_PATCH: FontAdvancedPatch = {
  charScale: 150,
  letterSpacingPt: -20,
  positionPt: 10,
  kerningPt: 12,
  dstrike: true,
  outline: true,
  shadow: true,
  emboss: true,
  imprint: true,
  vanish: false,
  fontSizeCsPt: 12.5,
  boldCs: true,
  italicCs: true,
  complexScript: true,
  rtl: true,
  complexFontName: 'David',
  proofingLangBidi: 'he-IL',
};

describe('שמות המפתחות של buildInlinePatch מול InlineRunPatch האמיתי', () => {
  it('`bCs`/`iCs`/`cs` הם מפתחות אמיתיים ב-InlineRunPatch (עוגן — אחרת הבדיקה חסרת משמעות)', () => {
    const keys = engineInlineKeys();
    expect(keys.has('bCs')).toBe(true);
    expect(keys.has('iCs')).toBe(true);
    expect(keys.has('cs')).toBe(true);
    // עוגן שלילי: ודאי ששלושת אלה *אינם* מפתחות של המנוע — כלומר לו נשלחו
    // כלשונם היו נדחים.
    expect(keys.has('boldCs')).toBe(false);
    expect(keys.has('italicCs')).toBe(false);
    expect(keys.has('complexScript')).toBe(false);
  });

  it('כל מפתח top-level שנבנה קיים ב-InlineRunPatch — אחרת format.apply דוחה את כל ה-patch', () => {
    const engineKeys = engineInlineKeys();
    const built = buildInlinePatch(FULL_PATCH);

    expect('error' in built).toBe(false);
    if ('error' in built) return;

    for (const key of Object.keys(built.inline)) {
      expect(engineKeys.has(key), `המפתח "${key}" אינו קיים ב-InlineRunPatch של המנוע`).toBe(true);
    }
  });

  it('boldCs/italicCs/complexScript מתורגמים ל-bCs/iCs/cs — לא נשלחים כלשונם', () => {
    const built = buildInlinePatch({ boldCs: true, italicCs: false, complexScript: true });

    expect('error' in built).toBe(false);
    if ('error' in built) return;

    expect(built.inline).toEqual({ bCs: true, iCs: false, cs: true });
    expect(built.inline).not.toHaveProperty('boldCs');
    expect(built.inline).not.toHaveProperty('italicCs');
    expect(built.inline).not.toHaveProperty('complexScript');
  });
});

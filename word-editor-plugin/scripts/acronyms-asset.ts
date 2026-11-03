/**
 * בניית נכס ראשי-התיבות — צד Node בלבד (vite.config.ts והבדיקות). ההיגיון
 * שצורך אותו ב-src/engine/acronyms.ts.
 *
 * קובץ המקור הוא מיזוג של מילון אוצריא הרשמי ושל KleiKodeshProject, ושתי
 * הפעולות כאן נובעות מכך שהמיזוג אינו הומוגני:
 *
 * ## סינון: לא כל מפתח הוא ראשי תיבות
 *
 * מ-25,363 המפתחות שבמקור, 6,508 אינם ראשי תיבות כלל — 3,862 מילים ארמיות
 * בלי שום גרש (מילון מונחים שנכנס לאותו קובץ), 2,604 ערכים מרובי מילים
 * ש-`lookup` אינו יכול להגיע אליהם ממילא (הוא בודק את המילה האחרונה
 * שהוקלדה), ועוד 42 עם סוגריים, מקפים או תווים שנשברו בקידוד. הם נופלים
 * בשער `looksLikeAcronym` — אותו שער עצמו שמונע טעינת הנכס אחרי מילה רגילה —
 * ולכן הכנסתם לנכס הייתה מחייבת לפתוח אותו, כלומר להזריק 0.7MB אחרי כל
 * הקלדה. עוד 522 נושרים כשאין להם אף פירוש שימושי — ר' `HAS_HEBREW` למטה.
 * נשארים 17,717.
 *
 * ## כיווץ: פירוש אחד לכל ערך
 *
 * ל-8,877 מפתחות יש יותר מפירוש אחד, וה-UI מציג את הראשון השימושי. הנכס
 * נושא רק אותו — קובץ המקור נשאר שלם, ולכן מחזור בין פירושים (כשימומש) יטען
 * נכס שני משם.
 *
 * המספרים אינם הערכה: `tests/unit/acronyms.test.ts` מאמת כל אחד מהם על הקובץ
 * עצמו, כדי שהחלפת נתונים שמזיזה אותם תיראה ולא תעבור בשקט.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { looksLikeAcronym, normalizeAcronym } from '../src/engine/acronyms';
import { ACRONYMS_GLOBAL } from '../src/engine/acronyms-constants';

/** מ-`cwd` ולא מ-`import.meta.url`: תחת jsdom הכתובת אינה file://, כמו ב-scripts/blank-docx.ts. */
const SOURCE = resolve('src/data/acronyms.json');

/**
 * פירוש חייב להכיל אות עברית אחת לפחות.
 *
 * 503 מפתחות ב-KleiKodesh הם לעזי רש"י שהפירוש היחיד שלהם הוא ממלא-מקום —
 * `*`, `/`, `/ ?`, `/ antenas`. בלי המבחן הזה הקלדת „ב''ג ” הייתה מציעה „*”,
 * ו-Tab היה כותב אותו למסמך.
 */
const HAS_HEBREW = /[א-ת]/;

export interface AcronymsPack {
  readonly packed: Record<string, string>;
  /** מה נשר, לפי סיבה. */
  readonly dropped: { readonly notAcronym: number; readonly noExpansion: number };
  /** מפתחות שנכנסו רק בצורתם המנורמלת (גרש כפול, ניקוד) ולא היו קיימים ישירות. */
  readonly fromVariants: number;
}

export function packAcronyms(data: Record<string, unknown>): AcronymsPack {
  const packed: Record<string, string> = Object.create(null);
  /** מפתחות שצורתם הכתובה אינה המנורמלת. נדחים לסוף כדי שלא ידרסו ערך ישיר. */
  const variants: [string, string][] = [];
  let notAcronym = 0;
  let noExpansion = 0;

  for (const [key, value] of Object.entries(data)) {
    if (!Array.isArray(value) || value.some((item) => typeof item !== 'string')) {
      throw new Error(`acronyms.json: הערך של "${key}" אינו מערך מחרוזות.`);
    }
    if (!looksLikeAcronym(key)) {
      notAcronym += 1;
      continue;
    }

    const normalized = normalizeAcronym(key);
    // „רש"י → רש"י” אינו פירוש, ואף „*” אינו. רווח בקצה ורווח כפול קיימים
    // גם הם בנתונים, ומסלול הר"ת מכניס את הפירוש כפי שהוא.
    const expansion = (value as string[])
      .map((item) => item.replace(/\s+/g, ' ').trim())
      .find((item) => item !== '' && HAS_HEBREW.test(item) && normalizeAcronym(item) !== normalized);
    if (expansion === undefined) {
      noExpansion += 1;
      continue;
    }

    if (normalized === key) packed[key] = expansion;
    else variants.push([normalized, expansion]);
  }

  let fromVariants = 0;
  for (const [key, expansion] of variants) {
    if (key in packed) continue;
    packed[key] = expansion;
    fromVariants += 1;
  }

  return { packed, dropped: { notAcronym, noExpansion }, fromVariants };
}

export function readAcronymsSource(source: string = SOURCE): Record<string, unknown> {
  return JSON.parse(readFileSync(source, 'utf8')) as Record<string, unknown>;
}

/** תוכן `assets/acronyms.js`: השמה יחידה של ליטרל אובייקט ל-`window`. */
export function buildAcronymsAsset(source: string = SOURCE): string {
  const { packed } = packAcronyms(readAcronymsSource(source));
  return `window.${ACRONYMS_GLOBAL} = ${JSON.stringify(packed)};\n`;
}

/**
 * כל מפתח הזרקה ב-`composables/keys.ts` חייב שמישהו יספק אותו.
 *
 * הבדיקה הזאת נכתבה אחרי באג אמיתי: `PAGE_MARKING` ו-`DRAFT_OPENER` סופקו
 * ב-App.vue, וקומיט מאוחר יותר מחק את שתי השורות. שום דבר לא נשבר בהידור
 * ושום בדיקה לא האדימה — מפני ש-`inject` נקרא עם **ברירת מחדל**, וברירת
 * המחדל היא no-op תקין לחלוטין. התוצאה הייתה ש„סימון עמודים” נכשל בכל ריצה
 * עם „לא נמצאו עמודים לסימון”, ו„פירוק מסמך” לא פתח טאב — שתי תכונות
 * שנעלמו בשקט מוחלט.
 *
 * ברירת המחדל של `inject` נחוצה (היא מה שמאפשר לבדוק לשונית לבדה), ולכן
 * המחיר שלה משולם כאן: הבדיקה קוראת את המקור עצמו ומוודאת שלכל מפתח יש
 * `provide` אחד לפחות. מפתח שנולד בלי ספק ייכשל כאן ביום שהוא נולד.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join, relative } from 'node:path';

// vitest רץ משורש המאגר, ולכן cwd הוא השורש.
const SRC = join(process.cwd(), 'src');
const KEYS = join(SRC, 'composables', 'keys.ts');

function sourceFiles(dir = SRC): string[] {
  const files: string[] = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...sourceFiles(full));
    else if (/\.(ts|vue)$/.test(entry.name) && !entry.name.endsWith('.d.ts')) files.push(full);
  }
  return files;
}

/**
 * הערות מוסרות: התיעוד של כל מפתח מזכיר במפורש מי מספק אותו ומי מזריק,
 * ואזכור בהערה אינו חיווט. השורות נשמרות כדי שמספרי השורות יישארו נכונים.
 */
function stripComments(text: string): string {
  return text
    .replace(/\/\*[\s\S]*?\*\//g, (block) => block.replace(/[^\n]/g, ' '))
    .replace(/\/\/.*$/gm, '');
}

const sources = sourceFiles().map((path) => ({
  path: relative(SRC, path).replace(/\\/g, '/'),
  text: stripComments(readFileSync(path, 'utf8')),
}));

/**
 * שתי צורות ההצהרה, ולא אחת: `const K: InjectionKey<T> = Symbol('k')` וגם
 * `const K = Symbol('k') as InjectionKey<T>`. השנייה מהדרת ומתנהגת בדיוק
 * כמו הראשונה, וכשהיא לא נקראה כאן היא נולדה בלי ספק והשער נשאר ירוק —
 * כלומר ההבטחה שבראש הקובץ הייתה שקר לצורה אחת מתוך שתיים.
 */
const KEY_DECLARATIONS = [
  /export const ([A-Z_][A-Z0-9_]*)\s*:\s*InjectionKey</g,
  /export const ([A-Z_][A-Z0-9_]*)\s*=[^;]*?\bas\s+InjectionKey</g,
];

function declaredKeys(source: string): string[] {
  return Array.from(
    new Set(
      KEY_DECLARATIONS.flatMap((pattern) =>
        Array.from(source.matchAll(new RegExp(pattern.source, 'g')), (match) => match[1]!),
      ),
    ),
  );
}

const keyNames = declaredKeys(stripComments(readFileSync(KEYS, 'utf8')));

/**
 * הקבצים שבהם המפתח מופיע בקריאה כזאת — `inject<T>(KEY` נתפס גם הוא.
 *
 * ‏`[^()\n]*` ולא `[^>]*`: גנריקה מקוננת כמו `provide<Ref<number>>(KEY, x)`
 * נעצרה ב-`>` הראשון, ולכן קוד תקין לחלוטין לא נתפס והשער האדים על כלום.
 * הסוגריים מוצאים מהמחלקה כדי שהחיפוש לא ידלג מעל קריאה אחרת עד ל-`(` רחוק,
 * והשורה החדשה כדי שלא ייסרק על פני כל הקובץ.
 */
function callPattern(verb: 'provide' | 'inject', key: string): RegExp {
  return new RegExp(`\\b${verb}\\s*(?:<[^()\\n]*>)?\\s*\\(\\s*${key}\\b`);
}

function callers(verb: 'provide' | 'inject', key: string): string[] {
  const pattern = callPattern(verb, key);
  return sources.filter(({ text }) => pattern.test(text)).map(({ path }) => path);
}

describe('מפתחות ההזרקה של המעטפת', () => {
  it('נקראו מהמקור', () => {
    // אם הרשימה התרוקנה, הבדיקה שמתחת עוברת על אפס מפתחות ואינה בודקת דבר.
    expect(keyNames.length).toBeGreaterThan(5);
    expect(keyNames).toContain('PAGE_MARKING');
    expect(keyNames).toContain('DRAFT_OPENER');
  });

  /*
    שתי הבדיקות הבאות הן על הכלים עצמם, על מקור סינתטי. בלעדיהן ה„חיפוש”
    יכול להפסיק לחפש בלי שאף בדיקה תאדים — וזה בדיוק מה שקרה: הצורה
    `as InjectionKey<` לא נקראה כלל, וגנריקה מקוננת לא נתפסה כספק.
  */
  it('שתי צורות ההצהרה נקראות', () => {
    const synthetic = [
      "export const FIRST: InjectionKey<Ref<number>> = Symbol('first');",
      "export const SECOND = Symbol('second') as InjectionKey<Map<string, number>>;",
      "const NOT_EXPORTED: InjectionKey<number> = Symbol('x');",
    ].join('\n');
    expect(declaredKeys(synthetic)).toEqual(['FIRST', 'SECOND']);
  });

  it('גנריקה מקוננת נתפסת כספק', () => {
    const key = 'SOME_KEY';
    expect(callPattern('provide', key).test(`provide<Ref<number>>(${key}, x);`)).toBe(true);
    expect(callPattern('provide', key).test(`provide(${key}, x);`)).toBe(true);
    expect(callPattern('inject', key).test(`inject<Ref<Map<string, number>>>(${key}, fallback);`)).toBe(true);
    // ואינה נמתחת מעבר לקריאה: `provide` של מפתח אחר אינו ספק של זה.
    expect(callPattern('provide', key).test(`provide(OTHER, x);\nfoo(${key});`)).toBe(false);
  });

  it('לכל מפתח יש מי שמספק אותו', () => {
    const orphans = keyNames.filter((key) => callers('provide', key).length === 0);
    expect(orphans).toEqual([]);
  });

  it('מי שמזריק מקבל ספק — גם כשיש ברירת מחדל', () => {
    const broken = keyNames
      .filter((key) => callers('inject', key).length > 0 && callers('provide', key).length === 0)
      .map((key) => `${key}: מוזרק ב-${callers('inject', key).join(', ')} ואינו מסופק`);
    expect(broken).toEqual([]);
  });
});

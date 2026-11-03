/**
 * תיקון אוטומטי בהקלדה, כמו ב-Word: ההחלטה **מה** לתקן — טהורה ונבדקת ביחידה.
 * החיווט למנוע (קריאת הסמן, ההחלה, הביטול) ב-`autocorrect-install.ts`.
 *
 * ## הקלט: הטקסט שלפני הסמן, והסמן בסופו
 *
 * `planAutoCorrect` מקבלת את הטקסט מתחילת חלון ועד **הסמן עצמו**, ומחזירה
 * רשימת תיקונים בהיסטים יחסיים אליו. היא אינה יודעת דבר על המנוע, על בלוקים או
 * על היסטים במסמך — וזה מה שמאפשר לבדוק כל חוק על מחרוזת ולא על DOM.
 *
 * ## „מילה שהושלמה”: ההבדל בין חוק סמל לחוק מילה
 *
 * Word מתקן מילה רק כשהמשתמש **סיים** אותה — הקליד אחריה רווח או פיסוק. מילה
 * שהסמן עדיין נוגע בה אינה מתוקנת: `teh` באמצע ההקלדה היה הופך ל-`the` לפני
 * שנכתבה ה-`n` של `tehran`. לכן:
 *
 * - **חוק מילה** (דפוס שמתחיל או מסתיים באות או ספרה: `teh`, `1/2`, `c++`) —
 *   רק על מילה שאחריה רווח, או שאחרי ליבתה הוקלד סימן פיסוק (`teh,`).
 * - **חוק סמל** (דפוס שהתו הראשון **והאחרון** שלו אינם אות או ספרה: `(c)`,
 *   `...`, `-->`) — מיד כשהוא שלם. אין בו „המשך מילה” שעלול לשנות את משמעותו,
 *   וב-Word `(c)` הופך ל-© ברגע שנסגר. האות שבתוך `(c)` אינה הופכת אותו למילה:
 *   הסוגריים סוגרים אותו משני צדיו. כך גם רצפים שהמשתמש מגדיר למקש שאינו
 *   בשימוש (`/` ← `"`, `==` ← `?`): הם מוחלפים מיד, גם בתוך מילה. שני חריגים —
 *   דפוס שהוא תחילתו של דפוס סמל ארוך ממנו ממתין לתו הבא (`mayStillGrow`),
 *   וטוקן שהוא כתובת אינו נוגע (`ADDRESS_TOKEN`).
 *
 * ## חוקים שמסתכלים קדימה דורשים תו אחד אחריהם
 *
 * `א--ב` ו-`"` בין שתי אותיות עבריות נקבעים לפי **שני** הצדדים, ולכן אינם
 * מתקנים עד שהתו שאחריהם הוקלד. תו הגרש בסוף מילה (`ר'`) זקוק לתו שאחריו מאותה
 * סיבה: בלעדיו אי אפשר לדעת אם זו מילה שהסתיימה או הקלדה באמצע.
 *
 * ## למה אין כאן דבר על ביטול
 *
 * ‏Ctrl+Z מחזיר את הטקסט המקורי, ואז הסריקה הבאה הייתה מתקנת אותו שוב. מי שזוכר
 * מה בוטל הוא המתקין (`autocorrect-install.ts`), כי רק הוא יודע היסטים אבסולוטיים.
 * כאן כל קריאה היא מחדש.
 */

import wordListRaw from '../data/autocorrect-word.tsv?raw';

/** צמד „החלף” → „ב”. */
export interface AutoCorrectEntry {
  readonly from: string;
  readonly to: string;
}

export interface AutoCorrectOptions {
  /** המתג הראשי. כבוי = דבר אינו מתוקן. */
  readonly enabled: boolean;
  /** טבלת „החלף טקסט תוך כדי הקלדה”. */
  readonly replaceAsYouType: boolean;
  /** `--` ← `—`, ו-` -- ` ← ` – `. */
  readonly dashes: boolean;
  /** ` - ` בין שתי מילים ← ` – `. כבוי כברירת מחדל: מקף עם רווחים נפוץ בעברית. */
  readonly spacedHyphen: boolean;
  /** `"` בין שתי אותיות ← ״, ו-`'` אחרי אות ← ׳. */
  readonly gershayim: boolean;
  /** כ ← ך בסוף מילה, ו-ך ← כ באמצעה. */
  readonly finalLetters: boolean;
  /** `TWo` ← `Two`. */
  readonly twoInitialCaps: boolean;
  /** אות ראשונה גדולה בתחילת משפט לטיני. */
  readonly capitalizeSentences: boolean;
  /** `monday` ← `Monday`. */
  readonly capitalizeDays: boolean;
  readonly replacements: readonly AutoCorrectEntry[];
}

/** תיקון אחד, בהיסטים יחסיים לטקסט שנמסר. */
export interface AutoCorrectFix {
  readonly start: number;
  readonly end: number;
  readonly original: string;
  readonly replacement: string;
  /** מזהה החוק — לבדיקות ולדיווח. */
  readonly rule: AutoCorrectRule;
}

export type AutoCorrectRule =
  | 'table'
  | 'dashes'
  | 'spaced-hyphen'
  | 'gershayim'
  | 'final-letters'
  | 'two-initial-caps'
  | 'capitalize-sentence'
  | 'capitalize-day';

/** מה שמגיע עם הטקסט ואינו ממנו. */
export interface AutoCorrectContext {
  /** הטקסט מתחיל בתחילת הבלוק — מילה ראשונה בו פותחת משפט. */
  readonly atBlockStart?: boolean;
  /** החלון מתחיל באמצע טוקן; אין להתייחס לסיומת שלו כמילה שלמה. */
  readonly startsInsideToken?: boolean;
  /**
   * אחרי הטקסט לא יוקלד תו נוסף באותה שורה — Enter או Tab. דפוס סמל שהמתין
   * כתחילית של דפוס ארוך ממנו (`mayStillGrow`) כבר לא יגדל, ומוחלף.
   */
  readonly ended?: boolean;
}

/**
 * תקרות הטבלה. התקרה מאפשרת גם רשימות אישיות גדולות שיוצאות מ-Word. חוקי
 * המילה במפה, וחוקי הסמל נסרקים על חלון של מאות תווים לכל היותר.
 */
const MAX_ENTRIES = 5000;
const MAX_FROM = 64;
const MAX_TO = 255;

/**
 * הרשימה שנשלחה עם המוצר, בשני חלקים: הסמלים המובנים כאן, והרשימה שיוצאה מ-Word
 * ב-`src/data/autocorrect-word.tsv`.
 *
 * ## למה קובץ טקסט ולא מערך בקוד
 *
 * הרשימה היא נתונים, לא לוגיקה: מי שמעדכן אותה מחליף קובץ, לא עורך TypeScript.
 * והקובץ הוא **בדיוק** מה שהמאקרו `scripts/word/ExportAutoCorrect.bas` מוציא מ-Word
 * (שורה לכל רשומה: ההחלפה, טאב, ומה שיבוא במקומה), כך שאין שלב המרה שאפשר לטעות
 * בו. `tests/unit/autocorrect-word-list.test.ts` נכשל אם שורה בו פגומה או כפולה.
 *
 * הסמלים מצורפים לברירת המחדל. הקיצורים העבריים הם רשימה אישית, ונטענים רק
 * בבחירה מפורשת: למשל „למד” יכול להיות מילה תקינה ולא למ״ד.
 */
export const BUILT_IN_REPLACEMENTS: readonly AutoCorrectEntry[] = [
  { from: '(c)', to: '©' },
  { from: '(r)', to: '®' },
  { from: '(tm)', to: '™' },
  { from: '...', to: '…' },
  { from: '-->', to: '→' },
  { from: '<--', to: '←' },
  { from: '==>', to: '⇒' },
  { from: '<==', to: '⇐' },
  { from: '1/2', to: '½' },
  { from: '1/4', to: '¼' },
  { from: '3/4', to: '¾' },
];

/** רשימה אישית שיוצאה מ-Word. נטענת רק בבחירה מפורשת בדיאלוג. */
export const WORD_REPLACEMENTS: readonly AutoCorrectEntry[] = parseWordList(wordListRaw).entries;

// קיצורים אישיים כמו למד/טס/סל הם גם מילים תקינות. משאירים אותם לבחירה
// מפורשת; הסמלים נשארים זמינים כברירת מחדל.
export const DEFAULT_REPLACEMENTS: readonly AutoCorrectEntry[] = mergeEntries(
  BUILT_IN_REPLACEMENTS,
  WORD_REPLACEMENTS.filter((entry) => !/[\u05D0-\u05EA]/u.test(entry.from)),
);

/**
 * ברירות המחדל. כבויים בכוונה: האותיות הסופיות, הגרשיים והמקף עם רווחים — כולם
 * משנים טקסט עברי תקין לפי הנחה על הכוונה, ומי שכותב ר״ת או תעתיק יפגע מהם.
 * `capitalizeSentences` כבוי אף שב-Word הוא דלוק: מסמך עברי עם ציטוטים לטיניים
 * מלאי קיצורים (`vol.`, `p.`, `ibid.`) היה מקבל אות גדולה באמצע הפניה.
 */
export const DEFAULT_AUTOCORRECT: AutoCorrectOptions = {
  enabled: true,
  replaceAsYouType: true,
  dashes: true,
  spacedHyphen: false,
  gershayim: false,
  finalLetters: false,
  twoInitialCaps: true,
  capitalizeSentences: false,
  capitalizeDays: true,
  replacements: DEFAULT_REPLACEMENTS,
};

/* ------------------------------------------------------------------ */
/* נרמול ממה שנשמר                                                      */
/* ------------------------------------------------------------------ */

/**
 * הגדרות מהאחסון → הגדרות תקינות.
 *
 * ערך פגום אינו סיבה להיכשל: הצורה שנשמרה נקראת כשדה-שדה, ושדה שאינו תקין נופל
 * לברירת המחדל שלו. רשומה שגויה בטבלה נזרקת **לבדה** ואינה מפילה את האחרות.
 */
export function normalizeAutoCorrect(raw: unknown): AutoCorrectOptions {
  if (raw === null || typeof raw !== 'object') return DEFAULT_AUTOCORRECT;
  const source = raw as Record<string, unknown>;

  const flag = (key: Exclude<keyof AutoCorrectOptions, 'replacements'>): boolean =>
    typeof source[key] === 'boolean' ? (source[key] as boolean) : DEFAULT_AUTOCORRECT[key];

  return {
    enabled: flag('enabled'),
    replaceAsYouType: flag('replaceAsYouType'),
    dashes: flag('dashes'),
    spacedHyphen: flag('spacedHyphen'),
    gershayim: flag('gershayim'),
    finalLetters: flag('finalLetters'),
    twoInitialCaps: flag('twoInitialCaps'),
    capitalizeSentences: flag('capitalizeSentences'),
    capitalizeDays: flag('capitalizeDays'),
    replacements: Array.isArray(source.replacements)
      ? migrateReplacements(normalizeEntries(source.replacements))
      : DEFAULT_REPLACEMENTS,
  };
}

/** רשומות תקינות בלבד, בלי כפילויות לפי „החלף”, ובתקרה. */
export function normalizeEntries(list: readonly unknown[]): AutoCorrectEntry[] {
  const seen = new Set<string>();
  const out: AutoCorrectEntry[] = [];
  for (const item of list) {
    if (item === null || typeof item !== 'object') continue;
    const { from, to } = item as { from?: unknown; to?: unknown };
    if (typeof from !== 'string' || typeof to !== 'string') continue;
    if (from.length === 0 || from.length > MAX_FROM || to.length > MAX_TO) continue;
    if (/[\r\n]/.test(from) || /[\r\n]/.test(to)) continue;
    // דפוס עם רווח בקצה לעולם אינו מתאים: מילה מוגדרת בין רווחים.
    if (from !== from.trim()) continue;
    if (seen.has(from)) continue;
    seen.add(from);
    out.push({ from, to });
    if (out.length >= MAX_ENTRIES) break;
  }
  return out;
}

/**
 * טבלה שמורה שזהה לברירת המחדל **הישנה** (הסמלים המובנים בלבד) לא נבחרה על ידי
 * המשתמש — היא נשמרה כשלחץ „אישור” בלי לגעת בה. אותה מחליפים בברירת המחדל
 * הנוכחית, אחרת מי שפתח את הדיאלוג פעם אחת לא היה מקבל לעולם את הרשימה שנוספה.
 * טבלה ששונתה נשמרת כמו שהיא.
 */
function migrateReplacements(saved: AutoCorrectEntry[]): readonly AutoCorrectEntry[] {
  return sameEntries(saved, BUILT_IN_REPLACEMENTS) ? DEFAULT_REPLACEMENTS : saved;
}

/** שוות באורך, בסדר ובתוכן. */
export function sameEntries(a: readonly AutoCorrectEntry[], b: readonly AutoCorrectEntry[]): boolean {
  return a.length === b.length && a.every((entry, index) => entry.from === b[index].from && entry.to === b[index].to);
}

export interface WordListReport {
  readonly entries: AutoCorrectEntry[];
  /** שורות שלא נכנסו: בלי טאב, פגומות או כפולות. הערות וריקות אינן נספרות. */
  readonly skipped: number;
}

/**
 * הקובץ שהמאקרו מוציא מ-Word → רשומות. שורה לכל רשומה, טאב בין ההחלפה למה שיבוא
 * במקומה. שורה שמתחילה ב-`#` היא הערה, ושורה ריקה נבלעת.
 *
 * ‏`trim()` על ה„החלף” בלבד ולא על כל הקובץ: ערך שנגמר ברווח נשמר, ורשומה
 * שערכה ריק (`xx`, טאב) היא החלפה למחיקה ולא שורה בלי טאב.
 */
export function parseWordList(raw: string): WordListReport {
  const candidates = raw
    .replace(/^\uFEFF/, '')
    .split(/\r?\n/)
    .filter((line) => line.trim() !== '' && !line.startsWith('#'))
    .map((line) => {
      const tab = line.indexOf('\t');
      return tab < 0 ? null : { from: line.slice(0, tab).trim(), to: line.slice(tab + 1) };
    });

  const entries = normalizeEntries(candidates);
  return { entries, skipped: candidates.length - entries.length };
}

/**
 * מיזוג: רשומה ב„עליון” גוברת על רשומה ב„בסיס” עם אותו „החלף” ונשארת במקומה,
 * וחדשה נוספת בסוף.
 */
export function mergeEntries(
  base: readonly AutoCorrectEntry[],
  top: readonly AutoCorrectEntry[],
): AutoCorrectEntry[] {
  const out = base.map((entry) => ({ ...entry }));
  const index = new Map(out.map((entry, position) => [entry.from, position]));
  for (const entry of top) {
    const at = index.get(entry.from);
    if (at === undefined) {
      index.set(entry.from, out.length);
      out.push({ ...entry });
    } else {
      out[at] = { ...entry };
    }
  }
  return normalizeEntries(out);
}

/* ------------------------------------------------------------------ */
/* עזרים                                                                */
/* ------------------------------------------------------------------ */

/** אות או ספרה. */
const HAS_WORD_CHAR = /[\p{L}\p{N}]/u;

/**
 * דפוס „סמל”: התו הראשון **והאחרון** שלו אינם אות או ספרה. `(c)` ו-`...` כן;
 * `teh`, `1/2` ו-`c++` לא (הם מתחילים באות או בספרה).
 *
 * הבדיקה על הקצוות ולא על כל התווים: בדיקה על כל הדפוס סיווגה את `(c)` כחוק
 * מילה בגלל האות שבתוכו, והוא חיכה לרווח שאחריו — בניגוד ל-Word, ש-`(c)` שלו
 * הופך ברגע שהסוגר נסגר (נתפס בבדיקת היחידה).
 */
function isSymbolPattern(pattern: string): boolean {
  const chars = [...pattern];
  return !HAS_WORD_CHAR.test(chars[0] ?? '') && !HAS_WORD_CHAR.test(chars[chars.length - 1] ?? '');
}

/**
 * האם הדפוס מוחלף **מיד** כשהוקלד, גם בתוך מילה (`/` ב-`רש/י`), או רק כשהמילה
 * הושלמה. זה אותו כלל של חוקי הסמל, והדיאלוג מציג אותו ליד כל שורה.
 *
 * ## למה כלל ולא מתג לכל שורה
 *
 * שני הכיוונים של מתג כזה שוברים משהו: דפוס מילה „מיידי” היה הופך `teh` באמצע
 * `tehran`, ודפוס סמל „בסוף מילה” לעולם לא היה מוחלף בתוך מילה — `רש/י` אינו
 * שווה ל-`/`. הכלל נגזר מהדפוס עצמו, ולכן אין מצב שהמשתמש צריך לנחש.
 */
export function isImmediatePattern(pattern: string): boolean {
  return pattern.length > 0 && isSymbolPattern(pattern);
}

/**
 * דפוס סמל שההחלפה שלו מכילה אותו (`/` ← `//`) היה מוחלף שוב בכל הקשה: הסריקה
 * של חוקי הסמל רצה על כל האזור שהוקלד, וההחלפה נשארת בתוכו.
 */
export function isSelfFeedingEntry(entry: AutoCorrectEntry): boolean {
  return isImmediatePattern(entry.from) && entry.to !== entry.from && entry.to.includes(entry.from);
}

/**
 * טוקן שהוא כתובת: `http:/`, `file:/` או `www.` בראשו. דפוס סמל שהמשתמש הגדיר
 * למקש שאינו בשימוש (`/` ← `"`) היה הורס כל כתובת שמוקלדת. נבדק כבר מ-`http:/`,
 * כי זה מה שלפני הסמן כשה-`/` הראשון הוקלד.
 */
const ADDRESS_TOKEN = /^(?:[A-Za-z][A-Za-z\d+.-]*:\/|www\.)/u;

const LEADING_PUNCT = /^[("'[{“‘«<]+/u;
const TRAILING_PUNCT = /[.,;:!?)"'\]}”’»…>׃]+$/u;

interface Token {
  /** ההיסט של תחילת הטוקן (רצף של תווים שאינם רווח). */
  readonly start: number;
  /** תחילת הליבה, אחרי פיסוק פותח. */
  readonly coreStart: number;
  readonly coreEnd: number;
  readonly core: string;
  /** הליבה נסגרה: אחריה רווח או פיסוק. */
  readonly completed: boolean;
  /** הפיסוק שאחרי הליבה מתחיל בגרש (`'`) — כמו ב-`לכאו'`, שהגרש הוא חלק מהקיצור. */
  readonly geresh: boolean;
}

function tokenize(text: string, startsInsideToken = false): Token[] {
  const tokens: Token[] = [];
  for (const match of text.matchAll(/\S+/gu)) {
    const raw = match[0];
    const start = match.index ?? 0;
    const lead = raw.match(LEADING_PUNCT)?.[0].length ?? 0;
    const tailText = raw.slice(lead).match(TRAILING_PUNCT)?.[0] ?? '';
    const tail = tailText.length;
    const core = raw.slice(lead, raw.length - tail);
    if (core.length === 0) continue;

    const tokenEnd = start + raw.length;
    tokens.push({
      start,
      coreStart: start + lead,
      coreEnd: start + lead + core.length,
      core,
      // רווח אחריה, או פיסוק שהוקלד אחרי הליבה. טוקן שהסמן נוגע בו — לא.
      completed: (tokenEnd < text.length || tail > 0) && !(startsInsideToken && start === 0),
      geresh: tailText.startsWith("'"),
    });
  }
  return tokens;
}

/** אותיות גדולות/קטנות מהמקור עוברות להחלפה: `Teh` ← `The`, `TEH` ← `THE`. */
export function transferCase(original: string, pattern: string, replacement: string): string {
  if (original === pattern) return replacement;
  const letters = [...original].filter((char) => /\p{L}/u.test(char));
  if (letters.length === 0) return replacement;

  const upper = (char: string): boolean => char !== char.toLowerCase();
  if (letters.length > 1 && letters.every(upper)) return replacement.toUpperCase();
  if (upper(letters[0])) {
    const first = [...replacement][0] ?? '';
    return first.toUpperCase() + replacement.slice(first.length);
  }
  return replacement;
}

const HEBREW_LETTER = /[\u05D0-\u05EA]/u;
const NON_FINAL_TO_FINAL: Record<string, string> = { כ: 'ך', מ: 'ם', נ: 'ן', פ: 'ף', צ: 'ץ' };
const FINAL_TO_NON_FINAL: Record<string, string> = { ך: 'כ', ם: 'מ', ן: 'נ', ף: 'פ', ץ: 'צ' };

const ABBREVIATIONS = new Set([
  'e.g.', 'i.e.', 'etc.', 'vs.', 'cf.', 'ibid.', 'op.', 'cit.', 'mr.', 'mrs.', 'ms.', 'dr.',
  'prof.', 'no.', 'vol.', 'p.', 'pp.', 'ch.', 'fig.', 'st.', 'jr.', 'sr.',
]);

const DAYS = new Set(['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']);

/* ------------------------------------------------------------------ */
/* החוקים                                                               */
/* ------------------------------------------------------------------ */

interface CompiledTable {
  /** חוקי סמל, הארוך ראשון — `-->` לפני `--`. */
  readonly symbols: readonly AutoCorrectEntry[];
  /** חוקי מילה, מותאמים בדיוק. */
  readonly exact: ReadonlyMap<string, string>;
  /** דפוסים באותיות קטנות, מותאמים בלי רגישות לרישיות. */
  readonly folded: ReadonlyMap<string, string>;
  readonly hasPhrases: boolean;
}

/**
 * הטבלה מפוצלת ומסודרת פעם אחת לכל רשימה, לא בכל הערכה. ההערכה רצה אחרי כל
 * הקלדה, והרשימה שנשלחת עם המוצר היא אלפי רשומות — בנייה מחדש של שתי מפות וסינון
 * שלה בכל פעם היא עבודה שהתוצאה שלה זהה. המפתח הוא **הרשימה עצמה** (הפניה), ולכן
 * עריכה בדיאלוג, שמחליפה את המערך, מתחילה מטמון חדש מאליה.
 */
const compiledTables = new WeakMap<readonly AutoCorrectEntry[], CompiledTable>();

function compileTable(entries: readonly AutoCorrectEntry[]): CompiledTable {
  const cached = compiledTables.get(entries);
  if (cached) return cached;

  const words = entries.filter((entry) => !isSymbolPattern(entry.from));
  const compiled: CompiledTable = {
    symbols: entries
      .filter((entry) => isSymbolPattern(entry.from) && !isSelfFeedingEntry(entry))
      .sort((a, b) => b.from.length - a.from.length),
    exact: new Map(words.map((entry) => [entry.from, entry.to])),
    // דפוס שיש בו אות גדולה מותאם בדיוק בלבד.
    folded: new Map(words.filter((entry) => entry.from === entry.from.toLowerCase()).map((entry) => [entry.from, entry.to])),
    hasPhrases: words.some((entry) => /\s/u.test(entry.from)),
  };
  compiledTables.set(entries, compiled);
  return compiled;
}

/** האם ההיסט בתוך טוקן שהוא כתובת (`ADDRESS_TOKEN`). */
function insideAddress(text: string, index: number): boolean {
  const before = text.slice(0, index);
  const start = Math.max(before.lastIndexOf(' '), before.lastIndexOf('\t'), before.lastIndexOf(' ')) + 1;
  return ADDRESS_TOKEN.test(text.slice(start));
}

/**
 * מופע שנוגע בסמן, ודפוס סמל ארוך ממנו עדיין יכול להיות מושלם בתו הבא: `--`
 * כשבטבלה גם `-->`. בלי ההמתנה התוצאה הייתה תלויה בקצב — מי שעצר אחרי `--`
 * היה מקבל את הקצר, ומי שהקליד `-->` ברצף את הארוך. התו הבא (כל תו) מכריע.
 */
function mayStillGrow(
  text: string,
  index: number,
  hit: AutoCorrectEntry,
  symbols: readonly AutoCorrectEntry[],
): boolean {
  if (index + hit.from.length !== text.length) return false;
  const longest = symbols[0]?.from.length ?? 0;
  for (let from = Math.max(0, text.length - longest + 1); from <= index; from += 1) {
    const tail = text.slice(from);
    if (symbols.some((entry) => entry.from.length > tail.length && entry.from.startsWith(tail))) return true;
  }
  return false;
}

function tableFixes(
  text: string,
  tokens: readonly Token[],
  entries: readonly AutoCorrectEntry[],
  ended = false,
  waiting: AutoCorrectFix[] = [],
): AutoCorrectFix[] {
  const fixes: AutoCorrectFix[] = [];
  const { symbols, exact, folded, hasPhrases } = compileTable(entries);

  // חוקי סמל: מופע שלם בכל מקום, הארוך ראשון.
  for (let index = 0; index < text.length; ) {
    const hit = symbols.find((entry) => text.startsWith(entry.from, index));
    if (hit && insideAddress(text, index)) {
      index += 1;
    } else if (hit && !ended && mayStillGrow(text, index, hit, symbols)) {
      waiting.push({ start: index, end: index + hit.from.length, original: hit.from, replacement: hit.to, rule: 'table' });
      index += 1;
    } else if (hit) {
      fixes.push({
        start: index,
        end: index + hit.from.length,
        original: hit.from,
        replacement: hit.to,
        rule: 'table',
      });
      index += hit.from.length;
    } else {
      index += 1;
    }
  }

  // חוקי מילה: ליבת הטוקן שווה לדפוס, רק אחרי שהושלמה.
  const wordFixes: AutoCorrectFix[] = [];
  if (exact.size > 0) {
    for (const [position, token] of tokens.entries()) {
      if (!token.completed) continue;

      // ביטויים נבדקים בין גבולות טוקנים, עד MAX_FROM תווים. המפות נשארות
      // תלויות בטבלה, אך מספר המועמדים תלוי רק בחלון הטקסט המוגבל.
      if (hasPhrases) {
        for (let first = 0; first <= position; first += 1) {
          const startToken = tokens[first]!;
          if (first < position && !startToken.completed) continue;
          const start = startToken.coreStart;
          if (token.coreEnd - start > MAX_FROM) continue;
          const ends = token.geresh ? [token.coreEnd + 1, token.coreEnd] : [token.coreEnd];
          for (const end of ends) {
            const original = text.slice(start, end);
            const direct = exact.get(original);
            const to = direct ?? folded.get(original.toLowerCase());
            if (to === undefined) continue;
            const replacement = direct === undefined ? transferCase(original, original.toLowerCase(), to) : to;
            if (replacement !== original) wordFixes.push({ start, end, original, replacement, rule: 'table' });
            break;
          }
        }
        continue;
      }

      // קיצור שהגרש בסופו הוא חלק ממנו (`לכאו'`): הגרש נחתך מהליבה, ולכן מחפשים אותו גם איתו.
      const withGeresh = token.geresh ? exact.get(`${token.core}'`) : undefined;
      if (withGeresh !== undefined) {
        const original = `${token.core}'`;
        if (withGeresh !== original) {
          wordFixes.push({
            start: token.coreStart,
            end: token.coreEnd + 1,
            original,
            replacement: withGeresh,
            rule: 'table',
          });
        }
        continue;
      }

      const direct = exact.get(token.core);
      const to = direct ?? folded.get(token.core.toLowerCase());
      if (to === undefined) continue;
      const replacement = direct === undefined ? transferCase(token.core, token.core.toLowerCase(), to) : to;
      if (replacement === token.core) continue;
      wordFixes.push({
        start: token.coreStart,
        end: token.coreEnd,
        original: token.core,
        replacement,
        rule: 'table',
      });
    }
  }

  // ביטוי ארוך קודם להחלפה של מילה בתוכו. הסמלים שומרים על העדיפות הקיימת.
  return [...fixes, ...wordFixes.sort((a, b) => (b.end - b.start) - (a.end - a.start))];
}

function dashFixes(text: string, spacedHyphen: boolean): AutoCorrectFix[] {
  const fixes: AutoCorrectFix[] = [];
  const push = (index: number, length: number, replacement: string, rule: AutoCorrectRule): void => {
    fixes.push({ start: index, end: index + length, original: text.slice(index, index + length), replacement, rule });
  };

  // `א--ב` ← em dash. חוסמים `-->` בכך שמצפים לאות/ספרה אחרי.
  for (const match of text.matchAll(/(?<=[\p{L}\p{N}])--(?=[\p{L}\p{N}])/gu)) {
    push(match.index ?? 0, 2, '—', 'dashes');
  }
  // `א -- ב` ← en dash.
  for (const match of text.matchAll(/(?<=[\p{L}\p{N}]) -- (?=[\p{L}\p{N}])/gu)) {
    push((match.index ?? 0) + 1, 2, '–', 'dashes');
  }
  if (spacedHyphen) {
    for (const match of text.matchAll(/(?<=[\p{L}\p{N}]) - (?=[\p{L}\p{N}])/gu)) {
      push((match.index ?? 0) + 1, 1, '–', 'spaced-hyphen');
    }
  }
  return fixes;
}

function gershayimFixes(text: string): AutoCorrectFix[] {
  const fixes: AutoCorrectFix[] = [];
  // גרשיים: בין שתי אותיות עבריות בלבד. מרכאות סוגרות אחרי רווח אינן נוגעות.
  for (const match of text.matchAll(/(?<=[\u05D0-\u05EA])"(?=[\u05D0-\u05EA])/gu)) {
    const index = match.index ?? 0;
    fixes.push({ start: index, end: index + 1, original: '"', replacement: '״', rule: 'gershayim' });
  }
  // גרש: אחרי אות עברית, כשיש תו אחריו (אות — ג'ירף, או כל דבר אחר — ר').
  for (const match of text.matchAll(/(?<=[\u05D0-\u05EA])'(?=[\s\S])/gu)) {
    const index = match.index ?? 0;
    fixes.push({ start: index, end: index + 1, original: "'", replacement: '׳', rule: 'gershayim' });
  }
  return fixes;
}

function finalLetterFixes(tokens: readonly Token[]): AutoCorrectFix[] {
  const fixes: AutoCorrectFix[] = [];
  for (const token of tokens) {
    if (!token.completed) continue;
    const letters = [...token.core];
    // רק מילה שכולה אותיות עבריות, ובאורך שתיים לפחות: אות בודדת היא קיצור.
    if (letters.length < 2 || !letters.every((letter) => HEBREW_LETTER.test(letter))) continue;

    letters.forEach((letter, index) => {
      const isLast = index === letters.length - 1;
      const target = isLast ? NON_FINAL_TO_FINAL[letter] : FINAL_TO_NON_FINAL[letter];
      if (!target) return;
      const at = token.coreStart + index;
      fixes.push({ start: at, end: at + 1, original: letter, replacement: target, rule: 'final-letters' });
    });
  }
  return fixes;
}

function capsFixes(
  text: string,
  tokens: readonly Token[],
  options: AutoCorrectOptions,
  context: AutoCorrectContext,
): AutoCorrectFix[] {
  const fixes: AutoCorrectFix[] = [];

  tokens.forEach((token, position) => {
    if (!token.completed) return;
    const { core } = token;
    const fix = (replacement: string, rule: AutoCorrectRule): void => {
      fixes.push({ start: token.coreStart, end: token.coreEnd, original: core, replacement, rule });
    };

    if (options.twoInitialCaps && /^[A-Z]{2}[a-z]+$/.test(core) && !/^[A-Z]{2}s$/.test(core)) {
      fix(core[0] + core.slice(1).toLowerCase(), 'two-initial-caps');
      return;
    }

    if (options.capitalizeDays && DAYS.has(core)) {
      fix(core[0].toUpperCase() + core.slice(1), 'capitalize-day');
      return;
    }

    if (options.capitalizeSentences && /^[a-z]/.test(core)) {
      const previous = tokens[position - 1];
      const opensSentence =
        previous === undefined
          ? context.atBlockStart === true && token.start === 0
          : previous.completed && startsSentenceAfter(text, previous);
      if (opensSentence) fix(core[0].toUpperCase() + core.slice(1), 'capitalize-sentence');
    }
  });

  return fixes;
}

/** האם הטוקן הקודם סיים משפט, ואינו קיצור או סימן רשימה. */
function startsSentenceAfter(text: string, previous: Token): boolean {
  // הסוף **בפועל** של הטוקן, כולל הפיסוק שאחרי הליבה.
  let end = previous.coreEnd;
  while (end < text.length && !/\s/u.test(text[end])) end += 1;
  const raw = text.slice(previous.start, end);
  if (!/[.!?]$/.test(raw)) return false;
  const lower = raw.toLowerCase();
  if (ABBREVIATIONS.has(lower)) return false;
  // „1.” ו-„a.” הם סימני רשימה, לא סוף משפט.
  if (/^[\p{N}]+[.)]$/u.test(raw) || /^[a-z][.)]$/i.test(raw)) return false;
  return true;
}

/* ------------------------------------------------------------------ */
/* התכנון                                                               */
/* ------------------------------------------------------------------ */

/**
 * התיקונים לטקסט שלפני הסמן, ממוינים מההיסט הגבוה לנמוך — הסדר שבו מחילים אותם
 * כדי שהיסט שכבר חושב לא יזוז. תיקונים שחופפים: הראשון לפי עדיפות החוקים מנצח.
 */
export function planAutoCorrect(
  text: string,
  options: AutoCorrectOptions,
  context: AutoCorrectContext = {},
): AutoCorrectFix[] {
  return planAutoCorrectWithPending(text, options, context).fixes;
}

/** The installer also needs the deferred suffix; derive it during the same bounded planning pass. */
export function planAutoCorrectWithPending(
  text: string,
  options: AutoCorrectOptions,
  context: AutoCorrectContext = {},
): { fixes: AutoCorrectFix[]; pending: AutoCorrectFix | null } {
  if (!options.enabled || text.length === 0) return { fixes: [], pending: null };

  const tokens = tokenize(text, context.startsInsideToken);
  const candidates: AutoCorrectFix[] = [];

  const waiting: AutoCorrectFix[] = [];
  if (options.replaceAsYouType) candidates.push(...tableFixes(text, tokens, options.replacements, context.ended, waiting));
  if (options.dashes || options.spacedHyphen) {
    candidates.push(...dashFixes(text, options.spacedHyphen).filter((fix) => options.dashes || fix.rule === 'spaced-hyphen'));
  }
  if (options.gershayim) candidates.push(...gershayimFixes(text));
  if (options.finalLetters) candidates.push(...finalLetterFixes(tokens));
  if (options.twoInitialCaps || options.capitalizeDays || options.capitalizeSentences) {
    candidates.push(...capsFixes(text, tokens, options, context));
  }

  // בחירה לפי הסדר שבו החוקים נוספו (= עדיפות), ודילוג על כל מה שחופף.
  const chosen: AutoCorrectFix[] = [];
  for (const candidate of candidates) {
    if (candidate.original === candidate.replacement) continue;
    const overlaps = chosen.some((fix) => candidate.start < fix.end && fix.start < candidate.end);
    if (!overlaps) chosen.push(candidate);
  }

  const pending = waiting.find((candidate) =>
    candidate.original !== candidate.replacement &&
    !chosen.some((fix) => candidate.start < fix.end && fix.start < candidate.end),
  ) ?? null;
  return { fixes: chosen.sort((a, b) => b.start - a.start), pending };
}

/**
 * השלמה מרשימות סטטיות: ביטויים תלמודיים (325) ושמות מחברים (651). מודול
 * טהור — בלי DOM ובלי RPC. הטעינה ב-static-completion-dictionary.ts.
 *
 * מופעלת רק אחרי שההשלמה מהספר הפתוח בקורא לא מצאה התאמה לאותה הקשה, ולפני
 * ראשי-התיבות.
 *
 * ## למה אינדקס משלה, ולא `buildSectionCache` על הרשומות
 *
 * המסלול המתבקש היה `buildSectionCache(entries.join('\n'))` — לחבר את
 * הרשומות לטקסט אחד ולהשתמש במנוע ההתאמה של הספר. הוא **שובר את הגבולות**:
 * המטמון של הספר מחזיק מערך שטוח של מילים ואינו יודע איפה רשומה נגמרה, ולכן
 * `sliceWords` חוצה בחופשיות מרשומה לרשומה. שתי רשומות עוקבות —
 *
 *     אבן השתיה
 *     אבן שאין לה הופכין
 *
 * — הופכות לרצף אחד, והקלדת „אבן הש” יכולה להחזיר „השתיה אבן שאין לה”:
 * סוף הביטוי הראשון מודבק לתחילת הבא. הצעה כזאת אינה שגויה בקצה אחד, היא
 * ביטוי שלא קיים.
 *
 * לכן כאן כל רשומה היא יחידה נפרדת, וההשלמה נגזרת מתוכה בלבד — חריגה מעבר
 * לסוף רשומה אינה „נמנעת”, היא בלתי אפשרית לפי מבנה הנתונים.
 *
 * ## ולמה אינדקס תחיליות, ולא סריקה
 *
 * `matchWithContext` של הספר סורק את כל המילים בכל הקשה — סביר ל-section
 * אחד, בזבוז על שתי רשימות קבועות. כאן כל מילה נכנסת לדלי לפי שתי אותיותיה
 * הראשונות, וההתאמה בודקת דלי אחד קטן במקום 3,106 מילים.
 *
 * ## ולמה עד סוף הרשומה, ולא חמש מילים
 *
 * בספר יש „המשך” — `continueFrom` מציע את חמש המילים הבאות בכל Tab. לרשומה
 * אין המשך: מה שהמשתמש רוצה הוא הביטוי השלם או השם השלם, וקיצוץ באמצע היה
 * משאיר אותו בלי דרך להשלים את השאר. הרשומה הארוכה כאן היא 10 מילים.
 */
import { hasEnoughSignal, normalizeTypedWord, type TypedContext } from './book-completion';
import { WORD_INNER } from './word-selection';

const WORD_RUN = new RegExp(`(?:${WORD_INNER.source})+`, 'gu');

/**
 * מיקום מילה: אינדקס רשומה ואינדקס מילה בתוכה, ארוזים במספר אחד. מערך
 * מספרים ולא מערך אובייקטים — 3,106 מילים, ולכל אחת שני דליים.
 */
const WORD_BITS = 5;
const MAX_WORDS_PER_ENTRY = (1 << WORD_BITS) - 1;

const pack = (entryIndex: number, wordIndex: number): number => (entryIndex << WORD_BITS) | wordIndex;
const entryOf = (position: number): number => position >>> WORD_BITS;
const wordOf = (position: number): number => position & MAX_WORDS_PER_ENTRY;

interface StaticEntry {
  /** הטקסט כפי שהוא במקור — כולל פסיקים, סוגריים וניקוד שיש בשמות מחברים. */
  readonly text: string;
  /** המילים המנורמלות, להשוואה מול מה שהוקלד. */
  readonly words: readonly string[];
  /** ההיסטים ב-`text` של כל מילה, באותו סדר. */
  readonly spans: readonly { readonly start: number; readonly end: number }[];
}

export interface StaticIndex {
  /** שם המקור, לדיווח ולבדיקות. */
  readonly name: string;
  readonly entries: readonly StaticEntry[];
  /** שתי האותיות הראשונות של מילה (או אות אחת) → מיקומיה. */
  readonly byPrefix: ReadonlyMap<string, readonly number[]>;
}

/**
 * מה שנכנס למסמך חייב להיות נקי.
 *
 * שתי רשומות ב-`authors.json` מכילות U+200F (סימן כיווניות), ובאחת הוא יושב
 * **בין** מילים — כלומר הצעה שמתחילה באמצע השם הייתה נושאת אותו אל תוך
 * ה-DOCX. `\s+` תופס גם רווח כפול, שקיים גם הוא בנתונים.
 */
function sanitizeEntry(text: string): string {
  return text.replace(/[\u200E\u200F\u061C]/g, '').replace(/\s+/g, ' ').trim();
}

/**
 * מסירה סוגר שאין לו פותח בתוך ההשלמה עצמה.
 *
 * ההשלמה היא סיפא של הרשומה, ולכן הצעה שמתחילה **בתוך** הסוגריים נושאת את
 * הסוגר בלי הפותח: „בעל הק” ברשומה „אריה ליב בן יוסף הכהן (בעל הקצות)” הציע
 * „הקצות)”. הרשומות עצמן מאוזנות (נבדק: 0 חריגות ב-976), ולכן אחרי ההסרה
 * הזאת מה שנכתב למסמך מאוזן תמיד.
 */
function dropOrphanClosers(text: string): string {
  let depth = 0;
  let out = '';
  for (const ch of text) {
    if (ch === '(') depth += 1;
    else if (ch === ')') {
      if (depth === 0) continue;
      depth -= 1;
    }
    out += ch;
  }
  return out.trimEnd();
}

export function buildStaticIndex(name: string, sourceEntries: readonly string[]): StaticIndex {
  const entries: StaticEntry[] = [];
  const byPrefix = new Map<string, number[]>();

  for (const raw of sourceEntries) {
    const text = sanitizeEntry(raw);
    const runs = [...text.matchAll(WORD_RUN)];
    if (runs.length === 0) continue;
    if (runs.length > MAX_WORDS_PER_ENTRY) {
      throw new Error(`${name}: לרשומה "${text}" יש ${runs.length} מילים, מעל ${MAX_WORDS_PER_ENTRY}.`);
    }

    const spans = runs.map((run) => ({ start: run.index ?? 0, end: (run.index ?? 0) + run[0].length }));
    const words = runs.map((run) => normalizeTypedWord(run[0]));
    const entryIndex = entries.length;
    entries.push({ text, words, spans });

    words.forEach((word, wordIndex) => {
      if (word === '') return;
      const position = pack(entryIndex, wordIndex);
      // שני דליים לכל מילה: עוגן בן אות אחת („א”) מחפש בדלי של אות אחת,
      // וארוך ממנה — בדלי של שתיים. בלי הראשון אין לו שום דלי להיכנס אליו.
      for (const key of word.length >= 2 ? [word.slice(0, 1), word.slice(0, 2)] : [word.slice(0, 1)]) {
        const bucket = byPrefix.get(key);
        if (bucket) bucket.push(position);
        else byPrefix.set(key, [position]);
      }
    });
  }

  return { name, entries, byPrefix };
}

export interface StaticMatch {
  /** מה להציע — מהמילה שתאמה ועד סוף הרשומה, ובטקסט המקור שלה. */
  readonly text: string;
  /** כמה מילות הקשר קודמות תאמו בפועל. */
  readonly contextWordsUsed: number;
  /** שם המקור שבו נמצאה ההתאמה. */
  readonly source: string;
}

export interface StaticMatchOptions {
  /** כמה מילות הקשר קודמות לנסות, מהארוך לקצר. ברירת מחדל 3. */
  maxContextWords?: number;
  /**
   * אורך מזערי למילה חלקית שאין לפניה שום הקשר תואם. ברירת מחדל 2 — ולא 3
   * כמו בספר: רשומה כאן היא ביטוי מוכר או שם, ושתי אותיות בראשו הן אות
   * מספקת. בספר, שהוא טקסט רץ, שתי אותיות תואמות עשרות מקומות.
   */
  minStandalonePartial?: number;
}

/**
 * מנסה כל מקור לפי הסדר ומחזירה את ההתאמה הראשונה. אין מיזוג בין מקורות —
 * מקור שמצא, השאר אינם נבדקים לאותה הקשה.
 */
export function matchStaticCompletion(
  indexes: readonly StaticIndex[],
  context: TypedContext,
  options: StaticMatchOptions = {},
): StaticMatch | null {
  const maxContextWords = options.maxContextWords ?? 3;
  const minStandalone = options.minStandalonePartial ?? 2;

  const queryWords = context.precedingWords
    .map(normalizeTypedWord)
    .filter((word) => word !== '')
    .slice(-maxContextWords);
  const partial = normalizeTypedWord(context.partialWord);

  /*
   * אורך ההקשר הוא הלולאה החיצונית, והמקור הפנימית — ולא להפך.
   *
   * נמדד בשער: „אבן גב” החזיר את הביטוי „גבור הכובש את יצרו” במקום את המחבר
   * „אבן גבאי, מאיר בן יחזקאל”. הביטויים נבדקו ראשונים, ומשלא נמצא בהם דבר
   * עם ההקשר „אבן” הם נבדקו שוב **בלי** הקשר — ומצאו. כלומר התאמה חלשה
   * במקור מועדף גברה על התאמה חזקה במקור הבא. הקשר תואם הוא האות החזקה כאן
   * (אותו נימוק כמו ב-`matchAtCursor`), וסדר המקורות מכריע רק ביניהם.
   */
  for (let contextLen = queryWords.length; contextLen >= 0; contextLen -= 1) {
    if (!hasEnoughSignal(contextLen, partial, minStandalone)) continue;
    const contextWords = queryWords.slice(queryWords.length - contextLen);
    for (const index of indexes) {
      const found = matchInIndex(index, contextWords, partial);
      if (found) return found;
    }
  }
  return null;
}

function matchInIndex(index: StaticIndex, contextWords: readonly string[], partial: string): StaticMatch | null {
  const anchor = contextWords.length > 0 ? contextWords[0]! : partial;
  if (anchor === '') return null;

  const bucket = index.byPrefix.get(anchor.slice(0, 2));
  if (!bucket) return null;

  for (const position of bucket) {
    const entry = index.entries[entryOf(position)]!;
    const at = wordOf(position);

    let matches = true;
    for (let k = 0; k < contextWords.length; k += 1) {
      if (entry.words[at + k] !== contextWords[k]) {
        matches = false;
        break;
      }
    }
    if (!matches) continue;

    const target = at + contextWords.length;
    // ההשלמה חייבת להיות בתוך אותה רשומה. זה הגבול שהמסלול הקודם איבד.
    if (target >= entry.words.length) continue;
    if (partial !== '' && !entry.words[target]!.startsWith(partial)) continue;

    /*
     * עד סוף **הטקסט**, ולא עד סוף המילה האחרונה. נמדד: „בן יוסף הכ” החזיר
     * „הכהן (בעל הקצות” — הסוגר שאחרי המילה האחרונה נחתך, ומה שנכתב למסמך
     * היה סוגריים לא מאוזנים. 58 מקרים כאלה בנתונים.
     */
    const text = dropOrphanClosers(entry.text.slice(entry.spans[target]!.start));
    if (text === '') continue;
    return { text, contextWordsUsed: contextWords.length, source: index.name };
  }
  return null;
}

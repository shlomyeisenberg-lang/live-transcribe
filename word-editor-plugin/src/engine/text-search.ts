/**
 * לוגיקה טהורה לחיפוש-והחלפה עצמאיים: מציאת כל המופעים של מחרוזת בטקסט
 * הקנוני של בלוקי המסמך, וניתוב שלהם בין מופעים (הבא/קודם), ומיפוי מופע
 * ליעד שהמנוע מקבל להחלפה/בחירה. אין כאן תלות במנוע כלל — לא `SuperDoc`, לא
 * Document API, לא async — וזה מה שמאפשר לבדוק את הלוגיקה בלי להרים אותו.
 *
 * ## למה מודול נפרד, ולמה הוא לא קורא ל-`ui.search`/`ui.search.find` בכלל
 *
 * `search.ts` הקודם היה עטיפה דקה סביב `ui.search` (`handle.search()` /
 * `handle.find()`) — כלומר כל „מציאה” וכל „החלפה” עברו דרך מנגנון פנימי של
 * המנוע שנקרא שם „projection”. זה נמדד שבור על מסמך רב-פסקאות: 8 מופעים על
 * פני כמה פסקאות, והמנוע דיווח 4 בלבד (`search()` מצא ארבע התאמות, ו-
 * `replaceAll()` היה נאמן להן — החליף את מה שנמצא, לא פחות). המנוע עצמו רשם
 * את האבחנה: `projection-incomplete: exact-complete projection did not cover
 * the full document (contiguous 1 of 7 ordinals)`. זו תקלה במנגנון הפנימי
 * שבניית ה"projection" מסתמכת עליו — ולא נפתרה גם ב-superdoc 2.10.0. פירוט
 * מלא ומדוד ב-docs/superdoc-2.10-review.md (”'החלף הכל' — האבחנה הקודמת
 * הייתה שגויה”) וב-docs/button-audit.md (טבלת „קשה לתקן”, שורה א').
 *
 * הפתרון כאן עוקף את ה-projection לגמרי: `doc.blocks.list({includeText:true})`
 * כבר משמש בביטחון ב-`engine/formatting-marks-layer.ts` וב-
 * `engine/line-number-layer.ts` לקריאת הטקסט הקנוני **המלא** של המסמך —
 * טקסט שאומת שם, פסקה-פסקה, מול DOM אמיתי בהתאמת-רצף. המודול הזה מקבל את
 * רשימת הבלוקים (הקריאה עצמה, כולל הדפדוף על פני `offset`/`limit`, יושבת
 * ב-`engine/search.ts` — התלוי-מנוע) ומוצא בעצמו את כל המופעים בטקסט שלהם.
 * שום דבר כאן אינו שואל את המנוע „איפה יש התאמות” — רק „איפה יש התאמות
 * *בטקסט הזה*”, וזו שאלה שה-JavaScript עונה עליה נכון בכל פעם.
 *
 * ההחלפה עצמה (`engine/search.ts`) משתמשת ב-`doc.replace({target, text})`
 * הציבורי עם `SelectionTarget` מדויק — לא בקבוצת-התאמות של המנוע — ולכן היא
 * גם אינה תלויה ב-projection. `matchToTarget` כאן היא בדיוק המיפוי מופע ←
 * יעד-מנוע, ונבדקת ביחידה כמו כל השאר.
 */

/** בלוק אחד כפי ש-`doc.blocks.list({includeText:true})` מחזיר אותו, בחלק שנצרך כאן. */
export interface SearchableBlock {
  /** `nodeId` של הבלוק — הוא ה-`blockId` ביעד ההחלפה/הבחירה. */
  blockId: string;
  /** הטקסט הקנוני, השטוח, של הבלוק. */
  text: string;
  /**
   * סוג הבלוק (`'paragraph'`, `'heading'`, `'listItem'` וכו'). אופציונלי:
   * `findAllMatches` אינו נזקק לו כלל, וקוד ישן/בדיקות ישנות שבונות
   * `SearchableBlock` בלי השדה הזה נשארים תקפים. הוא נדרש רק ב-
   * `engine/search.ts`, לבניית `BlockNodeAddress` עבור „החלף הכל" האטומי
   * (`doc.mutations.apply`) — ראו שם.
   */
  nodeType?: string;
}

/** מופע יחיד שנמצא: הבלוק שבו הוא נמצא, וטווח בקואורדינטות-הטקסט של הבלוק (לא ב-DOM). */
export interface TextMatch {
  blockId: string;
  start: number;
  end: number;
}

export interface FindMatchesOptions {
  /**
   * ברירת המחדל `false` — לא רגיש לרישיות. תואם את ברירת המחדל שנמדדה של
   * `ui.search` (`caseSensitive:false` ב-snapshot), כדי שההתנהגות שהמשתמש
   * מכיר לא תשתנה — רק הכיסוי שלה.
   */
  caseSensitive?: boolean;
}

/**
 * מוצאת את כל המופעים של `query` בכל בלוקי המסמך, בסדר המסמך (הבלוקים
 * עצמם כבר בסדר הזה — זה מה ש-`blocks.list()` מבטיח). חיפוש מילולי
 * (substring) ולא regex, בדיוק כמו הדיאלוג הקיים שאינו מציע מתג regex.
 * מופעים אינם חופפים: התאמה "צורכת" את הטווח שלה לפני שהחיפוש בבלוק ממשיך.
 *
 * שאילתה ריקה מחזירה מערך ריק — לא כשל, בדיוק כמו `handle.clear()` הישן.
 */
export function findAllMatches(
  blocks: readonly SearchableBlock[],
  query: string,
  options: FindMatchesOptions = {},
): TextMatch[] {
  if (!query) return [];
  const caseSensitive = options.caseSensitive ?? false;
  const needle = caseSensitive ? query : query.toLowerCase();
  if (!needle) return [];

  const matches: TextMatch[] = [];
  for (const block of blocks) {
    const haystack = caseSensitive ? block.text : block.text.toLowerCase();
    let from = 0;
    for (;;) {
      const at = haystack.indexOf(needle, from);
      if (at < 0) break;
      matches.push({ blockId: block.blockId, start: at, end: at + needle.length });
      from = at + needle.length;
    }
  }
  return matches;
}

/**
 * היעד שהמנוע מקבל כדי להחליף/לבחור מופע ספציפי: `SelectionTarget` שתחילתו
 * וסופו הם נקודת-טקסט (`{kind:'text', blockId, offset}`) — בדיוק המודל
 * הציבורי שמתועד ב-`SelectionPoint` (superdoc/ui), ואותו מודל ש-
 * `engine/caret-anchor.ts` כבר קורא ובונה משם לשחזור מיקום הסמן. מוצהר כאן
 * כצורה מקומית (לא מיובא) כדי ש-`engine/search.ts` יוכל להטיל אותה ל-
 * `SelectionTarget` האמיתי בנקודת השימוש, ומודול זה יישאר טהור לגמרי.
 */
export interface MatchSelectionTarget {
  kind: 'selection';
  start: { kind: 'text'; blockId: string; offset: number };
  end: { kind: 'text'; blockId: string; offset: number };
}

export function matchToTarget(match: TextMatch): MatchSelectionTarget {
  return {
    kind: 'selection',
    start: { kind: 'text', blockId: match.blockId, offset: match.start },
    end: { kind: 'text', blockId: match.blockId, offset: match.end },
  };
}

/**
 * המופע הבא/קודם לפי כיוון, עם מעגליות (מהאחרון לראשון ולהפך). `current:-1`
 * (אין התאמה פעילה) עם `total > 0`, בכיוון `next`, נכנס להתאמה הראשונה,
 * ובכיוון `prev` — לאחרונה: כך גם „מצא קודם” על שאילתה חדשה מגיע למופע
 * הגיוני, ולא ל-`-1 - 1`.
 */
export function advanceActiveIndex(
  current: number,
  total: number,
  direction: 'next' | 'prev',
): number {
  if (total <= 0) return -1;
  if (current < 0 || current >= total) return direction === 'next' ? 0 : total - 1;
  if (direction === 'next') return (current + 1) % total;
  return (current - 1 + total) % total;
}

/**
 * אינדקס ההתאמה הפעילה אחרי שמופע הוחלף: המופע שהיה באינדקס `previousIndex`
 * אינו קיים עוד ברשימה החדשה (קוצרה באחד), ולכן מה שכעת יושב באותו מספור
 * הוא בדיוק המופע ש**היה** אחריו — התקדמות טבעית ל„הבא” בלי צעד נוסף. אם
 * האינדקס הקודם היה מעבר לסוף הרשימה החדשה (הוחלף המופע האחרון), נצמד לחדש
 * האחרון; רשימה ריקה מחזירה `-1`.
 */
export function activeIndexAfterReplace(previousIndex: number, newTotal: number): number {
  if (newTotal <= 0) return -1;
  return Math.min(Math.max(previousIndex, 0), newTotal - 1);
}

/**
 * סדר ביצוע ל„החלף הכל”: לכל בלוק, מהמופע **האחרון** לראשון. החלפה משנה את
 * אורך הבלוק (`delta = replacement.length - (match.end - match.start)`), וכל
 * מופע אחרי הנקודה שהוחלפה זז בהתאם בתוך אותו בלוק — לכן החלפה מהסוף
 * ואחורה היא היחידה ששומרת על ההיסטים המקוריים תקפים למופעים הבאים
 * *באותו בלוק*. מופעים בבלוקים שונים בלתי-תלויים לגמרי (`blockId`+`offset`
 * של האחד אינם משפיעים על רעהו), ולכן הסדר בין בלוקים אינו משנה תוצאה, ורק
 * הסדר *בתוך* כל בלוק צריך להיות מפורש. הסדר בין בלוקים בפלט הוא סדר
 * ההופעה הראשונה שלהם ב-`matches` (סדר המסמך, כש-`matches` מגיע מ-
 * `findAllMatches`), כדי שהתוצאה תהיה דטרמיניסטית ונוחה לבדיקה.
 */
export function matchesForReplacement(matches: readonly TextMatch[]): TextMatch[] {
  const byBlock = new Map<string, TextMatch[]>();
  for (const match of matches) {
    const existing = byBlock.get(match.blockId);
    if (existing) existing.push(match);
    else byBlock.set(match.blockId, [match]);
  }
  const ordered: TextMatch[] = [];
  for (const inBlock of byBlock.values()) {
    ordered.push(...[...inBlock].sort((a, b) => b.start - a.start));
  }
  return ordered;
}

/** מספר המופעים בבלוק אחד. */
export interface BlockMatchCount {
  blockId: string;
  count: number;
}

/**
 * מקבצת מופעים לפי בלוק — כמה יש בכל אחד — בסדר הופעתם הראשונה (סדר
 * המסמך, כש-`matches` מגיע מ-`findAllMatches`).
 *
 * קיימת עבור „החלף הכל" האטומי (`engine/search.ts`): לכל בלוק עם מופעים
 * נבנה **צעד אחד** של `doc.mutations.apply` (`text.rewrite` עם `within`
 * שמוגבל לבלוק הזה בדיוק) — לא צעד לכל מופע — וכל הצעדים יחד נכנסים
 * לקריאה אטומית אחת, כלומר `undo`-step יחיד. `count` הוא הערך שמולו
 * מאומתת התוצאה שהמנוע מחזיר (`matchCount` בקבלה): אם הם לא שווים, אין
 * לבטוח בתוצאה החלקית עבור אותו בלוק, וההחלפה בו חוזרת לנתיב המדויק-בודד.
 */
export function groupMatchesByBlock(matches: readonly TextMatch[]): BlockMatchCount[] {
  const counts = new Map<string, number>();
  for (const match of matches) {
    counts.set(match.blockId, (counts.get(match.blockId) ?? 0) + 1);
  }
  return [...counts.entries()].map(([blockId, count]) => ({ blockId, count }));
}

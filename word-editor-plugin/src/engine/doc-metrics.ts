/**
 * מה ששורת המצב מודדת: מספר מילים, מספר עמודים ועמוד הסמן.
 *
 * למה המודול הזה נכתב: שלושת הערכים היו `ref` שאותחלו ל-1/1/0 ולא התעדכנו
 * מעולם. שורת המצב הציגה „עמוד 1 מתוך 1” ו„0 מילים” על כל מסמך, כולל מסמך
 * בן שמונים עמודים. זה גרוע יותר משורת מצב ריקה: מספר שנראה כמו מדידה ואינו
 * מדידה.
 *
 * מכאן הכלל שקובע כל החלטה בקובץ: **מה שלא נמדד מוחזר כ-`null`.** אין ערך
 * ברירת מחדל שנראה אמיתי, ואין 1 במקום „לא ידוע”. שכבת התצוגה
 * (composables/shell-format.ts) יודעת לנסח „N עמודים” כשאין עמוד סמן, ולהציג
 * כלום כשאין מדידה בכלל.
 *
 * שלושת המקורות, וכל אחד מהם והמלכודת שלו:
 *
 * 1. **מספר עמודים — `onPaginationUpdate` בקונפיגורציה של המנוע.** אין getter
 *    ציבורי לשאול בו „כמה עמודים יש”; ראו engine/create-editor.ts.
 * 2. **מספר מילים — `doc.info({}).counts.words`.** בדפדפן הפאסדה
 *    א-סינכרונית ועשויה להחזיר הבטחה (`BrowserDocumentApi`), והיא גם עשויה
 *    לזרוק. שתי הצורות מטופלות, ומדידה שנכשלה משאירה את הערך הקודם ואינה
 *    מפילה כלום. הספירה **בהשקטה** ולא בכל הקשה: `info` סורק את המסמך כולו.
 * 3. **עמוד הסמן — `ui.selection.getAnchorRect()?.pageIndex`.** נמדד
 *    ב-superdoc@2.8.0: ה-host מספק `getSelectionAnchorRect`, וה-controller
 *    מחזיר ממנו `ViewportRect` עם `pageIndex` מאופס — כלומר זה עמוד הסמן
 *    האמיתי, מתוך פסי הפריסה. `null` הוא מצב רגיל ולא כשל (אין בחירה, או
 *    שהגיאומטריה עוד לא נפתרה), ואז מוצג „N עמודים” בלבד.
 *
 * ה-adapter מחזיק snapshot ומדווח רק על שינוי אמיתי: שינוי שאינו משנה ערך
 * היה מרנדר את שורת המצב על כל הקשה.
 */
import type { SuperDoc } from 'superdoc';
import type { MaybePromise } from './document-api';

/**
 * השקטת ספירת המילים. ארוכה מזו של החיפוש (250) בכוונה: `doc.info` סורק את
 * המסמך כולו ואינו מחפש התאמה בודדת, והמספר בשורת המצב אינו משהו שהמשתמש
 * ממתין לו בזמן ההקלדה.
 */
export const WORD_COUNT_DEBOUNCE_MS = 700;

/**
 * השקטת קריאת עמוד הסמן. קצרה — היא רק מיפוי של הסמן לפס פריסה, ומספר עמוד
 * שמתעדכן חצי שנייה אחרי הגלילה נראה תקוע.
 */
export const CURRENT_PAGE_DEBOUNCE_MS = 150;

export interface DocMetrics {
  /** מספר המילים במסמך, או `null` כשטרם נמדד. */
  words: number | null;
  /** מספר עמודי הפריסה, או `null` כשהעימוד טרם דיווח. */
  totalPages: number | null;
  /** עמוד הסמן, מבוסס 1. `null` = אין מקור אמין ברגע זה. */
  currentPage: number | null;
}

/** מה שנקרא מ-`doc.info({})`. הכול `unknown` — הערכים נבדקים ולא מונחים. */
export interface DocInfoLike {
  counts?: {
    words?: unknown;
    pages?: unknown;
  };
}

/** מה שנדרש מ-SuperDoc: רק `info` מהפאסדה של המסמך. ראו engine/page-setup.ts. */
export interface DocInfoHost {
  activeEditor?: {
    doc?: {
      info?: (input: Record<string, never>) => MaybePromise<DocInfoLike | null | undefined>;
    } | null;
  } | null;
}

/**
 * ה-union הוא מה שמאפשר להעביר גם את המופע האמיתי וגם כפיל בבדיקות — אותה
 * תבנית כמו `PageSetupTarget`, ומאותה סיבה: השוואה מבנית מול
 * `BrowserDocumentApi` המלא הייתה מחייבת לשכפל כאן את כל הפאסדה.
 */
export type DocInfoTarget = SuperDoc | DocInfoHost | null | undefined;

/**
 * `doc.info({})` של המסמך הפתוח. לעולם אינה זורקת סינכרונית: היעדר `doc`
 * (המסמך עדיין נטען) מוחזר כ-`null`, וזה מצב רגיל ולא כשל.
 */
export function readDocumentInfo(
  host: DocInfoTarget,
): MaybePromise<DocInfoLike | null | undefined> {
  const info = (host as DocInfoHost | null | undefined)?.activeEditor?.doc?.info;
  if (typeof info !== 'function') return null;
  return info({});
}

/** מה שנדרש מה-controller: `selection.getAnchorRect` בלבד. */
export interface AnchorRectSource {
  selection?: {
    getAnchorRect?: (input?: {
      placement?: 'start' | 'end' | 'center';
    }) => { pageIndex?: unknown } | null;
  };
}

/**
 * אינדקס העמוד של הסמן (מאופס), או `null`.
 *
 * `null` הוא מצב רגיל: המנוע מחזיר rect רק כשהבחירה נפתרת לפס פריסה מצויר.
 * לכן אין כאן הודעת כשל — שכבת התצוגה פשוט תציג „N עמודים” בלבד.
 */
export function anchorPageIndex(ui: AnchorRectSource | null | undefined): number | null {
  const read = ui?.selection?.getAnchorRect;
  if (typeof read !== 'function') return null;

  const rect = read.call(ui?.selection, { placement: 'start' });
  const index = rect?.pageIndex;
  return typeof index === 'number' && Number.isFinite(index) && index >= 0 ? index : null;
}

export interface DocMetricsSource {
  /**
   * `superdoc.activeEditor.doc.info({})`. מותר לה להחזיר הבטחה, להחזיר
   * `null` (אין מסמך) או לזרוק.
   */
  readInfo: () => MaybePromise<DocInfoLike | null | undefined>;
  /**
   * `ui.selection.getAnchorRect()?.pageIndex` — אינדקס **מאופס**, או `null`
   * כשאין גיאומטריה. ההמרה ל„עמוד מספר N” נעשית כאן, כדי שאתר הקריאה לא
   * יצטרך לזכור מי מאופס ומי לא.
   */
  readAnchorPageIndex: () => number | null;
  /** נקרא על כל שינוי אמיתי ב-snapshot. */
  onChange: (metrics: DocMetrics) => void;
}

export interface DocMetricsAdapter {
  getState(): DocMetrics;
  /** אחרי שינוי במסמך. סופר מילים בהשקטה. */
  noteDocumentChanged(): void;
  /** אחרי תזוזת סמן או שינוי בחירה. קורא את עמוד הסמן בהשקטה. */
  noteSelectionChanged(): void;
  /** מה-callback של העימוד. ערך שאינו מספר סביר נדחה. */
  notePaginationUpdate(totalPages: unknown): void;
  /** מדידה מיידית, בלי השקטה — לשימוש מיד אחרי שמסמך נפתח. */
  measureNow(): void;
  dispose(): void;
}

/** מה שמוצג לפני שנמדד משהו, וגם כשאין מסמך פתוח. */
export function emptyDocMetrics(): DocMetrics {
  return { words: null, totalPages: null, currentPage: null };
}

/** מספר שלם אי-שלילי, או `null`. ספירת מילים של מסמך ריק היא 0 חוקי. */
export function countValue(value: unknown): number | null {
  const parsed = typeof value === 'number' ? value : Number.NaN;
  if (!Number.isFinite(parsed) || parsed < 0) return null;
  return Math.round(parsed);
}

/** מספר עמודים: שלם חיובי, או `null`. מסמך אינו יכול להיות בן אפס עמודים. */
export function pageCountValue(value: unknown): number | null {
  const count = countValue(value);
  return count !== null && count > 0 ? count : null;
}

/** אינדקס עמוד מאופס → „עמוד מספר N”. `null` נשאר `null`. */
export function pageNumberFromIndex(index: number | null): number | null {
  if (index === null || !Number.isFinite(index) || index < 0) return null;
  return Math.round(index) + 1;
}

export function createDocMetrics(source: DocMetricsSource): DocMetricsAdapter {
  let words: number | null = null;
  let currentPage: number | null = null;
  /** מ-`onPaginationUpdate`. המקור המדויק, כי הוא נורה על כל מעבר פריסה. */
  let paginationPages: number | null = null;
  /**
   * מ-`doc.info().counts.pages`. גיבוי בלבד: התיעוד קובע שהשדה נעדר כשהעימוד
   * אינו פעיל או שהפריסה לא הסתיימה, ולכן הוא אינו יכול להחליף את ה-callback.
   */
  let infoPages: number | null = null;

  let wordsTimer: ReturnType<typeof setTimeout> | undefined;
  let pageTimer: ReturnType<typeof setTimeout> | undefined;
  /**
   * דור הקריאה מ-`info`. קריאה א-סינכרונית שמסתיימת אחרי קריאה חדשה יותר —
   * או אחרי הפירוק — נזרקת: אחרת ספירה של מסמך שנסגר הייתה נכתבת על החדש.
   */
  let generation = 0;
  let disposed = false;

  function snapshot(): DocMetrics {
    return { words, totalPages: paginationPages ?? infoPages, currentPage };
  }

  let published = snapshot();

  /** מדווח רק על שינוי אמיתי: כל הקשה מגיעה לכאן. */
  function publish(): void {
    if (disposed) return;
    const next = snapshot();
    if (
      next.words === published.words &&
      next.totalPages === published.totalPages &&
      next.currentPage === published.currentPage
    ) {
      return;
    }
    published = next;
    source.onChange(next);
  }

  async function readInfo(): Promise<void> {
    const mine = (generation += 1);
    let info: DocInfoLike | null | undefined;

    try {
      info = await source.readInfo();
    } catch (error) {
      // מדידה שנכשלה משאירה את המספר הקודם על המסך ואינה מפילה כלום: ספירת
      // מילים אינה שווה תוסף שנפל, וגם לא „0 מילים” שנראה כמו מסמך ריק.
      console.warn('[otzaria-word] ספירת המילים נכשלה', error);
      return;
    }

    if (disposed || mine !== generation) return;

    const counts = info?.counts;
    words = countValue(counts?.words);
    infoPages = pageCountValue(counts?.pages);
    publish();
  }

  function readCurrentPage(): void {
    let index: number | null = null;
    try {
      index = source.readAnchorPageIndex();
    } catch (error) {
      console.warn('[otzaria-word] קריאת עמוד הסמן נכשלה', error);
    }

    currentPage = pageNumberFromIndex(index);
    publish();
  }

  function scheduleWordCount(): void {
    if (disposed) return;
    clearTimeout(wordsTimer);
    wordsTimer = setTimeout(() => {
      wordsTimer = undefined;
      void readInfo();
    }, WORD_COUNT_DEBOUNCE_MS);
  }

  /**
   * פונקציות ולא מתודות על האובייקט המוחזר: הן נקראות זו מזו, וכל אתר קריאה
   * שהיה מפרק את ה-adapter (`const { noteDocumentChanged } = ...`) היה מאבד
   * את ה-`this`.
   */
  function scheduleCurrentPage(): void {
    if (disposed) return;
    clearTimeout(pageTimer);
    pageTimer = setTimeout(() => {
      pageTimer = undefined;
      readCurrentPage();
    }, CURRENT_PAGE_DEBOUNCE_MS);
  }

  return {
    getState: snapshot,

    noteDocumentChanged() {
      if (disposed) return;
      scheduleWordCount();
      // עריכה מזיזה את הסמן, ולעתים גם את העמוד שהוא נמצא בו.
      scheduleCurrentPage();
    },

    noteSelectionChanged: scheduleCurrentPage,

    notePaginationUpdate(totalPages) {
      if (disposed) return;
      const count = pageCountValue(totalPages);
      // דיווח פגום אינו מוחק מספר תקין שכבר יש: עדיף מספר עמודים מלפני מעבר
      // הפריסה האחרון מאשר שורת מצב שמתרוקנת.
      if (count === null) return;
      paginationPages = count;
      publish();

      // מעבר הפריסה הראשון הוא גם הרגע המוקדם ביותר שבו ספירה מוצלחת סבירה:
      // `measureNow` בפתיחה עשוי ליפול על פאסדת מסמך שעוד לא נבנתה, ובלי
      // הניסיון הזה מספר המילים היה מופיע רק אחרי ההקלדה הראשונה.
      if (words === null) scheduleWordCount();
    },

    measureNow() {
      if (disposed) return;
      void readInfo();
      readCurrentPage();
    },

    dispose() {
      disposed = true;
      // מעלה דור, ולכן קריאת `info` שבאוויר לא תדווח על מסמך שנסגר.
      generation += 1;
      clearTimeout(wordsTimer);
      clearTimeout(pageTimer);
      wordsTimer = undefined;
      pageTimer = undefined;
    },
  };
}

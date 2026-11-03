/**
 * מאזין ייעודי ל-Undo/Redo במקלדת, **בשלב ה-capture** — לפני שהעורך (ProseMirror)
 * מטפל בהם ומבטל את ברירת המחדל.
 *
 * ## למה זה קיים בנפרד מ-`createShortcutDispatcher`
 *
 * `createShortcutDispatcher` (dispatch.ts) מדלג בכוונה על אירוע שכבר
 * `defaultPrevented` — כדי לא להריץ קיצור שהמנוע כבר קשר בעצמו (כמו Ctrl+B)
 * פעמיים. `Ctrl+Z`/`Ctrl+Y`/`Ctrl+Shift+Z` הם בדיוק המקרה הזה: הם ה-`history`
 * המובנה של ProseMirror, קשור על אזור המסמך, ומבטל את ברירת המחדל **לפני**
 * שהאירוע מגיע ל-`window` בשלב ה-bubble ששם `createShortcutDispatcher` יושב.
 *
 * נמדד ישירות (Chrome headless, ה-dist הארוז): לחיצת Ctrl+Z אמיתית עם הפוקוס
 * בתוך המסמך שינתה את ה-DOCX בפועל (המנוע ביצע Undo), אבל `runCommand('undo')`
 * שלנו **לא רץ בכלל** — כלומר כל צרכן שרוצה לדעת „משהו כמו Undo/Redo קרה”
 * לא יכול להסתמך על `createShortcutDispatcher`/`runCommand`.
 *
 * הפתרון: מאזין נפרד, ב-**capture** ולא ב-bubble — כלומר רואה את האירוע
 * **לפני** שהוא מגיע ליעד, ולפני שהמנוע מספיק לבטל אותו. הוא אינו מריץ פקודה
 * ואינו קורא ל-`preventDefault` (זה היה משבש את ה-Undo האמיתי של המנוע) —
 * רק מודיע ש„נלחץ צירוף שנראה כמו Undo” או „כמו Redo”, בלי קשר למי בסוף
 * מטפל בו בפועל (המנוע, או אנחנו כשהפוקוס מחוץ למסמך).
 *
 * ## `isBlocked` — ולמה בלעדיו זה מזיק
 *
 * גרסה ראשונה תפסה **כל** `keydown` תואם ב-`window`, בלי לבדוק את
 * `event.target` בכלל — ולכן Ctrl+Z בתוך שדה טקסט של הממשק (שדה חיפוש,
 * למשל) ניקה את המעקב בלי שום קשר למסמך (נמדד ב-QA, פעמיים בעקביות).
 * `isBlocked` היא אותה בדיקה בדיוק ש-`createShortcutDispatcher` וגם
 * `createDirectionShortcut` כבר עושים לפני שהם מריצים משהו — App.vue מזריק
 * את אותו closure (`isModalDialogOpen() || (isTextEntryTarget(target) &&
 * !isDocumentSurface(target))`) לשלושתם.
 *
 * הצרכן הראשון: `PageBreakTracker` (engine/page-break.ts) — Undo/Redo יכולים
 * לשנות `pageBreakBefore` בלי לעבור דרך הכפתור ברצועה, וזה האיתות היחיד
 * שיש. ראו ההסבר המלא שם, „QA עצמאי” → „Undo/Redo”, כולל למה `onUndo`
 * ו-`onRedo` הם שתי קריאות נפרדות ולא אחת (א-סימטריית התצלום).
 */
import { matchAny, type KeyEventLike } from './match';
import { SHORTCUTS, type Shortcut } from './registry';

/** רשומות הרג'יסטרי, מוקלדות כ-`Shortcut[]` — `SHORTCUTS` עצמו הוא טופל צר מדי לסינון. */
const ALL_SHORTCUTS: readonly Shortcut[] = SHORTCUTS;

/** רשומות הרג'יסטרי ש„נראות כמו” Undo/Redo — נגזר מהרג'יסטרי, לא כפול לו. */
const UNDO_REDO_SHORTCUTS: readonly Shortcut[] = ALL_SHORTCUTS.filter(
  (shortcut) => shortcut.command === 'undo' || shortcut.command === 'redo',
);

export interface UndoRedoWatcher {
  /** מנתקת את המאזין. אידמפוטנטית. */
  dispose(): void;
}

export interface WatchUndoRedoKeysOptions {
  /** נקראת על לחיצה שתואמת רשומת `undo` ברג'יסטרי (Ctrl+Z). */
  onUndo: () => void;
  /** נקראת על לחיצה שתואמת רשומת `redo` (Ctrl+Y, Ctrl+Shift+Z). */
  onRedo: () => void;
  /** ברירת המחדל: `window`. */
  target?: Pick<EventTarget, 'addEventListener' | 'removeEventListener'>;
  /**
   * חוסמת התראה כש-`true` — שדה טקסט של הממשק (לא של המסמך) או דיאלוג
   * מודאלי. ברירת המחדל: אף פעם לא חוסמת (למי שלא מזריק בדיקה, למשל בדיקת
   * יחידה שאין לה מסך אמיתי).
   */
  isBlocked?: (target: EventTarget | null) => boolean;
}

/**
 * רושמת את המאזין. `onUndo`/`onRedo` נקראות פעם אחת לכל לחיצה תואמת ולא
 * חסומה — לפני שהמנוע מספיק לטפל בה, ובלי קשר אם הוא בסוף עשה משהו (לחיצה
 * על Undo במסמך ריק, שאין בו מה לבטל, עדיין מודיעה — ראו „המחיר” בהערת
 * הפתיחה של page-break.ts).
 */
export function watchUndoRedoKeys(options: WatchUndoRedoKeysOptions): UndoRedoWatcher {
  const { onUndo, onRedo, target = window, isBlocked = () => false } = options;

  const listener = (event: Event): void => {
    const shortcut = matchAny(event as unknown as KeyEventLike, UNDO_REDO_SHORTCUTS);
    if (!shortcut) return;
    if (isBlocked((event as { target?: EventTarget | null }).target ?? null)) return;
    if (shortcut.command === 'undo') onUndo();
    else if (shortcut.command === 'redo') onRedo();
  };

  target.addEventListener('keydown', listener, true);

  let disposed = false;
  return {
    dispose() {
      if (disposed) return;
      disposed = true;
      target.removeEventListener('keydown', listener, true);
    },
  };
}

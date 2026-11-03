/**
 * ‏`Ctrl+Backspace` / `Ctrl+Delete` — „לא תמיד מוחק מילה”.
 *
 * ## התסמין
 *
 * אחרי שפקודה ציירה מחדש את הפסקה — „מספור” או „תבליטים” ברצועה, או
 * `lists.create` דרך ה-API — מחיקת מילה אינה עושה דבר, בלי שגיאה ובלי הודעה,
 * עד שהסמן זז (לחיצה, חץ, Home/End). ‏Backspace רגיל באותו מקום עובד, והקלדה
 * אינה מרפאת. בפסקה ממוספרת זה חוזר אחרי כל מחיקה, ולכן מקש מוחזק מחק רק את
 * המילה הראשונה. המשתמש ראה „לפעמים עובד ולפעמים לא” בלי סיבה שאפשר לעלות עליה.
 *
 * באג במנוע: superdoc/docx-editor#4084.
 *
 * ## העקיפה
 *
 * לפני שהמנוע רואה את ההקשה, הסמן נכתב **לאותה נקודה בדיוק** דרך
 * `setSelectionTarget` — אותה פעולה שמזיזה סמן, ולכן מרפאת. ההקשה עצמה אינה
 * מבוטלת: המנוע מוחק כרגיל, עם הסמנטיקה שלו, במעקב שינויים ובהיסטוריה שלו.
 * הקלדה מהירה ומיד Ctrl+Backspace נותנת אותה תוצאה עם העקיפה ובלעדיה.
 *
 * ## מתי היא לא רצה
 *
 * - בחירה שאינה מכווצת: המנוע מוחק את הבחירה ולא מחשב מילה.
 * - ‏Shift/Alt/Meta: המנוע עצמו מחשב שם מחיקת תו, ואין מה לרפא.
 * - הרכבה במקלדת (IME).
 * - אין תצלום סינכרוני או אין `setSelectionTarget` — ההקשה עוברת למנוע כמו
 *   שהיא, כלומר ההתנהגות של היום ולא שבירה.
 * - התצלום מפגר אחרי ההקשה הקודמת (‏`isLaggingSnapshot`), למשל במקש מוחזק.
 *   כתיבה של היסט ישן באמצע פסקה הייתה מזיזה את הסמן ומוחקת מילה אחרת, ולכן
 *   ההקשה הזאת עוברת למנוע כמו שהיא: עדיף להחמיץ הקשה ממחיקה לא נכונה.
 *
 * ## העיצוב הממתין
 *
 * ‏`setSelectionTarget` מאפס את העיצוב הממתין (Ctrl+B בלי בחירה). העקיפה
 * קוראת אותו לפני הכתיבה ומשחזרת מיד אחריה (`setPendingInlineFormat`, אותו
 * מפתח וערך), אחרת Ctrl+B, ‏Ctrl+Backspace והקלדה היו יוצאים לא מודגשים.
 *
 * ## מה היא אינה מרפאת: תא טבלה
 *
 * סמן בסוף תא: Ctrl+Backspace השני והלאה אינו מוחק, עם העקיפה ובלעדיה, ורק
 * ניווט מהמסך (לחיצה, חצים) מרפא. גם זה ב-#4084.
 *
 * כשהמנוע יתוקן, למחוק את המודול הזה ואת השער שלו (`scripts/qa/word-delete-qa.mjs`).
 */

/** נקודת קצה של בחירה, כפי שהמנוע מחזיר אותה. */
interface TextPoint {
  kind?: string;
  blockId?: string;
  offset?: number;
  story?: unknown;
}

/** התצלום הסינכרוני של הבחירה. */
interface LiveSelectionSnapshot {
  selectionTarget?: {
    kind?: string;
    start?: TextPoint | null;
    end?: TextPoint | null;
    story?: unknown;
    coordinateSpace?: string;
  } | null;
}

/** מה שנצרך מ-`activeEditor`. מוגדר כאן ולא מיובא — ההסבר ב-document-api.ts. */
export interface WordDeleteEditor {
  host?: {
    readLiveSelectionSyncSnapshot?: () => LiveSelectionSnapshot | null | undefined;
    /** העיצוב הממתין (Ctrl+B בלי בחירה וכו'); `null` כשאין. */
    getPendingInlineFormat?: () => Record<string, unknown> | null | undefined;
    setPendingInlineFormat?: (key: string, value: unknown) => void;
    /** מונה המוטציות המקומיות; עולה ב-1 בכל commit. */
    getLocalMutationRevision?: () => number;
  } | null;
  authoring?: {
    setSelectionTarget?: (input: {
      target: {
        kind: 'selection';
        start: TextPoint;
        end: TextPoint;
        story: unknown;
        coordinateSpace?: string;
      };
      focus?: boolean;
    }) => unknown;
  } | null;
}

export interface WordDeleteHost {
  activeEditor?: WordDeleteEditor | null;
}

export interface WordDeleteRefreshOptions {
  /** ה-div שהמנוע מצייר בתוכו — `paintedHost(editor.ui)`. */
  host: HTMLElement | null;
  superdoc: WordDeleteHost | null | undefined;
}

export interface WordDeleteRefreshHandle {
  dispose(): void;
}

/**
 * ההקשה שהמנוע מתרגם למחיקת מילה: Backspace או Delete עם Ctrl **בלבד**.
 * זה התנאי של המנוע עצמו (Ctrl בלי Meta, Alt ו-Shift); כל צירוף אחר נמחק
 * אצלו כתו, ואינו עובר במסלול השבור.
 */
export function isWordDeleteKey(
  event: Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey'> &
    Partial<Pick<KeyboardEvent, 'isComposing' | 'keyCode'>>,
): boolean {
  if (event.key !== 'Backspace' && event.key !== 'Delete') return false;
  if (event.isComposing === true || event.keyCode === 229) return false;
  return event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey;
}

/** הסמן, כשהבחירה מכווצת בנקודת טקסט אחת; אחרת `null`. */
export function collapsedCaret(snapshot: LiveSelectionSnapshot | null | undefined): TextPoint | null {
  const selection = snapshot?.selectionTarget;
  const start = selection?.start;
  const end = selection?.end;
  if (!start || !end || start.kind !== 'text' || end.kind !== 'text') return null;
  if (typeof start.blockId !== 'string' || typeof start.offset !== 'number') return null;
  if (start.blockId !== end.blockId || start.offset !== end.offset) return null;
  return start;
}

/** הנקודה שנכתבה בהקשה הקודמת, ומונה המוטציות באותו רגע. */
export interface LastWrite {
  blockId: string;
  offset: number;
  revision: number | undefined;
  direction: 'backward' | 'forward';
}

/**
 * התצלום הסינכרוני עדיין מראה את הנקודה שנכתבה בהקשה הקודמת, והמחיקה שהיא
 * אפשרה טרם השתקפה בו — כלומר הוא מפגר, והסמן האמיתי כבר לפניה.
 *
 * מונה המוטציות מבדיל בין זה לבין חזרה אמיתית לאותה נקודה: בפיגור הוא זז
 * לכל היותר ב-1 (המחיקה עצמה, אם כבר נקלטה); כדי לחזור לאותו היסט אחרי
 * מחיקה אחורה צריך לפחות פעולה אחת נוספת (הקלדה, ביטול).
 * במחיקה קדימה ההיסט נשאר במקומו: אם המוטציה שלה כבר נקלטה, אותה נקודה
 * היא סמן תקין. הכיוון הוא של ההקשה הקודמת, גם אם המשתמש החליף כיוון כעת.
 */
export function isLaggingSnapshot(last: LastWrite | null, caret: TextPoint, revision: number | undefined): boolean {
  if (!last || last.blockId !== caret.blockId || last.offset !== caret.offset) return false;
  if (typeof revision !== 'number' || typeof last.revision !== 'number') return true;
  const mutations = revision - last.revision;
  if (last.direction === 'forward' && mutations === 1) return false;
  return mutations <= 1;
}

/** מתקין את העקיפה. המצב היחיד הוא הכתיבה האחרונה, והיא נפרקת עם המסמך. */
export function installWordDeleteRefresh({ host, superdoc }: WordDeleteRefreshOptions): WordDeleteRefreshHandle {
  if (!host) return { dispose() {} };

  let lastWrite: LastWrite | null = null;

  const onKeyDown = (event: KeyboardEvent): void => {
    if (!isWordDeleteKey(event)) return;

    const editor = superdoc?.activeEditor;
    const engineHost = editor?.host;
    const setSelectionTarget = editor?.authoring?.setSelectionTarget;
    const readSnapshot = engineHost?.readLiveSelectionSyncSnapshot;
    if (typeof setSelectionTarget !== 'function' || typeof readSnapshot !== 'function') return;

    let snapshot: LiveSelectionSnapshot | null | undefined;
    let revision: number | undefined;
    let pending: Record<string, unknown> | null | undefined;
    try {
      snapshot = readSnapshot.call(engineHost);
      revision = engineHost?.getLocalMutationRevision?.();
      pending = engineHost?.getPendingInlineFormat?.();
    } catch {
      return;
    }

    const caret = collapsedCaret(snapshot);
    if (!caret) return;

    // הקשה חוזרת (מקש מוחזק) לפני שהתצלום התעדכן: כתיבה
    // של ההיסט הישן אחרי מחיקה אחורה הייתה מזיזה את הסמן קדימה, אל מעבר למה שכבר נמחק —
    // ובאמצע פסקה זו מחיקה של המילה הלא נכונה. עדיף להחמיץ הקשה.
    if (isLaggingSnapshot(lastWrite, caret, revision)) return;

    const selection = snapshot?.selectionTarget;
    const story = selection?.story ?? caret.story;
    const coordinateSpace = selection?.coordinateSpace;
    const point: TextPoint = { kind: 'text', blockId: caret.blockId, offset: caret.offset, story };

    try {
      const written = setSelectionTarget.call(editor?.authoring, {
        target: {
          kind: 'selection',
          start: point,
          end: point,
          story,
          ...(typeof coordinateSpace === 'string' ? { coordinateSpace } : {}),
        },
        focus: true,
      });
      void Promise.resolve(written).catch(() => {});
      lastWrite = {
        blockId: caret.blockId as string,
        offset: caret.offset as number,
        revision,
        direction: event.key === 'Delete' ? 'forward' : 'backward',
      };
    } catch {
      /* סמן שלא נכתב מחדש אינו סיבה לעצור את המחיקה — המנוע עדיין מקבל אותה */
    }

    // כתיבת בחירה מאפסת את העיצוב הממתין. ‏Ctrl+B ואחריו Ctrl+Backspace והקלדה —
    // הטקסט חייב לצאת מודגש, כמו בלי העקיפה.
    if (pending && typeof engineHost?.setPendingInlineFormat === 'function') {
      for (const [key, value] of Object.entries(pending)) {
        try {
          engineHost.setPendingInlineFormat(key, value);
        } catch {
          /* מפתח שהמנוע אינו מקבל — השאר משוחזרים */
        }
      }
    }
  };

  host.addEventListener('keydown', onKeyDown, true);
  return {
    dispose() {
      host.removeEventListener('keydown', onKeyDown, true);
    },
  };
}

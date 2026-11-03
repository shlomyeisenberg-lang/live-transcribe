/**
 * תפיסת הבחירה לפני שדיאלוג גוזל את המיקוד.
 *
 * הבדיקות כאן מכסות בדיוק את המצבים שהופכים כפתור לשבור בשקט: מסמך שעדיין
 * נטען, מנוע בגרסה שאינה חושפת `selection`, קריאה שזורקת, ותשובה שאינה
 * אובייקט. בכל אחד מהם התשובה חייבת להיות „אין בחירה” ולא זריקה — פקד ברצועה
 * שזורק מפיל את רינדור הרצועה כולה.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  emptySelectionSnapshot,
  readDocSelection,
  type SelectionDocumentApi,
  type SelectionInfoLike,
} from '../../src/engine/doc-selection';

/**
 * מופע כפול: רק המשטח ש-`readDocSelection` נוגע בו. `current` מתקבל כ-`unknown`
 * בכוונה — חלק מהבדיקות מוסרות לו תשובות שאינן בחוזה (מחרוזת, מספר), וזה בדיוק
 * מה שהן בודקות.
 */
function hostWith(current: unknown) {
  const doc = { selection: { current } } as unknown as SelectionDocumentApi;
  return { activeEditor: { doc } };
}

function hostReturning(info: SelectionInfoLike | undefined) {
  const current = vi.fn(async () => info);
  return { host: hostWith(current), current };
}

/**
 * שני השדות יחד, כמו שהמנוע מחזיר. הפיקסטורה שהייתה כאן נשאה `target` בלבד,
 * וזה בדיוק מה שאיפשר ל-`insertCitation` לשלוח את הצורה הלא נכונה בלי שאף
 * בדיקה תרגיש: כפיל שאינו מדמה את התשובה המלאה מאשר גם קוד ששואל את השדה
 * הלא נכון.
 */
const RANGE_INFO: SelectionInfoLike = {
  empty: false,
  target: { kind: 'text', segments: [{ blockId: 'p7', range: { start: 3, end: 9 } }] },
  selectionTarget: {
    kind: 'selection',
    start: { kind: 'text', blockId: 'p7', offset: 3 },
    end: { kind: 'text', blockId: 'p7', offset: 9 },
  },
  text: 'בראשית',
};

const CARET_INFO: SelectionInfoLike = {
  empty: true,
  target: { kind: 'text', segments: [{ blockId: 'p7', range: { start: 4, end: 4 } }] },
};

describe('readDocSelection', () => {
  it('טווח מסומן מדווח hasRange, טקסט ו-blockId', () => {
    const { host } = hostReturning(RANGE_INFO);

    return expect(readDocSelection(host, { includeText: true })).resolves.toEqual({
      target: RANGE_INFO.target,
      selectionTarget: RANGE_INFO.selectionTarget,
      text: 'בראשית',
      hasRange: true,
      blockId: 'p7',
      blockIds: ['p7'],
      story: null,
      empty: false,
    });
  });

  /**
   * שני השדות נמסרים בנפרד, כי הצרכנים שלהם שונים: `hyperlinks.wrap` מקבל את
   * רשימת הקטעים, ו-`doc.insert` מקבל **רק** את ה-SelectionTarget. ערבוב
   * ביניהם נכשל סגור עם `target must be a SelectionTarget object.`
   */
  it('שני היעדים נמסרים בנפרד ואינם מתערבבים', async () => {
    const snapshot = await readDocSelection(hostReturning(RANGE_INFO).host);

    expect(snapshot.target).toEqual(RANGE_INFO.target);
    expect(snapshot.selectionTarget).toEqual(RANGE_INFO.selectionTarget);
    expect(snapshot.selectionTarget).not.toEqual(snapshot.target);
  });

  it('selectionTarget שאינו אובייקט נדחה ל-null', async () => {
    // מסירת ערך כזה ל-`insert` הייתה זריקת INVALID_INPUT במקום נפילה חזרה
    // לסוף המסמך.
    for (const bad of [null, undefined, 'selection', 7]) {
      const snapshot = await readDocSelection(
        hostReturning({ ...RANGE_INFO, selectionTarget: bad } as SelectionInfoLike).host,
      );
      expect(snapshot.selectionTarget).toBeNull();
    }
  });

  it('סמן בלבד אינו טווח — וזה מה שקובע את המסלול במנוע', async () => {
    // `hasRange: false` = `hyperlinks.insert`, `true` = `hyperlinks.wrap`.
    const snapshot = await readDocSelection(hostReturning(CARET_INFO).host);

    expect(snapshot.hasRange).toBe(false);
    expect(snapshot.blockId).toBe('p7');
    expect(snapshot.text).toBe('');
    expect(snapshot.empty).toBe(true);
  });

  it('בחירת אובייקט (תא טבלה/תמונה): target ריק אבל יש בחירה — `empty` מדווח את זה גם כש-hasRange ו-text שקריים', async () => {
    // `SelectionInfo.empty` במנוע: "True when the selection is empty (cursor
    // only, no range)" — ולא רק על טווח טקסט שאפשר לעטוף. בחירת אובייקט אינה
    // מייצרת target/segments בכלל, ולכן hasRange ו-text נשארים false.
    const { host } = hostReturning({ empty: false, target: null });

    const snapshot = await readDocSelection(host);

    expect(snapshot.empty).toBe(false);
    expect(snapshot.hasRange).toBe(false);
    expect(snapshot.text).toBe('');
  });

  it('מנוע שאינו מדווח `empty` — נופלים לאחור ל-hasRange||text', async () => {
    const { host: withRange } = hostReturning({
      target: { kind: 'text', segments: [{ blockId: 'p7', range: { start: 3, end: 9 } }] },
    });
    expect((await readDocSelection(withRange)).empty).toBe(false);

    const { host: withoutRange } = hostReturning({ target: null });
    expect((await readDocSelection(withoutRange)).empty).toBe(true);
  });

  it('מספר קטעים — די באחד עם טווח', async () => {
    const { host } = hostReturning({
      target: {
        kind: 'text',
        segments: [
          { blockId: 'p1', range: { start: 5, end: 5 } },
          { blockId: 'p2', range: { start: 0, end: 4 } },
        ],
      },
    });

    const snapshot = await readDocSelection(host);

    expect(snapshot.hasRange).toBe(true);
    // ה-blockId הוא של הקטע הראשון — שם הבחירה מתחילה.
    expect(snapshot.blockId).toBe('p1');
  });

  it('`includeText` נשלח רק כשהוא מבוקש — החילוץ עולה בביצועים', async () => {
    const { host, current } = hostReturning(RANGE_INFO);

    await readDocSelection(host);
    expect(current).toHaveBeenCalledWith(undefined);

    await readDocSelection(host, { includeText: true });
    expect(current).toHaveBeenLastCalledWith({ includeText: true });
  });

  it('ה-story נשמר — בלעדיו הקישור נכתב בגוף במקום בכותרת', async () => {
    const story = { kind: 'story', storyType: 'headerFooterSlot' };
    const { host } = hostReturning({
      target: { kind: 'text', segments: [{ blockId: 'h1', range: { start: 0, end: 2 } }], story },
    });

    await expect(readDocSelection(host)).resolves.toMatchObject({ story });
  });

  it('בחירה בלי קטעים אינה מוסרת target', async () => {
    // `{segments: []}` הוא אובייקט, ולכן היה עובר את
    // `linkPayloadHasExplicitTarget` ואז מייצר רשימת כתובות ריקה — פקודה
    // שנראית מוכנה ונכשלת סגור.
    const { host } = hostReturning({ target: { kind: 'text', segments: [] } });

    await expect(readDocSelection(host)).resolves.toEqual(emptySelectionSnapshot());
  });

  it('target שהוא null אינו מוסר', async () => {
    const { host } = hostReturning({ empty: true, target: null });

    await expect(readDocSelection(host)).resolves.toEqual(emptySelectionSnapshot());
  });

  it('קטע בלי blockId אינו נחשב', async () => {
    const { host } = hostReturning({
      target: { kind: 'text', segments: [{ range: { start: 0, end: 3 } }] },
    });

    await expect(readDocSelection(host)).resolves.toEqual(emptySelectionSnapshot());
  });

  it('אין Document API — „אין בחירה”, לא זריקה', async () => {
    await expect(readDocSelection(null)).resolves.toEqual(emptySelectionSnapshot());
    await expect(readDocSelection(undefined)).resolves.toEqual(emptySelectionSnapshot());
    await expect(readDocSelection({ activeEditor: null })).resolves.toEqual(
      emptySelectionSnapshot(),
    );
    await expect(readDocSelection({ activeEditor: { doc: null } })).resolves.toEqual(
      emptySelectionSnapshot(),
    );
  });

  it('גרסה שאינה חושפת selection אינה מפילה את הפקד', async () => {
    await expect(readDocSelection({ activeEditor: { doc: {} } })).resolves.toEqual(
      emptySelectionSnapshot(),
    );
    await expect(
      readDocSelection({ activeEditor: { doc: { selection: {} } } }),
    ).resolves.toEqual(emptySelectionSnapshot());
  });

  it('קריאה שזורקת נבלעת לטובת „אין בחירה”, ומדווחת ללוג', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const host = hostWith(
      vi.fn(async () => {
        throw new Error('boom');
      }),
    );

    await expect(readDocSelection(host)).resolves.toEqual(emptySelectionSnapshot());
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
  });

  it('תשובה שאינה אובייקט אינה תשובה', async () => {
    for (const value of [undefined, null, 'ready', 42] as unknown[]) {
      const host = hostWith(vi.fn(async () => value));
      await expect(readDocSelection(host)).resolves.toEqual(emptySelectionSnapshot());
    }
  });

  it('תשובה סינכרונית מטופלת כמו הבטחה', async () => {
    // הפאסדה בדפדפן א-סינכרונית, אבל החוזה מצהיר על החזרה סינכרונית.
    const host = hostWith(vi.fn(() => RANGE_INFO));

    await expect(readDocSelection(host)).resolves.toMatchObject({ blockId: 'p7' });
  });

  it('`emptySelectionSnapshot` מחזירה מופע חדש בכל קריאה', () => {
    // קבוע משותף היה מאפשר לצרכן אחד לשנות את התשובה של כולם.
    expect(emptySelectionSnapshot()).not.toBe(emptySelectionSnapshot());
    expect(emptySelectionSnapshot()).toEqual(emptySelectionSnapshot());
  });
});

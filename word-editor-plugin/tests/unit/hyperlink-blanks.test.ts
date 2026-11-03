/**
 * ניקוי צומתי קישור ריקים.
 *
 * הבדיקה מקבעת את מה שנמדד על המנוע: מחיקת הטקסט של קישור משאירה את הצומת
 * עם עוגן באורך אפס, הוא מגיע **גם** ב-`items[]` (בלי `text`) **וגם**
 * ב-`stories[].hyperlinks[]` (עם `text: ""`), ו-`remove` מקבל את `address`
 * כפי שהוא. מה שנשען על „ריק = עוגן באורך אפס” ולא על `text`, כי בצורה אחת
 * השדה פשוט נעדר.
 */
import { describe, expect, it, vi } from 'vitest';
import { removeBlankHyperlinks } from '../../src/engine/hyperlink-blanks';

function anchorAt(blockId: string, start: number, end: number) {
  return { start: { blockId, offset: start }, end: { blockId, offset: end } };
}

function address(blockId: string, start: number, end: number) {
  return { kind: 'inline' as const, nodeType: 'hyperlink' as const, anchor: anchorAt(blockId, start, end) };
}

/** התשובה כפי שהיא נמדדה: אותם קישורים בשתי הצורות. */
function listReply(links: Array<{ text?: string; address: ReturnType<typeof address> }>) {
  return {
    stories: [{ storyId: 'main', hyperlinks: links }],
    items: links.map((link) => ({ id: 'x', address: link.address, properties: { href: 'otzaria://x' } })),
    total: links.length,
  };
}

function fakeDoc(links: Array<{ text?: string; address: ReturnType<typeof address> }>, removeResult: unknown = { success: true }) {
  const remove = vi.fn(() => removeResult as never);
  return {
    remove,
    doc: { hyperlinks: { list: () => listReply(links), remove } },
  };
}

describe('removeBlankHyperlinks', () => {
  it('מסיר צומת ריק, פעם אחת בלבד', async () => {
    const blank = { text: '', address: address('b1', 8, 8) };
    const { doc, remove } = fakeDoc([blank]);

    expect(await removeBlankHyperlinks(doc)).toBe(1);
    // אותו צומת מגיע גם מ-`items` וגם מ-`stories` — הסרה כפולה היא באג.
    expect(remove).toHaveBeenCalledTimes(1);
    expect(remove).toHaveBeenCalledWith({ target: blank.address });
  });

  /**
   * נמדד: מחיקת הטקסט של שני קישורים באותה פסקה משאירה שני צמתים שונים על
   * **אותו עוגן בדיוק** (‏`b1:0-0` פעמיים), וגם `hyperlinkNodeId` שלהם זהה.
   * סינון כפילויות לפי מפתח כלשהו היה מסיר אחד ומשאיר את השני בקובץ.
   */
  it('שני צמתים ריקים על אותו עוגן מוסרים שניהם', async () => {
    const { doc, remove } = fakeDoc([
      { text: '', address: address('b1', 0, 0) },
      { text: '', address: address('b1', 0, 0) },
    ]);

    expect(await removeBlankHyperlinks(doc)).toBe(2);
    expect(remove).toHaveBeenCalledTimes(2);
  });

  it('קורא מקור אחד — `items` אינו מכפיל את `stories`', async () => {
    // הכפיל מחזיר את שניהם, כמו המנוע; ספירה של 2 כאן הייתה אומרת שהאיחוד חזר.
    const { doc, remove } = fakeDoc([{ text: '', address: address('b1', 5, 5) }]);

    expect(await removeBlankHyperlinks(doc)).toBe(1);
    expect(remove).toHaveBeenCalledTimes(1);
  });

  it('נופל ל-items כשאין stories', async () => {
    const remove = vi.fn(() => ({ success: true }) as never);
    const doc = {
      hyperlinks: {
        list: () => ({ items: [{ address: address('b1', 2, 2) }] }),
        remove,
      },
    };
    expect(await removeBlankHyperlinks(doc)).toBe(1);
    expect(remove).toHaveBeenCalledTimes(1);
  });

  it('אינו נוגע בקישור שיש בו טקסט', async () => {
    const { doc, remove } = fakeDoc([{ text: 'פסחים דף לד', address: address('b1', 8, 19) }]);

    expect(await removeBlankHyperlinks(doc)).toBe(0);
    expect(remove).not.toHaveBeenCalled();
  });

  /**
   * המסמך ולא הבלוק: קישור ריק בפסקה אחרת הוא בדיוק הקישור הבלתי-נראה שמגיע
   * לקובץ שיוצא, והקריאה ל-`list` ממילא מחזירה את כולם.
   */
  it('מסיר צומת ריק גם בבלוק אחר — הסריקה היא על המסמך', async () => {
    const { doc, remove } = fakeDoc([{ text: '', address: address('b2', 0, 0) }]);

    expect(await removeBlankHyperlinks(doc)).toBe(1);
    expect(remove).toHaveBeenCalledTimes(1);
  });

  it('במסלול כתיבה ממוקד מסיר רק את הבלוק המבוקש', async () => {
    const { doc, remove } = fakeDoc([
      { text: '', address: address('b1', 0, 0) },
      { text: '', address: address('b2', 0, 0) },
    ]);

    expect(await removeBlankHyperlinks(doc, { blockId: 'b1' })).toBe(1);
    expect(remove).toHaveBeenCalledTimes(1);
    expect(remove).toHaveBeenCalledWith({ target: address('b1', 0, 0) });
  });

  it('שני צמתים ריקים באותו היסט בבלוקים שונים אינם מתבלבלים', async () => {
    const { doc, remove } = fakeDoc([
      { text: '', address: address('b1', 0, 0) },
      { text: '', address: address('b2', 0, 0) },
    ]);

    expect(await removeBlankHyperlinks(doc)).toBe(2);
    expect(remove).toHaveBeenCalledTimes(2);
  });

  it('מזהה ריקנות מהעוגן גם כשאין שדה text', async () => {
    // זו בדיוק הצורה של `items[]`: על צומת ריק השדה נעדר לגמרי.
    const { doc, remove } = fakeDoc([{ address: address('b1', 3, 3) }]);

    expect(await removeBlankHyperlinks(doc)).toBe(1);
    expect(remove).toHaveBeenCalledTimes(1);
  });

  it('מסיר כמה צמתים ריקים באותו בלוק', async () => {
    const { doc, remove } = fakeDoc([
      { text: '', address: address('b1', 4, 4) },
      { text: 'קיים', address: address('b1', 10, 14) },
      { text: '', address: address('b1', 20, 20) },
    ]);

    expect(await removeBlankHyperlinks(doc)).toBe(2);
    expect(remove).toHaveBeenCalledTimes(2);
  });

  it('כשל של remove אינו נספר ואינו זורק', async () => {
    const { doc } = fakeDoc([{ text: '', address: address('b1', 8, 8) }], {
      success: false,
      failure: { code: 'TARGET_NOT_FOUND' },
    });

    expect(await removeBlankHyperlinks(doc)).toBe(0);
  });

  /**
   * ‏`remove` שזורק אינו „הצומת הזה נשאר” אלא „הניקוי נעצר”: בלי ה-catch
   * הראשון שנכשל מפיל את כל הלולאה, והצומת שבאמת חוסם את הכתיבה — זה שבפסקה
   * שמתייגים בה — נשאר במקומו. הענף הזה שרד סבב מוטציה בלי הבדיקה הזאת.
   */
  it('זריקה של remove אינה עוצרת את הניקוי ואינה נספרת', async () => {
    let call = 0;
    const remove = vi.fn(() => {
      if (++call === 1) throw new Error('busy');
      return { success: true } as never;
    });
    const links = [
      { text: '', address: address('b1', 0, 0) },
      { text: '', address: address('b2', 4, 4) },
    ];
    const doc = { hyperlinks: { list: () => listReply(links), remove } };

    expect(await removeBlankHyperlinks(doc)).toBe(1);
    expect(remove).toHaveBeenCalledTimes(2);
  });

  it('זריקה של list נבלעת — היא לא החמירה את המצב', async () => {
    const doc = {
      hyperlinks: {
        list: () => {
          throw new Error('busy');
        },
        remove: vi.fn(),
      },
    };
    expect(await removeBlankHyperlinks(doc)).toBe(0);
  });

  it('מנוע בלי הידיות אינו נוגע בכלום', async () => {
    expect(await removeBlankHyperlinks(null)).toBe(0);
    expect(await removeBlankHyperlinks({})).toBe(0);
    expect(await removeBlankHyperlinks({ hyperlinks: null })).toBe(0);
  });
});

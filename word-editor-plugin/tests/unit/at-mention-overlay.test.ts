/**
 * ה-overlay עצמו, עם כפיל `doc` ו-container אמיתי ב-jsdom. מה שנבדק כאן הוא
 * מה שאפשר לבדוק בלי מנוע: הטווח שנשלח לכתיבה, שני מסלולי הכתיבה, וניווט
 * המקלדת. המראה והמיקום נבדקים בשער ה-QA מול Chrome אמיתי.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { PASTE_STARTED, PASTE_FINISHED } from '../../src/engine/paste-events';
import { installAtMention } from '../../src/engine/at-mention-overlay';
import type { ResolvedRefHit } from '../../src/types/otzaria_plugin';

const { resolveRefMock } = vi.hoisted(() => ({ resolveRefMock: vi.fn() }));

vi.mock('../../src/host/otzaria-reader', () => ({
  resolveRef: resolveRefMock,
}));

function hit(overrides: Partial<ResolvedRefHit> = {}): ResolvedRefHit {
  return {
    id: 42,
    bookId: 'פסחים',
    bookUid: 'id:42',
    title: 'פסחים',
    reference: 'פסחים דף לד',
    index: 1234,
    isPdf: false,
    isSourceLine: true,
    isUserBook: false,
    bookPath: 'ש"ס, בבלי',
    ...overrides,
  };
}

const BLOCK = 'b1';

/**
 * כפיל מנוע: הסמן יושב בסוף `text`, וחלון הקריאה מחזיר את הטקסט כולו.
 *
 * שתי התנהגויות כאן אינן נוחות אלא **מדודות**, והבדיקות נשענות עליהן:
 *
 * - **`insert` אינו מזיז את הסמן.** נמדד: כתיבת „פסחים דף לה” בהיסט 24
 *   השאירה את הסמן ב-24. כפיל שמזיז אותו לסוף היה מסתיר בדיוק את הבאג
 *   שבגללו כל הקלדה נוספת נדחפה לפני הקישור.
 * - **הקישורים מוחזרים גם ב-`stories[]` וגם ב-`items[]`.** זו הצורה שנמדדה,
 *   וממנה נגזרת הדרישה ש-`removeBlankHyperlinks` לא יסיר את אותו צומת פעמיים.
 */
function fakeDoc(
  text: string,
  options: {
    docInsert?: () => unknown;
    wrap?: () => unknown;
    blanks?: Array<{ start: number; end: number; blockId?: string }>;
  } = {},
) {
  const calls = new Map<string, unknown[]>();
  /** סדר הקריאות, ולא רק כמותן — יש ענפים שבהם הסדר הוא כל ההבדל. */
  const order: string[] = [];
  const record = (name: string, input: unknown) => {
    order.push(name);
    calls.set(name, [...(calls.get(name) ?? []), input]);
  };

  let cursor = text.length;
  const blanks = [...(options.blanks ?? [])];
  const addressOf = (b: { start: number; end: number; blockId?: string }) => ({
    kind: 'inline',
    nodeType: 'hyperlink',
    anchor: {
      start: { blockId: b.blockId ?? BLOCK, offset: b.start },
      end: { blockId: b.blockId ?? BLOCK, offset: b.end },
    },
  });

  const hyperlinks: Record<string, unknown> = {
    wrap: (input: unknown) => {
      record('hyperlinks.wrap', input);
      return options.wrap?.() ?? { success: true };
    },
    list: () => {
      record('hyperlinks.list', null);
      const items = blanks.map((b) => ({ text: '', address: addressOf(b) }));
      return { stories: [{ storyId: 'main', hyperlinks: items }], items };
    },
    remove: (input: unknown) => {
      record('hyperlinks.remove', input);
      const { target } = input as { target: { anchor: { start: { offset: number } } } };
      const at = target.anchor.start.offset;
      const index = blanks.findIndex((b) => b.start === at);
      if (index >= 0) blanks.splice(index, 1);
      return { success: true };
    },
  };

  const doc = {
    selection: {
      current: async () => ({
        selectionTarget: {
          kind: 'selection',
          start: { kind: 'text', blockId: BLOCK, offset: cursor },
          end: { kind: 'text', blockId: BLOCK, offset: cursor },
        },
      }),
    },
    ranges: {
      resolve: async () => ({
        preview: { text, truncated: false },
        target: { start: { offset: 0 }, end: { offset: cursor } },
      }),
    },
    insert: (input: unknown) => {
      record('insert', input);
      const { target } = input as { target: { start: { offset: number } } };
      // נמדד: הסמן נשאר בתחילת הטווח שנכתב, ואינו זז לסופו.
      cursor = target.start.offset;
      return options.docInsert?.() ?? { success: true };
    },
    hyperlinks,
  };

  const applied: unknown[] = [];
  const host = {
    activeEditor: { doc },
    ui: {
      selection: {
        getAnchorRect: () => ({ left: 200, top: 100, width: 1, height: 18 }),
        apply: (target: unknown) => {
          applied.push(target);
          record('selection.apply', target);
          return { ok: true };
        },
      },
    },
  };
  return { host, calls, applied, order };
}

/** מריצה את מחזור ה-debounce וההערכה עד שהרשימה מצוירת. */
async function settle(): Promise<void> {
  await vi.advanceTimersByTimeAsync(200);
  await vi.advanceTimersByTimeAsync(0);
}

function popup(): HTMLElement | null {
  return document.querySelector('.otzaria-at-mention');
}

function options(): HTMLElement[] {
  return [...(popup()?.querySelectorAll('[role="option"]') ?? [])] as HTMLElement[];
}

let container: HTMLElement;

beforeEach(() => {
  vi.useFakeTimers();
  resolveRefMock.mockReset();
  resolveRefMock.mockResolvedValue({ ok: true, value: [hit()] });
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  vi.useRealTimers();
  container.remove();
  popup()?.remove();
});

describe('installAtMention', () => {
  it('פותח רשימה על אזכור, ושולח לאוצריא את ההפניה בלבד', async () => {
    const { host } = fakeDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    expect(resolveRefMock).toHaveBeenCalledWith('פסחים לד', 8);
    expect(options()).toHaveLength(1);
    expect(options()[0]!.textContent).toContain('פסחים דף לד');
    handle.dispose();
  });

  it('אינו נפתח על טקסט שאינו אזכור', async () => {
    const { host } = fakeDoc('סתם טקסט בלי סימן');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    expect(resolveRefMock).not.toHaveBeenCalled();
    expect(popup()).toBeNull();
    handle.dispose();
  });

  it('אינו נפתח כשאין התאמות', async () => {
    resolveRefMock.mockResolvedValue({ ok: true, value: [] });
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    expect(popup()).toBeNull();
    handle.dispose();
  });

  it('מחליף את הטווח מה-@ ועד הסמן, כותב את הטקסט ועוטף אותו', async () => {
    const { host, calls } = fakeDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const inserts = calls.get('insert') as Array<{
      value: string;
      target: { start: { offset: number }; end: { offset: number } };
    }>;
    // "ראה " הוא 4 תווים, ולכן ה-@ יושב ב-4 והסמן ב-13.
    expect(inserts[0]!.value).toBe('');
    expect(inserts[0]!.target.start.offset).toBe(4);
    expect(inserts[0]!.target.end.offset).toBe(13);

    // ואז הטקסט הנראה נכתב בנקודה שנפתחה.
    expect(inserts[1]!.value).toBe('פסחים דף לד');
    expect(inserts[1]!.target).toMatchObject({ start: { offset: 4 }, end: { offset: 4 } });

    // ורק אז הוא נעטף — `wrap` ולא `insert`, ראו הערת המודול.
    const wrapped = calls.get('hyperlinks.wrap')?.[0] as {
      target: { blockId: string; range: { start: number; end: number } };
      link: { destination: { href: string } };
    };
    expect(wrapped.target).toMatchObject({ blockId: BLOCK, range: { start: 4, end: 15 } });
    expect(wrapped.link.destination.href).toBe('otzaria://open/book/42?index=1234&uid=id%3A42');
    handle.dispose();
  });

  /**
   * הבאג: `hyperlinks.insert` השאירה את הסמן לפני הקישור, ולכן „ראה @פסחים לד”
   * ואז „ וכן …” יצא „ראה  וכן …פסחים דף לד” — ההמשך נדחף לפני הקישור.
   */
  it('הסמן מוצב אחרי הקישור, כדי שההקלדה הבאה תמשיך ממנו', async () => {
    const { host, applied } = fakeDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    // 4 (תחילת האזכור) + 11 (אורך „פסחים דף לד”).
    expect(applied).toEqual([
      {
        kind: 'selection',
        start: { kind: 'text', blockId: BLOCK, offset: 15 },
        end: { kind: 'text', blockId: BLOCK, offset: 15 },
      },
    ]);
    handle.dispose();
  });

  /**
   * הבאג: מחיקת קישור משאירה צומת ריק, והוא חוסם כתיבה חדשה באותו מקום
   * ב-`hyperlink-nested-unsupported` — על קישור שכבר לא רואים.
   */
  it('מנקה לפני הכתיבה רק את הבלוק החוסם, ואת שאר המסמך בזמן שקט', async () => {
    const { host, calls, order } = fakeDoc('@פסחים לד', {
      blanks: [
        { start: 0, end: 0 },
        { start: 4, end: 4, blockId: 'b2' },
      ],
    });
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const removedStarts = (): Array<{ blockId: string; offset: number }> =>
      (
        (calls.get('hyperlinks.remove') ?? []) as Array<{
          target: { anchor: { start: { blockId: string; offset: number } } };
        }>
      ).map((r) => r.target.anchor.start);
    // לפני הכתיבה אין להמתין על קישור ריק בפסקה אחרת: רק הריק שחוסם כאן
    // נדרש כדי ש-wrap תוכל להצליח.
    expect(removedStarts()).toEqual([
      { blockId: BLOCK, offset: 0 },
    ]);

    // והניקוי קודם ל**מחיקה**, לא רק לכתיבה: מרגע המחיקה הטקסט של המשתמש
    // אינו במסמך, וקריאת מנוע שאינה חוזרת בחלון הזה מאבדת אותו.
    expect(order.indexOf('hyperlinks.remove')).toBeLessThan(order.indexOf('insert'));
    expect(calls.get('hyperlinks.wrap')).toHaveLength(1);

    // הניקוי הרחב נשמר, אך אינו נועל את מחוות הבחירה: אחרי השהיה אחת הוא
    // מסיר גם את השארית הרחוקה, ובמספר אזכורים רצופים זו נשארת סריקה אחת.
    await vi.advanceTimersByTimeAsync(750);
    expect(removedStarts()).toEqual([
      { blockId: BLOCK, offset: 0 },
      { blockId: 'b2', offset: 4 },
    ]);
    handle.dispose();
  });

  it('אות שימוש נשארת במסמך — מוחלף רק מה-@', async () => {
    const { host, calls } = fakeDoc('כמובא ב@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const deleted = calls.get('insert')?.[0] as { target: { start: { offset: number } } };
    // "כמובא ב" הוא 7 תווים; המחיקה מתחילה ב-@ שאחריהם.
    expect(deleted.target.start.offset).toBe(7);
    handle.dispose();
  });

  it('בלי hyperlinks.wrap אין כתיבה, ויש דיווח', async () => {
    const onStatus = vi.fn();
    const { host, calls } = fakeDoc('@פסחים לד');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    delete (host.activeEditor.doc as any).hyperlinks.wrap;
    const handle = installAtMention(container, host as never, { onStatus });

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    expect(calls.has('insert')).toBe(false);
    expect(onStatus).toHaveBeenCalledWith('הוספת קישור אינה זמינה במסמך זה', true);
    handle.dispose();
  });

  it('כשל בעטיפה מחזיר את האזכור במקום הטקסט שנכתב', async () => {
    const { host, calls } = fakeDoc('@פסחים לד', {
      wrap: () => ({ success: false, failure: { code: 'document-readonly' } }),
    });
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const inserts = calls.get('insert') as Array<{
      value: string;
      target: { start: { offset: number }; end: { offset: number } };
    }>;
    expect(inserts).toHaveLength(3);
    // מחיקה, כתיבת הטקסט, ואז החלפתו בחזרה באזכור — על הטווח שנכתב.
    expect(inserts[0]).toMatchObject({ value: '', target: { start: { offset: 0 }, end: { offset: 9 } } });
    expect(inserts[1]).toMatchObject({ value: 'פסחים דף לד', target: { start: { offset: 0 } } });
    expect(inserts[2]).toMatchObject({
      value: '@פסחים לד',
      target: { start: { offset: 0 }, end: { offset: 11 } },
    });
    handle.dispose();
  });

  /**
   * הענף הזה לא היה עטוף: `doc.insert` של הטקסט ישב מחוץ ל-`try`, ולכן
   * **זריקה** בו מחקה את האזכור של המשתמש ולא שחזרה אותו — הטקסט שהוקלד
   * פשוט נעלם, ובלי אזכור לא היה גם מה לנסות שוב.
   */
  it('זריקה בכתיבת הטקסט מחזירה את האזכור ואינה בולעת את השגיאה', async () => {
    let call = 0;
    const { host, calls } = fakeDoc('@פסחים לד', {
      docInsert: () => {
        if (++call === 2) throw new Error('boom');
        return { success: true };
      },
    });
    const onStatus = vi.fn();
    const handle = installAtMention(container, host as never, { onStatus });

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const inserts = calls.get('insert') as Array<{
      value: string;
      target: { start: { offset: number }; end: { offset: number } };
    }>;
    expect(inserts.map((i) => i.value)).toEqual(['', 'פסחים דף לד', '@פסחים לד']);
    // השחזור הוא על הנקודה שנפתחה — טווח רחב ממנה היה מוחק טקסט שכן נשאר.
    expect(inserts[2]!.target).toMatchObject({ start: { offset: 0 }, end: { offset: 0 } });
    expect(onStatus).toHaveBeenCalledWith('הוספת הקישור נכשלה', true);
    handle.dispose();
  });

  it('כשל בכתיבת הטקסט עצמו מחזיר את האזכור לנקודה שנפתחה', async () => {
    let call = 0;
    const { host, calls } = fakeDoc('@פסחים לד', {
      // הראשון הוא המחיקה והוא מצליח; השני הוא כתיבת הטקסט ונכשל.
      docInsert: () => (++call === 2 ? { success: false, failure: { code: 'document-readonly' } } : { success: true }),
    });
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const inserts = calls.get('insert') as Array<{
      value: string;
      target: { start: { offset: number }; end: { offset: number } };
    }>;
    expect(inserts.map((i) => i.value)).toEqual(['', 'פסחים דף לד', '@פסחים לד']);
    // הנקודה שנפתחה, ולא הטווח שהטקסט **היה** תופס אילו נכתב.
    expect(inserts[2]!.target).toMatchObject({ start: { offset: 0 }, end: { offset: 0 } });
    expect(calls.has('hyperlinks.wrap')).toBe(false);
    handle.dispose();
  });

  /**
   * המנוע מחזיר את הקינון כ-**הודעה** ‏(`hyperlink-nested-unsupported`) ואת
   * `INVALID_CONTEXT` כקוד. בדיקה על הקוד לבדו לא התאימה מעולם, והמשתמש ראה
   * בשורת המצב את המחרוזת האנגלית הגולמית.
   */
  it('קינון מדווח בעברית, גם כשהוא מגיע כהודעה ולא כקוד', async () => {
    const onStatus = vi.fn();
    const { host } = fakeDoc('@פסחים לד', {
      wrap: () => ({
        success: false,
        failure: { code: 'INVALID_CONTEXT', message: 'hyperlink-nested-unsupported' },
      }),
    });
    const handle = installAtMention(container, host as never, { onStatus });

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    expect(onStatus).toHaveBeenCalledWith('אי אפשר להוסיף קישור בתוך קישור קיים', true);
    handle.dispose();
  });

  it('חפיפה שמדווחת בנוסח של wrap מתורגמת אף היא', async () => {
    const onStatus = vi.fn();
    const { host } = fakeDoc('@פסחים לד', {
      wrap: () => ({
        success: false,
        failure: {
          code: 'INVALID_TARGET',
          message: 'hyperlinks.wrap does not support ranges that overlap an existing hyperlink.',
        },
      }),
    });
    const handle = installAtMention(container, host as never, { onStatus });

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    expect(onStatus).toHaveBeenCalledWith('אי אפשר להוסיף קישור בתוך קישור קיים', true);
    handle.dispose();
  });

  it('החצים מדלגים מעגלית, ו-Tab מקבל את הפריט הפעיל', async () => {
    resolveRefMock.mockResolvedValue({
      ok: true,
      value: [hit(), hit({ id: 7, reference: 'פסחים דף לה', index: 99 })],
    });
    const { host, calls } = fakeDoc('@פסחים');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    expect(options()[0]!.getAttribute('aria-selected')).toBe('true');
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(options()[1]!.getAttribute('aria-selected')).toBe('true');
    // מעגלי: עוד צעד אחד חוזר לראשון.
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(options()[0]!.getAttribute('aria-selected')).toBe('true');
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
    expect(options()[1]!.getAttribute('aria-selected')).toBe('true');

    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const linked = calls.get('hyperlinks.wrap')?.[0] as { link: { destination: { href: string } } };
    expect(linked.link.destination.href).toBe('otzaria://open/book/7?index=99&uid=id%3A42');
    handle.dispose();
  });

  it('לחיצת עכבר בוחרת את השורה שנלחצה', async () => {
    resolveRefMock.mockResolvedValue({
      ok: true,
      value: [hit(), hit({ id: 7, reference: 'פסחים דף לה', index: 99 })],
    });
    const { host, calls } = fakeDoc('@פסחים');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    options()[1]!.click();
    await vi.advanceTimersByTimeAsync(0);

    const linked = calls.get('hyperlinks.wrap')?.[0] as { link: { destination: { href: string } } };
    expect(linked.link.destination.href).toBe('otzaria://open/book/7?index=99&uid=id%3A42');
    handle.dispose();
  });

  it('Escape סוגר בלי לכתוב', async () => {
    const { host, calls } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

    expect(popup()).toBeNull();
    expect(calls.has('insert')).toBe(false);
    handle.dispose();
  });

  it('keyup של מקשי הרשימה אינו מעריך מחדש', async () => {
    // נמדד בשער: בלי הסינון, ה-keyup של חץ למטה בנה session חדש שהחזיר את
    // הבחירה לפריט הראשון.
    resolveRefMock.mockResolvedValue({
      ok: true,
      value: [hit(), hit({ id: 7, reference: 'פסחים דף לה', index: 99 })],
    });
    const { host } = fakeDoc('@פסחים');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    container.dispatchEvent(new KeyboardEvent('keyup', { key: 'ArrowDown', bubbles: true }));
    await settle();

    expect(options()[1]!.getAttribute('aria-selected')).toBe('true');
    handle.dispose();
  });

  it('אחרי Escape אותו אזכור אינו נפתח מחדש', async () => {
    // נמדד בשער: ה-keyup של Escape עצמו פתח את הרשימה מיד מחדש.
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    container.dispatchEvent(new KeyboardEvent('keyup', { key: 'Escape', bubbles: true }));
    await settle();
    expect(popup()).toBeNull();

    // גם הקלדה נוספת באותו אזכור אינה מחזירה אותו.
    container.dispatchEvent(new Event('input'));
    await settle();
    expect(popup()).toBeNull();
    handle.dispose();
  });

  it('אזכור חדש נפתח גם אחרי שקודמו נדחה', async () => {
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await settle();
    expect(popup()).toBeNull();
    handle.dispose();

    // אזכור במיקום אחר — הדחייה נקשרת להיסט ה-@, לא לפיצ'ר כולו.
    const second = fakeDoc('ראה @ברכות ב');
    const handle2 = installAtMention(container, second.host as never);
    container.dispatchEvent(new Event('input'));
    await settle();
    expect(popup()).not.toBeNull();
    handle2.dispose();
  });

  it('מקש שאינו של הרשימה אינו נבלע', async () => {
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    const event = new KeyboardEvent('keydown', { key: 'a', bubbles: true, cancelable: true });
    container.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    handle.dispose();
  });

  it('Tab אינו נבלע כשאין רשימה פתוחה', async () => {
    const { host } = fakeDoc('סתם טקסט');
    const handle = installAtMention(container, host as never);

    const event = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
    container.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    handle.dispose();
  });

  it('כשל הרשאה מדווח פעם אחת ולא פותח רשימה', async () => {
    resolveRefMock.mockResolvedValue({
      ok: false,
      reason: 'permission-denied',
      message: 'חסרה הרשאה',
    });
    const onStatus = vi.fn();
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never, { onStatus });

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new Event('input'));
    await settle();

    expect(onStatus).toHaveBeenCalledTimes(1);
    expect(onStatus).toHaveBeenCalledWith('חסרה הרשאה', true);
    expect(popup()).toBeNull();
    // הרשאה חסרה אינה חולפת: אין טעם לשלוח קריאה נוספת בכל הקלדה.
    expect(resolveRefMock).toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('dispose מסיר את הרשימה ומפסיק להאזין', async () => {
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    expect(popup()).not.toBeNull();

    handle.dispose();
    expect(popup()).toBeNull();

    resolveRefMock.mockClear();
    container.dispatchEvent(new Event('input'));
    await settle();
    expect(resolveRefMock).not.toHaveBeenCalled();
  });

  /**
   * הסף מנקודת מבטו של המקליד: מה נשלח לאוצריא ומה נפתח על המסך. הדיווח
   * מהשדה היה שרק אחרי גרשיים מופיעות הצעות — כאן נבדק שכל תו שני מספיק.
   */
  describe('סף שני התווים בנתיב המלא', () => {
    it('אות אחת אחרי @ אינה שולחת שאילתה ואינה פותחת רשימה', async () => {
      const { host } = fakeDoc('ראה @ר');
      const handle = installAtMention(container, host as never);

      container.dispatchEvent(new Event('input'));
      await settle();

      expect(resolveRefMock).not.toHaveBeenCalled();
      expect(popup()).toBeNull();
      handle.dispose();
    });

    it('שתי אותיות פותחות רשימה בלי שום סימן פיסוק', async () => {
      const { host } = fakeDoc('ראה @רש');
      const handle = installAtMention(container, host as never);

      container.dispatchEvent(new Event('input'));
      await settle();

      expect(resolveRefMock).toHaveBeenCalledWith('רש', 8);
      expect(options()).toHaveLength(1);
      handle.dispose();
    });

    it('אות וגרשיים פותחים רשימה — הגרשיים אינו תנאי אלא תו שני', async () => {
      const { host } = fakeDoc('ראה @ר״');
      const handle = installAtMention(container, host as never);

      container.dispatchEvent(new Event('input'));
      await settle();

      expect(resolveRefMock).toHaveBeenCalledWith('ר״', 8);
      expect(options()).toHaveLength(1);
      handle.dispose();
    });

    it('הקלדת התו השני פותחת את הרשימה על אותו אזכור', async () => {
      // אותו אזכור בדיוק, שני מצבי הקלדה: „@ר” ואז „@רש”.
      const first = fakeDoc('ראה @ר');
      const handle = installAtMention(container, first.host as never);
      container.dispatchEvent(new Event('input'));
      await settle();
      expect(popup()).toBeNull();
      handle.dispose();

      const second = fakeDoc('ראה @רש');
      const handle2 = installAtMention(container, second.host as never);
      container.dispatchEvent(new Event('input'));
      await settle();
      expect(popup()).not.toBeNull();
      handle2.dispose();
    });
  });

  it('גלילה סוגרת את הרשימה — היא אינה נגררת עם הטקסט', async () => {
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new Event('scroll'));

    expect(popup()).toBeNull();
    handle.dispose();
  });
});

/**
 * החימוש — מה שמוציא את הפיצ'ר ממסלול ההקלדה.
 *
 * `evaluate` הוא שתי קריאות RPC למנוע, והוא רץ 180ms אחרי כל תו. מה שנמדד
 * כאן הוא שהוא **אינו** רץ על טקסט רגיל, וכן רץ בכל אחד מהמסלולים שבהם
 * הסמן עשוי להיות בתוך אזכור: „@” שהוקלד, הדבקה, תנועת סמן, לחיצה.
 */
describe('חימוש — מי בכלל שואל את המנוע', () => {
  /** כפיל שסופר את שתי קריאות ה-RPC של `evaluate`. */
  function countingDoc(text: string) {
    const { host } = fakeDoc(text);
    const doc = (host as { activeEditor: { doc: { selection: { current: () => unknown } } } })
      .activeEditor.doc;
    const current = doc.selection.current;
    let asked = 0;
    doc.selection.current = () => {
      asked += 1;
      return current();
    };
    return { host, asked: () => asked };
  }

  function typed(data: string): InputEvent {
    return new InputEvent('input', { data, inputType: 'insertText' });
  }

  function keyed(key: string): KeyboardEvent {
    return new KeyboardEvent('keyup', { key, bubbles: true });
  }

  it('תו רגיל אינו שואל את המנוע כלל', async () => {
    const { host, asked } = countingDoc('ראה פסחים לד');
    const handle = installAtMention(container, host as never);

    for (const letter of 'פסחים') container.dispatchEvent(typed(letter));
    await settle();

    expect(asked()).toBe(0);
    handle.dispose();
  });

  it('`keyup` הוא המסלול היחיד במנוע האמיתי — „@” בו מחמש', async () => {
    // נמדד על ה-dist הארוז: המנוע מטפל בהקלדה ב-`keydown` ומכניס בעצמו,
    // ולכן אין `input` ואין `beforeinput` בכלל — שלוש הקשות, שלושה `keyup`,
    // אפס `input`. חימוש שנשען על `InputEvent.data` לבדו היה כיבוי מוחלט.
    const { host, asked } = countingDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(keyed('א'));
    await settle();
    expect(asked(), 'תו רגיל ב-keyup אינו שואל').toBe(0);

    container.dispatchEvent(keyed('@'));
    await settle();
    expect(asked()).toBe(1);

    container.dispatchEvent(keyed('פ'));
    await settle();
    expect(asked()).toBe(2);
    expect(popup()).not.toBeNull();
    handle.dispose();
  });

  it('application paste waits for completion before arming an initially idle mention', async () => {
    const { host, asked } = countingDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);
    const operation = {};
    container.dispatchEvent(new CustomEvent(PASTE_STARTED, { detail: { operation } }));
    container.dispatchEvent(keyed('v'));
    await settle();
    expect(asked()).toBe(0);
    container.dispatchEvent(new CustomEvent(PASTE_FINISHED, { detail: { operation } }));
    await settle();
    expect(asked()).toBe(1);
    expect(popup()).not.toBeNull();
    handle.dispose();
  });

  it('application paste cancels stale suggestions and its listeners are disposed', async () => {
    const { host, asked } = countingDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);
    container.dispatchEvent(keyed('@'));
    await settle();
    expect(popup()).not.toBeNull();
    const operation = {};
    container.dispatchEvent(new CustomEvent(PASTE_STARTED, { detail: { operation } }));
    expect(popup()).toBeNull();
    await settle();
    expect(asked()).toBe(1);
    handle.dispose();
    container.dispatchEvent(new CustomEvent(PASTE_FINISHED, { detail: { operation } }));
    await settle();
    expect(asked()).toBe(1);
  });

  it('הדבקה מחמשת גם בלי `input` — אין לה מקש', async () => {
    const { host, asked } = countingDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('paste', { bubbles: true }));
    await settle();

    expect(asked()).toBe(1);
    handle.dispose();
  });

  it('„@” מחמש, ומכאן והלאה כל תו נבדק', async () => {
    const { host, asked } = countingDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(typed('@'));
    await settle();
    expect(asked()).toBe(1);

    container.dispatchEvent(typed('פ'));
    await settle();
    expect(asked()).toBe(2);
    expect(popup()).not.toBeNull();
    handle.dispose();
  });

  it('הדבקה מחמשת — אין ב-`data` מה לקרוא', async () => {
    const { host, asked } = countingDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new InputEvent('input', { inputType: 'insertFromPaste' }));
    await settle();

    expect(asked()).toBe(1);
    handle.dispose();
  });

  it('תנועת סמן ולחיצה מחמשות — כך נכנסים לאזכור שכבר בטקסט', async () => {
    const { host, asked } = countingDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(keyed('ArrowLeft'));
    await settle();
    expect(asked()).toBe(1);

    container.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    await settle();
    expect(asked()).toBe(2);
    handle.dispose();
  });

  it('אזכור שנסגר מפרק את החימוש — התו שאחריו אינו שואל', async () => {
    // „ראה מקור” — אין „@” בכלל, ולכן הבדיקה הראשונה מפרקת.
    const { host, asked } = countingDoc('ראה מקור');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(typed('@'));
    await settle();
    expect(asked(), 'הבדיקה שמפרקת').toBe(1);

    container.dispatchEvent(typed('ר'));
    container.dispatchEvent(typed('ק'));
    await settle();

    expect(asked()).toBe(1);
    handle.dispose();
  });

  it('מקש רגיל ב-keyup אינו מכפיל את הבדיקה של `input`', async () => {
    const { host, asked } = countingDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);
    container.dispatchEvent(typed('@'));
    await settle();
    const armed = asked();

    container.dispatchEvent(typed('פ'));
    container.dispatchEvent(keyed('פ'));
    await settle();

    expect(asked()).toBe(armed + 1);
    handle.dispose();
  });
});

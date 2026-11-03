/**
 * ‏`Ctrl+Backspace`/`Ctrl+Delete` שאינם מוחקים אחרי פקודת רצועה — העקיפה.
 *
 * הבאג עצמו (superdoc/docx-editor#4084) אינו ניתן לשחזור ב-jsdom, ונבדק
 * בשער `scripts/qa/word-delete-qa.mjs`.
 * כאן נבדק החוזה של המאזין: מתי הוא כותב את הסמן, לאן, ושהוא לעולם אינו
 * בולע את ההקשה.
 */
import { describe, it, expect, vi } from 'vitest';
import {
  collapsedCaret,
  installWordDeleteRefresh,
  isLaggingSnapshot,
  isWordDeleteKey,
  type WordDeleteHost,
} from '../../src/engine/word-delete-refresh';

const STORY = { kind: 'story', storyType: 'body' };

function key(over: Partial<KeyboardEvent> = {}) {
  return { key: 'Backspace', ctrlKey: true, metaKey: false, altKey: false, shiftKey: false, ...over } as KeyboardEvent;
}

describe('isWordDeleteKey', () => {
  it('Ctrl+Backspace ו-Ctrl+Delete', () => {
    expect(isWordDeleteKey(key())).toBe(true);
    expect(isWordDeleteKey(key({ key: 'Delete' }))).toBe(true);
  });

  it('בלי Ctrl, או עם מודיפייר נוסף — המנוע מוחק תו, ואינו במסלול השבור', () => {
    expect(isWordDeleteKey(key({ ctrlKey: false }))).toBe(false);
    expect(isWordDeleteKey(key({ shiftKey: true }))).toBe(false);
    expect(isWordDeleteKey(key({ altKey: true }))).toBe(false);
    expect(isWordDeleteKey(key({ metaKey: true }))).toBe(false);
  });

  it('מקש אחר עם Ctrl אינו שלנו', () => {
    expect(isWordDeleteKey(key({ key: 'z' }))).toBe(false);
    expect(isWordDeleteKey(key({ key: 'ArrowLeft' }))).toBe(false);
  });

  it('הרכבה במקלדת (IME) אינה שלנו', () => {
    expect(isWordDeleteKey(key({ isComposing: true }))).toBe(false);
    expect(isWordDeleteKey(key({ keyCode: 229 }))).toBe(false);
  });
});

describe('collapsedCaret', () => {
  const point = (offset: number, blockId = 'p1') => ({ kind: 'text', blockId, offset, story: STORY });

  it('סמן מכווץ — הנקודה', () => {
    expect(collapsedCaret({ selectionTarget: { kind: 'selection', start: point(14), end: point(14) } })).toEqual(point(14));
  });

  it('בחירה פתוחה, או בין פסקאות — null: המנוע מוחק את הבחירה ולא מילה', () => {
    expect(collapsedCaret({ selectionTarget: { start: point(9), end: point(14) } })).toBeNull();
    expect(collapsedCaret({ selectionTarget: { start: point(3, 'p1'), end: point(3, 'p2') } })).toBeNull();
  });

  it('תצלום ריק — `selectionTarget: null` הוא מה שהמנוע מחזיר לבחירה לא מכווצת', () => {
    expect(collapsedCaret(null)).toBeNull();
    expect(collapsedCaret({ selectionTarget: null })).toBeNull();
    expect(collapsedCaret({ selectionTarget: { start: { kind: 'node' }, end: { kind: 'node' } } })).toBeNull();
  });
});

describe('isLaggingSnapshot', () => {
  const at = (offset: number, blockId = 'p1') => ({ kind: 'text', blockId, offset });
  const last = { blockId: 'p1', offset: 13, revision: 7, direction: 'backward' as const };

  it('אין כתיבה קודמת, או נקודה אחרת — אינו מפגר', () => {
    expect(isLaggingSnapshot(null, at(13), 7)).toBe(false);
    expect(isLaggingSnapshot(last, at(11), 7)).toBe(false);
    expect(isLaggingSnapshot(last, at(13, 'p2'), 7)).toBe(false);
  });

  it('אותה נקודה, והמחיקה טרם נקלטה או נקלטה זה עתה — מפגר', () => {
    expect(isLaggingSnapshot(last, at(13), 7)).toBe(true);
    expect(isLaggingSnapshot(last, at(13), 8)).toBe(true);
  });

  it('אותה נקודה אחרי מחיקה ועוד פעולה (הקלדה חזרה, ביטול) — אמיתי', () => {
    expect(isLaggingSnapshot(last, at(13), 9)).toBe(false);
    expect(isLaggingSnapshot(last, at(13), 14)).toBe(false);
  });

  it('מחיקה קדימה נשארת באותו היסט — רק מוטציה שנקלטה מתירה כתיבה חוזרת', () => {
    const forward = { ...last, direction: 'forward' as const };
    expect(isLaggingSnapshot(forward, at(13), 7)).toBe(true);
    expect(isLaggingSnapshot(forward, at(13), 8)).toBe(false);
    expect(isLaggingSnapshot(forward, at(13), undefined)).toBe(true);
    expect(isLaggingSnapshot({ ...forward, revision: undefined }, at(13), 8)).toBe(true);
    expect(isLaggingSnapshot(forward, at(13), 6)).toBe(true);
  });

  it('בלי מונה מוטציות — אותה נקודה נחשבת מפגרת (עדיף להחמיץ הקשה ממחיקה שגויה)', () => {
    expect(isLaggingSnapshot(last, at(13), undefined)).toBe(true);
    expect(isLaggingSnapshot({ ...last, revision: undefined }, at(13), 9)).toBe(true);
  });
});

describe('installWordDeleteRefresh', () => {
  function setup(selectionTarget: unknown) {
    const host = document.createElement('div');
    const inner = document.createElement('textarea');
    host.appendChild(inner);
    document.body.appendChild(host);
    const setSelectionTarget = vi.fn((_input: { target: { coordinateSpace?: string } }) => Promise.resolve({ ok: true }));
    const superdoc: WordDeleteHost = {
      activeEditor: {
        host: { readLiveSelectionSyncSnapshot: () => ({ selectionTarget }) as never },
        authoring: { setSelectionTarget },
      },
    };
    const handle = installWordDeleteRefresh({ host, superdoc });
    // המנוע מאזין על היעד עצמו; המאזין שלנו חייב לרוץ לפניו (capture על ה-host).
    const engine = vi.fn((event: KeyboardEvent) => ({ calledBefore: setSelectionTarget.mock.calls.length, prevented: event.defaultPrevented }));
    const order: unknown[] = [];
    inner.addEventListener('keydown', (event) => order.push(engine(event)));
    const press = (over: KeyboardEventInit = {}) => {
      const event = new KeyboardEvent('keydown', { key: 'Backspace', ctrlKey: true, bubbles: true, cancelable: true, ...over });
      inner.dispatchEvent(event);
      return event;
    };
    return { host, setSelectionTarget, handle, press, order };
  }

  const caret = { kind: 'selection', start: { kind: 'text', blockId: 'p1', offset: 14, story: STORY }, end: { kind: 'text', blockId: 'p1', offset: 14, story: STORY }, story: STORY };

  it('Ctrl+Backspace בסמן מכווץ: כותב את הסמן לאותה נקודה, לפני המנוע, ואינו בולע', () => {
    const { setSelectionTarget, press, order } = setup(caret);
    const event = press();
    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    expect(setSelectionTarget).toHaveBeenCalledWith({
      target: {
        kind: 'selection',
        start: { kind: 'text', blockId: 'p1', offset: 14, story: STORY },
        end: { kind: 'text', blockId: 'p1', offset: 14, story: STORY },
        story: STORY,
      },
      focus: true,
    });
    expect(event.defaultPrevented, 'המחיקה עצמה שייכת למנוע').toBe(false);
    expect(order).toEqual([{ calledBefore: 1, prevented: false }]);
  });

  it('Ctrl+Delete — אותו מסלול', () => {
    const { setSelectionTarget, press } = setup(caret);
    press({ key: 'Delete' });
    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
  });

  it('coordinateSpace מהתצלום עובר הלאה (מעקב שינויים)', () => {
    const { setSelectionTarget, press } = setup({ ...caret, coordinateSpace: 'tracked' });
    press();
    expect(setSelectionTarget.mock.calls[0]?.[0].target.coordinateSpace).toBe('tracked');
  });

  it('Backspace רגיל, Ctrl+Shift+Backspace, בחירה פתוחה — לא נוגע', () => {
    const plain = setup(caret);
    plain.press({ ctrlKey: false });
    plain.press({ shiftKey: true });
    expect(plain.setSelectionTarget).not.toHaveBeenCalled();

    const range = setup({ ...caret, end: { ...caret.end, offset: 9 } });
    range.press();
    expect(range.setSelectionTarget).not.toHaveBeenCalled();
  });

  it('dispose מסיר את המאזין', () => {
    const { setSelectionTarget, press, handle } = setup(caret);
    handle.dispose();
    press();
    expect(setSelectionTarget).not.toHaveBeenCalled();
  });

  it('מנוע בלי התצלום או בלי setSelectionTarget — ההקשה עוברת כמו שהיא', () => {
    const host = document.createElement('div');
    const handle = installWordDeleteRefresh({ host, superdoc: { activeEditor: { host: null, authoring: null } } });
    const event = new KeyboardEvent('keydown', { key: 'Backspace', ctrlKey: true, cancelable: true });
    expect(() => host.dispatchEvent(event)).not.toThrow();
    expect(event.defaultPrevented).toBe(false);
    handle.dispose();
  });

  it('setSelectionTarget שזורק אינו עוצר את ההקשה', () => {
    const host = document.createElement('div');
    const superdoc: WordDeleteHost = {
      activeEditor: {
        host: { readLiveSelectionSyncSnapshot: () => ({ selectionTarget: caret }) as never },
        authoring: {
          setSelectionTarget: () => {
            throw new Error('boom');
          },
        },
      },
    };
    const handle = installWordDeleteRefresh({ host, superdoc });
    const event = new KeyboardEvent('keydown', { key: 'Backspace', ctrlKey: true, cancelable: true });
    expect(() => host.dispatchEvent(event)).not.toThrow();
    expect(event.defaultPrevented).toBe(false);
    handle.dispose();
  });
  it('Ctrl+B בלי בחירה ואז Ctrl+Backspace: העיצוב הממתין שאופס בכתיבה משוחזר', () => {
    const host = document.createElement('div');
    let pending: Record<string, unknown> | null = { bold: true, fontSize: 14 };
    // כמו המנוע: כתיבת בחירה מאפסת את העיצוב הממתין, סינכרונית.
    const setSelectionTarget = vi.fn(() => {
      pending = null;
      return Promise.resolve({ ok: true });
    });
    const setPendingInlineFormat = vi.fn((k: string, v: unknown) => {
      pending = { ...(pending ?? {}), [k]: v };
    });
    const superdoc: WordDeleteHost = {
      activeEditor: {
        host: {
          readLiveSelectionSyncSnapshot: () => ({ selectionTarget: caret }) as never,
          getPendingInlineFormat: () => pending,
          setPendingInlineFormat,
        },
        authoring: { setSelectionTarget },
      },
    };
    const handle = installWordDeleteRefresh({ host, superdoc });
    host.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace', ctrlKey: true, cancelable: true }));
    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    expect(pending).toEqual({ bold: true, fontSize: 14 });
    handle.dispose();
  });

  it('בלי עיצוב ממתין — אינו כותב עיצוב', () => {
    const host = document.createElement('div');
    const setPendingInlineFormat = vi.fn();
    const superdoc: WordDeleteHost = {
      activeEditor: {
        host: {
          readLiveSelectionSyncSnapshot: () => ({ selectionTarget: caret }) as never,
          getPendingInlineFormat: () => null,
          setPendingInlineFormat,
        },
        authoring: { setSelectionTarget: () => Promise.resolve({ ok: true }) },
      },
    };
    const handle = installWordDeleteRefresh({ host, superdoc });
    host.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace', ctrlKey: true, cancelable: true }));
    expect(setPendingInlineFormat).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('מקש מוחזק: תצלום שמפגר על הנקודה הקודמת אינו נכתב שוב (היסט ישן = מחיקת מילה אחרת)', () => {
    const host = document.createElement('div');
    let offset = 15;
    let revision = 3;
    const at = (o: number) => ({ kind: 'text', blockId: 'p1', offset: o, story: STORY });
    const setSelectionTarget = vi.fn((_input: { target: { start: { offset?: number } } }) => Promise.resolve({ ok: true }));
    const superdoc: WordDeleteHost = {
      activeEditor: {
        host: {
          readLiveSelectionSyncSnapshot: () => ({ selectionTarget: { kind: 'selection', start: at(offset), end: at(offset), story: STORY } }) as never,
          getLocalMutationRevision: () => revision,
        },
        authoring: { setSelectionTarget },
      },
    };
    const handle = installWordDeleteRefresh({ host, superdoc });
    const press = () => host.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace', ctrlKey: true, repeat: true, cancelable: true }));
    const written = () => setSelectionTarget.mock.calls.map((c) => c[0].target.start.offset);

    press(); // 15 → המנוע מוחק „ ח”
    revision += 1;
    offset = 13;
    press(); // 13 — התצלום התעדכן
    press(); // עדיין 13: המחיקה של ההקשה הקודמת טרם נקלטה
    revision += 1;
    press(); // 13, והמחיקה נקלטה — התצלום עדיין מפגר
    expect(written()).toEqual([15, 13]);

    offset = 11;
    press();
    expect(written()).toEqual([15, 13, 11]);
    handle.dispose();
  });

  it('Ctrl+Delete חוזר באותו היסט אחרי commit; החלפת כיוון אינה עוקפת תצלום אחורה מפגר', () => {
    const host = document.createElement('div');
    let revision = 3;
    const setSelectionTarget = vi.fn(() => Promise.resolve({ ok: true }));
    const superdoc: WordDeleteHost = {
      activeEditor: {
        host: { readLiveSelectionSyncSnapshot: () => ({ selectionTarget: caret }) as never, getLocalMutationRevision: () => revision },
        authoring: { setSelectionTarget },
      },
    };
    const handle = installWordDeleteRefresh({ host, superdoc });
    const press = (key: string) => host.dispatchEvent(new KeyboardEvent('keydown', { key, ctrlKey: true, cancelable: true }));
    press('Delete');
    press('Delete'); // no commit: still cannot trust the old snapshot
    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    revision += 1;
    press('Delete');
    revision += 1;
    press('Backspace'); // previous forward deletion legitimately kept this offset
    expect(setSelectionTarget).toHaveBeenCalledTimes(3);
    revision += 1;
    press('Delete'); // previous BACKWARD deletion should have moved the caret
    expect(setSelectionTarget).toHaveBeenCalledTimes(3);
    handle.dispose();
  });

  it('אותה נקודה אחרי מחיקה והקלדה חזרה — נכתב שוב (הפסקה הממוספרת מתיישנת מחדש)', () => {
    const host = document.createElement('div');
    let revision = 3;
    const setSelectionTarget = vi.fn(() => Promise.resolve({ ok: true }));
    const superdoc: WordDeleteHost = {
      activeEditor: {
        host: { readLiveSelectionSyncSnapshot: () => ({ selectionTarget: caret }) as never, getLocalMutationRevision: () => revision },
        authoring: { setSelectionTarget },
      },
    };
    const handle = installWordDeleteRefresh({ host, superdoc });
    const press = () => host.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace', ctrlKey: true, cancelable: true }));
    press();
    revision += 6; // המחיקה, ו-5 תווים שהוקלדו חזרה
    press();
    expect(setSelectionTarget).toHaveBeenCalledTimes(2);
    handle.dispose();
  });
});

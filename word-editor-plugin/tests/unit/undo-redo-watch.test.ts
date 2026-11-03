/**
 * `watchUndoRedoKeys` — ראו ההסבר המלא ב-src/ui/shortcuts/undo-redo-watch.ts
 * למה זה בכלל קיים: `createShortcutDispatcher` מדלג על אירוע `defaultPrevented`,
 * ו-Ctrl+Z/Ctrl+Y עם הפוקוס במסמך אמיתי כבר מטופלים ומבוטלים על ידי ה-`history`
 * המובנה של ProseMirror לפני שהם מגיעים אליו — כלומר צרכן שרוצה לדעת „משהו
 * כמו Undo/Redo נלחץ” לא יכול להסתמך על המנתב הרגיל, ובלי המאזין הנפרד הזה
 * לא שומע על זה בכלל.
 *
 * `onUndo`/`onRedo` נפרדות (לא `onUndoRedo` יחידה) כי QA מדד א-סימטריה
 * אמיתית: הידע שצריך לנקות ב-Undo שונה ממה שצריך לשחזר ב-Redo. `isBlocked`
 * נוסף אחרי ש-QA מדד ש-Ctrl+Z בתוך שדה טקסט של הממשק (לא של המסמך) ניקה
 * מעקב בטעות.
 */
import { describe, it, expect, vi } from 'vitest';
import { watchUndoRedoKeys } from '../../src/ui/shortcuts/undo-redo-watch';

function fakeTarget() {
  const listeners: Array<{ type: string; listener: EventListener; capture: unknown }> = [];
  return {
    calls: listeners,
    fire: (type: string, event: unknown) =>
      listeners
        .filter((entry) => entry.type === type)
        .forEach((entry) => entry.listener(event as Event)),
    addEventListener: (type: string, listener: unknown, capture?: unknown) => {
      listeners.push({ type, listener: listener as EventListener, capture });
    },
    removeEventListener: (type: string, listener: unknown) => {
      const index = listeners.findIndex((entry) => entry.type === type && entry.listener === listener);
      if (index !== -1) listeners.splice(index, 1);
    },
  };
}

/** אירוע מקלדת מינימלי, בצורה ש-`matchShortcut` צריך. `target` — לבדיקות `isBlocked`. */
function key(
  code: string,
  over: Partial<{ ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; altKey: boolean; target: unknown }> = {},
) {
  const { target, ...rest } = over;
  return { code, key: code, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, target, ...rest };
}

describe('watchUndoRedoKeys', () => {
  it('נרשמת ב-capture, לא ב-bubble — זה כל הטעם שלה', () => {
    const target = fakeTarget();
    watchUndoRedoKeys({ onUndo: vi.fn(), onRedo: vi.fn(), target });

    expect(target.calls).toHaveLength(1);
    expect(target.calls[0]).toMatchObject({ type: 'keydown', capture: true });
  });

  it('Ctrl+Z מפעילה את onUndo, לא את onRedo', () => {
    const onUndo = vi.fn();
    const onRedo = vi.fn();
    const target = fakeTarget();
    watchUndoRedoKeys({ onUndo, onRedo, target });

    target.fire('keydown', key('KeyZ', { ctrlKey: true }));

    expect(onUndo).toHaveBeenCalledTimes(1);
    expect(onRedo).not.toHaveBeenCalled();
  });

  it('Ctrl+Y מפעילה את onRedo, לא את onUndo', () => {
    const onUndo = vi.fn();
    const onRedo = vi.fn();
    const target = fakeTarget();
    watchUndoRedoKeys({ onUndo, onRedo, target });

    target.fire('keydown', key('KeyY', { ctrlKey: true }));

    expect(onRedo).toHaveBeenCalledTimes(1);
    expect(onUndo).not.toHaveBeenCalled();
  });

  it('Ctrl+Shift+Z (redo-shift) מפעילה את onRedo', () => {
    const onUndo = vi.fn();
    const onRedo = vi.fn();
    const target = fakeTarget();
    watchUndoRedoKeys({ onUndo, onRedo, target });

    target.fire('keydown', key('KeyZ', { ctrlKey: true, shiftKey: true }));

    expect(onRedo).toHaveBeenCalledTimes(1);
    expect(onUndo).not.toHaveBeenCalled();
  });

  it('Meta שקול ל-Ctrl (מקלדת Mac)', () => {
    const onUndo = vi.fn();
    const target = fakeTarget();
    watchUndoRedoKeys({ onUndo, onRedo: vi.fn(), target });

    target.fire('keydown', key('KeyZ', { metaKey: true }));

    expect(onUndo).toHaveBeenCalledTimes(1);
  });

  it('Z בלי Ctrl אינה מפעילה כלום', () => {
    const onUndo = vi.fn();
    const onRedo = vi.fn();
    const target = fakeTarget();
    watchUndoRedoKeys({ onUndo, onRedo, target });

    target.fire('keydown', key('KeyZ'));

    expect(onUndo).not.toHaveBeenCalled();
    expect(onRedo).not.toHaveBeenCalled();
  });

  it('קיצור לא-קשור (Ctrl+B) אינה מפעילה כלום', () => {
    const onUndo = vi.fn();
    const onRedo = vi.fn();
    const target = fakeTarget();
    watchUndoRedoKeys({ onUndo, onRedo, target });

    target.fire('keydown', key('KeyB', { ctrlKey: true }));

    expect(onUndo).not.toHaveBeenCalled();
    expect(onRedo).not.toHaveBeenCalled();
  });

  it('dispose מנתקת את המאזין', () => {
    const onUndo = vi.fn();
    const target = fakeTarget();
    const watcher = watchUndoRedoKeys({ onUndo, onRedo: vi.fn(), target });

    watcher.dispose();
    target.fire('keydown', key('KeyZ', { ctrlKey: true }));

    expect(onUndo).not.toHaveBeenCalled();
  });

  it('dispose פעמיים אינה זורקת', () => {
    const target = fakeTarget();
    const watcher = watchUndoRedoKeys({ onUndo: vi.fn(), onRedo: vi.fn(), target });

    watcher.dispose();
    expect(() => watcher.dispose()).not.toThrow();
  });

  describe('isBlocked', () => {
    it('בלי isBlocked — לעולם לא חוסמת (ברירת המחדל)', () => {
      const onUndo = vi.fn();
      const target = fakeTarget();
      watchUndoRedoKeys({ onUndo, onRedo: vi.fn(), target });

      target.fire('keydown', key('KeyZ', { ctrlKey: true, target: { fake: 'search-input' } }));

      expect(onUndo).toHaveBeenCalledTimes(1);
    });

    it('isBlocked שמחזירה true — Ctrl+Z בתוך שדה טקסט לא-קשור אינה מנקה', () => {
      // התרחיש שה-QA מדד: Ctrl+Z בתוך #fr-search-input לא אמור לגעת במעקב.
      const onUndo = vi.fn();
      const target = fakeTarget();
      const searchInput = { marker: 'search-input' } as unknown as EventTarget;
      watchUndoRedoKeys({
        onUndo,
        onRedo: vi.fn(),
        target,
        isBlocked: (t) => t === searchInput,
      });

      target.fire('keydown', key('KeyZ', { ctrlKey: true, target: searchInput }));

      expect(onUndo).not.toHaveBeenCalled();
    });

    it('isBlocked שמחזירה false — Ctrl+Z בתוך המסמך כן מטפלת', () => {
      const onUndo = vi.fn();
      const target = fakeTarget();
      const documentSurface = { marker: 'document' } as unknown as EventTarget;
      watchUndoRedoKeys({
        onUndo,
        onRedo: vi.fn(),
        target,
        isBlocked: (t) => t !== documentSurface,
      });

      target.fire('keydown', key('KeyZ', { ctrlKey: true, target: documentSurface }));

      expect(onUndo).toHaveBeenCalledTimes(1);
    });

    it('isBlocked נבדקת גם עבור Redo', () => {
      const onRedo = vi.fn();
      const target = fakeTarget();
      watchUndoRedoKeys({
        onUndo: vi.fn(),
        onRedo,
        target,
        isBlocked: () => true,
      });

      target.fire('keydown', key('KeyY', { ctrlKey: true }));

      expect(onRedo).not.toHaveBeenCalled();
    });
  });
});

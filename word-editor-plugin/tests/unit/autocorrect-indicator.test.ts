import { describe, it, expect, vi } from 'vitest';
import { applyFixesToText, buildMenu, createAutoCorrectIndicator, stopCorrecting } from '../../src/engine/autocorrect-indicator';
import { DEFAULT_AUTOCORRECT, planAutoCorrect } from '../../src/engine/autocorrect';

describe('סימון התיקון האוטומטי', () => {
  it('עריכה או פעולה בממשק פוסלות זהות ישנה, וגם את הקריאות שבאוויר', () => {
    const container = document.createElement('div');
    document.body.append(container);
    const invalidate = vi.fn();
    const handle = createAutoCorrectIndicator(container, { changeBack: () => {}, invalidate });
    const entry = { rule: 'table' as const, original: 'AA', replacement: 'X', blockId: 'b', story: null, start: 0 };
    try {
      handle.show([entry]);
      container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));
      expect(handle.active).toBe(false);
      expect(invalidate).toHaveBeenLastCalledWith(false, expect.any(Event));
      handle.show([entry]);
      container.dispatchEvent(new Event('input', { bubbles: true }));
      expect(handle.active).toBe(false);
      handle.show([entry]);
      document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
      expect(handle.active).toBe(false);
      expect(invalidate).toHaveBeenLastCalledWith(true);
      handle.dispose();
      invalidate.mockClear();
      container.dispatchEvent(new Event('input', { bubbles: true }));
      expect(invalidate).not.toHaveBeenCalled();
    } finally {
      handle.dispose();
      container.remove();
    }
  });
  it('התפריט כמו ב-Word: חזור, הפסק, אפשרויות', () => {
    const items = buildMenu({ rule: 'table', original: 'teh', replacement: 'the' }, { stop: true, options: true });
    expect(items.map((i) => i.id)).toEqual(['undo', 'stop', 'options']);
    expect(items[0]!.label).toBe('חזור ל„teh”');
    expect(items[1]!.label).toBe('הפסק לתקן אוטומטית את „teh”');
  });

  it('פריטים בלי מטפל אינם מוצגים', () => {
    const items = buildMenu({ rule: 'table', original: 'a', replacement: 'b' }, { stop: false, options: false });
    expect(items.map((i) => i.id)).toEqual(['undo']);
  });

  it('הפסק לתקן מוריד רשומה מהטבלה', () => {
    const options = { ...DEFAULT_AUTOCORRECT, replacements: [{ from: 'teh', to: 'the' }, { from: 'x', to: 'y' }] };
    const next = stopCorrecting(options, { rule: 'table', original: 'Teh', replacement: 'The' });
    expect(next.replacements).toEqual([{ from: 'x', to: 'y' }]);
  });

  it('הפסק לתקן מכבה דגל של חוק', () => {
    const next = stopCorrecting(DEFAULT_AUTOCORRECT, { rule: 'dashes', original: '--', replacement: '—' });
    expect(next.dashes).toBe(false);
    expect(stopCorrecting(next, { rule: 'dashes', original: '--', replacement: '—' })).toBe(next);
  });

  it('היסטים אחרי החלפה מוזזים לפי מה שלפניהם', () => {
    const options = { ...DEFAULT_AUTOCORRECT, replacements: [{ from: 'teh', to: 'the' }, { from: 'abcd', to: 'x' }] };
    const text = 'abcd teh ';
    const fixes = planAutoCorrect(text, options, {});
    const applied = applyFixesToText(text, fixes);
    expect(applied.text).toBe('x the ');
    fixes.forEach((fix, i) => expect(applied.text.startsWith(fix.replacement, applied.starts[i]!)).toBe(true));
  });
});

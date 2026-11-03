import { afterEach, describe, expect, it, vi } from 'vitest';

const pasteFromClipboard = vi.fn(async () => ({ ok: true }) as const);
vi.mock('../../src/engine/clipboard', () => ({ pasteFromClipboard }));

const { installPasteIntercept, payloadFromEvent } = await import('../../src/engine/paste-intercept');

/** `DataTransfer` אינו קיים ב-jsdom; זה מה שהאירוע נושא בדפדפן. */
function transfer(parts: Record<string, string>): DataTransfer {
  return {
    types: Object.keys(parts),
    getData: (type: string) => parts[type] ?? '',
  } as unknown as DataTransfer;
}

/** אירוע `paste` שיורד מה-container אל היעד שבתוכו — המקום שבו המנוע מאזין. */
function paste(target: HTMLElement, data: DataTransfer | null) {
  const event = new Event('paste', { bubbles: true, cancelable: true }) as ClipboardEvent;
  Object.defineProperty(event, 'clipboardData', { value: data });
  const engine = vi.fn();
  target.addEventListener('paste', engine);
  target.dispatchEvent(event);
  return { event, engine };
}

function setup(mode: 'keep-source' | 'merge' | 'text-only') {
  const container = document.createElement('div');
  const editor = document.createElement('div');
  container.appendChild(editor);
  document.body.appendChild(container);
  const report = vi.fn();
  const handle = installPasteIntercept({ container, host: null, mode: () => mode, report });
  return { container, editor, report, handle };
}

afterEach(() => {
  document.body.innerHTML = '';
  pasteFromClipboard.mockClear();
});

describe('installPasteIntercept — Ctrl+V לפי ברירת המחדל', () => {
  it('„שמירה על עיצוב המקור” אינה נחטפת: ההדבקה של המנוע כמו שהייתה', () => {
    const { editor } = setup('keep-source');
    const { event, engine } = paste(editor, transfer({ 'text/html': '<b>x</b>', 'text/plain': 'x' }));

    expect(event.defaultPrevented).toBe(false);
    expect(engine).toHaveBeenCalledTimes(1);
    expect(pasteFromClipboard).not.toHaveBeenCalled();
  });

  it('„מזג עיצוב” נחטף לפני המנוע, ומודבק מהתוכן של האירוע', () => {
    const { editor } = setup('merge');
    const { event, engine } = paste(editor, transfer({ 'text/html': '<b>x</b>', 'text/plain': 'x' }));

    expect(event.defaultPrevented).toBe(true);
    // המנוע אינו מקבל את האירוע — אחרת התוכן היה מודבק פעמיים.
    expect(engine).not.toHaveBeenCalled();
    expect(pasteFromClipboard).toHaveBeenCalledWith(null, {
      mode: 'merge',
      payload: {
        source: 'browser',
        items: [
          { type: 'text/html', kind: 'string', data: '<b>x</b>' },
          { type: 'text/plain', kind: 'string', data: 'x' },
        ],
      },
    });
  });

  it('מאזינים אחרים על ה-container עדיין שומעים שהייתה הדבקה', () => {
    const { container, editor } = setup('text-only');
    const sibling = vi.fn();
    container.addEventListener('paste', sibling, true);

    paste(editor, transfer({ 'text/plain': 'x' }));
    expect(sibling).toHaveBeenCalledTimes(1);
  });

  it('אירוע תמונה בלבד מדווח על אי-תמיכה בלי לשנות את המסמך', () => {
    const { editor, report } = setup('merge');
    const { event, engine } = paste(editor, transfer({ Files: '' }));

    expect(event.defaultPrevented).toBe(true);
    expect(engine).not.toHaveBeenCalled();
    expect(report).toHaveBeenCalledWith(expect.objectContaining({ ok: false }));
  });

  it('merge rejects mixed external media rather than silently dropping the image', () => {
    const { editor, report } = setup('merge');
    const { event } = paste(editor, transfer({ Files: '', 'text/plain': 'caption', 'text/html': '<p>caption<img src="cid:image"></p>' }));
    expect(event.defaultPrevented).toBe(true);
    expect(pasteFromClipboard).not.toHaveBeenCalled();
    expect(report).toHaveBeenCalledWith(expect.objectContaining({ ok: false }));
  });

  it('שדה קלט בתוך ה-container מדביק לעצמו', () => {
    const { container } = setup('merge');
    const input = document.createElement('input');
    container.appendChild(input);

    const { event } = paste(input, transfer({ 'text/plain': 'x' }));
    expect(event.defaultPrevented).toBe(false);
  });

  it('dispose מסיר את המאזין', () => {
    const { editor, handle } = setup('merge');
    handle.dispose();

    const { event } = paste(editor, transfer({ 'text/plain': 'x' }));
    expect(event.defaultPrevented).toBe(false);
  });
});

describe('installPasteIntercept — ה-keydown של Ctrl+V', () => {
  /** `keydown` שיורד מה-container אל המנוע, שמטפל ב-Ctrl+V בעצמו. */
  function keydown(target: HTMLElement, init: KeyboardEventInit) {
    const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
    const engine = vi.fn();
    target.addEventListener('keydown', engine);
    target.dispatchEvent(event);
    return { event, engine };
  }

  it('במצב שאינו ברירת המחדל — המנוע אינו רואה את Ctrl+V, והדפדפן ממשיך להדבקה', () => {
    const { editor } = setup('merge');
    const { event, engine } = keydown(editor, { code: 'KeyV', key: 'ה', ctrlKey: true });

    expect(engine).not.toHaveBeenCalled();
    // בלי `preventDefault` — אחרת הדפדפן אינו שולח את אירוע ה-`paste`.
    expect(event.defaultPrevented).toBe(false);
  });

  it('Ctrl+V on macOS uses the explicit reader because the native browser chord is Cmd+V', () => {
    const platform = navigator.platform;
    Object.defineProperty(navigator, 'platform', { configurable: true, value: 'MacIntel' });
    try {
      const { editor } = setup('merge');
      const { event, engine } = keydown(editor, { code: 'KeyV', ctrlKey: true });
      expect(event.defaultPrevented).toBe(true);
      expect(engine).not.toHaveBeenCalled();
      expect(pasteFromClipboard).toHaveBeenCalledWith(null, { mode: 'merge' });
    } finally {
      Object.defineProperty(navigator, 'platform', { configurable: true, value: platform });
    }
  });

  it('בברירת המחדל, ב-Ctrl+Shift+V ובמקשים אחרים — המנוע מקבל את ה-keydown', () => {
    const { editor } = setup('keep-source');
    expect(keydown(editor, { code: 'KeyV', ctrlKey: true }).engine).toHaveBeenCalled();

    document.body.innerHTML = '';
    const merge = setup('merge');
    expect(keydown(merge.editor, { code: 'KeyV', ctrlKey: true, shiftKey: true }).engine).toHaveBeenCalled();
    expect(keydown(merge.editor, { code: 'KeyC', ctrlKey: true }).engine).toHaveBeenCalled();
    expect(keydown(merge.editor, { code: 'KeyV' }).engine).toHaveBeenCalled();
  });
});

describe('payloadFromEvent', () => {
  it('ה-fragment הפרטי של המנוע ראשון, ואחריו HTML וטקסט', () => {
    const payload = payloadFromEvent(
      transfer({ 'text/plain': 'x', 'application/x-superdoc-v2-fragment': '{}', 'text/html': '<p>x</p>' }),
    );
    expect(payload?.items.map((item) => item.type)).toEqual([
      'application/x-superdoc-v2-fragment',
      'text/html',
      'text/plain',
    ]);
  });

  it('אין clipboardData — אין payload', () => {
    expect(payloadFromEvent(null)).toBeNull();
  });
});

/**
 * החיווט של התיקון האוטומטי מול מנוע מדומה.
 *
 * הכפיל מחזיק בלוק אחד כמחרוזת וסמן, וההכנסה מחליפה בו טווח — כמו
 * `doc.insert` האמיתי. מה שנבדק: מה נקרא, מה הוחלף ובאילו היסטים, ובעיקר
 * **מה לא**: טקסט שלא הוקלד, תיקון שבוטל, קריאה שהמשתמש עקף בהקלדה.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  AUTOCORRECT_DEBOUNCE_MS,
  installAutoCorrect,
  touchesTyped,
  type AutoCorrectHandle,
} from '../../src/engine/autocorrect-install';
import {
  BUILT_IN_REPLACEMENTS,
  DEFAULT_AUTOCORRECT,
  type AutoCorrectOptions,
} from '../../src/engine/autocorrect';

/** הסמלים המובנים בלבד — ראו `BASE` ב-`autocorrect.test.ts`: הרשימה שנשלחת היא נתונים שמתחלפים. */
const BASE: AutoCorrectOptions = { ...DEFAULT_AUTOCORRECT, replacements: BUILT_IN_REPLACEMENTS };

interface Model {
  text: string;
  caret: number;
  /** כשאינו null: הסמן מדומה כטווח נבחר. */
  selectionEnd: number | null;
}

function fakeEngine(initial = '') {
  const model: Model = { text: initial, caret: initial.length, selectionEnd: null };
  const inserts: Array<{ value: string; start: number; end: number }> = [];

  const doc = {
    selection: {
      current: async () => ({
        selectionTarget: {
          kind: 'selection',
          start: { kind: 'text', blockId: 'b1', offset: model.caret },
          end: { kind: 'text', blockId: 'b1', offset: model.selectionEnd ?? model.caret },
        },
      }),
    },
    ranges: {
      resolve: async (request: { start: { point: { offset: number } }; end: { point: { offset: number } } }) => {
        const from = request.start.point.offset;
        const to = request.end.point.offset;
        return {
          preview: { text: model.text.slice(from, to) },
          target: { start: { offset: from }, end: { offset: to } },
        };
      },
    },
    insert: vi.fn(
      async (input: { value: string; target: { start: { offset: number }; end: { offset: number } } }) => {
        const start = input.target.start.offset;
        const end = input.target.end.offset;
        inserts.push({ value: input.value, start, end });
        model.text = model.text.slice(0, start) + input.value + model.text.slice(end);
        if (model.caret >= end) model.caret += input.value.length - (end - start);
        return { success: true };
      },
    ),
  };

  return { model, doc, inserts, host: { activeEditor: { doc } } };
}

let container: HTMLElement;
let handle: AutoCorrectHandle | null = null;

function key(name: string, init: KeyboardEventInit = {}): void {
  container.dispatchEvent(new KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true, ...init }));
}

/** הקלדה כמו משתמש: keydown, התו נכנס למודל, ואז `input`. */
function type(engine: ReturnType<typeof fakeEngine>, text: string): void {
  for (const char of text) {
    key(char);
    engine.model.text =
      engine.model.text.slice(0, engine.model.caret) + char + engine.model.text.slice(engine.model.caret);
    engine.model.caret += 1;
    container.dispatchEvent(new Event('input', { bubbles: true }));
  }
}

async function settle(ms = AUTOCORRECT_DEBOUNCE_MS + 20): Promise<void> {
  await vi.advanceTimersByTimeAsync(ms);
}

function install(
  engine: ReturnType<typeof fakeEngine>,
  options: Partial<AutoCorrectOptions> = {},
  extra: { isSuppressed?: () => boolean } = {},
): AutoCorrectHandle {
  handle = installAutoCorrect({
    container,
    host: engine.host as never,
    getOptions: () => ({ ...BASE, ...options }),
    ...extra,
  });
  return handle;
}

describe('installAutoCorrect', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    container = document.createElement('div');
    document.body.append(container);
  });

  afterEach(() => {
    handle?.dispose();
    handle = null;
    container.remove();
    vi.useRealTimers();
  });

  it('מתקן סמל שהוקלד, בשקט, בהיסטים של הסמל עצמו', async () => {
    const engine = fakeEngine('זכויות ');
    install(engine);

    type(engine, '(c)');
    await settle();

    expect(engine.model.text).toBe('זכויות ©');
    expect(engine.inserts).toEqual([{ value: '©', start: 7, end: 10 }]);
  });

  it('הסמן נשאר אחרי התיקון, וההקלדה ממשיכה אחריו', async () => {
    const engine = fakeEngine('');
    install(engine);

    type(engine, '(c)');
    await settle();
    type(engine, ' ok');

    expect(engine.model.text).toBe('© ok');
  });

  it('אינו קורא מהמנוע כשלא הוקלד דבר — טקסט שנפתח עם המסמך אינו מתוקן', async () => {
    const engine = fakeEngine('(c) teh ');
    install(engine);

    container.dispatchEvent(new Event('input', { bubbles: true }));
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('לא מתקן טקסט ישן שלפני האזור שהוקלד: teh שנשאר בכוונה', async () => {
    const engine = fakeEngine('teh and more words here ');
    install(engine, { replacements: [{ from: 'teh', to: 'the' }] });

    type(engine, 'x');
    await settle();

    expect(engine.model.text).toBe('teh and more words here x');
  });

  it('מילה שהוקלדה מתוקנת רק אחרי הרווח', async () => {
    const engine = fakeEngine('');
    install(engine, { replacements: [{ from: 'teh', to: 'the' }] });

    type(engine, 'teh');
    await settle();
    expect(engine.model.text).toBe('teh');

    type(engine, ' ');
    await settle();
    expect(engine.model.text).toBe('the ');
  });

  it('הקלדה בתוך חלון ההמתנה פוסלת את הקריאה — מתקן רק כשהמשתמש נרגע', async () => {
    const engine = fakeEngine('');
    install(engine);

    type(engine, '(c)');
    await vi.advanceTimersByTimeAsync(AUTOCORRECT_DEBOUNCE_MS - 50);
    type(engine, 'x');
    await vi.advanceTimersByTimeAsync(AUTOCORRECT_DEBOUNCE_MS - 50);
    expect(engine.inserts).toHaveLength(0);

    await settle();
    expect(engine.model.text).toBe('©x');
  });

  it('keydown חדש פוסל קריאה באוויר גם לפני שהגיעו input או keyup', async () => {
    const engine = fakeEngine();
    install(engine);
    const resolve = engine.doc.ranges.resolve;
    let finish = () => {};
    engine.doc.ranges.resolve = async (request) => {
      const snapshot = await resolve(request);
      await new Promise<void>((done) => { finish = done; });
      return snapshot;
    };
    type(engine, '(c)');
    await settle();
    key('x');
    finish();
    await settle();
    expect(engine.inserts).toHaveLength(0);
  });

  it('מחוון מחזיר גם החלפה ארוכה מה-preview של 200 תווים', async () => {
    const engine = fakeEngine();
    const replacement = 'a'.repeat(255);
    install(engine, { replacements: [{ from: 'xx', to: replacement }] });
    const resolve = engine.doc.ranges.resolve;
    const reads: number[] = [];
    engine.doc.ranges.resolve = async (request) => {
      const result = await resolve(request);
      reads.push(request.end.point.offset - request.start.point.offset);
      return result.preview.text.length > 200
        ? { ...result, preview: { text: result.preview.text.slice(0, 200), truncated: true } }
        : result;
    };
    type(engine, 'xx ');
    await settle();
    expect(engine.model.text).toBe(replacement + ' ');
    container.innerHTML = `<p data-source-node-id="b1" data-pm-start="1" data-pm-end="257"><span data-pm-start="1" data-pm-end="257">${replacement} </span></p>`;
    const original = Object.getOwnPropertyDescriptor(Range.prototype, 'getClientRects');
    Object.defineProperty(Range.prototype, 'getClientRects', { configurable: true, value: () => [{ left: 0, right: 10, top: 0, bottom: 10, width: 10, height: 10 }] });
    try {
      container.firstElementChild!.dispatchEvent(new MouseEvent('pointermove', { bubbles: true, clientX: 5, clientY: 5 }));
      await vi.advanceTimersByTimeAsync(30);
      (document.querySelector('.otzaria-acorr-chip') as HTMLButtonElement).click();
      const undo = document.querySelector<HTMLButtonElement>('.otzaria-acorr-menu button');
      expect(undo).not.toBeNull();
      undo!.click();
      await settle();
      expect(engine.model.text).toBe('xx ');
      expect(Math.max(...reads)).toBeLessThanOrEqual(200);
    } finally {
      if (original) Object.defineProperty(Range.prototype, 'getClientRects', original);
      else Reflect.deleteProperty(Range.prototype, 'getClientRects');
    }
  });

  it('Ctrl+Z שהחזיר את המקור אינו מתוקן שוב בהקלדה הבאה', async () => {
    const engine = fakeEngine('');
    install(engine);

    type(engine, '(c)');
    await settle();
    expect(engine.model.text).toBe('©');

    // המשתמש מבטל: המנוע מחזיר את המקור. הקיצור מאפס את הספירה.
    key('z', { ctrlKey: true });
    engine.model.text = '(c)';
    engine.model.caret = 3;

    type(engine, ' ');
    await settle();

    expect(engine.model.text).toBe('(c) ');
    expect(engine.inserts).toHaveLength(1);
  });

  it('הדבקה מאפסת: טקסט מודבק אינו מתוקן אפילו כשמקלידים אחריו', async () => {
    const engine = fakeEngine('');
    install(engine, { replacements: [{ from: 'teh', to: 'the' }] });

    container.dispatchEvent(new Event('paste', { bubbles: true }));
    engine.model.text = 'xx teh';
    engine.model.caret = 6;
    type(engine, 'y');
    await settle();

    expect(engine.model.text).toBe('xx tehy');
  });

  it('אחרי הקלדה ארוכה ברצף החלון אינו עובר את גבול ה-200 של המנוע, והתיקון ממשיך', async () => {
    // פסקה ארוכה שכבר קיימת: אחרת החלון מוגבל בגודל הבלוק עצמו ולא בגבול של המנוע.
    const engine = fakeEngine('y'.repeat(300));
    install(engine);
    // כמו המנוע האמיתי: טווח ארוך מ-200 תווים חוזר קטום, וקטום נדחה.
    const resolve = engine.doc.ranges.resolve;
    engine.doc.ranges.resolve = async (request) => {
      const result = await resolve(request);
      const text = result.preview.text;
      return text.length > 200 ? { ...result, preview: { text: text.slice(0, 200), truncated: true } } : result;
    };

    type(engine, 'x'.repeat(190));
    type(engine, '(c)');
    await settle();

    expect(engine.model.text.endsWith('©')).toBe(true);
  });

  it('חלון שמתחיל באמצע מילה אינו מתקן את הסיומת שלה, גם בטבלה אישית', async () => {
    const engine = fakeEngine();
    install(engine, { replacements: [{ from: 'למד', to: 'למ"ד' }] });
    const text = 'מלמד ' + 'x'.repeat(196);
    type(engine, text);
    await settle();
    expect(engine.model.text).toBe(text);
    expect(engine.inserts).toHaveLength(0);
  });

  it('קיצור שהוקלד אינו מזיז את אזור ההקלדה אל מילה קיימת', async () => {
    const engine = fakeEngine('teh ');
    install(engine, { replacements: [...BUILT_IN_REPLACEMENTS, { from: 'teh', to: 'the' }] });
    type(engine, '(c)');
    await settle();
    expect(engine.model.text).toBe('teh ©');
    type(engine, ' ');
    await settle();
    expect(engine.model.text).toBe('teh © ');
  });

  it('מילה קיימת שהושלמה ברווח עדיין מתוקנת, ואפשר לתקן גם את ההקלדה הבאה', async () => {
    const engine = fakeEngine('longword');
    install(engine, { replacements: [{ from: 'longword', to: 'x' }, { from: 'teh', to: 'the' }] });
    type(engine, ' ');
    await settle();
    expect(engine.model.text).toBe('x ');
    type(engine, 'teh ');
    await settle();
    expect(engine.model.text).toBe('x the ');
  });

  describe('דפוס שממתין לתו הבא (-- מול -->)', () => {
    const KEYS = { replacements: [...BUILT_IN_REPLACEMENTS, { from: '--', to: ':' }] };
    /** Enter מדומה: הסמן עובר לבלוק חדש, והבלוק הקודם נשאר במודל (b1). */
    function enter(engine: ReturnType<typeof fakeEngine>): void {
      key('Enter');
      engine.model.caret = engine.model.text.length + 1;
      engine.model.text += ' ';
    }

    it('ממתין בשקט, ו-Enter מחליף אותו במקומו', async () => {
      const engine = fakeEngine('');
      install(engine, KEYS);
      type(engine, 'ואלו הן--');
      await settle();
      expect(engine.model.text).toBe('ואלו הן--');
      enter(engine);
      await settle();
      expect(engine.model.text).toBe('ואלו הן: ');
    });

    it('Tab מחליף כמו Enter', async () => {
      const engine = fakeEngine('');
      install(engine, KEYS);
      type(engine, 'א--');
      await settle();
      key('Tab');
      await settle();
      expect(engine.model.text).toBe('א:');
    });

    it('התו הבא מכריע לבדו, ו-Enter אחריו אינו מחליף שוב', async () => {
      const engine = fakeEngine('');
      install(engine, KEYS);
      type(engine, 'א--');
      await settle();
      type(engine, '>');
      await settle();
      expect(engine.model.text).toBe('א→');
      enter(engine);
      await settle();
      expect(engine.inserts).toHaveLength(1);
    });

    it('Enter לפני שההערכה רצה אינו מחליף — אין ניחוש בלי קריאה', async () => {
      const engine = fakeEngine('');
      install(engine, KEYS);
      type(engine, 'א--');
      enter(engine);
      await settle();
      expect(engine.inserts).toHaveLength(0);
    });

    it.each(['reset', 'undo', 'pointer', 'typing', 'delete', 'paste', 'composition', 'dispose'])(
      'מבטל החלפה מתוזמנת אחרי %s', async (action) => {
        const engine = fakeEngine('');
        const installed = install(engine, KEYS);
        type(engine, 'א--');
        await settle();
        enter(engine);
        if (action === 'reset') installed.reset();
        if (action === 'undo') {
          key('z', { ctrlKey: true });
          container.dispatchEvent(new InputEvent('input', { inputType: 'historyUndo', bubbles: true }));
        }
        if (action === 'pointer') container.dispatchEvent(new Event('pointerdown', { bubbles: true }));
        if (action === 'typing') type(engine, 'x');
        if (action === 'delete') container.dispatchEvent(new InputEvent('input', { inputType: 'deleteContentBackward', bubbles: true }));
        if (action === 'paste') container.dispatchEvent(new Event('paste', { bubbles: true }));
        if (action === 'composition') container.dispatchEvent(new Event('compositionstart', { bubbles: true }));
        if (action === 'dispose') installed.dispose();
        await settle();
        expect(engine.inserts).toHaveLength(0);
      },
    );

    it('אירוע הקלט של Enter עצמו אינו מבטל את ההחלפה', async () => {
      const engine = fakeEngine('');
      install(engine, KEYS);
      type(engine, 'א--');
      await settle();
      enter(engine);
      container.dispatchEvent(new InputEvent('input', { inputType: 'insertParagraph', bubbles: true }));
      await settle();
      expect(engine.model.text).toBe('א:\u0000');
    });

    it('קריאה ישנה אינה מוחקת טקסט חדש שהוקלד בזמן ההמתנה', async () => {
      const engine = fakeEngine('');
      install(engine, KEYS);
      type(engine, 'א--');
      await settle();
      const read = engine.doc.ranges.resolve;
      let finish!: () => void;
      engine.doc.ranges.resolve = async (request) => {
        const snapshot = await read(request);
        await new Promise<void>((resolve) => { finish = resolve; });
        return snapshot;
      };
      enter(engine);
      await settle();
      key('Backspace');
      key('X');
      engine.model.text = 'אXY';
      engine.model.caret = 3;
      finish();
      await settle();
      expect(engine.model.text).toBe('אXY');
      expect(engine.inserts).toHaveLength(0);
    });

    it('reset פוסל גם קריאה שכבר התחילה', async () => {
      const engine = fakeEngine('');
      const installed = install(engine, KEYS);
      type(engine, 'א--');
      await settle();
      const read = engine.doc.ranges.resolve;
      let finish!: () => void;
      engine.doc.ranges.resolve = async (request) => {
        const snapshot = await read(request);
        await new Promise<void>((resolve) => { finish = resolve; });
        return snapshot;
      };
      enter(engine);
      await settle();
      installed.reset();
      finish();
      await settle();
      expect(engine.inserts).toHaveLength(0);
    });

    it('שינוי ההגדרות אחרי Enter מבטל את הדפוס הישן', async () => {
      const engine = fakeEngine('');
      let settings = { ...BASE, ...KEYS };
      handle = installAutoCorrect({ container, host: engine.host as never, getOptions: () => settings });
      type(engine, 'א--');
      await settle();
      enter(engine);
      settings = { ...settings, replacements: [{ from: '--', to: 'different' }] };
      await settle();
      expect(engine.inserts).toHaveLength(0);
    });

    it('דיאלוג שנפתח בזמן קריאת הטווח מונע החלפה', async () => {
      const engine = fakeEngine('');
      let suppressed = false;
      install(engine, KEYS, { isSuppressed: () => suppressed });
      type(engine, 'א--');
      await settle();
      const read = engine.doc.ranges.resolve;
      let finish!: () => void;
      engine.doc.ranges.resolve = async (request) => {
        const snapshot = await read(request);
        await new Promise<void>((resolve) => { finish = resolve; });
        return snapshot;
      };
      enter(engine);
      await settle();
      suppressed = true;
      finish();
      await settle();
      expect(engine.inserts).toHaveLength(0);
    });

    it('חץ או עכבר אחרי ההמתנה אינם מחליפים, כמו ב-Word', async () => {
      const engine = fakeEngine('');
      install(engine, KEYS);
      type(engine, 'א--');
      await settle();
      key('ArrowLeft');
      key('Enter');
      await settle();
      expect(engine.inserts).toHaveLength(0);
    });

    it('טקסט שזז מאז ההערכה אינו מוחלף', async () => {
      const engine = fakeEngine('');
      install(engine, KEYS);
      type(engine, 'א--');
      await settle();
      engine.model.text = 'בא--';
      enter(engine);
      await settle();
      expect(engine.inserts).toHaveLength(0);
    });
  });

  it('Backspace מבטל בדיקה שכבר מתוזמנת', async () => {
    const engine = fakeEngine('');
    install(engine);

    type(engine, '(c)');
    key('Backspace');
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('חץ מאפס את הספירה ומבטל קריאה תלויה', async () => {
    const engine = fakeEngine('');
    install(engine);

    type(engine, '(c)');
    key('ArrowLeft');
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('לחיצת עכבר מאפסת: מה שלפני הלחיצה אינו „הוקלד”', async () => {
    const engine = fakeEngine('');
    install(engine);

    type(engine, '(c)');
    container.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('Backspace מקטין את הספירה ואינו מאפס אותה', async () => {
    const engine = fakeEngine('');
    install(engine, { replacements: [{ from: 'teh', to: 'the' }] });

    type(engine, 'tex');
    key('Backspace');
    engine.model.text = 'te';
    engine.model.caret = 2;
    type(engine, 'h ');
    await settle();

    expect(engine.model.text).toBe('the ');
  });

  it('מתג כבוי: אין תיקון ואין קריאה', async () => {
    const engine = fakeEngine('');
    install(engine, { enabled: false });

    type(engine, '(c)');
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('ההגדרות נקראות בכל הערכה: הדלקה אחרי ההתקנה פועלת בלי התקנה מחדש', async () => {
    const engine = fakeEngine('');
    let enabled = false;
    handle = installAutoCorrect({
      container,
      host: engine.host as never,
      getOptions: () => ({ ...BASE, enabled }),
    });

    type(engine, '(c)');
    await settle();
    expect(engine.inserts).toHaveLength(0);

    enabled = true;
    type(engine, ' ');
    await settle();
    // המקור נשאר: הספירה כבר נצרכה בהערכה שנדחתה, והסמל מעבר לה צמוד אליה.
    expect(engine.model.text).toBe('© ');
  });

  it('בחירת טווח אינה הקלדה: אין תיקון', async () => {
    const engine = fakeEngine('');
    install(engine);

    type(engine, '(c)');
    engine.model.selectionEnd = 3;
    engine.model.caret = 0;
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('טקסט שאורכו אינו מתיישב עם ההיסטים — אין תיקון, כדי לא למחוק טקסט אחר', async () => {
    const engine = fakeEngine('');
    install(engine);
    engine.doc.ranges.resolve = async () =>
      ({ preview: { text: '(c)' }, target: { start: { offset: 0 }, end: { offset: 9 } } }) as never;

    type(engine, '(c)');
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('חלון קטום אינו מתוקן', async () => {
    const engine = fakeEngine('');
    install(engine);
    engine.doc.ranges.resolve = async () =>
      ({ preview: { text: '(c)', truncated: true }, target: { start: { offset: 0 }, end: { offset: 3 } } }) as never;

    type(engine, '(c)');
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('כשל בהכנסה אינו זורק ואינו מנסה שוב', async () => {
    const engine = fakeEngine('');
    install(engine);
    engine.doc.insert.mockImplementationOnce(async () => ({ success: false }) as never);

    type(engine, '(c)');
    await settle();

    expect(engine.model.text).toBe('(c)');
    expect(engine.doc.insert).toHaveBeenCalledTimes(1);
  });

  it('כמה תיקונים באותה הערכה: מהסוף להתחלה, וההיסטים נשארים נכונים', async () => {
    const engine = fakeEngine('');
    install(engine);

    type(engine, '(c) (r) (tm)');
    await settle();

    expect(engine.model.text).toBe('© ® ™');
    expect(engine.inserts.map((insert) => insert.value)).toEqual(['™', '®', '©']);
  });

  it('IME: אין תיקון באמצע הרכבה, ויש בסיומה', async () => {
    const engine = fakeEngine('');
    install(engine);

    container.dispatchEvent(new Event('compositionstart', { bubbles: true }));
    engine.model.text = '(c)';
    engine.model.caret = 3;
    container.dispatchEvent(new Event('input', { bubbles: true }));
    await settle();
    expect(engine.inserts).toHaveLength(0);

    container.dispatchEvent(new Event('compositionend', { bubbles: true }));
    await settle();
    expect(engine.model.text).toBe('©');
  });

  it('דיאלוג מודאלי פתוח: אין תיקון', async () => {
    const engine = fakeEngine('');
    install(engine, {}, { isSuppressed: () => true });

    type(engine, '(c)');
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('dispose מנתק את המאזינים ומבטל קריאה תלויה', async () => {
    const engine = fakeEngine('');
    const installed = install(engine);

    type(engine, '(c)');
    installed.dispose();
    await settle();
    type(engine, 'x');
    await settle();

    expect(engine.inserts).toHaveLength(0);
  });

  it('מנוע בלי Document API: אינו זורק', async () => {
    handle = installAutoCorrect({
      container,
      host: {} as never,
      getOptions: () => BASE,
    });

    expect(() => {
      key('a');
      container.dispatchEvent(new Event('input', { bubbles: true }));
    }).not.toThrow();
    await settle();
  });
});

describe('installAutoCorrect — מנוע שאינו משגר input', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    container = document.createElement('div');
    document.body.append(container);
  });
  afterEach(() => {
    handle?.dispose();
    handle = null;
    container.remove();
    vi.useRealTimers();
    vi.restoreAllMocks();
  });

  /** הקלדה בלי אירוע input: רק keydown ו-keyup, כמו מנוע שמטפל בהקלדה בעצמו. */
  function typeWithoutInput(engine: ReturnType<typeof fakeEngine>, text: string): void {
    for (const char of text) {
      key(char);
      engine.model.text += char;
      engine.model.caret += 1;
      container.dispatchEvent(new KeyboardEvent('keyup', { key: char, bubbles: true }));
    }
  }

  it('keyup לבדו מתזמן הערכה — התיקון אינו תלוי ב-input', async () => {
    const engine = fakeEngine('');
    install(engine);

    typeWithoutInput(engine, '(c)');
    await settle();

    expect(engine.model.text).toBe('©');
  });

  it('ה-doc נקרא בכל הערכה: מנוע שהתאפס אחרי ההתקנה עדיין מתוקן', async () => {
    const engine = fakeEngine('');
    const host: { activeEditor: { doc: unknown } | null } = { activeEditor: null };
    handle = installAutoCorrect({
      container,
      host: host as never,
      getOptions: () => BASE,
    });

    host.activeEditor = { doc: engine.doc };
    typeWithoutInput(engine, '(c)');
    await settle();

    expect(engine.model.text).toBe('©');
  });

  it('מצב אבחון: כל שלב מדווח, וסיבת הדחייה כתובה', async () => {
    const info = vi.spyOn(console, 'info').mockImplementation(() => {});
    const engine = fakeEngine('');
    handle = installAutoCorrect({
      container,
      host: engine.host as never,
      getOptions: () => BASE,
      debug: () => true,
    });
    engine.doc.ranges.resolve = async () =>
      ({ preview: { text: '(c)' }, target: { start: { offset: 0 }, end: { offset: 9 } } }) as never;

    typeWithoutInput(engine, '(c)');
    await settle();

    const stages = info.mock.calls.map((call) => String(call[1]));
    expect(stages).toContain('הקשה');
    expect(stages).toContain('הערכה');
    const refusal = info.mock.calls.find((call) => call[1] === 'לא נקרא טקסט');
    expect(String(refusal?.[2])).toContain('ההיסטים אינם מתיישבים');
  });

  it('בלי מצב אבחון הקונסולה שקטה', async () => {
    const info = vi.spyOn(console, 'info').mockImplementation(() => {});
    const engine = fakeEngine('');
    install(engine);

    typeWithoutInput(engine, '(c)');
    await settle();

    expect(info).not.toHaveBeenCalled();
  });

  it('הכנסה שהמנוע דחה מזהירה בקונסולה ואינה שקטה', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const engine = fakeEngine('');
    install(engine);
    engine.doc.insert.mockImplementationOnce(async () => ({ success: false, failure: { code: 'INVALID_TARGET' } }) as never);

    typeWithoutInput(engine, '(c)');
    await settle();

    expect(warn).toHaveBeenCalled();
  });
});

describe('touchesTyped', () => {
  const fix = (end: number) => ({ start: end - 1, end, original: 'x', replacement: 'y', rule: 'table' as const });

  it('תיקון בתוך האזור שהוקלד, או צמוד אליו, נחשב; רחוק ממנו — לא', () => {
    // טקסט באורך 10, הוקלדו 3 תווים → האזור מתחיל ב-7.
    expect(touchesTyped(fix(10), 10, 3)).toBe(true);
    expect(touchesTyped(fix(7), 10, 3)).toBe(true);
    expect(touchesTyped(fix(6), 10, 3)).toBe(false);
  });
});

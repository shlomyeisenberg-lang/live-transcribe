/**
 * `End` בשורה עברית — הבאג שדווח („הנקודה עוברת לתחילת המסמך”) והתיקון.
 *
 * המספרים בבדיקות אינם מומצאים: הם מה שנמדד על ה-dist הארוז ב-Chrome אמיתי
 * (superdoc 2.12.0) — פסקה עברית שגולשת לשתי שורות, `data-pm-start`/
 * `data-pm-end` של השורות המצוירות 1..99 ו-99..192, ופסקה אנגלית שבה המנוע
 * עונה נכון (92 ו-157). הפירוט המלא בהערת הפתיחה של engine/rtl-line-end.ts.
 */
import { describe, it, expect, vi } from 'vitest';
import {
  installRtlLineEnd,
  isLineEndKey,
  lineEndOffset,
  type LineEndHost,
  type PaintedLine,
} from '../../src/engine/rtl-line-end';

/** הפסקה העברית שנמדדה: שתי שורות מצוירות, ה-fragment מתחיל ב-pm 1. */
const HEBREW: PaintedLine[] = [
  { pmStart: 1, pmEnd: 99, rtl: true },
  { pmStart: 99, pmEnd: 192, rtl: true },
];

/** אותה פסקה באנגלית, שבה `End` של המנוע תקין — ולכן אינה שלנו. */
const ENGLISH: PaintedLine[] = [
  { pmStart: 193, pmEnd: 285, rtl: false },
  { pmStart: 285, pmEnd: 350, rtl: false },
];

describe('isLineEndKey', () => {
  const key = (over: Partial<KeyboardEvent> = {}) =>
    ({ key: 'End', ctrlKey: false, metaKey: false, altKey: false, ...over }) as KeyboardEvent;

  it('End לבדו, וגם עם Shift', () => {
    expect(isLineEndKey(key())).toBe(true);
    expect(isLineEndKey(key({ shiftKey: true }))).toBe(true);
  });

  it('Ctrl+End נמדד תקין במנוע, ולכן אינו שלנו', () => {
    expect(isLineEndKey(key({ ctrlKey: true }))).toBe(false);
    expect(isLineEndKey(key({ metaKey: true }))).toBe(false);
    expect(isLineEndKey(key({ altKey: true }))).toBe(false);
  });

  it('Home נמדד תקין, ואינו נוגע', () => {
    expect(isLineEndKey(key({ key: 'Home' }))).toBe(false);
  });

  it('הרכבה במקלדת (IME): `End` שייך לחלונית ההרכבה ולא לסמן', () => {
    /*
     * ‏`End` בתוך הרכבה מזיז בתוך המועמד, לא במסמך — והמסלול כאן בולע את
     * האירוע ב-`stopImmediatePropagation`, כלומר חלונית ההרכבה לא הייתה רואה
     * אותו כלל. שני החצאים של התקן בריפו (`src/ui/shortcuts/match.ts`):
     * דפדפן שאינו מציב `isComposing` מדווח `keyCode === 229`.
     */
    expect(isLineEndKey(key({ isComposing: true }))).toBe(false);
    expect(isLineEndKey(key({ keyCode: 229 }))).toBe(false);
    expect(isLineEndKey(key({ isComposing: false, keyCode: 35 })), 'בלי הרכבה — שלנו').toBe(true);
  });
});

describe('lineEndOffset', () => {
  it('שורה עברית ראשונה: הסמן באמצע → סוף השורה, ולא 0 כמו שהמנוע עושה', () => {
    expect(lineEndOffset(HEBREW, 1, 48)).toBe(98);
  });

  it('שורה עברית שנייה: סוף השורה, ולא תחילתה', () => {
    expect(lineEndOffset(HEBREW, 1, 144)).toBe(191);
  });

  it('שורה אנגלית — `null`, המנוע מטפל בעצמו', () => {
    expect(lineEndOffset(ENGLISH, 193, 45)).toBeNull();
    expect(lineEndOffset(ENGLISH, 193, 141)).toBeNull();
  });

  it('שורה אנגלית בתוך פסקה עברית אינה נגררת אחרי הפסקה', () => {
    const mixed: PaintedLine[] = [
      { pmStart: 1, pmEnd: 99, rtl: true },
      { pmStart: 99, pmEnd: 192, rtl: false },
    ];
    expect(lineEndOffset(mixed, 1, 48), 'השורה העברית').toBe(98);
    expect(lineEndOffset(mixed, 1, 144), 'השורה האנגלית').toBeNull();
  });

  it('תפר בין שתי שורות נקרא כסופה של הקודמת — `End` חוזר אינו מטייל', () => {
    // אילו התפר היה נקרא כתחילת השורה הבאה, כל הקשה נוספת הייתה מעבירה את
    // הסמן לסוף השורה שאחריה. ראו „גבול גלישה” בהערת הפתיחה.
    expect(lineEndOffset(HEBREW, 1, 98)).toBe(98);
  });

  it('תפר אחרי Home שייך לתחילת השורה הבאה', () => {
    expect(lineEndOffset(HEBREW, 1, 98, true)).toBe(191);
  });

  it('בסוף הפסקה — מוחזר אותו היסט, כלומר ההקשה עדיין שלנו', () => {
    // חשוב: `null` כאן היה מחזיר את ההקשה למנוע, והוא היה מקפיץ לתחילת השורה.
    expect(lineEndOffset(HEBREW, 1, 191)).toBe(191);
  });

  it('היסט שאינו בשום שורה — `null`', () => {
    expect(lineEndOffset(HEBREW, 1, 500)).toBeNull();
    expect(lineEndOffset([], 1, 0)).toBeNull();
  });
});

/* ------------------------------------------------------------------ */
/* ההתקנה על ה-host                                                    */
/* ------------------------------------------------------------------ */

interface FakeSetup {
  host: HTMLElement;
  line: HTMLElement;
  superdoc: LineEndHost;
  setSelectionTarget: ReturnType<typeof vi.fn>;
  snapshot: { blockId: string; start: number; end: number; coordinateSpace?: string };
}

/** מסמך מצויר מזערי בצורה שנמדדה: fragment עם שורה אחת בתוכו. */
function setup({
  rtl = true,
  lines = [{ pmStart: 1, pmEnd: 99 }, { pmStart: 99, pmEnd: 192 }],
  coordinateSpace,
}: {
  rtl?: boolean;
  lines?: { pmStart: number; pmEnd: number }[];
  coordinateSpace?: string;
} = {}): FakeSetup {
  const host = document.createElement('div');
  const fragment = document.createElement('div');
  fragment.setAttribute('data-source-node-id', 'p1');
  fragment.setAttribute('data-pm-start', '1');
  fragment.setAttribute('data-pm-end', '192');

  let first: HTMLElement | null = null;
  for (const { pmStart, pmEnd } of lines) {
    const line = document.createElement('div');
    line.setAttribute('data-pm-start', String(pmStart));
    line.setAttribute('data-pm-end', String(pmEnd));
    if (rtl) line.setAttribute('dir', 'rtl');
    // הריצות נושאות גם הן `data-pm-*`, ולכן חייבות להיות מסוננות: רק ילד
    // ישיר של ה-fragment הוא שורה.
    const run = document.createElement('span');
    run.setAttribute('data-pm-start', String(pmStart));
    run.setAttribute('data-pm-end', String(pmEnd));
    line.append(run);
    fragment.append(line);
    first ??= line;
  }
  host.append(fragment);
  document.body.append(host);

  const snapshot = { blockId: 'p1', start: 48, end: 48, ...(coordinateSpace ? { coordinateSpace } : {}) };
  const setSelectionTarget = vi.fn();
  const story = { kind: 'story', storyType: 'body' };
  const superdoc: LineEndHost = {
    activeEditor: {
      host: {
        readLiveSelectionSyncSnapshot: () => ({
          selectionTarget: {
            kind: 'selection',
            start: { kind: 'text', blockId: snapshot.blockId, offset: snapshot.start, story },
            end: { kind: 'text', blockId: snapshot.blockId, offset: snapshot.end, story },
            story,
            ...(snapshot.coordinateSpace ? { coordinateSpace: snapshot.coordinateSpace } : {}),
          },
        }),
      },
      authoring: { setSelectionTarget },
    },
  };

  return { host, line: first!, superdoc, setSelectionTarget, snapshot };
}

function pressKey(target: HTMLElement, key: string, init: KeyboardEventInit = {}): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
}

function pressEnd(target: HTMLElement, init: KeyboardEventInit = {}): KeyboardEvent {
  return pressKey(target, 'End', init);
}

describe('installRtlLineEnd', () => {
  it('בשורה עברית: ההקשה נבלעת, והסמן נקבע לסוף השורה', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const handle = installRtlLineEnd({ host, superdoc });

    const event = pressEnd(line);

    expect(event.defaultPrevented, 'המנוע לא יראה את ההקשה').toBe(true);
    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    const input = setSelectionTarget.mock.calls[0]![0] as {
      target: { start: { offset: number }; end: { offset: number } };
    };
    expect(input.target.end.offset).toBe(98);
    expect(input.target.start.offset, 'בלי Shift הבחירה מתכווצת').toBe(98);
    handle.dispose();
  });

  it('בתוך הרכבה: ההקשה אינה נבלעת ואינה מתורגמת לתנועת סמן', () => {
    // בלי זה `End` באמצע הרכבה נעצר כאן ב-`stopImmediatePropagation`, וחלונית
    // ההרכבה — שמקש זה שייך לה — לא רואה אותו. אותה הגנה כבר קיימת בחצים.
    for (const composing of [{ isComposing: true }, { keyCode: 229 }]) {
      const { host, line, superdoc, setSelectionTarget } = setup();
      const handle = installRtlLineEnd({ host, superdoc });

      const event = pressEnd(line, composing);

      expect(event.defaultPrevented, JSON.stringify(composing)).toBe(false);
      expect(setSelectionTarget).not.toHaveBeenCalled();
      handle.dispose();
      document.body.innerHTML = '';
    }
  });

  it('‏`Home` של הרכבה אינו מזיז את רמז התפר', () => {
    /*
     * הרמז שייך ל-`Home` **של המשתמש**. הסמן כאן יושב בתפר (98), ושם ורק שם
     * הרמז משנה: בלי רמז ה-`End` חוזר לסוף השורה הראשונה (98), ועם רמז הוא
     * הולך לסוף השנייה (191). ‏`Home` שנשלח בתוך הרכבה שייך לחלונית ההרכבה,
     * ואם היה מציב את הרמז — ה-`End` שאחריו היה קופץ שורה.
     */
    const { host, line, superdoc, setSelectionTarget, snapshot } = setup();
    snapshot.start = 98;
    snapshot.end = 98;
    const handle = installRtlLineEnd({ host, superdoc });

    pressKey(line, 'Home', { isComposing: true });
    const event = pressEnd(line);

    expect(event.defaultPrevented).toBe(true);
    expect(setSelectionTarget, 'הסמן כבר בסוף השורה הראשונה — אין מה לכתוב').not.toHaveBeenCalled();
    handle.dispose();
  });

  it('ה-`story` נלקח מהתצלום, ואין ברירת מחדל מומצאת', () => {
    /*
     * הערך שהיה כאן — `{kind:'story',storyType:'body'}` — מעולם לא נמדד
     * כנכון, והיה נכתב למנוע דווקא במצב שבו המנוע עצמו לא דיווח story
     * (הערת שוליים? כותרת?). השער (`check:line-end`, „ה-story בתצלום”) מודד
     * שהשדה כן מגיע; כאן נמדד מה קורה בתצלום ההיפותטי שאין בו — המודול מעביר
     * ‏`undefined` הלאה, ואינו ממציא גוף מסמך.
     */
    const { host, line, superdoc, setSelectionTarget } = setup();
    superdoc.activeEditor!.host!.readLiveSelectionSyncSnapshot = () => ({
      selectionTarget: {
        kind: 'selection',
        start: { kind: 'text', blockId: 'p1', offset: 48 },
        end: { kind: 'text', blockId: 'p1', offset: 48 },
      },
    });
    const handle = installRtlLineEnd({ host, superdoc });

    pressEnd(line);

    const input = setSelectionTarget.mock.calls[0]![0] as {
      target: { story: unknown; end: { story: unknown } };
    };
    expect(input.target.story).toBeUndefined();
    expect(input.target.end.story).toBeUndefined();
    handle.dispose();
  });

  it('בשורה אנגלית: ההקשה ממשיכה למנוע כפי שהיא', () => {
    const { host, line, superdoc, setSelectionTarget } = setup({ rtl: false });
    const handle = installRtlLineEnd({ host, superdoc });

    const event = pressEnd(line);

    expect(event.defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('Shift+End מרחיב מהעוגן ולא בוחר אחורה', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const handle = installRtlLineEnd({ host, superdoc });

    pressEnd(line, { shiftKey: true });

    const input = setSelectionTarget.mock.calls[0]![0] as {
      target: { start: { offset: number }; end: { offset: number } };
    };
    expect(input.target.start.offset, 'העוגן נשאר במקום שהיה').toBe(48);
    expect(input.target.end.offset).toBe(98);
    handle.dispose();
  });

  it('שומר coordinateSpace של בחירת tracked', () => {
    const { host, line, superdoc, setSelectionTarget } = setup({ coordinateSpace: 'tracked' });
    const handle = installRtlLineEnd({ host, superdoc });

    pressEnd(line);

    const input = setSelectionTarget.mock.calls[0]![0] as {
      target: { coordinateSpace?: string };
    };
    expect(input.target.coordinateSpace).toBe('tracked');
    handle.dispose();
  });

  it('דחיית כתיבת הבחירה אינה יוצרת rejection לא מטופלת', async () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    setSelectionTarget.mockRejectedValueOnce(new Error('editor disposed'));
    const handle = installRtlLineEnd({ host, superdoc });

    pressEnd(line);
    await Promise.resolve();

    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('`End` פעמיים: השנייה נבלעת אך אינה מזיזה לשורה הבאה', () => {
    const { host, line, superdoc, setSelectionTarget, snapshot } = setup();
    const handle = installRtlLineEnd({ host, superdoc });

    pressEnd(line);
    // המנוע הזיז את הסמן לאן שביקשנו, וההקשה הבאה יוצאת משם.
    snapshot.start = 98;
    snapshot.end = 98;
    const second = pressEnd(line);

    expect(second.defaultPrevented, 'עדיין שלנו — אחרת המנוע יקפיץ לתחילת השורה').toBe(true);
    expect(setSelectionTarget, 'ואין כתיבה שנייה').toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('Home ואז End בתפר מגיעים לסוף השורה השנייה', () => {
    const { host, line, superdoc, setSelectionTarget, snapshot } = setup();
    const handle = installRtlLineEnd({ host, superdoc });
    snapshot.start = 98;
    snapshot.end = 98;

    pressKey(line, 'Home');
    pressEnd(line);

    const input = setSelectionTarget.mock.calls[0]![0] as {
      target: { end: { offset: number } };
    };
    expect(input.target.end.offset).toBe(191);
    handle.dispose();
  });

  it('לחיצת עכבר אחרי Home מבטלת את הרמז על התפר', () => {
    const { host, line, superdoc, setSelectionTarget, snapshot } = setup();
    const handle = installRtlLineEnd({ host, superdoc });
    snapshot.start = 98;
    snapshot.end = 98;

    pressKey(line, 'Home');
    line.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    const event = pressEnd(line);

    expect(event.defaultPrevented).toBe(true);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('בלי `authoring` או בלי תצלום סינכרוני — ההקשה נשארת של המנוע', () => {
    const { host, line } = setup();
    const handle = installRtlLineEnd({ host, superdoc: { activeEditor: {} } });

    expect(pressEnd(line).defaultPrevented).toBe(false);
    handle.dispose();
  });

  it('פסקה שאינה מצוירת (אין fragment תואם) — ההקשה נשארת של המנוע', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    host.querySelector('[data-source-node-id]')?.setAttribute('data-source-node-id', 'other');
    const handle = installRtlLineEnd({ host, superdoc });

    expect(pressEnd(line).defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('אחרי `dispose` המאזין אינו נוגע יותר', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    installRtlLineEnd({ host, superdoc }).dispose();

    expect(pressEnd(line).defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
  });

  it('בלי host אין התקנה, ואין קריסה', () => {
    expect(() => installRtlLineEnd({ host: null, superdoc: null }).dispose()).not.toThrow();
  });
});

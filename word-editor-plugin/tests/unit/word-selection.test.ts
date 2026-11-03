/**
 * לחיצה כפולה שבוחרת מילה שלמה.
 *
 * שלוש שכבות נמדדות כאן, וכל אחת עונה על שאלה אחרת:
 *
 *   1. **גבול המילה** — `wordBoundsIn` היא פונקציה טהורה, ולכן הניקוד,
 *      הטעמים, הגרשיים, המקף והמרכאות נמדדים בלי DOM ובלי מנוע. זה הלב:
 *      הליקוי שנמדד במנוע הוא שסימני ניקוד נספרים אצלו כמפרידי מילה.
 *   2. **המסלול מול המנוע** — `selectWordAtSelection` מול כפיל, כדי לוודא
 *      שהוא שולח את החלון הנכון, מרכיב `SelectionTarget` תקין, **ולא** נוגע
 *      בבחירה כשאין לו תשובה בטוחה.
 *   3. **המאזין** — `installWordSelection` על אלמנט אמיתי ב-jsdom, כולל
 *      המקרה שהמאזין נועד לו: לחיצה שלישית שמגיעה באמצע העבודה
 *      הא-סינכרונית.
 *
 * הטקסטים כאן הם מה שנמדד בפועל ב-Chrome מול המנוע: הפסוק המנוקד והמוטעם
 * שהלחיצה הכפולה בחרה בו אות אחת בלבד.
 */
import { describe, it, expect, afterEach, vi } from 'vitest';
import {
  installWordSelection,
  markSyntheticPointer,
  selectBlockAtSelection,
  selectWordAtSelection,
  wordBoundsIn,
  BLOCK_LENGTH_PROBE,
  WORD_WINDOW_RADIUS,
  type TextWindow,
  type WordSelectionHandle,
  type WordSelectionHost,
} from '../../src/engine/word-selection';

/** פסוק מנוקד ומוטעם — בדיוק הטקסט שנמדד. */
const VERSE = 'בְּרֵאשִׁ֖ית בָּרָ֣א אֱלֹהִ֑ים אֵ֥ת הַשָּׁמַ֖יִם';

/** חלון שהוא הבלוק כולו: שני קצותיו הם קצות הבלוק. */
function whole(text: string, base = 0): TextWindow {
  return { text, base, atBlockStart: base === 0, atBlockEnd: true };
}

/** גבול המילה שבמיקום `at`, כמחרוזת — כדי שהציפייה תהיה קריאה. */
function wordAt(text: string, at: number, window: TextWindow = whole(text)): string | null {
  const bounds = wordBoundsIn(window, at + window.base);
  return bounds === null ? null : text.slice(bounds.start - window.base, bounds.end - window.base);
}

describe('גבול המילה', () => {
  it('מילה עברית לא מנוקדת', () => {
    const text = 'שלום עולם גדול';
    expect(wordAt(text, text.indexOf('עולם') + 1)).toBe('עולם');
  });

  it('מילה מנוקדת ומוטעמת — הליקוי שבגללו המודול קיים', () => {
    const word = 'בָּרָ֣א';
    const at = VERSE.indexOf(word);
    // המנוע בוחר כאן אות אחת, ולכן נמדד גם המצב שהזרע הוא אמצע המילה.
    expect(wordAt(VERSE, at + 2)).toBe(word);
    expect(wordAt(VERSE, at)).toBe(word);
  });

  it('כל מילה בפסוק מוחזרת שלמה, מכל היסט שבתוכה', () => {
    for (const word of VERSE.split(' ')) {
      const at = VERSE.indexOf(word);
      for (let offset = 0; offset < word.length; offset += 1) {
        expect(wordAt(VERSE, at + offset)).toBe(word);
      }
    }
  });

  it('גרשיים באמצע מילה הם חלק ממנה, בשתי צורות התו', () => {
    for (const text of ['רמב״ם הלכות תשובה', 'רמב"ם הלכות תשובה']) {
      expect(wordAt(text, 2)).toBe(text.slice(0, 5));
    }
  });

  it('גרש בסוף מילה הוא חלק ממנה', () => {
    expect(wordAt('ר׳ יוחנן אמר', 0)).toBe('ר׳');
    expect(wordAt("ר' יוחנן אמר", 0)).toBe("ר'");
  });

  it('מרכאות שעוטפות ציטוט אינן חלק מהמילה', () => {
    const text = 'אמר "שלום" לכולם';
    expect(wordAt(text, text.indexOf('שלום') + 1)).toBe('שלום');
  });

  it('מקף מפריד בין שתי מילים, כמו ב-Word', () => {
    const text = 'עַל־כֵּן יֵלֵךְ';
    expect(wordAt(text, 1)).toBe('עַל');
    expect(wordAt(text, text.indexOf('כֵּן') + 1)).toBe('כֵּן');
  });

  it('סוף פסוק ופסק אינם נבלעים במילה', () => {
    const text = 'הָאָֽרֶץ׃ וְהָאָ֗רֶץ';
    expect(wordAt(text, 1)).toBe('הָאָֽרֶץ');
  });

  it('היסט שאינו על מילה מחזיר null', () => {
    const text = 'שלום עולם';
    expect(wordAt(text, text.indexOf(' '))).toBe('שלום');
    expect(wordBoundsIn(whole('   '), 1)).toBeNull();
    expect(wordBoundsIn(whole(''), 0)).toBeNull();
  });

  it('היסט שמצביע לאחרי התו האחרון עדיין בוחר את המילה', () => {
    // זה בדיוק קצה הזרע: המנוע מחזיר טווח, וקצהו יושב אחרי התו האחרון.
    expect(wordAt('שלום עולם', 4)).toBe('שלום');
  });

  it('לטינית וספרות', () => {
    expect(wordAt('פרק abc123 ראשון', 5)).toBe('abc123');
  });

  it('מילה שנוגעת בקצה חלון שאינו קצה בלוק אינה מוחזרת', () => {
    const cut: TextWindow = { text: 'עולם גדול', base: 40, atBlockStart: false, atBlockEnd: true };
    // „עולם” פותחת את החלון, כלומר ייתכן שהיא נמשכת אחורה מעבר למה שנקרא.
    expect(wordBoundsIn(cut, 41)).toBeNull();
    // „גדול” סוגרת אותו, והקצה הזה כן קצה בלוק.
    expect(wordBoundsIn(cut, 46)).toEqual({ start: 45, end: 49 });
  });

  it('מילה שנוגעת בקצה שהוא קצה הבלוק מוחזרת', () => {
    const window: TextWindow = { text: 'עולם גדול', base: 0, atBlockStart: true, atBlockEnd: false };
    expect(wordBoundsIn(window, 1)).toEqual({ start: 0, end: 4 });
    expect(wordBoundsIn(window, 6)).toBeNull();
  });

  it('ההיסטים שמוחזרים הם של הבלוק ולא של החלון', () => {
    const window: TextWindow = { text: ' שלום עולם ', base: 100, atBlockStart: false, atBlockEnd: false };
    expect(wordBoundsIn(window, 102)).toEqual({ start: 101, end: 105 });
  });
});

/** כפיל של המנוע: בחירה נתונה, טקסט נתון, ו-`apply` שנרשם. */
function engineDouble(options: {
  text: string;
  /** הטווח שהמנוע „בחר” בלחיצה הכפולה, בהיסטים של הבלוק. */
  seed: { start: number; end: number };
  story?: unknown;
  coordinateSpace?: string;
  truncated?: boolean;
  blockId?: string;
}): {
  host: WordSelectionHost;
  applied: unknown[];
  requests: unknown[];
} {
  const blockId = options.blockId ?? 'B1';
  const applied: unknown[] = [];
  const requests: unknown[] = [];
  const point = (offset: number) => ({
    kind: 'text',
    blockId,
    offset,
    ...(options.story ? { story: options.story } : {}),
  });

  const host: WordSelectionHost = {
    activeEditor: {
      doc: {
        selection: {
          current: () => ({
            selectionTarget: {
              kind: 'selection',
              start: point(options.seed.start),
              end: point(options.seed.end),
              ...(options.story ? { story: options.story } : {}),
              ...(options.coordinateSpace ? { coordinateSpace: options.coordinateSpace } : {}),
            },
          }),
        },
        ranges: {
          resolve: (input: unknown) => {
            requests.push(input);
            const request = input as {
              start: { point: { offset: number } };
              end: { point: { offset: number } };
            };
            const from = Math.max(0, request.start.point.offset);
            // המנוע חותך את הקצה לאורך הבלוק — וזה מה שמסמן „הגענו לסוף”.
            const to = Math.min(options.text.length, request.end.point.offset);
            return {
              preview: { text: options.text.slice(from, to), truncated: options.truncated === true },
              target: { start: { offset: from }, end: { offset: to } },
            };
          },
        },
      },
    },
    ui: { selection: { apply: (target: unknown) => applied.push(target) } },
  };

  return { host, applied, requests };
}

describe('הרחבת הבחירה מול המנוע', () => {
  it('מרחיבה זרע של אות אחת למילה מנוקדת שלמה', async () => {
    const word = 'בָּרָ֣א';
    const at = VERSE.indexOf(word);
    const { host, applied } = engineDouble({ text: VERSE, seed: { start: at + 2, end: at + 3 } });

    await expect(selectWordAtSelection(host)).resolves.toBe(true);
    expect(applied).toEqual([
      {
        kind: 'selection',
        start: { kind: 'text', blockId: 'B1', offset: at },
        end: { kind: 'text', blockId: 'B1', offset: at + word.length },
      },
    ]);
  });

  it('החלון נקרא ברדיוס סביב הזרע ואינו יורד מתחת לאפס', async () => {
    const { host, requests } = engineDouble({ text: VERSE, seed: { start: 2, end: 3 } });
    await selectWordAtSelection(host);
    expect(requests).toEqual([
      {
        start: { kind: 'point', point: { kind: 'text', blockId: 'B1', offset: 0 } },
        end: {
          kind: 'point',
          point: { kind: 'text', blockId: 'B1', offset: 3 + WORD_WINDOW_RADIUS },
        },
      },
    ]);
  });

  it('אינה נוגעת בבחירה כשהמנוע כבר בחר את המילה', async () => {
    const text = 'שלום עולם גדול';
    const at = text.indexOf('עולם');
    const { host, applied } = engineDouble({ text, seed: { start: at, end: at + 4 } });

    await expect(selectWordAtSelection(host)).resolves.toBe(false);
    expect(applied).toEqual([]);
  });

  it('אינה מרחיבה מזרע ארוך מהחלון — זו בחירה קודמת ולא לחיצה', async () => {
    const long = `${VERSE} ${VERSE}`;
    const { host, applied } = engineDouble({
      text: long,
      seed: { start: 0, end: WORD_WINDOW_RADIUS + 1 },
    });

    await expect(selectWordAtSelection(host)).resolves.toBe(false);
    expect(applied).toEqual([]);
  });

  it('אינה נוגעת בבחירה כשהטקסט שהוחזר נחתך', async () => {
    const at = VERSE.indexOf('בָּרָ֣א');
    const { host, applied } = engineDouble({
      text: VERSE,
      seed: { start: at + 2, end: at + 3 },
      truncated: true,
    });

    await expect(selectWordAtSelection(host)).resolves.toBe(false);
    expect(applied).toEqual([]);
  });

  it('אינה נוגעת בבחירה במרחב היסטים שאינו הנראה', async () => {
    const at = VERSE.indexOf('בָּרָ֣א');
    const { host, applied } = engineDouble({
      text: VERSE,
      seed: { start: at + 2, end: at + 3 },
      coordinateSpace: 'tracked',
    });

    await expect(selectWordAtSelection(host)).resolves.toBe(false);
    expect(applied).toEqual([]);
  });

  it('ה-story נשמר גם בחלון וגם בבחירה שנקבעת', async () => {
    const story = { kind: 'story', storyType: 'header' };
    const at = VERSE.indexOf('בָּרָ֣א');
    const { host, applied, requests } = engineDouble({
      text: VERSE,
      seed: { start: at + 2, end: at + 3 },
      story,
    });

    await expect(selectWordAtSelection(host)).resolves.toBe(true);
    expect((requests[0] as { in?: unknown }).in).toEqual(story);
    expect(applied[0]).toMatchObject({ story, start: { story }, end: { story } });
  });

  it('`isStale` מבטל לפני שהבחירה נשלחת', async () => {
    const at = VERSE.indexOf('בָּרָ֣א');
    const { host, applied } = engineDouble({ text: VERSE, seed: { start: at + 2, end: at + 3 } });

    await expect(selectWordAtSelection(host, { isStale: () => true })).resolves.toBe(false);
    expect(applied).toEqual([]);
  });

  it('אינה זורקת כשהמנוע זורק', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const host: WordSelectionHost = {
      activeEditor: {
        doc: {
          selection: {
            current: () => {
              throw new Error('נפל');
            },
          },
          ranges: { resolve: () => undefined },
        },
      },
      ui: { selection: { apply: () => undefined } },
    };

    await expect(selectWordAtSelection(host)).resolves.toBe(false);
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
  });

  it('מארח שאין בו את הפעולות מחזיר false בלי לזרוק', async () => {
    await expect(selectWordAtSelection(null)).resolves.toBe(false);
    await expect(selectWordAtSelection({})).resolves.toBe(false);
    await expect(selectWordAtSelection({ activeEditor: { doc: {} } })).resolves.toBe(false);
  });
});

let installed: WordSelectionHandle | null = null;

afterEach(() => {
  installed?.dispose();
  installed = null;
  document.body.innerHTML = '';
});

/** ממתין לשתי הקריאות הא-סינכרוניות של המסלול. */
function flush(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function container(): HTMLElement {
  const element = document.createElement('div');
  document.body.append(element);
  return element;
}

interface ClickAt {
  x?: number;
  y?: number;
  button?: number;
}

function press(target: HTMLElement, { x = 10, y = 10, button = 0 }: ClickAt = {}): void {
  target.dispatchEvent(
    new MouseEvent('mousedown', { bubbles: true, button, clientX: x, clientY: y }),
  );
}

function release(target: HTMLElement, { x = 10, y = 10, button = 0 }: ClickAt = {}): void {
  target.dispatchEvent(new MouseEvent('click', { bubbles: true, button, clientX: x, clientY: y }));
}

/** רצף של `times` לחיצות באותו מקום, כמו משתמש שלוחץ מהר. */
function clickTimes(target: HTMLElement, times: number, at: ClickAt = {}): void {
  for (let i = 0; i < times; i += 1) {
    press(target, at);
    release(target, at);
  }
}

/** מקימה מאזין מול כפיל שהזרע שלו הוא אות אחת בתוך מילה מנוקדת. */
function installed2(): { root: HTMLElement; applied: unknown[] } {
  const at = VERSE.indexOf('בָּרָ֣א');
  const { host, applied } = engineDouble({ text: VERSE, seed: { start: at + 2, end: at + 3 } });
  const root = container();
  installed = installWordSelection(root, host);
  return { root, applied };
}

describe('המאזין על ה-container', () => {
  it('שתי לחיצות בוחרות את המילה', async () => {
    const { root, applied } = installed2();
    clickTimes(root, 2);
    await flush();

    const word = 'בָּרָ֣א';
    const at = VERSE.indexOf(word);
    expect(applied).toEqual([
      {
        kind: 'selection',
        start: { kind: 'text', blockId: 'B1', offset: at },
        end: { kind: 'text', blockId: 'B1', offset: at + word.length },
      },
    ]);
  });

  it('שלוש לחיצות בוחרות את הפסקה כולה', async () => {
    const { root, applied } = installed2();
    clickTimes(root, 3);
    await flush();

    // הלחיצה השנייה בחרה מילה, השלישית מרחיבה לפסקה — וזו האחרונה שקובעת.
    expect(applied[applied.length - 1]).toEqual({
      kind: 'selection',
      start: { kind: 'text', blockId: 'B1', offset: 0 },
      end: { kind: 'text', blockId: 'B1', offset: VERSE.length },
    });
  });

  /**
   * תפריט ההקשר מזיז את הסמן בלחיצה שהוא עצמו משגר, מפני שהמנוע מתעלם מכפתור
   * ימני. לחיצה כזאת אינה של אדם, ואסור שתיספר: בלי זה שתי לחיצות ימניות
   * באותה נקודה היו בוחרות מילה שאיש לא ביקש.
   */
  it('לחיצה שהתוסף שיגר אינה נספרת ברצף', async () => {
    const { root, applied } = installed2();

    press(root);
    release(root);
    root.dispatchEvent(
      markSyntheticPointer(new MouseEvent('mousedown', { bubbles: true, clientX: 10, clientY: 10 })),
    );
    root.dispatchEvent(
      markSyntheticPointer(new MouseEvent('click', { bubbles: true, clientX: 10, clientY: 10 })),
    );
    await flush();

    expect(applied).toEqual([]);
  });

  it('לחיצה בודדת אינה נוגעת בבחירה', async () => {
    const { root, applied } = installed2();
    clickTimes(root, 1);
    await flush();
    expect(applied).toEqual([]);
  });

  it('שתי לחיצות במקומות שונים אינן רצף', async () => {
    const { root, applied } = installed2();
    clickTimes(root, 1, { x: 10, y: 10 });
    clickTimes(root, 1, { x: 400, y: 300 });
    await flush();
    expect(applied).toEqual([]);
  });

  it('גרירה אינה לחיצה — הבחירה של המשתמש נשארת', async () => {
    const { root, applied } = installed2();
    clickTimes(root, 1);
    // הלחיצה השנייה גוררת: היא יורדת במקום אחד ומשתחררת באחר.
    press(root, { x: 10, y: 10 });
    release(root, { x: 120, y: 10 });
    await flush();
    expect(applied).toEqual([]);
  });

  it('לחיצה נוספת שמגיעה באמצע מבטלת את ההרחבה', async () => {
    const { root, applied } = installed2();
    clickTimes(root, 2);
    // הלחיצה הבאה, לפני שההרחבה הא-סינכרונית הספיקה לחזור.
    press(root);
    await flush();
    expect(applied).toEqual([]);
  });

  it('כפתור שאינו הראשי אינו מפעיל כלום', async () => {
    const { root, applied } = installed2();
    clickTimes(root, 2, { button: 2 });
    await flush();
    expect(applied).toEqual([]);
  });

  it('`dispose` מסיר את המאזינים', async () => {
    const at = VERSE.indexOf('בָּרָ֣א');
    const { host, applied } = engineDouble({ text: VERSE, seed: { start: at + 2, end: at + 3 } });
    const root = container();
    const handle = installWordSelection(root, host);
    handle.dispose();

    clickTimes(root, 2);
    await flush();
    expect(applied).toEqual([]);
  });
});

describe('בחירת פסקה', () => {
  it('בוחרת מאפס עד אורך הבלוק', async () => {
    const { host, applied, requests } = engineDouble({ text: VERSE, seed: { start: 5, end: 5 } });
    await expect(selectBlockAtSelection(host)).resolves.toBe(true);
    expect(applied).toEqual([
      {
        kind: 'selection',
        start: { kind: 'text', blockId: 'B1', offset: 0 },
        end: { kind: 'text', blockId: 'B1', offset: VERSE.length },
      },
    ]);
    // אורך הבלוק נשאל דרך היסט שחורג ממנו, והמנוע חותך אותו לאורך האמיתי.
    expect((requests[0] as { end: { point: { offset: number } } }).end.point.offset).toBe(
      BLOCK_LENGTH_PROBE,
    );
  });

  it('אינה שולחת בחירה שנייה כשהפסקה כבר מסומנת', async () => {
    const { host, applied } = engineDouble({ text: VERSE, seed: { start: 0, end: VERSE.length } });
    await expect(selectBlockAtSelection(host)).resolves.toBe(false);
    expect(applied).toEqual([]);
  });

  it('`isStale` מבטל לפני שהבחירה נשלחת', async () => {
    const { host, applied } = engineDouble({ text: VERSE, seed: { start: 5, end: 5 } });
    await expect(selectBlockAtSelection(host, { isStale: () => true })).resolves.toBe(false);
    expect(applied).toEqual([]);
  });

  it('מארח שאין בו את הפעולות מחזיר false בלי לזרוק', async () => {
    await expect(selectBlockAtSelection(null)).resolves.toBe(false);
    await expect(selectBlockAtSelection({ activeEditor: { doc: {} } })).resolves.toBe(false);
  });
});

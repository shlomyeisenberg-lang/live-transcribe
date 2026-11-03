/**
 * הניתוב המלא: צירוף → מה שרץ. הבדיקה רצה על **הרג'יסטרי האמיתי** ועל המנתב
 * האמיתי, ומחליפה רק את שני הקצוות (מה שמריץ פקודה, ומה שמריץ פעולה).
 *
 * למה כאן ולא על המעטפת: המעטפת כבר מוכיחה שהחיווט קיים (`shortcuts-core`),
 * ומה שנשאר להוכיח הוא שכל צירוף מגיע למזהה הנכון עם ה-payload הנכון —
 * שתים-עשרה בדיקות מעטפת לאותה שאלה היו איטיות ולא מדויקות יותר. כאן כל
 * צירוף נבדק, כולל ה-payload, בלי להרכיב דבר.
 */
import { describe, it, expect, vi } from 'vitest';
import { createShortcutDispatcher } from '../../src/ui/shortcuts/dispatch';
import { SHORTCUTS, type ShellAction, type Shortcut } from '../../src/ui/shortcuts/registry';

/** הרשומות כטיפוס הרחב — `as const` הופך כל אחת לליטרל משלה. */
const ENTRIES: readonly Shortcut[] = SHORTCUTS;

interface Ran {
  commands: Array<{ id: string; payload: unknown }>;
  actions: ShellAction[];
}

function press(init: Partial<KeyboardEvent> & { code?: string; key?: string }): Ran {
  const ran: Ran = { commands: [], actions: [] };
  const dispatcher = createShortcutDispatcher({
    runCommand: (id, payload) => ran.commands.push({ id, payload }),
    runAction: (action) => {
      ran.actions.push(action);
      return true;
    },
    target: { addEventListener: () => {}, removeEventListener: () => {} },
  });

  const keyEvent = {
    key: '',
    code: '',
    ctrlKey: false,
    metaKey: false,
    shiftKey: false,
    altKey: false,
    target: null,
    preventDefault: vi.fn(),
    // הבליעה עוצרת גם את המסע אל המנוע — ראו `swallow` ב-dispatch.ts.
    stopPropagation: vi.fn(),
    ...init,
  } as unknown as KeyboardEvent;

  /*
   * שני השלבים, בסדר שבו הדפדפן מעביר אותם: capture ואז bubble. רשומות
   * `Alt` בלי `Ctrl` רצות בראשון (ראו `runsBeforeEngine`), וכל השאר בשני —
   * ובדיקה שקוראת רק ל-`handle` הייתה מדווחת על 11 רשומות חיות כמתות.
   */
  if (!dispatcher.handleCapture(keyEvent)) dispatcher.handle(keyEvent);

  dispatcher.dispose();
  return ran;
}

/**
 * באיזה שלב הרשומה רצה **בפועל**, על הרג'יסטרי האמיתי.
 *
 * שני אירועים ולא אחד: `press` מחקה את הדפדפן, ולכן ברגע שה-capture טיפל הוא
 * אינו קורא ל-`handle` כלל — כלומר הוא אינו יכול להבחין בין „רצה ב-capture”
 * לבין „רצה בשני השלבים”, וזו בדיוק ההרצה הכפולה שיש לשמור מפניה.
 */
function phasesOf(init: Partial<KeyboardEvent> & { code?: string; key?: string }): {
  capture: boolean;
  bubble: boolean;
} {
  const build = (): KeyboardEvent =>
    ({
      key: '',
      code: '',
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
      altKey: false,
      target: null,
      preventDefault: vi.fn(),
      stopPropagation: vi.fn(),
      ...init,
    }) as unknown as KeyboardEvent;

  const make = (): ReturnType<typeof createShortcutDispatcher> =>
    createShortcutDispatcher({
      runCommand: () => {},
      runAction: () => true,
      target: { addEventListener: () => {}, removeEventListener: () => {} },
    });

  const capture = make();
  const bubble = make();
  const result = { capture: capture.handleCapture(build()), bubble: bubble.handle(build()) };
  capture.dispose();
  bubble.dispose();
  return result;
}

describe('עיצוב תו', () => {
  it('Ctrl+Shift+X — קו חוצה', () => {
    expect(press({ code: 'KeyX', ctrlKey: true, shiftKey: true }).commands).toEqual([
      { id: 'strikethrough', payload: undefined },
    ]);
  });

  it('Ctrl+Space — ניקוי עיצוב', () => {
    expect(press({ code: 'Space', ctrlKey: true }).commands).toEqual([
      { id: 'clear-formatting', payload: undefined },
    ]);
  });

  it('Ctrl+Shift+C — מברשת עיצוב', () => {
    expect(press({ code: 'KeyC', ctrlKey: true, shiftKey: true }).commands).toEqual([
      { id: 'copy-format', payload: undefined },
    ]);
  });

  it('Ctrl+Shift+C אינו Ctrl+C — ההעתקה נשארת של הדפדפן', () => {
    const copy = press({ code: 'KeyC', ctrlKey: true });
    expect(copy.commands).toEqual([]);
    expect(copy.actions).toEqual([]);
  });

  it('Ctrl+] ו-Ctrl+[ הם פעולות, כי הגודל תלוי במנוע', () => {
    expect(press({ key: ']', code: 'BracketRight', ctrlKey: true }).actions).toEqual(['font-grow']);
    expect(press({ key: '[', code: 'BracketLeft', ctrlKey: true }).actions).toEqual(['font-shrink']);
  });

  it('הסוגריים לפי המקש הפיזי — בפריסה עברית התווים מתהפכים', () => {
    // מדוד: בפריסה העברית `BracketLeft` מפיק „]” ו-`BracketRight` מפיק „[”.
    // התאמה לפי תו הייתה מחליפה בין „הגדל” ל„הקטן” בדיוק בפריסה שהתוסף נועד
    // לה. המקש הפיזי נשאר במקומו, כמו בכל שאר הקיצורים.
    expect(press({ key: '[', code: 'BracketRight', ctrlKey: true }).actions).toEqual(['font-grow']);
    expect(press({ key: ']', code: 'BracketLeft', ctrlKey: true }).actions).toEqual(['font-shrink']);
  });

  it('Ctrl+= ו-Ctrl+Shift+= — תחתי ועילי', () => {
    expect(press({ key: '=', code: 'Equal', ctrlKey: true }).actions).toEqual(['subscript']);
    // עם Shift הדפדפן מדווח `key: '+'` בכל הפריסות שנמדדו. רשומה שנשענה על
    // התו לא הייתה יורה לעולם — וזה בדיוק מה שקרה כאן עד שנמדד.
    expect(press({ key: '+', code: 'Equal', ctrlKey: true, shiftKey: true }).actions).toEqual([
      'superscript',
    ]);
  });
});

describe('פסקה', () => {
  const alignments = [
    ['KeyR', 'right'],
    ['KeyE', 'center'],
    ['KeyL', 'left'],
    ['KeyJ', 'justify'],
  ] as const;

  for (const [code, alignment] of alignments) {
    it(`Ctrl+${code.slice(3)} — יישור ${alignment}`, () => {
      expect(press({ code, ctrlKey: true }).commands).toEqual([
        { id: 'text-align', payload: { alignment } },
      ]);
    });
  }

  it('Ctrl+M ו-Ctrl+Shift+M — כניסה', () => {
    expect(press({ code: 'KeyM', ctrlKey: true }).commands).toEqual([
      { id: 'indent-increase', payload: undefined },
    ]);
    expect(press({ code: 'KeyM', ctrlKey: true, shiftKey: true }).commands).toEqual([
      { id: 'indent-decrease', payload: undefined },
    ]);
  });

  it('ריווח שורות שולח מספר, לא מחרוזת', () => {
    // `unwrapScalar` של המנוע מכיר את המפתח `lineHeight` עם מספר. מחרוזת
    // נדחית בשקט — וזה בדיוק סוג הכשל שהכפתורים סבלו ממנו קודם.
    expect(press({ code: 'Digit1', ctrlKey: true }).commands).toEqual([
      { id: 'line-height', payload: { lineHeight: 1 } },
    ]);
    expect(press({ code: 'Digit2', ctrlKey: true }).commands).toEqual([
      { id: 'line-height', payload: { lineHeight: 2 } },
    ]);
    expect(press({ code: 'Digit5', ctrlKey: true }).commands).toEqual([
      { id: 'line-height', payload: { lineHeight: 1.5 } },
    ]);
  });

  it('Ctrl+Shift+N — סגנון רגיל', () => {
    expect(press({ code: 'KeyN', ctrlKey: true, shiftKey: true }).commands).toEqual([
      { id: 'linked-style', payload: { style: 'Normal' } },
    ]);
  });

  it('Ctrl+1 אינו Ctrl+Alt+1 — ריווח מול כותרת', () => {
    expect(press({ code: 'Digit1', ctrlKey: true }).commands[0]?.id).toBe('line-height');
    expect(press({ code: 'Digit1', ctrlKey: true, altKey: true }).commands[0]?.id).toBe(
      'linked-style',
    );
  });

  it('Ctrl+Shift+N אינו Ctrl+N — סגנון מול מסמך חדש', () => {
    expect(press({ code: 'KeyN', ctrlKey: true }).actions).toEqual(['new-document']);
    expect(press({ code: 'KeyN', ctrlKey: true, shiftKey: true }).commands[0]?.id).toBe(
      'linked-style',
    );
  });
});

describe('הרשימה כולה', () => {
  it('לכל רשומה שאינה של הדפדפן יש ניתוב בפועל', () => {
    // שער נגד רשומה שנוספה ונשכחה: היא תיראה בדיאלוג העזרה ולא תעשה דבר.
    const dead: string[] = [];

    for (const shortcut of ENTRIES) {
      if (shortcut.native) continue;
      // כיווניות מזוהה בשחרור מודיפייר — נבדקת ב-shortcut-direction.test.ts.
      if (shortcut.onKeyUp) continue;
      const codes =
        shortcut.code === undefined
          ? ['']
          : typeof shortcut.code === 'string'
            ? [shortcut.code]
            : [...shortcut.code];

      for (const code of codes) {
        const ran = press({
          code,
          key: shortcut.key ?? '',
          ctrlKey: shortcut.ctrl === true,
          shiftKey: shortcut.shift === true,
          altKey: shortcut.alt === true,
        });
        if (ran.commands.length + ran.actions.length !== 1) dead.push(`${shortcut.id}:${code}`);
      }
    }

    expect(dead).toEqual([]);
  });

  /**
   * השלב שכל רשומה רצה בו — על הרג'יסטרי האמיתי ולא על רשימה מומצאת.
   *
   * הבדיקה הזאת נוספה בסבב QA, אחרי שנמדד שהיא הייתה חסרה: החזרת
   * `runsBeforeEngine` ל-`false` — כלומר בדיוק הבאג שבגללו עשר הרשומות האלה
   * לא עבדו כשהסמן במסמך — השאירה את **כל** קובץ הניתוב ירוק, ואת שער
   * „לכל רשומה יש ניתוב בפועל” שמעל ירוק במיוחד: הוא מריץ את שני השלבים
   * ברצף, ולכן רשומה שנדדה מ-capture ל-bubble עדיין „רצה”.
   *
   * ולכן המצפן כאן הוא **רשימת המזהים**, ולא הפונקציה: בדיקה שקוראת ל-
   * `runsBeforeEngine` כדי לדעת מה לצפות הייתה מסכימה עם כל מוטציה שלה.
   * הרשימה היא מה שנמדד בדפדפן (scripts/qa/custom-shortcut-focus-qa.mjs):
   * המאזין של המנוע מטפל ב-`Alt` בלי `Ctrl` כהקלדת תו. עשר מהן נבלעו בפועל,
   * ו-`macro-manage` (מקש פונקציה, בלי תו להקליד) מצטרפת לשלב מכוח צורת
   * הצירוף — ראו `runsBeforeEngine`. רשומה חדשה שתיפול לכאן תצבע את הבדיקה
   * באדום, וזה הרצוי: „האם גם את זאת המנוע בולע” היא שאלה שעונים עליה
   * במדידה, פעם אחת.
   */
  it('רשומות ה-`Alt` רצות ב-capture, וכל השאר ב-bubble', () => {
    const BEFORE_ENGINE = [
      'tab-goto-1',
      'tab-goto-2',
      'tab-goto-3',
      'tab-goto-4',
      'tab-goto-5',
      'tab-goto-6',
      'tab-goto-7',
      'tab-goto-8',
      'tab-goto-last',
      'tell-me',
      'macro-manage',
      // `Ctrl+Shift` שהמנוע קושר לפעולה אחרת (מירכוז, רשימת תבליטים) — נמדד
      // ב-scripts/qa/ctrl-shift-survey-probe.mjs.
      'track-changes',
      'formatting-marks',
    ];

    const inCapture: string[] = [];
    const inBoth: string[] = [];

    for (const shortcut of ENTRIES) {
      if (shortcut.native) continue;
      if (shortcut.onKeyUp) continue;
      // קוד אחד די: השלב נקבע מהמודיפיירים, ורשומה עם כמה קודים (Enter
      // ו-NumpadEnter) אינה יכולה להתפצל ביניהם.
      const code = typeof shortcut.code === 'string' ? shortcut.code : (shortcut.code?.[0] ?? '');

      const { capture, bubble } = phasesOf({
        code,
        key: shortcut.key ?? '',
        ctrlKey: shortcut.ctrl === true,
        shiftKey: shortcut.shift === true,
        altKey: shortcut.alt === true,
      });

      if (capture && bubble) inBoth.push(shortcut.id);
      if (capture) inCapture.push(shortcut.id);
    }

    expect([...inCapture].sort()).toEqual([...BEFORE_ENGINE].sort());
    // הרצה בשני השלבים פירושה הדגשה שנדלקת ונכבית באותה הקשה.
    expect(inBoth).toEqual([]);
  });
});

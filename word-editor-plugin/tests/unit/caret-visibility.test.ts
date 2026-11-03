/**
 * החשבון שמביא את הסמן לתצוגה בחלון צר.
 *
 * המספרים נמדדו בכרום על ה-dist הארוז, מסמך עברי בחלון 600px: מיכל הגלילה
 * מדווח `scrollWidth` 794 מול `clientWidth` 600 (כלומר 209px גלילה זמינה),
 * הסמן בתחילת השורה יושב על `x = 698`, ואחרי שלושה תווים על `672` — שניהם
 * מעבר לקצה התצוגה. ההנמקה המלאה בראש engine/caret-visibility.ts.
 *
 * „הדף באמת זז” אינו נמדד ב-jsdom — אין בו פריסה, `getBoundingClientRect`
 * מחזיר אפסים ו-`scrollWidth` הוא 0 תמיד, כלומר כל נוסחה שהיא עוברת שם ירוק.
 * את זה מודד scripts/qa/caret-visibility-qa.mjs. מה שכן נמדד כאן הוא הנוסחה
 * ו**ההתקנה** — מי קורא לה, עם מה, ומה הוא עושה בתשובה — מול כפילים שמחזיקים
 * את המספרים שנמדדו בכרום.
 */
import { describe, expect, it } from 'vitest';
import {
  CARET_MARGIN_PX,
  horizontalScrollDelta,
  installCaretVisibility,
  isCaretFollowScroll,
  type CaretVisibilityHost,
  type CaretVisibilityUi,
} from '../../src/engine/caret-visibility';

/** התצוגה כפי שנמדדה: מאגס 600px שמתחיל בקצה החלון. */
const VIEW = { left: 0, right: 600 };

/** סמן מכווץ — `width: 0`, כפי שהמנוע מדווח אותו. */
const caretAt = (x: number) => ({ left: x, right: x });

describe('horizontalScrollDelta', () => {
  it('סמן בתחילת שורה עברית, מעבר לקצה — גולל עד שהוא נכנס עם המרווח', () => {
    // 698 הוא המקום שנמדד; הגבול הוא 600 − 48.
    expect(horizontalScrollDelta(caretAt(698), VIEW, CARET_MARGIN_PX)).toBe(698 - (600 - 48));
  });

  it('גם אחרי שלושה תווים הסמן עדיין בחוץ, ועדיין נגלל', () => {
    expect(horizontalScrollDelta(caretAt(672), VIEW, CARET_MARGIN_PX)).toBe(672 - (600 - 48));
  });

  it('סמן שכבר בתצוגה, הרחק משני הקצוות — אין לזוז', () => {
    expect(horizontalScrollDelta(caretAt(300), VIEW, CARET_MARGIN_PX)).toBe(0);
  });

  /*
   * הקצה השני הוא סוף השורה בעברית, ושם הדלתא **שלילית**: המיכל `ltr`, ולכן
   * הקטנת `scrollLeft` היא החשיפה שמאלה. בלי הסימן הזה ההקלדה לקצה השמאלי
   * הייתה גוררת את הדף לכיוון ההפוך בדיוק.
   */
  it('סמן שחרג בקצה השמאלי נגלל אחורה', () => {
    expect(horizontalScrollDelta(caretAt(10), VIEW, CARET_MARGIN_PX)).toBe(10 - 48);
  });

  it('בדיוק על גבול המרווח — עוד אין תנועה', () => {
    expect(horizontalScrollDelta(caretAt(600 - CARET_MARGIN_PX), VIEW, CARET_MARGIN_PX)).toBe(0);
    expect(horizontalScrollDelta(caretAt(CARET_MARGIN_PX), VIEW, CARET_MARGIN_PX)).toBe(0);
  });

  /*
   * החיתוך לרבע מרוחב התצוגה. בלי אחד כזה, מאגס צר מ-‎2·48 היה מקבל `min`
   * גדול מ-`max`: כל מיקום סמן מפר את שני הגבולות, וכל תיקון של האחד מפר את
   * השני — כלומר גלילה שמתנדנדת ואינה נחה. עם החיתוך נשאר תמיד חצי מרוחב
   * התצוגה שבו הסמן „בסדר”, ולכן כל גלילה מגיעה למנוחה בצעד אחד.
   */
  it('מאגס צר: המרווח נחתך, והגבולות אינם מתהפכים', () => {
    const narrow = { left: 0, right: 80 };
    // pad נחתך ל-20, כלומר הטווח הבטוח הוא 20..60.
    expect(horizontalScrollDelta(caretAt(100), narrow, CARET_MARGIN_PX)).toBe(40);
    expect(horizontalScrollDelta(caretAt(40), narrow, CARET_MARGIN_PX)).toBe(0);
    expect(horizontalScrollDelta(caretAt(0), narrow, CARET_MARGIN_PX)).toBe(-20);
  });

  it('גלילה מביאה למנוחה: הפעלה שנייה על התוצאה מחזירה אפס', () => {
    const delta = horizontalScrollDelta(caretAt(698), VIEW, CARET_MARGIN_PX);
    // אחרי שהמיכל זז ב-delta, הסמן זז באותו שיעור אל תוך התצוגה.
    expect(horizontalScrollDelta(caretAt(698 - delta), VIEW, CARET_MARGIN_PX)).toBe(0);
  });

  /* מדידה פגומה אינה מזיזה את הדף: עדיף לא לגלול על פני לגלול לשום מקום. */
  it('מידות שאינן מספרים אינן גוררות תנועה', () => {
    expect(horizontalScrollDelta(caretAt(Number.NaN), VIEW, CARET_MARGIN_PX)).toBe(0);
    expect(horizontalScrollDelta(caretAt(698), { left: 0, right: 0 }, CARET_MARGIN_PX)).toBe(0);
    expect(horizontalScrollDelta(caretAt(698), { left: 600, right: 0 }, CARET_MARGIN_PX)).toBe(0);
  });
});

/*
 * -------------------------------------------------------------------------
 * וההתקנה עצמה.
 *
 * הנוסחה שלמעלה היא חצי מהמודול; החצי השני הוא מי קורא לה, עם מה, ומה הוא
 * עושה בתשובה. סבב מוטציה על הגוף של `installCaretVisibility` — החלפתו
 * ב-`return { follow() {}, dispose() {} }` — השאיר את שמונה הבדיקות שמעל
 * ירוקות, כלומר לא היה שם דבר שמחזיק אותו.
 *
 * הכפיל הוא `{ scrollWidth, clientWidth, clientLeft, scrollLeft,
 * getBoundingClientRect }` ולא אלמנט אמיתי, מאותה סיבה שכתובה בראש הקובץ.
 * המספרים הם אלה שנמדדו בכרום: מאגס `clientWidth` 585 מול `scrollWidth` 794
 * (209px גלילה), והסמן בתחילת שורה עברית על `x = 698`.
 * -------------------------------------------------------------------------
 */

/**
 * המיכל כפי שנמדד בחלון 600px.
 *
 * ‏`scrollLeft` **מעגל** את מה שנכתב לו, כמו אלמנט אמיתי: נמדד ב-Chrome 152
 * ‏‎161.47→161, 176.46875→176, 176.6→177, 0.5→1, 224→224. זה אינו פרט טכני —
 * גרסה קודמת של `isCaretFollowScroll` השוותה את `scrollLeft` למספר השבור
 * שנכתב, וכפיל שלא מעגל היה מדווח עליה ירוק בעוד שבדפדפן היא כמעט לא נתפסה.
 *
 * ‏`boxReads` ו-`widthReads` סופרים את שתי קריאות הפריסה — זה מה שמבדיל „יצא
 * מוקדם” מ„חישב הכול והגיע לאפס”, וזה מה שמקבע את **סדר** הקריאות.
 */
function fakeHost(
  over: Partial<CaretVisibilityHost> = {},
): CaretVisibilityHost & { boxReads: number; widthReads: number } {
  const base = { scrollWidth: 794, clientWidth: 585, clientLeft: 0, scrollLeft: 0, ...over };
  let scrollLeft = base.scrollLeft;
  const host = {
    clientWidth: base.clientWidth,
    clientLeft: base.clientLeft,
    boxReads: 0,
    widthReads: 0,
    get scrollLeft() {
      return scrollLeft;
    },
    set scrollLeft(value: number) {
      scrollLeft = Math.round(value);
    },
    get scrollWidth() {
      host.widthReads += 1;
      return base.scrollWidth;
    },
    getBoundingClientRect:
      over.getBoundingClientRect ??
      (() => {
        host.boxReads += 1;
        return { left: 0 };
      }),
  };
  return host;
}

/** אירוע `scroll` כפי ש-`isCaretFollowScroll` צורך אותו: ה-`target` בלבד. */
const scrollOf = (target: unknown) => ({ target });

/** ממתינה לפריים הבא, שבו הסימן של המודול נמחק. */
const nextFrame = () => new Promise((resolve) => requestAnimationFrame(() => resolve(null)));

/**
 * סמן מכווץ בקואורדינטות חלון, בצורה שנמדדה מהמנוע: `left` ו-`right` קיימים
 * שניהם (המלבן המלא הוא `{pageIndex, left, right, top, bottom, width, height}`).
 */
const rectAt = (x: number) => ({ left: x, right: x });

type Placement = { placement?: 'start' | 'end' | 'center' };

interface FakeUi extends CaretVisibilityUi {
  /** משחררת את הבחירה, כלומר מפעילה את מה ש-`subscribe` רשם. */
  fire(): void;
  /** הארגומנטים שכל קריאה ל-`getAnchorRect` קיבלה. */
  readonly placements: unknown[];
  /** כמה ביטולים נקראו בפועל. */
  unsubscribed(): number;
}

function fakeUi(
  anchor: (input?: Placement) => unknown,
  { subscribeReturns = 'function' as 'function' | 'object' }: { subscribeReturns?: 'function' | 'object' } = {},
): FakeUi {
  const placements: unknown[] = [];
  let listener: (() => void) | null = null;
  let unsubscribed = 0;

  return {
    selection: {
      subscribe: (next: () => void) => {
        listener = next;
        if (subscribeReturns === 'function') {
          // מנתק באמת, כמו המנוע: בלי זה „dispose מנתק” עובר ירוק גם על ידית
          // שאינה קוראת לביטול כלל.
          return () => {
            unsubscribed += 1;
            listener = null;
          };
        }
        return {
          dispose: () => {
            unsubscribed += 1;
          },
        };
      },
      getAnchorRect: (input?: Placement) => {
        placements.push(input);
        return anchor(input) as ReturnType<
          NonNullable<NonNullable<CaretVisibilityUi['selection']>['getAnchorRect']>
        >;
      },
    },
    fire: () => listener?.(),
    placements,
    unsubscribed: () => unsubscribed,
  };
}

describe('installCaretVisibility', () => {
  it('נרשם לבחירה, ותזוזה שלה מגלגלת את המיכל', () => {
    const host = fakeHost();
    const ui = fakeUi(() => rectAt(698));
    const handle = installCaretVisibility({ host, ui });

    // עד שהבחירה לא זזה, איש לא נגע בגלילה.
    expect(host.scrollLeft).toBe(0);

    ui.fire();

    // 698 מול תצוגה שנגמרת ב-585, פחות מרווח של 48.
    expect(host.scrollLeft).toBe(698 - (585 - CARET_MARGIN_PX));
    handle.dispose();
  });

  /*
   * מה שנשלח, ורק הוא. נמדד ש-2.15.0 **מתעלם** מ-`placement` ומחזיר את המלבן
   * שעוטף את הבחירה כולה, ולכן הכפיל כאן מתעלם ממנו גם הוא — כפיל שמכבד
   * ארגומנט שהמנוע זורק היה נותן לנו לקבע מספר שאינו קיים בדפדפן.
   *
   * מה שכן מקובע הוא הצד שלנו בחוזה: אותה חתימה כמו ארבעת אתרי הקריאה האחרים
   * בריפו. `'end'` הוא הקצה **האחרון בסדר המסמך**, ולא „ראש הבחירה”.
   */
  it('שולח את אותה חתימה כמו שאר אתרי הקריאה', () => {
    const host = fakeHost();
    const ui = fakeUi(() => rectAt(698));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    expect(ui.placements).toEqual([{ placement: 'end' }]);
    handle.dispose();
  });

  /*
   * מלבן של **טווח**, ולא של סמן מכווץ. נמדד שזה בדיוק מה שהמנוע מחזיר כשיש
   * בחירה: על 12 תווים המלבן היה 453.5..539.1 (וסמן מכווץ נמדד כ-`width: 0`,
   * `left === right`). הקצה הימני הוא שמכריע — בעברית שם מתחילה השורה — ומי
   * שיקרא את `left` במקומו יראה מלבן שכולו „בתצוגה” ולא יזיז דבר.
   *
   * החשש שהקצה הימני הוא דווקא **העוגן** בבחירה שנפתחת שמאלה נבדק בדפדפן,
   * ולא התממש: בחלון 600, אחרי `Home` בשורה עברית ארוכה (הגלילה כבר עקבה,
   * `scrollLeft` 161/209), עשר הרחבות `Shift+שמאלה` הזיזו את הקצה השמאלי
   * מ-537 ל-462 והשאירו את `scrollLeft` על 161 — לא נורה ולו אירוע גלילה
   * אחד. אותו דבר בכיוון ההפוך מ-`End` (‏`scrollLeft` 109, המלבן 48..136).
   * הסיבה נראית מהמספרים: אחרי מעקב הסמן הקצה הקבוע יושב **על** גבול המרווח
   * ולא מעבר לו, ולכן `caret.right > max` אינו מתקיים וההכרעה נופלת לקצה שזז.
   */
  it('מלבן של טווח — הקצה הימני הוא שמכריע', () => {
    const host = fakeHost();
    const ui = fakeUi(() => ({ left: 453.5, right: 539.1 }));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    // הדפדפן מעגל את מה שנכתב, וכך גם הכפיל: 539.1 − 537 = 2.1 → 2.
    expect(host.scrollLeft).toBe(Math.round(539.1 - (585 - CARET_MARGIN_PX)));
    handle.dispose();
  });

  /*
   * „אין גלישה” אינו נמדד דרך `scrollLeft` בלבד: הוא נשאר 0 גם בלי היציאה,
   * מפני שהחיתוך לטווח `[0, 0]` מחזיר אותו לשם ממילא. מה שמבדיל הוא **שלא
   * נמדד כלום** — התצוגה לא נקראה, והמיכל לא סומן כמי שהמודול הזיז.
   */
  it('אין גלישה — יציאה מוקדמת, בלי מדידה ובלי סימון', () => {
    const host = fakeHost({ scrollWidth: 585, clientWidth: 585 });
    const ui = fakeUi(() => rectAt(698));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    expect(host.scrollLeft).toBe(0);
    expect(host.boxReads).toBe(0);
    expect(isCaretFollowScroll(scrollOf(host))).toBe(false);
    handle.dispose();
  });

  /*
   * נמדד: באירוע הראשון של כל הקשה המנוע עדיין מחזיר `null`.
   *
   * המיכל מתחיל **גלול** ולא באפס, ובכוונה: מי שיתרגם „אין מלבן” ל-`{0, 0}`
   * יגרור את הדף בחזרה לקצה השמאלי בכל הקשה שנייה, ומיכל שנח על אפס אינו
   * מבחין בזה כלל.
   */
  it('אין מלבן — יציאה שקטה, בלי לגעת בגלילה', () => {
    const host = fakeHost({ scrollLeft: 120 });
    const ui = fakeUi(() => null);
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    expect(host.scrollLeft).toBe(120);
    handle.dispose();
  });

  /*
   * וזה מקבע את **הסדר**: המנוע נשאל לפני שנמדדת הגלישה.
   *
   * הסדר נבחר לפי מדידה ולא לפי תחושה — ‏~200 אירועי בחירה אמיתיים בכל רוחב
   * חלון, והמספרים במודול: מי שקורא ראשון משלם על הפריסה שההקשה השאירה
   * פתוחה, ו-`getAnchorRect` משלם פחות (0.068ms מול 0.169ms בחלון רחב,
   * 0.089ms מול 0.281ms בחלון צר). מחצית מהאירועים מחזירים `null`, והם
   * יוצאים כאן **בלי אף קריאת פריסה**. היפוך הסדר מאדים את שתי השורות.
   */
  it('אין מלבן — ואף קריאת פריסה לא נעשתה', () => {
    const host = fakeHost({ scrollLeft: 120 });
    const ui = fakeUi(() => null);
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    expect(ui.placements.length).toBe(1);
    expect(host.widthReads).toBe(0);
    expect(host.boxReads).toBe(0);
    handle.dispose();
  });

  it('מלבן שאינו מספרים אינו מזיז את הדף', () => {
    const host = fakeHost({ scrollLeft: 120 });
    const ui = fakeUi(() => ({ left: '698', right: '698' }));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    expect(host.scrollLeft).toBe(120);
    handle.dispose();
  });

  /*
   * המיכל אינו מתחיל בקצה החלון, ויש לו גבול. `getAnchorRect` מדווח
   * בקואורדינטות **חלון**, ולכן ההשוואה חייבת להיות מול תיבת התוכן שלו באותו
   * ציר: `box.left + clientLeft`. בלי `clientLeft` התצוגה נמדדת 5px שמאלה
   * ממקומה — בדיוק הרצועה שהגבול תופס.
   */
  it('התצוגה נמדדת מתיבת התוכן: `box.left + clientLeft`', () => {
    const host = fakeHost({ clientLeft: 5, getBoundingClientRect: () => ({ left: 100 }) });
    const ui = fakeUi(() => rectAt(698));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    // תיבת התוכן היא 105..690, ולכן הגבול הוא 690 − 48.
    expect(host.scrollLeft).toBe(698 - (690 - CARET_MARGIN_PX));
    handle.dispose();
  });

  /*
   * החיתוך ל-`[0, overflow]`. הדפדפן חותך ממילא — ולכן זו אינה התנהגות שנראית
   * למשתמש — אבל בלעדיו המודול מוסר למיכל מספר שאינו במרחב שלו, וכל מי שיקרא
   * את הקוד יצטרך לדעת בעל פה שמישהו אחר מתקן אותו.
   */
  it('הגלילה נחתכת לטווח של המיכל', () => {
    const far = fakeHost();
    const right = fakeUi(() => rectAt(1_400));
    const handle = installCaretVisibility({ host: far, ui: right });
    right.fire();
    expect(far.scrollLeft).toBe(794 - 585);
    handle.dispose();

    const back = fakeHost({ scrollLeft: 30 });
    const left = fakeUi(() => rectAt(-900));
    const handle2 = installCaretVisibility({ host: back, ui: left });
    left.fire();
    expect(back.scrollLeft).toBe(0);
    handle2.dispose();
  });

  it('dispose מנתק — תזוזת בחירה אחריו אינה מגלגלת', () => {
    const host = fakeHost();
    const ui = fakeUi(() => rectAt(698));
    const handle = installCaretVisibility({ host, ui });

    handle.dispose();
    expect(ui.unsubscribed()).toBe(1);

    ui.fire();
    expect(host.scrollLeft).toBe(0);
  });

  /*
   * נמדד ש-`subscribe` מחזיר פונקציה. ערך אחר פירושו גרסת מנוע שאיננו
   * יודעים לפרק ממנה — והידית מוותרת על הביטול במקום לנחש ולזרוק.
   */
  it('subscribe שאינו מחזיר פונקציה — dispose אינו זורק', () => {
    const host = fakeHost();
    const ui = fakeUi(() => rectAt(698), { subscribeReturns: 'object' });
    const handle = installCaretVisibility({ host, ui });

    expect(() => handle.dispose()).not.toThrow();
    expect(ui.unsubscribed()).toBe(0);
  });

  it('בלי מיכל ובלי ידית בחירה — ההתקנה שקטה', () => {
    const noHost = installCaretVisibility({ host: null, ui: fakeUi(() => rectAt(698)) });
    expect(() => noHost.follow()).not.toThrow();
    noHost.dispose();

    const host = fakeHost();
    const empty = installCaretVisibility({ host, ui: { selection: null } });
    expect(() => empty.follow()).not.toThrow();
    expect(host.scrollLeft).toBe(0);
    empty.dispose();

    const none = installCaretVisibility({ host, ui: undefined });
    expect(() => none.follow()).not.toThrow();
    none.dispose();
  });

  /* הידית מוחזרת כדי שהשער יוכל למדוד קריאה יזומה, ולא רק אירוע. */
  it('follow הגלויה עושה את אותו דבר בדיוק', () => {
    const host = fakeHost();
    const ui = fakeUi(() => rectAt(698));
    const handle = installCaretVisibility({ host, ui });

    handle.follow();

    expect(host.scrollLeft).toBe(698 - (585 - CARET_MARGIN_PX));
    handle.dispose();
  });
});

/*
 * ההבחנה בין „המשתמש גלל משם” ל„הגלילה היא המשך של ההקלדה”. בלעדיה רשימת
 * המקורות של „@” נסגרת באמצע הקלדת השאילתה בחלון צר — ההסבר המלא במודול.
 */
describe('isCaretFollowScroll', () => {
  it('מיכל שלא נגעו בו אינו מסומן', () => {
    expect(isCaretFollowScroll(scrollOf(fakeHost()))).toBe(false);
  });

  /*
   * ‏**המלבן כאן שבור בכוונה.** הגרסה הקודמת של ההבחנה השוותה את `scrollLeft`
   * למספר שנכתב, והדפדפן מעגל אותו: `caret.right` הוא שבר בפועל (נמדד
   * `left: 777.47`), ולכן היא נתפסה רק כשהכתיבה נחתכה ל-0 או ל-`overflow`.
   * ‏698.47 − 537 = 161.47, והמיכל נח על 161 — כלומר השורה הזאת אדומה על כל
   * הבחנה שמשווה מיקומים, וירוקה רק על אחת שאינה משווה.
   */
  it('גלילת מעקב עם יעד שבור — האירוע שלה מזוהה', () => {
    const host = fakeHost();
    const ui = fakeUi(() => rectAt(698.47));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    expect(host.scrollLeft).toBe(161);
    expect(isCaretFollowScroll(scrollOf(host))).toBe(true);
    // שני מאזינים קוראים את אותו אירוע; הקריאה הראשונה אינה צורכת את הסימן.
    expect(isCaretFollowScroll(scrollOf(host))).toBe(true);
    handle.dispose();
  });

  /*
   * והסימן **פג**, בפריים שאחרי הכתיבה. נמדד בדפדפן שאירוע ה-`scroll` שהכתיבה
   * יוצרת מגיע לפני ה-`requestAnimationFrame` שנקבע מיד אחריה (`scroll → raf1
   * → raf2`, גם כשהכתיבה עצמה מתוך rAF). בלי הפקיעה, מיכל שנח על מקומו נשאר
   * מסומן לתמיד — וכל גלילה אנכית בגלגל, כל גלילה של צאצא וכל `resize` נבלעים.
   */
  it('הסימן חי פריים אחד בלבד', async () => {
    const host = fakeHost();
    const ui = fakeUi(() => rectAt(698.47));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();
    expect(isCaretFollowScroll(scrollOf(host))).toBe(true);

    await nextFrame();

    expect(isCaretFollowScroll(scrollOf(host))).toBe(false);
    handle.dispose();
  });

  /*
   * ולא רק הזמן: גם ה-`target`. נמדד שגלילה של צאצא נשמעת על שלב ה-capture של
   * המיכל עם ה-`target` של הצאצא, ושה-`resize` — שאותו מאזין רשום לו — מגיע
   * עם ה-`window`. שניהם „המשתמש עשה משהו”, ושניהם נופלים כאן גם באמצע הפריים
   * שבו המודול גלל.
   */
  it('אותו פריים, אבל אירוע של מישהו אחר — לא מסומן', () => {
    const host = fakeHost();
    const child = { name: 'צאצא של המיכל' };
    const view = { name: 'window' };
    const ui = fakeUi(() => rectAt(698.47));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    expect(isCaretFollowScroll(scrollOf(host))).toBe(true);
    expect(isCaretFollowScroll(scrollOf(child))).toBe(false);
    expect(isCaretFollowScroll(scrollOf(view))).toBe(false);
    expect(isCaretFollowScroll(scrollOf(null))).toBe(false);
    handle.dispose();
  });

  it('אירוע שלא הזיז דבר אינו מסמן', () => {
    const host = fakeHost();
    const ui = fakeUi(() => rectAt(300));
    const handle = installCaretVisibility({ host, ui });

    ui.fire();

    expect(host.scrollLeft).toBe(0);
    expect(isCaretFollowScroll(scrollOf(host))).toBe(false);
    handle.dispose();
  });

  /* הסימן הוא פר-מיכל: לטאב אחר יש מיכל משלו, והוא לא נגלל. */
  it('הסימן אינו זולג בין מיכלים', () => {
    const mine = fakeHost();
    const other = fakeHost();
    const ui = fakeUi(() => rectAt(698));
    const handle = installCaretVisibility({ host: mine, ui });

    ui.fire();

    expect(isCaretFollowScroll(scrollOf(mine))).toBe(true);
    expect(isCaretFollowScroll(scrollOf(other))).toBe(false);
    handle.dispose();
  });
});

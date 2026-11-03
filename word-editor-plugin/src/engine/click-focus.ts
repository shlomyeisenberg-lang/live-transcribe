/**
 * לחיצה באזור המסמך שאינה נוחתת על עמוד — והמיקוד שנשאר במסמך.
 *
 * ## הפער שנמדד
 *
 * סביב העמוד יש שטח אפור, והוא גדול: כל רוחב אזור העריכה שמשני צדי העמוד,
 * מעליו ומתחתיו, ובחלון צר גם הרצועה שנשארת אחרי הזום. נמדד ב-Chrome אמיתי
 * (מסמך של 60 פסקאות, superdoc 2.12):
 *
 *   | הלחיצה                       | `activeElement` אחריה  | ההקלדה שאחריה |
 *   |------------------------------|------------------------|---------------|
 *   | על טקסט                       | ה-`textarea` של המנוע  | נכנסת         |
 *   | בשולי העמוד, בגובה הטקסט      | ה-`textarea` של המנוע  | נכנסת         |
 *   | מתחת לשורה האחרונה, בעמוד     | ה-`textarea` של המנוע  | נכנסת         |
 *   | **בשטח האפור שסביב העמוד**    | **`<body>`**           | **נבלעת**     |
 *
 * והבחירה של המנוע **שורדת** את הלחיצה הזאת: אותו `blockId` ואותו טווח לפני
 * ואחרי, והסמן ממשיך להיות מצויר. כלומר המשתמש רואה סמן מהבהב בתוך המסמך,
 * מקליד — ושום דבר לא קורה, עד שהוא לוחץ שוב, הפעם על העמוד. זה מה שדווח:
 * „לפעמים במעבר מהמסמך לכפתורים, כשחוזרים למסמך צריך ללחוץ כמה פעמים בשביל
 * שיקליד”. הלחיצה שאחרי הרצועה נופלת באפור, וכפתורי הרצועה עצמם אינם
 * מעורבים — הם מבטלים את `pointerdown` ולכן המיקוד כלל אינו יוצא מהמסמך
 * (ui/ribbon/common/RibbonButton.vue).
 *
 * ב-Word אין לזה מקבילה: השטח שסביב העמוד אינו לוקח מיקוד, הסמן נשאר במקומו
 * וההקלדה ממשיכה להיכנס.
 *
 * ## שני מצבים, ושתי תשובות שונות
 *
 *   - **המיקוד במסמך** → `preventDefault` על ה-`pointerdown`. ביטול
 *     ה-`pointerdown` מדכא גם את ה-`mousedown` התואם (Pointer Events §11),
 *     והוא זה שמזיז מיקוד — כלומר שום דבר אינו מאבד אותו, והסמן נשאר בדיוק
 *     במקום שהיה. זה עדיף על החזרת מיקוד בדיעבד: אין `restoreSelection`, אין
 *     גלילה, ואין הבהוב.
 *   - **המיקוד מחוץ למסמך** (תיבת הרצועה, לשונית שנלחצה, שכבה שנסגרה) →
 *     נותנים לדפדפן לעשות את שלו, ומיד אחריו מחזירים את המיקוד למסמך. כאן
 *     `preventDefault` היה מזיק: הוא היה משאיר את רשימת הגופנים פתוחה, כי
 *     היא נסגרת על `blur` של השדה.
 *
 * המצב השני נמדד גם הוא, ולא הוסק: מיקוד בתיבת „גודל גופן” + לחיצה באפור
 * הותירו `activeElement` על `<body>` וההקלדה לא נכנסה; ולחיצה **על** רצועת
 * הכותרת של העמוד (שהיא בתוך העמוד, ולכן אינה „אפור”) לא הזיזה מיקוד בכלל —
 * כלומר מי שהיה על כפתור ברצועה נשאר שם. לכן הבדיקה שאחרי הלחיצה היא „האם
 * המיקוד בתוך המסמך”, ולא „האם הלחיצה הייתה מחוץ לעמוד”.
 *
 * ## מה נמדד שהתיקון **אינו** שובר
 *
 * ה-container שהמנוע מצייר בתוכו הוא גם מיכל הגלילה, ולכן לחיצה על פס
 * הגלילה מגיעה ל-DOM עם הקואורדינטות של הרצועה — כלומר „מחוץ לעמוד”, כלומר
 * מבוטלת. נמדד שזה חסר משמעות לגלילה: לחיצה על מסלול הפס גללה (0→859) וגרירת
 * הידית גללה (859→0) גם כשהשומר ביטל את האירוע, מפני שפס גלילה מקומי אינו
 * „פעולת ברירת המחדל” של האירוע. תופעת לוואי מבורכת: מעכשיו לחיצה על הפס גם
 * אינה מוציאה את המיקוד מהמסמך.
 *
 * כפתור שאינו ראשי אינו נוגע בכלל: `contextmenu` ולחיצה אמצעית ממשיכים
 * כרגיל.
 */
import { focusDocument, type FocusTarget } from './focus';
import { pointOverPage } from './page-ruler';

export interface ClickFocusHandle {
  dispose(): void;
}

/**
 * מה לעשות בלחיצה.
 *
 *   - `keep` — לבטל את האירוע; המיקוד והסמן נשארים.
 *   - `ensure` — לא לבטל, ואחרי שהדפדפן סיים לוודא שהמיקוד חזר למסמך.
 *   - `none` — לא לגעת.
 */
export type ClickFocusAction = 'keep' | 'ensure' | 'none';

export interface ClickFocusInput {
  /** האם הנקודה על עמוד מצויר. `null` = אין עמודים, אין תשובה. */
  readonly onPage: boolean | null;
  /** האם המיקוד כרגע בתוך אזור המסמך. */
  readonly focused: boolean;
}

/**
 * ההכרעה עצמה, בלי DOM.
 *
 * `onPage === null` הוא „עוד לא יודעים”: בין פתיחת המסמך לציור הראשון אין
 * עמודים בעץ, וכל לחיצה הייתה נראית „מחוץ לעמוד”. במצב הזה עדיף לא לגעת —
 * ביטול לחיצה שהמנוע כן היה מטפל בה גרוע יותר מלחיצה שנשארת כפי שהיא.
 */
export function clickFocusAction({ onPage, focused }: ClickFocusInput): ClickFocusAction {
  if (!focused) return 'ensure';
  if (onPage === false) return 'keep';
  return 'none';
}

export interface ClickFocusOptions {
  /**
   * מתי לבדוק שהמיקוד חזר. ברירת המחדל: `setTimeout(0)` — הזזת המיקוד של
   * הדפדפן קורית בתוך אותו task של ה-`pointerdown` (נמדד: כל שרשרת
   * ה-`focusout`/`focusin` באותה מילישנייה), ולכן task אחד אחריו מספיק.
   * הפרמטר קיים כדי שבדיקה תריץ סינכרונית.
   */
  readonly schedule?: (run: () => void) => void;
  /** דריסת מדידת העמודים, לבדיקות. */
  readonly onPageAt?: (x: number, y: number) => boolean | null;
}

/**
 * מתקינה את השומר על ה-container של המנוע. נרשמת ב-capture על ה-`window`,
 * כמו pointer-snap.ts, אבל בלי דרישת סדר: `preventDefault` פועל על פעולת
 * ברירת המחדל ולא על מאזינים אחרים, ולכן אין הבדל אם המנוע נרשם לפנינו.
 */
export function installClickFocus(
  container: HTMLElement,
  host: FocusTarget,
  options: ClickFocusOptions = {},
): ClickFocusHandle {
  const view = container.ownerDocument?.defaultView ?? window;
  const schedule = options.schedule ?? ((run: () => void) => void setTimeout(run, 0));
  const onPageAt = options.onPageAt ?? ((x: number, y: number) => pointOverPage(container, x, y));

  const focused = (): boolean => {
    const active = container.ownerDocument?.activeElement ?? null;
    return active !== null && container.contains(active);
  };

  const onDown = (event: MouseEvent): void => {
    if (event.button !== 0) return;
    const target = event.target;
    if (!(target instanceof Node) || !container.contains(target)) return;

    switch (clickFocusAction({ onPage: onPageAt(event.clientX, event.clientY), focused: focused() })) {
      case 'keep':
        event.preventDefault();
        return;
      case 'ensure':
        schedule(() => {
          if (!focused()) focusDocument(host);
        });
        return;
      default:
        return;
    }
  };

  const listener = onDown as EventListener;
  view.addEventListener('pointerdown', listener, true);

  return {
    dispose() {
      view.removeEventListener('pointerdown', listener, true);
    },
  };
}

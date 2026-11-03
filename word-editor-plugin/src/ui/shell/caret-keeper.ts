/**
 * לחיצה על **רקע** של פס מעטפת אינה לוקחת את הסמן מהמסמך.
 *
 * ## מה נמדד
 *
 * בכל פס של המעטפת יש שטח שאינו פקד — הרווח בין קבוצות ברצועה, השטח שמימין
 * ללשוניות, הקצה של שורת המצב. לחיצה בכל אחד מהם נמדדה ב-Chrome אמיתי, אחרי
 * לחיצה בטקסט (כלומר עם סמן במסמך), וההקלדה שאחריה **בלי לחיצה נוספת**:
 *
 *   | הנקודה הריקה                 | `activeElement` אחריה | ההקלדה |
 *   |------------------------------|------------------------|--------|
 *   | קצה שורת המצב                 | `<body>`               | נבלעת  |
 *   | הפס של לשוניות המסמכים        | `<body>`               | נבלעת  |
 *   | סרגל הכותרת                   | `<body>`               | נבלעת  |
 *   | שורת לשוניות הרצועה           | `<body>`               | נבלעת  |
 *   | **גוף הרצועה — בין הכפתורים** | `<body>`               | נבלעת  |
 *
 * זה אותו כשל שבשטח האפור שסביב העמוד (engine/click-focus.ts), רק מהצד השני:
 * הסמן ממשיך להיות מצויר במסמך, המשתמש מקליד, ושום דבר לא נכנס. הפעם הפגיעה
 * דווקא בתנועה שהמשתמש עושה הכי הרבה — יד אל הרצועה, לחיצה שפספסה כפתור
 * בשני פיקסלים.
 *
 * ב-Word אין לזה מקבילה: הרצועה, פס הכותרת ושורת המצב אינם לוקחים מיקוד.
 * הכפתורים שלנו כבר מתנהגים כך בעצמם (`@pointerdown.prevent` ב-RibbonButton),
 * וזה מה שמשלים את התמונה לרקע שביניהם.
 *
 * ## הכלל, וגבולותיו
 *
 * ‏`pointerdown` בכפתור ראשי שנוחת על המעטפת, **לא** על פקד, **לא** באזור
 * המסמך, וכשהמיקוד כרגע במסמך — מבוטל. ביטול ה-`pointerdown` מדכא גם את
 * ה-`mousedown` התואם (Pointer Events §11), והוא זה שמזיז מיקוד.
 *
 *   - **פקד** — לפי `CARET_KEEPER_CONTROL_SELECTOR`: לו המיקוד שייך, והוא זה
 *     שיחזיר אותו כשיסיים (ui/ribbon/common/RibbonCombo.vue).
 *   - **אזור המסמך** — יוצא מהכלל הזה במפורש: שם לחיצה כן צריכה להזיז סמן,
 *     ומה שקורה בשטח האפור שבתוכו מטופל ב-engine/click-focus.ts.
 *   - **המיקוד כבר לא במסמך** — אין מה לשמור, והביטול היה מזיק: שכבה
 *     שנסגרת על `blur` (רשימת הגופנים) הייתה נשארת פתוחה.
 *
 * הביטול אינו עוצר מאזינים — רק את פעולת ברירת המחדל — ולכן גרירת סמן טאב
 * בסרגל, גרירת דיאלוג וכל טיפול אחר בלחיצה ממשיכים לרוץ כרגיל.
 */

/**
 * מה נחשב פקד שהמיקוד שייך לו.
 *
 * `[tabindex]` נכלל בכוונה, גם `-1`: אלמנט שמישהו סימן במפורש כמקבל מיקוד
 * תכנותי (שורש של דיאלוג, פריט ברשימה) אינו „רקע”.
 *
 * משטח עריכה אינו ברשימה: המשטח היחיד בתוסף הוא זה של המנוע, והוא בתוך אזור
 * המסמך — שכבר יוצא מהכלל. הרשימה מתארת את **פסי המעטפת** ולא את הדף.
 */
export const CARET_KEEPER_CONTROL_SELECTOR =
  'button, input, select, textarea, a, [role="button"], [role="tab"], [role="combobox"], [role="menuitem"], [role="option"], [tabindex]';

export interface CaretKeeperInput {
  /** הלחיצה נחתה על פקד. */
  readonly onControl: boolean;
  /** הלחיצה נחתה באזור המסמך. */
  readonly inDocument: boolean;
  /** המיקוד כרגע בתוך אזור המסמך. */
  readonly documentFocused: boolean;
}

/** האם לבטל את הלחיצה כדי לשמור את הסמן. */
export function shouldKeepCaret({ onControl, inDocument, documentFocused }: CaretKeeperInput): boolean {
  if (!documentFocused) return false;
  if (inDocument) return false;
  return !onControl;
}

export interface CaretKeeperOptions {
  /** אזור המסמך: מה שבתוכו אינו „מעטפת”. `null` = אין מסמך פתוח. */
  readonly documentArea: () => HTMLElement | null;
}

/**
 * המאזין עצמו. נמסר ל-`@pointerdown.capture` על שורש המעטפת — ב-capture, כדי
 * שרכיב שעוצר הפצה בדרך לא ישתיק אותו.
 */
export function createCaretKeeper({ documentArea }: CaretKeeperOptions): (event: PointerEvent) => void {
  return (event: PointerEvent): void => {
    if (event.button !== 0) return;
    const target = event.target;
    if (!(target instanceof Element)) return;

    const area = documentArea();
    if (!area) return;

    const active = target.ownerDocument?.activeElement ?? null;
    if (
      shouldKeepCaret({
        onControl: target.closest(CARET_KEEPER_CONTROL_SELECTOR) !== null,
        inDocument: area.contains(target),
        documentFocused: active !== null && area.contains(active),
      })
    ) {
      event.preventDefault();
    }
  };
}

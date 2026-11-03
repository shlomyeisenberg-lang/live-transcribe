/**
 * Enter בדיאלוג = לחיצה על כפתור ברירת המחדל, כמו ב-Word.
 *
 * ## מה זה פותר
 *
 * הדיאלוגים כאן אינם `<form>`, ולכן אין להם „שליחה משתמעת” — ההתנהגות
 * שבדפדפן נותנת ל-Enter להפעיל את כפתור ברירת המחדל. מה שהיה במקומה: שורת
 * `@keydown.enter` על **חלק** מהשדות בכל דיאלוג. כלומר Enter עבד בשדה מספר,
 * ולא עבד בבורר, לא בתיבת סימון, ובעיקר לא כשהמיקוד על שורש הדיאלוג — וזה
 * בדיוק המצב שבו הוא נפתח („גופן מתקדם” ממקד את השורש בפתיחה). המשתמש מילא
 * שדה, לחץ Enter, ולא קרה כלום.
 *
 * מכאן שהמנגנון יושב על **השורש** ולא על השדות: כל Enter בדיאלוג עולה אליו.
 *
 * ## למה „ללחוץ על הכפתור” ולא „לקרוא ל-onSubmit”
 *
 * לכל דיאלוג שם אחר לפעולה (`onSubmit`, `onAdd`, `onMark`, `onApply`…) ותנאי
 * `:disabled` משלו. לחיצה על הכפתור עצמה יורשת את שניהם בחינם: פעולה שאסורה
 * ברגע זה פשוט לא תרוץ, בלי שהמנגנון הזה יצטרך לדעת דבר על הוולידציה. זה גם
 * מה שהופך את `[data-default-action]` להצהרה יחידה בכל דיאלוג.
 */

/**
 * מה ש-Enter עליו הוא של האלמנט עצמו, ולא של הדיאלוג.
 *
 * כפתור וקישור מופעלים כבר על ידי הדפדפן, ופריט בתפריט/ברשימה בוחר את עצמו.
 * לחטוף מהם את Enter פירושו כפתור „ביטול” שמאשר. וכך גם `<summary>`: Enter
 * עליו פותח את ה-`<details>` — ב„דיווח על באג” זה „מה יישלח”, ולחטוף ממנו את
 * Enter פירושו לשלוח את הדיווח במקום להציג אותו.
 */
const SELF_ACTIVATING = 'button, a[href], summary, [role="button"], [role="menuitem"], [role="option"], [role="tab"]';

export interface DialogDefaultAction {
  /** נקשר ב-`@keydown.enter` על שורש הדיאלוג. */
  onDialogEnter: (event: KeyboardEvent) => void;
}

export function useDialogDefaultAction(): DialogDefaultAction {
  function onDialogEnter(event: KeyboardEvent): void {
    // הרכבה במקלדת (IME): ה-Enter סוגר את הבחירה בחלונית ההרכבה, ואינו אישור.
    if (event.isComposing) return;
    // Enter „נקי” בלבד. Shift+Enter ו-Ctrl+Enter הם קיצורים אחרים, ובשדה
    // רב-שורות גם שורה חדשה.
    if (event.shiftKey || event.ctrlKey || event.metaKey || event.altKey) return;

    const target = event.target as HTMLElement | null;
    if (!target) return;
    // טקסט רב-שורות: Enter הוא שורה חדשה, וזו כל מהותו של השדה.
    if (target.closest('textarea') || target.isContentEditable) return;
    if (target.closest(SELF_ACTIVATING)) return;

    const root = event.currentTarget as HTMLElement | null;
    const action = root?.querySelector<HTMLButtonElement>('[data-default-action]');
    if (!action || action.disabled || action.getAttribute('aria-disabled') === 'true') return;

    event.preventDefault();
    // הדיאלוג בלע את ה-Enter: מה שמאזין מעליו (קיצורי המעטפת) לא יקבל אותו,
    // בדיוק כמו ה-Escape של אותו שורש.
    event.stopPropagation();
    action.click();
  }

  return { onDialogEnter };
}

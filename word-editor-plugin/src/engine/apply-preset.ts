/**
 * החלת ערכת עיצוב על הבחירה — הצד שמדבר עם המנוע.
 *
 * הצעדים עצמם נגזרים ב-ui/shortcuts/format-preset.ts (`presetSteps`), פונקציה
 * טהורה; כאן נשארת ההרצה בלבד, דרך אותו `CommandAdapter` שכל כפתור ברצועה
 * עובר בו. אין מסלול עוקף למנוע — קיצור אישי שנכשל מדבר עברית בשורת המצב
 * בדיוק כמו כפתור שנכשל.
 *
 * ## שתי הכרעות
 *
 * **מתג נבדק לפני שהוא נלחץ.** `bold` אינו „הדגש” אלא „הפוך הדגשה”, ולכן
 * ערכה שמבקשת „מודגש” על טקסט שכבר מודגש הייתה **מבטלת** את ההדגשה. זה הבאג
 * שהפרדת סוגי הצעדים נועדה למנוע, וכאן הוא נמנע בפועל: מתג שמצבו כבר המבוקש
 * אינו מורץ כלל.
 *
 * **ברצף ולא במקביל.** כל צעד ממתין לקודמו. `Promise.all` היה שולח ארבע
 * פקודות עיצוב לאותה בחירה בו-זמנית, ומצב הבחירה שכל אחת מהן קוראת היה תלוי
 * בסדר שבו ה-controller הספיק לנתב אותן.
 */
import type { CommandAdapter, CommandOutcome } from './command-adapter';
import { presetSteps, type FormatPreset } from '../ui/shortcuts/format-preset';

/** מה שנצרך מהמתאם. */
export type PresetTarget = Pick<CommandAdapter, 'has' | 'getState' | 'run'>;

/**
 * מריצה את הערכה. מחזירה את מספר הצעדים שרצו בפועל — `0` פירושו שלא היה מה
 * לעשות (ערכה ריקה, או ערכה שכל מתגיה כבר במצב המבוקש), וזה מה שמאפשר לקורא
 * להימנע מהודעה על פעולה שלא קרתה.
 *
 * `report` הוא אותו מדווח של הרצועה: הוא מציג כשל בעברית ומנקה שגיאה קודמת
 * בהצלחה. הכשל הראשון **אינו** עוצר את השאר — ערכה שהצבע שלה נדחה עדיין
 * אמורה להחיל את הגופן, בדיוק כמו שלחיצה על שני כפתורים ברצועה אינה נזנחת
 * באמצע.
 */
export async function applyPreset(
  target: PresetTarget,
  preset: FormatPreset,
  report: (outcome: CommandOutcome, commandId: string) => void,
): Promise<number> {
  let ran = 0;

  for (const step of presetSteps(preset)) {
    if (!target.has(step.command)) {
      // פקודה שהמנוע אינו מכיר: `run` היה מחזיר „הפעולה אינה מוכרת למנוע”,
      // שהיא הודעה על באג בקוד שלנו ולא על מה שהמשתמש ביקש. ערכה שנשמרה
      // בגרסה עדכנית ונקראת בגרסת מנוע ישנה אינה באג.
      report(
        { ok: false, message: 'הפעולה אינה נתמכת בגרסה הזאת של המנוע', reason: 'command-unsupported' },
        step.command,
      );
      continue;
    }

    if (step.kind === 'toggle') {
      if (target.getState(step.command).active === step.to) continue;
      report(await target.run(step.command), step.command);
      ran += 1;
      continue;
    }

    report(await target.run(step.command, step.payload), step.command);
    ran += 1;
  }

  return ran;
}

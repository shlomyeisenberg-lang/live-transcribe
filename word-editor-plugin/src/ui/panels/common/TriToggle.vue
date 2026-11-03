<template>
  <button
    type="button"
    class="tri-toggle"
    :class="`tri-${state}`"
    :data-state="state"
    :aria-pressed="pressed"
    :aria-label="`${label} — ${STATE_TEXT[state]}`"
    :data-tip-title="label"
    :data-tip-desc="tip"
    @pointerdown.prevent
    @click="cycle"
  >
    <span
      class="tri-mark"
      aria-hidden="true"
    >{{ MARK[state] }}</span>
    <span class="tri-label">{{ label }}</span>
  </button>
</template>

<script setup lang="ts">
/**
 * כפתור אפקט בדיאלוג — נדלק, נכבה, וחוזר ל„לא נגעתי”.
 *
 * ## למה שלושה מצבים ולא שניים
 *
 * הפקד הזה מרכיב patch: מפתח שאינו נשלח אינו נוגע בעיצוב קיים. כפתור
 * דו-מצבי אינו יודע לומר את זה — הוא אומר או „כן” או „לא”, ו„לא” על טקסט
 * שהמשתמש לא התכוון לגעת בו **מוחק ממנו אפקט**. זה לא היפותטי: „אישור” על
 * דיאלוג שנפתח ונסגר היה מסיר צל, מסגרת וחריטה מכל מה שסומן.
 *
 * הפתרון הקל — כפתור דו-מצבי שמציג את מצב הבחירה — אינו זמין, וזו מדידה ולא
 * הערכה: אף אחד משלושת המשטחים שהרצועה קוראת מהם אינו מדווח את **הערך** של
 * `dstrike`/`outline`/`shadow`/`emboss`/`imprint`/`vanish` ושל משפחת ה-CS.
 * `query.match()` מדווח bold/italic/underline/strike, צבע, הדגשה, גופן וגודל
 * — וזה הכול. הרישום המלא ב-engine/vert-align.ts („למה אין מצב דלוק”)
 * וב-docs/engine-gaps.md.
 *
 * כלומר: כפתור דו-מצבי כאן היה **משקר** על טקסט שכבר נושא את האפקט, ומצב
 * שלישי אינו קישוט אלא מה שמפריד בין „כבה” לבין „איני יודע”.
 *
 * ## ומה יסיר אותו
 *
 * ברגע שהמנוע ידווח את מצב הריצה, הכפתור יוכל להיפתח על המצב האמיתי ולהיות
 * דו-מצבי כמו B/I/U ברצועה. זה מעקב אחר
 * https://github.com/superdoc/docx-editor/issues/3982 — עד שייסגר, המצב
 * הנייטרלי הוא מה שמחזיק את ההבטחה „מה שלא נגעת בו לא ישתנה”.
 *
 * ## המחזור, והסימן
 *
 * לחיצה: „ללא שינוי” → „דלוק” → „כבוי” → „ללא שינוי”. הסימן משמאל לתווית
 * מצויר תמיד — גם במצב הנייטרלי — כדי שהכפתור לא ישנה רוחב תוך כדי מחזור:
 * שורת כפתורים שזזה מתחת לאצבע היא שורה שאי אפשר ללחוץ בה פעמיים ברצף.
 * והוא טקסט ולא צבע בלבד, מפני שצבע לבדו אינו הבדל שכולם רואים.
 */
import { computed } from 'vue';

/** שלושת המצבים, בערכים שה-patch מדבר בהם. `''` = לא נשלח. */
export type TriState = '' | 'yes' | 'no';

const MARK: Record<TriState, string> = { '': '–', yes: '✓', no: '✕' };
const STATE_TEXT: Record<TriState, string> = { '': 'ללא שינוי', yes: 'דלוק', no: 'כבוי' };
const TIP_DESC = 'לחיצה: הדלקה ← כיבוי ← ללא שינוי';

/** הסדר שלחיצה מתקדמת בו. */
const CYCLE: Record<TriState, TriState> = { '': 'yes', yes: 'no', no: '' };

const props = defineProps<{
  modelValue: TriState;
  label: string;
  /**
   * מה שהאפקט עושה — נוסף להסבר המחזור בטולטיפ.
   *
   * התוויות בדיאלוג קצרות בכוונה („מודגש”, „שקוע”), מפני שהן יושבות ברשת
   * כפתורים ולא בשורת תווית. מה שנחתך מהן — ובעברית זה בדיוק ההבדל שקובע,
   * `w:bCs` מול `w:b` — חוזר כאן, למי שעוצר על הכפתור.
   */
  description?: string;
}>();

const emit = defineEmits<{ 'update:modelValue': [value: TriState] }>();

const tip = computed(() => (props.description ? `${props.description} · ${TIP_DESC}` : TIP_DESC));

const state = computed<TriState>(() => props.modelValue);

/**
 * `aria-pressed="mixed"` למצב הנייטרלי — הערך שהתקן מייעד לכפתור מיתוג
 * תלת-מצבי, ובדיוק המשמעות כאן: „אינו לחוץ ואינו משוחרר”. התווית מוסיפה את
 * המצב במילים, מפני ש-`mixed` לבדו אינו אומר „לא ייגע”.
 */
const pressed = computed(() => (state.value === '' ? 'mixed' : state.value === 'yes' ? 'true' : 'false'));

function cycle(): void {
  emit('update:modelValue', CYCLE[state.value]);
}
</script>

<style scoped>
.tri-toggle {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  min-width: 0;
  padding: 3px 8px;
  font-family: var(--font-main);
  font-size: 11px;
  line-height: 1.4;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface-variant);
  cursor: pointer;
  transition: background 0.08s, border-color 0.08s;
}

.tri-toggle:hover {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

.tri-toggle:focus-visible {
  outline: none;
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

/* רוחב קבוע לסימן: מחזור המצבים אינו מזיז את התווית שלידו. */
.tri-mark {
  flex: 0 0 auto;
  width: 9px;
  text-align: center;
  font-size: 10px;
}

.tri-label {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

/* „דלוק” — אותה שפה של פקד דלוק ברצועה, ומאותם טוקנים. */
.tri-yes {
  background: var(--word-btn-active);
  border-color: var(--word-btn-active-border);
  color: var(--color-on-surface);
  font-weight: 600;
}

.tri-yes:hover {
  background: var(--word-btn-active-hover);
}

.tri-yes .tri-mark {
  color: var(--word-blue);
}

/*
  „כבוי” אינו מצב חיובי, ולכן אינו נצבע כמוהו — אבל הוא כן **נבחר**, ולכן
  אינו יכול להיראות כמו הנייטרלי שלידו. המסגרת המלאה היא ההבדל, והסימן הוא
  מה שאומר אותו במילים.
*/
.tri-no {
  border-color: var(--color-outline);
  color: var(--color-on-surface);
  font-weight: 600;
}

.tri-no .tri-mark {
  color: var(--color-error);
}
</style>

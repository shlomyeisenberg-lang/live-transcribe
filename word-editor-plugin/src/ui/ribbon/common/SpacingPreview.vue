<template>
  <div class="sp-preview">
    <div
      v-if="geometry"
      class="sp-page"
      :dir="rtl ? 'rtl' : 'ltr'"
      aria-hidden="true"
    >
      <div
        class="sp-para sp-context"
        :style="contextStyle"
      >
        <span
          v-for="line in PREVIEW_NEIGHBOUR_LINES"
          :key="`before-${line}`"
          class="sp-line"
        />
      </div>
      <div
        class="sp-para sp-target"
        :style="targetStyle"
      >
        <span
          v-for="line in PREVIEW_TARGET_LINES"
          :key="`target-${line}`"
          class="sp-line"
          :class="{ 'sp-line-last': line === PREVIEW_TARGET_LINES }"
        />
      </div>
      <div
        class="sp-para sp-context"
        :style="contextStyle"
      >
        <span
          v-for="line in PREVIEW_NEIGHBOUR_LINES"
          :key="`after-${line}`"
          class="sp-line"
        />
      </div>
    </div>
    <p
      v-else
      class="sp-note"
      role="note"
    >
      {{ menuString(NO_GEOMETRY_HINT) }}
    </p>
    <span class="sp-caption">{{ caption }}</span>
  </div>
</template>

<script setup lang="ts">
/**
 * פס התצוגה המקדימה של תפריט „מרווח שורות וריווח” — שלוש פסקאות במיניאטורה,
 * שמראות מה הלחיצה תעשה **לפני** שהיא נעשית.
 *
 * ## למה פס בתפריט, ולא תצוגה על הדף עצמו
 *
 * זו הייתה הכוונה הראשונה, והיא נמדדה ונשללה. הרשומה המלאה ב-
 * `docs/engine-gaps.md` („תצוגה מקדימה אנכית של ריווח”), והגשש עצמו ב-
 * `scripts/spacing-preview-probe.mjs`. בקצרה: תיבת הפסקה במנוע היא
 * `position: absolute` בגובה קשיח, ולכן שינוי חזותי של מרווח השורות מגדיל את
 * הטקסט **בתוך** תיבה שאינה גדלה — נמדדה חפיפה של 40.8px על הפסקה הבאה. פיצוי
 * בקסקייד (הזזת כל האחיות שאחרי היעד) כן נותן תמונה מדויקת לפיקסל בתוך העמוד,
 * ובכל זאת נשלל: הצירוף `~` אינו חוצה עמודים, והשינוי האמיתי כן — נמדד שהרכב
 * העמודים עבר מ-17/16 ל-16/16, כלומר פסקה נדחקה לעמוד הבא. התצוגה הייתה
 * מראה טקסט גולש מתחת לשולי העמוד במקום עמוד שמתעמד מחדש, ופס הבחירה הכחול
 * נשאר במקומו בזמן שהטקסט זז מתחתיו.
 *
 * זו אותה הכרעה שננקטה בתצוגה החיה של הגופן (`engine/font-preview.ts`): תצוגה
 * שמראה משהו אחר ממה שיקרה בפועל גרועה מאין תצוגה, ולכן היא יוצאת מהמסמך אל
 * הפקד.
 *
 * ## ולמה הוא גדול מזה שבדיאלוג
 *
 * בדיאלוג הפס הוא עד אחד מני שישה: הכניסות, „מיוחד” והריווח יושבים בשדות
 * מספריים שמעליו, והוא מאשר אותם. כאן הוא **הדבר היחיד שאומר כמה**: התפריט
 * נוקב בשמות („הוסף רווח לפני הפסקה”) ולא במספרים, ומרווח שורות הוא יחס שאין
 * לו משמעות בלי לראות אותו. לכן `block-size` גדול יותר וגופן גדול יותר.
 *
 * ## החישוב אינו כאן
 *
 * `paragraphPreviewGeometry` (engine/paragraph-preview.ts) הוא אותו חישוב
 * שהדיאלוג מצייר לפיו, על כל ההנמקה שלו — למה אחוזים ו-`cqw` ולא פיקסלים,
 * ולמה פסים ולא אותיות. שכפול שלו כאן היה מייצר שתי תצוגות שמתפצלות ביום
 * שהראשונה תתוקן.
 *
 * הכניסות נמסרות כאפס תמיד: התפריט אינו נוגע בהן, ופס שהיה מצייר את הכניסות
 * של הפסקה היה אומר שהתפריט משנה גם אותן.
 */
import { computed, type CSSProperties } from 'vue';
import {
  paragraphPreviewGeometry,
  PREVIEW_NEIGHBOUR_LINES,
  PREVIEW_TARGET_LINES,
} from '../../../engine/paragraph-preview';
import { TWIPS_PER_PT, type LineSpacingRule } from '../../../engine/paragraph-format';
import { menuString } from '../i18n';

const NO_GEOMETRY_HINT = 'התצוגה תופיע כשהמסמך ייטען';

const props = defineProps<{
  /** רוחב עמודת הטקסט של המקטע, ב-twips. `0` = אין גיאומטריה, ואז אין פס. */
  textWidthTwips: number;
  /** גודל הגופן שבסמן, בנקודות — הסרגל של הציר האנכי. */
  fontSizePt: number;
  beforeTwips: number;
  afterTwips: number;
  lineTwips: number;
  lineRule: LineSpacingRule;
  rtl: boolean;
}>();

const geometry = computed(() =>
  paragraphPreviewGeometry({
    textWidthTwips: props.textWidthTwips,
    fontSizePt: props.fontSizePt,
    startTwips: 0,
    endTwips: 0,
    special: 'none',
    amountTwips: 0,
    beforeTwips: props.beforeTwips,
    afterTwips: props.afterTwips,
    lineTwips: props.lineTwips,
    lineRule: props.lineRule,
  }),
);

const targetStyle = computed<CSSProperties>(() => {
  const g = geometry.value;
  if (!g) return {};
  return {
    '--pv-font': `${g.fontCqw}cqw`,
    marginBlockStart: `${g.beforePct}%`,
    marginBlockEnd: `${g.afterPct}%`,
    '--pv-lh': String(g.lineHeight),
  };
});

const contextStyle = computed<CSSProperties>(() => {
  const g = geometry.value;
  if (!g) return {};
  return { '--pv-font': `${g.fontCqw}cqw`, '--pv-lh': '1' };
});

/**
 * מה שהפס מראה, במילים. הוא אינו קישוט: „לפני 12 נק'” הוא ההבדל היחיד שנראה
 * בין „הוסף רווח לפני” ל„הוסף רווח אחרי” כשהמרחק קטן, ו„בודדת”/„1.5” הוא מה
 * שהופך יחס לשם מוכר.
 *
 * המרחק נאמר בנקודות גם כשהוא אפס — „לפני 0” הוא תשובה, ופס שהיה משמיט אותו
 * היה מציג את אותה שורה בדיוק לפני הלחיצה ואחריה.
 */
const caption = computed(
  () =>
    `${lineText.value} · ${menuString('לפני')} ${pt(props.beforeTwips)} · ` +
    `${menuString('אחרי')} ${pt(props.afterTwips)} ${menuString('נק׳')}`,
);

/** נקודות, בלי שבר מיותר: „12”, ולא „12.0”. */
function pt(twips: number): string {
  const value = twips / TWIPS_PER_PT;
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

/** „שורה בודדת”, „1.5 שורות”, או „בדיוק 18 נק׳” — בדיוק השמות שבדיאלוג. */
const lineText = computed(() => {
  if (props.lineRule !== 'auto') {
    const label = props.lineRule === 'exact' ? menuString('בדיוק') : menuString('לפחות');
    return `${label} ${pt(props.lineTwips)} ${menuString('נק׳')}`;
  }
  const multiple = props.lineTwips / 240;
  if (multiple === 1) return menuString('שורה בודדת');
  if (multiple === 2) return menuString('שורה כפולה');
  const shown = Number.isInteger(multiple) ? String(multiple) : multiple.toFixed(2).replace(/0$/, '');
  // „מרובה 1.5” ולא „1.5 שורות”: `שורות` כבר תפוס במילון כ-`rows` של הטבלה,
  // ושני מובנים לאותו מפתח הם תרגום שגוי באחד משניהם. זה גם השם ש-Word נותן
  // לכלל הזה בעצמו.
  return `${menuString('מרובה')} ${shown}`;
});
</script>

<style scoped>
.sp-preview {
  display: flex;
  flex-direction: column;
  gap: 3px;
}

/*
 * `container-type: inline-size` הוא מה שהופך את הרוחב הזה למכנה של `cqw` —
 * אותה מכניקה כמו בדיאלוג, וההנמקה ב-engine/paragraph-preview.ts.
 *
 * גובה קבוע: הפס משתנה בכל ריחוף על פריט אחר, וגובה שנגזר מהתוכן היה מזיז את
 * תחתית התפריט מתחת לעכבר — כלומר פריט שבורח בדיוק כשמכוונים אליו.
 */
.sp-page {
  container-type: inline-size;
  block-size: 150px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  padding-inline: 8px;
  overflow: hidden;
  border-radius: var(--radius-xs);
  background: var(--color-surface-container-high);
}

.sp-para {
  margin: 0;
  display: flex;
  flex-direction: column;
  /* דרך משתנה ולא ישירות: `cqw` נבלעת ב-jsdom בשקט. ההנמקה המלאה בדיאלוג. */
  font-size: var(--pv-font);
}

.sp-target {
  color: var(--color-on-surface);
}

.sp-context {
  color: var(--color-on-surface-variant);
  opacity: 0.45;
}

.sp-line {
  block-size: calc(1em * var(--pv-lh, 1));
  display: flex;
  align-items: center;
}

.sp-line::before {
  content: '';
  display: block;
  inline-size: 100%;
  block-size: 0.42em;
  background: currentcolor;
  border-radius: 1px;
}

.sp-line-last::before {
  inline-size: 62%;
}

.sp-context .sp-line:last-child::before {
  inline-size: 74%;
}

.sp-caption {
  font-size: 10px;
  color: var(--color-on-surface-variant);
  text-align: center;
}

.sp-note {
  margin: 0;
  padding: 12px 8px;
  font-size: 10px;
  text-align: center;
  color: var(--color-on-surface-variant);
}
</style>

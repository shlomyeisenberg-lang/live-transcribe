<template>
  <Teleport to="body">
    <div
      v-if="isOpen"
      ref="rootRef"
      class="para-dialog"
      :style="dragStyle"
      role="dialog"
      aria-modal="true"
      :aria-label="DIALOG_TITLE"
      tabindex="-1"
      @keydown.esc.stop="$emit('close')"
      @keydown.enter="onDialogEnter"
    >
      <div
        class="pd-header dialog-drag-handle"
        @pointerdown="startDialogDrag"
      >
        <span class="pd-title">{{ DIALOG_TITLE }}</span>
        <button
          type="button"
          class="pd-close-btn"
          data-tip-title="סגור"
          data-tip-shortcut="Esc"
          aria-label="סגור את תפריט הפסקה"
          @click="$emit('close')"
        >
          ✕
        </button>
      </div>

      <div class="pd-body">
        <div class="pd-columns">
          <!-- עמודה א׳: המידות — כניסות וריווח -->
          <div class="pd-column">
            <!-- כניסות -->
            <fieldset class="pd-group">
              <legend>כניסה</legend>
              <div class="pd-row">
                <label
                  for="pd-ind-left"
                  class="pd-label"
                >לפני טקסט:</label>
                <input
                  id="pd-ind-left"
                  v-model="leftCm"
                  class="pd-number"
                  type="number"
                  min="0"
                  max="55.87"
                  step="0.1"
                  aria-label="כניסה לפני הטקסט, בסנטימטרים"
                >
                <span class="pd-unit">ס"מ</span>
              </div>
              <div class="pd-row">
                <label
                  for="pd-ind-right"
                  class="pd-label"
                >אחרי טקסט:</label>
                <input
                  id="pd-ind-right"
                  v-model="rightCm"
                  class="pd-number"
                  type="number"
                  min="0"
                  max="55.87"
                  step="0.1"
                  aria-label="כניסה אחרי הטקסט, בסנטימטרים"
                >
                <span class="pd-unit">ס"מ</span>
              </div>
              <div class="pd-row">
                <label
                  for="pd-special"
                  class="pd-label"
                >מיוחד:</label>
                <select
                  id="pd-special"
                  v-model="special"
                  class="pd-select"
                >
                  <option value="none">ללא</option>
                  <option value="firstLine">שורה ראשונה</option>
                  <option value="hanging">תלויה</option>
                </select>
                <input
                  v-model="amountCm"
                  class="pd-number"
                  type="number"
                  min="0"
                  step="0.1"
                  :disabled="special === 'none'"
                  aria-label="מידת הכניסה המיוחדת, בסנטימטרים"
                >
                <span class="pd-unit">ס"מ</span>
              </div>
            </fieldset>

            <!-- ריווח -->
            <fieldset class="pd-group">
              <legend>ריווח</legend>
              <div class="pd-row">
                <label
                  for="pd-sp-before"
                  class="pd-label"
                >לפני:</label>
                <input
                  id="pd-sp-before"
                  v-model="beforePt"
                  class="pd-number"
                  type="number"
                  min="0"
                  step="1"
                  aria-label="ריווח לפני הפסקה, בנקודות"
                >
                <span class="pd-unit">נק'</span>
                <label
                  for="pd-sp-after"
                  class="pd-label pd-label-after"
                >אחרי:</label>
                <input
                  id="pd-sp-after"
                  v-model="afterPt"
                  class="pd-number"
                  type="number"
                  min="0"
                  step="1"
                  aria-label="ריווח אחרי הפסקה, בנקודות"
                >
                <span class="pd-unit">נק'</span>
              </div>
              <div class="pd-row">
                <label
                  for="pd-line"
                  class="pd-label"
                >מרווח שורות:</label>
                <select
                  id="pd-line"
                  v-model="lineMode"
                  class="pd-select"
                >
                  <option value="240">בודדת</option>
                  <option value="360">1.5 שורות</option>
                  <option value="480">כפולה</option>
                  <option value="exact">מדויקת</option>
                  <option value="atLeast">לפחות</option>
                </select>
                <template v-if="lineMode === 'exact' || lineMode === 'atLeast'">
                  <input
                    v-model="linePt"
                    class="pd-number"
                    type="number"
                    min="0"
                    step="1"
                    aria-label="גובה השורה, בנקודות"
                  >
                  <span class="pd-unit">נק'</span>
                </template>
              </div>
              <!--
                „אל תוסיף רווח בין פסקאות מאותו סגנון” (`w:contextualSpacing`).
                ב-Word זו תיבת סימון, וכאן היא פקד תלת-מצבי — לא בחירת עיצוב
                אלא מדידה: המנוע **אינו מחזיר** את הערך (נמדד; ראו
                `applyParagraphContextualSpacing`), ותיבה דו-מצבית הייתה
                נפתחת תמיד ריקה ו„אישור” היה מכבה הגדרה שהמשתמש קבע ב-Word.
                „ללא שינוי” אינו שולח כלום, והפעולה במנוע היא patch.
              -->
              <div class="pd-row pd-row-tri">
                <TriToggle
                  v-model="contextual"
                  label="אל תוסיף רווח בין פסקאות מאותו סגנון"
                  description="מבטל את הריווח שלפני ואחרי בין פסקאות רצופות שיש להן אותו סגנון"
                />
              </div>
            </fieldset>
          </div>

          <!-- עמודה ב׳: מה שאינו מידה — שמירה וטאבים -->
          <div class="pd-column">
            <!-- אפשרויות שמירה -->
            <fieldset class="pd-group">
              <legend>אפשרויות שמירה</legend>
              <div class="pd-checks">
                <label class="pd-check">
                  <input
                    v-model="keepNext"
                    type="checkbox"
                  >
                  השאר עם הבא
                </label>
                <label class="pd-check">
                  <input
                    v-model="keepLines"
                    type="checkbox"
                  >
                  השאר שורות יחד
                </label>
                <label class="pd-check">
                  <input
                    v-model="widowControl"
                    type="checkbox"
                  >
                  בקרת אלמנות ויתומים
                </label>
              </div>
            </fieldset>

            <!-- טאבים -->
            <fieldset
              v-if="tabsEnabled"
              class="pd-group"
            >
              <legend>עצירות טאב</legend>
              <p
                class="pd-note"
                role="note"
              >
                הוספה והסרה מוחלות מיד על הפסקה שבה הסמן.
              </p>
              <ul
                v-if="tabs.length > 0"
                class="pd-tabs"
              >
                <li
                  v-for="(tab, index) in tabs"
                  :key="`${tab.positionTwips}-${index}`"
                  class="pd-tab-row"
                >
                  <span>{{ formatTab(tab) }}</span>
                  <button
                    type="button"
                    class="pd-btn pd-btn-small"
                    :disabled="busy"
                    aria-label="הסר עצירת טאב"
                    @pointerdown.prevent
                    @click="$emit('tab-remove', { positionTwips: tab.positionTwips })"
                  >
                    הסר
                  </button>
                </li>
              </ul>
              <p
                v-else
                class="pd-note"
              >
                אין עצירות טאב בפסקה זו.
              </p>
              <div class="pd-row pd-row-wrap">
                <input
                  v-model="newTabCm"
                  class="pd-number"
                  type="number"
                  min="0.1"
                  step="0.1"
                  placeholder="מיקום"
                  aria-label="מיקום עצירת טאב חדשה, בסנטימטרים"
                >
                <select
                  v-model="newTabAlignment"
                  class="pd-select"
                  aria-label="יישור עצירת הטאב"
                >
                  <option value="left">שמאל</option>
                  <option value="center">מרכז</option>
                  <option value="right">ימין</option>
                  <option value="decimal">עשרוני</option>
                </select>
                <select
                  v-model="newTabLeader"
                  class="pd-select"
                  aria-label="מוביל עצירת הטאב"
                >
                  <option value="">ללא</option>
                  <option value="dot">........</option>
                  <option value="hyphen">--------</option>
                  <option value="underscore">________</option>
                </select>
                <button
                  type="button"
                  class="pd-btn"
                  :disabled="busy || !canAddTab"
                  @pointerdown.prevent
                  @click="onAddTab"
                >
                  הוסף
                </button>
              </div>
              <button
                v-if="tabs.length > 0"
                type="button"
                class="pd-btn"
                :disabled="busy"
                @pointerdown.prevent
                @click="$emit('tabs-clear')"
              >
                נקה את כל העצירות
              </button>
            </fieldset>
          </div>
        </div>

        <!--
          פס התצוגה המקדימה — מיניאטורה של עמודת הטקסט, בקנה מידה אחד לשני
          הצירים. ההנמקה כולה — למה אחוזים ו-`cqw` ולא פיקסלים, ולמה פסי שורה
          ולא אותיות — ב-engine/paragraph-preview.ts.

          `aria-hidden` כמו פס הדגימה של בורר הגופן: כל שש ההגדרות שהוא מצייר
          יושבות בשדות שמעליו, והפסים עצמם אינם מוסיפים דבר למי שקורא מסך.
        -->
        <div class="pd-preview">
          <span class="pd-preview-caption">
            {{ PREVIEW_CAPTION }}<template v-if="scaleText !== ''"> · {{ scaleText }}</template>
          </span>
          <div
            v-if="preview"
            class="pd-preview-page"
            :dir="previewDir"
            aria-hidden="true"
          >
            <div
              class="pd-pv-para pd-pv-context"
              :style="contextStyle"
            >
              <span
                v-for="line in PREVIEW_NEIGHBOUR_LINES"
                :key="`before-${line}`"
                class="pd-pv-line"
              />
            </div>
            <div
              class="pd-pv-para pd-pv-target"
              :style="targetStyle"
            >
              <span
                v-for="line in PREVIEW_TARGET_LINES"
                :key="`target-${line}`"
                class="pd-pv-line"
                :class="{ 'pd-pv-line-last': line === PREVIEW_TARGET_LINES }"
                :style="line === 1 ? firstLineStyle : lineStyle"
              />
            </div>
            <div
              class="pd-pv-para pd-pv-context"
              :style="contextStyle"
            >
              <span
                v-for="line in PREVIEW_NEIGHBOUR_LINES"
                :key="`after-${line}`"
                class="pd-pv-line"
              />
            </div>
          </div>
          <p
            v-else
            class="pd-note pd-preview-note"
            role="note"
          >
            {{ NO_GEOMETRY_HINT }}
          </p>
          <p
            v-if="preview && !preview.leavesRoom"
            class="pd-note pd-preview-note pd-note-warn"
            role="note"
          >
            {{ NO_ROOM_HINT }}
          </p>
        </div>

        <p
          v-if="showError"
          class="pd-error"
          role="alert"
        >
          {{ INVALID_HINT }}
        </p>

      </div>

      <div class="pd-footer">
        <button
          type="button"
          class="pd-btn pd-btn-primary"
          data-default-action
          :disabled="busy || !canSubmit"
          @pointerdown.prevent
          @click="onSubmit"
        >
          אישור
        </button>
        <button
          type="button"
          class="pd-btn"
          @pointerdown.prevent
          @click="$emit('close')"
        >
          ביטול
        </button>
      </div>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
/**
 * „פסקה” — כניסות, ריווח, אפשרויות שמירה וטאבים; ההנמקות ב-engine/paragraph-format.ts.
 *
 * שתי הכרעות מבניות:
 *
 * 1. **מילוי מוקדם מהמסמך.** `setIndentation`/`setSpacing` מחליפים את
 *    `<w:ind>`/`<w:spacing>` כולו (נמדד), ולכן הדיאלוג נפתח על ה-prop
 *    `snapshot` — תצלום שקרא `readParagraphFormat` — ואישור שולח מצב מלא.
 *    דיאלוג שנפתח ריק היה מוחק בשקט כניסות קיימות.
 * 2. **הטאבים מיידיים** ושאר הסעיפים ב„אישור”. `setTabStop` מוסיף בלי לגעת
 *    באחרות (נמדד), כך שהוספה והסרה הן פעולות עצמאיות בטוחות; הכניסות
 *    והריווח הן replace ודורשות אישור אחד ששולח את המצב כולו.
 *
 * ## הפריסה: שתי עמודות ופס מתחתן
 *
 * שתי עמודות מ„מתקדם” של הגופן (FontAdvancedDialog), ומאותו טעם בדיוק:
 * הרוחב גדל, והגובה — שאליו נוסף עכשיו פס התצוגה המקדימה — יורד מתחת
 * לתקרה `calc(100vh - 200px)` במקום לגלול. החלוקה היא לפי סוג ההגדרה ולא
 * לפי גודל הסעיף: **המידות** (כניסות, ריווח) בעמודה אחת — הן מה שהפס
 * מצייר — ומה שאינו מידה (אפשרויות שמירה, עצירות טאב) בשנייה.
 *
 * וסדר ה-DOM של השדות נשמר בדיוק כשהיה — `left, right, amount, before,
 * after, linePt, newTabCm` — מפני ש-`scripts/qa/home-paragraph-qa.mjs` פונה
 * לשלושה מהם לפי מקומם ברשימה (`input[type=number]` באינדקסים 2 ו-5, והטאב
 * ב„אחרון”) ולא לפי מזהה, ואין להם מזהה לפנות אליו.
 *
 * רשימת הדיאלוג (checklist): tabindex="-1" + focus() ב-nextTick, Enter על
 * שדות קלט, prop `busy` שמנטרל את פעולות המסמך ומשאיר את „ביטול”/Esc חיים,
 * RTL לוגי בלבד, ו-@pointerdown.prevent על כל כפתור.
 */
import { computed, nextTick, ref, watch, type CSSProperties } from 'vue';
import {
  TWIPS_PER_CM,
  TWIPS_PER_PT,
  type LineSpacingRule,
  type TabAlignment,
  type TabLeader,
} from '../../engine/paragraph-format';
import {
  PREVIEW_NEIGHBOUR_LINES,
  PREVIEW_TARGET_LINES,
  paragraphPreviewGeometry,
} from '../../engine/paragraph-preview';
import TriToggle, { type TriState } from './common/TriToggle.vue';
import { useDialogDrag } from '../../composables/dialog-drag';
import { useDialogDefaultAction } from '../../composables/dialog-default-action';

/* הדיאלוג נגרר בכותרת שלו — composables/dialog-drag.ts. */
const { dragStyle, startDialogDrag } = useDialogDrag();
/* Enter = הכפתור הראשי — composables/dialog-default-action.ts. */
const { onDialogEnter } = useDialogDefaultAction();

const DIALOG_TITLE = 'פסקה';
const INVALID_HINT = 'הערכים חייבים להיות מספרים חוקיים ולא-שליליים.';
const PREVIEW_CAPTION = 'תצוגה מקדימה';
/**
 * מה שנאמר במקום הפס כשאין ממה לגזור קנה מידה. „עדיין לא דיווח” ולא
 * „שגיאה”: `readPageMargins` מחזיר `null` בדיוק במצב הזה, וכך גם הסרגל
 * מתנהג — הוא פשוט אינו מצייר.
 *
 * ובלי להבטיח שהפס יופיע מעצמו: המידות נקראות **בפתיחה** (ראו
 * `pageTextWidthTwips`), ולכן דיאלוג שנפתח בלעדיהן לא יקבל אותן עד הפתיחה
 * הבאה. מה שנאמר הוא הסיבה ומה שכן עובד, ולא הבטחה.
 */
const NO_GEOMETRY_HINT = 'המסמך עדיין לא דיווח מידות עמוד, ואין ממה לגזור את הפס. שאר ההגדרות עובדות כרגיל.';
const NO_ROOM_HINT = 'שתי הכניסות יחד אינן משאירות רוחב שאפשר לכתוב בו.';

const props = defineProps<{
  isOpen: boolean;
  busy: boolean;
  tabsEnabled: boolean;
  snapshot: {
    indentation: { leftTwips: number; rightTwips: number; firstLineTwips: number; hangingTwips: number };
    spacing: { beforeTwips: number; afterTwips: number; lineTwips: number; rule: string };
    keepNext: boolean;
    keepLines: boolean;
    widowControl: boolean;
    tabs: readonly { positionTwips: number; alignment: string; leader?: string }[];
    /**
     * `<w:bidi>` של הפסקה — הצד שבו „לפני טקסט” יושב בפס. `null` = הפסקה
     * אינה מצהירה, ואז הכיוון יורש מ-`sectionRtl`.
     */
    bidi: boolean | null;
  };
  /**
   * רוחב עמודת הטקסט של המקטע — רוחב הדף פחות שני השוליים, מ-`readPageMargins`.
   * `0` = אין גיאומטריה, ואז אין פס. זהו המכנה של כל האחוזים בפס.
   */
  pageTextWidthTwips: number;
  /** גודל הגופן שבסמן, בנקודות — הסרגל של הציר האנכי בפס. */
  fontSizePt: number;
  /**
   * כיוון המקטע — מה שפסקה **שאינה מצהירה** יורשת. ראו `previewDir`.
   */
  sectionRtl: boolean;
}>();

const emit = defineEmits<{
  close: [];
  submit: [
    payload: {
      leftCm: number;
      rightCm: number;
      special: 'none' | 'firstLine' | 'hanging';
      amountCm: number;
      beforePt: number;
      afterPt: number;
      lineTwips: number;
      lineRule: LineSpacingRule;
      keepNext: boolean;
      keepLines: boolean;
      widowControl: boolean;
      /** `null` = „ללא שינוי”, ואז ההורה אינו קורא למנוע כלל. */
      contextualSpacing: boolean | null;
    },
  ];
  'tab-add': [tab: { positionTwips: number; alignment: TabAlignment; leader?: TabLeader }];
  'tab-remove': [payload: { positionTwips: number }];
  'tabs-clear': [];
}>();

const cmToTwips = (cm: number): number => Math.round(cm * TWIPS_PER_CM);
const twipsToCmText = (twips: number): string => (twips / TWIPS_PER_CM).toFixed(2);
const ptToTwips = (pt: number): number => Math.round(pt * TWIPS_PER_PT);
const twipsToPtText = (twips: number): string => String(Math.round(twips / TWIPS_PER_PT));

const rootRef = ref<HTMLElement | null>(null);

/** השדות הנערכים. מתאפסים מה-snapshot בכל פתיחה — ראו ה-watch למטה. */
const leftCm = ref('0');
const rightCm = ref('0');
const special = ref<'none' | 'firstLine' | 'hanging'>('none');
const amountCm = ref('1.25');
const beforePt = ref('0');
const afterPt = ref('0');
const lineMode = ref<string>('240');
const linePt = ref('12');
/**
 * „אל תוסיף רווח בין פסקאות מאותו סגנון”. נפתח תמיד ב„ללא שינוי” — ולא
 * מהתצלום — מפני שהמנוע אינו מחזיר את הערך כלל (נמדד; ראו
 * `applyParagraphContextualSpacing`). מצב שאינו „ללא שינוי” הוא היחיד שנשלח.
 */
const contextual = ref<TriState>('');

const keepNext = ref(false);
const keepLines = ref(false);
const widowControl = ref(true);

/** עותק מקומי של רשימת הטאבים; ההוספות/הסרות עצמן מופעלות אצל ההורה. */
const tabs = ref<Array<{ positionTwips: number; alignment: string; leader?: string }>>([]);

watch(
  () => props.isOpen,
  async (open) => {
    if (!open) return;
    const snap = props.snapshot;
    leftCm.value = twipsToCmText(snap.indentation.leftTwips);
    rightCm.value = twipsToCmText(snap.indentation.rightTwips);
    // „מיוחד”: שתי התכונות לעולם לא קיימות יחד ב-OOXML; מה שקיים מוצג.
    if (snap.indentation.firstLineTwips > 0) {
      special.value = 'firstLine';
      amountCm.value = twipsToCmText(snap.indentation.firstLineTwips);
    } else if (snap.indentation.hangingTwips > 0) {
      special.value = 'hanging';
      amountCm.value = twipsToCmText(snap.indentation.hangingTwips);
    } else {
      special.value = 'none';
      amountCm.value = twipsToCmText(567);
    }
    beforePt.value = twipsToPtText(snap.spacing.beforeTwips);
    afterPt.value = twipsToPtText(snap.spacing.afterTwips);
    if (snap.spacing.rule === 'exact' || snap.spacing.rule === 'atLeast') {
      lineMode.value = snap.spacing.rule;
      linePt.value = twipsToPtText(snap.spacing.lineTwips);
    } else {
      const known = ['240', '360', '480'];
      lineMode.value = known.includes(String(snap.spacing.lineTwips)) ? String(snap.spacing.lineTwips) : '240';
      linePt.value = twipsToPtText(snap.spacing.lineTwips || 240);
    }
    contextual.value = '';
    keepNext.value = snap.keepNext;
    keepLines.value = snap.keepLines;
    widowControl.value = snap.widowControl;
    tabs.value = [...snap.tabs];

    // פריט 1 ב-checklist: בלי מיקוד על השורש, Escape מפסיק לעבוד אחרי קליק
    // על גוף הדיאלוג.
    await nextTick();
    rootRef.value?.focus();
  },
);

function parseNonNegative(value: string): number | null {
  const parsed = Number(value);
  return typeof parsed === 'number' && Number.isFinite(parsed) && parsed >= 0 ? parsed : null;
}

function parsePositive(value: string): number | null {
  const parsed = Number(value);
  return typeof parsed === 'number' && Number.isFinite(parsed) && parsed > 0 ? parsed : null;
}

const newTabCm = ref('');
const newTabAlignment = ref<TabAlignment>('left');
const newTabLeader = ref('');

const canAddTab = computed(() => parsePositive(newTabCm.value) !== null);

const showError = computed(() => {
  for (const value of [leftCm.value, rightCm.value, amountCm.value, beforePt.value, afterPt.value, linePt.value]) {
    if (parseNonNegative(value) === null) return true;
  }
  return false;
});

const canSubmit = computed(() => !showError.value);

/** הכלל שנשלח למנוע, גם לפס וגם ב„אישור” — שני נתיבים לאותה הכרעה היו נפרדים. */
const lineRule = computed<LineSpacingRule>(() =>
  lineMode.value === 'exact' || lineMode.value === 'atLeast' ? lineMode.value : 'auto',
);

/** ה-`w:line` שנשלח: כפולה×240 ב-`auto`, ומרחק בנקודות בשני האחרים. */
const lineTwips = computed(() =>
  lineRule.value === 'auto'
    ? Number(lineMode.value)
    : ptToTwips(parseNonNegative(linePt.value) ?? 0),
);

/**
 * הפס מצויר מהערכים ש**בשדות** ולא מה-snapshot: הוא נועד להראות מה יקרה
 * ב„אישור”, וזה בדיוק ההבדל בין תצוגה מקדימה לבין חזרה על מה שכבר במסמך.
 *
 * ערך פסול בשדה (מינוס, טקסט) נקרא כאפס ולא מפיל את הפס: `showError` ממילא
 * נועל את „אישור” ואומר את הסיבה, ופס שנעלם באמצע הקלדה היה תזוזה שנייה
 * באותו דיאלוג על אותה שגיאה אחת.
 */
const preview = computed(() =>
  paragraphPreviewGeometry({
    textWidthTwips: props.pageTextWidthTwips,
    fontSizePt: props.fontSizePt,
    startTwips: cmToTwips(parseNonNegative(leftCm.value) ?? 0),
    endTwips: cmToTwips(parseNonNegative(rightCm.value) ?? 0),
    special: special.value,
    amountTwips: cmToTwips(parseNonNegative(amountCm.value) ?? 0),
    beforeTwips: ptToTwips(parseNonNegative(beforePt.value) ?? 0),
    afterTwips: ptToTwips(parseNonNegative(afterPt.value) ?? 0),
    lineTwips: lineTwips.value,
    lineRule: lineRule.value,
  }),
);

/**
 * הכיוון של הפס: של הפסקה כשהיא מצהירה, ושל המקטע כשהיא שותקת.
 *
 * `w:start`/`w:end`, שהכניסות נכתבות בהם, לוגיים **לפסקה** — פסקה LTR בתוך
 * מקטע RTL מקבלת „לפני טקסט” בצד השמאלי — ולכן ההצהרה של הפסקה גוברת.
 *
 * אבל היעדר הצהרה אינו LTR: מסמך עברי שנוצר ב-Word אינו מצהיר `<w:bidi>` על
 * כל פסקה אלא יורש, ופס שהיה קורא היעדר כ-LTR היה מצייר את הכניסה בצד ההפוך
 * על מסמכים רגילים לגמרי. שלושת המצבים אפשריים מפני שהמודל מבדיל ביניהם —
 * נמדד; ראו `bidi` ב-`ParagraphFormatSnapshot`.
 */
const previewDir = computed(() => ((props.snapshot.bidi ?? props.sectionRtl) ? 'rtl' : 'ltr'));

/** „עמודה 16.00 ס״מ · גופן 11 נק׳” — קנה המידה שהפס מצייר לפיו, במפורש. */
const scaleText = computed(() => {
  if (!preview.value) return '';
  const cm = (props.pageTextWidthTwips / TWIPS_PER_CM).toFixed(2);
  return `עמודה ${cm} ס"מ · גופן ${props.fontSizePt} נק'`;
});

/**
 * הפסקה שנערכת: הריווח לפניה ואחריה, ומרווח השורות כמשתנה שהפסים נמדדים בו.
 *
 * הכניסות **אינן** כאן אלא על הפסים עצמם, וזו הכרעה ולא סידור: אחוז ב-CSS
 * נמדד מול הרוחב הפנימי של האב, ולכן `margin-inline-start` על הפסקה היה
 * מקטין את האב של הפסים ומשנה את המכנה של כל אחוז שבתוכה. עם הכניסות על
 * הפסים כל שש ההגדרות נמדדות מול אותו מכנה אחד — רוחב עמודת הטקסט.
 */
const targetStyle = computed<CSSProperties>(() => {
  const geometry = preview.value;
  if (!geometry) return {};
  return {
    '--pv-font': `${geometry.fontCqw}cqw`,
    marginBlockStart: `${geometry.beforePct}%`,
    marginBlockEnd: `${geometry.afterPct}%`,
    '--pv-lh': String(geometry.lineHeight),
  };
});

/**
 * השכנות: קיימות רק כדי שיהיה למרחק „לפני”/„אחרי” גוף להיפרד ממנו.
 *
 * ברוחב מלא ובמרווח בודד, ובכוונה: לשכנות יש כניסות ומרווח **משל עצמן**,
 * שאינם נקראים כאן ואינם משתנים ב„אישור”. פס שהיה מצייר עליהן את ההגדרות
 * של הפסקה שנערכת היה אומר שהאישור נוגע גם בהן.
 */
const contextStyle = computed<CSSProperties>(() => {
  const geometry = preview.value;
  if (!geometry) return {};
  return { '--pv-font': `${geometry.fontCqw}cqw`, '--pv-lh': '1' };
});

/** שורות 2 ואילך: שני צדי הכניסה. */
const lineStyle = computed<CSSProperties>(() => {
  const geometry = preview.value;
  if (!geometry) return {};
  return {
    marginInlineStart: `${geometry.startPct}%`,
    marginInlineEnd: `${geometry.endPct}%`,
  };
});

/** השורה הראשונה: אותם צדדים, ועליהם „מיוחד” — חיובי לראשונה, שלילי לתלויה. */
const firstLineStyle = computed<CSSProperties>(() => {
  const geometry = preview.value;
  if (!geometry) return {};
  return {
    marginInlineStart: `${geometry.startPct + geometry.firstLinePct}%`,
    marginInlineEnd: `${geometry.endPct}%`,
  };
});

function onAddTab(): void {
  const cm = parsePositive(newTabCm.value);
  if (cm === null) return;
  emit('tab-add', {
    positionTwips: cmToTwips(cm),
    alignment: newTabAlignment.value,
    ...(newTabLeader.value !== '' ? { leader: newTabLeader.value as TabLeader } : {}),
  });
  // אופטימי: הרשימה מתעדכנת מיד; כשל של המנוע יגיע דרך ה-reporter של ההורה,
  // והתצלום בפתיחה הבאה יחזיר את האמת מהמסמך.
  tabs.value.push({
    positionTwips: cmToTwips(cm),
    alignment: newTabAlignment.value,
    ...(newTabLeader.value !== '' ? { leader: newTabLeader.value } : {}),
  });
  newTabCm.value = '';
}

function formatTab(tab: { positionTwips: number; alignment: string; leader?: string }): string {
  const ALIGNMENT_LABELS: Record<string, string> = {
    left: 'שמאל',
    center: 'מרכז',
    right: 'ימין',
    decimal: 'עשרוני',
    bar: 'קו אנכי',
  };
  const position = (tab.positionTwips / TWIPS_PER_CM).toFixed(2);
  return `${position} ס"מ · ${ALIGNMENT_LABELS[tab.alignment] ?? tab.alignment}${tab.leader ? ' · מוביל' : ''}`;
}

function onSubmit(): void {
  if (props.busy || !canSubmit.value) return;
  const amount = parseNonNegative(amountCm.value) ?? 0;
  emit('submit', {
    leftCm: parseNonNegative(leftCm.value) ?? 0,
    rightCm: parseNonNegative(rightCm.value) ?? 0,
    special: special.value,
    amountCm: amount,
    beforePt: parseNonNegative(beforePt.value) ?? 0,
    afterPt: parseNonNegative(afterPt.value) ?? 0,
    lineTwips: lineTwips.value,
    lineRule: lineRule.value,
    keepNext: keepNext.value,
    keepLines: keepLines.value,
    widowControl: widowControl.value,
    contextualSpacing: contextual.value === '' ? null : contextual.value === 'yes',
  });
}
</script>

<style scoped>
.para-dialog {
  position: fixed;
  top: 140px;
  inset-inline-start: 40px;
  z-index: 2000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
  /* שתי עמודות ופס תצוגה — אותו רוחב של „מתקדם” של הגופן, ומאותו טעם: ראו
     „הפריסה” בהערת הפתיחה. */
  width: 560px;
  max-width: calc(100vw - 32px);
  max-height: calc(100vh - 200px);
  /* הגוף גולל, לא הדיאלוג — אותה הכרעה, ומאותה סיבה, כמו ב-FontAdvancedDialog:
     `overflow-block: auto` על השורש דחק את הפוטר („אישור” / „ביטול”) אל מתחת
     לתקרה `calc(100vh - 200px)`, בלי דרך להגיע אליו. שם זה נמדד בשער. */
  display: flex;
  flex-direction: column;
  overflow: hidden;
  font-family: var(--font-main);
  user-select: none;
}

.pd-header {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-block-end: 1px solid var(--color-outline-variant);
  padding: 6px 8px 6px 12px;
  background: var(--color-surface-container-high);
  border-radius: var(--radius-sm) var(--radius-sm) 0 0;
}

.pd-title {
  font-size: 12px;
  font-weight: 600;
  color: var(--color-on-surface);
}

.pd-close-btn {
  background: none;
  border: none;
  font-size: 14px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  padding: 4px 6px;
  border-radius: var(--radius-xs);
}

.pd-close-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-error);
}

.pd-body {
  flex: 1 1 auto;
  min-height: 0;
  overflow-y: auto;
  padding: 10px 12px;
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.pd-columns {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 8px;
  align-items: start;
}

.pd-column {
  display: flex;
  flex-direction: column;
  gap: 10px;
  /* בלעדיו עמודת גריד אינה מתכווצת מתחת לרוחב תוכנה, ושורת הטאב הארוכה
     הייתה מרחיבה את הדיאלוג מעבר ל-560. */
  min-width: 0;
}

.pd-group {
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  padding: 6px 8px 8px;
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.pd-group legend {
  font-size: 11px;
  font-weight: 600;
  color: var(--color-on-surface-variant);
  padding-inline: 4px;
}

.pd-row {
  display: flex;
  align-items: center;
  gap: 6px;
}

/* הפקד התלת-מצבי הוא כפתור ברוחב מלא, ולא שדה בשורת תווית: התווית שלו היא
   משפט („אל תוסיף רווח בין פסקאות מאותו סגנון”), ושורת `flex` רגילה הייתה
   דוחסת אותו לרוחב של מספר. */
.pd-row-tri {
  margin-block-start: 2px;
}

.pd-row-tri > * {
  inline-size: 100%;
}

/* שורת הטאב היא ארבעה פקדים בעמודה של חצי דיאלוג — היא נשברת ולא גולשת. */
.pd-row-wrap {
  flex-wrap: wrap;
}

.pd-label {
  font-size: 11px;
  color: var(--color-on-surface);
  flex-shrink: 0;
}

.pd-label-after {
  margin-inline-start: auto;
}

.pd-unit {
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.pd-number,
.pd-select {
  padding: 3px 6px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 12px;
  outline: none;
  width: 72px;
}

.pd-select {
  width: auto;
}

.pd-number:focus,
.pd-select:focus {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

.pd-number:disabled {
  opacity: 0.45;
}

.pd-checks {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.pd-check {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: var(--color-on-surface);
}

.pd-tabs {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.pd-tab-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  font-size: 11px;
  color: var(--color-on-surface);
  padding-block: 2px;
}

.pd-note {
  margin: 0;
  font-size: 10.5px;
  line-height: 1.4;
  color: var(--color-on-surface-variant);
}

.pd-note-warn {
  color: var(--color-error);
}

/* ---- פס התצוגה המקדימה ---- */

.pd-preview {
  display: flex;
  flex-direction: column;
  gap: 3px;
}

.pd-preview-caption {
  font-size: 10px;
  color: var(--color-on-surface-variant);
}

.pd-preview-page {
  /*
   * `container-type: inline-size` הוא מה שהופך את הרוחב הזה למכנה של `cqw`,
   * וכך גודל הגופן בפס נשמר בקנה מידה אחד עם האחוזים — בלי מדידה של ה-DOM
   * ובלי `ResizeObserver`. ההנמקה המלאה ב-engine/paragraph-preview.ts.
   *
   * גובה קבוע ותוכן ממורכז, כמו בפס של „מתקדם”: הפס משנה מרווח שורות וריווח
   * תוך כדי הקלדה, וגובה שנגזר מהתוכן היה מזיז את הפוטר בכל הקשה — כלומר
   * „אישור” שבורח מתחת לעכבר. מה שאינו נכנס נחתך בקצוות, והפסקה שנערכת
   * נשארת במרכז.
   */
  container-type: inline-size;
  block-size: 116px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  padding-inline: 10px;
  overflow: hidden;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface-container-high);
}

.pd-pv-para {
  margin: 0;
  display: flex;
  flex-direction: column;
  /*
   * גודל הגופן דרך משתנה, ולא ישירות מה-`:style`, ובכוונה: `cqw` היא יחידה
   * ש-jsdom אינו מכיר, ו-`style.setProperty('font-size', '2.44cqw')` **נבלע
   * שם בשקט** — נמדד. כלומר בדיקת רכיב שהייתה נוקבת בגודל הגופן הייתה עוברת
   * ירוק גם על מספר שגוי לגמרי. משתנה מותאם כן נשמר ב-jsdom, ולכן המספר
   * נבדק שם, וההסבה שלו ל-`font-size` היא השורה הסטטית הזאת — שנמדדת בשער
   * הדפדפן. ראו tests/component/paragraph-dialog.test.ts.
   */
  font-size: var(--pv-font);
}

.pd-pv-target {
  color: var(--color-on-surface);
}

/* השכנות אינן מתחרות על תשומת הלב עם הפסקה שנערכת — הן רק ההקשר שלה. */
.pd-pv-context {
  color: var(--color-on-surface-variant);
  opacity: 0.45;
}

.pd-pv-line {
  /*
   * תיבת השורה, ובתוכה הפס. הגובה הוא `1em × --pv-lh`: ‏`1em` הוא גודל
   * הגופן שהפס צויר בו (ב-`cqw`), ו-`--pv-lh` הוא בדיוק ה-`line-height`
   * שהמסמך יקבל — כלומר מרווח השורות נראה כמרחק בין הפסים, ולא כעיבוי שלהם.
   */
  block-size: calc(1em * var(--pv-lh, 1));
  display: flex;
  align-items: center;
}

.pd-pv-line::before {
  content: '';
  display: block;
  inline-size: 100%;
  block-size: 0.42em;
  background: currentcolor;
  border-radius: 1px;
}

/* השורה האחרונה של פסקה אינה מגיעה לקצה — בלעדיה הפסים נראים כטבלה. */
.pd-pv-line-last::before {
  inline-size: 62%;
}

.pd-pv-context .pd-pv-line:last-child::before {
  inline-size: 74%;
}

.pd-error {
  margin: 0;
  font-size: 11px;
  color: var(--color-error);
}

.pd-footer {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  padding: 8px 12px;
  border-block-start: 1px solid var(--color-outline-variant);
  background: var(--color-surface-container-high);
  border-radius: 0 0 var(--radius-sm) var(--radius-sm);
}

.pd-btn {
  padding: 4px 10px;
  font-size: 11px;
  font-family: var(--font-main);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  cursor: pointer;
  transition: all 0.08s;
}

.pd-btn:hover:not(:disabled) {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

/* הטקסט על הכפתור הממולא הוא `--color-on-primary` ולא לבן קבוע — ראו
   LinkDialog.vue. */
.pd-btn-primary {
  background: var(--word-blue);
  color: var(--color-on-primary);
  border-color: var(--word-blue);
}

.pd-btn-primary:hover:not(:disabled) {
  background: var(--word-blue-dark);
}

.pd-btn:disabled {
  opacity: 0.4;
  cursor: default;
}
</style>

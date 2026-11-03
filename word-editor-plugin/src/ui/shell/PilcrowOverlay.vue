<template>
  <div
    ref="rootRef"
    class="pilcrow-layer"
    aria-hidden="true"
  >
    <span
      v-for="mark in marks"
      :key="mark.nodeId"
      class="pilcrow-layer__mark"
      :class="{ 'pilcrow-layer__mark--rtl': mark.direction === 'rtl' }"
      :style="{
        left: `${mark.anchorXPx}px`,
        top: `${mark.topPx}px`,
        height: `${mark.heightPx}px`,
        lineHeight: `${mark.heightPx}px`,
      }"
      >¶</span
    >
  </div>
</template>

<script setup lang="ts">
/**
 * שכבת „סימני עיצוב" — מציירת ¶ בסוף כל פסקה, כשהפקד „הצג/הסתר סימני
 * עיצוב" דלוק.
 *
 * ## למה השכבה הזאת קיימת בכלל
 *
 * `docs/superdoc-2.10-review.md` ("סימני עיצוב (¶/·/→) — נחקר לעומק") מתעד:
 * הפקודה `formatting-marks` מתהפכת כמו שצריך במנוע, אבל נתיב הציור המובנה
 * שלו מת מבנית ב-2.10.0 (`getDocumentRuntime()` מוקלד `() => null`, ואף
 * קורא ל-`setDocumentRuntime` לא קיים בכל ה-bundle). בדיוק כמו „גבולות
 * עמוד"/„מספרי שורות" — זו לא באג בקוד שלנו, והשכבה הזאת היא ציור **שלנו**,
 * מעל מה שהמנוע כבר מסמן כדלוק, ואינה נוגעת בשום דבר בשכבת הכתיבה/ייצוא.
 *
 * ## למה ¶ בלבד, ולא גם · (רווח) ו-→ (טאב)
 *
 * ראו `docs/superdoc-2.10-review.md`: טאבים אינם קיימים כתו בשום Text node,
 * וסמן רשימה מזריק טקסט שאינו קיים בטקסט הקנוני של `blocks.list()` — שני
 * חסמים שפוסלים מיקום מדויק. ¶ עוקף אותם כי הוא נוגע רק בסוף הפסקה. ההיקף
 * המצומצם הזה (וגם: רק `nodeType === 'paragraph'`, לא כותרות/פריטי-רשימה)
 * מתועד ונבדק — לא ניחוש חלקי. ראו engine/formatting-marks-layer.ts.
 *
 * ## שני מקורות
 *
 *   1. **מה לצייר** — `blocks` (`FormattingMarksBlock[]`, engine/formatting-marks.ts),
 *      נקרא מ-`doc.blocks.list({includeText:true})` — Document API ציבורי,
 *      לא selector. `null`/ריק כשאין Document API — אין מה לצייר.
 *   2. **איפה לצייר** — `runs` (`PageTextRun[]`, engine/page-ruler.ts,
 *      `watchAllPageTextRuns`), נמדד כאן פנימית מ-`host`/`viewport-source` —
 *      אותם props בדיוק כמו הסרגל וגבולות העמוד. העיגון (`data-page-index`)
 *      יושב אך ורק ב-page-ruler.ts, לא כאן.
 *
 * ההרכבה (התאמת-הרצף בין הבלוקים לריצות) היא engine/formatting-marks-layer.ts —
 * כאן רק ציור.
 *
 * ## `visible` עוצר את המדידה, לא רק את הציור
 *
 * בשונה מ-`PageBorderOverlay`/`LineNumberOverlay` (שתמיד עוקבים אחרי
 * הגיאומטריה, ורק `reading:null` מצייר ריק), כאן `visible:false` **מפרקת**
 * את המעקב אחרי ריצות הטקסט לגמרי: המדידה כאן (TreeWalker+Range על כל צומת,
 * ועוד `getComputedStyle` לכיוון) יקרה יותר מזו של שתי השכבות האחרות, וסימני
 * עיצוב כבויים הם המצב השכיח (ברירת המחדל של הפקד — כבוי). אין טעם למדוד
 * מה שאיש לא רואה.
 *
 * ## עדכון חי
 *
 * הצגה/הסתרה מגיעה מ-`adapter.observe('formatting-marks', …)` ב-App.vue —
 * בשונה מ„גבולות עמוד"/„מספרי שורות", הפקודה הזו **כן** מפעילה עדכון מצב
 * תקין (נמדד: `active` מתהפך כראוי), ולכן אין צורך ברענון מפורש על הפקודה
 * עצמה; App.vue קורא `formattingMarksModel.setEnabled(active)` על כל שינוי.
 * עריכת טקסט/reflow מגיעים דרך `noteDocumentChanged()` (App.vue, `onUpdate`)
 * וגם דרך `watchAllPageTextRuns` עצמו (גלילה, שינוי גודל, `viewport.observe`).
 */
import { computed, shallowRef, watch } from 'vue';
import {
  watchAllPageTextRuns,
  type PageTextRun,
  type PageRectWatch,
  type ViewportSource,
} from '../../engine/page-ruler';
import { buildPilcrowMarks, type FormattingMarksBlock, type PilcrowMark } from '../../engine/formatting-marks-layer';

const props = withDefaults(
  defineProps<{
    /** ה-host המצויר של המסמך הפתוח, מ-`paintedHost(ui)`. */
    host?: HTMLElement | null;
    /** ה-controller, בשביל `viewport.observe` — אותו prop כמו שתי השכבות האחרות. */
    viewportSource?: ViewportSource | null;
    /** בלוקי המסמך הנוכחיים, או `null` כשאין (מסמך סגור, או עדיין לא נקרא). */
    blocks?: readonly FormattingMarksBlock[] | null;
    /** מצב הפקד „הצג/הסתר סימני עיצוב". `false` = לא מוצג ולא נמדד. */
    visible?: boolean;
  }>(),
  {
    host: null,
    viewportSource: null,
    blocks: null,
    visible: false,
  },
);

const rootRef = shallowRef<HTMLElement | null>(null);
const runs = shallowRef<readonly PageTextRun[]>([]);

let watcher: PageRectWatch | null = null;

function stopWatching(): void {
  watcher?.dispose();
  watcher = null;
}

watch(
  [() => props.host, rootRef, () => props.visible],
  ([host, root, visible]) => {
    stopWatching();
    runs.value = [];
    if (!host || !root || !visible) return;
    watcher = watchAllPageTextRuns({
      host,
      reference: root,
      ui: props.viewportSource,
      onChange: (next) => {
        runs.value = next;
      },
    });
  },
  { immediate: true, flush: 'post' },
);

// בלוקים חדשים (פתיחת מסמך, עריכה שכבר עודכנה ב-Document API) — למדוד שוב
// את הגיאומטריה גם אם אף אחד מגלילה/שינוי-גודל/`viewport.observe` לא ירה:
// רשת ביטחון לאותה מלכודת שכבר נתפסה ב-DocumentRuler/PageBorderOverlay.
watch(() => props.blocks, () => watcher?.measure());

const marks = computed<PilcrowMark[]>(() => {
  if (!props.visible || !props.blocks) return [];
  return buildPilcrowMarks(props.blocks, runs.value);
});
</script>

<style scoped>
.pilcrow-layer {
  position: absolute;
  inset: 0;
  overflow: hidden;
  pointer-events: none;
  /* מעל תוכן המסמך שהמנוע מצייר, מתחת לכל דיאלוג/תפריט — אותו z-index כמו
     שתי השכבות האחרות. */
  z-index: 1;
}

.pilcrow-layer__mark {
  position: absolute;
  box-sizing: border-box;
  font-family: var(--font-main);
  color: var(--color-on-surface-variant);
  white-space: nowrap;
  user-select: none;
}

/* RTL: הקצה שנמדד (`anchorXPx`) הוא ה-`left` של התו האחרון, כלומר ה-¶ צריך
   להסתיים שם ולא להתחיל שם — הזזה מלאה של רוחב עצמו שמאלה. */
.pilcrow-layer__mark--rtl {
  transform: translateX(-100%);
}
</style>

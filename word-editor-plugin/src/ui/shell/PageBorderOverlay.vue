<template>
  <div
    ref="rootRef"
    class="page-border-layer"
    aria-hidden="true"
  >
    <div
      v-for="box in boxes"
      :key="box.pageIndex"
      class="page-border-layer__page"
      :style="{
        left: `${box.leftPx}px`,
        top: `${box.topPx}px`,
        width: `${box.widthPx}px`,
        height: `${box.heightPx}px`,
        borderTop: `${box.top.widthPx}px ${box.top.style} ${box.top.color}`,
        borderRight: `${box.right.widthPx}px ${box.right.style} ${box.right.color}`,
        borderBottom: `${box.bottom.widthPx}px ${box.bottom.style} ${box.bottom.color}`,
        borderLeft: `${box.left.widthPx}px ${box.left.style} ${box.left.color}`,
      }"
    />
  </div>
</template>

<script setup lang="ts">
/**
 * שכבת „גבולות עמוד” — מציירת מסגרת סביב כל עמוד מרונדר, לפי `<w:pgBorders>`
 * הנוכחי של המסמך.
 *
 * ## למה השכבה הזאת קיימת בכלל
 *
 * `@superdoc/docx-engine` כותב `<w:pgBorders>` קנוני ותקין ל-docx (נמדד —
 * docs/button-audit.md, docs/engine-gaps.md), אבל אינו מצייר אותו בעורך:
 * אפס אזכורים של `pgBorders` בשכבת ה-DOM החיה (נמדד ב-scripts/qa/layout-qa.mjs,
 * `anyBorderEl: 0`). זו לא באג בקוד שלנו — המנוע סגור-קוד ואינו מממש ציור
 * של page borders. השכבה הזאת היא ציור **שלנו**, מעל מה שהמנוע כבר מצייר,
 * ואינה נוגעת בשום דבר בשכבת הכתיבה/ייצוא של ה-docx.
 *
 * ## שני המקורות, ולמה כל אחד דווקא שם
 *
 *   1. **מה לצייר** — `reading` (`PageBordersReading`, engine/page-setup.ts),
 *      נקרא ישירות מ-`sections.list()` של המסמך ולא ממצב נפרד: ראה
 *      `createPageBorderModel` שם. `null` כשאין `<w:pgBorders>` — הסרת גבול
 *      מתאפסת כאן בדיוק כמו שהיא נעלמת מה-docx.
 *   2. **איפה לצייר** — `host`/`viewport-source`, אותם props בדיוק שהסרגל
 *      מקבל (App.vue מזין את שניהם מ-`paintedHost`/`rulerViewport`). המדידה
 *      עצמה — `watchAllPageRects` — יושבת ב-engine/page-ruler.ts, לא כאן:
 *      זה המקום היחיד שרשאי לגעת ב-`data-page-index`
 *      (tests/unit/engine-boundaries.test.ts).
 *
 * ## איך זה ממוקם על המסך
 *
 * שורש הרכיב הוא `position: absolute; inset: 0;` בתוך `.editor-area`
 * (`position: relative`, App.vue) — לא בתוך `.editor-stack` שהמנוע מצייר
 * לתוכו: כתיבה לתוך ה-container שהועבר ל-`createEditor` היא בדיוק העריכה
 * הישנה של ה-DOM שהתכנית באה להחליף, וגם מסוכנת בפני עצמה — `createEditor`
 * עשוי לנקות את תוכן ה-container שלו במעברי מסמך.
 *
 * המיקום של כל עמוד מחושב **ביחס לשורש הזה עצמו** (`reference: rootRef`),
 * בדיוק כמו ש-`DocumentRuler.vue`/`VerticalRuler.vue` עושים: `leftPx`/`topPx`
 * הם ההפרש האמיתי בפיקסלים בין מלבן העמוד למלבן השורש
 * (`getBoundingClientRect`), ולכן הם נכונים גם כש-`.editor-stack` גולל
 * מתחת לשורש הבלתי-נגלל, בלי שום חישוב CSS יחסי-לגלילה. `overflow: hidden`
 * על השורש חוסם עמוד שגלל חלקית מחוץ לתחום — בלי זה, עמוד שהתחלתו מעל
 * לאזור הנגלל היה מצייר גבול שמשתרע אל הרצועה שמעליו.
 *
 * ## עדכון חי
 *
 * שינוי הגדרות דרך תפריט „גבולות עמוד” מפעיל את אותה תחנת `onUpdate` שכל
 * מוטציה במסמך מפעילה (App.vue), וזו מריצה `noteDocumentChanged` על מודד
 * הגבולות — בדיוק כמו הסרגל. אין כאן האזנה נפרדת לפקודה הספציפית: כל שינוי
 * במסמך, מכל מקור, גורם לקריאה מחדש.
 */
import { computed, shallowRef, watch } from 'vue';
import {
  watchAllPageRects,
  type IndexedPageRect,
  type PageRectWatch,
  type ViewportSource,
} from '../../engine/page-ruler';
import { buildPageBorderBoxes } from '../../engine/page-border-layer';
import type { PageBordersReading } from '../../engine/page-setup';

const props = withDefaults(
  defineProps<{
    /** ה-host המצויר של המסמך הפתוח, מ-`paintedHost(ui)`. */
    host?: HTMLElement | null;
    /** ה-controller, בשביל `viewport.observe` — אותו prop כמו הסרגל. */
    viewportSource?: ViewportSource | null;
    /** מצב „גבולות עמוד” הנוכחי, או `null` כשאין גבול. */
    reading?: PageBordersReading | null;
  }>(),
  {
    host: null,
    viewportSource: null,
    reading: null,
  },
);

const rootRef = shallowRef<HTMLElement | null>(null);
const pageRects = shallowRef<readonly IndexedPageRect[]>([]);

let watcher: PageRectWatch | null = null;

function stopWatching(): void {
  watcher?.dispose();
  watcher = null;
}

/**
 * האם יש בכלל גבול. **התלוי הזה הוא מה שהופך את הכיבוי לחינמי:** בלעדיו
 * המעקב מדד את המלבן של **כל** עמוד על כל גלילה, כל שינוי גודל וכל
 * `viewport.observe` — גם במסמך שאין בו `<w:pgBorders>` כלל, כלומר ברוב
 * המסמכים. הציור ממילא היה ריק (`buildPageBorderBoxes` מחזיר `[]` על
 * `reading` שהוא `null`).
 *
 * `computed` על „האם קיים” ולא `() => props.reading`: שינוי בתוך הגבול עצמו
 * (צבע, עובי) אינו מפרק ובונה את המעקב מחדש — זו אותה גיאומטריה בדיוק.
 */
const bordered = computed(() => props.reading !== null);

watch(
  [() => props.host, rootRef, bordered],
  ([host, root, on]) => {
    stopWatching();
    pageRects.value = [];
    if (!host || !root || !on) return;
    watcher = watchAllPageRects({
      host,
      reference: root,
      ui: props.viewportSource,
      onChange: (rects) => {
        pageRects.value = rects;
      },
    });
    // מדידות ההתיישבות מיד עם ההתקנה: זו הרשת שהייתה קודם ב-`watch` על
    // `reading` למקרה שהגבול הראשון שנקרא הגיע בדיוק כשעימוד המסמך עדיין
    // באוויר (מסמך שנפתח עם `<w:pgBorders>` מ-Word, לפני שהעמוד הראשון סיים
    // להיצייר) — ומאז שההתקנה מותנית ב-`reading`, זה הרגע שבו זה מגיע לכאן.
    watcher.measure();
  },
  { immediate: true, flush: 'post' },
);

// שינוי **בתוך** גבול קיים (צבע, עובי, אילו צדדים) אינו מזיז את העמודים —
// `boxes` למטה כבר מגיב לו לבד — והמדידה כאן היא אותה רשת ביטחון לעימוד
// שעוד באוויר. הדלקה/כיבוי אינם עוברים כאן: הם ב-`watch` שמעל.
watch(() => props.reading, () => watcher?.measure());

const boxes = computed(() => buildPageBorderBoxes(pageRects.value, props.reading));
</script>

<style scoped>
.page-border-layer {
  position: absolute;
  inset: 0;
  overflow: hidden;
  pointer-events: none;
  /* מעל תוכן המסמך שהמנוע מצייר, מתחת לכל דיאלוג/תפריט (שיושבים מחוץ
     ל-.editor-area לגמרי). */
  z-index: 1;
}

.page-border-layer__page {
  position: absolute;
  box-sizing: border-box;
}
</style>

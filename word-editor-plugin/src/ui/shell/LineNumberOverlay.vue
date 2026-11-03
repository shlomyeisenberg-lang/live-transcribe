<template>
  <div
    ref="rootRef"
    class="line-number-layer"
    aria-hidden="true"
  >
    <span
      v-for="(box, index) in numberBoxes"
      :key="`${box.pageIndex}:${index}`"
      class="line-number-layer__num"
      :style="{
        left: `${gutterOf(box.pageIndex).leftPx}px`,
        width: `${gutterOf(box.pageIndex).widthPx}px`,
        top: `${box.topPx}px`,
        height: `${box.heightPx}px`,
        lineHeight: `${box.heightPx}px`,
        textAlign: textAlignSide,
      }"
      >{{ box.value }}</span
    >
  </div>
</template>

<script setup lang="ts">
/**
 * שכבת „מספרי שורות” — מציירת מספר בשולי כל שורת טקסט בגוף המסמך, לפי
 * `<w:lnNumType>` הנוכחי של המסמך.
 *
 * ## למה השכבה הזאת קיימת בכלל
 *
 * `@superdoc/docx-engine` כותב `<w:lnNumType>` קנוני ותקין ל-docx (נמדד —
 * docs/button-audit.md, docs/engine-gaps.md), אבל אינו מצייר אותו בעורך:
 * לא מספר שורה אחד על המסך בכל ארבעת הפריטים בתפריט „מספרי שורות”. בדיוק
 * כמו „גבולות עמוד” (ui/shell/PageBorderOverlay.vue) — זו לא באג בקוד
 * שלנו, והשכבה הזאת היא ציור **שלנו**, מעל מה שהמנוע כבר מצייר, בלי לגעת
 * בשכבת הכתיבה/ייצוא של ה-docx.
 *
 * ## שלושה מקורות
 *
 *   1. **מה לצייר** — `reading` (`LineNumberingReading`, engine/page-setup.ts),
 *      נקרא ישירות מ-`sections.list()`. ראה `createLineNumberingModel` שם.
 *      `null` כשאין `<w:lnNumType>` — כיבוי מתאפס כאן בדיוק כמו שהוא נעלם
 *      מה-docx.
 *   2. **איפה השורות** — מלבני הטקסט הגולמיים של כל עמוד
 *      (`watchAllPageContentRects`, engine/page-ruler.ts) — `Range.getClientRects()`
 *      על תוכן העמוד, לא selector אל מבנה פנימי של המנוע. ראו הערת הפתיחה
 *      של `measurePageContentRects` שם: זו בדיוק הסיבה שהטכניקה הזאת, ולא
 *      class name פנימי, היא היחידה שמותר לנו להישען עליה.
 *   3. **מלבן כל עמוד** — `watchAllPageRects`, אותו מקור בדיוק כמו גבולות
 *      עמוד — נחוץ כדי לדעת את פס השוליים (ימין/שמאל) לציור העמודה עצמה,
 *      ולתחום את פס „גוף הטקסט” (לא כותרת/שוליים) יחד עם `reading.page`.
 *
 * ההרכבה עצמה (סינון כותרת/שוליים, מיזוג ריצות-טקסט לשורה, מספור לפי
 * countBy/start/restart) היא engine/line-number-layer.ts — כאן רק ציור.
 *
 * ## מיקום ומיושר — אותו מבנה בדיוק כמו PageBorderOverlay.vue
 *
 * שורש הרכיב `position: absolute; inset: 0` בתוך `.editor-area`, וכל מספר
 * ממוקם ביחס לשורש הזה עצמו (`getBoundingClientRect`, לא CSS יחסי-לגלילה) —
 * ראו ההנמקה המלאה שם.
 *
 * ## צד השוליים — הכרעת עיצוב, לא נמדדת מול Word
 *
 * המנוע אינו מצייר `lnNumType` בכלל (ראו למעלה), ולכן אין ציור-ייחוס
 * לבדוק מולו. ההכרעה כאן: מסמך RTL (`reading.page.direction === 'rtl'`) —
 * טור המספרים בשוליים **הימניים**; LTR — בשוליים **השמאליים**, כמו ב-Word
 * הלועזי המתועד. זו מראה שכיוב-הגיון (אותה מיפוי כיוון שכבר נהוג בכל שאר
 * השכבות בתוסף — סרגל, כניסות פסקה) ולא אימות מול Word עצמו: אין לנו Word
 * להריץ מולו. אם זה יתברר כשגוי, התיקון מקומי לשורת ה-`textAlignSide`/
 * `gutterOf` למטה בלבד — אינו נוגע בכתיבת ה-docx (`applyLineNumbering`
 * כבר כותב `<w:lnNumType>` תקני בלי קשר לציור הזה).
 *
 * ## עדכון חי
 *
 * בדיוק כמו „גבולות עמוד”: שינוי דרך תפריט „מספרי שורות” הוא קריאת
 * section-level שאינה מפעילה `onUpdate` בעצמה (נמדד שם — ראו App.vue,
 * `reportCommand`), ולכן `lineNumberModel?.refreshNow()` נקרא במפורש משם
 * על הפקודה `page-line-numbering`. עריכת טקסט/reflow רגיל מגיע דרך
 * `onUpdate` הרגיל (`noteDocumentChanged`), וגם דרך `watchAllPageContentRects`/
 * `watchAllPageRects` עצמם (גלילה, שינוי גודל, `viewport.observe`).
 */
import { computed, shallowRef, watch } from 'vue';
import {
  watchAllPageRects,
  watchAllPageContentRects,
  type IndexedPageRect,
  type PageContentRects,
  type PageRectWatch,
  type ViewportSource,
} from '../../engine/page-ruler';
import { buildLineNumberBoxes, marginGutterPx, type GutterBand } from '../../engine/line-number-layer';
import type { LineNumberingReading } from '../../engine/page-setup';

const props = withDefaults(
  defineProps<{
    /** ה-host המצויר של המסמך הפתוח, מ-`paintedHost(ui)`. */
    host?: HTMLElement | null;
    /** ה-controller, בשביל `viewport.observe` — אותו prop כמו הסרגל וגבולות העמוד. */
    viewportSource?: ViewportSource | null;
    /** מצב „מספרי שורות” הנוכחי, או `null` כשהם כבויים. */
    reading?: LineNumberingReading | null;
  }>(),
  {
    host: null,
    viewportSource: null,
    reading: null,
  },
);

const rootRef = shallowRef<HTMLElement | null>(null);
const pageRects = shallowRef<readonly IndexedPageRect[]>([]);
const contentRects = shallowRef<readonly PageContentRects[]>([]);

let pageWatcher: PageRectWatch | null = null;
let contentWatcher: PageRectWatch | null = null;

function stopWatching(): void {
  pageWatcher?.dispose();
  pageWatcher = null;
  contentWatcher?.dispose();
  contentWatcher = null;
}

/**
 * האם יש בכלל מספור. **התלוי הזה הוא מה שהופך את הכיבוי לחינמי**, וזה נמדד:
 * המדידה של `watchAllPageContentRects` היא `Range.getClientRects()` על תוכן
 * כל עמוד, והיא רצה על כל גלילה, כל שינוי גודל וכל `viewport.observe` — גם
 * כשלמסמך אין `<w:lnNumType>` בכלל, כלומר ברוב המסמכים ורוב הזמן. בפרופיל
 * הקלדה נמדדו 275 קריאות `Range.getClientRects` ב-40 תווים על מסמך **בלי**
 * מספור שורות. הציור עצמו ממילא היה ריק (`buildLineNumberBoxes` מחזיר `[]`
 * על `reading` שהוא `null`), ולכן כל אותה גיאומטריה נמדדה בשביל כלום.
 *
 * `computed` על „האם קיים” ולא `() => props.reading`: שינוי בתוך המצב עצמו
 * (countBy, start) אינו מפרק ובונה את המעקב מחדש — זו אותה גיאומטריה בדיוק.
 *
 * אותו שיקול בדיוק כמו `visible` ב-PilcrowOverlay.vue, שכבר עוצר שם את
 * המדידה ולא רק את הציור.
 */
const numbering = computed(() => props.reading !== null);

watch(
  [() => props.host, rootRef, numbering],
  ([host, root, on]) => {
    stopWatching();
    pageRects.value = [];
    contentRects.value = [];
    if (!host || !root || !on) return;
    pageWatcher = watchAllPageRects({
      host,
      reference: root,
      ui: props.viewportSource,
      onChange: (rects) => {
        pageRects.value = rects;
      },
    });
    contentWatcher = watchAllPageContentRects({
      host,
      reference: root,
      ui: props.viewportSource,
      onChange: (rects) => {
        contentRects.value = rects;
      },
    });
    // מדידות ההתיישבות מיד עם ההתקנה, ולא רק המדידה היחידה שבבנייה: זו
    // הרשת שהייתה קודם ב-`watch` על `reading` — ומאז שההתקנה עצמה מותנית
    // ב-`reading`, זה הרגע שבו „המספור נדלק” מגיע לכאן. ראו SETTLE_DELAYS_MS.
    pageWatcher.measure();
    contentWatcher.measure();
  },
  { immediate: true, flush: 'post' },
);

// שינוי **בתוך** מצב קיים (countBy, start, restart) אינו מזיז עמודים, אבל כן
// עשוי להגיע בדיוק כשעימוד המסמך עוד באוויר — אותה רשת ביטחון כמו
// ב-PageBorderOverlay.vue. הדלקה/כיבוי אינם עוברים כאן: הם ב-`watch` שמעל.
watch(() => props.reading, () => {
  pageWatcher?.measure();
  contentWatcher?.measure();
});

const numberBoxes = computed(() => buildLineNumberBoxes(contentRects.value, pageRects.value, props.reading));

const textAlignSide = computed<'left' | 'right'>(() =>
  props.reading?.page.direction === 'rtl' ? 'left' : 'right',
);

const pageByIndex = computed(() => new Map(pageRects.value.map((p) => [p.pageIndex, p] as const)));

/** פס השוליים לעמוד נתון — ראו הערת הפתיחה, „צד השוליים”. */
function gutterOf(pageIndex: number): GutterBand {
  const reading = props.reading;
  const page = pageByIndex.value.get(pageIndex);
  if (!reading || !page) return { leftPx: 0, widthPx: 0 };
  const side = reading.page.direction === 'rtl' ? 'right' : 'left';
  return marginGutterPx(page, reading.page, side);
}
</script>

<style scoped>
.line-number-layer {
  position: absolute;
  inset: 0;
  overflow: hidden;
  pointer-events: none;
  /* מעל תוכן המסמך שהמנוע מצייר, מתחת לכל דיאלוג/תפריט — אותו z-index
     כמו PageBorderOverlay.vue: שתי השכבות אינן חופפות (גבול על קצה העמוד,
     מספרים בשוליים), אבל אין נזק גם אם ייצא שיחפפו. */
  z-index: 1;
}

.line-number-layer__num {
  position: absolute;
  box-sizing: border-box;
  font-family: var(--font-main);
  font-size: 10px;
  color: var(--color-on-surface-variant);
  white-space: nowrap;
  overflow: hidden;
  padding-inline: 2px;
  user-select: none;
}
</style>

/**
 * גיאומטריית „מספרי שורות” — הופכת מלבנים גולמיים (`RawTextRect`,
 * engine/page-ruler.ts) יחד עם מה שנקרא מהמסמך (`LineNumberingReading`,
 * engine/page-setup.ts) לרשימת (מיקום, מספר) לציור. אין כאן DOM ואין
 * selector אל המנוע — רק חשבון, בדיוק כמו engine/page-border-layer.ts.
 *
 * ## שלושה שלבים, ולמה כל אחד נפרד
 *
 * 1. **`groupLinesFromRects`** — `Range.getClientRects()` מחזיר הרבה יותר
 *    מלבנים ממספר השורות האמיתי: כמה ריצות-טקסט על אותה שורה (נמדד: עד
 *    2 בממוצע), ומלבן-קונטיינר אחד לכל פסקה/כותרת שלמה. השלב הזה מפריד
 *    „שורה” מ„קונטיינר” לפי גובה חריג ביחס לחציון, ומאחד ריצות על אותה
 *    שורה לפי `top` כמעט זהה. נמדד (page-ruler.ts, הערת הפתיחה של
 *    `measurePageContentRects`): 107 מלבנים גולמיים → 51 קבוצות, תואם
 *    אחד-לאחד למספר השורות האמיתי שנמדד בנפרד.
 * 2. **`filterBodyLines`** — מוציא כותרת עליונה/תחתונה מהספירה: Word אינו
 *    ממספר אותן, ואין להן class ייעודי שמותר לנו לחפש (ראו הערת הפתיחה של
 *    page-ruler.ts). הפתרון: פס הגובה שגוף הטקסט תופס בעמוד כבר ידוע —
 *    `readPageMargins` (engine/page-setup.ts) מחזיר את השוליים
 *    ה**אפקטיביים** (מה שכותרת עליונה כבר דוחקת), ואותו מקור בדיוק שסרגל
 *    המדידה כבר משתמש בו. שורה שמרכזה מחוץ לפס — כותרת/שוליים, לא גוף.
 * 3. **`buildLineNumberBoxes`** — ממספר את השורות שנשארו לפי `countBy`
 *    (מציג רק כפולות), `start` (המספר הראשון) ו-`restart`.
 *
 * ## `restart: 'newSection'` — קירוב מכוון, ומתועד
 *
 * אין ב-DOM שנמדד דרך למפות שורה למקטע שלה בלי selector אל המנוע (ראו
 * page-ruler.ts). לכן `'newSection'` מטופל כאן **כמו `'continuous'`** —
 * אינו מאפס בין עמודים. זה בדיוק נכון במסמך חד-מקטעי (הרוב המכריע: זה כל
 * מה שממשק „מעבר מקטע” של התוסף הזה יודע ליצור), ומחמיר-בזהירות במסמך
 * שהגיע מ-Word עם כמה מקטעים: המספור ימשיך לרוץ במקום לאפס, שנראה פחות
 * „שבור” מאיפוס במקום הלא-נכון. ה-`w:lnNumType` שנכתב ל-docx אינו מושפע —
 * Word יציג נכון בפתיחה; זה קירוב של שכבת התצוגה שלנו בלבד.
 *
 * ## פער שנחקר ולא נסגר: טקסט בתוך תא טבלה מקבל מספר שורה משלו
 *
 * **Word אינו ממספר טקסט בתוך טבלה בכלל — לא כאן.** נמדד (Chrome headless
 * על ה-dist הארוז, טבלה 1×2 עם תוכן בתא אחד, פסקה אחת לפניה ואחת אחריה,
 * „רציף” מודלק): 4 מספרים במקום 2. הסיבה שורשית שונה לגמרי מ„כותרת
 * עליונה” (סעיף 2 למעלה), ולכן לא נפתרת באותה טכניקה:
 *
 * `filterBodyLines` מוציא כותרת/שוליים כי הם יושבים ב**פס גובה קבוע**
 * שכבר ידוע (`readPageMargins`). טבלה אין לה פס כזה — היא באמצע גוף
 * הטקסט, בכל גובה שהמסמך קובע, ואין API ציבורי (Document API, לא DOM)
 * שמחזיר גיאומטריית מסך לבלוק: `doc.blocks.list()` מחזיר `nodeType`
 * ('tbl' לעומת 'paragraph') לפי **סדר מסמך**, בלי שום `top`/`rect`, וגם
 * אין דרך פומבית לשאול "מה המלבן על המסך של הבלוק הזה" בשום namespace
 * שנמדד (engine-gaps.md).
 *
 * **מה שנבדק ונדחה, ולמה:**
 *
 *   1. **גובה חריג (כמו קונטיינר של פסקה).** נכשל: לתא טבלה בודד-שורה כמעט
 *      אין ריפוד/מרווח פסקה, ולכן מלבן-הקונטיינר שלו (`19.7px`) קרוב
 *      מאוד לגובה השורה עצמה (`17px`) — לא כמו פסקה רגילה (`92px` מול
 *      `17px`, נמדד). סף יחס-גובה שהיה תופס את זה היה תופס גם שורות רגילות.
 *   2. **הכלה גיאומטרית** (מלבן A "מכיל" מלבן B → A קונטיינר). נכשל: שתי
 *      ריצות-טקסט על **אותה שורה** (בידי-RTL) יכולות לשאת גובה שונה במעט
 *      זה מזה (`18.39px` מול `17px`, נמדד — הבדל טבעי בין סקריפטים), וההכלה
 *      הייתה פוסלת ריצה אמיתית מאותה שורה בטעות.
 *   3. **שני מלבנים נפרדים (רווח אמיתי) על אותו `top`, לזיהוי "שתי עמודות
 *      טבלה".** זה כן נמדד כסימן אמיתי: תא ריק ותא עם טקסט, שניהם ב-`top`
 *      זהה בדיוק, ברווח של כ-295px ביניהם. אבל הסימן הזה **אינו ניתן
 *      להבחנה מפסקה רגילה עם עצירות טאב** (`Name:\tJohn`) — גם שם הטווח
 *      חוצה שני מקטעי טקסט נפרדים על אותה שורה עם רווח ריק ביניהם, וזו
 *      תבנית שכיחה במסמכים עבריים (רשימות, טבלאות ידניות עם טאבים). סף
 *      רוחב-רווח שהיה מבדיל ביניהם הוא ניחוש, לא מדידה — ונפילה כאן
 *      הייתה מדלגת שורה אמיתית שמגיעה מספרה.
 *
 * **הדרך היחידה שכן הייתה עובדת — ואסורה בכוונה.** `data-block-id`/
 * `data-source-node-id`/`data-layout-story` על ה-DOM כן היו נותנים מיפוי
 * מדויק בלוק↔מסך (נמדד: `data-layout-story="body"` לעומת `"header:rId7"`,
 * ו-`data-block-id` נושא את אינדקס הפסקה). אבל אלה בדיוק אותם attributes
 * שהמנוע עצמו מתעד כ„חוזה DOM פנימי... שינוי כאן הוא breaking change”
 * (`../layout-engine/dom-contract/src/data-attrs.ts`, מקובץ בתוך
 * `node_modules/superdoc/dist/chunks/create-super-doc-ui-*.es.js`) — אותה
 * משפחה בדיוק כמו class names כמו `.superdoc-line`, ו-
 * tests/unit/engine-boundaries.test.ts אוסר במפורש selector שנוגע בהם.
 * ההסתמכות עליהם הייתה פותרת את הבאג הזה ושוברת את הכלל שמנע מהתוסף
 * להיות תלוי במימוש הפנימי של המנוע — לא נבחרה.
 *
 * **המסקנה, ולא סתם "עוד לא הספקנו":** זו מגבלה אמיתית של הטכניקה
 * (`Range.getClientRects()` על עמוד שלם, בלי selector פנימי), לא חוסר
 * זמן. `buildLineNumberBoxes` **אינו** מסנן טקסט בתוך טבלה, ומספר אותו
 * כאילו היה גוף טקסט רגיל — נמדד ומתועד ב-`tests/unit/line-number-layer.test.ts`
 * (`'טקסט בתוך תא טבלה — פער ידוע, לא מסונן'`) וב-
 * `scripts/qa/line-number-overlay-qa.mjs` (שלב „טבלה”, מדווח `fail`
 * במפורש — לא מוסתר). מי שירצה לסגור את הפער יצטרך או API ציבורי חדש
 * מהמנוע (בלוק→מלבן מסך), או להכריע מחדש שהסתמכות על ה-dom-contract
 * הפנימי שווה את השבירות שהיא חושפת.
 */
import type { IndexedPageRect, RawTextRect } from './page-ruler';
import type { LineNumberingReading } from './page-setup';

/** שורה חזותית אחת, אחרי איחוד ריצות-טקסט וסינון קונטיינרים. */
export interface LineBox {
  topPx: number;
  heightPx: number;
  leftPx: number;
  rightPx: number;
}

/** קונטיינר של פסקה/כותרת שלמה גבוה מזה בפחות מפי הזה מהשורה החציונית. */
const HEIGHT_OUTLIER_RATIO = 1.6;

/** שתי ריצות-טקסט על אותה שורה נמדדות בהפרש `top` זניח, לא זהה לגמרי. */
const TOP_MERGE_TOLERANCE_PX = 0.75;

/**
 * מלבנים גולמיים → שורות חזותיות. ראו הערת הפתיחה, שלב 1.
 *
 * לא רגיש לסדר הקלט: ממיין לפי `top` לפני הקיבוץ, ומאחד רק לתוך הקבוצה
 * **האחרונה** שנפתחה — תקין כשהקלט ממוין, וזול (`O(n log n)`) במקום
 * להשוות מול כל הקבוצות שנפתחו עד כה.
 */
export function groupLinesFromRects(rects: readonly RawTextRect[]): LineBox[] {
  const positive = rects.filter((r) => r.widthPx > 0 && r.heightPx > 0);
  if (positive.length === 0) return [];

  const heights = positive.map((r) => r.heightPx).sort((a, b) => a - b);
  const median = heights[Math.floor(heights.length / 2)]!;
  const lineLevel = positive.filter((r) => r.heightPx <= median * HEIGHT_OUTLIER_RATIO);

  const sorted = [...lineLevel].sort((a, b) => a.topPx - b.topPx);
  const groups: LineBox[] = [];
  for (const r of sorted) {
    const last = groups[groups.length - 1];
    if (last && Math.abs(last.topPx - r.topPx) < TOP_MERGE_TOLERANCE_PX) {
      last.leftPx = Math.min(last.leftPx, r.leftPx);
      last.rightPx = Math.max(last.rightPx, r.leftPx + r.widthPx);
      last.heightPx = Math.max(last.heightPx, r.heightPx);
    } else {
      groups.push({ topPx: r.topPx, heightPx: r.heightPx, leftPx: r.leftPx, rightPx: r.leftPx + r.widthPx });
    }
  }
  return groups;
}

/** מה שנדרש מ-`LineNumberingReading.page` לחישוב הפס האנכי/האופקי — טוויפס בלבד. */
export interface PageGeometryTwips {
  pageWidthTwips: number;
  pageHeightTwips: number;
  leftTwips: number;
  rightTwips: number;
  effectiveTopTwips: number;
  effectiveBottomTwips: number;
}

/** הפס האנכי (בפיקסלים על המסך) שגוף הטקסט תופס בעמוד הזה — לא כותרת/שוליים. */
export function bodyBandPx(
  page: IndexedPageRect,
  geometry: PageGeometryTwips,
): { topPx: number; bottomPx: number } {
  if (!(geometry.pageHeightTwips > 0)) {
    return { topPx: page.topPx, bottomPx: page.topPx + page.heightPx };
  }
  const topFraction = geometry.effectiveTopTwips / geometry.pageHeightTwips;
  const bottomFraction = geometry.effectiveBottomTwips / geometry.pageHeightTwips;
  return {
    topPx: page.topPx + topFraction * page.heightPx,
    bottomPx: page.topPx + page.heightPx - bottomFraction * page.heightPx,
  };
}

/** שורות שמרכזן בתוך פס גוף הטקסט בלבד — ראו הערת הפתיחה, שלב 2. */
export function filterBodyLines(
  lines: readonly LineBox[],
  page: IndexedPageRect,
  geometry: PageGeometryTwips,
): LineBox[] {
  const { topPx, bottomPx } = bodyBandPx(page, geometry);
  return lines.filter((line) => {
    const center = line.topPx + line.heightPx / 2;
    return center >= topPx && center <= bottomPx;
  });
}

/** פס השוליים (בפיקסלים על המסך) שבו מצויר טור מספרי השורות. */
export interface GutterBand {
  leftPx: number;
  widthPx: number;
}

/**
 * פס השוליים בצד המבוקש — שמאל או ימין. `side` אינו נגזר מכיוון המסמך כאן;
 * הקורא (הרכיב) הוא זה שבוחר איזה צד לפי `direction` — ראו הערת הפתיחה של
 * ui/shell/LineNumberOverlay.vue.
 */
export function marginGutterPx(
  page: IndexedPageRect,
  geometry: Pick<PageGeometryTwips, 'pageWidthTwips' | 'leftTwips' | 'rightTwips'>,
  side: 'left' | 'right',
): GutterBand {
  if (!(geometry.pageWidthTwips > 0)) return { leftPx: page.leftPx, widthPx: 0 };
  const marginTwips = side === 'left' ? geometry.leftTwips : geometry.rightTwips;
  const widthPx = (marginTwips / geometry.pageWidthTwips) * page.widthPx;
  const leftPx = side === 'left' ? page.leftPx : page.leftPx + page.widthPx - widthPx;
  return { leftPx, widthPx };
}

/** מלבני הטקסט הגולמיים של עמוד אחד — הקלט הגולמי ל-`buildLineNumberBoxes`. */
export interface PageLineSource {
  pageIndex: number;
  rects: readonly RawTextRect[];
}

/** מיקום ומספר לציור עבור שורה אחת שקיבלה תווית. */
export interface LineNumberBox {
  pageIndex: number;
  topPx: number;
  heightPx: number;
  value: number;
}

/**
 * ההרכבה השלמה: מלבנים גולמיים של כל עמוד + מלבני העמודים עצמם + מה
 * שנקרא מהמסמך → רשימת התוויות לציור, בכל העמודים יחד וברצף מסמך אחד
 * (המונה ממשיך מעמוד לעמוד אלא אם `restart: 'newPage'`).
 *
 * `null`/ריק בכל קלט → בלי תוויות, לא שגיאה: בדיוק כמו `buildPageBorderBoxes`.
 */
export function buildLineNumberBoxes(
  sources: readonly PageLineSource[],
  pages: readonly IndexedPageRect[],
  reading: LineNumberingReading | null,
): LineNumberBox[] {
  if (!reading || sources.length === 0 || pages.length === 0) return [];

  const pageByIndex = new Map(pages.map((p) => [p.pageIndex, p] as const));
  const ordered = [...sources].sort((a, b) => a.pageIndex - b.pageIndex);

  const out: LineNumberBox[] = [];
  let counter = reading.start;

  for (const { pageIndex, rects } of ordered) {
    const page = pageByIndex.get(pageIndex);
    if (!page) continue;

    // `'newSection'` אינו כאן בכוונה — ראו הערת הפתיחה. רק `'newPage'` מאפס.
    if (reading.restart === 'newPage') counter = reading.start;

    const lines = filterBodyLines(groupLinesFromRects(rects), page, reading.page);
    lines.sort((a, b) => a.topPx - b.topPx);

    for (const line of lines) {
      const isDisplayed = (counter - reading.start) % reading.countBy === 0;
      if (isDisplayed) {
        out.push({ pageIndex, topPx: line.topPx, heightPx: line.heightPx, value: counter });
      }
      counter += 1;
    }
  }

  return out;
}

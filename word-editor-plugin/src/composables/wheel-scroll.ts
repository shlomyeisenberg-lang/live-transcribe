/**
 * גלילה אופקית בגלגלת העכבר.
 *
 * ## מה זה פותר
 *
 * `.word-ribbon-body` — הפס שמחזיק את כל קבוצות הלשונית — הוא `overflow-x:
 * auto`, ובחלון צר הוא גולש ומצייר פס גלילה. נמדד בכרום (חלון 900px, לשונית
 * „בית”): `scrollWidth` 1227 מול `clientWidth` 900, כלומר 327px שאין דרך
 * להגיע אליהם — הגלגלת מעליו לא הזיזה את `scrollLeft` מאפס, לא מטה ולא מעלה.
 * כרום **אינו** מתרגם גלגלת אנכית לגלילה אופקית בעצמו, ולכן הדרך היחידה לפקד
 * שנשאר מחוץ למסך הייתה גרירת הפס עצמו.
 *
 * ## הכיוון: `rtl` הופך את הסימן, ולא את הכוונה
 *
 * המעטפת כולה `dir="rtl"`, ובמכולה כזאת `scrollLeft` מתחיל ב-0 בקצה הימני
 * ויורד לשלילי — נמדד טווח `[-327, 0]`. „גלגלת מטה” פירושו תמיד להתקדם אל
 * המשך התוכן, וב-RTL ההמשך נמצא **שמאלה**, כלומר delta שלילי. delta שלילי הוא
 * „שמאלה על המסך” בשני מודלי ה-`scrollLeft` שדפדפנים חיים מיישמים, כמו
 * ב-`galleryScrollDelta` (src/engine/style-gallery.ts) — ולכן החשבון כאן נשען
 * עליו ולא על סימן ה-`scrollLeft` שהמכולה מדווחת.
 *
 * ## הקצה מחזיר את האירוע, ולא בולע אותו
 *
 * כשאין לאן לגלול בכיוון המבוקש הפונקציה מחזירה `null` ולא מבטלת את האירוע.
 * זה מה שמאפשר קינון: גלגלת מעל גלריית הסגנונות גוללת את הגלריה, וכשהגלריה
 * מגיעה לסופה אותו גלגול ממשיך לבעבע ומזיז את הרצועה כולה — בדיוק כמו
 * שרשרת הגלילה הרגילה של הדפדפן. מי שכן טיפל מסמן זאת ב-`preventDefault`,
 * וההורה מדלג על אירוע שכבר טופל.
 *
 * ## מה לא מטופל כאן, בכוונה
 *
 * `deltaX` (מחווה אופקית במשטח מגע, או Shift+גלגלת שכרום ממיר לציר X) עובר
 * הלאה לדפדפן: את הציר האופקי הוא כן גולל בעצמו. `Ctrl+גלגלת` הוא זום, וגם
 * הוא אינו נלקח. `deltaMode` אינו נבדק מפני שהמארח הוא Chromium בלבד
 * (WebView2 באוצריא, כרום בפיתוח) והוא מדווח פיקסלים תמיד.
 */

/** סבילות במידות תת-פיקסליות. בלעדיה „יש לאן לגלול” נכון על 0.4 פיקסל. */
const EDGE_TOLERANCE_PX = 1;

/** המידות שמכולת הגלילה מדווחת. */
export interface WheelScrollMetrics {
  scrollLeft: number;
  scrollWidth: number;
  clientWidth: number;
}

/**
 * מה שנקרא מאירוע הגלגלת. `WheelEvent` אמיתי מקיים את זה מאליו, ובדיקה יכולה
 * למסור אובייקט פשוט.
 */
export interface WheelGesture {
  deltaX: number;
  deltaY: number;
  ctrlKey: boolean;
}

/**
 * כמה להוסיף ל-`scrollLeft`, או `null` כשאין להתערב.
 *
 * `null` בארבעה מצבים, וכולם „תן לדפדפן”: זום, מחווה שהציר האופקי שלה שולט,
 * מכולה שאינה גולשת כלל, וקצה שאין ממנו המשך בכיוון המבוקש.
 */
export function wheelScrollDelta(
  gesture: WheelGesture,
  metrics: WheelScrollMetrics,
  rtl: boolean,
): number | null {
  if (gesture.ctrlKey) return null;

  const { deltaX, deltaY } = gesture;
  if (!Number.isFinite(deltaY) || deltaY === 0) return null;
  // מחווה אופקית אמיתית: הדפדפן כבר יודע לגלול אותה, ואין מה לתרגם.
  if (Number.isFinite(deltaX) && Math.abs(deltaX) >= Math.abs(deltaY)) return null;

  const overflow = metrics.scrollWidth - metrics.clientWidth;
  if (!Number.isFinite(overflow) || overflow <= EDGE_TOLERANCE_PX) return null;

  // המרחק מתחילת הרשימה כערך מוחלט: ב-RTL הטווח הוא `[-overflow, 0]` וב-LTR
  // הוא `[0, overflow]`, ואותו חשבון מכסה את שניהם בלי ענף.
  const fromStart = Math.abs(metrics.scrollLeft);
  const toward = deltaY > 0 ? 'end' : 'start';
  if (toward === 'end' && fromStart >= overflow - EDGE_TOLERANCE_PX) return null;
  if (toward === 'start' && fromStart <= EDGE_TOLERANCE_PX) return null;

  return deltaY * (rtl ? -1 : 1);
}

/**
 * כיווניות המכולה. ברירת המחדל היא `rtl` כמו המעטפת, ורק `ltr` מפורש מבטל
 * אותה: ב-jsdom ובמשטחים שלא פתרו כיווניות הערך יוצא ריק.
 */
export function isRtlContainer(element: Element): boolean {
  return globalThis.getComputedStyle?.(element)?.direction !== 'ltr';
}

/**
 * המטפל עצמו: `@wheel="handleWheelScroll"` על מכולת הגלילה, בלי ref.
 *
 * `currentTarget` ולא `target`: האירוע מגיע מהפקד שהעכבר עמד עליו — כפתור,
 * תווית, אייקון — והמכולה היא זו שרשומה למאזין.
 *
 * מחזירה `true` אם גללה בפועל.
 */
export function handleWheelScroll(event: WheelEvent): boolean {
  // מישהו פנימי כבר טיפל: גלריית הסגנונות יושבת בתוך הרצועה, ואירוע שהיא
  // לקחה אינו אמור להזיז גם את הרצועה מתחתיה.
  if (event.defaultPrevented) return false;

  const element = event.currentTarget as HTMLElement | null;
  if (!element) return false;

  const delta = wheelScrollDelta(
    event,
    {
      scrollLeft: element.scrollLeft,
      scrollWidth: element.scrollWidth,
      clientWidth: element.clientWidth,
    },
    isRtlContainer(element),
  );
  if (delta === null) return false;

  // „זזתי” נמדד ואינו מנובא, וזה תיקון של באג שנמדד: גלריית הסגנונות היא
  // `scroll-snap-type: inline mandatory` עם 2px ריפוד, ולכן היא **נחה** על
  // `scrollLeft = -2` ולא על 0. החשבון שמעל ראה שם 2px של מרחק מההתחלה,
  // לקח את האירוע, ביקש לחזור — וההצמדה החזירה אותה מיד לאותם -2. התוצאה
  // הייתה גלגלת מתה: לא הגלריה זזה, ולא הרצועה שמתחתיה קיבלה את האירוע.
  //
  // ההצמדה מיידית, וזה מה שמאפשר את המדידה כאן: נמדד בכרום ש-`scrollLeft = 0`
  // על הגלריה מוחזר `-2` כבר בקריאה הבאה, ו-`-50` מוחזר `-73` (גבול כרטיס).
  const before = element.scrollLeft;
  element.scrollLeft += delta;
  if (element.scrollLeft === before) return false;

  // בלי זה כרום ממשיך לשרשרת הגלילה שמעל — ובחלון שגולל, הדף כולו היה זז
  // יחד עם הרצועה.
  event.preventDefault();
  return true;
}

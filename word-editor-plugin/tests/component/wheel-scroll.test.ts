/**
 * הגלגלת כפי שהיא מחווטת בפועל על הפסים הנגללים.
 *
 * מה שנמדד כאן אינו החשבון (הוא נבדק ב-tests/unit/wheel-scroll.test.ts, על
 * המידות שנמדדו בכרום) אלא **החוט**: שהמאזין יושב על מכולת הגלילה עצמה, שהוא
 * תופס גם גלגול שהתחיל מעל פקד בתוכה, ושקינון גלריה-בתוך-רצועה מתנהג כמו
 * שרשרת גלילה — הפנימי ראשון, וכשנגמר לו, החיצוני.
 *
 * ל-jsdom אין פריסה: כל `scrollWidth` ו-`clientWidth` הם אפס, ולכן כל מכולה
 * כאן „נכנסת כולה” והמאזין לא היה עושה דבר. המידות נמסרות לכן במפורש — הן
 * אותן מידות שנמדדו בכרום — ו-`scrollLeft` הוא הדבר היחיד ש-jsdom כן שומר.
 *
 * הסריקה בסוף הקובץ היא השער האמיתי: הפס הבא שייכתב עם `overflow-x: auto`
 * ייראה תקין לחלוטין בלי המאזין, פשוט לא יגיב לגלגלת.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import Ribbon from '../../src/ui/ribbon/Ribbon.vue';
import DocumentTabsBar from '../../src/ui/shell/DocumentTabsBar.vue';
import { autoUnmount, mountUi, settle } from './harness';

autoUnmount();

/** הרצועה כפי שנמדדה בכרום בחלון 900px, לשונית „בית”. */
const RIBBON_SIZE = { scrollWidth: 1227, clientWidth: 900 };
/** גלריית הסגנונות באותה מדידה. */
const GALLERY_SIZE = { scrollWidth: 498, clientWidth: 356 };

function find(selector: string): HTMLElement {
  const element = document.querySelector<HTMLElement>(selector);
  if (!element) throw new Error(`אין ${selector} בגוף הדף`);
  return element;
}

/** נותן למכולה מידות אמיתיות, ומחזיר אותה לתחילת הגלילה. */
function size(element: HTMLElement, box: { scrollWidth: number; clientWidth: number }): HTMLElement {
  Object.defineProperty(element, 'scrollWidth', { value: box.scrollWidth, configurable: true });
  Object.defineProperty(element, 'clientWidth', { value: box.clientWidth, configurable: true });
  element.scrollLeft = 0;
  return element;
}

/**
 * מכולה שההצמדה מחזירה למקומה — כמו גלריית הסגנונות, שהיא
 * `scroll-snap-type: inline mandatory` עם 2px ריפוד ולכן **נחה** על
 * `scrollLeft = -2`. נמדד בכרום שההצמדה מיידית: `scrollLeft = 0` מוחזר `-2`
 * כבר בקריאה הבאה, ו-`-50` מוחזר `-73` (גבול כרטיס, 68+3).
 */
function snapAt(element: HTMLElement, resting: number, overflow: number, step = 71): void {
  let value = resting;
  Object.defineProperty(element, 'scrollLeft', {
    configurable: true,
    get: () => value,
    set: (next: number) => {
      // המכולה `rtl`, ולכן הטווח הוא `[-overflow, 0]`. ההצמדה פועלת על מה
      // שנשאר אחרי החיתוך לטווח — וזה הצירוף שהרג את האירוע בכרום: בקשה
      // לחזור אחורה נחתכה ל-0, וההצמדה החזירה אותה משם ל--2.
      const clamped = Math.min(0, Math.max(-overflow, next));
      value = Math.abs(clamped - resting) < step ? resting : clamped;
    },
  });
}

/**
 * גלגול אמיתי, שמתחיל מהפקד שהעכבר עומד עליו ולא מהמכולה.
 *
 * זה ההבדל היחיד בין מאזין על המכולה למאזין על כל כפתור בנפרד, והוא מה
 * שהבדיקה הזאת אמורה להבחין בו. מוחזר האירוע — `defaultPrevented` שלו הוא
 * התשובה לשאלה „מי לקח אותו”.
 */
function wheel(from: Element, deltaY: number, init: Partial<WheelEventInit> = {}): WheelEvent {
  const event = new WheelEvent('wheel', {
    deltaY,
    deltaX: 0,
    bubbles: true,
    cancelable: true,
    ...init,
  });
  from.dispatchEvent(event);
  return event;
}

/**
 * ההרכבה הראשונה של הרצועה בקובץ יקרה: `mountUi` בונה את רשימת הגופנים
 * (`fallbackFontOptions`) לפני שהוא מגיע לתבנית, ונמדדו כאן 6.4 שניות מול
 * תקרת ברירת המחדל של 5. התקרה מורמת על הבדיקות שמרכיבות את הרצועה בלבד,
 * כמו ב-tests/component/ribbon-tabs.test.ts, ולא גלובלית.
 */
const MOUNT_TIMEOUT = 20_000;

describe('גלילת הרצועה בגלגלת', () => {
  it('גלגלת מעל פקד ברצועה גוללת את הרצועה', { timeout: MOUNT_TIMEOUT }, async () => {
    mountUi(Ribbon, { props: { hasDocument: true } });
    await settle();

    const body = size(find('.word-ribbon-body'), RIBBON_SIZE);
    const control = body.querySelector('button');
    if (!control) throw new Error('אין פקד בלשונית');

    const event = wheel(control, 100);

    // המעטפת `rtl`, ולכן ההמשך של הרצועה נמצא שמאלה — כלומר `scrollLeft` שלילי.
    expect(body.scrollLeft).toBe(-100);
    expect(event.defaultPrevented).toBe(true);
  });

  it('בקצה הרצועה האירוע נשאר של הדפדפן', { timeout: MOUNT_TIMEOUT }, async () => {
    mountUi(Ribbon, { props: { hasDocument: true } });
    await settle();

    const body = size(find('.word-ribbon-body'), RIBBON_SIZE);
    const event = wheel(body, -100);

    expect(body.scrollLeft).toBe(0);
    expect(event.defaultPrevented).toBe(false);
  });

  it('Ctrl+גלגלת אינו נלקח מהזום', { timeout: MOUNT_TIMEOUT }, async () => {
    mountUi(Ribbon, { props: { hasDocument: true } });
    await settle();

    const body = size(find('.word-ribbon-body'), RIBBON_SIZE);
    const event = wheel(body, 100, { ctrlKey: true });

    expect(body.scrollLeft).toBe(0);
    expect(event.defaultPrevented).toBe(false);
  });

  it('לשוניות הרצועה נגללות אף הן', { timeout: MOUNT_TIMEOUT }, async () => {
    mountUi(Ribbon, { props: { hasDocument: true } });
    await settle();

    const strip = size(find('.word-tab-strip'), { scrollWidth: 700, clientWidth: 400 });
    wheel(find('.word-tab-btn'), 100);

    expect(strip.scrollLeft).toBe(-100);
  });

  it('טאבי המסמכים נגללים אף הם', async () => {
    mountUi(DocumentTabsBar, {
      props: {
        tabs: [
          { id: 'a', title: 'מסמך א', isDirty: false },
          { id: 'b', title: 'מסמך ב', isDirty: true },
        ],
        activeId: 'a',
      },
    });
    await settle();

    const strip = size(find('.word-doctabs-strip'), { scrollWidth: 900, clientWidth: 300 });
    wheel(find('.word-doctab'), 100);

    expect(strip.scrollLeft).toBe(-100);
  });

  /*
   * הקינון הוא מה שמצדיק את `defaultPrevented` בראש המטפל ואת `null` בקצה:
   * גלריית הסגנונות יושבת **בתוך** הרצועה, ובלי שניהם גלגול מעליה היה מזיז
   * את שתיהן בבת אחת.
   */
  it('גלגלת מעל גלריית הסגנונות גוללת את הגלריה בלבד', { timeout: MOUNT_TIMEOUT }, async () => {
    mountUi(Ribbon, { props: { hasDocument: true } });
    await settle();

    const body = size(find('.word-ribbon-body'), RIBBON_SIZE);
    const gallery = size(find('.style-cards-scroll'), GALLERY_SIZE);

    wheel(gallery, 100);

    expect(gallery.scrollLeft).toBe(-100);
    expect(body.scrollLeft).toBe(0);
  });

  it('כשנגמרה הגלריה, אותו גלגול ממשיך לרצועה', { timeout: MOUNT_TIMEOUT }, async () => {
    mountUi(Ribbon, { props: { hasDocument: true } });
    await settle();

    const body = size(find('.word-ribbon-body'), RIBBON_SIZE);
    const gallery = size(find('.style-cards-scroll'), GALLERY_SIZE);
    // הגלריה בסופה: `-(498 - 356)`.
    gallery.scrollLeft = -142;

    wheel(gallery, 100);

    expect(gallery.scrollLeft).toBe(-142);
    expect(body.scrollLeft).toBe(-100);
  });

  /*
   * מה שנמדד בכרום ולא ב-jsdom, ולכן נכתב כאן רק אחרי שהשער תפס אותו
   * (scripts/qa/ribbon-wheel-qa.mjs): הגלריה נחה על -2 בגלל ההצמדה, החשבון
   * ראה שם 2px של מרחק מההתחלה, והאירוע נבלע בלי שאיש זז.
   */
  it('גלריה שההצמדה מחזירה למקומה אינה בולעת את הגלגול', { timeout: MOUNT_TIMEOUT }, async () => {
    mountUi(Ribbon, { props: { hasDocument: true } });
    await settle();

    const body = size(find('.word-ribbon-body'), RIBBON_SIZE);
    const gallery = size(find('.style-cards-scroll'), GALLERY_SIZE);
    snapAt(gallery, -2, GALLERY_SIZE.scrollWidth - GALLERY_SIZE.clientWidth);
    body.scrollLeft = -327;

    const event = wheel(gallery, -100);

    expect(gallery.scrollLeft).toBe(-2);
    expect(body.scrollLeft).toBe(-227);
    expect(event.defaultPrevented).toBe(true);
  });
});

/* ------------------------------------------------------------------ */
/* השער: כל פס נגלל מגיב לגלגלת                                        */
/* ------------------------------------------------------------------ */

const SRC = join(__dirname, '..', '..', 'src');

/** כל קובץ `.vue` ו-`.css` תחת src, רקורסיבית. */
function sources(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) return sources(path);
    return /\.(vue|css)$/.test(entry.name) ? [path] : [];
  });
}

/**
 * מחלקות שהוכרזו `overflow-x: auto` (או `scroll`) — כלומר פסים שנגללים
 * אופקית. הכלל אינו קורא סלקטורים מורכבים בכוונה: מחלקה בודדת היא הצורה
 * שכל ארבעת הפסים כתובים בה, וסלקטור אחר ייפול כאן כ„לא נמצא” ויידרש טיפול.
 */
function horizontalScrollers(): Map<string, string> {
  const found = new Map<string, string>();
  for (const file of sources(SRC)) {
    const text = readFileSync(file, 'utf8');
    const rules = text.matchAll(/\.([\w-]+)\s*(?:,[^{]*)?\{([^}]*)\}/g);
    for (const rule of rules) {
      if (/overflow-x:\s*(auto|scroll)/.test(rule[2])) found.set(rule[1], file);
    }
  }
  return found;
}

describe('כל פס נגלל מגיב לגלגלת', () => {
  it('לכל מכולה עם `overflow-x: auto` יש `@wheel`', () => {
    const scrollers = horizontalScrollers();
    // אם המפה ריקה, הביטוי לא תפס דבר — והשער היה עובר בירוק על אפס בדיקות.
    expect(scrollers.size).toBeGreaterThanOrEqual(4);

    const templates = sources(SRC).filter((file) => file.endsWith('.vue'));
    const missing: string[] = [];

    for (const [cls] of scrollers) {
      // התג שנושא את המחלקה, עד סוגר הפתיחה שלו: `@wheel` חייב לשבת בו ולא
      // בשכן. `class="x"` בלבד — כל ארבעת הפסים כתובים כך.
      const tag = new RegExp(`<[a-zA-Z][^>]*\\bclass="${cls}"[^>]*>`);
      const declared = templates.some((file) => {
        const match = tag.exec(readFileSync(file, 'utf8'));
        return match ? match[0].includes('@wheel') : false;
      });
      if (!declared) missing.push(cls);
    }

    expect(missing, `פסים נגללים בלי מאזין גלגלת: ${missing.join(', ')}`).toEqual([]);
  });
});

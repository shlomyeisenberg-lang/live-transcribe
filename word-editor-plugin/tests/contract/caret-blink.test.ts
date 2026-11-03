/**
 * חוזה סמן כתיבה מהבהב מול המנוע וגיליונות העיצוב.
 *
 * ## מה הבעיה שהחוזה מאמת
 *
 * סמן העריכה של SuperDoc מרונדר כ-`div` הנושא את המחלקה `sd-v2-local-selection-caret`.
 * באריזת המנוע קיימת הגדרת הבהוב (`sd-v2-local-caret-blink`), אולם המנוע מכבה אותה
 * תחת `@media (prefers-reduced-motion: reduce)`. בסביבות Windows רבות (ובפרט במחשבים
 * שבהם אפקטי הנפשה כבויים בהגדרות הנגישות או ביצועי המערכת), הגדרה זו מופעלת כברירת מחדל,
 * והסמן במסמך נשאר קפוא ולא מהבהב.
 *
 * ב-Word (ובכל מעבד תמלילים אמיתי), סמן העריכה תמיד מהבהב כאינדיקטור מיקום והקלדה חיוני.
 * הכלל ב-styles/shell.css כופה את האנימציה עם `!important` ומגדיר את ה-keyframes,
 * ו-styles/print.css מסתיר את הסמן בהדפסה.
 *
 * ## והצד השני: כשאין מיקוד אין סמן
 *
 * ‏„תמיד מהבהב” הוא כל עוד ההקלדה באמת נכנסת למסמך. כשהחלון עבר לחלון אחר, או
 * כשפקד של הממשק לקח את המיקוד, סמן מהבהב מבטיח דבר שאינו נכון — וב-Word הוא
 * נעלם. ההכרעה ב-`ui/shell/caret-focus.ts`, ההחלה היא `[data-caret-idle]`
 * שהמעטפת נושאת, והכלל שמכבה יושב באותו קובץ CSS.
 *
 * ## מה נמדד כאן, ומה **לא**
 *
 * הגיליון נטען לתוך המסמך של הבדיקה והתוצאה נקראת מ-`getComputedStyle` על
 * אלמנט סמן אמיתי — כלומר מה שנמדד הוא **הכרעת המפל**, ולא הטקסט של הכלל.
 * ההבדל אינו אקדמי: בדיקה שרק מוודאת שהבורר המכבה „נראה כמו שצריך” נשארת
 * ירוקה גם אחרי שהכלל המדליק מקבל סגוליות גבוהה ממנו, כלומר בדיוק במצב שבו
 * הסמן ממשיך להבהב בלי מיקוד.
 *
 * מה שאינו יכול להימדד כאן: שהתכונה באמת מודבקת למעטפת בזמן אמת, ושהסמן
 * שהמנוע מצייר באמת נמצא מתחת ל-`.word-app-shell`. שניהם נמדדים בדפדפן אמיתי
 * ב-scripts/qa/caret-focus-qa.mjs, וזו הסיבה שהוא רץ ב-`verify`.
 *
 * מה שנמדד:
 *   1. המחלקה `sd-v2-local-selection-caret` עדיין קיימת באריזת המנוע.
 *   2. שם האנימציה `sd-v2-local-caret-blink` עדיין קיים באריזת המנוע.
 *   3. הכלל ב-shell.css מחיל את האנימציה עם `!important` ומגדיר את ה-keyframes.
 *   4. print.css מסתיר את הסמן ואת שכבת הבחירה במדיית print.
 *   5. תחת `[data-caret-idle]` המפל מכבה בפועל — גם את האנימציה וגם את
 *      הנראות — ובלעדיה הוא מדליק.
 *   6. ‏App.vue מקשר את התכונה לשורש המעטפת, **באותו ערך** שהבורר דורש.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const ENGINE = join(process.cwd(), 'node_modules/@superdoc/docx-engine/dist/docx-engine.es.js');
const SHELL_CSS = join(process.cwd(), 'src/styles/shell.css');
const PRINT_CSS = join(process.cwd(), 'src/styles/print.css');
const APP = join(process.cwd(), 'src/App.vue');

const engineBundle = readFileSync(ENGINE, 'utf8');
const shellCss = readFileSync(SHELL_CSS, 'utf8');
const printCss = readFileSync(PRINT_CSS, 'utf8');
const app = readFileSync(APP, 'utf8');

export const CARET_CLASS = 'sd-v2-local-selection-caret';
export const CARET_BLINK_ANIMATION = 'sd-v2-local-caret-blink';
export const CARET_IDLE_ATTRIBUTE = 'data-caret-idle';

describe('חוזה סמן כתיבה מהבהב', () => {
  it('המחלקות ושמות האנימציה של המנוע עדיין קיימים באריזה', () => {
    expect(engineBundle).toContain(CARET_CLASS);
    expect(engineBundle).toContain(CARET_BLINK_ANIMATION);
  });

  it('shell.css מחיל אנימציית הבהוב על סמן הכתיבה עם !important', () => {
    expect(shellCss).toContain(`.${CARET_CLASS}`);
    expect(shellCss).toMatch(new RegExp(`animation:[^;]*${CARET_BLINK_ANIMATION}[^;]*!important`));
    expect(shellCss).toContain(`@keyframes ${CARET_BLINK_ANIMATION}`);
  });

  it('print.css מסתיר את סמן העריכה בהדפסה', () => {
    expect(printCss).toContain(`.${CARET_CLASS}`);
    expect(printCss).toMatch(new RegExp(`\\.${CARET_CLASS}[^{]*{[^}]*display:\\s*none\\s*!important`));
  });
});

/**
 * סמן אמיתי מתחת למעטפת אמיתית, עם `shell.css` כולו טעון — ההיררכיה היא זו
 * שה-DOM החי בונה (`App.vue` ← `sessions/editor-swap.ts`), ולא הצמדה של
 * הבורר לעצמו.
 */
function caretStyle(idle: boolean): CSSStyleDeclaration {
  const style = document.createElement('style');
  style.textContent = shellCss;
  document.head.appendChild(style);

  const shell = document.createElement('div');
  shell.className = 'word-app-shell';
  if (idle) shell.setAttribute(CARET_IDLE_ATTRIBUTE, 'true');

  const stack = document.createElement('div');
  stack.className = 'editor-stack';
  const host = document.createElement('div');
  host.className = 'editor-stack__host';
  const caret = document.createElement('div');
  caret.className = CARET_CLASS;
  /* המנוע מציב את התכונה דרך `dataset.v2LocalSelectionCaret` — ראו הבורר השני. */
  caret.setAttribute('data-v2-local-selection-caret', 'true');

  host.appendChild(caret);
  stack.appendChild(host);
  shell.appendChild(stack);
  document.body.appendChild(shell);

  return getComputedStyle(caret);
}

/** הכלל ב-shell.css שבגופו יש את `needle`, על בורריו. */
function ruleWith(needle: string): string {
  const rules = [...shellCss.matchAll(/(^|\n)([^\n@}][^{}]*)\{([^}]*)\}/g)];
  const hit = rules.find((rule) => rule[3].includes(needle) || rule[2].includes(needle));
  if (!hit) throw new Error(`אין ב-shell.css כלל עם ${needle}`);
  return hit[2];
}

/** רשימת הבוררים של כלל, נקייה מהערות ומרווח. */
function selectorsOf(selectorList: string): string[] {
  return selectorList
    .replace(/\/\*[\s\S]*?\*\//g, ' ')
    .split(',')
    .map((part) => part.trim())
    .filter(Boolean);
}

/**
 * סגוליות של בורר, כמספר יחיד להשוואה.
 *
 * מחלקה, תכונה ופסאודו-מחלקה שוות ערך (העמודה האמצעית), ושם אלמנט נמוך מהן
 * (העמודה הימנית). מזהים ופסאודו-אלמנטים אינם נתמכים כאן, והבדיקה מוודאת
 * שאין כאלה בבוררים שהיא סופרת.
 *
 * **וזו הרשת המשנית, לא ההכרעה.** הספירה הזאת היא ידנית, והיא אינה מבינה
 * `:where()` (סגוליות אפס), `:is()`/`:not()` (סגוליות הארגומנט החזק ביותר)
 * ולא `@layer` — כלומר ברגע שאחד מאלה ייכנס ל-shell.css היא תוכל להיפרד
 * מהדפדפן בלי להאדים. מי שמכריע באמת הוא
 * `scripts/qa/caret-focus-qa.mjs`: הוא קורא `getComputedStyle` על הסמן החי
 * במנוע אמיתי ודורש `animation: none` יחד עם `opacity: 0`. הספירה כאן קיימת
 * מפני שהיא זולה ורצה בכל commit, ומפני ש-`getComputedStyle` של jsdom מכריע
 * את המפל לפי סדר הכתיבה בגיליון ולא לפי סגוליות — כלומר בלעדיה הבדיקה
 * שמעליה עיוורת לגמרי לשאלה מי גובר על מי.
 */
function specificity(selector: string): number {
  const attributes = selector.match(/\[[^\]]*\]/g)?.length ?? 0;
  const rest = selector.replace(/\[[^\]]*\]/g, ' ');
  const classes = rest.match(/\.[A-Za-z_-][\w-]*/g)?.length ?? 0;
  const pseudoClasses = rest.match(/:[A-Za-z-]+/g)?.length ?? 0;
  const elements = rest
    .replace(/\.[A-Za-z_-][\w-]*/g, ' ')
    .replace(/:[A-Za-z-]+/g, ' ')
    .split(/[\s>+~]+/)
    .filter((token) => /^[A-Za-z]/.test(token)).length;
  return (attributes + classes + pseudoClasses) * 1_000 + elements;
}

describe('חוזה „אין מיקוד — אין סמן”', () => {
  afterEach(() => {
    document.head.innerHTML = '';
    document.body.innerHTML = '';
  });

  /* בקרה, ובלעדיה כל השורה שאחריה חסרת משמעות: כשאין תכונה — הסמן מהבהב. */
  it('בקרה: בלי התכונה המפל מדליק את ההבהוב', () => {
    expect(caretStyle(false).animation).toContain(CARET_BLINK_ANIMATION);
  });

  /*
   * וזו ההכרעה עצמה. שני הכללים `!important`, ולכן מה שמכריע הוא הסגוליות —
   * וזה בדיוק מה שנמדד כאן, במקום „הבורר נראה כמו שצריך”. כלל מדליק שיקבל
   * סגוליות גבוהה יותר הופך את השורה הזאת לאדומה.
   *
   * שתי ההצהרות ולא רק כיבוי האנימציה: אנימציה שנעצרת משאירה את מה שהיה
   * בפריים האחרון, כלומר סמן גלוי בחצי מהמקרים.
   */
  it('עם התכונה המפל מכבה — גם את האנימציה וגם את הנראות', () => {
    const idle = caretStyle(true);
    expect(idle.animation).toBe('none');
    expect(idle.opacity).toBe('0');
  });

  /*
   * ההכרעה שלמעלה נמדדת ב-jsdom, ונמדד שם גם מה שהיא **אינה** מכסה: סבב
   * מוטציה שהעלה את הכלל המדליק לארבע מחלקות השאיר אותה ירוקה, מפני
   * ש-`getComputedStyle` של jsdom מכריע לפי סדר הכתיבה בגיליון ולא לפי
   * סגוליות. בדפדפן אמיתי אותו שינוי היה מחזיר את הסמן להבהב בלי מיקוד.
   *
   * לכן הסגוליות נספרת כאן במפורש. אין בבוררים האלה מזהים ואין פסאודו-
   * אלמנטים — נבדק — ולכן הספירה היא מחלקות, תכונות ופסאודו-מחלקות מול
   * שמות אלמנטים.
   */
  it('הכלל המכבה גובר על המדליק גם בספירת סגוליות', () => {
    const on = selectorsOf(ruleWith(CARET_BLINK_ANIMATION));
    const off = selectorsOf(ruleWith(`[${CARET_IDLE_ATTRIBUTE}`));
    expect(on.length).toBeGreaterThan(0);
    expect(off.length).toBeGreaterThan(0);

    for (const selector of [...on, ...off]) {
      expect(selector, `מזהה או פסאודו-אלמנט בבורר: ${selector}`).not.toMatch(/#|::/);
    }

    const strongestOn = Math.max(...on.map(specificity));
    const weakestOff = Math.min(...off.map(specificity));
    expect(weakestOff, `${off.join(', ')} חייב לגבור על ${on.join(', ')}`).toBeGreaterThan(strongestOn);
  });

  /*
   * הבורר השני, `[data-v2-local-selection-caret]`, הוא התכונה שהמנוע מציב על
   * אותו `div` דרך `dataset` — ולכן היא אינה מופיעה בטקסט של האריזה. שני
   * הבוררים חייבים להסכים: מי שיתקן אחד מהם וישכח את השני משאיר חצי סמן.
   */
  it('גם הבורר של התכונה, ולא רק זה של המחלקה', () => {
    const style = document.createElement('style');
    style.textContent = shellCss;
    document.head.appendChild(style);

    const shell = document.createElement('div');
    shell.className = 'word-app-shell';
    shell.setAttribute(CARET_IDLE_ATTRIBUTE, 'true');
    const caret = document.createElement('div');
    /* בלי המחלקה בכלל: מה שמכבה כאן הוא הבורר של התכונה. */
    caret.setAttribute('data-v2-local-selection-caret', 'true');
    shell.appendChild(caret);
    document.body.appendChild(shell);

    expect(getComputedStyle(caret).opacity).toBe('0');
  });

  /*
   * הקישור בין הגיליון לתבנית, ולא רק „יש קישור”: הבורר דורש ערך מסוים
   * (`[data-caret-idle='true']`), ותבנית שתקשור ערך אחר — או `null` — תשאיר
   * את הסמן מהבהב בלי שאיש יבחין. את החיים עצמם מודד scripts/qa/caret-focus-qa.mjs.
   */
  it('App.vue קושר את התכונה בערך שהבורר דורש', () => {
    const required = new RegExp(`\\[${CARET_IDLE_ATTRIBUTE}=['"]([^'"\\]]+)['"]\\]`).exec(shellCss);
    expect(required, `אין ב-shell.css בורר על [${CARET_IDLE_ATTRIBUTE}] עם ערך`).not.toBeNull();

    const binding = new RegExp(`:${CARET_IDLE_ATTRIBUTE}="([^"]*)"`).exec(app);
    expect(binding, `App.vue אינו קושר :${CARET_IDLE_ATTRIBUTE}`).not.toBeNull();
    expect(binding?.[1]).toContain(`'${required?.[1]}'`);

    expect(app).toContain('watchCaretFocus');
  });
});

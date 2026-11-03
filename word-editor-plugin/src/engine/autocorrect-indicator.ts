/**
 * סימון התיקון האוטומטי, כמו ב-Word: אחרי שמילה הוחלפה מופיע, כשמרחפים מעליה,
 * פס כחול קטן מתחת לאות הראשונה שלה; ריחוף על הפס הופך אותו לכפתור (ברק), ולחיצה
 * עליו פותחת תפריט:
 *
 * - **חזור ל„המקור”** — מחזיר את מה שהוקלד.
 * - **הפסק לתקן אוטומטית את „המקור”** — מחזיר, ומוריד את הרשומה מהטבלה (או מכבה
 *   את החוק, אם התיקון בא מחוק ולא מהטבלה).
 * - **שליטה באפשרויות התיקון האוטומטי…** — פותח את הדיאלוג.
 *
 * ## למה overlay חיצוני
 *
 * כמו `book-completion-overlay.ts`: אין API ציבורי לקשט טקסט בתוך המסמך, ולכן
 * הסימון הוא אלמנטים חיצוניים ב-`document.body`, ממוקמים בקואורדינטות מסך.
 *
 * ## המיקום נקרא מה-DOM המצויר, בכל ריחוף
 *
 * הסימון שומר היכן המילה בבלוק (בלוק + היסט). בכל ריחוף נאספים נתוני הבלוק פעם
 * אחת, וכל מחוון בודק את מלוא הטקסט שלו ומודד אותו עם `Range`.
 * ‏`ui.selection.getAnchorRect` מחזיר מלבן של הסמן בלבד, ולכן אינו משמש כאן.
 *
 * ## מתי הסימון נעלם
 *
 * המחוונים מההערכה האחרונה נשארים עד לעריכה חדשה, פעולה בממשק או החלפת המסמך.
 * הקשר ייחודי בטקסט
 * הנוכחי אינו מוכיח זהות אחרי מחיקה, ולכן אין שימוש חוזר במחוון אחרי עריכה.
 * ‏„חזור” בודק שוב לפני שהוא מחליף.
 */
import type { AutoCorrectFix, AutoCorrectOptions, AutoCorrectRule } from './autocorrect';
import { ICONS } from '../ui/icons/icons';
import { blockIdAt, createTextLocator, locateText, type LocatedWord, type TextContext } from './painted-lines';

/* ------------------------------------------------------------------ */
/* החלק הטהור: מה כתוב בתפריט, ומה „הפסק לתקן” משנה                       */
/* ------------------------------------------------------------------ */

/** מה שהסימון צריך לדעת על תיקון — בלי היסטים. */
export interface FixSummary {
  readonly rule: AutoCorrectRule;
  readonly original: string;
  readonly replacement: string;
}

export type FlagKey = Exclude<keyof AutoCorrectOptions, 'enabled' | 'replacements'>;

/** החוק → הדגל שמכבה אותו. `table` אינו כאן: הוא מוריד רשומה, לא דגל. */
export const RULE_FLAG: Readonly<Record<Exclude<AutoCorrectRule, 'table'>, FlagKey>> = {
  dashes: 'dashes',
  'spaced-hyphen': 'spacedHyphen',
  gershayim: 'gershayim',
  'final-letters': 'finalLetters',
  'two-initial-caps': 'twoInitialCaps',
  'capitalize-sentence': 'capitalizeSentences',
  'capitalize-day': 'capitalizeDays',
};

const STOP_LABEL: Readonly<Record<Exclude<AutoCorrectRule, 'table'>, string>> = {
  dashes: 'הפסק להחליף מקפים בקו מפריד',
  'spaced-hyphen': 'הפסק להחליף מקף עם רווחים',
  gershayim: 'הפסק להחליף מירכאות וגרש בגרשיים',
  'final-letters': 'הפסק לתקן אותיות סופיות',
  'two-initial-caps': 'הפסק לתקן שתי אותיות ראשונות גדולות',
  'capitalize-sentence': 'הפסק להגדיל אות ראשונה בתחילת משפט',
  'capitalize-day': 'הפסק להגדיל אות ראשונה בשמות ימים',
};

export type MenuItemId = 'undo' | 'stop' | 'options';

export interface MenuItem {
  readonly id: MenuItemId;
  readonly label: string;
}

/** פותח ב„ ומסיים ב” — כך שמקור לטיני בתוך משפט עברי לא מתהפך בתפריט. */
function quoted(text: string): string {
  return `„${text}”`;
}

/**
 * פריטי התפריט, בסדר של Word. „הפסק” ו„אפשרויות” מופיעים רק אם יש מי שיטפל
 * בהם — תפריט עם פריט שאינו עושה דבר גרוע מתפריט קצר.
 */
export function buildMenu(fix: FixSummary, can: { stop: boolean; options: boolean }): MenuItem[] {
  const items: MenuItem[] = [{ id: 'undo', label: `חזור ל${quoted(fix.original)}` }];
  if (can.stop) {
    items.push({
      id: 'stop',
      label:
        fix.rule === 'table' ? `הפסק לתקן אוטומטית את ${quoted(fix.original)}` : STOP_LABEL[fix.rule],
    });
  }
  if (can.options) items.push({ id: 'options', label: 'שליטה באפשרויות התיקון האוטומטי…' });
  return items;
}

/**
 * ההגדרות אחרי „הפסק לתקן”. תיקון מהטבלה מוריד את הרשומה (ההשוואה בלי הבדל
 * רישיות: `Teh` שהוקלדה תואמת לרשומה `teh`); תיקון מחוק מכבה את הדגל שלו.
 * מחזירה את אותו אובייקט כשאין מה לשנות.
 */
export function stopCorrecting(options: AutoCorrectOptions, fix: FixSummary): AutoCorrectOptions {
  if (fix.rule === 'table') {
    const key = fix.original.toLowerCase();
    const kept = options.replacements.filter((entry) => entry.from.toLowerCase() !== key);
    return kept.length === options.replacements.length ? options : { ...options, replacements: kept };
  }
  const flag = RULE_FLAG[fix.rule];
  return options[flag] ? { ...options, [flag]: false } : options;
}

/** תוצאת החלת התיקונים על חלון הטקסט: הטקסט החדש, והיסט ההתחלה החדש של כל תיקון. */
export interface AppliedText {
  readonly text: string;
  /** באותו סדר שבו התיקונים נמסרו. */
  readonly starts: number[];
}

/**
 * מה יהיה כתוב בחלון אחרי שכל התיקונים הוחלו, ואיפה מתחילה כל מילה שהוחלפה.
 * התיקונים אינם חופפים (`planAutoCorrect` מבטיחה זאת). תיקון שלפניו קוצר או
 * האריך את הטקסט מזיז את ההיסט של אלה שאחריו — זה מה שמחושב כאן.
 */
export function applyFixesToText(text: string, fixes: readonly AutoCorrectFix[]): AppliedText {
  const order = fixes.map((_, index) => index).sort((a, b) => fixes[a]!.start - fixes[b]!.start);
  const starts = new Array<number>(fixes.length).fill(0);
  let out = '';
  let cursor = 0;
  for (const index of order) {
    const fix = fixes[index]!;
    out += text.slice(cursor, fix.start);
    starts[index] = out.length;
    out += fix.replacement;
    cursor = fix.end;
  }
  out += text.slice(cursor);
  return { text: out, starts };
}

/* ------------------------------------------------------------------ */
/* הציור                                                                */
/* ------------------------------------------------------------------ */

/**
 * תיקון שהוחל. מכיל רק **היכן הוא בבלוק** (לא היכן הוא על המסך): המיקום על
 * המסך נקרא מחדש מה-DOM המצויר בכל ריחוף (`locateWord`), ולכן הוא נכון גם אחרי
 * גלילה, שינוי גודל וגלישת שורה, וגם אחרי שהמשתמש המשיך להקליד.
 */
export interface IndicatorEntry extends FixSummary {
  readonly blockId: string;
  readonly story: unknown;
  /** היסט אבסולוטי בבלוק שבו מתחילה ההחלפה. מתעדכן כשעריכה לפניה מזיזה אותה. */
  readonly start: number;
  /** מה שהיה סביבה, כדי למצוא אותה אחרי עריכה לפניה. */
  readonly context?: TextContext;
}

export type { Box, LocatedWord } from './painted-lines';

/**
 * היכן המילה שהוחלפה מצוירת עכשיו, או `null`. הקריאה עצמה ב-`painted-lines.ts`
 * (`locateText`): הכלל ב-engine-boundaries.test.ts הוא שרק שם נוגעים בעיגוני ה-pm
 * של השורות המצוירות. כאן רק מה שהסימון מוסיף: הטקסט שנבדק הוא מה שהוחלף.
 */
export function locateWord(
  entry: Pick<IndicatorEntry, 'blockId' | 'start' | 'replacement' | 'context'>,
  root: ParentNode = document,
): LocatedWord | null {
  return locateText(entry.blockId, entry.start, entry.replacement, root, entry.context);
}

export interface IndicatorCallbacks {
  /** „חזור ל…”. מחזירה אחרי שההחלפה נעשתה או נדחתה. */
  changeBack: (entry: IndicatorEntry) => Promise<void> | void;
  /** „הפסק לתקן”. חסר = הפריט אינו מוצג. */
  stop?: (entry: IndicatorEntry) => Promise<void> | void;
  /** „שליטה באפשרויות”. חסר = הפריט אינו מוצג. */
  openOptions?: () => void;
  /** עריכה פוסלת קריאות באוויר; שינוי הקשר מאפס גם את רצף ההקלדה. */
  invalidate?: (resetTyping: boolean, event?: Event) => void;
}

export interface IndicatorHandle {
  show(entries: readonly IndicatorEntry[]): void;
  clear(): void;
  dispose(): void;
  /** יש סימון חי (מוצג או ממתין לריחוף). */
  readonly active: boolean;
}

const STYLE_ID = 'otzaria-acorr-indicator-style';
const CHIP_CLASS = 'otzaria-acorr-chip';
const MENU_CLASS = 'otzaria-acorr-menu';
/**
 * אין תפוגה בזמן: המחוונים של אותה הערכה נשארים עד לאינטראקציה הבאה.
 * עריכה חדשה פוסלת אותם, כי אין כאן עוגני זהות שעוקבים אחרי שינויי המסמך.
 * התקרה מגבילה את הזיכרון גם לקוראים שמוסיפים מחוונים רבים.
 */
const MAX_ENTRIES = 1000;
/** כמה פיקסלים מתחת לשורה עדיין נחשבים „מעל המילה” לצורך הריחוף. */
const HOVER_SLACK_BELOW = 6;
const CHIP_SIZE = 22;

/*
 * הצבעים, הגופן, הרדיוסים והמעברים כולם טוקנים של `styles/tokens.css`, ולכן
 * הסימון עוקב אחרי ערכת הנושא של אוצריא (בהיר/כהה/צבעוני) בלי `@media` משלו.
 * התפריט משחזר את `.ctx-menu` / `.ctx-btn--item` של תפריט ההקשר (ui/menu): אותו
 * כרטיס, אותה מסגרת, אותו צל, אותו גובה פריט ואותו hover. z-index 2500 הוא של
 * תפריט ההקשר (ראו ContextMenu.vue), והסימון עצמו מתחתיו.
 */
const CSS = `
.${CHIP_CLASS}{position:fixed;z-index:2400;display:none;box-sizing:border-box;margin:0;padding:0;border:0;
  height:2px;min-width:5px;background:var(--color-primary);border-radius:var(--radius-xs);cursor:pointer;
  color:var(--color-primary);line-height:0;overflow:hidden}
.${CHIP_CLASS}::before{content:"";position:absolute;inset:-8px -4px -10px -4px}
.${CHIP_CLASS} .svg-icon{display:none;width:16px;height:16px}
.${CHIP_CLASS} svg{width:100%;height:100%;fill:currentColor;shape-rendering:geometricPrecision}
.${CHIP_CLASS}:hover,.${CHIP_CLASS}[data-open="true"]{width:${CHIP_SIZE}px!important;height:${CHIP_SIZE}px;
  background:var(--color-surface);border:1px solid var(--color-primary);border-radius:var(--radius-sm);
  box-shadow:0 4px 16px rgba(0,0,0,.16);display:flex!important;align-items:center;justify-content:center;overflow:visible}
.${CHIP_CLASS}:hover .svg-icon,.${CHIP_CLASS}[data-open="true"] .svg-icon{display:inline-flex}
.${CHIP_CLASS}:focus-visible{outline:2px solid var(--color-primary);outline-offset:1px}
.${MENU_CLASS}{position:fixed;z-index:2500;display:none;box-sizing:border-box;min-width:240px;max-width:320px;padding:4px;
  background:var(--color-surface);color:var(--color-on-surface);border:1px solid var(--color-outline);
  border-radius:var(--radius-md);box-shadow:0 4px 16px rgba(0,0,0,.16);font-family:var(--font-main)}
.${MENU_CLASS}:focus{outline:none}
.${MENU_CLASS} button{display:flex;align-items:center;gap:8px;width:100%;height:30px;box-sizing:border-box;margin:0;
  padding:0 10px;border:1px solid transparent;border-radius:var(--radius-sm);background:transparent;
  color:var(--color-on-surface);font:inherit;font-size:13px;text-align:start;cursor:pointer;white-space:nowrap}
.${MENU_CLASS} button:hover{background:var(--word-btn-hover)}
.${MENU_CLASS} button:focus-visible{outline:2px solid var(--color-primary);outline-offset:-2px}
.${MENU_CLASS} .svg-icon{display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;width:16px;height:16px}
.${MENU_CLASS} .svg-icon svg{width:100%;height:100%;fill:currentColor}
.${MENU_CLASS} hr{height:1px;margin:4px 6px;border:0;background:var(--color-outline-variant)}
`;

/** האייקונים הם של ספריית הסרגל (`ui/icons/icons.ts`, Fluent System Icons) — לא ציור משלנו. */
const MENU_ICON: Readonly<Record<MenuItemId, string>> = { undo: 'undo', stop: 'reject', options: 'replace' };

function iconSpan(name: string): string {
  return `<span class="svg-icon" aria-hidden="true">${ICONS[name] ?? ''}</span>`;
}

function ensureStyle(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}

export function createAutoCorrectIndicator(
  container: HTMLElement,
  callbacks: IndicatorCallbacks,
  now: () => number = () => Date.now(),
): IndicatorHandle {
  ensureStyle();

  let entries: { entry: IndicatorEntry; at: number }[] = [];
  let current: IndicatorEntry | null = null;
  let menuOpen = false;
  let disposed = false;

  const chip = document.createElement('button');
  chip.type = 'button';
  chip.className = CHIP_CLASS;
  chip.innerHTML = iconSpan('replace');
  chip.setAttribute('aria-haspopup', 'menu');
  chip.setAttribute('aria-label', 'אפשרויות תיקון אוטומטי');
  chip.tabIndex = -1;

  const menu = document.createElement('div');
  menu.className = MENU_CLASS;
  menu.setAttribute('role', 'menu');

  document.body.appendChild(chip);
  document.body.appendChild(menu);

  /** לחיצה על הסימון אינה לוקחת פוקוס מהמסמך: ההקלדה ממשיכה במקום שהייתה. */
  const keepFocus = (event: Event): void => event.preventDefault();
  chip.addEventListener('pointerdown', keepFocus);
  chip.addEventListener('mousedown', keepFocus);
  menu.addEventListener('pointerdown', keepFocus);
  menu.addEventListener('mousedown', keepFocus);

  function hideChip(): void {
    chip.style.display = 'none';
    chip.removeAttribute('data-open');
    current = null;
  }

  function closeMenu(): void {
    menuOpen = false;
    menu.style.display = 'none';
    menu.textContent = '';
    chip.removeAttribute('data-open');
  }

  /** פס כחול מתחת לאות הראשונה של המילה — בימין ב-RTL, בשמאל ב-LTR, כפי שהדפדפן מצייר אותה. */
  function placeChip(entry: IndicatorEntry, found: LocatedWord): void {
    const width = Math.max(5, Math.round(found.first.right - found.first.left));
    chip.style.width = `${width}px`;
    chip.style.left = `${Math.round(found.first.left)}px`;
    chip.style.top = `${Math.round(found.first.bottom - 2)}px`;
    chip.style.display = 'block';
    current = entry;
  }

  function isOver(found: LocatedWord, x: number, y: number): boolean {
    return found.rects.some(
      (box) => x >= box.left - 1 && x <= box.right + 1 && y >= box.top && y <= box.bottom + HOVER_SLACK_BELOW,
    );
  }

  /** המילה שמתחת לעכבר, מתוך הסימונים החיים. כל מיקום נקרא עכשיו מה-DOM. */
  function entryAt(
    x: number,
    y: number,
    target: Node | null,
  ): { entry: IndicatorEntry; found: LocatedWord } | null {
    // רק סימונים של הבלוק שמתחת לעכבר: אחרת כל תזוזה סורקת את כל הסימונים החיים.
    const block = blockIdAt(target);
    if (block === null) return null;
    const matching = entries.filter((item) => item.entry.blockId === block);
    if (matching.length === 0) return null;
    const locate = createTextLocator(block, container);
    for (const item of matching) {
      const found = locate(item.entry.start, item.entry.replacement, item.entry.context);
      if (!found || !isOver(found, x, y)) continue;
      // עריכה לפני המילה הזיזה אותה: ההיסט החדש נשמר, ו"חזור" פועל עליו.
      if (found.start !== item.entry.start) item.entry = { ...item.entry, start: found.start };
      return { entry: item.entry, found };
    }
    return null;
  }

  /** ריחוף: קריאה אחת בפריים, ולא על כל אירוע עכבר. */
  let pending: { x: number; y: number; target: Node | null } | null = null;
  let frame = 0;

  function scan(): void {
    frame = 0;
    const move = pending;
    pending = null;
    if (!move || disposed || entries.length === 0 || menuOpen) return;
    if (move.target && (move.target === chip || chip.contains(move.target))) return;
    // מחוץ למסמך (סרגל, תפריטים) אין מה לסמן.
    if (!move.target || !container.contains(move.target)) {
      if (current) hideChip();
      return;
    }
    const hit = entryAt(move.x, move.y, move.target);
    if (hit) placeChip(hit.entry, hit.found);
    else if (current) hideChip();
  }

  function onPointerMove(event: PointerEvent): void {
    if (disposed || entries.length === 0 || menuOpen) return;
    pending = { x: event.clientX, y: event.clientY, target: event.target as Node | null };
    if (!frame) frame = requestAnimationFrame(scan);
  }

  function openMenu(entry: IndicatorEntry): void {
    const items = buildMenu(entry, { stop: !!callbacks.stop, options: !!callbacks.openOptions });
    menu.textContent = '';
    items.forEach((item, index) => {
      if (item.id === 'options' && index > 0) menu.appendChild(document.createElement('hr'));
      const button = document.createElement('button');
      button.type = 'button';
      button.setAttribute('role', 'menuitem');
      button.innerHTML = iconSpan(MENU_ICON[item.id]);
      const label = document.createElement('span');
      label.textContent = item.label;
      button.appendChild(label);
      button.addEventListener('click', () => void run(item.id, entry));
      menu.appendChild(button);
    });

    menu.style.display = 'block';
    const chipRect = chip.getBoundingClientRect();
    const menuRect = menu.getBoundingClientRect();
    // מיושר לפס, ונדחס פנימה אם חורג מהחלון.
    const rtl = getComputedStyle(document.documentElement).direction === 'rtl';
    let left = rtl ? chipRect.right - menuRect.width : chipRect.left;
    left = Math.max(4, Math.min(left, window.innerWidth - menuRect.width - 4));
    let top = chipRect.bottom + 2;
    if (top + menuRect.height > window.innerHeight - 4) top = Math.max(4, chipRect.top - menuRect.height - 2);
    menu.style.left = `${Math.round(left)}px`;
    menu.style.top = `${Math.round(top)}px`;

    menuOpen = true;
    chip.setAttribute('data-open', 'true');
  }

  async function run(id: MenuItemId, entry: IndicatorEntry): Promise<void> {
    // רק הסימון שנבחר מוסר; סימונים אחרים של מילים שעדיין במקומן נשארים.
    entries = entries.filter((item) => item.entry !== entry);
    closeMenu();
    hideChip();
    try {
      if (id === 'undo') await callbacks.changeBack(entry);
      else if (id === 'stop') {
        await callbacks.changeBack(entry);
        await callbacks.stop?.(entry);
      } else callbacks.openOptions?.();
    } catch (error) {
      console.warn('[otzaria-word] תיקון אוטומטי: פעולת הסימון נכשלה', error);
    }
  }

  chip.addEventListener('click', () => {
    if (!current) return;
    if (menuOpen) closeMenu();
    else openMenu(current);
  });

  function clear(): void {
    entries = [];
    closeMenu();
    hideChip();
  }

  const onDocPointerDown = (event: PointerEvent): void => {
    const target = event.target as Node | null;
    if (target && (chip.contains(target) || menu.contains(target))) return;
    callbacks.invalidate?.(true);
    clear();
  };
  const onEdit = (event: Event): void => {
    callbacks.invalidate?.(false, event);
    clear();
  };
  const onKeyDown = (event: KeyboardEvent): void => {
    const target = event.target as Node | null;
    if (target && container.contains(target) &&
      (event.key.length === 1 || ['Backspace', 'Delete', 'Enter', 'Tab'].includes(event.key))) onEdit(event);
    else { closeMenu(); hideChip(); }
  };
  const onBlur = (): void => {
    callbacks.invalidate?.(true);
    clear();
  };

  document.addEventListener('pointermove', onPointerMove, { passive: true });
  document.addEventListener('pointerdown', onDocPointerDown, true);
  document.addEventListener('keydown', onKeyDown, true);
  for (const name of ['input', 'compositionstart', 'paste', 'cut', 'drop']) container.addEventListener(name, onEdit, true);
  window.addEventListener('blur', onBlur);

  return {
    show(next) {
      if (disposed) return;
      const at = now();
      // תיקון חדש באותו מקום מחליף את הישן; אחרת הסימון הישן היה מצביע על טקסט שהוחלף.
      const fresh = next.map((entry) => ({ entry, at }));
      const kept = entries.filter(
        (old) => !fresh.some((n) => n.entry.blockId === old.entry.blockId && n.entry.start === old.entry.start),
      );
      entries = [...kept, ...fresh].slice(-MAX_ENTRIES);
    },
    clear,
    dispose() {
      disposed = true;
      clear();
      document.removeEventListener('pointermove', onPointerMove);
      document.removeEventListener('pointerdown', onDocPointerDown, true);
      document.removeEventListener('keydown', onKeyDown, true);
      for (const name of ['input', 'compositionstart', 'paste', 'cut', 'drop']) container.removeEventListener(name, onEdit, true);
      window.removeEventListener('blur', onBlur);
      if (frame) cancelAnimationFrame(frame);
      chip.remove();
      menu.remove();
    },
    get active() {
      return entries.length > 0;
    },
  };
}

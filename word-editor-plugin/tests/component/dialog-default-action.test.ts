/**
 * Enter בדיאלוג = הכפתור הראשי, כמו ב-Word.
 *
 * מה שנמדד: שה-Enter עולה מכל מקום בדיאלוג — גם משורש הדיאלוג עצמו, שהוא
 * מה שממוקד ברגע הפתיחה ב„גופן מתקדם” — ושאינו חוטף Enter ממי שהוא שלו
 * (כפתור „ביטול” שממוקד, שדה רב-שורות, IME).
 *
 * הבדיקה מודדת דרך הכפתור: `[data-default-action]` הוא ההצהרה היחידה בכל
 * דיאלוג, וההסתמכות עליו היא מה שנותן ל„מנוטרל” ול„busy” של כל דיאלוג לחסום
 * את ה-Enter בלי שהמנגנון יכיר את הוולידציה שלו.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import FontAdvancedDialog from '../../src/ui/panels/FontAdvancedDialog.vue';
import CitationSourceDialog from '../../src/ui/panels/CitationSourceDialog.vue';
import { autoUnmount, createSuperdocDouble, mountUi, settle } from './harness';

autoUnmount();

const PANELS = join(__dirname, '..', '..', 'src', 'ui', 'panels');

function dialog(selector = '.fontadv-dialog'): HTMLElement {
  const element = document.querySelector<HTMLElement>(selector);
  if (!element) throw new Error(`הדיאלוג ${selector} אינו בגוף הדף`);
  return element;
}

function enter(target: EventTarget, init: KeyboardEventInit = {}): void {
  target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, ...init }));
}

/**
 * „גופן מתקדם" **על טקסט מסומן** — התנאי המוקדם שלו.
 *
 * הדיאלוג נועל את „אישור" כשאין בחירה ואומר למה (ראו
 * `hasRangeSelection` ב-engine/font-advanced.ts), ו-Enter הוא לחיצה על
 * הכפתור הראשי — כלומר בלי בחירה גם הוא, בצדק, אינו עושה דבר. מה שנמדד כאן
 * הוא המנגנון של Enter, ולכן הבחירה מוצהרת ואינה נסמכת על ברירת המחדל של
 * הדמה.
 */
function mountFontAdvanced(props: { isOpen: boolean; busy: boolean }) {
  return mountUi(FontAdvancedDialog, {
    props,
    superdoc: createSuperdocDouble({ selection: { hasRange: true, text: 'טקסט מסומן' } }),
  });
}

describe('Enter מפעיל את הכפתור הראשי', () => {
  it('Enter על שורש הדיאלוג — המצב שבו הוא נפתח — שולח', async () => {
    // פתיחה אמיתית ולא הרכבה פתוחה: זה מה שמעביר את המיקוד לשורש, וזו הנקודה
    // שבה Enter לא עשה כלום קודם — אף שדה אינו ממוקד.
    const harness = mountFontAdvanced({ isOpen: false, busy: false });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();

    /*
     * שדה אחד מלא, ובכוונה: „אישור" נעול על דיאלוג שאין בו מה להחיל, ולכן
     * Enter על דיאלוג ריק **אמור** לא לעשות דבר. מה שנמדד כאן הוא שה-Enter
     * מגיע לכפתור, ולא שהוא מגיע גם כשאין לו מה לעשות.
     */
    const scale = dialog().querySelector<HTMLInputElement>('#fa-scale');
    if (!scale) throw new Error('אין שדה מתיחה');
    scale.value = '120';
    scale.dispatchEvent(new Event('input'));
    await settle();

    const root = dialog();
    expect(document.activeElement).toBe(root);
    enter(root);
    await settle();

    expect(harness.wrapper.emitted('submit')).toHaveLength(1);
  });

  it('Enter בשדה מספר שולח פעם אחת בלבד', async () => {
    const harness = mountFontAdvanced({ isOpen: true, busy: false });
    await settle();

    const input = dialog().querySelector<HTMLInputElement>('#fa-scale');
    if (!input) throw new Error('אין שדה מתיחה');
    input.value = '120';
    input.dispatchEvent(new Event('input'));
    await settle();
    enter(input);
    await settle();

    // אילו נשאר גם ה-`@keydown.enter` על השדה, הפעולה הייתה יוצאת פעמיים —
    // כלומר שתי פקודות למנוע על לחיצה אחת.
    expect(harness.wrapper.emitted('submit')).toHaveLength(1);
    expect(harness.wrapper.emitted('submit')?.[0]).toEqual([{ charScale: 120 }]);
  });

  it('Enter בבורר שולח — שם לא היה מטפל כלל', async () => {
    const harness = mountFontAdvanced({ isOpen: true, busy: false });
    await settle();

    // בורר שפת ההגהה הוא ה-`<select>` היחיד שנשאר בדיאלוג; שאר הבוררים הפכו
    // לכפתורי מיתוג, ועל כפתור ה-Enter הוא של הכפתור עצמו (SELF_ACTIVATING).
    const select = dialog().querySelector<HTMLSelectElement>('#fa-lang');
    if (!select) throw new Error('אין בורר');
    select.value = 'he-IL';
    select.dispatchEvent(new Event('change'));
    await settle();
    enter(select);
    await settle();

    expect(harness.wrapper.emitted('submit')).toHaveLength(1);
  });

  it('כפתור ממוקד שומר את ה-Enter של עצמו', async () => {
    const harness = mountFontAdvanced({ isOpen: true, busy: false });
    await settle();

    const cancel = [...dialog().querySelectorAll('.fa-footer .fa-btn')].find(
      (button) => button.textContent?.trim() === 'ביטול',
    );
    if (!cancel) throw new Error('אין „ביטול”');
    enter(cancel);
    await settle();

    expect(harness.wrapper.emitted('submit')).toBeUndefined();
  });

  it('כפתור ראשי מנוטרל (busy) אינו נלחץ', async () => {
    const harness = mountFontAdvanced({ isOpen: true, busy: true });
    await settle();

    enter(dialog());
    await settle();

    expect(harness.wrapper.emitted('submit')).toBeUndefined();
  });

  it('Shift+Enter ו-Ctrl+Enter אינם אישור', async () => {
    const harness = mountFontAdvanced({ isOpen: true, busy: false });
    await settle();

    enter(dialog(), { shiftKey: true });
    enter(dialog(), { ctrlKey: true });
    await settle();

    expect(harness.wrapper.emitted('submit')).toBeUndefined();
  });

  /**
   * IME: ה-Enter שסוגר בחירה בחלונית ההרכבה אינו אישור של הדיאלוג. בלי
   * הבדיקה הזאת כל הקלדה בעברית מנוקדת או בכל שיטת קלט אחרת הייתה יכולה
   * לסגור את הדיאלוג באמצע מילה.
   */
  it('Enter של IME אינו אישור', async () => {
    const harness = mountFontAdvanced({ isOpen: true, busy: false });
    await settle();

    enter(dialog(), { isComposing: true } as KeyboardEventInit);
    await settle();

    expect(harness.wrapper.emitted('submit')).toBeUndefined();
  });

  /**
   * שדה רב-שורות: Enter הוא שורה חדשה. `CitationSourceDialog` הוא הדיאלוג
   * היחיד במשפחה עם `<textarea>` בגוף („מחבר: שם לכל שורה”), ולכן הוא זה
   * שנבדק — ובאותה הרכבה נמדד גם הצד השני: שדה חד-שורתי כן מאשר.
   */
  it('Enter ב-textarea אינו אישור, ובשדה שלידו כן', async () => {
    const harness = mountUi(CitationSourceDialog, { props: { isOpen: true } });
    await settle();

    const root = dialog('.citation-source-dialog');
    const title = root.querySelector<HTMLInputElement>('#cs-title');
    const authors = root.querySelector('textarea');
    if (!title || !authors) throw new Error('אין שדות במקור');

    title.value = 'שולחן ערוך';
    title.dispatchEvent(new Event('input'));
    await settle();

    enter(authors);
    await settle();
    expect(harness.wrapper.emitted('add')).toBeUndefined();

    enter(title);
    await settle();
    expect(harness.wrapper.emitted('add')).toHaveLength(1);
  });
});

/**
 * הדיאלוגים שיש להם כפתור אישור אחד — שם Enter חייב לעבוד. שלושת החריגים
 * מוצהרים כאן, וכל אחד מהם הוא החלטה ולא שכחה.
 */
describe('משפחת ברירת המחדל', () => {
  const EXEMPT: Record<string, string> = {
    'FindReplaceDialog.vue': 'ה-Enter שלו כבר „מצא הבא”, ו-Shift+Enter אחורה',
    'MacrosDialog.vue': 'שלושה כפתורים ראשיים — אין ברירת מחדל אחת',
    'ShulchanUnclosedDialog.vue': 'רשימת ממצאים בלי כפתור אישור',
  };

  const anchored = readdirSync(PANELS)
    .filter((name) => name.endsWith('.vue'))
    .map((name) => ({ name, source: readFileSync(join(PANELS, name), 'utf8') }))
    .filter(({ source }) => source.includes('inset-inline-start: 40px'));

  it.each(anchored.map(({ name }) => name))('%s — Enter מפעיל את הראשי', (name) => {
    const source = anchored.find((file) => file.name === name)!.source;
    if (EXEMPT[name]) {
      expect(source).not.toContain('data-default-action');
      return;
    }
    expect(source).toContain('data-default-action');
    expect(source).toContain('@keydown.enter="onDialogEnter"');
    expect(source).toContain(
      `import { useDialogDefaultAction } from '../../composables/dialog-default-action';`,
    );
    // Enter כפול על שדה בודד היה שולח את אותה פעולה פעמיים.
    expect(source).not.toMatch(/@keydown\.enter="(?!onDialogEnter)/);
  });
});

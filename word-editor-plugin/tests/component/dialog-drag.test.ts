/**
 * הגרירה כפי שהיא מחווטת בפועל: לחיצה בכותרת, תזוזה, ושחרור.
 *
 * מה שנמדד כאן אינו הגאומטריה (זו נבדקת ב-tests/unit/dialog-drag.test.ts, על
 * מלבנים אמיתיים) אלא **החוט**: שה-`@pointerdown` יושב על הכותרת, שה-`:style`
 * יושב על השורש, ושהמאזינים על `window` נעלמים בשחרור. jsdom מחזיר מלבן של
 * אפסים מכל `getBoundingClientRect`, ולכן התזוזה כאן היא בדיוק תזוזת המצביע.
 *
 * הסריקה בסוף הקובץ היא השער האמיתי: עשרים וארבעה דיאלוגים יושבים על אותו
 * עוגן קבוע, והדיאלוג העשרים-וחמישי שייכתב יישכח — הוא נראה תקין לחלוטין בלי
 * הידית, פשוט לא זז.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import FontAdvancedDialog from '../../src/ui/panels/FontAdvancedDialog.vue';
import { autoUnmount, mountUi, settle } from './harness';

autoUnmount();

const PANELS = join(__dirname, '..', '..', 'src', 'ui', 'panels');

function dialog(): HTMLElement {
  const element = document.querySelector<HTMLElement>('.fontadv-dialog');
  if (!element) throw new Error('הדיאלוג אינו בגוף הדף');
  return element;
}

/**
 * לחיצה־גרירה־שחרור, כמו שהמשתמש עושה.
 *
 * `MouseEvent` ולא `PointerEvent`: jsdom אינו מממש את השני, ושדות המיקום —
 * מה שהגרירה קוראת — זהים בשניהם.
 */
function pointer(type: string, target: EventTarget, x: number, y: number): void {
  target.dispatchEvent(new MouseEvent(type, { bubbles: true, clientX: x, clientY: y, button: 0 }));
}

describe('גרירת דיאלוג בכותרת', () => {
  it('גרירה בכותרת מזיזה את הדיאלוג', async () => {
    mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false } });
    await settle();

    const header = dialog().querySelector('.fa-header');
    if (!header) throw new Error('אין כותרת');

    pointer('pointerdown', header, 100, 200);
    pointer('pointermove', window, 220, 290);
    await settle();

    expect(dialog().style.left).toBe('120px');
    expect(dialog().style.top).toBe('90px');
    // בלי ניטרול הקצוות הלוגיים ה-CSS היה גובר על `left` בכיוון ימין-לשמאל.
    expect(dialog().style.insetInlineStart).toBe('auto');
    expect(dialog().style.insetInlineEnd).toBe('auto');
  });

  it('שחרור מנתק את המאזינים — הסמן ממשיך והדיאלוג לא', async () => {
    mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false } });
    await settle();

    const header = dialog().querySelector('.fa-header');
    if (!header) throw new Error('אין כותרת');

    pointer('pointerdown', header, 100, 200);
    pointer('pointermove', window, 200, 250);
    pointer('pointerup', window, 200, 250);
    pointer('pointermove', window, 600, 600);
    await settle();

    expect(dialog().style.left).toBe('100px');
    expect(dialog().style.top).toBe('50px');
  });

  /**
   * כפתור הסגירה יושב בתוך הידית. גרירה שהייתה מתחילה ממנו הייתה בולעת את
   * הלחיצה שלו — כלומר „✕” שלא סוגר.
   */
  it('לחיצה על כפתור הסגירה שבכותרת אינה גרירה', async () => {
    const harness = mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false } });
    await settle();

    const close = dialog().querySelector('.fa-close-btn');
    if (!close) throw new Error('אין כפתור סגירה');

    pointer('pointerdown', close, 300, 150);
    pointer('pointermove', window, 500, 400);
    await settle();

    expect(dialog().style.left).toBe('');
    close.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(harness.wrapper.emitted('close')).toHaveLength(1);
  });

  it('לחצן ימני אינו גורר', async () => {
    mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false } });
    await settle();

    const header = dialog().querySelector('.fa-header');
    if (!header) throw new Error('אין כותרת');

    header.dispatchEvent(
      new MouseEvent('pointerdown', { bubbles: true, clientX: 100, clientY: 200, button: 2 }),
    );
    pointer('pointermove', window, 400, 400);
    await settle();

    expect(dialog().style.left).toBe('');
  });

  /**
   * המיקום שנגרר שורד סגירה ופתיחה, כמו ב-Word: מי שהזיז את הדיאלוג הצידה כדי
   * לראות את הפסקה שמתחתיו לא רוצה למצוא אותו שוב מעליה בפעם הבאה.
   */
  it('המיקום נשמר בין פתיחות', async () => {
    const harness = mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false } });
    await settle();

    const header = dialog().querySelector('.fa-header');
    if (!header) throw new Error('אין כותרת');
    pointer('pointerdown', header, 100, 200);
    pointer('pointermove', window, 300, 400);
    pointer('pointerup', window, 300, 400);
    await settle();

    await harness.wrapper.setProps({ isOpen: false });
    await settle();
    await harness.wrapper.setProps({ isOpen: true });
    await settle();

    expect(dialog().style.left).toBe('200px');
    expect(dialog().style.top).toBe('200px');
  });
});

/**
 * חלון שקטן **בזמן שהדיאלוג סגור** הוא המקרה שאין לו תיקון עצמי: `v-if` מחק
 * את האלמנט, ואין מה למדוד בפתיחה הבאה — הדיאלוג היה נפתח מחוץ למסך, בלי
 * כותרת לתפוס בה. ההצמדה מחדש נשענת על המידות של הגרירה האחרונה.
 */
describe('הקטנת החלון אחרי גרירה', () => {
  it('מחזירה את הדיאלוג לתוך המסך', async () => {
    const width = window.innerWidth;
    const height = window.innerHeight;
    try {
      mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false } });
      await settle();

      const header = dialog().querySelector('.fa-header');
      if (!header) throw new Error('אין כותרת');
      pointer('pointerdown', header, 0, 0);
      pointer('pointermove', window, 600, 500);
      pointer('pointerup', window, 600, 500);
      await settle();
      expect(dialog().style.left).toBe('600px');

      Object.defineProperty(window, 'innerWidth', { value: 400, configurable: true });
      Object.defineProperty(window, 'innerHeight', { value: 300, configurable: true });
      window.dispatchEvent(new Event('resize'));
      await settle();

      expect(dialog().style.left).toBe('400px');
      expect(dialog().style.top).toBe('300px');
    } finally {
      Object.defineProperty(window, 'innerWidth', { value: width, configurable: true });
      Object.defineProperty(window, 'innerHeight', { value: height, configurable: true });
    }
  });
});

/**
 * העוגן המשותף: `top: 140px; inset-inline-start: 40px`. כל דיאלוג שנפתח עליו
 * הוא דיאלוג צף שמסתיר את המסמך, ולכן חייב להיות נגרר.
 */
describe('משפחת הדיאלוגים הצפים', () => {
  const anchored = readdirSync(PANELS)
    .filter((name) => name.endsWith('.vue'))
    .map((name) => ({ name, source: readFileSync(join(PANELS, name), 'utf8') }))
    .filter(({ source }) => source.includes('inset-inline-start: 40px'));

  it('המשפחה אינה ריקה — אחרת הסריקה מאשרת כלום', () => {
    expect(anchored.length).toBeGreaterThan(20);
  });

  it.each(anchored.map(({ name }) => name))('%s נגרר בכותרת שלו', (name) => {
    const source = anchored.find((file) => file.name === name)!.source;
    expect(source).toContain(`import { useDialogDrag } from '../../composables/dialog-drag';`);
    expect(source).toContain('const { dragStyle, startDialogDrag } = useDialogDrag();');
    expect(source).toContain(':style="dragStyle"');
    expect(source).toContain('dialog-drag-handle');
    expect(source).toContain('@pointerdown="startDialogDrag"');
  });
});

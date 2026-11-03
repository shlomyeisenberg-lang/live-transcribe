/**
 * ניהול סגנונות מהגלריה, מקצה הרצועה: תפריט הלחצן הימני על כרטיס, כפתור
 * „עוד”, הדיאלוג, ומה שנשלח למעטפת. הכתיבה ל-`styles.xml` והטעינה מחדש הן של
 * המעטפת — ה-harness רושם את העריכות (`styleEdits`), והן עצמן נבדקות ב-
 * tests/unit/docx-style-*.test.ts.
 */
import { describe, expect, it } from 'vitest';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import type { StyleGalleryState } from '../../src/engine/style-gallery';
import { autoUnmount, mountUi, settle, createSuperdocDouble } from './harness';

autoUnmount();

const GALLERY: StyleGalleryState = {
  items: [
    { id: 'Normal', label: 'רגיל', previewText: 'AaBbCc', previewStyle: {}, custom: false },
    { id: 'Heading1', label: 'כותרת 1', previewText: 'כותרת 1', previewStyle: {}, custom: false },
    { id: 'Mine', label: 'שלי', previewText: 'AaBbCc', previewStyle: {}, custom: true },
  ],
  activeId: 'Normal',
  fromDocument: true,
};

function mount() {
  return mountUi(HomeTab, { styleGallery: GALLERY });
}

function card(harness: ReturnType<typeof mount>, id: number) {
  return harness.wrapper.findAll('.style-card')[id];
}

/** התפריט והדיאלוג מטולפרטים ל-body. */
const menu = () => document.querySelector<HTMLElement>('[data-testid="style-menu"]');
const dialog = () => document.querySelector<HTMLElement>('[data-testid="style-dialog"]');

function menuItem(action: string): HTMLButtonElement {
  const button = menu()?.querySelector<HTMLButtonElement>(`[data-action="${action}"]`);
  if (!button) throw new Error(`אין פריט ${action}`);
  return button;
}

async function rightClick(harness: ReturnType<typeof mount>, index: number): Promise<void> {
  await card(harness, index).trigger('contextmenu', { clientX: 500, clientY: 100 });
  await settle();
}

async function primary(): Promise<void> {
  dialog()?.querySelector<HTMLButtonElement>('.sd-btn-primary')?.click();
  await settle();
  await settle();
}

describe('תפריט הסגנון בגלריה', () => {
  it('שינוי שם זמין רק לסגנון מותאם, ומחיקה אינה מוצגת', async () => {
    const harness = mount();
    await settle();

    await rightClick(harness, 1);
    expect(menuItem('modify').disabled).toBe(false);
    expect(menuItem('rename').disabled).toBe(true);
    expect(menu()?.querySelector('[data-action=delete]')).toBeNull();

    document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    await settle();
    expect(menu()).toBeNull();

    await rightClick(harness, 2);
    expect(menuItem('rename').disabled).toBe(false);
    expect(menu()?.querySelector('[data-action=delete]')).toBeNull();
  });

  it('„שנה” פותח את הדיאלוג על העיצוב האפקטיבי, ושולח רק את מה ששונה', async () => {
    const harness = mount();
    await settle();
    await rightClick(harness, 1);
    menuItem('modify').click();
    await settle();
    await settle();

    const root = dialog();
    expect(root).not.toBeNull();
    // כותרת 1 של הגיליון: 16 נקודות ומודגש, כולל מה שהיא יורשת מ„רגיל”.
    const size = root!.querySelector<HTMLInputElement>('input[aria-label="גודל, בנקודות"]')!;
    expect(size.value).toBe('16');
    const bold = root!.querySelector<HTMLButtonElement>('button[aria-label="מודגש"]')!;
    expect(bold.getAttribute('aria-pressed')).toBe('true');

    size.value = '20';
    size.dispatchEvent(new Event('input'));
    root!.querySelector<HTMLButtonElement>('button[aria-label="מרכז"]')!.click();
    await settle();
    await primary();

    expect(harness.styleEdits).toHaveLength(1);
    expect(harness.styleEdits[0].edit).toEqual({
      kind: 'modify',
      id: 'Heading1',
      basedOn: undefined,
      next: undefined,
      format: { fontSizePt: 20, alignment: 'center' },
    });
    expect(dialog()).toBeNull();
  });

  it('„אישור” בלי שינוי אינו שולח דבר', async () => {
    const harness = mount();
    await settle();
    await rightClick(harness, 1);
    menuItem('modify').click();
    await settle();
    await settle();
    await primary();
    expect(harness.styleEdits).toHaveLength(0);
    expect(dialog()).toBeNull();
  });

  it('„צור סגנון” מכפתור „עוד” יוצר סגנון ומחיל אותו על הבחירה', async () => {
    const harness = mount();
    await settle();
    await harness.wrapper.find('[data-testid="style-gallery-more"]').trigger('click');
    await settle();
    menuItem('create').click();
    await settle();
    await settle();

    const name = dialog()!.querySelector<HTMLInputElement>('#sd-name')!;
    expect(name.value).toBe('סגנון1');
    name.value = 'ציטוט מקור';
    name.dispatchEvent(new Event('input'));
    await settle();
    await primary();

    expect(harness.styleEdits).toHaveLength(1);
    const { edit, options } = harness.styleEdits[0];
    expect(edit).toMatchObject({ kind: 'create', name: 'ציטוט מקור', type: 'paragraph', basedOn: 'Normal' });
    expect(options).toEqual({ applyToSelection: true });
  });

  it('שינוי שם שולח את השם החדש', async () => {
    const harness = mount();
    await settle();
    await rightClick(harness, 2);
    menuItem('rename').click();
    await settle();
    await settle();

    const name = dialog()!.querySelector<HTMLInputElement>('#sd-name')!;
    expect(name.value).toBe('שלי');
    name.value = 'שלי החדש';
    name.dispatchEvent(new Event('input'));
    await settle();
    await primary();
    expect(harness.styleEdits.map((entry) => entry.edit)).toEqual([{ kind: 'rename', id: 'Mine', name: 'שלי החדש' }]);
  });

  it('בלי מסמך פתוח — ניהול הסגנונות מנוטרל', async () => {
    const harness = mountUi(HomeTab, { styleGallery: GALLERY, superdoc: null });
    await settle();
    expect(harness.wrapper.find('[data-testid="style-gallery-more"]').attributes('disabled')).toBeDefined();
  });

  it('רמת חלוקה נכתבת לסגנון — מה ש-Word קורא כדי לזהות כותרת', async () => {
    const harness = mount();
    await settle();
    await rightClick(harness, 2);
    menuItem('modify').click();
    await settle();
    await settle();

    const select = dialog()!.querySelector<HTMLSelectElement>('#sd-outline')!;
    // „שלי” אינו מצהיר רמה — גוף טקסט.
    expect(select.value).toBe('9');
    select.value = '0';
    select.dispatchEvent(new Event('change'));
    await settle();
    await primary();

    expect(harness.styleEdits.map((entry) => entry.edit)).toEqual([
      { kind: 'modify', id: 'Mine', basedOn: undefined, next: undefined, format: { outlineLevel: 0 } },
    ]);
  });

  /** לוכד הצירוף בדיאלוג: לחיצה ואז הקשה. */
  async function recordCombo(init: KeyboardEventInit): Promise<void> {
    const recorder = dialog()!.querySelector<HTMLButtonElement>('[data-testid="style-shortcut-recorder"]')!;
    recorder.click();
    await settle();
    // ההקשה הולכת לאלמנט **שבמיקוד**, כמו במקלדת אמיתית — לא ישירות ללוכד.
    // שליחה ישירה הסתירה את הבאג שבו לחיצה בעכבר לא העבירה אליו מיקוד.
    expect(document.activeElement).toBe(recorder);
    (document.activeElement as HTMLElement).dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }),
    );
    await settle();
  }

  it('קיצור מקשים בלבד נשמר בלי לגעת במסמך', async () => {
    const harness = mount();
    await settle();
    await rightClick(harness, 1);
    menuItem('modify').click();
    await settle();
    await settle();

    await recordCombo({ key: 'q', code: 'KeyQ', ctrlKey: true, altKey: true, shiftKey: true });
    expect(dialog()!.textContent).toContain('Ctrl');
    await primary();

    expect(harness.styleEdits).toHaveLength(0);
    expect(harness.shortcuts.list).toEqual([
      expect.objectContaining({
        kind: 'apply-style',
        style: { id: 'Heading1', name: 'כותרת 1' },
        combo: { code: 'KeyQ', ctrl: true, shift: true, alt: true },
      }),
    ]);
    expect(dialog()).toBeNull();
  });

  it('צירוף תפוס חוסם את „אישור” ואינו נשמר', async () => {
    const harness = mountUi(HomeTab, {
      styleGallery: GALLERY,
      customShortcuts: [
        { id: 'cs-1', name: 'אחר', kind: 'apply-style', combo: { code: 'KeyQ', ctrl: true, shift: true, alt: true }, style: { id: 'Mine', name: 'שלי' } },
      ],
    });
    await settle();
    await rightClick(harness, 1);
    menuItem('modify').click();
    await settle();
    await settle();

    await recordCombo({ key: 'q', code: 'KeyQ', ctrlKey: true, altKey: true, shiftKey: true });
    const button = dialog()!.querySelector<HTMLButtonElement>('.sd-btn-primary')!;
    expect(button.disabled).toBe(true);
    expect(dialog()!.textContent).toContain('הצירוף הזה משמש כבר');
    expect(harness.shortcuts.list).toHaveLength(1);
  });

  it('שינוי שם מעדכן את הקיצור שלו', async () => {
    const combo = { code: 'KeyQ', ctrl: true, shift: true, alt: true };
    const harness = mountUi(HomeTab, {
      styleGallery: GALLERY,
      customShortcuts: [{ id: 'cs-1', name: 'שלי', kind: 'apply-style', combo, style: { id: 'Mine', name: 'שלי' } }],
    });
    await settle();

    await rightClick(harness, 2);
    menuItem('rename').click();
    await settle();
    await settle();
    const name = dialog()!.querySelector<HTMLInputElement>('#sd-name')!;
    name.value = 'שלי החדש';
    name.dispatchEvent(new Event('input'));
    await settle();
    await primary();
    expect(harness.shortcuts.list[0]).toMatchObject({ name: 'שלי החדש', style: { id: 'Mine', name: 'שלי החדש' } });

  });
  it('a dialog closes when the document changes and cannot mutate the new document', async () => {
    const harness = mount();
    await settle();
    await rightClick(harness, 1);
    menuItem('modify').click();
    await settle(); await settle();
    expect(dialog()).not.toBeNull();
    await harness.setSuperdoc(createSuperdocDouble());
    expect(dialog()).toBeNull();
    expect(harness.styleEdits).toEqual([]);
  });

});

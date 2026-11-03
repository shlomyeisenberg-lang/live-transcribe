/**
 * דיאלוג „ניהול קיצורים” — הטופס שמצמיד צירוף מקשים לערכת עיצוב.
 *
 * מה שרק הרכבה יכולה להוכיח כאן, ושהמודולים הטהורים אינם מכסים:
 *
 * 1. **הלוכד לוכד.** הצירוף אינו מוקלד אלא נלכד מהקשה, ולכן „האם `event.code`
 *    אכן נשמר” היא שאלה על מאזין ב-DOM ולא על פונקציה. לוכד שאינו מאזין הוא
 *    טופס שאי אפשר למלא.
 * 2. **הלוכד אינו גוזל את המקלדת.** `Escape` בזמן לכידה מבטל אותה ואינו סוגר
 *    את הדיאלוג — מי שנכנס בטעות צריך דרך לצאת בלי לאבד את הטופס — ו-`Tab`
 *    נשאר מקש ניווט. שני אלה הם `stopPropagation`/`preventDefault` על אירוע
 *    אמיתי, כלומר בדיוק מה שנשמט בשקט.
 * 3. **האימות שהכפתור מציית לו הוא אותו אימות שהשמירה אוכפת.** אם הם נפרדים,
 *    נוצר כפתור פעיל ששמירתו מסורבת.
 */
import { describe, expect, it } from 'vitest';
import { DOMWrapper } from '@vue/test-utils';
import ShortcutManagerDialog from '../../src/ui/panels/ShortcutManagerDialog.vue';
import { autoUnmount, mountUi, settle } from './harness';
import { shortcutTextOwners } from '../../src/ui/shortcuts/combo';
import type { CustomShortcut } from '../../src/ui/shortcuts/custom-shortcuts';

autoUnmount();

const EXISTING: CustomShortcut = {
  id: 'cs-1',
  name: 'כותרת קטע',
  kind: 'format-preset',
  combo: { code: 'KeyH', ctrl: true, shift: false, alt: true },
  preset: { fontFamily: 'David', fontSizePt: 16 },
};

function dialog(): HTMLElement {
  const element = document.querySelector<HTMLElement>('.shortmgr-dialog');
  if (!element) throw new Error('הדיאלוג אינו בגוף הדף');
  return element;
}

function recorder(): HTMLButtonElement {
  const element = dialog().querySelector<HTMLButtonElement>('.sm-recorder');
  if (!element) throw new Error('אין לוכד צירוף בדיאלוג');
  return element;
}

function footerButton(label: string): DOMWrapper<Element> {
  const buttons = [...dialog().querySelectorAll('.sm-footer .sm-btn')];
  const found = buttons.find((button) => button.textContent?.trim() === label);
  if (!found) throw new Error(`לא נמצא הכפתור „${label}” בדיאלוג`);
  return new DOMWrapper(found);
}

/** הודעת השגיאה או ההסבר שמוצגים בטופס. */
function noteText(): string {
  return dialog().querySelector('.sm-error, .sm-note')?.textContent?.trim() ?? '';
}

/**
 * הקשה על הלוכד. `cancelable` כדי ש-`defaultPrevented` יהיה מדיד — זה מה
 * שקובע אם ההקשה נגזלה מהדפדפן ומהמנתב.
 */
function pressOnRecorder(init: KeyboardEventInit & { code: string }): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { cancelable: true, bubbles: true, ...init });
  recorder().dispatchEvent(event);
  return event;
}

async function open(props: Record<string, unknown> = {}) {
  const harness = mountUi(ShortcutManagerDialog, {
    props: { isOpen: true, list: [], ...props },
  });
  await settle();
  return harness;
}

describe('דיאלוג ניהול הקיצורים', () => {
  it('נפתח מטולפרט לגוף הדף, ומצהיר על עצמו כמודאלי', async () => {
    // `aria-modal` אינו קוסמטי: `isModalDialogOpen` ב-App.vue מאתר את
    // הדיאלוג דרכו, וזה מה שחוסם קיצורים בזמן שהוא פתוח.
    await open();
    expect(dialog().getAttribute('aria-modal')).toBe('true');
    expect(dialog().getAttribute('role')).toBe('dialog');
  });

  it('הרשימה מציגה שם, צירוף וסיכום ערכה', async () => {
    await open({ list: [EXISTING] });
    const item = dialog().querySelector('.sm-item');
    expect(item?.textContent).toContain('כותרת קטע');
    expect(item?.textContent).toContain('Ctrl+Alt+H');
    expect(item?.textContent).toContain('David');
  });

  it('רשימה ריקה אומרת מה לעשות, ולא נראית שבורה', async () => {
    await open();
    expect(dialog().querySelector('.sm-empty')?.textContent).toContain('קיצור חדש');
  });

  it('**הלוכד שומר את המקש הפיזי**, גם בפריסה עברית', async () => {
    // `key: 'ל'` הוא מה שהמקש KeyK מפיק בעברית. תווית שנגזרת מה-`key` הייתה
    // מציגה „Ctrl+Alt+ל”, וההתאמה בזמן ריצה הייתה מחפשת מקש אחר.
    await open();
    recorder().click();
    await settle();

    pressOnRecorder({ code: 'KeyK', key: 'ל', ctrlKey: true, altKey: true });
    await settle();

    expect(recorder().textContent?.trim()).toBe('Ctrl+Alt+K');
  });

  it('ההקשה בזמן לכידה אינה מגיעה לדפדפן ולא למנתב', async () => {
    // בלי `preventDefault` הלכידה של `Ctrl+O` הייתה גם פותחת קובץ.
    await open();
    recorder().click();
    await settle();

    const event = pressOnRecorder({ code: 'KeyO', ctrlKey: true, altKey: true });
    expect(event.defaultPrevented).toBe(true);
  });

  it('מודיפייר לבדו אינו סוגר את הלכידה', async () => {
    // אחרת הלוכד היה „נסגר” על Ctrl ברגע שהמשתמש מתחיל להקיש את הצירוף,
    // והמשתמש היה רואה שדה שלא נקלט בו דבר.
    await open();
    recorder().click();
    await settle();

    pressOnRecorder({ code: 'ControlLeft', key: 'Control', ctrlKey: true });
    await settle();
    expect(recorder().classList.contains('sm-recorder--armed')).toBe(true);

    pressOnRecorder({ code: 'KeyK', ctrlKey: true, altKey: true });
    await settle();
    expect(recorder().classList.contains('sm-recorder--armed')).toBe(false);
  });

  it('`Escape` בזמן לכידה מבטל אותה — ואינו סוגר את הדיאלוג', async () => {
    const harness = await open();
    recorder().click();
    await settle();

    pressOnRecorder({ code: 'Escape', key: 'Escape' });
    await settle();

    expect(recorder().classList.contains('sm-recorder--armed')).toBe(false);
    expect(harness.wrapper.emitted('close')).toBeUndefined();
  });

  it('`Tab` נשאר מקש ניווט גם בזמן לכידה', async () => {
    // בליעתו הייתה משאירה את המשתמש נעול בלוכד בלי דרך לצאת במקלדת.
    await open();
    recorder().click();
    await settle();

    const event = pressOnRecorder({ code: 'Tab', key: 'Tab' });
    expect(event.defaultPrevented).toBe(false);
  });

  it('**דיאלוג שנפתח אינו מציג שגיאה** — הוא מציג מה הקיצור עושה', async () => {
    // כלל מתועד של הבית (LinkDialog.vue): „השגיאה מוצגת רק אחרי שהמשתמש
    // הקליד משהו. הצגה על שדה ריק פירושה דיאלוג שנפתח עם הודעת שגיאה, בלי
    // שאיש עשה כלום.” הדיאלוג הזה הפר אותו — הוא נפתח עם „יש לתת שם לקיצור”
    // באדום ובתוך `role="alert"`.
    await open();
    expect(dialog().querySelector('.sm-error')).toBeNull();
    expect(noteText()).toContain('לחיצה נוספת מחזירה');
    // והכפתור עדיין נעול: „אין שגיאה” אינו „אפשר לשמור”.
    expect(footerButton('הוספה').attributes('disabled')).toBeDefined();
  });

  it('**צירוף תפוס מדווח גם לפני שיש שם**, ובאדום', async () => {
    // מדוד: מי שלוכד לפני שהוא נותן שם ראה „יש לתת שם לקיצור” בעוד הלוכד
    // מציג „Ctrl+S” — צירוף תפוס בלי שום סימן שהוא תפוס.
    await open();
    recorder().click();
    await settle();
    pressOnRecorder({ code: 'KeyS', ctrlKey: true });
    await settle();

    const error = dialog().querySelector('.sm-error');
    expect(error).not.toBeNull();
    expect(error!.getAttribute('role')).toBe('alert');
    expect(error!.textContent).toContain('פעולה מובנית');
  });

  it('טופס חסר אינו נשמר, וההנחיה השקטה אומרת מה חסר', async () => {
    const harness = await open();
    const save = footerButton('הוספה');
    expect(save.attributes('disabled')).toBeDefined();

    await new DOMWrapper(dialog().querySelector('#sm-name')!).setValue('כותרת');
    await settle();
    // אחרי שהמשתמש התחיל — ההנחיה מוצגת, אך אינה שגיאה.
    expect(noteText()).toContain('צירוף');
    expect(dialog().querySelector('.sm-error')).toBeNull();

    recorder().click();
    await settle();
    pressOnRecorder({ code: 'KeyK', ctrlKey: true, altKey: true });
    await settle();
    // צירוף יש, ערכה עוד אין.
    expect(noteText()).toContain('עיצוב');
    expect(footerButton('הוספה').attributes('disabled')).toBeDefined();
    expect(harness.wrapper.emitted('save')).toBeUndefined();
  });

  it('טופס שלם נשמר, והכוונה שנפלטת היא מה שהוקלד', async () => {
    const harness = await open();

    await new DOMWrapper(dialog().querySelector('#sm-name')!).setValue('כותרת');
    recorder().click();
    await settle();
    pressOnRecorder({ code: 'KeyK', ctrlKey: true, altKey: true });
    await new DOMWrapper(dialog().querySelector('#sm-size')!).setValue('18');
    await settle();

    expect(noteText()).not.toContain('יש ל');
    await footerButton('הוספה').trigger('click');

    const saved = harness.wrapper.emitted('save');
    expect(saved).toHaveLength(1);
    expect(saved![0]![0]).toEqual({
      name: 'כותרת',
      combo: { code: 'KeyK', ctrl: true, shift: false, alt: true },
      preset: { fontSizePt: 18 },
    });
  });

  it('**צירוף של פעולה מובנית נדחה**, בשמה', async () => {
    await open();
    await new DOMWrapper(dialog().querySelector('#sm-name')!).setValue('כותרת');
    await new DOMWrapper(dialog().querySelector('#sm-size')!).setValue('18');
    recorder().click();
    await settle();
    pressOnRecorder({ code: 'KeyS', ctrlKey: true });
    await settle();

    expect(noteText()).toContain('פעולה מובנית');
    expect(footerButton('הוספה').attributes('disabled')).toBeDefined();
  });

  it('צירוף שמאקרו מחזיק נדחה בשם המאקרו', async () => {
    await open({ taken: shortcutTextOwners([{ name: 'בסד', shortcut: 'Ctrl+Alt+K' }]) });
    await new DOMWrapper(dialog().querySelector('#sm-name')!).setValue('כותרת');
    await new DOMWrapper(dialog().querySelector('#sm-size')!).setValue('18');
    recorder().click();
    await settle();
    pressOnRecorder({ code: 'KeyK', ctrlKey: true, altKey: true });
    await settle();

    expect(noteText()).toContain('מאקרו „בסד”');
  });

  it('בחירת רשומה טוענת אותה לטופס, והשמירה נושאת את המזהה', async () => {
    const harness = await open({ list: [EXISTING] });
    await new DOMWrapper(dialog().querySelector('.sm-item')!).trigger('click');
    await settle();

    expect(dialog().querySelector<HTMLInputElement>('#sm-name')!.value).toBe('כותרת קטע');
    expect(recorder().textContent?.trim()).toBe('Ctrl+Alt+H');
    expect(dialog().querySelector<HTMLInputElement>('#sm-size')!.value).toBe('16');
    // הצירוף של הרשומה שנערכת אינו מתנגש בעצמה.
    expect(noteText()).not.toContain('משמש כבר');

    await footerButton('שמירה').trigger('click');
    expect((harness.wrapper.emitted('save')![0]![0] as { id?: string }).id).toBe('cs-1');
  });

  it('„מחיקה” פעילה רק על רשומה שנבחרה', async () => {
    const harness = await open({ list: [EXISTING] });
    expect(footerButton('מחיקה').attributes('disabled')).toBeDefined();

    await new DOMWrapper(dialog().querySelector('.sm-item')!).trigger('click');
    await settle();
    await footerButton('מחיקה').trigger('click');

    expect(harness.wrapper.emitted('remove')![0]).toEqual(['cs-1']);
  });

  it('„ללא שינוי” הוא באמת ללא שינוי — שדה שלא נגעו בו אינו נכנס לערכה', async () => {
    // זו ההבחנה שכל מודל הערכה עומד עליה: ערכה שמגדילה גופן אינה אמורה
    // למחוק בדרך את הצבע.
    const harness = await open();
    await new DOMWrapper(dialog().querySelector('#sm-name')!).setValue('רק גודל');
    recorder().click();
    await settle();
    pressOnRecorder({ code: 'KeyM', ctrlKey: true, altKey: true });
    await new DOMWrapper(dialog().querySelector('#sm-size')!).setValue('11');
    await settle();
    await footerButton('הוספה').trigger('click');

    const draft = harness.wrapper.emitted('save')![0]![0] as { preset: Record<string, unknown> };
    expect(Object.keys(draft.preset)).toEqual(['fontSizePt']);
  });

  it('מתג נכנס לערכה בשני הכיוונים — „כן” ו„לא”', async () => {
    // ערכה שאינה מכבה מודגש נראית שבורה כשמחילים אותה בתוך טקסט מודגש,
    // ולכן „לא” הוא ערך ולא היעדר ערך.
    const harness = await open();
    await new DOMWrapper(dialog().querySelector('#sm-name')!).setValue('גוף');
    recorder().click();
    await settle();
    pressOnRecorder({ code: 'KeyM', ctrlKey: true, altKey: true });

    const selects = [...dialog().querySelectorAll('.sm-toggles select')];
    await new DOMWrapper(selects[0]!).setValue('on');
    await new DOMWrapper(selects[1]!).setValue('off');
    await settle();
    await footerButton('הוספה').trigger('click');

    const draft = harness.wrapper.emitted('save')![0]![0] as { preset: Record<string, unknown> };
    expect(draft.preset).toEqual({ bold: true, italic: false });
  });

  it('סגירה ופתיחה מנקות את הטופס', async () => {
    // טופס שנשאר מלא היה מציע בפתיחה הבאה לשמור רשומה שהמשתמש נטש.
    const harness = await open();
    await new DOMWrapper(dialog().querySelector('#sm-name')!).setValue('נטוש');
    await settle();

    await harness.wrapper.setProps({ isOpen: false });
    await settle();
    await harness.wrapper.setProps({ isOpen: true });
    await settle();

    expect(dialog().querySelector<HTMLInputElement>('#sm-name')!.value).toBe('');
    expect(recorder().textContent?.trim()).toContain('לחצו כאן');
  });
});

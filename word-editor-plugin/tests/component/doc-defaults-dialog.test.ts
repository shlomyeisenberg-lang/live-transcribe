/**
 * דיאלוג „ברירות מחדל למסמך" — גופן וגודל ברירת המחדל של המסמך כולו.
 *
 * הבאג: `<input id="dd-size" type="number" v-model="fontSize">` — Vue 3
 * ממיר את הערך שנקלט מ-`v-model` על `type="number"` **למספר**, גם בלי
 * modifier מפורש (`.number`), ברגע שהמחרוזת נפרסת בהצלחה כמספר
 * (`runtime-dom`: `castToNumber = number || vnode.props?.type === 'number'`).
 * `fontSize` הוקלד `ref('')`, כלומר `string` בעיני TypeScript, אבל בזמן
 * ריצה קיבל `number` ברגע שהוקלד תו ראשון תקין. `parseSize` קרא ל-
 * `value.trim()` על מה שהגיע — וקרס עם `TypeError: f.trim is not a
 * function` בתוך computed, מה שהעיף את כל הדיאלוג מהעץ (שום דבר לא
 * נכתב ל-styles.xml). ראו את אותו כשל בדיוק, כבר מתועד ומתוקן, ב-
 * FontAdvancedDialog.vue (`asText`).
 *
 * הבדיקה כאן מקלידה ערך אמיתי לשדה (לא מזריקה מספר ידנית) כדי לשחזר בדיוק
 * את ההמרה שקורה בדפדפן.
 */
import { describe, expect, it } from 'vitest';
import { DOMWrapper } from '@vue/test-utils';
import DocDefaultsDialog from '../../src/ui/panels/DocDefaultsDialog.vue';
import { autoUnmount, mountUi, settle } from './harness';

autoUnmount();

function teleported(selector: string): DOMWrapper<Element> {
  const element = document.querySelector(selector);
  if (!element) throw new Error(`לא נמצא ${selector} בגוף הדף`);
  return new DOMWrapper(element);
}

function footerButton(label: string): DOMWrapper<Element> {
  const buttons = [...document.querySelectorAll('.docdef-dialog .dd-footer .dd-btn')];
  const found = buttons.find((button) => button.textContent?.trim() === label);
  if (!found) throw new Error(`לא נמצא הכפתור „${label}" בדיאלוג`);
  return new DOMWrapper(found);
}

/**
 * בחירת גופן בבורר — כפי שהמשתמש בוחר: פתיחה בעכבר ולחיצה על שורה.
 *
 * שם הגופן היה תיבת טקסט חופשי, ולכן הבדיקות כאן היו `setValue`. הבורר אינו
 * מחיל טקסט שהוקלד ולא אושר (`onBlur` מוותר בכוונה), ולכן „בחירה” חייבת
 * להיות בחירה. `'ללא שינוי'` היא השורה שמחזירה ל„לא ייגע”.
 */
async function chooseFont(label: string): Promise<void> {
  const input = document.querySelector<HTMLInputElement>('.docdef-dialog .dd-combo input');
  if (!input) throw new Error('אין בורר גופן בדיאלוג');
  input.dispatchEvent(new FocusEvent('focus'));
  await settle();

  const row = [...document.querySelectorAll('.docdef-dialog .ribbon-combo-option')].find(
    (option) => option.textContent?.trim() === label,
  );
  if (!row) throw new Error(`אין שורה „${label}” ברשימת הגופנים`);
  await new DOMWrapper(row).trigger('pointerdown');
  await settle();
}

describe('DocDefaultsDialog', () => {
  it('הקלדת תו בודד בשדה הגודל אינה מפילה את הדיאלוג', async () => {
    mountUi(DocDefaultsDialog, { props: { isOpen: true, busy: false, currentSizePt: null } });
    await settle();

    // הקלדה אמיתית לשדה — `type="number"`, בדיוק כמו בדפדפן — ולא הזרקת
    // מספר ל-ref ידנית. זה מה ששחזר את ה-TypeError בפועל.
    await teleported('#dd-size').setValue('2');
    await settle();

    // לפני התיקון: `f.trim is not a function` נזרק מתוך computed, והדיאלוג
    // נעלם מה-DOM לגמרי.
    expect(document.querySelector('.docdef-dialog')).not.toBeNull();
    expect(document.querySelector('.dd-error')).toBeNull();
  });

  it('הערך שהוקלד נקרא נכון ונשלח כמספר ב-submit', async () => {
    const harness = mountUi(DocDefaultsDialog, {
      props: { isOpen: true, busy: false, currentSizePt: null },
    });
    await settle();

    await teleported('#dd-size').setValue('24');
    await settle();
    await footerButton('אישור').trigger('click');
    await settle();

    const emissions = harness.wrapper.emitted('submit');
    expect(emissions).toHaveLength(1);
    expect(emissions?.[0]?.[0]).toEqual({ fontSizePt: 24 });
  });

  it('ערך לא-תקין (0) מציג שגיאה ומנטרל את „אישור", בלי לקרוס', async () => {
    mountUi(DocDefaultsDialog, { props: { isOpen: true, busy: false, currentSizePt: null } });
    await settle();

    await teleported('#dd-size').setValue('0');
    await settle();

    expect(document.querySelector('.docdef-dialog')).not.toBeNull();
    expect(teleported('.dd-error').exists()).toBe(true);
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
  });

  it('ריקון השדה אחרי הקלדה חוזר ל"ללא שינוי" — אינו נשלח ב-patch', async () => {
    const harness = mountUi(DocDefaultsDialog, {
      props: { isOpen: true, busy: false, currentSizePt: null },
    });
    await settle();

    await teleported('#dd-size').setValue('12');
    await teleported('#dd-size').setValue('');
    await settle();

    expect(document.querySelector('.docdef-dialog')).not.toBeNull();
    await chooseFont('Assistant');
    await footerButton('אישור').trigger('click');
    await settle();

    const emissions = harness.wrapper.emitted('submit');
    expect(emissions?.[0]?.[0]).toEqual({ fontFamily: 'Assistant' });
  });

  /**
   * הכפתור היה פעיל על דיאלוג ריק, ולחיצה עליו סגרה בלי לעשות דבר — „אישור”
   * שמתנהג כמו „ביטול”. נמדד בשער הפריסה: `אישור disabled=false`.
   */
  it('שני שדות ריקים: „אישור” נעול, ובלי הודעת שגיאה', async () => {
    mountUi(DocDefaultsDialog, {
      props: { isOpen: true, busy: false, currentSizePt: 12 },
    });
    await settle();

    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    // „לא הוקלד כלום” אינו שגיאה, ולכן אינו ראוי להודעה אדומה. נבדק ישירות
    // מול ה-DOM: העוזר `teleported` זורק כשאין אלמנט, ולכן אינו יכול לבטא היעדר.
    expect(document.querySelector('.dd-error')).toBeNull();
  });

  it('מילוי שדה אחד בלבד משחרר את „אישור”', async () => {
    mountUi(DocDefaultsDialog, {
      props: { isOpen: true, busy: false, currentSizePt: 12 },
    });
    await settle();

    await chooseFont('David');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeUndefined();

    // וריקון חוזר נועל אותו שוב — המצב נגזר מהשדות, ולא נקבע פעם אחת.
    await chooseFont('ללא שינוי');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
  });

  /**
   * טקסט שהוקלד ולא אושר אינו „מילוי”. זה מה שהחליף את המקרה של „רווחים
   * בלבד”: בתיבת טקסט חופשי אפשר היה להשאיר רווחים בשדה, ובבורר אי אפשר
   * להשאיר בו דבר שלא נבחר — `onBlur` מוותר, בדיוק כדי לא להחיל בטעות.
   */
  it('הקלדה בלי בחירה אינה „מילוי” — הכפתור נשאר נעול', async () => {
    mountUi(DocDefaultsDialog, {
      props: { isOpen: true, busy: false, currentSizePt: 12 },
    });
    await settle();

    const input = document.querySelector<HTMLInputElement>('.docdef-dialog .dd-combo input');
    if (!input) throw new Error('אין בורר גופן בדיאלוג');
    input.dispatchEvent(new FocusEvent('focus'));
    input.value = 'Dav';
    input.dispatchEvent(new Event('input'));
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();

    input.dispatchEvent(new FocusEvent('blur'));
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
  });
});

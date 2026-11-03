/**
 * דיאלוג „גופן מתקדם", וההגעה אליו מלשונית „בית".
 *
 * למה קובץ ייעודי: הסורק הגנרי לא בודק את תוכן ה-payload. מה שנבדק כאן:
 * רק שדות שמולאו יוצאים למנוע („ללא שינוי" = לא נשלח), אזהרת ה-vanish
 * מופיעה בזמן, והכשל של המנוע מדווח בעברית. הדיאלוג נבדק דרך ה-DOM —
 * Teleport לגוף הדף, כמו BookmarkDialog.
 */
import { describe, expect, it } from 'vitest';
import { DOMWrapper } from '@vue/test-utils';
import FontAdvancedDialog from '../../src/ui/panels/FontAdvancedDialog.vue';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import { autoUnmount, createSuperdocDouble, mountUi, settle, tipSelector } from './harness';

autoUnmount();

function teleported(selector: string): DOMWrapper<Element> {
  const element = document.querySelector(selector);
  if (!element) throw new Error(`לא נמצא ${selector} בגוף הדף`);
  return new DOMWrapper(element);
}

function footerButton(label: string): DOMWrapper<Element> {
  const buttons = [...document.querySelectorAll('.fontadv-dialog .fa-footer .fa-btn')];
  const found = buttons.find((button) => button.textContent?.trim() === label);
  if (!found) throw new Error(`לא נמצא הכפתור „${label}" בדיאלוג`);
  return new DOMWrapper(found);
}

/** כפתור אפקט לפי התווית שעליו — הפקד שהחליף את בוררי „ללא שינוי/כן/לא". */
function toggle(label: string): DOMWrapper<Element> {
  const buttons = [...document.querySelectorAll('.fontadv-dialog .tri-toggle')];
  const found = buttons.find((button) => button.querySelector('.tri-label')?.textContent?.trim() === label);
  if (!found) throw new Error(`לא נמצא כפתור האפקט „${label}" בדיאלוג`);
  return new DOMWrapper(found);
}

/** מחזור: לחיצה אחת = „דלוק", שתיים = „כבוי", שלוש = חזרה ל„ללא שינוי". */
async function clickToggle(label: string, times = 1): Promise<void> {
  for (let i = 0; i < times; i += 1) await toggle(label).trigger('click');
  await settle();
}

const ADVANCED_BUTTON = tipSelector('מתקדם');

/**
 * הרכבה של הדיאלוג **על טקסט מסומן** — התנאי המוקדם שלו.
 *
 * `format.apply` מקבל SelectionTarget בלבד, ולכן דיאלוג שנפתח על סמן מכווץ
 * נועל את „אישור" ואומר למה (ראו „בלי בחירה" למטה). זה המצב הרגיל של כל
 * הבדיקות כאן — מי שפתח את „גופן מתקדם" סימן קודם — ולכן הוא כתוב פעם אחת
 * בעזר, ולא בכל הרכבה. הבחירה מוצהרת ואינה ברירת מחדל שקטה של הדמה.
 */
function mountDialog(props: { isOpen: boolean; busy: boolean }, hasRange = true) {
  return mountUi(FontAdvancedDialog, {
    props,
    superdoc: createSuperdocDouble({ selection: { hasRange, text: hasRange ? 'טקסט מסומן' : '' } }),
  });
}

describe('FontAdvancedDialog (בדיד)', () => {
  it('סגור אינו מרונדר בכלל', () => {
    mountDialog({ isOpen: false, busy: false });
    expect(document.querySelector('.fontadv-dialog')).toBeNull();
  });

  it('פתיחה ממקדת את שורש הדיאלוג', async () => {
    const harness = mountDialog({ isOpen: false, busy: false });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();

    expect(document.activeElement).toBe(document.querySelector('.fontadv-dialog'));
  });

  it('Escape סוגר', async () => {
    const harness = mountDialog({ isOpen: true, busy: false });
    await settle();

    await teleported('.fontadv-dialog').trigger('keydown.esc');
    expect(harness.wrapper.emitted('close')).toHaveLength(1);
  });

  /**
   * שדה מלא לפני הבדיקה, ובכוונה: על דיאלוג ריק „אישור” נעול ממילא מפני שאין
   * מה להחיל, ובלי המילוי הבדיקה הייתה עוברת גם אילו `busy` לא היה נבדק כלל —
   * כלומר השומר היחיד מפני שליחה כפולה היה נשאר בלי כיסוי.
   */
  it('busy מנטרל „אישור" גם כשיש מה להחיל, ומשאיר „ביטול" חי', async () => {
    const harness = mountDialog({ isOpen: true, busy: false });
    await settle();

    await teleported('#fa-scale').setValue('120');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeUndefined();

    await harness.wrapper.setProps({ busy: true });
    await settle();

    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(footerButton('ביטול').attributes('disabled')).toBeUndefined();
  });

  it('שדות לא-ממולאים אינם נשלחים — רק מה שהמשתמש מילא', async () => {
    const harness = mountDialog({ isOpen: true, busy: false });
    await settle();

    await teleported('#fa-scale').setValue('150');
    await clickToggle('קו חוצה כפול');
    await settle();
    await footerButton('אישור').trigger('click');
    await settle();

    const emissions = harness.wrapper.emitted('submit');
    expect(emissions).toHaveLength(1);
    expect(emissions?.[0]?.[0]).toEqual({ charScale: 150, dstrike: true });
  });

  /**
   * שלושת המצבים הם ההבטחה המרכזית של הפקד: „כבוי" חייב להיות **נבדל**
   * מ„לא נגעתי", אחרת אישור על דיאלוג שנפתח ונסגר היה מסיר אפקטים מהטקסט
   * המסומן. ההנמקה ב-ui/panels/common/TriToggle.vue.
   */
  it('הכפתור מחזורי: דלוק, כבוי, ואז אינו נשלח כלל', async () => {
    const harness = mountDialog({ isOpen: true, busy: false });
    await settle();

    await clickToggle('צל');
    expect(toggle('צל').attributes('aria-pressed')).toBe('true');
    await footerButton('אישור').trigger('click');
    await settle();
    expect(harness.wrapper.emitted('submit')?.[0]?.[0]).toEqual({ shadow: true });

    await clickToggle('צל');
    expect(toggle('צל').attributes('aria-pressed')).toBe('false');
    await footerButton('אישור').trigger('click');
    await settle();
    expect(harness.wrapper.emitted('submit')?.[1]?.[0]).toEqual({ shadow: false });

    // הסיבוב השלישי מחזיר ל„ללא שינוי": אין מה לשלוח, ולכן גם אין מה ללחוץ.
    await clickToggle('צל');
    expect(toggle('צל').attributes('aria-pressed')).toBe('mixed');
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(harness.wrapper.emitted('submit')).toHaveLength(2);
  });

  /**
   * „אישור" על דיאלוג שאין בו מה להחיל היה נלחץ, סוגר, ולא עושה דבר — אותה
   * תקלה בדיוק שתוקנה ב„ברירות מחדל למסמך".
   */
  it('„אישור" נעול עד שיש מה להחיל, והמונה אומר כמה', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(teleported('.fa-count').text()).toBe('אין מה להחיל');

    await teleported('#fa-scale').setValue('120');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeUndefined();
    expect(teleported('.fa-count').text()).toBe('שינוי אחד יוחל');

    await clickToggle('צל');
    expect(teleported('.fa-count').text()).toBe('2 שינויים יוחלו');
  });

  /**
   * הפער שדווח: המשתמש מילא את הדיאלוג, לחץ „אישור", והכול נעלם — הדיאלוג
   * נסגר ולקח איתו את מה שהוקלד, ובשורת המצב הופיע „יש לסמן טקסט תחילה".
   * התשובה הזאת קיימת מהרגע הראשון (`format.apply` מקבל SelectionTarget
   * בלבד), ולכן היא נאמרת בפתיחה.
   */
  it('בלי טקסט מסומן — נאמר בפתיחה, ו„אישור" נעול גם אחרי שמולא שדה', async () => {
    mountDialog({ isOpen: true, busy: false }, false);
    await settle();

    expect(teleported('.fa-notice-blocking').text()).toContain('מסומן');
    expect(teleported('.fa-count').text()).toBe('אין טקסט מסומן');

    await teleported('#fa-scale').setValue('120');
    await settle();

    // מולא שדה — ובכל זאת נעול, ובכל זאת אותה הודעה: אין על מה להחיל.
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(teleported('.fa-count').text()).toBe('אין טקסט מסומן');
  });

  /**
   * הדיאלוג אינו חוסם את המסמך — אין מאחוריו רקע, והוא נגרר בכוונה. תשובה
   * שנקראה פעם אחת בפתיחה הייתה מתיישנת בשני הכיוונים, ושניהם מחזירים בדיוק
   * את הבאג שדווח: „אישור” פתוח על בחירה שהתכווצה, או נעול לנצח אחרי שהמשתמש
   * סימן טקסט כמו שההודעה ביקשה.
   */
  it('סימון טקסט בזמן שהדיאלוג פתוח פותח את „אישור" בלי לסגור דבר', async () => {
    const superdoc = createSuperdocDouble({ selection: { hasRange: false } });
    mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false }, superdoc });
    await settle();

    expect(teleported('.fa-notice-blocking').text()).toContain('מסומן');
    await teleported('#fa-scale').setValue('120');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();

    // המשתמש סימן טקסט במסמך שמאחורי הדיאלוג.
    superdoc.setSelection({ hasRange: true, text: 'טקסט מסומן' });
    document.dispatchEvent(new Event('selectionchange'));
    await new Promise((resolve) => setTimeout(resolve, 200));
    await settle();

    expect(document.querySelector('.fa-notice-blocking')).toBeNull();
    // והשדה שמולא לפני הסימון עדיין שם — לא נדרשה סגירה, ולכן לא אבד דבר.
    expect(footerButton('אישור').attributes('disabled')).toBeUndefined();
    expect(teleported('.fa-count').text()).toBe('שינוי אחד יוחל');
  });

  it('בחירה שהתכווצה בזמן שהדיאלוג פתוח נועלת את „אישור" מחדש', async () => {
    const superdoc = createSuperdocDouble({ selection: { hasRange: true, text: 'טקסט' } });
    mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false }, superdoc });
    await settle();

    await teleported('#fa-scale').setValue('120');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeUndefined();

    superdoc.setSelection({ hasRange: false });
    document.dispatchEvent(new Event('selectionchange'));
    await new Promise((resolve) => setTimeout(resolve, 200));
    await settle();

    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(teleported('.fa-notice-blocking').text()).toContain('מסומן');
  });

  /**
   * הפס נקרא פעם אחת לסבב, ולכן בלי שחרור מפורש הוא היה נשאר על מה שנקרא
   * ברגע הפתיחה — בדיוק בזרימה שההודעה החוסמת מבקשת, ומתחת לכיתוב שמתחייב
   * „כך ייראה ב-Word”.
   */
  it('פס התצוגה המקדימה נקרא מחדש כשהבחירה משתנה תחת דיאלוג פתוח', async () => {
    const superdoc = createSuperdocDouble({ selection: { hasRange: false } });
    mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false }, superdoc });
    await settle();

    // בלי בחירה הפס מציג את פסוק ברירת המחדל.
    expect(teleported('.fa-preview-strip').text()).not.toBe('טקסט חדש');

    superdoc.setSelection({ hasRange: true, text: 'טקסט חדש' });
    document.dispatchEvent(new Event('selectionchange'));
    await new Promise((resolve) => setTimeout(resolve, 260));
    await settle();

    expect(teleported('.fa-preview-strip').text()).toBe('טקסט חדש');
  });

  it('פס התצוגה המקדימה מתחלף גם בין שני טווחים מסומנים', async () => {
    const superdoc = createSuperdocDouble({ selection: { hasRange: true, text: 'טקסט ראשון' } });
    mountUi(FontAdvancedDialog, { props: { isOpen: true, busy: false }, superdoc });
    await settle();

    expect(teleported('.fa-preview-strip').text()).toBe('טקסט ראשון');

    // שתי הבחירות הן טווחים, לכן readiness נשאר 'range'. זה המקרה שבו
    // רענון שתלוי רק במעבר none↔range היה משאיר את הטווח הקודם בפס.
    superdoc.setSelection({ hasRange: true, text: 'טקסט שני' });
    document.dispatchEvent(new Event('selectionchange'));
    await new Promise((resolve) => setTimeout(resolve, 260));
    await settle();

    expect(teleported('.fa-preview-strip').text()).toBe('טקסט שני');
  });

  /**
   * `selectionchange` אינו נורה כשהמסמך עצמו מתחלף — סגירת לשונית או מעבר
   * ביניהן. בלי המעקב על המסמך „אישור” היה נשאר פתוח על תשובה של מסמך שכבר
   * אינו על המסך.
   */
  it('החלפת המסמך תחת דיאלוג פתוח נקראת מחדש', async () => {
    const withRange = createSuperdocDouble({ selection: { hasRange: true, text: 'טקסט' } });
    const harness = mountUi(FontAdvancedDialog, {
      props: { isOpen: true, busy: false },
      superdoc: withRange,
    });
    await settle();
    await teleported('#fa-scale').setValue('120');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeUndefined();

    harness.setSuperdoc(createSuperdocDouble({ selection: { hasRange: false } }));
    await settle();
    await new Promise((resolve) => setTimeout(resolve, 20));
    await settle();

    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(teleported('.fa-notice-blocking').text()).toContain('מסומן');
  });

  /**
   * „אין מסמך” אינו „אין יכולת”. בין מסמכים, ובזמן טעינה, המסמך הפעיל הוא
   * `null` — והודעה שאומרת שם „אינו זמין בגרסה זו של המנוע” היא טענה שקרית
   * על הבניין.
   */
  it('בלי מסמך פעיל — אין הודעה חוסמת ואין טענה על המנוע', async () => {
    const harness = mountUi(FontAdvancedDialog, {
      props: { isOpen: true, busy: false },
      superdoc: createSuperdocDouble({ selection: { hasRange: true } }),
    });
    await settle();

    harness.setSuperdoc(null);
    await settle();
    await new Promise((resolve) => setTimeout(resolve, 20));
    await settle();

    expect(document.querySelector('.fa-notice-blocking')).toBeNull();
  });

  /**
   * `HomeTab` מצהיר („הפתיחה מסבירה”) שכפתור „מתקדם” נשאר לחיץ גם בלי
   * Document API. בלי המצב הזה הדיאלוג היה נפתח שקט, נמלא, ונכשל אחרי הסגירה.
   */
  it('בלי format.apply — נאמר בפתיחה שהפעולה אינה זמינה', async () => {
    mountUi(FontAdvancedDialog, {
      props: { isOpen: true, busy: false },
      superdoc: createSuperdocDouble({ missing: ['format.apply'], selection: { hasRange: true } }),
    });
    await settle();

    expect(teleported('.fa-notice-blocking').text()).toContain('אינו זמין');
    expect(teleported('.fa-count').text()).toBe('אינו זמין');
    await teleported('#fa-scale').setValue('120');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
  });

  it('עם טקסט מסומן אין הודעה חוסמת, ו„אישור" נפתח כרגיל', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    expect(document.querySelector('.fa-notice-blocking')).toBeNull();

    await teleported('#fa-scale').setValue('120');
    await settle();
    expect(footerButton('אישור').attributes('disabled')).toBeUndefined();
  });

  /**
   * ההודעה על מה שהעורך אינו מצייר.
   *
   * היא אינה קישוט: זה בדיוק המקרה שדווח — „בחרתי עיצוב, לחצתי, לא קרה
   * כלום". ארבעת האפקטים נכתבים ל-docx ומוצגים ב-Word, והעורך אינו מצייר
   * אותם (נמדד; docs/engine-gaps.md, והתיקון ב-superdoc/docx-editor#3983).
   * מה שנבדק כאן הוא החיווט; שההודעה **תואמת את הפיקסלים** נמדד ב-Chrome
   * אמיתי ב-scripts/qa/font-advanced-qa.mjs, וזו הבדיקה שתיפול כשהמנוע
   * יתוקן ותאמר להוריד אפקט מהרשימה.
   */
  it('אפקט שאינו מצויר אומר את זה, בשמו, ורק כשהוא נבחר', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    expect(document.querySelector('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)')).toBeNull();

    await clickToggle('צל');
    const notice = teleported('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)').text();
    expect(notice).toContain('צל');
    expect(notice).toContain('נשמר בקובץ');
    expect(notice).toContain('אינו מצייר');
    // מה שלא נבחר אינו מוזכר — ההודעה מתארת את הבחירה, לא את הפקד.
    expect(notice).not.toContain('חרוט');

    await clickToggle('חרוט');
    expect(teleported('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)').text()).toContain('חרוט');

    // „כבוי" אינו „דלוק": ההודעה נעלמת יחד עם הבחירה.
    await clickToggle('צל', 2);
    await clickToggle('חרוט', 2);
    expect(document.querySelector('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)')).toBeNull();
  });

  /**
   * שם הפקד הוא נתון, ולכן כל נוסח שמטה פועל אחריו נשען על ידיעת מין ומספר
   * שאין כאן: „מסגרת לתו” היא נקבה ו„צל” זכר. הכותרת מוציאה את השמות מהמשפט.
   */
  it('ההודעה מונה את מה שנבחר, בשמו, בלי להטות פועל אחריו', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    await clickToggle('צל');
    const single = teleported('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)').text();
    expect(single).toContain('נשמר בקובץ');
    expect(single).toContain('צל');

    await clickToggle('חרוט');
    const both = teleported('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)').text();
    expect(both).toContain('צל, חרוט');
  });

  /**
   * הרשימה אינה „האפקטים” אלא כל מה שנמדד כנכתב-ואינו-מצויר. „קרנינג” ו„גודל”
   * של הגופן המורכב יושבים בשתי העמודות האחרות, ושתיקה עליהם הייתה חצי אמת על
   * אותו דיאלוג בדיוק.
   */
  it('גם פקדים שאינם אפקטים נמנים — קרנינג, וגודל הגופן המורכב', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    await teleported('#fa-kerning').setValue('12');
    await settle();
    expect(teleported('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)').text()).toContain('קרנינג');

    await teleported('#fa-sizecs').setValue('24');
    await settle();
    const both = teleported('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)').text();
    expect(both).toContain('גודל הגופן המורכב');
    // ומה שכן מצויר אינו נכנס לרשימה, גם כשהוא מולא באותו רגע.
    await teleported('#fa-scale').setValue('120');
    await settle();
    expect(teleported('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)').text()).not.toContain('מתיחה');
  });

  it('„קו חוצה כפול" מקבל נוסח משלו — הוא מצויר, אך כקו בודד', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    await clickToggle('קו חוצה כפול');
    const notice = teleported('.fontadv-dialog .fa-notice:not(.fa-notice-blocking)').text();
    expect(notice).toContain('בודד');
    // ואינו נמנה עם ארבעת ה„אינם מצוירים": הוא כן מצויר.
    expect(notice).not.toContain('אינו מצייר');
  });

  it('„נקה הכל" מחזיר כל שדה ל„ללא שינוי"', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    await teleported('#fa-scale').setValue('120');
    await clickToggle('חרוט');
    await teleported('#fa-lang').setValue('he-IL');
    await settle();
    expect(teleported('.fa-count').text()).toBe('3 שינויים יוחלו');

    await footerButton('נקה הכל').trigger('click');
    await settle();

    expect(teleported('.fa-count').text()).toBe('אין מה להחיל');
    expect(toggle('חרוט').attributes('aria-pressed')).toBe('mixed');
    expect((teleported('#fa-scale').element as HTMLInputElement).value).toBe('');
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
  });

  /**
   * הגופן המורכב היה תיבת טקסט חופשי. הבדיקה כאן היא שהבורר באמת מחובר
   * לרשימה של הרצועה — הרשימה שהמעטפת מזריקה — ולא לרשימה מקומית.
   */
  it('בורר הגופן שואב מרשימת הגופנים, והבחירה יוצאת ב-patch', async () => {
    const harness = mountDialog({ isOpen: true, busy: false });
    await settle();

    const input = document.querySelector<HTMLInputElement>('.fontadv-dialog .fa-combo input');
    if (!input) throw new Error('אין בורר גופן בדיאלוג');
    expect(input.placeholder).toBe('ללא שינוי');

    input.dispatchEvent(new FocusEvent('focus'));
    await settle();
    const rows = [...document.querySelectorAll('.fontadv-dialog .ribbon-combo-option')];
    const labels = rows.map((row) => row.textContent?.trim());
    expect(labels).toContain('ללא שינוי');
    expect(labels).toContain('Assistant');

    const assistant = rows.find((row) => row.getAttribute('data-value') === 'Assistant');
    if (!assistant) throw new Error('אין שורה ל-Assistant ברשימה');
    // `pointerdown` ולא `click`: זה מה שהרשימה מאזינה לו — ראו RibbonCombo.
    await new DOMWrapper(assistant).trigger('pointerdown');
    await settle();

    await footerButton('אישור').trigger('click');
    await settle();
    expect(harness.wrapper.emitted('submit')?.[0]?.[0]).toEqual({ complexFontName: 'Assistant' });
  });

  it('פס התצוגה המקדימה מצייר את הטקסט המסומן', async () => {
    const harness = mountUi(FontAdvancedDialog, {
      props: { isOpen: false, busy: false },
      superdoc: createSuperdocDouble({ selection: { hasRange: true, text: 'טקסט נבחר' } }),
    });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();

    expect(teleported('.fa-preview-strip').text()).toBe('טקסט נבחר');

    // מה שנבחר בדיאלוג מצויר על אותו טקסט — בלי לגעת במסמך.
    await teleported('#fa-spacing').setValue('3');
    await settle();
    const style = teleported('.fa-preview-text').attributes('style') ?? '';
    expect(style).toContain('letter-spacing: 3pt');
    expect(harness.superdoc.inputs('format.apply')).toEqual([]);
  });

  /**
   * `transform` פועל **אחרי** הפריסה, ולכן מתיחה אופקית מציירת מעבר לתיבה
   * שנמדדה — ומכיוון שהפס ממורכז, החריגה יוצאת לשני הצדדים וה-`overflow`
   * חותך את שניהם. נמדד בכרום ב-140%: מצויר 686..1234 מול מסגרת 693..1227,
   * כלומר גם ההתחלה וגם הסוף נעלמו. התקרה ההפוכה היא מה שמחזיר אותו פנימה.
   */
  it('„מתיחה אופקית" מקבלת תקרת רוחב הפוכה, אחרת הפס נחתך משני הצדדים', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    await teleported('#fa-scale').setValue('140');
    await settle();

    const style = teleported('.fa-preview-text').attributes('style') ?? '';
    expect(style).toContain('scaleX(1.4)');
    // 10000/140: התיבה **אחרי** המתיחה שווה בדיוק לרוחב המסגרת.
    expect(style).toContain('max-width: 71.43%');
  });

  it('„טקסט מוסתר = כן" מציג אזהרה גלויה', async () => {
    mountDialog({ isOpen: true, busy: false });
    await settle();

    expect(document.querySelector('.fa-warning')).toBeNull();
    await clickToggle('טקסט מוסתר');

    expect(teleported('.fa-warning').text()).toContain('יוסתר');
  });
});

describe('„גופן מתקדם" בלשונית „בית"', () => {
  it('הלחיצה פותחת את הדיאלוג', async () => {
    const harness = mountUi(HomeTab, { superdoc: createSuperdocDouble() });
    await settle();

    await harness.wrapper.find(ADVANCED_BUTTON).trigger('click');
    await settle();

    expect(document.querySelector('.fontadv-dialog')).not.toBeNull();
  });

  it('„אישור" מגיע ל-format.apply פעם אחת, עם ה-inline שנבנה', async () => {
    const harness = mountUi(HomeTab, {
      superdoc: createSuperdocDouble({ selection: { hasRange: true } }),
    });
    await settle();
    await harness.wrapper.find(ADVANCED_BUTTON).trigger('click');
    await settle();

    await teleported('#fa-scale').setValue('125');
    await teleported('#fa-position').setValue('3');
    await settle();
    await footerButton('אישור').trigger('click');
    await settle();

    const inputs = harness.superdoc.inputs('format.apply') as { inline: Record<string, unknown> }[];
    expect(inputs).toHaveLength(1);
    expect(inputs[0]!.inline).toEqual({ charScale: 125, position: 3 });
    expect(harness.failures()).toEqual([]);
  });

  it('כשל של המנוע מדווח בעברית ואינו מפיל את הלשונית', async () => {
    const superdoc = createSuperdocDouble({
      failures: { 'format.apply': { code: 'PRECONDITION_FAILED' } },
      selection: { hasRange: true },
    });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();
    await harness.wrapper.find(ADVANCED_BUTTON).trigger('click');
    await settle();

    await teleported('#fa-scale').setValue('200');
    await settle();
    await footerButton('אישור').trigger('click');
    await settle();

    const failures = harness.failures();
    expect(failures).toHaveLength(1);
    expect(failures[0]!.commandId).toBe('font-advanced');
    expect(failures[0]!.outcome.ok === false && failures[0]!.outcome.message).toContain(
      'החלת עיצוב הגופן נכשלה',
    );
    // הלשונית חיה: פתיחה חוזרת עובדת.
    await harness.wrapper.find(ADVANCED_BUTTON).trigger('click');
    await settle();
    expect(document.querySelector('.fontadv-dialog')).not.toBeNull();
  });

  /**
   * החצי השני של הבאג שדווח.
   *
   * הדיאלוג היה נסגר **לפני** שההחלה יצאה לדרך, ולכן כשלון השאיר הודעה בשורת
   * המצב ולקח איתו את שבעה-עשר השדות שמולאו. לא רק שלא קרה דבר — גם לא נשאר
   * מה לתקן ולנסות שוב.
   */
  it('כשל משאיר את הדיאלוג פתוח, עם הערכים שהוקלדו', async () => {
    const superdoc = createSuperdocDouble({
      failures: { 'format.apply': { code: 'PRECONDITION_FAILED' } },
      selection: { hasRange: true },
    });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();
    await harness.wrapper.find(ADVANCED_BUTTON).trigger('click');
    await settle();

    await teleported('#fa-scale').setValue('200');
    await settle();
    await footerButton('אישור').trigger('click');
    await settle();

    expect(document.querySelector('.fontadv-dialog')).not.toBeNull();
    expect((teleported('#fa-scale').element as HTMLInputElement).value).toBe('200');
    expect(teleported('.fa-count').text()).toBe('שינוי אחד יוחל');
  });

  it('הצלחה סוגרת אותו, כמו קודם', async () => {
    const harness = mountUi(HomeTab, {
      superdoc: createSuperdocDouble({ selection: { hasRange: true } }),
    });
    await settle();
    await harness.wrapper.find(ADVANCED_BUTTON).trigger('click');
    await settle();

    await teleported('#fa-scale').setValue('125');
    await settle();
    await footerButton('אישור').trigger('click');
    await settle();

    expect(document.querySelector('.fontadv-dialog')).toBeNull();
  });
});

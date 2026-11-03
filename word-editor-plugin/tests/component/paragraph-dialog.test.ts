/**
 * דיאלוג „פסקה”, וההגעה אליו מלשונית „בית”.
 *
 * למה קובץ ייעודי: הסורק הגנרי לוחץ על פותח הדיאלוג ומוודא שמשהו קרה, ולא
 * יודע להשוות את מה שנשלח למנוע. מה שנבדק כאן:
 * - מילוי מוקדם מהמסמך — הערכים ב**נקודות** מ-`doc.get` מוצגים בס"מ;
 * - „אישור” שולח **מצב מלא** לשלוש פעולות, ב-twips — היחידות שנמדדו;
 * - טאבים מיידיים: „הוסף” מגיע ל-`setTabStop` בלחיצה, ולא באישור;
 * - פקד טאבים נעלם כשהיכולת נדחית — ולא מוצג חסר תפקוד.
 *
 * הדיאלוג נבדק דרך ה-DOM של ה-document ולא דרך ה-wrapper, מפני שהוא מרונדר
 * ב-Teleport לגוף הדף — כמו BookmarkDialog, ומאותו טעם.
 */
import { describe, expect, it } from 'vitest';
import { DOMWrapper } from '@vue/test-utils';
import ParagraphDialog from '../../src/ui/panels/ParagraphDialog.vue';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import { autoUnmount, createSuperdocDouble, mountUi, settle, tipStartsSelector } from './harness';

autoUnmount();

function teleported(selector: string): DOMWrapper<Element> {
  const element = document.querySelector(selector);
  if (!element) throw new Error(`לא נמצא ${selector} בגוף הדף`);
  return new DOMWrapper(element);
}

function footerButton(label: string): DOMWrapper<Element> {
  const buttons = [...document.querySelectorAll('.para-dialog .pd-footer .pd-btn')];
  const found = buttons.find((button) => button.textContent?.trim() === label);
  if (!found) throw new Error(`לא נמצא הכפתור „${label}” בדיאלוג`);
  return new DOMWrapper(found);
}

const PARAGRAPH_BUTTON = tipStartsSelector('תפריט פסקה');

const EMPTY_SNAPSHOT = {
  indentation: { leftTwips: 0, rightTwips: 0, firstLineTwips: 0, hangingTwips: 0 },
  spacing: { beforeTwips: 0, afterTwips: 0, lineTwips: 240, rule: 'auto' },
  keepNext: false,
  keepLines: false,
  widowControl: true,
  tabs: [],
  bidi: true,
};

/**
 * A4 פחות שוליים של אינץ' לכל צד — 9026 twips, כלומר עמודה של 15.92 ס״מ —
 * וגופן של 11 נק'. אלה בדיוק המספרים של `tests/unit/paragraph-preview.test.ts`,
 * ולכן האחוזים כאן הם אותם אחוזים שנגזרו שם בחשבון יד: 1 ס״מ = 6.28%.
 */
const PREVIEW_PROPS = { pageTextWidthTwips: 11906 - 1440 - 1440, fontSizePt: 11, sectionRtl: false };

function dialogProps(patch: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    isOpen: true,
    busy: false,
    tabsEnabled: true,
    snapshot: EMPTY_SNAPSHOT,
    ...PREVIEW_PROPS,
    ...patch,
  };
}

/** הפסים של הפסקה שנערכת, בסדר שהם מצוירים בו. */
function targetLines(): HTMLElement[] {
  return [...document.querySelectorAll<HTMLElement>('.para-dialog .pd-pv-target .pd-pv-line')];
}

function previewPage(): HTMLElement | null {
  return document.querySelector<HTMLElement>('.para-dialog .pd-preview-page');
}

function previewTarget(): HTMLElement | null {
  return document.querySelector<HTMLElement>('.para-dialog .pd-pv-target');
}

const SPECIAL_AMOUNT = '.para-dialog input[aria-label="מידת הכניסה המיוחדת, בסנטימטרים"]';
const LINE_HEIGHT_FIELD = '.para-dialog input[aria-label="גובה השורה, בנקודות"]';

describe('ParagraphDialog (בדיד)', () => {
  it('סגור אינו מרונדר בכלל', () => {
    mountUi(ParagraphDialog, {
      props: dialogProps({ isOpen: false }),
    });
    expect(document.querySelector('.para-dialog')).toBeNull();
  });

  it('פתיחה ממקדת את שורש הדיאלוג — בלעדיה Escape מפסיק לעבוד', async () => {
    const harness = mountUi(ParagraphDialog, {
      props: dialogProps({ isOpen: false }),
    });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();

    expect(document.activeElement).toBe(document.querySelector('.para-dialog'));
  });

  it('Escape סוגר', async () => {
    const harness = mountUi(ParagraphDialog, {
      props: dialogProps(),
    });
    await settle();

    await teleported('.para-dialog').trigger('keydown.esc');
    expect(harness.wrapper.emitted('close')).toHaveLength(1);
  });

  it('busy מנטרל „אישור” ומשאיר „ביטול” חי', () => {
    mountUi(ParagraphDialog, {
      props: dialogProps({ busy: true }),
    });

    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(footerButton('ביטול').attributes('disabled')).toBeUndefined();
  });

  it('ערך שלילי חוסם את „אישור” עם הסבר', async () => {
    const harness = mountUi(ParagraphDialog, {
      props: dialogProps(),
    });
    await settle();

    await teleported('#pd-ind-left').setValue('-5');
    await settle();

    expect(teleported('.pd-error').text()).toContain('לא-שליליים');
    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(harness.wrapper.emitted('submit')).toBeUndefined();
  });
});

/**
 * פס התצוגה המקדימה.
 *
 * מה שנבדק כאן הוא **ההגעה ל-DOM**: שהמספרים שהגאומטריה חישבה מגיעים
 * לתכונות הנכונות, על האלמנטים הנכונים, ומהערכים ש**בשדות** ולא מהתצלום.
 * החישוב עצמו נבדק ב-tests/unit/paragraph-preview.test.ts.
 *
 * ⚠️ מה ש**אינו** נבדק כאן, ולמה: `font-size` בפועל. `cqw` היא יחידה
 * ש-jsdom אינו מכיר, ו-`setProperty('font-size', '2.44cqw')` נבלע שם בשקט
 * (נמדד) — כלומר בדיקה שהייתה נוקבת בגודל הגופן הייתה עוברת ירוק גם על
 * מספר שגוי. לכן הרכיב פולט `--pv-font`, שכן נשמר ב-jsdom, וההסבה שלו
 * ל-`font-size` היא שורת CSS סטטית שנמדדת בשער הדפדפן.
 */
describe('ParagraphDialog — פס התצוגה המקדימה', () => {
  it('מצויר מהערכים שבשדות ולא מהתצלום — זה כל ההבדל בין תצוגה מקדימה לחזרה', async () => {
    // התצלום אומר אפס כניסה; השדה נערך ל-1 ס"מ, והפס חייב לזוז.
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    expect(targetLines()[1]?.style.getPropertyValue('margin-inline-start')).toBe('0%');

    await teleported('#pd-ind-left').setValue('1');
    await settle();

    // 567 twips מתוך 9026 = 6.28%, על כל השורות מלבד הראשונה.
    for (const line of targetLines().slice(1)) {
      expect(line.style.getPropertyValue('margin-inline-start')).toBe('6.28%');
    }
  });

  it('הכניסה בסוף היא השוליים בצד השני של אותו פס', async () => {
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    await teleported('#pd-ind-right').setValue('2');
    await settle();

    // 1134 twips = 12.56%.
    expect(targetLines()[1]?.style.getPropertyValue('margin-inline-end')).toBe('12.56%');
  });

  it('„שורה ראשונה” מזיזה את הפס הראשון בלבד, ומוסיפה על הכניסה', async () => {
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    await teleported('#pd-ind-left').setValue('1');
    await teleported('#pd-special').setValue('firstLine');
    await teleported(SPECIAL_AMOUNT).setValue('1');
    await settle();

    const lines = targetLines();
    // 6.28% של הכניסה + 6.28% של השורה הראשונה.
    expect(lines[0]?.style.getPropertyValue('margin-inline-start')).toBe('12.56%');
    expect(lines[1]?.style.getPropertyValue('margin-inline-start')).toBe('6.28%');
  });

  it('„תלויה” מזיזה את הפס הראשון לכיוון ההפוך — ולשלילי, כשאין כניסה', async () => {
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    await teleported('#pd-special').setValue('hanging');
    await teleported(SPECIAL_AMOUNT).setValue('1');
    await settle();

    const lines = targetLines();
    expect(lines[0]?.style.getPropertyValue('margin-inline-start')).toBe('-6.28%');
    expect(lines[1]?.style.getPropertyValue('margin-inline-start')).toBe('0%');
  });

  it('הריווח לפני ואחרי הוא שוליים בציר הבלוק — על הפסקה, לא על הפסים', async () => {
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    await teleported('#pd-sp-before').setValue('12');
    await teleported('#pd-sp-after').setValue('6');
    await settle();

    // 240 twips = 2.66%, 120 = 1.33% — מול רוחב העמודה, כך CSS מודד שוליים
    // באחוזים גם בציר הזה.
    const target = previewTarget();
    expect(target?.style.getPropertyValue('margin-block-start')).toBe('2.66%');
    expect(target?.style.getPropertyValue('margin-block-end')).toBe('1.33%');
  });

  it('לפסקה עצמה אין שוליים בציר האופקי — שם נשבר המכנה', async () => {
    /*
     * זו ההכרעה המבנית של הפס, והמוטציה שחשפה את הפער: כניסה על הפסקה
     * **בנוסף** לפסים עוברת כל בדיקה שרק מאשרת שהפס זז. אחוז ב-CSS נמדד מול
     * הרוחב הפנימי של האב, ולכן שוליים על הפסקה מקטינים את האב של הפסים —
     * והאחוז שלהם מפסיק להיות אחוז מעמודת הטקסט. נמדד: כניסה של 1 ס״מ בשני
     * המקומות מציירת 6.28% מ-93.72%, כלומר 5.89% מהעמודה, ופס שמראה 12.17%
     * במקום 6.28%.
     */
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    await teleported('#pd-ind-left').setValue('1');
    await teleported('#pd-ind-right').setValue('1');
    await settle();

    const target = previewTarget();
    expect(target?.style.getPropertyValue('margin-inline-start')).toBe('');
    expect(target?.style.getPropertyValue('margin-inline-end')).toBe('');
    expect(target?.style.getPropertyValue('padding-inline-start')).toBe('');
    // ומה שכן עליה הוא הציר האנכי בלבד.
    expect(target?.style.getPropertyValue('margin-block-start')).toBe('0%');
  });

  it('מרווח השורות מגיע כמשתנה שגובה תיבת הפס נמדד בו', async () => {
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    expect(previewTarget()?.style.getPropertyValue('--pv-lh')).toBe('1');

    await teleported('#pd-line').setValue('480');
    await settle();
    expect(previewTarget()?.style.getPropertyValue('--pv-lh')).toBe('2');

    // „מדויקת” נמדד מול הגופן ולא מול 240: 22 נקודות על גופן 11 הן פי שניים.
    await teleported('#pd-line').setValue('exact');
    await settle();
    await teleported(LINE_HEIGHT_FIELD).setValue('22');
    await settle();
    expect(previewTarget()?.style.getPropertyValue('--pv-lh')).toBe('2');
  });

  it('גודל הגופן מגיע ב-cqw — אותו מכנה של האחוזים', async () => {
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    // 11 נקודות = 220 twips מתוך 9026 = 2.44%.
    expect(previewTarget()?.style.getPropertyValue('--pv-font')).toBe('2.44cqw');
  });

  it('השכנות מצוירות במרווח בודד וברוחב מלא — האישור אינו נוגע בהן', async () => {
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    await teleported('#pd-ind-left').setValue('3');
    await teleported('#pd-line').setValue('480');
    await settle();

    const contexts = [...document.querySelectorAll<HTMLElement>('.para-dialog .pd-pv-context')];
    expect(contexts).toHaveLength(2);
    for (const context of contexts) {
      expect(context.style.getPropertyValue('--pv-lh')).toBe('1');
    }
    for (const line of document.querySelectorAll<HTMLElement>(
      '.para-dialog .pd-pv-context .pd-pv-line',
    )) {
      expect(line.style.getPropertyValue('margin-inline-start')).toBe('');
    }
  });

  it('הכיוון הוא של הפסקה — `w:start` לוגי לה, לא למקטע', async () => {
    // התצלום מצהיר RTL והמקטע LTR: מה שגובר הוא הפסקה.
    const harness = mountUi(ParagraphDialog, { props: dialogProps({ sectionRtl: false }) });
    await settle();
    expect(previewPage()?.getAttribute('dir')).toBe('rtl');

    await harness.wrapper.setProps({ snapshot: { ...EMPTY_SNAPSHOT, bidi: false } });
    await settle();
    expect(previewPage()?.getAttribute('dir')).toBe('ltr');
  });

  it('פסקה שאינה מצהירה יורשת את כיוון המקטע — היעדר `w:bidi` אינו LTR', async () => {
    /*
     * מסמך עברי שנוצר ב-Word אינו מצהיר `<w:bidi>` על כל פסקה אלא יורש,
     * ופס שקורא היעדר כ-LTR מצייר את הכניסה בצד ההפוך על מסמכים רגילים
     * לגמרי. `null` הוא „שתקה”, ואז המקטע מכריע.
     */
    const silent = { ...EMPTY_SNAPSHOT, bidi: null };

    const harness = mountUi(ParagraphDialog, {
      props: dialogProps({ snapshot: silent, sectionRtl: true }),
    });
    await settle();
    expect(previewPage()?.getAttribute('dir')).toBe('rtl');

    await harness.wrapper.setProps({ sectionRtl: false });
    await settle();
    expect(previewPage()?.getAttribute('dir')).toBe('ltr');
  });

  it('הצהרה של הפסקה גוברת על המקטע — בשני הכיוונים', async () => {
    // פסקה LTR בתוך מקטע RTL מקבלת „לפני טקסט” בצד השמאלי, ולהפך.
    const harness = mountUi(ParagraphDialog, {
      props: dialogProps({
        snapshot: { ...EMPTY_SNAPSHOT, bidi: false },
        sectionRtl: true,
      }),
    });
    await settle();
    expect(previewPage()?.getAttribute('dir')).toBe('ltr');

    await harness.wrapper.setProps({ snapshot: { ...EMPTY_SNAPSHOT, bidi: true } });
    await settle();
    expect(previewPage()?.getAttribute('dir')).toBe('rtl');
  });

  it('בלי מידות עמוד אין פס, יש הסבר, וכל שאר הדיאלוג עובד', async () => {
    mountUi(ParagraphDialog, { props: dialogProps({ pageTextWidthTwips: 0 }) });
    await settle();

    expect(previewPage()).toBeNull();
    expect(teleported('.pd-preview-note').text()).toContain('מידות עמוד');
    // „אישור” אינו תלוי בפס.
    expect(footerButton('אישור').attributes('disabled')).toBeUndefined();
  });

  it('כניסות שאינן משאירות מקום נאמרות, ולא מצוירות כפסקה ברוחב אפס', async () => {
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    // הדיאלוג מתיר 55.87 ס"מ לכל שדה בנפרד ואינו מצליב ביניהם.
    await teleported('#pd-ind-left').setValue('8');
    await teleported('#pd-ind-right').setValue('8');
    await settle();

    const notes = [...document.querySelectorAll('.para-dialog .pd-preview-note')];
    expect(notes.map((note) => note.textContent?.trim()).join(' ')).toContain(
      'רוחב שאפשר לכתוב בו',
    );
  });

  it('ערך פסול נועל את „אישור” אבל אינו מעלים את הפס', async () => {
    // תזוזה שנייה באותו דיאלוג על אותה שגיאה אחת: השדה כבר אומר את הסיבה.
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();

    await teleported('#pd-ind-left').setValue('-5');
    await settle();

    expect(footerButton('אישור').attributes('disabled')).toBeDefined();
    expect(previewPage()).not.toBeNull();
  });
});

describe('ParagraphDialog — הפריסה בשתי עמודות', () => {
  it('המידות בעמודה אחת, ומה שאינו מידה בשנייה', () => {
    mountUi(ParagraphDialog, { props: dialogProps() });

    const columns = [...document.querySelectorAll('.para-dialog .pd-columns > .pd-column')];
    expect(columns).toHaveLength(2);

    const legends = (column: Element): string[] =>
      [...column.querySelectorAll('legend')].map((legend) => legend.textContent?.trim() ?? '');

    expect(legends(columns[0]!)).toEqual(['כניסה', 'ריווח']);
    expect(legends(columns[1]!)).toEqual(['אפשרויות שמירה', 'עצירות טאב']);
  });

  it('סדר שדות המספר לא זז — שער ה-QA פונה לשלושה מהם לפי מקומם', async () => {
    // scripts/qa/home-paragraph-qa.mjs ממלא את „מיוחד” באינדקס 2, את גובה
    // השורה באינדקס 5 ואת מיקום הטאב ב„אחרון”, ולשלושתם אין id לפנות אליו.
    mountUi(ParagraphDialog, { props: dialogProps() });
    await settle();
    // אותו סדר של השער עצמו: הוא בוחר „מדויקת” ואז ניגש לאינדקס 5.
    await teleported('#pd-line').setValue('exact');
    await settle();

    const fields = [...document.querySelectorAll('.para-dialog input[type=number]')].map(
      (input) => input.getAttribute('id') ?? input.getAttribute('placeholder'),
    );

    expect(fields).toEqual([
      'pd-ind-left',
      'pd-ind-right',
      null, // „מיוחד” — אינדקס 2
      'pd-sp-before',
      'pd-sp-after',
      null, // גובה השורה — אינדקס 5
      'מיקום', // הטאב, ואחרון
    ]);
  });
});

describe('„תפריט פסקה” בלשונית „בית”', () => {
  it('מקטע RTL ופסקה שאינה מצהירה — הפס מצייר מימין לשמאל', async () => {
    /*
     * זו ההשחלה, ולא החישוב: מוטציה שביטלה את קריאת `sectionDirection`
     * ב-HomeTab עברה ירוק על כל שאר הבדיקות — כלומר התיקון היה מנוטרל
     * באפליקציה בזמן שהדיאלוג עצמו נבדק כתקין. וזה גם המסמך הטיפוסי:
     * `paragraphProps` כאן אינו כולל `bidi`, בדיוק כמו פסקה שיורשת ב-Word.
     */
    const superdoc = createSuperdocDouble({
      paragraphProps: { indentation: { left: 36 } },
      sections: {
        pageSize: { width: 8.5, height: 11 },
        margins: { top: 1, right: 1, bottom: 1, left: 1 },
        direction: 'rtl',
      },
    });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();

    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    expect(document.querySelector('.para-dialog .pd-preview-page')?.getAttribute('dir')).toBe('rtl');
  });

  it('מקטע LTR ואותה פסקה — הפס מתהפך', async () => {
    const superdoc = createSuperdocDouble({
      paragraphProps: { indentation: { left: 36 } },
      sections: {
        pageSize: { width: 8.5, height: 11 },
        margins: { top: 1, right: 1, bottom: 1, left: 1 },
        direction: 'ltr',
      },
    });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();

    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    expect(document.querySelector('.para-dialog .pd-preview-page')?.getAttribute('dir')).toBe('ltr');
  });

  /**
   * הפס בא מהמסמך ולא מקנה מידה מומצא, וזו הבדיקה שאומרת את זה: הכפיל מחזיר
   * דף של 8.5 אינץ' עם שוליים של אינץ' לכל צד, כלומר עמודת טקסט של 9360
   * twips = 16.51 ס״מ, והכיתוב מעל הפס נוקב בה. שינוי בקריאה עצמה — או
   * מכנה שנלקח מרוחב הדף במקום מרוחב העמודה — נופל כאן.
   */
  it('הפתיחה קוראת את מידות המקטע, והפס מצויר על עמודת הטקסט האמיתית', async () => {
    const superdoc = createSuperdocDouble({
      sections: {
        pageSize: { width: 8.5, height: 11 },
        margins: { top: 1, right: 1, bottom: 1, left: 1 },
      },
    });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();

    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    expect(harness.superdoc.ops()).toContain('sections.list');
    const caption = document.querySelector('.para-dialog .pd-preview-caption');
    expect(caption?.textContent).toContain('16.51 ס"מ');
    // גודל הגופן הוא זה שבבורר הגודל שברצועה, בלי קריאה נוספת למנוע.
    expect(caption?.textContent).toContain("גופן 12 נק'");
    expect(document.querySelector('.para-dialog .pd-preview-page')).not.toBeNull();
  });

  it('מסמך שאינו מחזיר שוליים — אין פס, והדיאלוג עצמו שלם', async () => {
    // `readPageMargins` הוא „הכול או כלום”, וזה מצבו של מסמך שעדיין נטען.
    const harness = mountUi(HomeTab, { superdoc: createSuperdocDouble({ sections: { margins: null } }) });
    await settle();

    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    expect(document.querySelector('.para-dialog')).not.toBeNull();
    expect(document.querySelector('.para-dialog .pd-preview-page')).toBeNull();
    expect(document.querySelector('.para-dialog .pd-preview-note')?.textContent).toContain(
      'מידות עמוד',
    );
  });

  it('הלחיצה קוראת את מצב הפסקה (`get`) ופותחת את הדיאלוג', async () => {
    const harness = mountUi(HomeTab, { superdoc: createSuperdocDouble() });
    await settle();

    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    expect(harness.superdoc.ops()).toContain('get');
    expect(document.querySelector('.para-dialog')).not.toBeNull();
  });

  it('המילוי המוקדם בס"מ — 36 נקודות (720 twips) מוצגות כ-1.27', async () => {
    const superdoc = createSuperdocDouble({
      paragraphProps: { indentation: { left: 36 }, keepWithNext: true },
    });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();

    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    const left = teleported('#pd-ind-left').element as HTMLInputElement;
    expect(left.value).toBe('1.27');
  });

  it('„אישור” שולח מצב מלא לשלוש הפעולות, ב-twips', async () => {
    const harness = mountUi(HomeTab, { superdoc: createSuperdocDouble() });
    await settle();
    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    // כניסה שמאלית 1 ס"מ = 567 twips; ריווח לפני 12 נק' = 240.
    await teleported('#pd-ind-left').setValue('1');
    await teleported('#pd-sp-before').setValue('12');
    await settle();
    await footerButton('אישור').trigger('click');
    await settle();

    const inputs = (op: string) => harness.superdoc.inputs(op) as Record<string, unknown>[];
    expect(inputs('format.paragraph.setIndentation')).toHaveLength(1);
    expect(inputs('format.paragraph.setIndentation')[0]).toMatchObject({ left: 567, right: 0 });
    expect(inputs('format.paragraph.setSpacing')[0]).toMatchObject({ before: 240, after: 0 });
    expect(inputs('format.paragraph.setKeepOptions')).toHaveLength(1);
    expect(inputs('format.paragraph.setKeepOptions')[0]).toMatchObject({ widowControl: true });
  });

  it('הוספת טאב מיידית — 2 ס"מ מגיעים ל-`setTabStop` בלחיצה, ולא באישור', async () => {
    const harness = mountUi(HomeTab, { superdoc: createSuperdocDouble() });
    await settle();
    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    const position = document.querySelector('.para-dialog input[placeholder="מיקום"]') as HTMLInputElement;
    await new DOMWrapper(position).setValue('2');
    await settle();

    const addButtons = [...document.querySelectorAll('.para-dialog .pd-btn')].filter(
      (button) => button.textContent?.trim() === 'הוסף',
    );
    await new DOMWrapper(addButtons[0]!).trigger('click');
    await settle();

    const inputs = harness.superdoc.inputs('format.paragraph.setTabStop') as Record<string, unknown>[];
    expect(inputs).toHaveLength(1);
    expect(inputs[0]).toMatchObject({ position: 1134, alignment: 'left' });
  });

  it('יכולת טאבים שנדחית — סעיף הטאבים אינו מוצג כלל', async () => {
    const superdoc = createSuperdocDouble({ denied: ['format.paragraph.setTabStop'] });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();
    await harness.wrapper.find(PARAGRAPH_BUTTON).trigger('click');
    await settle();

    expect(document.querySelector('.para-dialog')).not.toBeNull();
    for (const fieldset of document.querySelectorAll('.para-dialog fieldset')) {
      expect(fieldset.textContent).not.toContain('עצירות טאב');
    }
  });
});

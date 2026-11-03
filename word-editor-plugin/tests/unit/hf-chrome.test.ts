/**
 * עברות שכבת הכותרות של המנוע.
 *
 * שלוש שאלות נמדדות כאן, ולא אחת:
 *
 *   1. **הנוסח.** `hebrewLabel` היא פונקציה טהורה, ולכן שמונת השילובים של
 *      סוג×ווריאנט וסיומת המקטע נמדדים בלי DOM בכלל.
 *   2. **מה שקורה בעץ.** aria-label, שני הכפתורים, יחידת המידה ותג ההמשך —
 *      וגם מה ש**אינו** קורה: ארבע התוויות בפאנל מתורגמות ב-CSS, ואם ה-JS גם
 *      יגע בהן יהיו שני מקורות לאותה תווית. הבדיקה שהן נשארות אנגלית כאן היא
 *      השער על ההפרדה הזאת, ולא פספוס.
 *   3. **מה שקורה אחרי ש-Vue כותב מחדש.** זה כל הטעם ב-MutationObserver:
 *      המנוע כותב את האנגלית בחזרה בכל patch, והבדיקה מדמה בדיוק את זה.
 *   4. **שפת המשתמש.** באנגלית העברות כולה כבויה — האנגלית של המנוע היא
 *      הנכונה שם, ורצועה אנגלית עם שכבת כותרות עברית היא רגרסיה.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { hebrewLabel, localizeEngineChrome, HF_TEXTS } from '../../src/engine/hf-chrome';
import { setMenuLocale } from '../../src/ui/ribbon/i18n';
import { headerFooterChrome, textOf, type ChromeOptions } from '../support/hf-chrome-dom';
import type { EngineChromeLocalizer } from '../../src/engine/hf-chrome';

let active: EngineChromeLocalizer | null = null;

/** בונה עץ, מתקין עליו עברות, ומחזיר את שניהם. הפירוק ב-afterEach. */
function localize(options: ChromeOptions = {}): {
  root: HTMLElement;
  localizer: EngineChromeLocalizer;
} {
  const root = headerFooterChrome(options);
  document.body.append(root);
  const localizer = localizeEngineChrome(root);
  active = localizer;
  return { root, localizer };
}

/** המתנה למעבר של ה-observer: jsdom מוסר רשומות ב-microtask. */
function flush(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

afterEach(() => {
  active?.dispose();
  active = null;
  document.body.innerHTML = '';
  // השפה גלובלית, ולכן חייבת לחזור: בדיקה שמשאירה 'en' מפילה את הבאות.
  setMenuLocale('he');
});

describe('נוסח התווית', () => {
  it('שמונת השילובים של סוג וווריאנט', () => {
    expect(hebrewLabel('Header', 'header', 'default')).toBe('כותרת עליונה');
    expect(hebrewLabel('First Page Header', 'header', 'first')).toBe('כותרת עליונה של עמוד ראשון');
    expect(hebrewLabel('Even Page Header', 'header', 'even')).toBe('כותרת עליונה של עמוד זוגי');
    expect(hebrewLabel('Odd Page Header', 'header', 'odd')).toBe('כותרת עליונה של עמוד אי-זוגי');
    expect(hebrewLabel('Footer', 'footer', 'default')).toBe('כותרת תחתונה');
    expect(hebrewLabel('First Page Footer', 'footer', 'first')).toBe('כותרת תחתונה של עמוד ראשון');
    expect(hebrewLabel('Even Page Footer', 'footer', 'even')).toBe('כותרת תחתונה של עמוד זוגי');
    expect(hebrewLabel('Odd Page Footer', 'footer', 'odd')).toBe('כותרת תחתונה של עמוד אי-זוגי');
  });

  it('מספר המקטע נשמר — הוא קיים רק בתוך הטקסט', () => {
    // המנוע מוסיף אותו למסמך בעל יותר ממקטע אחד, ואין לו תכונה משלו. תרגום
    // שמחליף מחרוזת קבועה היה מוחק אותו.
    expect(hebrewLabel('Header -Section 2-', 'header', 'default')).toBe('כותרת עליונה — מקטע 2');
    expect(hebrewLabel('First Page Footer -Section 11-', 'footer', 'first')).toBe(
      'כותרת תחתונה של עמוד ראשון — מקטע 11',
    );
  });

  it('בלי תכונות הנוסח נגזר מהטקסט — זה מסלול תגי ההמשך', () => {
    expect(hebrewLabel('Footer', null, null)).toBe('כותרת תחתונה');
    expect(hebrewLabel('Header -Section 3-', null, null)).toBe('כותרת עליונה — מקטע 3');
  });

  it('ווריאנט או סוג שאינם מוכרים נשארים באנגלית, ולא מקבלים נוסח שגוי', () => {
    // ווריאנט חמישי בגרסה עתידית: „כותרת עליונה” סתם היה שקר על תוכן המסמך.
    expect(hebrewLabel('Fifth Page Header', 'header', 'fifth')).toBeNull();
    expect(hebrewLabel('Sidenote', 'sidenote', 'default')).toBeNull();
  });

  it('טקסט שאינו מזוהה ואין עליו תכונות אינו מתורגם', () => {
    expect(hebrewLabel('כותרת עליונה', null, null)).toBeNull();
    expect(hebrewLabel('', null, null)).toBeNull();
  });
});

describe('העברות בעץ של המנוע', () => {
  it('התג הצף, שני הכפתורים ויחידת המידה', () => {
    const { root } = localize({ region: 'header', variant: 'first', label: 'First Page Header' });

    expect(textOf(root, '[data-sd-hf-label] > span')).toBe('כותרת עליונה של עמוד ראשון');
    expect(textOf(root, '[data-sd-hf-options]')).toBe(HF_TEXTS.optionsButton);
    expect(root.querySelector('[data-sd-hf-options]')?.getAttribute('title')).toBe(
      HF_TEXTS.optionsTitle,
    );
    expect(root.querySelector('[data-sd-hf-exit]')?.getAttribute('title')).toBe(
      HF_TEXTS.exitTitle,
    );
    // ה-„×” עצמו נשאר: הוא סימן, לא טקסט לתרגום.
    expect(textOf(root, '[data-sd-hf-exit]')).toBe('×');
    expect(textOf(root, '[data-sd-hf-option="header-from-top"] .v2-hf-distance-unit')).toBe('ס"מ');
  });

  it('aria-label של קבוצת הפקדים', () => {
    const { root } = localize();

    expect(root.querySelector('[data-sd-header-footer-active]')?.getAttribute('aria-label')).toBe(
      HF_TEXTS.groupLabel,
    );
  });

  it('תג ההמשך, שאין עליו תכונות', () => {
    const { root } = localize({ continuation: 'Footer -Section 2-' });

    expect(textOf(root, '[data-sd-hf-continuation-label] > span')).toBe('כותרת תחתונה — מקטע 2');
  });

  it('אינץ\' כשהמנוע מודד באינצ\'ים', () => {
    const { root } = localize({ unit: 'in' });

    expect(textOf(root, '[data-sd-hf-option="footer-from-bottom"] .v2-hf-distance-unit')).toBe(
      "אינץ'",
    );
  });

  it('ארבע התוויות בפאנל אינן נוגעות ב-JS — הן של CSS', () => {
    // שער ההפרדה: תווית שמתורגמת בשני מקומות היא תווית שאין לה מקור אחד,
    // ובפועל היא הופכת ל„מי כתב אחרון” בין כלל CSS לבין observer.
    const { root } = localize();

    expect(textOf(root, '[data-sd-hf-option="different-first-page"] > span')).toBe(
      'Different First Page',
    );
    expect(textOf(root, '[data-sd-hf-option="different-odd-even"] > span')).toBe(
      'Different Odd & Even Pages',
    );
    expect(textOf(root, '[data-sd-hf-option="header-from-top"] > span:first-child')).toBe(
      'Header from Top',
    );
    expect(textOf(root, '[data-sd-hf-option="footer-from-bottom"] > span:first-child')).toBe(
      'Footer from Bottom',
    );
  });

  it('ווריאנט שאינו מוכר משאיר את התג כפי שהוא', () => {
    const { root } = localize({ variant: 'fifth', label: 'Fifth Page Header' });

    expect(textOf(root, '[data-sd-hf-label] > span')).toBe('Fifth Page Header');
  });
});

describe('החזקה מול patch של המנוע', () => {
  it('אנגלית שנכתבה מחדש מתורגמת שוב', async () => {
    const { root } = localize();
    const chip = root.querySelector('[data-sd-hf-label] > span')!;
    const host = root.querySelector('[data-sd-header-footer-active]')!;
    expect(chip.textContent).toBe('כותרת עליונה');

    // בדיוק מה ש-Vue עושה כשהסמן עובר לכותרת של עמוד ראשון.
    host.setAttribute('data-sd-hf-variant', 'first');
    chip.textContent = 'First Page Header';
    await flush();

    expect(chip.textContent).toBe('כותרת עליונה של עמוד ראשון');
  });

  it('הפאנל שנוצר מחדש מקבל עברות — הוא v-if, ולא נשאר בעץ', async () => {
    const { root } = localize();
    const panel = root.querySelector('[data-sd-hf-options-panel]')!;
    const holder = panel.parentElement!;
    const clone = panel.cloneNode(true) as HTMLElement;
    panel.remove();
    // המנוע משבט vnodes שהוקפאו בבנייה, כלומר הפאנל שנפתח שוב נושא אנגלית.
    // בלי האיפוס הזה הבדיקה הייתה מודדת את התרגום שהעתקנו איתו.
    for (const unit of [...clone.querySelectorAll('.v2-hf-distance-unit')]) {
      unit.textContent = 'cm';
    }

    holder.append(clone);
    await flush();

    expect(textOf(clone, '.v2-hf-distance-unit')).toBe('ס"מ');
  });

  it('יחידה שהתחלפה בזמן ריצה מתורגמת מחדש', async () => {
    const { root } = localize({ unit: 'cm' });
    const unit = root.querySelector('[data-sd-hf-option="header-from-top"] .v2-hf-distance-unit')!;

    unit.textContent = 'in';
    await flush();

    expect(unit.textContent).toBe("אינץ'");
  });

  it('מעבר חוזר אינו משנה דבר', () => {
    const { root, localizer } = localize();
    const before = root.innerHTML;

    localizer.refresh();
    localizer.refresh();

    expect(root.innerHTML).toBe(before);
  });

  it('אחרי dispose המנוע נשאר עם האנגלית שלו', async () => {
    const { root, localizer } = localize();
    const chip = root.querySelector('[data-sd-hf-label] > span')!;

    localizer.dispose();
    chip.textContent = 'Footer';
    await flush();

    expect(chip.textContent).toBe('Footer');
  });
});

describe('שפת המשתמש', () => {
  it('באנגלית שום דבר אינו מתורגם — גם לא התג ולא היחידה', () => {
    // הרצועה כולה אנגלית שם (ui/ribbon/i18n.ts), והאנגלית של המנוע היא בדיוק
    // מה שצריך. עברות כאן הייתה יוצרת מסך בשתי שפות.
    setMenuLocale('en');
    const { root } = localize({ region: 'header', variant: 'first', label: 'First Page Header' });

    expect(textOf(root, '[data-sd-hf-label] > span')).toBe('First Page Header');
    expect(textOf(root, '[data-sd-hf-options]')).toBe('Options ▾');
    expect(root.querySelector('[data-sd-hf-options]')?.getAttribute('title')).toBe(
      'Header and footer options',
    );
    expect(root.querySelector('[data-sd-header-footer-active]')?.getAttribute('aria-label')).toBe(
      'Header and footer controls',
    );
    expect(textOf(root, '.v2-hf-distance-unit')).toBe('cm');
  });

  it('באנגלית גם ה-observer אינו מותקן — patch של המנוע אינו מתורגם', async () => {
    setMenuLocale('en');
    const { root } = localize();
    const chip = root.querySelector('[data-sd-hf-label] > span')!;

    chip.textContent = 'Footer';
    await flush();

    expect(chip.textContent).toBe('Footer');
  });

  it("'he-IL' ושפה שלא דווחה הם עברית — התוסף מעברת בברירת מחדל", () => {
    setMenuLocale('he-IL');
    expect(textOf(localize().root, '[data-sd-hf-label] > span')).toBe('כותרת עליונה');

    active?.dispose();
    document.body.innerHTML = '';
    setMenuLocale(undefined);
    expect(textOf(localize().root, '[data-sd-hf-label] > span')).toBe('כותרת עליונה');
  });
});

describe('השער שחוסך מעברים', () => {
  it('עץ בלי שכבת כותרות אינו נסרק — ה-observer יושב על משטח העריכה', () => {
    // ה-observer מתעורר על כל הקלדה במסמך; השכבה קיימת רק כשהסמן בכותרת.
    // הבדיקה מודדת את השער עצמו: סלקטור אחד במקום שמונה מעברים.
    const root = document.createElement('div');
    root.innerHTML = '<p>פסקה רגילה במסמך</p>';
    document.body.append(root);
    const scanned: string[] = [];
    const original = root.querySelectorAll.bind(root);
    root.querySelectorAll = ((selector: string) => {
      scanned.push(selector);
      return original(selector);
    }) as typeof root.querySelectorAll;

    active = localizeEngineChrome(root);

    expect(scanned).toEqual([]);
  });

  it('הקלדה במסמך שאין בו שכבת כותרות אינה סורקת כלל', async () => {
    // ה-observer יושב על משטח העריכה, ולכן כל תו מעיר אותו. בלי מסננת
    // הרשומות גם `hasChrome` לבדו — `querySelector` על שורש המסמך — נמדד
    // 120 פעמים ב-40 תווים. עכשיו הרשומה נבדקת, ולא העץ.
    const root = document.createElement('div');
    const paragraph = document.createElement('p');
    paragraph.textContent = 'פסקה רגילה במסמך';
    root.append(paragraph);
    document.body.append(root);

    active = localizeEngineChrome(root);

    const scanned: string[] = [];
    const one = root.querySelector.bind(root);
    const all = root.querySelectorAll.bind(root);
    root.querySelector = ((selector: string) => {
      scanned.push(selector);
      return one(selector);
    }) as typeof root.querySelector;
    root.querySelectorAll = ((selector: string) => {
      scanned.push(selector);
      return all(selector);
    }) as typeof root.querySelectorAll;

    paragraph.append(document.createTextNode('ועוד תו'));
    paragraph.firstChild!.nodeValue = 'פסקה רגילה במסמך!';
    await flush();

    expect(scanned).toEqual([]);
  });

  it('שכבה שנכנסת לעץ כן מעירה את המעבר — הרשומה נוגעת בה', async () => {
    const root = document.createElement('div');
    root.innerHTML = '<p>פסקה רגילה במסמך</p>';
    document.body.append(root);
    active = localizeEngineChrome(root);

    // בדיוק מה שקורה כשהסמן נכנס לכותרת: המנוע מוסיף את השכבה לעץ.
    root.append(headerFooterChrome());
    await flush();

    expect(textOf(root, '[data-sd-hf-label] > span')).toBe('כותרת עליונה');
  });
});

/**
 * פס הדגימה של בורר הגופן — מה מוצג בו, ומתי הוא נקרא.
 *
 * הפס הוא מה שנשלח **במקום** תצוגה חיה שצובעת את המסמך, אחרי שזו נמדדה כמטעה
 * (`docs/engine-gaps.md`, „תצוגה חיה של גופן”). ולכן מה שנבדק כאן אינו „הטקסט
 * מוצג” אלא שתי ההכרעות שיש בו, ושתיהן בלתי נראות בעין:
 *
 * 1. **קריאה אחת לסבב.** קריאה בכל ריחוף הייתה מחליפה את הדגימה מתחת לאצבע
 *    אילו משהו אחר הזיז את הסמן בזמן שהרשימה פתוחה.
 * 2. **תשובה שאיחרה נזרקת.** הקריאה אסינכרונית, והרשימה יכולה להיסגר בזמן
 *    שהיא באוויר — ואז הטקסט שלה שייך לסבב שנגמר, לא לסבב הבא.
 *
 * הבדיקה של הפס עצמו — מה מרונדר, באיזה גופן, ומה קורה לניווט המקלדת — היא
 * `tests/component/font-combo.test.ts`.
 */
import { describe, it, expect, vi } from 'vitest';
import {
  SAMPLE_SIZE_MAX_PX,
  SAMPLE_SIZE_MIN_PX,
  sampleSizePx,
} from '../../src/composables/use-font-controls';
import {
  createFontSample,
  FONT_SAMPLE_FALLBACK,
  FONT_SAMPLE_MAX_CHARS,
} from '../../src/composables/font-sample';

/**
 * ריקון תור המשימות. `setTimeout` ולא ספירת מיקרו-משימות: מספר ה-`await`ים
 * שנדרשים תלוי בשאלה אם `deps.read` הוא `async` או מחזיר Promise ידני, וזה
 * פרט מימוש של הבדיקה ולא של הקוד.
 */
const flush = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

/** סבב שלם: הרשימה נפתחה, הקריאה חזרה. */
async function beginAndSettle(sample: ReturnType<typeof createFontSample>): Promise<void> {
  sample.begin();
  await flush();
}

describe('מה הפס מציג', () => {
  it('לפני שנקרא דבר — משפט ברירת המחדל, ולא ריק', () => {
    const sample = createFontSample({ read: async () => 'לא נקרא' });
    expect(sample.text.value).toBe(FONT_SAMPLE_FALLBACK);
  });

  it('אחרי סבב — הטקסט שהמשתמש סימן', async () => {
    const sample = createFontSample({ read: async () => 'בראשית ברא' });
    await beginAndSettle(sample);
    expect(sample.text.value).toBe('בראשית ברא');
  });

  it('בחירה ריקה או רווחים בלבד — ברירת המחדל, ולא פס ריק', async () => {
    for (const text of ['', '   ', '\n\t ']) {
      const sample = createFontSample({ read: async () => text });
      await beginAndSettle(sample);
      expect(sample.text.value).toBe(FONT_SAMPLE_FALLBACK);
    }
  });

  it('רווחים בקצוות נחתכים — דגימה שמתחילה ברווח אינה מראה אותיות', async () => {
    const sample = createFontSample({ read: async () => '   שלום עולם   ' });
    await beginAndSettle(sample);
    expect(sample.text.value).toBe('שלום עולם');
  });
});

describe('תקרת האורך', () => {
  it('בחירה ארוכה נחתכת לתקרה', async () => {
    const long = 'א'.repeat(FONT_SAMPLE_MAX_CHARS + 40);
    const sample = createFontSample({ read: async () => long });
    await beginAndSettle(sample);
    expect(sample.text.value).toHaveLength(FONT_SAMPLE_MAX_CHARS);
  });

  it('בחירה שבתקרה בדיוק אינה נחתכת', async () => {
    const exact = 'ב'.repeat(FONT_SAMPLE_MAX_CHARS);
    const sample = createFontSample({ read: async () => exact });
    await beginAndSettle(sample);
    expect(sample.text.value).toBe(exact);
  });

  /**
   * הרגרסיה שהנימוק ל-`Array.from` קיים בשבילה: `slice` לפי יחידות קוד היה
   * מפצל את התו האחרון לחצי זוג surrogate, וזה מרונדר כ-‎`U+FFFD`‎.
   */
  it('חיתוך אינו מפצל זוג surrogate', async () => {
    // 59 אותיות ואחריהן תו מחוץ ל-BMP: החיתוך נופל בדיוק עליו.
    const text = 'ג'.repeat(FONT_SAMPLE_MAX_CHARS - 1) + '\u{1D504}\u{1D505}';
    const sample = createFontSample({ read: async () => text });
    await beginAndSettle(sample);

    expect(Array.from(sample.text.value)).toHaveLength(FONT_SAMPLE_MAX_CHARS);
    expect(sample.text.value).not.toContain('�');
    expect(sample.text.value.endsWith('\u{1D504}')).toBe(true);
  });
});

describe('קריאה אחת לסבב', () => {
  it('ריחוף על כמה שורות באותה פתיחה קורא פעם אחת', async () => {
    const read = vi.fn(async () => 'טקסט');
    const sample = createFontSample({ read });

    await beginAndSettle(sample);
    sample.begin();
    sample.begin();
    await beginAndSettle(sample);

    expect(read).toHaveBeenCalledTimes(1);
  });

  it('סגירה ופתיחה מחדש קוראות שוב — הבחירה יכולה להיות אחרת', async () => {
    let answer = 'ראשון';
    const read = vi.fn(async () => answer);
    const sample = createFontSample({ read });

    await beginAndSettle(sample);
    expect(sample.text.value).toBe('ראשון');

    sample.end();
    answer = 'שני';
    await beginAndSettle(sample);

    expect(read).toHaveBeenCalledTimes(2);
    expect(sample.text.value).toBe('שני');
  });

  it('סגירה מנקה את הפס — הפתיחה הבאה אינה מציגה את הדגימה של הקודמת', async () => {
    const sample = createFontSample({ read: async () => 'של הסבב שנגמר' });
    await beginAndSettle(sample);
    expect(sample.text.value).toBe('של הסבב שנגמר');

    sample.end();
    expect(sample.text.value).toBe(FONT_SAMPLE_FALLBACK);
  });
});

describe('תשובה שאיחרה', () => {
  /**
   * החתימה שהמונה קיים בשבילה: הקריאה יצאה, הרשימה נסגרה, והתשובה חוזרת אחר
   * כך. בלי המונה היא הייתה נוחתת בפס — ומוצגת בפתיחה הבאה עד שהקריאה החדשה
   * תחזור.
   */
  it('קריאה שחוזרת אחרי הסגירה אינה נכתבת לפס', async () => {
    let land: (text: string) => void = () => {};
    const sample = createFontSample({
      read: () => new Promise<string>((resolve) => (land = resolve)),
    });

    sample.begin();
    await flush();
    sample.end();

    land('איחרתי');
    await flush();

    expect(sample.text.value).toBe(FONT_SAMPLE_FALLBACK);
  });

  it('ותשובה של הסבב הפתוח כן נכתבת', async () => {
    let land: (text: string) => void = () => {};
    const sample = createFontSample({
      read: () => new Promise<string>((resolve) => (land = resolve)),
    });

    sample.begin();
    await flush();

    land('הגעתי בזמן');
    await flush();

    expect(sample.text.value).toBe('הגעתי בזמן');
  });
});

describe('משפט ברירת המחדל', () => {
  /** אות סופית היא אותה אות — לצורך „האם האלפבית מכוסה”. */
  const FINALS: Readonly<Record<string, string>> = {
    ך: 'כ',
    ם: 'מ',
    ן: 'נ',
    ף: 'פ',
    ץ: 'צ',
  };
  const ALPHABET = 'אבגדהוזחטיכלמנסעפצקרשת';

  /**
   * למה בדיקה ולא סתם קבוע: כל מה שהמשפט הזה עושה הוא **להראות את הגופן**,
   * וקיצור שלו ל„שלום עולם” היה משאיר פס שאינו מראה כלום — ועובר בשקט.
   */
  it('פנגרם: כל 22 אותיות האלפבית מופיעות בו', () => {
    const seen = new Set(
      Array.from(FONT_SAMPLE_FALLBACK.replace(/\s/g, '')).map((char) => FINALS[char] ?? char),
    );
    expect(Array.from(ALPHABET).filter((letter) => !seen.has(letter))).toEqual([]);
  });

  /**
   * הפס במצב הדגימה מחזיק שלוש שורות של 227 פיקסלים, כלומר 681 לכל היותר.
   * נמדד בכרום, ב-14px: הפסוק (63 תווים) דורש 373px ב-David, 394 ב-Arial
   * ו-434 ב-Segoe UI. Courier New, לעומת זאת, זקוק לשלוש שורות; לכן התקרה
   * כאן היא שלוש ולא שתיים.
   *
   * 70 היא התקרה: היא משאירה מרווח מעל 63 התווים שנמדדו, ועדיין נופלת לפני
   * שורה רביעית. `FONT_SAMPLE_MAX_CHARS` אינו התקרה כאן — הוא חוסם טקסט
   * שנקרא מהמסמך (`Ctrl+A` על ספר שלם), וזו מחרוזת קבועה שנמדדה.
   *
   * מה שאין כאן, ובכוונה: „הפסוק נכנס בפס”. זו טענה על פריסה, והיא נמדדת
   * במקום היחיד שיש בו פריסה — `scripts/qa/font-list-layout-qa.mjs`.
   */
  it('נכנס בשלוש השורות של הפס', () => {
    expect(Array.from(FONT_SAMPLE_FALLBACK).length).toBeLessThanOrEqual(70);
  });
});

/**
 * `isFallback` — מה שמפריד בין „הטקסט שלך, בגודל שבמסמך” לבין „הנה האותיות”.
 *
 * בלעדיו הפס היה מצייר את הפסוק בגודל הבחירה (עד 24px), כלומר חותך אותו בכל
 * גופן; ועם דגל **נפרד** לאותה ידיעה שני הצדדים היו יכולים לסתור זה את זה.
 */
describe('האם מוצג הפסוק או הטקסט של המשתמש', () => {
  it('לפני שנקרא משהו — הפסוק', () => {
    const sample = createFontSample({ read: async () => '' });
    expect(sample.isFallback.value).toBe(true);
    expect(sample.text.value).toBe(FONT_SAMPLE_FALLBACK);
  });

  it('טקסט מסומן מכבה אותו', async () => {
    const sample = createFontSample({ read: async () => 'בראשית ברא' });
    sample.begin();
    await flush();
    expect(sample.isFallback.value).toBe(false);
    expect(sample.text.value).toBe('בראשית ברא');
  });

  it('סגירת הרשימה מחזירה אותו', async () => {
    const sample = createFontSample({ read: async () => 'בראשית ברא' });
    sample.begin();
    await flush();
    sample.end();
    expect(sample.isFallback.value).toBe(true);
  });

  it('בחירה של רווחים בלבד אינה טקסט — הפס אינו נשאר ריק', async () => {
    const sample = createFontSample({ read: async () => '   ' });
    sample.begin();
    await flush();
    expect(sample.isFallback.value).toBe(true);
    expect(sample.text.value).toBe(FONT_SAMPLE_FALLBACK);
  });
});

/* ------------------------------------------------------------------ */
/* הגודל שבו הפס מצייר                                                */
/* ------------------------------------------------------------------ */

/**
 * זה מה שהופך את הדגימה מ„הנה האותיות” ל„הנה איך הטקסט שלי ייראה”: גופן נראה
 * אחרת לגמרי ב-11pt וב-20pt, ומי שבוחר כתב לספר בוחר את הצירוף. לפני כן הפס
 * צייר ב-14px קבועים, ולכן הבדיקה הראשונה כאן היא שהגודל בכלל נגזר מהמסמך.
 */
describe('sampleSizePx — הגודל שהפס מצייר בו', () => {
  it('נגזר מגודל הבחירה, ואינו קבוע', () => {
    expect(sampleSizePx(12)).toBe('16px');
    expect(sampleSizePx(18)).toBe('24px');
    expect(sampleSizePx(12)).not.toBe(sampleSizePx(18));
  });

  it('ההמרה היא `pt × 4/3` — אותו יחס שבו הדפדפן מצייר `pt`', () => {
    // 15pt = 20px בדיוק, ולכן זו השורה שתופסת יחס שגוי בלי עיגול שמסתיר אותו.
    expect(sampleSizePx(15)).toBe('20px');
  });

  it('גודל חריג נחסם, כדי שלא ישבור את הרשימה', () => {
    // 72pt הם 96px. פס בגובה קבוע לא היה מכיל אותם.
    expect(sampleSizePx(72)).toBe(`${SAMPLE_SIZE_MAX_PX}px`);
    // 6pt הם 8px — קטן מכדי להראות צורת אות.
    expect(sampleSizePx(6)).toBe(`${SAMPLE_SIZE_MIN_PX}px`);
  });

  it('הגדלים הנפוצים נשארים בתוך הטווח, ולכן יוצאים מדויקים', () => {
    for (const pt of [10, 11, 12, 14, 16, 18]) {
      const px = Number.parseFloat(sampleSizePx(pt));
      expect(px).toBe(Math.round(pt * (4 / 3)));
      expect(px).toBeGreaterThanOrEqual(SAMPLE_SIZE_MIN_PX);
      expect(px).toBeLessThanOrEqual(SAMPLE_SIZE_MAX_PX);
    }
  });

  it('גודל שאינו מספר אינו הופך ל-`NaNpx`', () => {
    // `font-size: NaNpx` הוא ערך פסול שהדפדפן מתעלם ממנו בשקט — כלומר הפס היה
    // חוזר לגודל של הרשימה בלי שאיש ידע למה.
    expect(sampleSizePx(Number.NaN)).toBe(`${SAMPLE_SIZE_MIN_PX}px`);
    expect(sampleSizePx(Number.POSITIVE_INFINITY)).toBe(`${SAMPLE_SIZE_MIN_PX}px`);
  });

  it('תמיד מחזיר מחרוזת עם `px`, כי היא נכתבת ישר ל-`style`', () => {
    for (const pt of [6, 12, 20, 72]) expect(sampleSizePx(pt)).toMatch(/^\d+px$/);
  });
});

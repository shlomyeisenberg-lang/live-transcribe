/**
 * מילון ראשי-התיבות: הפענוח (engine/acronyms.ts), הטעינה העצלה
 * (engine/acronym-dictionary.ts), והחוזה מול הנכס שהבנייה מייצרת בפועל.
 *
 * הקבוצה האחרונה היא זו שהייתה חסרה: הבדיקה הקודמת מסרה לטוען
 * `JSON.stringify(...)` — מחרוזת מלאכותית שאינה מה שהבנייה מציבה — ולכן עברה
 * בירוק גם כשהטוען ציפה למחרוזת, הנכס הציב אובייקט, וההשלמה לא עבדה מעולם.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createAcronymDictionary, looksLikeAcronym, type PackedAcronyms } from '../../src/engine/acronyms';
import { loadAcronymDictionary, resetAcronymDictionary } from '../../src/engine/acronym-dictionary';
import { ACRONYMS_GLOBAL } from '../../src/engine/acronyms-constants';
import { buildAcronymsAsset, packAcronyms, readAcronymsSource } from '../../scripts/acronyms-asset';

const PACKED: PackedAcronyms = { 'א"א': 'אי אפשר', 'רמב"ן': 'רבי משה בן נחמן' };

describe('createAcronymDictionary', () => {
  const dictionary = createAcronymDictionary(PACKED);

  it('מחזירה את הפירוש לר"ת מוכר', () => {
    expect(dictionary.lookup('א"א')).toBe('אי אפשר');
    expect(dictionary.lookup('רמב"ן')).toBe('רבי משה בן נחמן');
  });

  it('גרשיים של מקלדת עברית (U+05F4) מוצאים את הערך', () => {
    expect(dictionary.lookup('רמב״ן')).toBe('רבי משה בן נחמן');
  });

  it('שני גרשים נפרדים שקולים לגרשיים', () => {
    expect(dictionary.lookup("רמב''ן")).toBe('רבי משה בן נחמן');
  });

  it('ניקוד על המילה שהוקלדה אינו מונע התאמה', () => {
    expect(dictionary.lookup('רַמְבַּ"ן')).toBe('רבי משה בן נחמן');
  });

  it('null לר"ת שאינו ברשימה', () => {
    expect(dictionary.lookup('זזזזז')).toBeNull();
  });

  it('null גם למפתחות של Object.prototype', () => {
    expect(dictionary.lookup('constructor')).toBeNull();
    expect(dictionary.lookup('__proto__')).toBeNull();
  });
});

describe('looksLikeAcronym', () => {
  it('מזהה גרשיים בכל צורות הכתיב', () => {
    expect(looksLikeAcronym('א"א')).toBe(true);
    expect(looksLikeAcronym('רמב״ן')).toBe(true);
    expect(looksLikeAcronym("רמב''ן")).toBe(true);
    expect(looksLikeAcronym('רַמְבַּ"ן')).toBe(true);
  });

  it('מזהה קיצור בגרש בסוף המילה', () => {
    expect(looksLikeAcronym("ר'")).toBe(true);
    expect(looksLikeAcronym("וכו'")).toBe(true);
    expect(looksLikeAcronym('וכו׳')).toBe(true);
  });

  it('דוחה מילה רגילה וגרשיים בקצה', () => {
    // אלה המילים שבלי השער הזה היו מזריקות את הנכס לחינם — ובראשן 3,862
    // המילים הארמיות שבקובץ המקור, שאינן ראשי תיבות.
    expect(looksLikeAcronym('אמר')).toBe(false);
    expect(looksLikeAcronym('אבדיקציה')).toBe(false);
    expect(looksLikeAcronym('שלום"')).toBe(false);
    expect(looksLikeAcronym('"שלום')).toBe(false);
    expect(looksLikeAcronym('')).toBe(false);
  });

  it('דוחה ערך מרובה מילים — `lookup` בודק מילה אחת', () => {
    expect(looksLikeAcronym('א"ס ב"ה')).toBe(false);
  });
});

describe('loadAcronymDictionary', () => {
  beforeEach(() => resetAcronymDictionary());

  it('נטען פעם אחת בלבד, גם בקריאות מקבילות', async () => {
    let calls = 0;
    const loader = async () => {
      calls += 1;
      return PACKED;
    };

    const [first, second] = await Promise.all([loadAcronymDictionary(loader), loadAcronymDictionary(loader)]);
    const third = await loadAcronymDictionary(loader);
    expect(calls).toBe(1);
    expect(first).toBe(second);
    expect(third).toBe(first);
  });

  it('כשל מחזיר null, וההמתנה מונעת ניסיון חוזר בכל השהיה', async () => {
    let calls = 0;
    const failThenSucceed = async () => {
      calls += 1;
      return calls === 1 ? null : PACKED;
    };

    expect(await loadAcronymDictionary(failThenSucceed)).toBeNull();
    // בלי ההמתנה כל הקלדה הייתה מזריקה את הנכס מחדש.
    expect(await loadAcronymDictionary(failThenSucceed)).toBeNull();
    expect(calls).toBe(1);

    resetAcronymDictionary();
    expect((await loadAcronymDictionary(failThenSucceed))?.lookup('א"א')).toBe('אי אפשר');
    expect(calls).toBe(2);
  });

  it('loader שזורק נחשב כשל ואינו מפיל את ההשלמה', async () => {
    let calls = 0;
    const thrower = async () => {
      calls += 1;
      throw new Error('נכס פגום');
    };

    expect(await loadAcronymDictionary(thrower)).toBeNull();
    expect(await loadAcronymDictionary(thrower)).toBeNull();
    expect(calls).toBe(1);
  });
});

describe('הנכס שהבנייה מייצרת', () => {
  const globals = globalThis as Record<string, unknown>;

  beforeEach(() => resetAcronymDictionary());
  afterEach(() => {
    delete globals[ACRONYMS_GLOBAL];
  });

  /** מריצה את הנכס כפי שהדפדפן מריץ אותו, ומחזירה את מה שנשאר ב-`window`. */
  function runAsset(): unknown {
    const asset = buildAcronymsAsset();
    new Function('window', asset)(globals);
    return globals[ACRONYMS_GLOBAL];
  }

  it('נטען דרך המסלול האמיתי ומשלים ר"ת מהמילון', async () => {
    runAsset();

    // ה-loader המוגדר כברירת מחדל מוצא את הגלובל ואינו צריך <script> כלל.
    const dictionary = await loadAcronymDictionary();
    expect(dictionary).not.toBeNull();
    expect(dictionary?.lookup('חז"ל')).toBe('חכמינו זכרונם לברכה');
    expect(dictionary?.lookup('חז״ל')).toBe('חכמינו זכרונם לברכה');
    // מ-KleiKodesh, ואינו במילון אוצריא הרשמי כלל.
    expect(dictionary?.lookup('רמב"ם')).toBe('רבי משה בן מימון');
    // מפתח שנכתב בשני גרשים במקור, ונכנס לנכס בצורתו המנורמלת.
    expect(dictionary?.lookup("ר'")).toBe('רבי');
  });

  it('כל מפתח בנכס עובר את השער שלפני הטעינה', () => {
    const packed = runAsset() as Record<string, string>;
    const keys = Object.keys(packed);

    expect(keys.filter((key) => !looksLikeAcronym(key))).toEqual([]);
    // אף פירוש אינו הר"ת עצמו — הצעה כזאת אינה משלימה דבר.
    expect(keys.filter((key) => packed[key] === key)).toEqual([]);
  });

  /**
   * מה שהמסלול הזה מכניס למסמך נכתב כפי שהוא (בשונה מהמסלול הסטטי, שעובר
   * דרך `normalizeSelectedText`), ולכן תוכן הפירוש הוא חלק מהחוזה. 503
   * מפתחות ב-KleiKodesh הם לעזי רש"י שהפירוש היחיד שלהם „*” או „/”.
   */
  it('אין פירוש שהוא אשפה או רווח חריג', () => {
    const packed = runAsset() as Record<string, string>;
    const keys = Object.keys(packed);

    expect(keys.filter((key) => !/[א-ת]/.test(packed[key]!))).toEqual([]);
    expect(keys.filter((key) => packed[key]!.trim() !== packed[key])).toEqual([]);
    expect(keys.filter((key) => /\s\s/.test(packed[key]!))).toEqual([]);
  });

  /**
   * הספירה מדויקת ולא „לפחות”: החלפת קובץ הנתונים היא בדיוק הדבר שיזיז את
   * המספרים האלה בשקט — ערך שיחדל להיות ר"ת פשוט ייעלם מההשלמה.
   */
  it('הסינון מדויק ומדווח', () => {
    const source = readAcronymsSource();
    const pack = packAcronyms(source);

    expect(Object.keys(source).length).toBe(25_363);
    expect(Object.keys(pack.packed).length).toBe(17_717);
    expect(pack.dropped.notAcronym).toBe(6_508);
    expect(pack.dropped.noExpansion).toBe(522);
    expect(pack.fromVariants).toBe(232);
  });

  it('הנכס קטן מקובץ המקור, למרות שיש בו יותר ערכים מקודם', () => {
    const source = readAcronymsSource();
    const { packed } = packAcronyms(source);

    // 17,717 ערכים ב-710KB, מול 13,067 ב-902KB לפני המיזוג — הכיווץ לפירוש
    // יחיד מחזיר יותר ממה שההרחבה לוקחת.
    expect(Buffer.byteLength(JSON.stringify(packed))).toBeLessThan(800_000);
    expect(Buffer.byteLength(JSON.stringify(source))).toBeGreaterThan(2_000_000);
  });
});

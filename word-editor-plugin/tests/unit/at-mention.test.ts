import { describe, expect, it } from 'vitest';
import {
  buildLinkText,
  buildRefHref,
  isQueryable,
  parseAtTrigger,
  suggestionSubtitle,
  type ResolvedRefHit,
} from '../../src/engine/at-mention';

function hit(overrides: Partial<ResolvedRefHit> = {}): ResolvedRefHit {
  return {
    id: 42,
    bookId: 'פסחים',
    bookUid: 'id:42',
    title: 'פסחים',
    reference: 'פסחים דף לד',
    index: 1234,
    isPdf: false,
    isSourceLine: true,
    isUserBook: false,
    bookPath: 'ש"ס, בבלי',
    ...overrides,
  };
}

describe('parseAtTrigger', () => {
  it('תופס אזכור בתחילת פסקה', () => {
    expect(parseAtTrigger('@פסחים לד')).toEqual({ query: 'פסחים לד', atIndex: 0 });
  });

  it('תופס אזכור אחרי רווח, ומחזיר את היסט ה-@ עצמו', () => {
    const trigger = parseAtTrigger('כמו שכתוב ב@פסחים לד');
    expect(trigger).toEqual({ query: 'פסחים לד', atIndex: 11 });
  });

  it('תופס אזכור אחרי אות שימוש — הצורה הטבעית בעברית', () => {
    // "ראה ב@פסחים לד" הוא איך שכותבים את זה בפועל; אות השימוש נשארת במסמך
    // ורק ה-@ מוחלף.
    expect(parseAtTrigger('ראה ב@פסחים לד')).toEqual({ query: 'פסחים לד', atIndex: 5 });
    expect(parseAtTrigger('כמובא ל@רמב"ם')?.query).toBe('רמב"ם');
    expect(parseAtTrigger('ו@ברכות')?.query).toBe('ברכות');
  });

  it('אות שימוש נספרת רק כשהיא עומדת לבדה', () => {
    // "כתוב@" — ב' היא סופה של מילה שלמה, לא אות שימוש.
    expect(parseAtTrigger('כתוב@פסחים')).toBeNull();
  });

  it('אינו תופס כתובת דוא"ל', () => {
    // הבעיה שהבדיקה הזו מגנה עליה: כל הקלדה של כתובת הייתה פותחת רשימת הצעות.
    expect(parseAtTrigger('שלחו אלי dev@example.com')).toBeNull();
    expect(parseAtTrigger('dev@exa')).toBeNull();
  });

  it('תופס אזכור אחרי סוגר או מירכאה פותחת', () => {
    expect(parseAtTrigger('(@פסחים')?.query).toBe('פסחים');
    expect(parseAtTrigger('"@פסחים')?.query).toBe('פסחים');
  });

  it('לוקח את ה-@ האחרון', () => {
    expect(parseAtTrigger('@ברכות ב, @פסחים')).toEqual({ query: 'פסחים', atIndex: 10 });
  });

  it('@ בלבד הוא טריגר פתוח, אך אינו בשל לשאילתה', () => {
    const trigger = parseAtTrigger('@');
    expect(trigger).toEqual({ query: '', atIndex: 0 });
    expect(isQueryable(trigger!)).toBe(false);
  });

  it('נסגר על שבירת שורה', () => {
    expect(parseAtTrigger('@פסחים\nלד')).toBeNull();
  });

  it('נסגר על רווח כפול — המשפט המשיך', () => {
    expect(parseAtTrigger('@פסחים  לד')).toBeNull();
  });

  it('נסגר כשההפניה ארוכה מדי', () => {
    expect(parseAtTrigger(`@${'א'.repeat(49)}`)).toBeNull();
    expect(parseAtTrigger(`@${'א'.repeat(48)}`)).not.toBeNull();
  });

  it('נסגר על יותר מדי מילים', () => {
    expect(parseAtTrigger('@תלמוד ירושלמי עירובין פרק ו הלכה ז')).not.toBeNull();
    expect(parseAtTrigger('@א ב ג ד ה ו ז ח')).toBeNull();
  });

  it('רווח נגרר נשמר בטריגר אך אינו מונע שאילתה', () => {
    const trigger = parseAtTrigger('@פסחים ');
    expect(trigger?.query).toBe('פסחים ');
    expect(isQueryable(trigger!)).toBe(true);
  });

  it('סף השאילתה הוא שני תווים', () => {
    expect(isQueryable({ query: 'פ', atIndex: 0 })).toBe(false);
    expect(isQueryable({ query: 'פס', atIndex: 0 })).toBe(true);
    expect(isQueryable({ query: ' פ ', atIndex: 0 })).toBe(false);
  });
});

/**
 * הסף נמדד באורך בלבד: גרשיים אינו „מפעיל” ואינו נדרש. הבדיקות כאן רצות על
 * המחרוזות כפי שהן נראות בזמן הקלדה — parseAtTrigger ואז isQueryable, כמו
 * ב-overlay — כדי שלא נאמת סף על אובייקט טריגר שנבנה ביד.
 */
describe('סף שני התווים על טקסט שהוקלד', () => {
  const queryable = (beforeCaret: string): boolean => {
    const trigger = parseAtTrigger(beforeCaret);
    return trigger !== null && isQueryable(trigger);
  };

  it('אות אחת אינה מפעילה', () => {
    expect(queryable('@ר')).toBe(false);
    expect(queryable('ראה @ר')).toBe(false);
    expect(queryable('ב@ר')).toBe(false);
  });

  it('שתי אותיות מפעילות — בלי שום סימן פיסוק', () => {
    expect(queryable('@רש')).toBe(true);
    expect(queryable('ראה @רש')).toBe(true);
    expect(queryable('ב@רש')).toBe(true);
    expect(queryable('@פס')).toBe(true);
  });

  it('אות וגרשיים מפעילים — התו השני הוא שממלא את הסף', () => {
    expect(queryable('@ר״')).toBe(true);
    expect(queryable('@ר"')).toBe(true);
    expect(queryable('@ר׳')).toBe(true);
  });

  it('הגרשיים נשמר בשאילתה כלשונו', () => {
    expect(parseAtTrigger('@ר״')?.query).toBe('ר״');
    expect(parseAtTrigger('@רש״י')?.query).toBe('רש״י');
  });

  it('רווח אינו נספר לסף — „@ר ” עדיין תו אחד', () => {
    expect(queryable('@ר ')).toBe(false);
    expect(queryable('@ר ב')).toBe(true);
  });
});

describe('buildRefHref', () => {
  it('בונה קישור עומק לספר טקסט, עם המזהה היציב', () => {
    expect(buildRefHref(hit(), 'פסחים לד')).toBe('otzaria://open/book/42?index=1234&uid=id%3A42');
  });

  it('בלי bookUid הקישור נשאר כפי שהיה — הוא אופציונלי', () => {
    expect(buildRefHref(hit({ bookUid: undefined }), 'פסחים לד')).toBe(
      'otzaria://open/book/42?index=1234',
    );
  });

  it('בונה קישור PDF עם מספר עמוד', () => {
    expect(buildRefHref(hit({ isPdf: true, index: 17 }), 'x')).toBe(
      'otzaria://open/pdf/42?index=17&uid=id%3A42',
    );
  });

  it('עמוד PDF לעולם אינו קטן מ-1 — הראוטר דוחה 0', () => {
    expect(buildRefHref(hit({ isPdf: true, index: 0, bookUid: undefined }), 'x')).toBe(
      'otzaria://open/pdf/42?index=1',
    );
  });

  it('ספר אישי נופל לאיתור מקורות ולא לקישור עומק', () => {
    // user_books.db מקצה מזהים באותו טווח כמו ספריית הבסיס: קישור לפי id
    // היה נפתר לספר אחר לגמרי.
    const href = buildRefHref(hit({ isUserBook: true, reference: 'הערות שלי' }), 'הערות שלי');
    expect(href).toBe('otzaria://open/detection?q=%D7%94%D7%A2%D7%A8%D7%95%D7%AA%20%D7%A9%D7%9C%D7%99');
  });

  /**
   * זה הבאג שדווח: המשתמש הקליד „בסוגי”, בחר מהרשימה „בסוגיא דדיורים בריבית”,
   * והקישור נשמר עם `q=בסוגי` — כלומר „איתור מקורות” חיפש קידומת חתוכה ולא
   * מצא דבר, בזמן שהטקסט הנראה בקישור היה מלא ותקין.
   */
  it('הכתובת נבנית מההתאמה שנבחרה, לא מהקידומת שהוקלדה', () => {
    const href = buildRefHref(
      hit({ id: null, isUserBook: true, reference: 'בסוגיא דדיורים בריבית' }),
      'בסוגי',
    );
    const q = decodeURIComponent(new URL(href).searchParams.get('q') ?? '');
    expect(q).toBe('בסוגיא דדיורים בריבית');
  });

  it('הכתובת והטקסט הנראה נבנים מאותו מקור ואינם יכולים לסטות', () => {
    const match = hit({ id: null, reference: 'תלמוד ירושלמי עירובין פרק ו' });
    const href = buildRefHref(match, 'ירוש עיר');
    expect(decodeURIComponent(new URL(href).searchParams.get('q') ?? '')).toBe(
      buildLinkText(match, 'ירוש עיר'),
    );
  });

  it('היעדר id נופל לאיתור מקורות', () => {
    expect(buildRefHref(hit({ id: null }), 'ספר סרוק ג')).toContain('open/detection?q=');
    expect(buildRefHref(hit({ id: undefined }), 'ספר סרוק ג')).toContain('open/detection?q=');
  });

  it('רווחים מיותרים ב-bookUid אינם נכתבים לקישור', () => {
    expect(buildRefHref(hit({ bookUid: '  id:42  ' }), 'x')).toBe(
      'otzaria://open/book/42?index=1234&uid=id%3A42',
    );
    expect(buildRefHref(hit({ bookUid: '   ' }), 'x')).toBe('otzaria://open/book/42?index=1234');
  });

  it('כשאין להתאמה שם — מה שהוקלד, מנורמל', () => {
    const href = buildRefHref(hit({ id: null, reference: '  ', title: '' }), '  בראשית א  ');
    expect(href).toBe('otzaria://open/detection?q=%D7%91%D7%A8%D7%90%D7%A9%D7%99%D7%AA%20%D7%90');
    expect(decodeURIComponent(new URL(href).searchParams.get('q') ?? '')).toBe('בראשית א');
  });
});

describe('buildLinkText', () => {
  /**
   * `doc.insert` עם `type: 'text'` מכניס טקסט לפסקה, ושבר שורה בתוכו אינו
   * נכנס כפי שהוא; וטווח העטיפה מחושב כ-`start + text.length`, ולכן תו
   * שנספר אחרת ממה שנכתב היה מזיז את קצה הקישור.
   */
  it('מנרמל רווחים לשורה אחת', () => {
    const withBreak = 'בראשית\nפרק  א';
    expect(buildLinkText(hit({ reference: withBreak }), 'x')).toBe('בראשית פרק א');
    const withTab = '  פסחים\tלד  ';
    expect(buildLinkText(hit({ reference: '', title: withTab }), 'x')).toBe('פסחים לד');
  });

  it('מעדיף את ההפניה שנפתרה על מה שהוקלד', () => {
    expect(buildLinkText(hit(), 'פסחים לד')).toBe('פסחים דף לד');
  });

  it('נופל לכותרת ואז למה שהוקלד', () => {
    expect(buildLinkText(hit({ reference: '   ' }), 'פסחים לד')).toBe('פסחים');
    expect(buildLinkText(hit({ reference: '', title: '' }), ' פסחים לד ')).toBe('פסחים לד');
  });
});

describe('suggestionSubtitle', () => {
  it('מציג את נתיב הקטגוריה', () => {
    expect(suggestionSubtitle(hit())).toBe('ש"ס, בבלי');
  });

  it('מסמן התאמה ברמת פרק בלבד', () => {
    expect(suggestionSubtitle(hit({ isSourceLine: false }))).toBe('ש"ס, בבלי · לרמת הפרק');
  });

  it('מסמן ספר אישי ו-PDF', () => {
    expect(suggestionSubtitle(hit({ isUserBook: true }))).toContain('ספר אישי');
    expect(suggestionSubtitle(hit({ isPdf: true, isSourceLine: false }))).toBe('ש"ס, בבלי · PDF');
  });

  it('ריק כשאין מה לומר', () => {
    expect(suggestionSubtitle(hit({ bookPath: '' }))).toBe('');
  });
});

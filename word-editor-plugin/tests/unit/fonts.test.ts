import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import {
  BUNDLED_FAMILIES,
  BUNDLED_FILES,
  bundledFontFaceCss,
  installBundledFonts,
} from '../../src/styles/fonts';

/** vitest רץ משורש המאגר, ולכן public/ נמצא ביחס ל-cwd. */
const FONT_DIR = join(process.cwd(), 'public', 'fonts');

/**
 * האם ה-cmap של קובץ ה-TTF ממפה מתו מהבלוק העברי (U+0590-05FF)? נקרא ישירות
 * מהטבלאות ולא דרך ספרייה: הבדיקה צריכה להעיד על הבייטים שנארזים בפועל.
 * נתמכים פורמט 4 (BMP) ופורמט 12, שני הפורמטים שגופני Google Fonts מוציאים.
 */
function hasHebrewCmap(font: Buffer): boolean {
  const numTables = font.readUInt16BE(4);
  let cmapOffset = 0;
  for (let i = 0; i < numTables; i++) {
    const record = 12 + i * 16;
    if (font.toString('ascii', record, record + 4) === 'cmap') {
      cmapOffset = font.readUInt32BE(record + 8);
      break;
    }
  }
  if (!cmapOffset) return false;

  const numSubtables = font.readUInt16BE(cmapOffset + 2);
  for (let i = 0; i < numSubtables; i++) {
    const sub = cmapOffset + font.readUInt32BE(cmapOffset + 4 + i * 8 + 4);
    const format = font.readUInt16BE(sub);
    if (format === 4) {
      const segCount = font.readUInt16BE(sub + 6) / 2;
      for (let s = 0; s < segCount; s++) {
        const end = font.readUInt16BE(sub + 14 + s * 2);
        const start = font.readUInt16BE(sub + 16 + segCount * 2 + s * 2);
        if (start <= 0x5d0 && 0x5d0 <= end) return true;
      }
    } else if (format === 12) {
      const nGroups = font.readUInt32BE(sub + 12);
      for (let g = 0; g < nGroups; g++) {
        const start = font.readUInt32BE(sub + 16 + g * 12);
        const end = font.readUInt32BE(sub + 20 + g * 12);
        if (start <= 0x5d0 && 0x5d0 <= end) return true;
      }
    }
  }
  return false;
}

describe('bundledFontFaceCss', () => {
  it('מצהיר על כל משקל בכל שם', () => {
    const css = bundledFontFaceCss();

    expect(css.match(/@font-face/g)).toHaveLength(
      BUNDLED_FILES.length * BUNDLED_FAMILIES.length,
    );
    for (const family of BUNDLED_FAMILIES) {
      for (const file of BUNDLED_FILES) {
        expect(css).toContain(`font-family:'${family}';font-style:normal;font-weight:${file.weight}`);
      }
    }
  });

  it('מצהיר על „Assistant” — השם שבו --font-main קורא לגופן', () => {
    // אי-התאמה בין השם כאן לשם ב-tokens.css פירושה גופן ארוז שאף אחד לא
    // מבקש, כלומר ממשק שנופל לגופן המערכת.
    expect(BUNDLED_FAMILIES).toEqual(['Assistant']);
  });

  it('אותו קובץ משרת את שני השמות', () => {
    const css = bundledFontFaceCss();

    for (const file of BUNDLED_FILES) {
      const uses = css.split(`url('${file.url}')`).length - 1;
      expect(uses).toBe(BUNDLED_FAMILIES.length);
    }
  });

  it('מבקש truetype ו-swap', () => {
    const css = bundledFontFaceCss();

    expect(css).toContain("format('truetype')");
    // הגופן לא יעכב את הצגת המסמך.
    expect(css).toContain('font-display:swap');
  });

  it('אינו מצהיר unicode-range — Assistant מכסה גם עברית וגם לטינית', () => {
    // הגופן הקודם (Selawik) חייב טווחים, מפני שהוא הוצהר גם בשם „Segoe UI”
    // ואין בו עברית. Assistant מוצהר בשמו בלבד ומכסה את שני הכתבים, ולכן טווח
    // כאן היה רק מונע ממתווים שיש בגופן לקבל אותו.
    expect(bundledFontFaceCss()).not.toContain('unicode-range');
  });

  it('מכסה בפועל את הבלוק העברי', () => {
    // זו הסיבה להחלפה: העברית — כמעט כל מה שייכתב בתוסף — חייבת לבוא מהגופן
    // הארוז ולא מ-fallback שרירותי של המערכת.
    const font = readFileSync(join(FONT_DIR, 'Assistant-Regular.ttf'));
    expect(hasHebrewCmap(font)).toBe(true);
  });

  it('אינו מצהיר על נטוי — האריזה כוללת רק פנים זקופות', () => {
    // הצהרת italic שמצביעה על הפנים הרגילה הייתה מונעת מהדפדפן להטות אותה
    // סינתטית, כלומר טקסט נטוי היה נראה זקוף.
    expect(bundledFontFaceCss()).not.toContain('italic');
  });

  it('כל הנתיבים יחסיים ל-document', () => {
    // נתיב מוחלט (`/fonts/...`) נפתר ל-שורש הדיסק ב-file://, ואוצריא טוענת
    // תוסף ארוז בדיוק משם.
    for (const file of BUNDLED_FILES) {
      expect(file.url.startsWith('./fonts/')).toBe(true);
    }
  });

  it('אין הצהרה כפולה לאותו משקל', () => {
    const weights = BUNDLED_FILES.map((file) => file.weight);
    expect(new Set(weights).size).toBe(weights.length);
  });
});

describe('installBundledFonts', () => {
  it('מזריק style אחד ל-head', () => {
    const doc = document.implementation.createHTMLDocument('t');
    installBundledFonts(doc);

    const styles = doc.head.querySelectorAll('style');
    expect(styles).toHaveLength(1);
    expect(styles[0]!.textContent).toContain("font-family:'Assistant'");
  });

  it('אידמפוטנטי — קריאה שנייה אינה מכפילה את ההצהרות', () => {
    const doc = document.implementation.createHTMLDocument('t');
    installBundledFonts(doc);
    installBundledFonts(doc);

    expect(doc.head.querySelectorAll('style')).toHaveLength(1);
  });
});

describe('הקבצים שנארזים', () => {
  it('כל קובץ שמוצהר קיים בפועל ואינו ריק', () => {
    for (const file of BUNDLED_FILES) {
      const path = join(FONT_DIR, file.url.replace('./fonts/', ''));
      expect(existsSync(path), `חסר ${file.url}`).toBe(true);
      expect(statSync(path).size).toBeGreaterThan(10_000);
    }
  });

  it('אין בתיקייה קובץ גופן שאינו מוצהר', () => {
    // קובץ שנשאר שם בלי הצהרה הוא נפח מבוזבז בחבילה — וזה בדיוק מה שקרה
    // בתוסף שממנו הגיע הגופן הקודם.
    const declared = new Set(BUNDLED_FILES.map((file) => file.url.replace('./fonts/', '')));
    const onDisk = readdirSync(FONT_DIR).filter((name) => name.endsWith('.ttf'));

    expect(onDisk.sort()).toEqual([...declared].sort());
  });

  it('נוסח ה-OFL נארז לצד הגופן', () => {
    // סעיף 2 ב-OFL: אין להפיץ את הגופן בלי נוסח הרישיון.
    expect(existsSync(join(process.cwd(), 'public/third-party/ASSISTANT-LICENSE.txt'))).toBe(true);
  });
});

describe('שרשרת --font-main', () => {
  const CHAIN =
    /--font-main:\s*([^;]+);/.exec(
      readFileSync(join(process.cwd(), 'src', 'styles', 'tokens.css'), 'utf8'),
    )?.[1] ?? '';

  it('פותחת בגופן הארוז', () => {
    // אף אחד לא דורס את `--font-main` בזמן ריצה (host/theme.ts), ולכן השרשרת
    // כאן היא מה שהממשק מקבל בפועל. חוליה ראשונה שאינה הגופן הארוז פירושה
    // ממשק שתלוי במערכת ההפעלה.
    expect(CHAIN).not.toBe('');
    expect(CHAIN.trimStart().startsWith(`'${BUNDLED_FAMILIES[0]}'`)).toBe(true);
  });

  it('אינה נופלת לגופן ספרים בשום חוליה', () => {
    // הרגרסיה: גופן ספרים בכפתור בן 12px נמרח. אין די בכך שהחוליה הראשונה
    // תקינה — חוליה שנייה כזאת היא אותו באג ברגע שהגופן הארוז לא נטען.
    const BOOK_FONTS = [
      'FrankRuhlCLM',
      'Frank Ruhl',
      'TaameyDavidCLM',
      'TaameyAshkenaz',
      'NotoRashiHebrew',
      'NotoSerifHebrew',
      'KeterYG',
      'Shofar',
      'Tinos',
      'David',
      'serif',
    ];

    for (const font of BOOK_FONTS) {
      // `sans-serif` הוא ההיפך מ-serif ולכן אינו נחשב פגיעה.
      const hits = CHAIN.replace(/sans-serif/g, '').includes(font);
      expect(hits, `${font} בשרשרת`).toBe(false);
    }
  });
});

/**
 * השער של `create.image` על `src`, מול המנוע **הארוז** ולא מול הזיכרון.
 *
 * ## למה הבדיקה הזאת קיימת
 *
 * הפקד „תמונות” קרא `imageCmd.run()` בלי payload, כלומר נכשל סגור תמיד. בדרך
 * לתיקון התברר שהחוזה אינו מה שנראה מהטיפוסים: `CreateImageInput.src` מוצהר
 * `string`, ומי שקורא רק את ה-`.d.ts` מסיק שאפשר למסור לו URL. **אי אפשר.**
 * המימוש דורש data URI בבסיס 64, ודוחה כל דבר אחר.
 *
 * זו אינה החמרה תיאורטית: בורר הקבצים של אוצריא מחזיר `url` של שרת loopback,
 * וזה מה שהיה מתבקש להעביר. אילו המנוע היה מקבל אותו, התמונה הייתה נשמרת
 * במסמך כהפניה לכתובת שהפורט שלה משתנה בכל הפעלה של אוצריא ושאינה קיימת בכלל
 * במכונה של מי שיקבל את המסמך — כלומר אובדן נתונים שקט, שנראה תקין עד הפתיחה
 * הבאה. העובדה שהמנוע דוחה URL היא מה שהפך את זה לכשל גלוי במקום שקט.
 *
 * ## למה מול המקור הארוז
 *
 * הטיפוסים לא מתעדים את השער, ולכן שינוי שלו בגרסה עתידית לא יפיל שום
 * typecheck. הבדיקה קוראת את הקוד שהחבילה שולחת ומקבעת שני דברים: את הביטוי
 * הרגולרי עצמו, ואת קבוצת סוגי ה-mime שמתקבלים. אם אחד מהם ישתנה — למשל אם
 * `create.image` יתחיל לקבל URL, או יתחיל לקבל WebP — הבדיקה תיפול, וזה בדיוק
 * מה שאנחנו רוצים לדעת: אז אפשר להרחיב את בורר הקבצים.
 *
 * ההשוואה אינה „הרשימה שלנו שווה לרשימה שלהם” אלא **הכלה**: כל מה ש-
 * `isEmbeddableImageSrc` מאשר חייב לעבור את השער של המנוע. זה מה שמונע את
 * הכשל השקט, וזה גם מה שמאפשר לנו להיות מחמירים יותר מהמנוע במקום שבו הוא
 * מקבל ואז דוחה בשלב הבא (`image/jpg`, `image/webp`).
 */
import { describe, it, expect } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { imagePayload, isEmbeddableImageSrc } from '../../src/engine/payloads';

const ENGINE_DIR = join(process.cwd(), 'node_modules/@superdoc/docx-engine/dist');

/**
 * שני ה-bundles שנושאים את המימוש. `document-runtime.js` הוא הקטן, ולכן הוא
 * נבדק ראשון; שניהם מכילים את אותו שער. שם הקובץ אינו נושא hash — הוא בשדה
 * ה-exports של החבילה — ולכן הוא נכתב ולא מחופש.
 */
const BUNDLES = ['document-runtime.js', 'docx-engine.es.js'] as const;

function readEngineBundle(): { name: string; source: string } {
  for (const name of BUNDLES) {
    const path = join(ENGINE_DIR, name);
    if (existsSync(path)) return { name, source: readFileSync(path, 'utf8') };
  }
  throw new Error(`לא נמצא bundle של מנוע ה-DOCX ב-${ENGINE_DIR}`);
}

const { name: BUNDLE_NAME, source: ENGINE } = readEngineBundle();

/**
 * הודעות הכשל של השער, כפי שהן במקור. ב-bundle הן מקודדות עם `\x20` במקום
 * רווח, ולכן זו הצורה שמחופשת.
 */
const REQUIRES_DATA_URI = 'create.image\\x20currently\\x20requires\\x20a\\x20base64\\x20data\\x20URI.';
const REQUIRES_PNG_OR_JPEG = 'create.image\\x20currently\\x20requires\\x20PNG\\x20or\\x20JPEG\\x20input.';

/**
 * הביטוי הרגולרי שהשער מפעיל על `src`, נחלץ מהמקור ומורכב מחדש.
 *
 * החילוץ לפי הליטרל בקוד ולא לפי שם הפונקציה: שמות הפונקציות ב-bundle
 * מעורפלים ומשתנים בין builds, בעוד הליטרל הוא חלק מהחוזה.
 */
function extractSrcGate(): RegExp {
  const match = ENGINE.match(/\/(\^data:[^/]*?base64[^/]*?\$)\/i/);
  if (!match) {
    throw new Error(
      `השער של create.image על src לא נמצא ב-${BUNDLE_NAME} — חוזה ה-src של superdoc השתנה`,
    );
  }
  return new RegExp(match[1], 'i');
}

const srcGate = extractSrcGate();

/**
 * קטע הקוד שמיד אחרי הביטוי הרגולרי — שרשרת ה-`if` שממפה mime לפורמט.
 * 400 תווים מכסים אותה בשלמותה בשני ה-bundles, ועוצרים לפני הקוד שאחריה.
 */
function srcGateRegion(): string {
  const index = ENGINE.search(/\/\^data:[^/]*?base64[^/]*?\$\/i/);
  expect(index, 'קטע השער לא נמצא').toBeGreaterThan(-1);
  return ENGINE.slice(index, index + 400);
}

describe('השער של create.image על src', () => {
  it('המנוע דורש data URI בבסיס 64, ואומר זאת במפורש', () => {
    expect(ENGINE).toContain(REQUIRES_DATA_URI);
  });

  it('הביטוי שנחלץ הוא זה שדוחה URL ומקבל data URI', () => {
    // אילו התיעוד היה מדויק אפשר היה לדלג על זה. `src: string` בטיפוסים אינו
    // מרמז על שום דבר מזה.
    expect(srcGate.test('data:image/png;base64,iVBORw==')).toBe(true);
    expect(srcGate.test('http://127.0.0.1:51763/file/abc')).toBe(false);
    expect(srcGate.test('file:///C:/Users/a/b.png')).toBe(false);
    expect(srcGate.test('https://example.com/a.png')).toBe(false);
  });

  it('WebP נדחה על ידי create.image גם אחרי שהוא נפרס', () => {
    // ולכן הוא אינו ברשימת הסיומות של הבורר. `images.replaceSource` כן מקבל
    // אותו (`allowWebp`), וזו הסיבה שהוא נפרס בכלל.
    expect(ENGINE).toContain(REQUIRES_PNG_OR_JPEG);
  });

  it('סוגי ה-mime שהפרסר ממפה הם png, jpeg ו-webp בלבד', () => {
    // נבדק על **קטע השער עצמו** ולא על ה-bundle כולו: `'image/gif'` מופיע
    // במקומות אחרים בחבילה (ייבוא, תצוגה), ובדיקה על כל הקובץ הייתה מאשרת
    // אותו בטעות. מה שקובע את `EMBEDDABLE_IMAGE_EXTENSIONS` הוא רשימת הענפים
    // שכאן: כל השאר נופל ל-else שמחזיר כשל.
    const gate = srcGateRegion();

    for (const mime of ['image/png', 'image/jpeg', 'image/jpg', 'image/webp']) {
      expect(gate, mime).toContain(mime);
    }
    for (const mime of ['image/gif', 'image/bmp', 'image/svg', 'image/tif']) {
      expect(gate, mime).not.toContain(mime);
    }
  });
});

describe('isEmbeddableImageSrc מוכל בשער של המנוע', () => {
  const CANDIDATES = [
    'data:image/png;base64,iVBORw==',
    'data:image/jpeg;base64,/9j/4AA=',
    'data:image/jpg;base64,/9j/4AA=',
    'data:image/webp;base64,UklGRg==',
    'data:image/gif;base64,R0lGOD==',
    'data:image/svg+xml;base64,PHN2Zz4=',
    'data:image/png,abc',
    'data:image/png;base64,',
    'http://127.0.0.1:51763/file/abc',
    'https://example.com/a.png',
    'file:///C:/a.png',
    '',
  ];

  it('כל src שאנחנו מאשרים עובר את השער של המנוע', () => {
    // זו הטענה שמונעת את הכשל השקט. אם היא נשברת, המנוע דוחה payload שבנינו.
    const approvedButRejected = CANDIDATES.filter(
      (src) => isEmbeddableImageSrc(src) && !srcGate.test(src),
    );
    expect(approvedButRejected).toEqual([]);
  });

  it('אנחנו מחמירים בדיוק במקום שהמנוע מקבל ואז דוחה', () => {
    // `image/webp` עובר את הביטוי ונדחה במשפט הבא, ו-`image/jpg` מתקבל אך
    // אינו צורה שיש סיבה לייצר. שניהם נעצרים אצלנו, עם הודעה בעברית.
    expect(srcGate.test('data:image/webp;base64,UklGRg==')).toBe(true);
    expect(isEmbeddableImageSrc('data:image/webp;base64,UklGRg==')).toBe(false);
    expect(isEmbeddableImageSrc('data:image/jpg;base64,/9j/4AA=')).toBe(false);
  });

  it('imagePayload אינו מייצר payload שהשער ידחה', () => {
    for (const src of CANDIDATES) {
      const payload = imagePayload({ src });
      if (payload) expect(srcGate.test(payload.src), src).toBe(true);
    }
  });
});

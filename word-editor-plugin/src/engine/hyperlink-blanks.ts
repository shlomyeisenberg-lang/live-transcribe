/**
 * ניקוי צומתי קישור שנשארו בלי טקסט.
 *
 * ## מאיפה הם מגיעים
 *
 * מחיקת הטקסט של קישור — בסימון ו-Backspace, כמו כל טקסט אחר — **אינה מוחקת
 * את צומת הקישור**. נמדד על 2.15.0: אחרי מחיקת „ברכות ב” נשאר
 * `<w:hyperlink r:id="rId8"><w:r><w:rPr>…</w:rPr></w:r></w:hyperlink>` —
 * ‏`<w:r>` בלי `<w:t>` בכלל — ו-`hyperlinks.list()` ממשיך להחזיר את הקישור עם
 * `text: ""` ועוגן באורך אפס.
 *
 * ## למה זה לא קוסמטי
 *
 * שני נזקים, ושניהם דווחו מהשטח:
 *
 * 1. **בקובץ שיוצא יש קישור בלתי-נראה.** הוא תקין מבחינת OOXML, אין לו טקסט
 *    לראות או ללחוץ עליו, והוא מופיע ברשימת הקישורים של Word.
 * 2. **הוא חוסם כתיבת קישור חדש באותו מקום.** גם `hyperlinks.insert` וגם
 *    `hyperlinks.wrap` דוחים טווח שנוגע בקישור קיים, וצומת ריק הוא קישור
 *    לכל דבר — כך „מחקתי את הקישור ואני מנסה לתייג מחדש באותה שורה” נכשל
 *    ב-`hyperlink-nested-unsupported` על קישור שכבר לא רואים.
 *
 * נמדד שהסרת הצומת פותרת את שניהם: `hyperlinks.remove` על הכתובת מחזירה
 * `success: true`, ה-`<w:hyperlink>` נעלם מה-OOXML, וההכנסה שאחריה מצליחה.
 *
 * ## למה לא בייצוא
 *
 * `docx-postflight.ts` מצהיר על עצמו כמי ש**מוסיף בלבד** — אף אחד משלושת
 * התיקונים שם אינו מוחק צומת. מחיקת צמתים ב-regex על ה-XML היוצא היא סיכון
 * אחר לגמרי, והיא גם לא הייתה פותרת את החסימה בתוך העורך. הניקוי נעשה במודל,
 * ברגע שממילא כותבים למסמך — כלומר כשהמשתמש מתייג.
 */
import type { DocReceipt, MaybePromise } from './document-api';

/** מיקום בתוך המסמך, כפי ש-`InlineAnchor` מחזיק אותו. */
interface AnchorPoint {
  blockId?: string;
  offset?: number;
}

/** כתובת צומת קישור — הצורה ש-`remove` מקבלת, כפי ש-`list` מחזירה אותה. */
export interface HyperlinkNodeAddress {
  kind: 'inline';
  nodeType: 'hyperlink';
  anchor: { start: AnchorPoint; end: AnchorPoint };
  story?: unknown;
}

interface RawHyperlink {
  text?: string;
  address?: HyperlinkNodeAddress;
}

export interface HyperlinkBlanksDoc {
  hyperlinks?: {
    list?: (input?: Record<string, unknown>) => MaybePromise<unknown>;
    remove?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
  } | null;
}

/**
 * צמצום הניקוי לבלוק שבו עומדים לכתוב.
 *
 * מסלול כתיבת אזכור זקוק רק לצומת הריק שנמצא באותו בלוק; ניקוי המסמך המלא
 * נעשה אחר כך, כשהעורך פנוי. כך מסמך עם הרבה שאריות ישנות אינו נועל את
 * ההקלדה רק מפני שנבחרה הצעת „@”.
 */
export interface BlankHyperlinkScope {
  blockId?: string;
}

/**
 * הקישורים שבתשובת `list`, **ממקור אחד**.
 *
 * `list` מחזיר את אותם קישורים פעמיים: `stories[].hyperlinks[]` היא הצורה
 * המפורטת, ו-`items[]` היא מעטפת ה-discovery הסטנדרטית. איחוד שתיהן היה
 * מכפיל כל צומת, וסינון הכפילויות לפי העוגן **אינו אפשרי**: נמדד ששני
 * צמתים ריקים שונים יושבים על אותו עוגן בדיוק (‏`41964671:0-0` פעמיים, אחרי
 * שנמחק הטקסט של שני קישורים באותה פסקה), וגם `hyperlinkNodeId` שלהם זהה.
 * כלומר אין מפתח שמבדיל ביניהם — ולכן קוראים מקור אחד ולא מנכים ממנו דבר.
 */
function allHyperlinks(raw: unknown): RawHyperlink[] {
  if (!raw || typeof raw !== 'object') return [];
  const source = raw as { items?: unknown; stories?: unknown };

  const fromStories: RawHyperlink[] = [];
  if (Array.isArray(source.stories)) {
    for (const story of source.stories as Array<{ hyperlinks?: unknown }>) {
      if (Array.isArray(story?.hyperlinks)) fromStories.push(...(story.hyperlinks as RawHyperlink[]));
    }
  }
  if (fromStories.length > 0) return fromStories;

  return Array.isArray(source.items) ? (source.items as RawHyperlink[]) : [];
}

/**
 * צומת ריק = עוגן באורך אפס, **וגם** בלי טקסט.
 *
 * העוגן הוא הסימן העיקרי, והוא היחיד שקיים בשתי הצורות שהמנוע מחזיר
 * (ב-`items[]` השדה `text` פשוט נעדר על צומת ריק). תנאי הטקסט מוסיף עליו
 * מקרה אחד בלבד: קישור שהמנוע מדווח עליו `text` **לא-ריק** אינו נמחק גם אם
 * דיווח עוגן באורך אפס.
 *
 * **ומה שהוא אינו מכסה, כי הסרה כאן היא שקטה ועל כל המסמך:** קישור שעוטף
 * צומת שאינו טקסט — תמונה, הפניית הערה — מדווח `text` ריק בדיוק כמו צומת
 * נטוש, ולכן העוגן לבדו מכריע עליו. מה שמחזיק את הקו הוא שלא נמדד קישור כזה
 * עם עוגן באורך אפס, וניסיון לבנות אחד נעצר במנוע: `hyperlinks.wrap` על טווח
 * שכולל הפניית הערה חזרה `INVALID_TARGET` / `TARGET_NOT_FOUND`. אם יימדד כזה,
 * מה שיבדיל אותו הוא `nodeType` של התוכן — לא עוד ניחוש על `text`.
 */
function isBlank(link: RawHyperlink): boolean {
  const anchor = link?.address?.anchor;
  const start = anchor?.start;
  const end = anchor?.end;
  if (typeof link?.text === 'string' && link.text !== '') return false;
  return (
    typeof start?.blockId === 'string' &&
    start.blockId === end?.blockId &&
    typeof start.offset === 'number' &&
    start.offset === end.offset
  );
}

function isInScope(link: RawHyperlink, scope: BlankHyperlinkScope | undefined): boolean {
  if (!scope?.blockId) return true;
  const anchor = link.address?.anchor;
  return anchor?.start?.blockId === scope.blockId && anchor?.end?.blockId === scope.blockId;
}

/**
 * מסירה כל צומת קישור שאין בו טקסט. מחזירה כמה הוסרו.
 *
 * **המסמך ולא הבלוק**, ובמכוון: הקריאה ל-`list` היא ממילא על המסמך כולו,
 * ולכן הסריקה הרחבה אינה עולה קריאת מנוע נוספת — רק קריאות `remove` לפי
 * מספר הצמתים הריקים, שהוא אפס ברוב המסמכים. וצומת ריק אינו נראה לעין בשום
 * מקום: אין לו טקסט, אי אפשר ללחוץ עליו, והסרתו אינה משנה אף פיקסל. הגבלה
 * לבלוק אחד הייתה פותרת את החסימה כאן ומשאירה קישורים בלתי-נראים בקובץ
 * שיוצא — וזה מה שדווח.
 *
 * כשל של `remove` נבלע: הוא לא הרע את המצב, והכתיבה שבדרך תדווח בעצמה אם
 * היא נחסמה.
 */
export async function removeBlankHyperlinks(
  doc: HyperlinkBlanksDoc | null | undefined,
  scope?: BlankHyperlinkScope,
): Promise<number> {
  const list = doc?.hyperlinks?.list;
  const remove = doc?.hyperlinks?.remove;
  if (typeof list !== 'function' || typeof remove !== 'function') return 0;

  let raw: unknown;
  try {
    raw = await list();
  } catch {
    return 0;
  }

  // כל צומת מוסר בקריאה משלו, גם כששניים חולקים עוגן: נמדד ש-`remove` על
  // אותה כתובת פעמיים מסירה שני צמתים ומחזירה `success` בשתי הפעמים. הסרת
  // צומת ריק אינה מזיזה טקסט, ולכן ההיסטים של הנותרים נשארים תקפים לאורך
  // הלולאה.
  let removed = 0;

  for (const link of allHyperlinks(raw)) {
    if (!isBlank(link) || !isInScope(link, scope)) continue;

    try {
      const receipt = await remove({ target: link.address });
      if (receipt?.success !== false) removed += 1;
    } catch {
      /* לא החמרנו — הכתיבה שבדרך תדווח אם היא נחסמה */
    }
  }
  return removed;
}

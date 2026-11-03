/**
 * „כיתובים” — הקבוצה השישית בלשונית „הפניות” של Word העברי, דרך
 * `doc.captions`.
 *
 * כיתוב הוא הפסקה „איור 1: שרטוט המשכן” שמעליה או מתחתיה של תמונה, טבלה או
 * תרשים, ובלב שלה שדה `SEQ` שמספר את הרצף. בספר תורני זו הפעולה שמלווה כל
 * לוח, כל שרטוט וכל טבלת דורות.
 *
 * ## מה נמדד בדפדפן, ולמה הקבוצה בכל זאת נשלחת
 *
 * Chrome headless, `file://`, ה-dist האמיתי, שישה סבבים, וכל אחד מהם פורק גם
 * ברמת ה-zip של `export.toDocx`. זה הלקח של גל 6, ובגל הזה הוא נדרש שוב —
 * כי כאן **שתיים מתוך שש** הפעולות של ה-namespace נפסלו במדידה.
 *
 * ### מה שנכתב נכון, ולכן נשלח
 *
 * `captions.insert` כותב פסקת כיתוב קנונית של Word, אחד לאחד:
 *
 *     <w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>
 *       <w:r><w:t xml:space="preserve">איור </w:t></w:r>
 *       <w:fldSimple w:instr="SEQ איור \* ARABIC"><w:r><w:t>1</w:t></w:r></w:fldSimple>
 *       <w:r><w:t>: שרטוט המשכן</w:t></w:r></w:p>
 *
 * שלושת הדברים שנבדקו כאן הם בדיוק אלה שהפילו את הגלים הקודמים:
 *
 * 1. **התווית בעברית עוברת שלמה אל תוך קוד השדה.** `SEQ איור \* ARABIC` —
 *    לא `Figure`, לא תרגום, ולא רשימה סגורה. התווית היא `label: string`
 *    חופשי, וכל מחרוזת מתקבלת. זה ההפך מ-`TA` של גל 6, שהפיל את טבלת
 *    המקורות בדיוק על הנקודה הזאת.
 * 2. **המספור אמיתי ולפי סדר המסמך.** כיתוב שני באותה תווית מקבל 2,
 *    ומחיקה של הראשון מורידה את השני ל-1 (נמדד גם ב-`list` וגם בערך
 *    ה-cached שבתוך `fldSimple`). כל תווית מנהלת רצף משלה: „טבלה 1” ו„איור 1”
 *    חיים זה לצד זה.
 * 3. **גרשיים ולוכסן בתווית מוברחים כהלכה.** `א"ב` נכתב
 *    `SEQ "א\"ב" \* ARABIC`, ו-`איור \* MERGEFORMAT` נכתב עם `\\*` מוברח —
 *    כלומר אי אפשר להזריק מתג דרך התווית. זה ההפך מ-`TA`, ששם הגרשיים נכתבו
 *    כמות שהם.
 *
 * גם ההסרה נקייה: `captions.remove` מפיל את הפסקה כולה, `blocks.list` יורד
 * בבלוק אחד, ואין שיירים — ההפך מתוכן העניינים של גל 4. הסרה חוזרת על אותה
 * כתובת מוחזרת `TARGET_NOT_FOUND`, וכך גם כתובת של פסקה רגילה: הכתובות
 * מאומתות ואינן נבלעות.
 *
 * ### `captions.update` — לא נשלח כמו שהוא
 *
 * זה הממצא המרכזי של הגל. `update({ patch: { text } })` **אינו מחליף את
 * הטקסט אלא מוסיף עליו**, ונמדד בשלושה צעדים רצופים על אותו כיתוב:
 *
 *     insert 'אלף'   → 'אלף'
 *     update 'בית'   → 'אלף: בית'
 *     update 'גימל'  → 'אלף: בית: גימל'
 *
 * ואומת ב-docx עצמו, לא רק בקבלה: הריצה שאחרי השדה מכילה
 * `<w:t xml:space="preserve">: שרטוט המשכן: </w:t>` — הטקסט הישן, המפריד,
 * והחדש. `update` עם מחרוזת ריקה אינו מוחק אלא מוסיף מפריד ריק. כלומר טופס
 * עריכה שקורא לפעולה הזאת מייצר, בלחיצה אחת, כיתוב שכתוב בו פעמיים.
 *
 * **הדרך שכן נשלחת:** עריכה = `remove` ואז `insert` באותו מקום. העוגן הוא
 * הבלוק שלפני הכיתוב (או שאחריו, כשהוא הראשון במסמך), והוא נקרא **לפני**
 * ההסרה. נמדד שהתוצאה זהה תו-בתו לכיתוב שנוצר מאפס, שהמיקום בתוך המסמך
 * נשמר, ושהמספור מתעדכן — כלומר הקובץ שנכתב הוא קנוני, ולא „כמעט”.
 * המחיר מוצהר, והוא בשני מישורים. **בהיסטוריה:** שני צעדים במקום אחד.
 * **בעוגן:** `captions.insert` מקבל פסקה או טבלה עליונה, ולא כל בלוק.
 * הטבלה נוספה במנוע: עד 2.14.0-next.5 עוגן `tbl:…` הוחזר
 * `TARGET_NOT_FOUND`, וב-2.15.0-next.15 נמדד שהוא מתקבל בשני הכיוונים —
 * ובלבד שהכתובת נושאת `nodeType: 'table'`. לכן הכיתוב שמתחת ללוח, המקרה
 * השכיח, נערך עכשיו במסלול הישיר, וגם „טבלה │ כיתוב │ טבלה”. מה שנשאר:
 * כיתוב שהוא הבלוק היחיד במסמך, וכיתוב ששכניו אינם פסקה ואינם טבלה (תוכן
 * עניינים, למשל — נמדד `INVALID_TARGET`). בשניהם הפעולה מסרבת **לפני**
 * ההסרה ומסבירה.
 *
 * ומעל שניהם רשת ביטחון: הסרה שהצליחה והוספה שנכשלה אחריה מנסה להחזיר את
 * התוכן הישן למקומו, ורק כשגם השחזור נכשל ההודעה אומרת שהכיתוב הוסר
 * ומפנה ל-Ctrl+Z. „נכשלה” לבדה על תוכן שנמחק היא בדיוק סוג ההודעה שהתוסף
 * נבנה כדי לא לומר.
 *
 * ### `captions.configure` — לא נשלח כלל
 *
 * הפעולה אינרטית לגמרי. `configure({ label:'איור', format:'upperRoman' })`
 * חזר `success: true`, והכיתוב הבא נכתב `SEQ איור \* ARABIC` — כלומר
 * ה-`format` אינו מגיע לשום מקום. `format: 'zigzag'`, ערך שאינו בחוזה כלל,
 * חזר גם הוא `success: true`. `includeChapter` הוא היחיד שכן אומר את האמת:
 * `CAPABILITY_UNAVAILABLE / caption-include-chapter-unsupported`.
 *
 * פקד „מספור” שכל מה שהוא עושה הוא להחזיר „בוצע” ולא לשנות דבר גרוע מפקד
 * שאינו קיים, ולכן אין כאן פקד כזה ואין שאלת יכולת בשבילו. המספור נשאר
 * ערבי — מה שגם Word מתחיל בו.
 *
 * ### `images.insertCaption` — לא נשלח
 *
 * שלוש סיבות, וכל אחת מספיקה. **הראשונה:** התווית שלו קשיחה ובאנגלית. אין
 * ב-`InsertCaptionInput` שדה `label` בכלל (`{ imageId, text }`), והמימוש
 * במנוע כותב `SEQ Figure` וטקסט `Figure <n> <text>` — בלי נקודתיים ובלי
 * `\* ARABIC`. כלומר אין דרך לקבל ממנו „איור”. **השנייה:** `create.image`
 * אינה זמינה במנוע כלל, ו-`doc.insert` של HTML עם `<img src="data:…">`
 * נדחה („HTML produced no safe canonical content”) — כלומר התוסף אינו יכול
 * להביא תמונה למסמך מלכתחילה. **השלישית:** `captions.insert` עושה את אותו
 * דבר טוב יותר — הוא מקבל תווית, והעוגן שלו הוא הפסקה, כולל הפסקה שהתמונה
 * יושבת בה.
 *
 * ### „טבלת איורים” — אינה כאן, ולא בגלל שהיא שבורה
 *
 * הפקד השני בקבוצה של Word הוא „הוסף טבלת איורים”, והמסלול אליו הוא
 * `create.tableOfContents({ instruction: 'TOC \c "איור" \h \z' })`. הוא
 * נמדד ועובד: השדה נכתב קנונית לתוך ה-sdt של תוכן העניינים, ו-`toc.list`
 * מחזיר אותו עם `preserved.seqFieldIdentifier: 'איור'`. מה שהמנוע **אינו**
 * עושה הוא לאסוף את הכיתובים — `entryCount` נשאר 0 גם אחרי `toc.update`,
 * ו-Word הוא שימלא אותה בפתיחה. הפקד עצמו שייך ל-engine/toc.ts, שאינו
 * בהיקף הגל הזה, ולכן הוא מדווח ולא נכתב.
 *
 * ## מה המנוע בולע בשקט, ולכן נבדק כאן
 *
 * | מה נשלח | מה נכתב |
 * |---|---|
 * | `label: '   '` | `SEQ "   " \* ARABIC` — שדה עם תווית של רווחים |
 * | `label: 'אי\nור'` | ירידת שורה **גולמית בתוך קוד השדה** |
 * | `text: '   '` | `<w:t xml:space="preserve">:    </w:t>`, ו-`get` מחזיר `''` |
 * | `text: 'שורה\nשנייה'` | ירידת שורה גולמית בתוך `<w:t>` |
 * | `text: 5` | `success: true`, והטקסט נעלם בלי זכר |
 *
 * `label: ''` הוא היחיד שהמנוע **זורק** עליו („requires a non-empty label
 * string”) ואינו בולע. כלומר שתי ההתנהגויות קיימות זו לצד זו, ומהן נגזרת
 * אותה מסקנה: הוולידציה יושבת כאן, לפני הקריאה.
 *
 * ## כיוון הפסקה — צעד שהמנוע אינו עושה
 *
 * פסקת הכיתוב שהמנוע כותב היא `<w:pPr><w:pStyle w:val="Caption"/></w:pPr>`
 * ותו לא — **בלי `<w:bidi/>`**, בעוד שהפסקה הרגילה שלצידה במסמך העברי כן
 * נושאת אותו. כיתוב כזה ייפתח ב-Word כפסקה משמאל לימין, והשורה „איור 1:
 * שרטוט המשכן” תיערך בסדר שאינו הסדר שהמשתמש הקליד.
 *
 * לכן כל כיתוב שנוצר כאן מקבל אחריו `paragraphs.setDirection({direction:'rtl'})`,
 * שנמדד ככותב `<w:bidi/>` לתוך אותה `pPr` ומשאיר את סגנון ה-`Caption` על
 * מקומו. הצעד הזה **אינו** מפיל את הפעולה כשהוא נכשל, וזו הכרעה מודעת:
 * הכיתוב כבר במסמך והוא תקין, וההודעה „הוספת הכיתוב נכשלה” על כיתוב שנמצא
 * על המסך הייתה שקר. `NO_OP` הוא התשובה כשהפסקה כבר ימין-לשמאל.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';

/* ------------------------------------------------------------------ */
/* צורת ה-API. מוגדרת כאן ואינה מיובאת — ראו document-api.ts            */
/* ------------------------------------------------------------------ */

/**
 * סוגי הבלוקים שהמנוע מקבל כעוגן ל-`captions.insert` — ורק הם.
 *
 * נמדד על 2.15.0-next.15: טבלה התקבלה כעוגן בשני הכיוונים, ובלוק מסוג
 * `tableOfContents` הוחזר `INVALID_TARGET` עם נוסח המנוע עצמו — „paragraph-like
 * blocks or top-level tables”. שליחת אותו בלוק כ„פסקה” מחזירה `success: true`
 * וכותבת את הכיתוב **בתוך** ה-sdt, כלומר מסמך שגוי בשקט. לכן הרשימה סגורה,
 * ואינה „כל מה שאינו פסקה”.
 */
type AnchorNodeType = 'paragraph' | 'table';

/** `CaptionAddress` — מה שכל פעולת כיתוב מקבלת כ-`target`. */
interface CaptionAddress {
  kind: 'block';
  nodeType: AnchorNodeType;
  nodeId: string;
}

/** `CaptionDomain` עטוף ב-`DiscoveryItem`. */
interface RawCaption {
  address?: { nodeId?: string };
  label?: string;
  number?: number;
  text?: string;
}

/**
 * `BlockNodeInfo` בחלק שנצרך כאן: „מי בא לפני מי”, ו**מאיזה סוג**.
 *
 * `nodeType` אינו קישוט: הכתובת חייבת לשאת את הסוג **האמיתי** של הבלוק.
 * נמדד: טבלה שנשלחה כ-`paragraph` הוחזרה `TARGET_NOT_FOUND`, פסקה שנשלחה
 * כ-`table` הוחזרה `TARGET_NOT_FOUND`, וכתובת בלי `nodeType` הוחזרה
 * `INVALID_TARGET`. ומעבר לזה: סוג שאינו ברשימת ה-`AnchorNodeType` חייב
 * להיחסם לפני ההסרה, אחרת ההסרה כבר מחקה את הכיתוב.
 */
interface RawBlock {
  nodeId?: string;
  nodeType?: string;
}

interface DiscoveryPage<T> {
  items?: readonly T[];
  total?: number;
}

/**
 * `blocks.list` מחזיר מעטפה **אחרת** מכל שאר ה-discovery: `blocks` ולא
 * `items`. נמדד — ולא ניחוש: הקריאה הראשונה שנכתבה כאן קראה `items` וקיבלה
 * `undefined` על מסמך שיש בו פסקה.
 */
interface BlocksPage {
  blocks?: readonly RawBlock[];
  total?: number;
}

export interface CaptionsDocumentApi {
  selection?: {
    current?: () => MaybePromise<
      { target?: { segments?: readonly { blockId?: string }[] } | null } | undefined
    >;
  };
  blocks?: {
    list?: (query?: { limit?: number; offset?: number }) => MaybePromise<BlocksPage | undefined>;
  };
  paragraphs?: {
    setDirection?: (input: {
      target: CaptionAddress;
      direction: 'rtl' | 'ltr';
    }) => MaybePromise<DocReceipt>;
  };
  captions?: {
    list?: (query?: {
      label?: string;
      limit?: number;
      offset?: number;
    }) => MaybePromise<DiscoveryPage<RawCaption> | undefined>;
    insert?: (input: {
      adjacentTo: CaptionAddress;
      position: 'above' | 'below';
      label: string;
      text?: string;
    }) => MaybePromise<DocReceipt>;
    remove?: (input: { target: CaptionAddress }) => MaybePromise<DocReceipt>;
  };
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו page-setup.ts. */
export interface CaptionsHost {
  activeEditor?: { doc?: CaptionsDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type CaptionsTarget = SuperDoc | CaptionsHost | null | undefined;

/* ------------------------------------------------------------------ */
/* נוסחי הכשל                                                          */
/* ------------------------------------------------------------------ */

/**
 * ביטוי שלם עם הטיית הכשל בעברית תקנית, ולא שם עצם: „הכיתוב” זכר. ראו
 * document-api.ts.
 */
const INSERT_FAILED = 'הוספת הכיתוב נכשלה';
const UPDATE_FAILED = 'עריכת הכיתוב נכשלה';
const REMOVE_FAILED = 'הסרת הכיתוב נכשלה';
const READ_FAILED = 'קריאת הכיתובים נכשלה';

const LOADING_DETAIL = 'המסמך עדיין נטען';
const NO_ANCHOR_DETAIL = 'יש למקם את הסמן בפסקה שלידה יופיע הכיתוב';
const NOT_FOUND_DETAIL = 'הכיתוב אינו נמצא במסמך';

/**
 * המצב היחיד שבו עריכה מסרבת: כיתוב שהוא הבלוק היחיד במסמך. העריכה היא
 * הסרה והוספה מחדש (ראו הערת הפתיחה), וההוספה דורשת בלוק שכן להיצמד אליו.
 * בלי הסירוב הזה ההסרה הייתה מצליחה וההוספה נכשלת — כלומר טקסט שנעלם.
 */
const NO_NEIGHBOUR_DETAIL =
  'הכיתוב הוא הפסקה היחידה במסמך, ואין לידה פסקה שאליה אפשר להצמיד אותו מחדש';

/**
 * המצב השני שבו עריכה מסרבת: כיתוב ששני שכניו אינם פסקה ואינם טבלה — למשל
 * כיתוב שנצמד לתוכן עניינים.
 *
 * טבלה **אינה** נופלת לכאן מאז 2.15.0-next.15 (ראו `AnchorNodeType`), אבל
 * בלוק מסוג `tableOfContents` נמדד `INVALID_TARGET`, ושליחתו כ„פסקה” כותבת
 * את הכיתוב בתוך ה-sdt. בלי הסירוב הזה ההסרה מצליחה, ההוספה נכשלת או כותבת
 * במקום הלא נכון, והכיתוב שהמשתמש ביקש לערוך פשוט נעלם.
 */
const NOT_ANCHORABLE_DETAIL =
  'הבלוק שליד הכיתוב אינו פסקה ואינו טבלה, ואי אפשר להצמיד אליו את הכיתוב מחדש';

/**
 * הנוסח היחיד שמודה באובדן. הוא נאמר רק כשההסרה הצליחה, ההוספה נכשלה,
 * וגם השחזור נכשל — ואז „נכשלה” לבדה הייתה שקר: המשתמש היה חוזר למסמך
 * לחפש כיתוב שכבר אינו שם.
 */
const LOST_MESSAGE = `${UPDATE_FAILED}, והכיתוב הוסר מהמסמך. אפשר להחזירו בעזרת ביטול (Ctrl+Z).`;

function unavailable(failedAction: string, detail: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${detail}`, reason };
}

/** הנוסח שהתכנית קובעת ב-§12, וזהה לזה שהיכולת מחזירה. ראו footnotes.ts. */
function unsupported(failedAction: string): CommandOutcome {
  return {
    ok: false,
    message: `${failedAction}: אינו זמין בגרסה זו`,
    reason: 'command-unsupported',
  };
}

/** קריאה למנוע שאינה זורקת החוצה. ראו הערת ה„לעולם לא זורקת” ב-footnotes.ts. */
async function attempt<T>(
  failedAction: string,
  call: () => MaybePromise<T>,
): Promise<{ ok: true; value: T } | { ok: false; outcome: CommandOutcome }> {
  try {
    return { ok: true, value: await call() };
  } catch (error) {
    return {
      ok: false,
      outcome: { ok: false, message: thrownText(failedAction, error), reason: 'threw' },
    };
  }
}

/** כשל הקבלה, או `null` כשהיא הצליחה. `NO_OP` נחשב הצלחה — ראו header-footer.ts. */
function failureOf(failedAction: string, receipt: DocReceipt | undefined): CommandOutcome | null {
  const code = receipt?.failure?.code;
  if (receipt?.success !== false || code === 'NO_OP') return null;
  return { ok: false, message: receiptFailureText(failedAction, receipt), reason: code };
}

function docOf(host: CaptionsTarget): CaptionsDocumentApi | null {
  return (host as CaptionsHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** גודל העמוד בכל שאיבה, ובלם מפני מנוע שיחזיר `total` שאינו יורד. ראו fields.ts. */
const PAGE_SIZE = 200;
const PAGE_GUARD = 1000;

/**
 * כל הפריטים של פעולת discovery, בשאיבת עמודים עד `total`.
 *
 * `total` ולא `items.length`: `captions.list` הוא `DiscoveryOutput`, כלומר
 * `items` הוא עמוד תחת `limit`/`offset`. ספר תורני עם מאות לוחות ושרטוטים
 * הוא בדיוק המסמך שבו העמוד הראשון אינו הכול — וזה התרחיש של אוצריא, לא
 * מקרה קצה.
 */
async function collectPages<T>(
  failedAction: string,
  list: (query: { limit: number; offset: number }) => MaybePromise<
    { total?: number; page?: readonly T[] } | undefined
  >,
): Promise<{ ok: boolean; items: T[]; outcome?: CommandOutcome }> {
  const items: T[] = [];
  let offset = 0;
  let guard = 0;

  for (;;) {
    const listed = await attempt(failedAction, () => list({ limit: PAGE_SIZE, offset }));
    if (!listed.ok) return { ok: false, items, outcome: listed.outcome };

    const page = listed.value?.page ?? [];
    items.push(...page);
    if (page.length === 0) return { ok: true, items };

    offset += page.length;

    const total = listed.value?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) return { ok: true, items };
    if (++guard > PAGE_GUARD) return { ok: true, items };
  }
}

/* ------------------------------------------------------------------ */
/* ולידציה — כאן ולא במנוע                                             */
/* ------------------------------------------------------------------ */

/**
 * התוויות המובנות של Word העברי, בסדר שבו הן מופיעות שם.
 *
 * שלוש ולא יותר, וזו הרשימה של Word עצמו. היא **אינה** רשימה סגורה: התווית
 * היא מחרוזת חופשית במנוע (נמדד — `SEQ איור`, `SEQ טבלה` ו-`SEQ א"ב` נכתבו
 * כולם), ולכן הדיאלוג מציע את השלוש ומאפשר להקליד אחרת. ספר תורני שירצה
 * „לוח” או „תרשים” יקבל אותם.
 */
export const CAPTION_LABELS = ['איור', 'טבלה', 'משוואה'] as const;

export const DEFAULT_CAPTION_LABEL = CAPTION_LABELS[0];

/** מה שמוצג כשהתווית נדחית. */
export const CAPTION_LABEL_HINT = 'יש להקליד תווית לכיתוב, בשורה אחת';

/**
 * התווית כפי שתישלח, או `null` כשהיא פסולה.
 *
 * שני דברים נבדקים, ושניהם נמדדו. ה-`trim` הוא אותה מלכודת של `XE "   "`
 * מגל 5: המנוע מקבל `label: '   '` בהצלחה וכותב `SEQ "   " \* ARABIC` —
 * שדה שאי אפשר לזהות ואי אפשר למצוא. ירידת השורה חמורה יותר: היא נכתבת
 * **גולמית לתוך קוד השדה** (`SEQ "אי\nור"`), וקוד שדה של Word אינו יכול
 * להכיל אותה. שניהם מוחזרים `success: true`.
 */
export function normalizeCaptionLabel(raw: string): string | null {
  if (typeof raw !== 'string') return null;
  const label = raw.trim();
  if (label === '') return null;
  return /[\n\r\t]/.test(label) ? null : label;
}

/**
 * טקסט הכיתוב כפי שיישלח.
 *
 * ריק הוא ערך חוקי — Word מתיר „איור 1” בלי תיאור, והמנוע כותב אז פסקה
 * שנגמרת בשדה. מה שאינו חוקי הוא ירידת שורה, שנמדדה נכתבת גולמית לתוך
 * `<w:t>`; כיתוב הוא שורה אחת מעצם הגדרתו, ולכן היא מכווצת לרווח ולא
 * חוסמת את הפעולה. רווחים בקצוות יורדים — המנוע כותב `: ` ואז ארבעה
 * רווחים, ו-`get` מחזיר על אותו כיתוב מחרוזת ריקה.
 */
export function normalizeCaptionText(raw: string): string {
  if (typeof raw !== 'string') return '';
  return raw.replace(/[\n\r\t]+/g, ' ').trim();
}

/* ------------------------------------------------------------------ */
/* המודל שהממשק עובד מולו                                              */
/* ------------------------------------------------------------------ */

/** איפה הכיתוב יושב ביחס לפסקה שאליה הוא נצמד. שני הערכים של Word. */
export type CaptionPosition = 'above' | 'below';

/** מה שהטופס מחזיק. מחרוזות בלבד — זו הצורה שקלט טקסט נותן. */
export interface CaptionDraft {
  label: string;
  text: string;
  position: CaptionPosition;
}

export function emptyCaptionDraft(): CaptionDraft {
  return { label: DEFAULT_CAPTION_LABEL, text: '', position: 'below' };
}

/** כיתוב אחד כפי שהדיאלוג מציג אותו. */
export interface CaptionSummary {
  /** `nodeId` של פסקת הכיתוב. גם מפתח `v-for` וגם היעד של כל פעולה. */
  id: string;
  label: string;
  /** המספר שהמנוע חישב ברצף של אותה תווית. */
  number: number;
  text: string;
  /** מה שמוצג ברשימה: „איור 1: שרטוט המשכן”, כמו שזה נראה במסמך. */
  display: string;
}

/** מה שהממשק צריך לדעת. תצלום ולא מנוי, כמו header-footer.ts. */
export interface CaptionsState {
  captions: readonly CaptionSummary[];
  /**
   * התוויות שכבר בשימוש במסמך, בלי כפילות. הדיאלוג מציע אותן לצד השלוש
   * המובנות: מסמך שכבר יש בו „לוח” לא אמור לכפות על המשתמש להקליד אותו שוב
   * בכל כיתוב, וטעות הקלדה שם מייצרת רצף מספור שני ונפרד.
   */
  labels: readonly string[];
}

export function emptyCaptionsState(): CaptionsState {
  return { captions: [], labels: [] };
}

/** „איור 1: שרטוט המשכן” — בדיוק מה שנכתב בפסקה שבמסמך. */
export function captionDisplay(label: string, number: number, text: string): string {
  const head = `${label} ${number}`;
  return text === '' ? head : `${head}: ${text}`;
}

/**
 * כתובת של בלוק. ברירת המחדל `paragraph`, וזה הסוג של כל **יעד** כאן —
 * `captions.remove` ו-`paragraphs.setDirection` מקבלים את פסקת הכיתוב עצמה.
 * העוגן של `captions.insert` הוא היחיד שיכול להיות גם טבלה.
 */
function addressOf(nodeId: string, nodeType: AnchorNodeType = 'paragraph'): CaptionAddress {
  return { kind: 'block', nodeType, nodeId };
}

/* ------------------------------------------------------------------ */
/* קריאה                                                               */
/* ------------------------------------------------------------------ */

/**
 * הכיתובים שבמסמך, בסדר הופעתם. לעולם אינה זורקת: כשל של קריאה מחזיר רשימה
 * ריקה, כלומר הדיאלוג יאמר „אין כיתובים” — ולא ימציא רשומות.
 *
 * כיתוב בלי `nodeId` מדולג ואינו מוחזר: הצגתו הייתה מייצרת שורה שלחיצה
 * עליה שולחת `undefined` ל-`captions.remove`.
 */
async function collectCaptions(
  host: CaptionsTarget,
  failedAction: string,
): Promise<{ ok: boolean; captions: CaptionSummary[]; outcome?: CommandOutcome }> {
  const list = docOf(host)?.captions?.list;
  if (typeof list !== 'function') {
    return { ok: false, captions: [], outcome: unsupported(failedAction) };
  }

  const listed = await collectPages<RawCaption>(failedAction, async (query) => {
    const page = await list(query);
    return { total: page?.total, page: page?.items };
  });

  const captions: CaptionSummary[] = [];
  for (const raw of listed.items) {
    const id = raw.address?.nodeId;
    if (typeof id !== 'string' || id === '') continue;
    const label = typeof raw.label === 'string' ? raw.label : '';
    const number = typeof raw.number === 'number' ? raw.number : 0;
    const text = typeof raw.text === 'string' ? raw.text : '';
    captions.push({ id, label, number, text, display: captionDisplay(label, number, text) });
  }
  return { ok: listed.ok, captions, outcome: listed.outcome };
}

/** „רשימת הכיתובים” לדיאלוג. */
export async function listCaptions(host: CaptionsTarget): Promise<CaptionSummary[]> {
  return (await collectCaptions(host, READ_FAILED)).captions;
}

/**
 * קוראת את מצב הכיתובים במסמך. לעולם אינה זורקת: כשל של קריאה מחזיר „אין”,
 * כלומר ה-tooltip יאמר שאין מה לערוך — ולא ימציא מספר.
 */
export async function readCaptionsState(host: CaptionsTarget): Promise<CaptionsState> {
  const doc = docOf(host);
  if (!doc) return emptyCaptionsState();

  const listed = await collectCaptions(host, READ_FAILED);
  const labels: string[] = [];
  for (const caption of listed.captions) {
    if (caption.label !== '' && !labels.includes(caption.label)) labels.push(caption.label);
  }
  return { captions: listed.captions, labels };
}

/* ------------------------------------------------------------------ */
/* עוגנים                                                              */
/* ------------------------------------------------------------------ */

/**
 * הפסקה שהסמן יושב בה, או כשל מנומק.
 *
 * המקטע ה**אחרון** התקין ולא הראשון, מאותו טעם כמו ב-citations.ts: בחירה
 * שחוצה פסקאות מדווחת מקטע לכל פסקה, וסופה של הבחירה הוא סוף האחרון. כאן
 * זה גם מה ש-Word עושה — הכיתוב נצמד לאובייקט שבסוף הבחירה.
 */
async function caretBlock(
  host: CaptionsTarget,
  failedAction: string,
): Promise<{ ok: true; nodeId: string } | { ok: false; outcome: CommandOutcome }> {
  const current = docOf(host)?.selection?.current;
  if (typeof current !== 'function') return { ok: false, outcome: unsupported(failedAction) };

  const read = await attempt(failedAction, () => current());
  if (!read.ok) return { ok: false, outcome: read.outcome };

  let nodeId: string | undefined;
  for (const segment of read.value?.target?.segments ?? []) {
    const blockId = segment?.blockId;
    if (typeof blockId === 'string' && blockId !== '') nodeId = blockId;
  }
  if (nodeId !== undefined) return { ok: true, nodeId };

  return { ok: false, outcome: unavailable(failedAction, NO_ANCHOR_DETAIL, 'no-anchor') };
}

/** בלוק אחד כפי שהעריכה צריכה אותו: מזהה, וסוג שקובע אם אפשר להיצמד אליו. */
interface OrderedBlock {
  nodeId: string;
  nodeType: string;
}

/** בלוקי המסמך, בסדרם. ראו `BlocksPage` על צורת המעטפה. */
async function blockOrder(
  host: CaptionsTarget,
  failedAction: string,
): Promise<{ ok: true; blocks: OrderedBlock[] } | { ok: false; outcome: CommandOutcome }> {
  const list = docOf(host)?.blocks?.list;
  if (typeof list !== 'function') return { ok: false, outcome: unsupported(failedAction) };

  const listed = await collectPages<RawBlock>(failedAction, async (query) => {
    const page = await list(query);
    return { total: page?.total, page: page?.blocks };
  });
  if (!listed.ok) return { ok: false, outcome: listed.outcome ?? unsupported(failedAction) };

  const blocks: OrderedBlock[] = [];
  for (const block of listed.items) {
    if (typeof block.nodeId !== 'string' || block.nodeId === '') continue;
    // סוג שאינו מחרוזת נחשב „לא ניתן להיצמדות”, ולכן יחסום את העריכה: מנוע
    // שאינו מדווח סוג אינו נותן ראיה שאפשר להיצמד לבלוק, וניחוש כאן הוא מחיקה.
    blocks.push({ nodeId: block.nodeId, nodeType: typeof block.nodeType === 'string' ? block.nodeType : '' });
  }
  return { ok: true, blocks };
}

/** הסוג שהעוגן יישלח בו, או `null` כשהמנוע אינו מקבל בלוק כזה כעוגן. */
function anchorTypeOf(block: OrderedBlock | null): AnchorNodeType | null {
  if (block?.nodeType === 'paragraph') return 'paragraph';
  if (block?.nodeType === 'table') return 'table';
  return null;
}

/**
 * העוגן שמחזיר כיתוב שהוסר אל אותו מקום בדיוק, או סירוב מנומק.
 *
 * הבלוק שלפניו ו-`below`, ואם הוא הראשון — הבלוק שאחריו ו-`above`. נקרא
 * **לפני** ההסרה, כי אחריה הכיתוב כבר אינו ברשימה ואי אפשר לדעת איפה היה.
 *
 * טבלה היא עוגן קביל, וזה **שינוי במנוע**: עד 2.14.0-next.5 הוחזר על
 * `tbl:…` ‏`TARGET_NOT_FOUND`, ולכן הכיתוב שמתחת ללוח — הצורה השכיחה
 * ביותר, וזו שמגיעה מכל docx מיובא — נצמד חזרה אל הפסקה שאחריו. ב-
 * 2.15.0-next.15 נמדד שעוגן טבלה מתקבל בשני הכיוונים, ובלבד שהכתובת נושאת
 * `nodeType: 'table'` ולא `'paragraph'`. לכן הכיתוב חוזר עכשיו אל הלוח
 * עצמו, ו„טבלה │ כיתוב │ טבלה” — שסורב עד כה — נערך.
 *
 * הנפילה-לאחור אל הבלוק שאחרי נשארת בשביל מה שעדיין נדחה (תוכן עניינים,
 * למשל), ואחריה הסירוב — לפני שנגעו במסמך.
 */
function neighbourAnchor(
  blocks: readonly OrderedBlock[],
  nodeId: string,
):
  | { ok: true; anchor: CaptionAddress; position: CaptionPosition }
  | { ok: false; detail: string; reason: string } {
  const index = blocks.findIndex((block) => block.nodeId === nodeId);
  if (index === -1) return { ok: false, detail: NOT_FOUND_DETAIL, reason: 'caption-not-found' };

  const before = index > 0 ? blocks[index - 1] : null;
  const after = index < blocks.length - 1 ? blocks[index + 1] : null;

  // שני העוגנים מחזירים את הכיתוב לאותו רווח בדיוק: „מתחת לקודם” ו„מעל
  // הבא” הם אותו מקום. נמדד בדפדפן על `פסקה │ tbl │ כיתוב │ טבלה` — ההסרה
  // וההוספה מחדש בעוגן הטבלה שלפני החזירו את הכיתוב בין שתי הטבלאות, עם
  // אותו מספר. התיעוד ב-docs/engine-gaps.md.
  const beforeType = anchorTypeOf(before);
  if (before && beforeType) {
    return { ok: true, anchor: addressOf(before.nodeId, beforeType), position: 'below' };
  }
  const afterType = anchorTypeOf(after);
  if (after && afterType) {
    return { ok: true, anchor: addressOf(after.nodeId, afterType), position: 'above' };
  }

  if (!before && !after) return { ok: false, detail: NO_NEIGHBOUR_DETAIL, reason: 'no-neighbour' };
  return { ok: false, detail: NOT_ANCHORABLE_DETAIL, reason: 'anchor-kind-unsupported' };
}

/* ------------------------------------------------------------------ */
/* כתיבה                                                               */
/* ------------------------------------------------------------------ */

/**
 * מסובבת את פסקת הכיתוב לימין-לשמאל.
 *
 * אינה מחזירה כשל בכוונה — ראו „כיוון הפסקה” בהערת הפתיחה: הכיתוב כבר
 * במסמך בשלב הזה, והודעת כשל עליו הייתה שולחת את המשתמש לחפש משהו שאינו
 * קיים. פעולה חסרה או קבלה שנכשלה משאירות כיתוב בכיוון ברירת המחדל, וזה
 * מצב שאפשר לתקן בעורך.
 */
async function faceRtl(host: CaptionsTarget, receipt: DocReceipt | undefined): Promise<void> {
  const nodeId = (receipt as { caption?: { nodeId?: string } } | undefined)?.caption?.nodeId;
  if (typeof nodeId !== 'string' || nodeId === '') return;

  const setDirection = docOf(host)?.paragraphs?.setDirection;
  if (typeof setDirection !== 'function') return;

  try {
    await setDirection({ target: addressOf(nodeId), direction: 'rtl' });
  } catch {
    /* כיוון פסקה אינו שווה הפלה של כיתוב שכבר נכתב */
  }
}

/** הבדיקות שקודמות לכל שליחה של כיתוב. ראו „מה המנוע בולע בשקט”. */
function validateDraft(
  failedAction: string,
  draft: CaptionDraft,
): { ok: true; label: string; text: string } | { ok: false; outcome: CommandOutcome } {
  const label = normalizeCaptionLabel(draft.label);
  if (label === null) {
    return {
      ok: false,
      outcome: { ok: false, message: `${failedAction}: ${CAPTION_LABEL_HINT}`, reason: 'invalid-label' },
    };
  }
  return { ok: true, label, text: normalizeCaptionText(draft.text) };
}

/** הצעד המשותף להוספה ולעריכה: כתיבת הכיתוב, ואז סיבוב הפסקה. */
async function writeCaption(
  host: CaptionsTarget,
  failedAction: string,
  input: { adjacentTo: CaptionAddress; position: CaptionPosition; label: string; text: string },
): Promise<CommandOutcome> {
  const insert = docOf(host)?.captions?.insert;
  if (typeof insert !== 'function') return unsupported(failedAction);

  const inserted = await attempt(failedAction, () =>
    insert({
      adjacentTo: input.adjacentTo,
      position: input.position,
      label: input.label,
      text: input.text,
    }),
  );
  if (!inserted.ok) return inserted.outcome;

  const failure = failureOf(failedAction, inserted.value);
  if (failure) return failure;

  await faceRtl(host, inserted.value);
  return { ok: true };
}

/**
 * „הוסף כיתוב” — מוסיפה פסקת כיתוב מעל הפסקה שהסמן בה או מתחתיה.
 *
 * העוגן הוא הפסקה שבסמן ולא האובייקט שנבחר, וזה ההבדל היחיד מ-Word: אין
 * ב-Document API דרך לשאול על מה הסמן עומד מלבד `blockId` (ראו
 * docs/engine-gaps.md). מעשית זה אותו דבר — תמונה או טבלה בפסקה משלה, והסמן
 * בה — ולמשתמש נשאר לבחור „מעל” או „מתחת”.
 */
export async function insertCaption(
  host: CaptionsTarget,
  draft: CaptionDraft,
): Promise<CommandOutcome> {
  const valid = validateDraft(INSERT_FAILED, draft);
  if (!valid.ok) return valid.outcome;

  const doc = docOf(host);
  if (!doc) return unavailable(INSERT_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const caret = await caretBlock(host, INSERT_FAILED);
  if (!caret.ok) return caret.outcome;

  // הסמן יושב תמיד בפסקה — גם כשהיא פסקה בתוך תא של טבלה (נמדד: לחיצה
  // בתוך לוח מדווחת את פסקת התא, לא את הלוח). לכן הכתובת כאן היא פסקה.
  return writeCaption(host, INSERT_FAILED, {
    adjacentTo: addressOf(caret.nodeId),
    position: draft.position === 'above' ? 'above' : 'below',
    label: valid.label,
    text: valid.text,
  });
}

/**
 * „ערוך כיתוב” — מחליפה את התווית ואת הטקסט של כיתוב קיים.
 *
 * הסרה והוספה מחדש, ולא `captions.update`. זו ההכרעה המרכזית של הגל, והיא
 * נמדדה: `update` מוסיף את הטקסט החדש על הישן במקום להחליף אותו, ובלחיצה
 * שנייה מייצר „אלף: בית: גימל”. ההנמקה המלאה, כולל מה שנכתב ל-docx, בהערת
 * הפתיחה.
 *
 * העוגן נקרא לפני ההסרה, והסירוב על שכן שאינו קביל קודם לה: הסרה
 * שהצליחה והוספה שנכשלה אחריה היא טקסט שנמחק בלי דרך חזרה, וזה בדיוק המצב
 * שהסדר הזה מונע. שני סירובים כאלה: כיתוב בלי שכן כלל, ושכן שאינו פסקה
 * ואינו טבלה — תוכן עניינים, למשל.
 *
 * ואם ההוספה בכל זאת נכשלה אחרי הסרה שהצליחה, התוכן הישן מוחזר למקומו,
 * ורק שחזור שגם הוא נכשל מדווח כאובדן. ראו `LOST_MESSAGE`.
 */
export async function updateCaption(
  host: CaptionsTarget,
  captionId: string,
  draft: CaptionDraft,
): Promise<CommandOutcome> {
  const valid = validateDraft(UPDATE_FAILED, draft);
  if (!valid.ok) return valid.outcome;

  const doc = docOf(host);
  if (!doc) return unavailable(UPDATE_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const remove = doc.captions?.remove;
  if (typeof remove !== 'function' || typeof doc.captions?.insert !== 'function') {
    return unsupported(UPDATE_FAILED);
  }

  const order = await blockOrder(host, UPDATE_FAILED);
  if (!order.ok) return order.outcome;

  const spot = neighbourAnchor(order.blocks, captionId);
  if (!spot.ok) return unavailable(UPDATE_FAILED, spot.detail, spot.reason);

  // התוכן הישן נקרא לפני ההסרה, ורק בשבילה: אם ההוספה תיכשל, זה מה
  // שיוחזר למסמך. אחרי ההסרה כבר אין מאיפה לקרוא אותו.
  const original = (await collectCaptions(host, UPDATE_FAILED)).captions.find(
    (caption) => caption.id === captionId,
  );

  const removed = await attempt(UPDATE_FAILED, () => remove({ target: addressOf(captionId) }));
  if (!removed.ok) return removed.outcome;
  const failure = failureOf(UPDATE_FAILED, removed.value);
  if (failure) return failure;

  const written = await writeCaption(host, UPDATE_FAILED, {
    adjacentTo: spot.anchor,
    position: spot.position,
    label: valid.label,
    text: valid.text,
  });
  if (written.ok) return written;

  // ההסרה כבר קרתה. שחזור מוצלח מחזיר את המסמך למה שהיה, ואז „העריכה
  // נכשלה” היא האמת השלמה; שחזור שנכשל מחייב לומר שהכיתוב אבד — הודעת
  // „נכשלה” לבדה על תוכן שנמחק היא בדיוק מה שהתוסף נבנה כדי לא לעשות.
  // תווית ריקה אינה משוחזרת: `captions.insert` **זורק** עליה (נמדד), כלומר
  // ניסיון שחזור כזה הוא רעש בלבד.
  if (original && original.label !== '') {
    const restored = await writeCaption(host, UPDATE_FAILED, {
      adjacentTo: spot.anchor,
      position: spot.position,
      label: original.label,
      text: original.text,
    });
    if (restored.ok) return written;
  }
  return { ok: false, message: LOST_MESSAGE, reason: 'caption-lost' };
}

/**
 * „הסר כיתוב” — מוחקת את פסקת הכיתוב.
 *
 * צעד אחד, בלי ניקוי שיירים: נמדד שאחרי `remove` הפסקה כולה יורדת
 * מ-`blocks.list` — ההפך מתוכן העניינים של גל 4. שאר הכיתובים באותה תווית
 * מתמספרים מחדש מיד, גם בערך שבתוך השדה.
 */
export async function removeCaption(
  host: CaptionsTarget,
  captionId: string,
): Promise<CommandOutcome> {
  if (typeof captionId !== 'string' || captionId === '') {
    return unavailable(REMOVE_FAILED, 'יש לבחור כיתוב', 'no-caption');
  }

  const doc = docOf(host);
  if (!doc) return unavailable(REMOVE_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const remove = doc.captions?.remove;
  if (typeof remove !== 'function') return unsupported(REMOVE_FAILED);

  const removed = await attempt(REMOVE_FAILED, () => remove({ target: addressOf(captionId) }));
  if (!removed.ok) return removed.outcome;

  return failureOf(REMOVE_FAILED, removed.value) ?? { ok: true };
}

/**
 * „לבדוק capability בעת boot” (docs/remaining-waves.md)
 * מול הדרך הקנונית שהמנוע נותן לזה: `doc.capabilities.get()`.
 *
 * למה עטיפה ולא קריאה ישירה מכל פקד:
 *
 * 1. **הצורה אינה הצורה שהממשק שואל.** המנוע מחזיר
 *    `operations: Record<OperationId, {available, tracked, dryRun, reasons}>`
 *    ו-`global` — כלומר תשובה לפי מזהה פעולה. פקד ב-Ribbon שואל שאלה אחרת:
 *    „האם אפשר להוסיף הערת שוליים”. פקד שיתרגם את זה בעצמו יקודד את מזהה
 *    הפעולה של המנוע לתוך התבנית, וכל שינוי בגרסה יתפזר על עשרות אתרי קריאה.
 * 2. **חלק מהיכולות תלויות בשני מקומות.** „תגובה חדשה” דורשת גם
 *    `operations['comments.create'].available` וגם `global.comments.enabled`;
 *    הראשון אומר שהפעולה קיימת, השני שהיא מאופשרת במסמך הזה.
 * 3. **נכשל סגור.** אם אין Document API, אם `capabilities` חסר, אם הקריאה
 *    זורקת או מחזירה משהו שאינו אובייקט — כל התשובות `false` עם סיבה, ולא
 *    „אולי כן”. פקד שמוצג פעיל ואינו עובד הוא בדיוק התקלה שהמודול הזה נכתב
 *    בגללה: תשע כפתורים בלי `@click` שנראו כמו כפתור עובד.
 *
 * מרחב השאלות מכסה גם את לשונית „הוספה” (תמונה, קישור, מעבר מקטע, תוכן
 * עניינים) ולא רק „פריסה”/„סקירה”: אותה בדיקה תידרש שם, ושני מודולים כאלה
 * היו נותנים שתי תשובות לאותה שאלה.
 */
import type { SuperDoc } from 'superdoc';
import type { MaybePromise } from './document-api';

/**
 * קודי הסיבה של המנוע (`CAPABILITY_REASON_CODES` ב-2.8.0), ועוד אחד שלנו:
 * `DOCUMENT_API_UNAVAILABLE` נוצר כאן כשאין בכלל מה לשאול, כדי שגם המצב הזה
 * יגיע לממשק כסיבה מפורשת ולא כ-`false` בלי הסבר.
 */
export const DOC_CAPABILITY_REASONS = [
  'COMMAND_UNAVAILABLE',
  'HELPER_UNAVAILABLE',
  'OPERATION_UNAVAILABLE',
  'TRACKED_MODE_UNAVAILABLE',
  'DRY_RUN_UNAVAILABLE',
  'NAMESPACE_UNAVAILABLE',
  'STYLES_PART_MISSING',
  'COLLABORATION_ACTIVE',
  'DOCUMENT_API_UNAVAILABLE',
] as const;

export type DocCapabilityReasonCode = (typeof DOC_CAPABILITY_REASONS)[number];

/** ההסבר שמגיע ל-tooltip. הנוסח של `NAMESPACE_UNAVAILABLE` הוא זה שהתכנית קובעת ב-§12. */
const REASON_TEXT: Record<DocCapabilityReasonCode, string> = {
  COMMAND_UNAVAILABLE: 'הפקודה אינה זמינה במנוע',
  HELPER_UNAVAILABLE: 'רכיב שהפעולה נשענת עליו אינו זמין במנוע',
  OPERATION_UNAVAILABLE: 'הפעולה אינה זמינה בגרסה הזאת של המנוע',
  TRACKED_MODE_UNAVAILABLE: 'הפעולה אינה נתמכת במצב מעקב אחר שינויים',
  DRY_RUN_UNAVAILABLE: 'הרצה יבשה אינה נתמכת בפעולה הזאת',
  NAMESPACE_UNAVAILABLE: 'אינו זמין בגרסה זו',
  STYLES_PART_MISSING: 'חסר חלק הסגנונות במסמך',
  COLLABORATION_ACTIVE: 'הפעולה אינה זמינה בזמן עבודה משותפת',
  DOCUMENT_API_UNAVAILABLE: 'המסמך עדיין נטען',
};

/**
 * מה כל שאלה בודקת. `operation` הוא מזהה הפעולה בקטלוג של המנוע, `global` הוא
 * דגל ברמת ה-namespace. שאלה עם שניהם דורשת את שניהם.
 */
interface CapabilitySpec {
  /** מזהה אחד, או רשימה שכולם נדרשים — יכולת שנשענת על שתי פעולות. */
  operation?: string | readonly string[];
  global?: 'trackChanges' | 'comments' | 'lists' | 'dryRun' | 'history';
}

const CAPABILITY_SPECS = {
  // פריסה
  canSetPageMargins: { operation: 'sections.setPageMargins' },
  canSetPageSetup: { operation: 'sections.setPageSetup' },
  canSetColumns: { operation: 'sections.setColumns' },
  canSetSectionDirection: { operation: 'sections.setSectionDirection' },
  canSetSectionBreak: { operation: 'sections.setBreakType' },
  // „גבולות עמוד” הוא כפתור תפריט אחד שגם מקיף וגם מסיר, ולכן הוא שואל על
  // שתי הפעולות. השאלה כבר הייתה כאן על `setPageBorders` לבדה, מהזמן שלא היה
  // לה פקד; מרגע שיש פקד, השאלה חייבת למנות את מה שהוא **באמת** מריץ —
  // אחרת „ללא גבול” נלחץ על מנוע שאין בו `clearPageBorders` ומחזיר כשל.
  canSetPageBorders: { operation: ['sections.setPageBorders', 'sections.clearPageBorders'] },
  // ארבע השאלות הבאות הן ארבעת הפקדים החדשים של „פריסה”. אף אחת מהן אינה
  // מונה את `sections.list`, בדיוק כמו ארבע השאלות שמעליהן: `list` ו-
  // `set*` הם אותו adapter אופציונלי, ומנוע שמצהיר על האחד ולא על השני אינו
  // צורה שקיימת — הצהרה חסרה מסמנת את **כל** ה-namespace יחד
  // (`NAMESPACE_UNAVAILABLE`). המודול עצמו נכשל סגור על היעדר `list` בלי
  // קשר לשאלה כאן.
  canSetLineNumbering: { operation: 'sections.setLineNumbering' },
  canSetVerticalAlign: { operation: 'sections.setVerticalAlign' },
  canSetPageNumbering: { operation: 'sections.setPageNumbering' },
  canSetHeaderFooterMargins: { operation: 'sections.setHeaderFooterMargins' },
  // גופן. `format.vertAlign` הוא alias ציבורי של `format.apply` על מפתח אחד
  // ב-`InlineRunPatch`, והוא `OperationId` בקטלוג — ולכן הוא נשאל כמו כל פעולה
  // אחרת. ראו engine/vert-align.ts: אין לו פקודה ב-registry של ה-controller,
  // והמסלול היחיד אליו הוא ה-Document API.
  canSetVertAlign: { operation: 'format.vertAlign' },
  // הפניות והוספה
  canInsertFootnote: { operation: 'footnotes.insert' },
  // „נהל הערות” — שאלה אחת לפקד אחד, כמו ב„סימנייה”: הדיאלוג מציג, עורך
  // ומסיר, ופקד מנוטרל למחצה אינו מצב שאפשר להציג. שתי הבחנות כאן, ושתיהן
  // נמדדו:
  //
  // 1. `footnotes.get` **הוא** חלק מהשאלה ואינו קישוט. כתובת ההערה אינה
  //    נושאת את סוגה (`entityType` הוא `'footnote'` גם עבור הערת סיום),
  //    ו-`get` הוא הדרך היחידה לשאול את המנוע לאיזו הערה הכתובת נפתרת לפני
  //    שנוגעים במסמך. בלעדיו „הסר” על הערת סיום היה מוחק הערת שוליים.
  // 2. `footnotes.configure` **אינו** נשאל: אין לו פקד. הוא כותב OOXML
  //    קנוני, אבל אין דרך לקרוא את ההגדרות שבמסמך וכל קריאה מחליפה את
  //    `w:footnotePr` כולו — כלומר טופס היה מוחק בשקט מה שהוגדר ב-Word.
  //    ההנמקה המלאה ב-engine/footnotes.ts.
  canManageNotes: {
    operation: ['footnotes.list', 'footnotes.get', 'footnotes.update', 'footnotes.remove'],
  },
  canSetPageBreakBefore: { operation: 'format.paragraph.setFlowOptions' },
  // גל 11 — תפריט „פסקה”. כל שאלה מונה את **כל** הפעולות שהפקד שלה מריץ,
  // בדיוק כמו `canSetPageBorders` למעלה: „נקה” בדיאלוג הוא כפתור לכל דבר,
  // ומנוע שיודע להגדיר ואינו יודע לנקות היה משאיר אותו פעיל על כשל מובטח.
  // `setKeepOptions` אין לו clear — `widowControl:false` הוא הניקוי שלו —
  // ולכן השאלה שלו על פעולה אחת.
  canSetParagraphIndentation: {
    operation: ['format.paragraph.setIndentation', 'format.paragraph.clearIndentation'],
  },
  canSetParagraphSpacing: {
    operation: ['format.paragraph.setSpacing', 'format.paragraph.clearSpacing'],
  },
  canSetParagraphKeepOptions: { operation: 'format.paragraph.setKeepOptions' },
  canManageParagraphTabs: {
    operation: [
      'format.paragraph.setTabStop',
      'format.paragraph.clearTabStop',
      'format.paragraph.clearAllTabStops',
    ],
  },
  // גל 12 — „גופן מתקדם". הדיאלוג כולו עובר בקריאה אחת ל-`format.apply`
  // (patch מרובה מפתחות = undo אחד), ולכן שאלה אחת לפעולה אחת. ה-alias
  // הבודדים (format.letterSpacing וכו') אינם נשאלים: איש מהם אינו נשלח.
  canApplyFontAdvanced: { operation: 'format.apply' },
  // גל 13 — „ברירות מחדל למסמך". styles.apply על docDefaults, ערוץ run.
  canSetDocDefaults: { operation: 'styles.apply' },
  // גל 14א — פעולות רשימה שאין להן פקודה ברצועה. המספור העברי
  // (setLevelNumberStyle + numberStyle:'hebrew1') הוא העיקרי.
  canManageLists: {
    operation: [
      'lists.setLevelNumberStyle',
      'lists.restartAt',
      'lists.continuePrevious',
      'lists.convertToText',
    ],
  },
  // `insert` ולא `create.text`: הכנסת טקסט היא פעולת הליבה של הפאסדה
  // (`memberPath: 'insert'` בקטלוג), אחותה של `delete`, ולא אחת מ-`create.*`.
  canInsertText: { operation: 'insert' },
  canInsertImage: { operation: 'create.image' },
  canInsertLink: { operation: 'hyperlinks.insert' },
  canInsertSectionBreak: { operation: 'create.sectionBreak' },
  canInsertTableOfContents: { operation: 'create.tableOfContents' },
  canInsertTable: { operation: 'create.table' },
  // כותרת עליונה ותחתונה. שאלה אחת לכל פקד, ולא אחת לכל פעולה: „עריכה”
  // ו„הסרה” יושבות באותו כפתור תפריט, ופקד מנוטרל למחצה אינו מצב שאפשר להציג.
  // `parts.create` הוא הפעולה שבלעדיה אין מה ליצור, ולכן היא זו שנשאלת.
  canEditHeaderFooter: { operation: 'headerFooters.parts.create' },
  canSetTitlePage: { operation: 'sections.setTitlePage' },
  canSetOddEvenHeaders: { operation: 'sections.setOddEvenHeadersFooters' },
  canLinkToPrevious: { operation: 'headerFooters.refs.setLinkedToPrevious' },
  // שדות. „מספר עמוד” ו„תאריך” דורשים את **שתי** הפעולות ולא רק את ההכנסה:
  // `fields.insert` מכניס שדה עם תוצאה ריקה, ו-`fields.rebuild` הוא שמחשב
  // אותה. מנוע שיודע להכניס ואינו יודע לחשב מחדש היה מכניס למסמך שדה בלתי
  // נראה, מדווח „בוצע”, והמשתמש לא היה רואה כלום — שדה שאי אפשר לראות אינו
  // פיצ'ר, וכפתור מנוטרל עם הסבר עדיף על הצלחה מדומה. (ההערה הקודמת כאן טענה
  // שההכנסה „אינה תלויה ב-rebuild”; היא כן — ראו engine/fields.ts.)
  //
  // שתי השאלות עדיין נפרדות, כי אינן זהות: מנוע עם `rebuild` בלי `insert`
  // משאיר את „עדכן שדות” פעיל על שדות שכבר במסמך, ומנטרל רק את ההכנסה.
  //
  // `fields.remove` אינו נשאל: אין לו פקד ברצועה — ב-Word מוחקים שדה כמו
  // שמוחקים טקסט — ושאלה בלי פקד היא הצהרת יכולת שאיש אינו קורא.
  canInsertField: { operation: ['fields.insert', 'fields.rebuild'] },
  canRebuildFields: { operation: 'fields.rebuild' },
  // סימניות. שאלה אחת לפקד אחד, כמו ב„כותרת עליונה”: „סימנייה” הוא כפתור
  // שפותח דיאלוג שמוסיף, מוחק ומשנה שם, ופקד מנוטרל למחצה אינו מצב שאפשר
  // להציג. `list` הוא חלק מהשאלה ולא קישוט — בלי רשימת השמות אין מה למחוק
  // ואין ממה לשנות שם, והדיאלוג היה נפתח ריק על מסמך מלא סימניות.
  canManageBookmarks: {
    operation: ['bookmarks.list', 'bookmarks.insert', 'bookmarks.rename', 'bookmarks.remove'],
  },
  // הפניות מקושרות. שתי הפעולות נדרשות מאותו טעם כמו ב-`canInsertField`:
  // `rebuild` מקבל כתובת של הפניה מסוימת, ואין דרך אחרת להשיג אותה מלבד
  // `list`. מנוע שיודע לחשב מחדש ואינו יודע למנות — הכפתור אצלו ילחץ ולא
  // יעדכן דבר, ויחזיר „בוצע”.
  //
  // `crossRefs.insert` **אינו** נשאל: הוא מוצהר זמין, מחזיר `success: true`,
  // וכותב שדה שגם Word וגם המנוע עצמו אינם יודעים לפתור. אין לו פקד, ושאלה
  // בלי פקד היא הצהרת יכולת שאיש אינו קורא. ההנמקה המלאה, כולל המדידה,
  // ב-engine/cross-refs.ts.
  canRebuildCrossRefs: { operation: ['crossRefs.list', 'crossRefs.rebuild'] },
  // תוכן עניינים. כל שאלה כאן מונה את הפעולות שהפקד שלה **באמת** מריץ, ולא
  // את ה-namespace כולו: מנוע שיודע לעדכן ואינו יודע להסיר צריך להשאיר את
  // „עדכן טבלה” פעיל ולנטרל רק את „הסר”.
  //
  // `toc.list` נדרש בשלוש מהן מפני ש-`update`/`remove`/`configure` מקבלים
  // כתובת של טבלה מסוימת, ואין דרך אחרת להשיג אותה.
  canUpdateTableOfContents: { operation: ['toc.list', 'toc.update'] },
  // `blocks.*` הם חלק מהשאלה ולא קישוט: `toc.remove` מוחק את הבלוק הראשון
  // של הטבלה בלבד ומשאיר את שאר השורות במסמך כפסקאות `TOC1`…`TOC9` (נמדד),
  // ו-`blocks.deleteRange` הוא מה שמנקה אותן. מנוע בלעדיו היה מציג „הסר”
  // שמחזיר „בוצע” ומשאיר את הטבלה על המסך. ההנמקה המלאה ב-engine/toc.ts.
  canRemoveTableOfContents: {
    operation: ['toc.list', 'toc.remove', 'blocks.list', 'blocks.deleteRange'],
  },
  canConfigureTableOfContents: { operation: ['toc.list', 'toc.configure'] },
  // שאלה אחת לפקד אחד, כמו ב„סימנייה”: „סמן ערך” הוא כפתור שפותח דיאלוג
  // שמסמן, מציג את הערכים הקיימים ומבטל סימון, ופקד מנוטרל למחצה אינו מצב
  // שאפשר להציג.
  canMarkTocEntry: {
    operation: ['toc.markEntry', 'toc.unmarkEntry', 'toc.listEntries'],
  },
  // מפתח ערכים. אותו עיקרון כמו בתוכן העניינים — כל שאלה מונה את הפעולות
  // שהפקד שלה **באמת** מריץ — אבל שתי הבחנות כאן שונות ממנו, ושתיהן נמדדו:
  //
  // 1. `index.rebuild` **אינו** חלק מ-`canInsertIndex`. שלא כמו `fields.insert`,
  //    שמכניס שדה עם תוצאה ריקה, `index.insert` מרנדר את המפתח מלא כבר
  //    ביצירה (נמדד: `getText` קיבל מיד את כל שמונת הערכים). מנוע שיודע
  //    להכניס ואינו יודע לבנות מחדש עדיין נותן למשתמש מפתח שרואים.
  // 2. `blocks.*` **אינם** חלק מ-`canRemoveIndex`, בניגוד ל„הסר תוכן
  //    עניינים”. המפתח הוא בלוק יחיד, ו-`index.remove` מוחק אותו כולו בלי
  //    להשאיר פסקאות יתומות (נמדד). ההנמקה המלאה ב-engine/index-field.ts.
  canInsertIndex: { operation: 'index.insert' },
  canRebuildIndex: { operation: ['index.list', 'index.rebuild'] },
  canRemoveIndex: { operation: ['index.list', 'index.remove'] },
  canConfigureIndex: { operation: ['index.list', 'index.configure'] },
  // שאלה אחת לפקד אחד, כמו ב„סימנייה”: „סמן ערך למפתח” הוא כפתור שפותח
  // דיאלוג שמסמן, מציג את הערכים הקיימים ומבטל סימון, ופקד מנוטרל למחצה
  // אינו מצב שאפשר להציג.
  canMarkIndexEntry: {
    operation: ['index.entries.list', 'index.entries.insert', 'index.entries.remove'],
  },
  // ציטוטים וביבליוגרפיה. שלוש הבחנות כאן, וכולן נמדדו:
  //
  // 1. `citations.list` הוא חלק מ-`canManageCitationSources` ואינו קישוט.
  //    „מחק מקור” מסרב כשיש למקור ציטוט במסמך, והרשימה הזאת היא הדרך
  //    היחידה לדעת. מנוע בלעדיה היה מוחק מקור ומשאיר שדה `CITATION` מצביע
  //    לתג שאינו קיים — נמדד שזה בדיוק מה שקורה. ההנמקה ב-engine/citations.ts.
  // 2. `fields.list` הוא חלק מהעדכון ומההסרה, וגם הוא אינו קישוט: אין
  //    ל-`citations.bibliography` פעולת `list`, ו-`blocks.list` מציג את
  //    הביבליוגרפיה כפסקה רגילה (נמדד). `fields.list` הוא המסלול היחיד
  //    שמחזיר `fieldType: 'BIBLIOGRAPHY'` ואת הכתובת שאפשר לפעול עליה.
  // 3. `citations.bibliography.configure` **אינו** נשאל: הוא כותב
  //    `BIBLIOGRAPHY \sdStyle "…"`, ו-`\sdStyle` אינו מתג של Word. אין לו
  //    פקד, ושאלה בלי פקד היא הצהרת יכולת שאיש אינו קורא.
  canManageCitationSources: {
    operation: [
      'citations.list',
      'citations.sources.list',
      'citations.sources.insert',
      'citations.sources.update',
      'citations.sources.remove',
    ],
  },
  canInsertCitation: { operation: ['citations.sources.list', 'citations.insert'] },
  // כיתובים. שאלה אחת לפקד אחד, כמו ב„סימנייה”: „הוסף כיתוב” הוא כפתור
  // שפותח דיאלוג שמוסיף, עורך ומסיר, ופקד מנוטרל למחצה אינו מצב שאפשר
  // להציג. שלוש הבחנות כאן, וכולן נמדדו:
  //
  // 1. `captions.update` **אינו** ברשימה, ולא מפני ששכחנו: נמדד שהוא מוסיף
  //    את הטקסט החדש על הישן במקום להחליף אותו („אלף” → „אלף: בית”), ולכן
  //    העריכה כאן היא `remove`+`insert`. שאלה על פעולה שהפקד אינו מריץ
  //    הייתה מנטרלת אותו בגלל משהו שאינו נוגע לו.
  // 2. `blocks.list` **הוא** חלק מהשאלה ואינו קישוט: העריכה מחזירה את
  //    הכיתוב למקומו לפי הבלוק שלפניו, ובלי הרשימה אין דרך לדעת מיהו.
  // 3. `captions.configure` אינו נשאל: הוא מחזיר `success: true` ואינו
  //    משנה דבר בקוד השדה (נמדד — `format: 'upperRoman'` נבלע). אין לו
  //    פקד, ושאלה בלי פקד היא הצהרת יכולת שאיש אינו קורא.
  canManageCaptions: {
    operation: ['captions.list', 'captions.insert', 'captions.remove', 'blocks.list'],
  },
  canInsertBibliography: { operation: 'citations.bibliography.insert' },
  canRebuildBibliography: { operation: ['fields.list', 'citations.bibliography.rebuild'] },
  canRemoveBibliography: { operation: ['fields.list', 'citations.bibliography.remove'] },
  // סקירה
  canAddComment: { operation: 'comments.create', global: 'comments' },
  canTrackChanges: { global: 'trackChanges' },
  // לוח ובחירה. `clipboard` הוא adapter אופציונלי בחוזה, בדיוק כמו `footnotes`,
  // ולכן אותה בדיקה נדרשת כאן. „גזור” הוא שתי יכולות (סדרוּר ומחיקה) ולא אחת,
  // ולכן הן שתי שאלות: מנוע שיודע להעתיק ואינו יודע למחוק צריך להשאיר את
  // „העתק” פעיל ולנטרל רק את „גזור”.
  canCopySelection: { operation: 'clipboard.serializeSelection' },
  canPasteContent: { operation: 'clipboard.insert' },
  canDeleteSelection: { operation: 'delete' },
  canResolveRange: { operation: 'ranges.resolve' },
} as const satisfies Record<string, CapabilitySpec>;

export type DocCapabilityQuestion = keyof typeof CAPABILITY_SPECS;

export const DOC_CAPABILITY_QUESTIONS = Object.keys(CAPABILITY_SPECS) as DocCapabilityQuestion[];

/** התשובה לכל השאלות, כפי שנקראה פעם אחת. */
export interface DocCapabilityReport {
  /** האם היה בכלל Document API לשאול. `false` = כל התשובות `false`. */
  readonly available: boolean;
  /** האם היכולת זמינה. */
  can(question: DocCapabilityQuestion): boolean;
  /** קודי הסיבה שהמנוע נתן. ריק כשהיכולת זמינה. */
  reasons(question: DocCapabilityQuestion): readonly DocCapabilityReasonCode[];
  /** הסבר בעברית, מוכן ל-tooltip. מחרוזת ריקה כשהיכולת זמינה. */
  explain(question: DocCapabilityQuestion): string;
}

/** הצורה הגולמית שנקראת מהמנוע. כל שדה אופציונלי: גרסה אחרת עשויה לא לחשוף אותו. */
interface RawFlag {
  enabled?: boolean;
  reasons?: readonly string[];
}

interface RawOperation {
  available?: boolean;
  reasons?: readonly string[];
}

interface RawCapabilities {
  global?: Partial<Record<string, RawFlag | undefined>>;
  operations?: Partial<Record<string, RawOperation | undefined>>;
}

/** הצורה שנצרכת מ-`doc`. ראו ההסבר ב-document-defaults.ts למה מוגדרת ולא מיובאת. */
export interface CapabilitiesDocumentApi {
  capabilities?: {
    get?: () => MaybePromise<RawCapabilities | undefined>;
  };
}

export interface CapabilitiesHost {
  activeEditor?: { doc?: CapabilitiesDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type CapabilitiesTarget = SuperDoc | CapabilitiesHost | null | undefined;

/** הסיבה כשהמנוע חסם ולא אמר למה. */
const FALLBACK_REASON: DocCapabilityReasonCode = 'OPERATION_UNAVAILABLE';

interface Answer {
  available: boolean;
  reasons: DocCapabilityReasonCode[];
}

/** קוד סיבה שהמנוע החזיר ואיננו מכירים אינו נזרק — הוא פשוט אינו מתורגם. */
function knownReasons(raw: readonly string[] | undefined): DocCapabilityReasonCode[] {
  if (!Array.isArray(raw)) return [];
  return raw.filter((code): code is DocCapabilityReasonCode =>
    (DOC_CAPABILITY_REASONS as readonly string[]).includes(code),
  );
}

/** התשובה כשאין מה לשאול. */
function closedAnswers(): Map<DocCapabilityQuestion, Answer> {
  const answers = new Map<DocCapabilityQuestion, Answer>();
  for (const question of DOC_CAPABILITY_QUESTIONS) {
    answers.set(question, { available: false, reasons: ['DOCUMENT_API_UNAVAILABLE'] });
  }
  return answers;
}

function buildReport(answers: Map<DocCapabilityQuestion, Answer>, available: boolean): DocCapabilityReport {
  /** שאלה שאינה במפה מקבלת את אותה תשובה סגורה — גם אם היא הגיעה מקוד לא מטופס. */
  const answerFor = (question: DocCapabilityQuestion): Answer =>
    answers.get(question) ?? { available: false, reasons: ['DOCUMENT_API_UNAVAILABLE'] };

  return {
    available,
    can: (question) => answerFor(question).available,
    reasons: (question) => answerFor(question).reasons,
    explain: (question) => {
      const answer = answerFor(question);
      if (answer.available) return '';
      // סיבות כפולות קורות כשהפעולה וה-namespace מדווחים את אותו קוד.
      const texts = [...new Set(answer.reasons.map((code) => REASON_TEXT[code]))];
      return texts.length > 0 ? texts.join('; ') : 'הפעולה אינה זמינה כרגע';
    },
  };
}

/**
 * קוראת את היכולות פעם אחת ומחזירה תצלום.
 *
 * לעולם אינה זורקת: `capabilities.get()` הוא קריאה אל המנוע, וכשל שלה אינו
 * סיבה להפיל את רינדור הרצועה. תצלום ולא reactive בכוונה — הצרכן קורא אותו
 * מחדש כשהמסמך מתחלף, ולא מחזיק מנוי על המנוע.
 */
export async function readDocCapabilities(
  host: CapabilitiesTarget,
): Promise<DocCapabilityReport> {
  const get = (host as CapabilitiesHost | null | undefined)?.activeEditor?.doc?.capabilities?.get;
  if (typeof get !== 'function') return buildReport(closedAnswers(), false);

  let raw: RawCapabilities | undefined;
  try {
    raw = await get();
  } catch (error) {
    console.warn('[otzaria-word] קריאת יכולות ה-Document API נכשלה', error);
    return buildReport(closedAnswers(), false);
  }

  // `null`, מחרוזת, מספר — כל תשובה שאינה אובייקט אינה תשובה.
  if (!raw || typeof raw !== 'object') return buildReport(closedAnswers(), false);

  const answers = new Map<DocCapabilityQuestion, Answer>();
  for (const question of DOC_CAPABILITY_QUESTIONS) {
    const spec: CapabilitySpec = CAPABILITY_SPECS[question];
    const reasons: DocCapabilityReasonCode[] = [];
    let available = true;

    const operations =
      typeof spec.operation === 'string' ? [spec.operation] : (spec.operation ?? []);
    for (const operation of operations) {
      const entry = raw.operations?.[operation];
      // פעולה שאינה בטבלה כלל = גרסה שאינה מכירה אותה. גם פעולה שכן בטבלה
      // אך לא נתנה סיבה מוכרת מקבלת את הקוד הגנרי, כדי שלא תישאר בלי הסבר.
      if (!entry || entry.available !== true) {
        available = false;
        const entryReasons = entry ? knownReasons(entry.reasons) : [];
        reasons.push(...(entryReasons.length > 0 ? entryReasons : [FALLBACK_REASON]));
      }
    }

    if (spec.global) {
      const flag = raw.global?.[spec.global];
      if (!flag || flag.enabled !== true) {
        available = false;
        const flagReasons = flag ? knownReasons(flag.reasons) : [];
        reasons.push(...(flagReasons.length > 0 ? flagReasons : [FALLBACK_REASON]));
      }
    }

    answers.set(question, { available, reasons });
  }

  return buildReport(answers, true);
}

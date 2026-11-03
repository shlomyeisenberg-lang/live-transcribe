/**
 * גלריית הסגנונות של המסמך.
 *
 * למה לא רשימה קשיחה בקומפוננטה, כפי שהיה: רשימה של שישה סגנונות בקוד אינה
 * יודעת מה יש **במסמך**. `ui.styles` חושף בדיוק את מה ש-Word מציג בלשונית
 * „בית” — `getQuickGallery()` הוא ה-Quick Styles של המסמך הפתוח, לפי
 * `w:qFormat`, בסדר של Word, עם השמות **כפי שהם במסמך**. במסמך עברי אלה השמות
 * העבריים; הרשימה הקשיחה תרגמה שישה מזהים אנגליים ולכל השאר לא הייתה תשובה,
 * וסגנון מותאם של המשתמש לא הופיע בגלריה בכלל.
 *
 * הקטלוג נפתר **אסינכרונית** אחרי פתיחת המסמך (`status: 'pending'` עד ההתייצבות
 * הראשונה), ולכן קריאה חד-פעמית ברגע הפתיחה מחזירה רשימה ריקה. זו הסיבה
 * ש-`observeStyleGallery` היא הדרך שהמעטפת אמורה להשתמש בה, ולא `readStyleGallery`
 * לבדה — בלי הרשמה הגלריה קופאת על מה שהיה זמין בשנייה הראשונה.
 *
 * נכשלת סגור: כל מצב שבו המנוע אינו יכול לענות — מצב תצוגה, עורך מבוסס worker,
 * עורך שעוד לא נטען, גרסה בלי `styles` — מחזיר את רשת הביטחון (`FALLBACK_STYLES`).
 * הרשימה הזאת נשארת בקוד **לא** כמקור אמת אלא כדי שגלריה ריקה לא תוצג למשתמש:
 * ששת המזהים שבה הם סגנונות בסיס של Word, והפקודה `linked-style` מקבלת אותם
 * גם כשהקטלוג לא נקרא.
 *
 * הטיפוסים נגזרים מ-`BorrowedSuperDocUI['styles']` ואינם מועתקים לכאן, מאותה
 * סיבה שב-engine/search.ts: כך החוזה הוא של המנוע, והוא נשבר ב-typecheck אם
 * המנוע ישתנה. `StyleGalleryHost` מרשה `Partial` בכוונה — superdoc עצמו קורא
 * למשטח הזה בהגנה (`this.ui?.styles?.getQuickGallery?.()`), כי הוא אינו קיים
 * בכל מצב עורך.
 *
 * גם גאומטריית הגלילה יושבת כאן ולא בקומפוננטה: היפוך הכיוון ב-RTL הוא
 * חשבון טהור שאי אפשר לראות בעין, ובקומפוננטה הוא היה נבדק רק בדפדפן.
 */
import type { BorrowedSuperDocUI } from 'superdoc';
import type { StyleCatalogItem } from 'superdoc/ui';

/** ה-handle של המנוע, כפי שהוא מוצהר על `SuperDocUI.styles`. */
type StylesHandle = BorrowedSuperDocUI['styles'];

/** ה-slice שהמנוע מחזיר. נגזר מה-handle כדי שלא תהיה כאן העתקה של החוזה. */
export type StylesSlice = ReturnType<StylesHandle['getSnapshot']>;

/** מה שנצרך מ-`superdoc.ui`. הכול אופציונלי — ראו הערת הראש. */
export interface StyleGalleryHost {
  styles?: Partial<StylesHandle> | null;
}

/* ------------------------------------------------------------------ */
/* הפריט שהגלריה מציגה                                                 */
/* ------------------------------------------------------------------ */

/**
 * תפקיד הסגנון בטיפוגרפיה של המסמך. הוא זה שקובע את **הצבע** של התצוגה
 * המקדימה, ולא הצבע שהמנוע מדווח: `preview.css` מביא את צבע הסגנון מהמסמך
 * (ב-Word ברירת המחדל של „כותרת 1” היא כחול כהה), והרצועה עשויה לשבת על משטח
 * כהה — צבע מסמך על משטח כהה הוא טקסט בלתי קריא. הצבע כאן הוא תפקיד מה-API
 * של אוצריא (`var(--color-*)`) ולכן מתהפך עם הערכת הנושא, והטיפוגרפיה — משקל,
 * נטייה, גופן — נלקחת מהמנוע.
 */
export type StyleRole = 'body' | 'heading' | 'subtle';

/** פריט אחד בגלריה, כבר מנורמל לתצוגה. */
export interface StyleGalleryItem {
  /** המזהה שנשלח לפקודה `linked-style`. */
  id: string;
  /** השם להצגה — מהמסמך, כשהוא זמין. */
  label: string;
  /** הטקסט בתוך הכרטיס. */
  previewText: string;
  /** טוקני CSS בטוחים לנושא. `style` בתגית מקבל אותם כמו שהם. */
  previewStyle: Readonly<Record<string, string>>;
  /** סגנון שהמשתמש יצר — רק אותו אפשר למחוק ולשנות את שמו, כמו ב-Word. */
  custom?: boolean;
}

/** מצב הגלריה כולה. */
export interface StyleGalleryState {
  items: readonly StyleGalleryItem[];
  /** הסגנון הפעיל בבחירה. `null` בבחירה מעורבת או כשאין תשובה. */
  activeId: string | null;
  /** `true` = הרשימה מהמסמך; `false` = רשת הביטחון. הקומפוננטה מכריעה לפיו למי להאמין לגבי `activeId`. */
  fromDocument: boolean;
}

/**
 * רשת הביטחון. שישה סגנונות בסיס של Word — מה שהוצג עד עכשיו תמיד, וממשיך
 * להוצג כשהקטלוג אינו זמין. מזהים בלבד: התווית העברית והתפקיד נגזרים מהמזהה
 * באותן פונקציות שמשרתות את הקטלוג האמיתי, ולכן אין כאן רשימת תוויות שנייה
 * שיכולה להיפרד ממנה.
 */
export const FALLBACK_STYLE_IDS: readonly string[] = [
  'Normal',
  'NoSpacing',
  'Heading1',
  'Heading2',
  'Subtitle',
  'Quote',
];

/** הטקסט בכרטיס של סגנון גוף. „AaBbCc” הוא מה ש-Word מציג, ולכן מזוהה. */
const BODY_PREVIEW_TEXT = 'AaBbCc';

/**
 * המזהים שמקבלים תפקיד לא-גוף. המזהים של Word אנגליים ויציבים גם במסמך עברי
 * (`w:styleId`), ולכן הזיהוי עליהם ולא על השם המוצג — שהוא מתורגם ומשתנה.
 */
const HEADING_IDS = /^(heading[1-9]?|title)$/;
const SUBTLE_IDS = /^(subtitle|quote|intensequote|caption|footnotetext|endnotetext)$/;

/**
 * המזהה בצורה שאפשר להשוות. הנרמול אינו קוסמטי: התבניות של Word אינן עקביות
 * באות ראשית ובריווח — בקטלוג של המסמך הריק נמדד `heading 1` באות קטנה ועם
 * רווח, בזמן שה-`w:styleId` הקנוני הוא `Heading1`. בלי הנרמול חצי מהטבלה
 * הייתה מפספסת.
 */
function styleKey(value: string): string {
  return value.trim().toLowerCase().replace(/[\s_-]+/g, '');
}

/** התפקיד של פריט קטלוג. */
export function styleRole(id: string): StyleRole {
  const key = styleKey(id);
  if (HEADING_IDS.test(key)) return 'heading';
  if (SUBTLE_IDS.test(key)) return 'subtle';
  return 'body';
}

/**
 * התוויות העבריות של הסגנונות המובנים של Word.
 *
 * למה זה נחוץ: `item.name` הוא השם כפי שהוא ב-`word/styles.xml`, והמסמך הריק
 * שהמנוע יוצר הוא תבנית **אנגלית** (`w:pgSz w:w="12240"` — Letter). כלומר
 * במסמך חדש הגלריה הציגה `Normal`, `heading 1`, `Title` — רגרסיה מול הרשימה
 * הקשיחה שקדמה לה, שלפחות הייתה בעברית.
 *
 * מה ש-Word עצמו עושה, וזה הכלל כאן: שם של סגנון **מובנה** מתורגם לשפת
 * הממשק, ושם של סגנון **מותאם אישית** מוצג כפי שהמשתמש נתן אותו. לכן ההתאמה
 * על המזהה הקנוני ולא על השם, ולכן `custom: true` אינו מתורגם בשום מצב.
 */
const BUILT_IN_LABELS: Readonly<Record<string, string>> = {
  normal: 'רגיל',
  nospacing: 'ללא מרווח',
  bodytext: 'גוף טקסט',
  title: 'כותרת',
  subtitle: 'כותרת משנה',
  heading1: 'כותרת 1',
  heading2: 'כותרת 2',
  heading3: 'כותרת 3',
  heading4: 'כותרת 4',
  heading5: 'כותרת 5',
  heading6: 'כותרת 6',
  heading7: 'כותרת 7',
  heading8: 'כותרת 8',
  heading9: 'כותרת 9',
  quote: 'ציטוט',
  intensequote: 'ציטוט מודגש',
  listparagraph: 'פסקת רשימה',
  caption: 'כתובית',
  strong: 'חזק',
  emphasis: 'הדגשה',
  subtleemphasis: 'הדגשה עדינה',
  intenseemphasis: 'הדגשה מודגשת',
  subtlereference: 'הפניה עדינה',
  intensereference: 'הפניה מודגשת',
  booktitle: 'שם ספר',
  header: 'כותרת עליונה',
  footer: 'כותרת תחתונה',
  footnotetext: 'טקסט הערת שוליים',
  endnotetext: 'טקסט הערת סיום',
  commenttext: 'טקסט הערה',
  hyperlink: 'היפר-קישור',
};

/** התווית העברית של סגנון מובנה, או `undefined` כשהוא אינו בטבלה. */
export function builtInStyleLabel(value: string | undefined): string | undefined {
  if (typeof value !== 'string' || value.trim() === '') return undefined;
  return BUILT_IN_LABELS[styleKey(value)];
}

/**
 * השם שהכרטיס מציג.
 *
 * שלושה מסלולי התאמה, בסדר יורד של קנוניות: המזהה, השם המוצג, ואז ה-aliases
 * (`w:aliases`). השם וה-aliases נבדקים גם הם מפני שיש תבניות שבהן ה-`styleId`
 * הוא מקומי ודווקא השם הוא האנגלי הקנוני; במסמך עברי הם פשוט לא מתאימים לאף
 * מפתח, וזה בסדר — השם שבמסמך הוא כבר התשובה הנכונה.
 */
export function styleDisplayLabel(item: StyleCatalogItem): string {
  const id = typeof item?.id === 'string' ? item.id.trim() : '';
  const name = typeof item?.name === 'string' ? item.name.trim() : '';

  // סגנון שהמשתמש יצר אינו מתורגם — גם אם קרא לו „Quote”.
  if (item?.custom !== true) {
    const translated =
      builtInStyleLabel(id) ??
      builtInStyleLabel(name) ??
      (item?.aliases ?? []).map((alias) => builtInStyleLabel(alias)).find(Boolean);
    if (translated) return translated;
  }

  return name !== '' ? name : id;
}

/**
 * טוקני ה-CSS של תפקיד. גדלים בפיקסלים ולא ביחסיים: הכרטיס בגלריה גבוה 68
 * פיקסלים בדיוק, וגודל יחסי לגופן המשתמש היה שובר את הפריסה של הרצועה.
 */
const ROLE_STYLE: Record<StyleRole, Readonly<Record<string, string>>> = {
  body: { fontSize: '13px', fontWeight: '400', color: 'var(--color-on-surface)' },
  heading: { fontSize: '14px', fontWeight: '700', color: 'var(--color-primary)' },
  subtle: { fontSize: '12px', fontStyle: 'italic', color: 'var(--color-on-surface-variant)' },
};

/**
 * מה שמותר לקחת מ-`preview.css`. רשימה סגורה ולא העברה גורפת: `preview.css`
 * הוא „small UI-safe CSS tokens” של המנוע, אבל UI-safe שם נמדד ביחס לגוף
 * המסמך — `color`, `background`, `margin` ו-`lineHeight` של סגנון מסמך היו
 * צובעים או מנפחים כרטיס ברצועה. הצבע מגיע מהתפקיד (ראו `StyleRole`).
 */
const SAFE_PREVIEW_KEYS = [
  'fontWeight',
  'fontStyle',
  'fontFamily',
  'fontVariant',
  'textDecoration',
  'textTransform',
  'letterSpacing',
] as const;

/** גבולות גודל התצוגה המקדימה, בפיקסלים. „כותרת 1” היא 16 נקודות במסמך. */
const PREVIEW_MIN_PX = 11;
const PREVIEW_MAX_PX = 16;

/** גודל בסיס להמרת יחידות יחסיות. אותו 13px של סגנון הגוף בכרטיס. */
const PREVIEW_BASE_PX = 13;

/**
 * גודל התצוגה המקדימה, מוקטן לגבולות הכרטיס. בלי ההקטנה כותרת של 32 נקודות
 * (Title) הייתה בולעת את הכרטיס ואת שאר הגלריה איתו.
 */
export function clampPreviewFontSize(raw: string | number | undefined): string | undefined {
  if (raw === undefined || raw === null) return undefined;
  const text = String(raw).trim();
  const match = /^(-?\d*\.?\d+)\s*(px|pt|em|rem|%)?$/.exec(text);
  if (!match) return undefined;

  const value = Number.parseFloat(match[1]);
  if (!Number.isFinite(value) || value <= 0) return undefined;

  const unit = match[2] ?? 'px';
  const px =
    unit === 'pt' ? (value * 4) / 3
    : unit === 'em' || unit === 'rem' ? value * PREVIEW_BASE_PX
    : unit === '%' ? (value / 100) * PREVIEW_BASE_PX
    : value;

  const clamped = Math.min(PREVIEW_MAX_PX, Math.max(PREVIEW_MIN_PX, Math.round(px)));
  return `${clamped}px`;
}

/**
 * טוקני התצוגה של פריט: התפקיד כבסיס, ומעליו הטיפוגרפיה של המנוע כשהיא זמינה
 * (`preview.available`). כך „כותרת 1” מוצגת במשקל ובגופן שבמסמך, ולא בניחוש.
 */
export function previewStyleFor(item: StyleCatalogItem): Readonly<Record<string, string>> {
  const style: Record<string, string> = { ...ROLE_STYLE[styleRole(item.id)] };
  const css = item.preview?.available ? item.preview.css : undefined;
  if (!css || typeof css !== 'object') return style;

  for (const key of SAFE_PREVIEW_KEYS) {
    const value = css[key];
    if (value === undefined || value === null) continue;
    const text = String(value).trim();
    if (text !== '') style[key] = text;
  }

  const size = clampPreviewFontSize(css.fontSize);
  if (size) style.fontSize = size;

  return style;
}

/** הפריטים כפי שהגלריה מציגה אותם. פריט בלי מזהה שמיש נשמט ולא מוצג ריק. */
export function toGalleryItems(
  items: readonly StyleCatalogItem[] | undefined,
): readonly StyleGalleryItem[] {
  const out: StyleGalleryItem[] = [];
  const seen = new Set<string>();

  for (const item of items ?? []) {
    const id = typeof item?.id === 'string' ? item.id.trim() : '';
    if (id === '' || seen.has(id)) continue;
    // `quickGallery` כבר מסונן, אבל `getCatalog` עם view אחר אינו — ופריט
    // מוסתר בגלריה של Word אינו מוצג.
    if (item.visibility?.effectivelyHidden) continue;
    seen.add(id);

    const label = styleDisplayLabel(item);
    const role = styleRole(id);
    out.push({
      id,
      label,
      // Word מציג את השם עצמו בכרטיס של כותרת, ו„AaBbCc” בכרטיס של סגנון גוף.
      previewText: role === 'body' ? BODY_PREVIEW_TEXT : label,
      previewStyle: previewStyleFor(item),
      custom: item.custom === true,
    });
  }

  return withHeadingLadder(out);
}

/* ------------------------------------------------------------------ */
/* סולם הכותרות                                                        */
/* ------------------------------------------------------------------ */

/**
 * עד איזו רמת כותרת הגלריה מציעה גם כשהמסמך אינו מציע אותה.
 *
 * שלוש, מפני שזה טווח ברירת המחדל של תוכן עניינים ב-Word (`TOC \o "1-3"`) —
 * וזה בדיוק מה שנשבר בלי זה. ראו `withHeadingLadder`.
 */
export const HEADING_LADDER_FLOOR = 3;

/** `Heading1`…`Heading9` בלבד. `Title` הוא תפקיד כותרת אבל אינו רמה. */
const HEADING_LEVEL_ID = /^heading([1-9])$/;

/** הרמה שהמזהה מייצג, או `null` כשאינו רמת כותרת. */
function headingLevelOf(id: string): number | null {
  const match = HEADING_LEVEL_ID.exec(styleKey(id));
  return match ? Number(match[1]) : null;
}

/**
 * משלימה את רמות הכותרת החסרות בגלריה.
 *
 * ## הבאג שזה מתקן
 *
 * דווח שתוכן העניינים נבנה מרמה אחת בלבד, גם אחרי „התאמה אישית” שמבקשת
 * רמות 1–3. נמדד על ה-dist הארוז, ומה שנמצא **אינו** בתוכן העניינים:
 *
 * 1. `toc.configure` דווקא כותב את המתג נכון — ה-`instruction` נעשה
 *    `TOC \o "1-3" \h`.
 * 2. וגם האיסוף עובד: פסקה שקיבלה `Heading2` נאספה אל הטבלה כשורת `TOC2`.
 * 3. אבל **הגלריה הציעה „כותרת 1” בלבד**: `quickGallery` של המנוע החזיר
 *    `["Normal","Heading1","Title","Subtitle","Quote","IntenseQuote","ListParagraph"]`.
 *    ב-`word/styles.xml` של התבנית `Heading2`…`Heading9` נושאים
 *    `<w:semiHidden/>` ו-`<w:unhideWhenUsed/>`, והמנוע מסנן אותם החוצה.
 *
 * כלומר לא הייתה שום דרך במסך להחיל „כותרת 2”, ולכן לא היו במסמך כותרות
 * ברמה 2 שהטבלה תאסוף. „רק כותרת אחת” היה תיאור מדויק של המסמך.
 *
 * ## למה כאן, ולא בתבנית „מסמך חדש”
 *
 * הפיתוי הוא לנקות את `w:semiHidden` מהתבנית (engine/blank-document.ts) ולתת
 * למנוע להציע אותם בעצמו. זה היה מתקן מסמך חדש ולא את השאר: נמדד ש-
 * `unhideWhenUsed` **אינו ממומש** — אחרי שהוחל `Heading2` על פסקה
 * (`success: true`, ה-pStyle נכתב) הגלריה של המנוע נשארה בת אותם שבעה
 * פריטים בדיוק. כלומר גם מסמך שהגיע מ-Word או מאוצריא **ומשתמש** ברמה 2
 * אינו מציע אותה. השכבה שלנו היא היחידה שרואה את כל המסמכים.
 *
 * ## מה נוסף, ומה לא
 *
 * הרמות עד `HEADING_LADDER_FLOOR`, ומעליהן רק מה שהמסמך עצמו כבר מציע (מסמך
 * Word שסגנונותיו העמוקים גלויים ממשיך להציג אותן). רמה עמוקה שהמסמך אינו
 * מציע נשארת בחוץ: הגלריה של Word אינה תשע כותרות, והפקודה
 * `linked-style` מקבלת כל מזהה — נמדד שהיא מחילה `Heading2` בהצלחה **גם
 * כשהסגנון semiHidden**, ולכן הכרטיס אינו מבטיח משהו שלא יעבוד.
 *
 * הכרטיס המסונתז נושא את התווית העברית מ-`BUILT_IN_LABELS` ואת טוקני התפקיד,
 * בלי טיפוגרפיה מהמסמך: `preview` מגיע מהקטלוג, וסגנון שאינו בקטלוג אין ממה
 * לגזור אותו. זו אותה הכרעה שרשת הביטחון כבר עושה.
 *
 * המקום: מיד אחרי הרמה שקדמה לה, כדי שהסדר יהיה 1,2,3 כמו ב-Word. כשהרמה
 * שקדמה לה אינה בגלריה — **לפני** הרמה העמוקה ביותר הבאה שכן בה, ורק בהיעדר
 * שתיהן אחרי הפריט הראשון, שהוא „רגיל”. הסעיף האמצעי אינו קוסמטי: גלריה
 * שמציעה רמה עמוקה בלבד הייתה מקבלת אותה **לפני** הרמות הרדודות שהושלמו,
 * כלומר „כותרת 9” ואחריה „כותרת 1” (נמדד בבדיקה).
 */
export function withHeadingLadder(
  items: readonly StyleGalleryItem[],
): readonly StyleGalleryItem[] {
  if (items.length === 0) return items;

  const levels = new Map<number, number>();
  items.forEach((item, index) => {
    const level = headingLevelOf(item.id);
    if (level !== null && !levels.has(level)) levels.set(level, index);
  });

  const out = [...items];
  for (let level = 1; level <= HEADING_LADDER_FLOOR; level++) {
    if (levels.has(level)) continue;

    const id = `Heading${level}`;
    const label = builtInStyleLabel(id) ?? id;
    const card: StyleGalleryItem = {
      id,
      label,
      previewText: label,
      previewStyle: { ...ROLE_STYLE.heading },
    };

    // `levels` מתעדכן כדי שהרמה הבאה תמצא את זו שהרגע נוספה, וההיסטים שאחרי
    // נקודת ההוספה נדחפים.
    const previous = levels.get(level - 1);
    let at: number;
    if (previous !== undefined) {
      at = previous + 1;
    } else {
      const deeper = [...levels.entries()]
        .filter(([existing]) => existing > level)
        .sort((a, b) => a[0] - b[0])[0];
      at = deeper === undefined ? 1 : deeper[1];
    }
    out.splice(at, 0, card);
    levels.set(level, at);
    for (const [existing, index] of levels) {
      if (existing !== level && index >= at) levels.set(existing, index + 1);
    }
  }

  return out;
}

/** רשת הביטחון, בצורת מצב גלריה. */
export function fallbackStyleGallery(): StyleGalleryState {
  return {
    // אותו סולם כותרות כמו בגלריה מהמסמך: רשת הביטחון נושאת „כותרת 1” ו-
    // „כותרת 2”, ובלי ההשלמה היא הייתה מציעה טווח רמות אחר מזה שהמסמך מציע.
    items: withHeadingLadder(
      FALLBACK_STYLE_IDS.map((id) => {
        const label = builtInStyleLabel(id) ?? id;
        const role = styleRole(id);
        return {
          id,
          label,
          previewText: role === 'body' ? BODY_PREVIEW_TEXT : label,
          previewStyle: { ...ROLE_STYLE[role] },
        };
      }),
    ),
    activeId: null,
    fromDocument: false,
  };
}

/**
 * המצב מתוך slice של המנוע. `status: 'pending'` עם גלריה ריקה הוא המסלול
 * האמיתי בשנייה שאחרי פתיחת המסמך, ולכן הוא נופל לרשת הביטחון ולא לרשימה
 * ריקה — הגלריה נראית עובדת עד שהקטלוג מתייצב, ואז מתחלפת.
 */
export function toGalleryState(slice: StylesSlice | null | undefined): StyleGalleryState {
  const items = toGalleryItems(slice?.quickGallery);
  if (items.length === 0) return fallbackStyleGallery();

  return {
    items,
    // בחירה שפורסת כמה סגנונות אינה מסמנת כרטיס: Word מציג גלריה בלי בחירה,
    // ולא את הסגנון של הפסקה הראשונה כאילו הוא של כולן.
    activeId: slice?.mixedSelection ? null : (slice?.activeParagraphStyleId ?? null),
    fromDocument: true,
  };
}

/** קריאה בהגנה: מתודה חסרה או זורקת אינה סיבה להפיל את הרצועה. */
function safeRead<T>(read: (() => T) | undefined): T | undefined {
  if (typeof read !== 'function') return undefined;
  try {
    return read();
  } catch (error) {
    console.warn('[otzaria-word] קריאת גלריית הסגנונות מהמנוע נכשלה', error);
    return undefined;
  }
}

/**
 * המצב ברגע זה. `getSnapshot` הוא המסלול הראשי — הוא מביא את הגלריה ואת
 * הסגנון הפעיל בקריאה אחת. `getQuickGallery`/`getActiveParagraphStyle` הם
 * מסלול גיבוי לגרסה שחושפת רק אותם.
 */
export function readStyleGallery(ui: StyleGalleryHost | null | undefined): StyleGalleryState {
  const handle = ui?.styles;
  if (!handle) return fallbackStyleGallery();

  const slice = safeRead(handle.getSnapshot?.bind(handle));
  if (slice) {
    const state = toGalleryState(slice);
    if (state.fromDocument) return state;
  }

  const items = toGalleryItems(safeRead(handle.getQuickGallery?.bind(handle)));
  if (items.length === 0) return fallbackStyleGallery();

  const active = safeRead(handle.getActiveParagraphStyle?.bind(handle));
  return {
    items,
    activeId: active?.mixed ? null : (active?.styleId ?? null),
    fromDocument: true,
  };
}

/**
 * מאזינה לגלריה. `observe` של המנוע יורה מיד עם ה-snapshot הנוכחי ואז על כל
 * שינוי, ולכן אין צורך בקריאה נפרדת לפניה — וזה מה שמכסה את המסלול האמיתי:
 * הקטלוג נפתר אחרי הפתיחה, וההרשמה היא מה שמביא אותו למסך.
 *
 * מחזירה disposer גם כשאין `observe`, כדי שאתר הקריאה לא יצטרך להבחין.
 */
export function observeStyleGallery(
  ui: StyleGalleryHost | null | undefined,
  listener: (state: StyleGalleryState) => void,
): () => void {
  const handle = ui?.styles;
  const observe = handle?.observe;

  if (typeof observe !== 'function') {
    listener(readStyleGallery(ui));
    return () => {};
  }

  try {
    const off = observe.call(handle, (slice) => listener(toGalleryState(slice)));
    return typeof off === 'function' ? off : () => {};
  } catch (error) {
    console.warn('[otzaria-word] האזנה לגלריית הסגנונות נכשלה', error);
    listener(readStyleGallery(ui));
    return () => {};
  }
}

/* ------------------------------------------------------------------ */
/* גאומטריית הגלילה                                                    */
/* ------------------------------------------------------------------ */

/**
 * צעד גלילה אחד — שני כרטיסים שלמים (`--style-card-width` + `--style-card-gap`
 * ב-StyleGallery.vue, 68+3). כפולה שלמה של רוחב כרטיס ולא מספר עגול, כדי
 * שהגלילה תיעצר על גבול כרטיס ולא באמצעו; `scroll-snap` שם מאכף את זה גם על
 * גלילה בגלגלת.
 */
export const GALLERY_SCROLL_STEP_PX = 142;

/** סבילות במידות תת-פיקסליות. בלעדיה כפתור נראה זמין על 0.4 פיקסל גלילה. */
const EDGE_TOLERANCE_PX = 1;

/** לאן לגלול — ביחס לרשימה ולא למסך. `end` = לעבר הסגנונות הבאים. */
export type ScrollToward = 'start' | 'end';

/**
 * ה-delta ל-`scrollBy({ left })`.
 *
 * ההיפוך כאן הוא הבאג שהיה: הקוד גלל `left: +120` וקרא לזה „ימינה”, בזמן
 * שבמכולה RTL `scrollLeft` מתחיל ב-0 בקצה הימני ויורד לשלילי כשגוללים שמאלה.
 * כלומר `+120` ב-RTL מחזיר **ימינה**, לתחילת הרשימה — הפוך מהתווית. delta
 * שלילי הוא תמיד „שמאלה על המסך” בשני מודלי ה-scrollLeft שדפדפנים חיים
 * מיישמים, ולכן החשבון כאן נשען עליו ולא על סימן ה-scrollLeft.
 */
export function galleryScrollDelta(
  toward: ScrollToward,
  rtl: boolean,
  step: number = GALLERY_SCROLL_STEP_PX,
): number {
  const forward = toward === 'end' ? 1 : -1;
  return forward * (rtl ? -1 : 1) * step;
}

/** המידות שהמכולה מדווחת. */
export interface GalleryScrollMetrics {
  scrollLeft: number;
  scrollWidth: number;
  clientWidth: number;
}

/** אילו כפתורי גלילה יש להציג. שניהם `false` = הגלריה נכנסת כולה, ואין כפתורים. */
export interface GalleryScrollAvailability {
  canScrollStart: boolean;
  canScrollEnd: boolean;
}

/**
 * האם יש לאן לגלול לכל כיוון. כפתור שאין לאן לגלול איתו מוסתר, כמו ב-Word,
 * ולא מוצג כפעיל ולא עושה כלום.
 *
 * המרחק מתחילת הרשימה מחושב כערך מוחלט של `scrollLeft` דווקא: ב-RTL הטווח הוא
 * `[-(scrollWidth - clientWidth), 0]` וב-LTR הוא `[0, scrollWidth - clientWidth]`,
 * ולכן אותו חשבון מכסה את שניהם בלי ענף.
 */
export function galleryScrollAvailability(
  metrics: GalleryScrollMetrics,
): GalleryScrollAvailability {
  const overflow = metrics.scrollWidth - metrics.clientWidth;
  if (!Number.isFinite(overflow) || overflow <= EDGE_TOLERANCE_PX) {
    return { canScrollStart: false, canScrollEnd: false };
  }

  const fromStart = Math.abs(metrics.scrollLeft);
  return {
    canScrollStart: fromStart > EDGE_TOLERANCE_PX,
    canScrollEnd: fromStart < overflow - EDGE_TOLERANCE_PX,
  };
}

/**
 * האייקון של כפתור הגלילה. ב-RTL „הסגנונות הבאים” נמצאים שמאלה, ולכן החץ
 * שמאלה — וזו בדיוק הנקודה שבה התוויות והחצים היו הפוכים.
 */
export function galleryScrollIcon(toward: ScrollToward, rtl: boolean): 'chevronLeft' | 'chevronRight' {
  const towardVisualLeft = rtl ? toward === 'end' : toward === 'start';
  return towardVisualLeft ? 'chevronLeft' : 'chevronRight';
}

/**
 * „גופן מתקדם" — ריווח תווים, מיקום, קו חוצה כפול, אפקטים, טקסט מוסתר
 * והליבה העברית (CS), דרך `doc.format.apply` — patch מרובה מפתחות בקריאה אחת.
 *
 * ## מה שנמדד לפני שנכתבה כאן השורה הראשונה
 *
 * Chrome headless על ה-dist הארוז; כל סבב מלווה בפירוק ה-zip של
 * `export.toDocx`. שתי סדרות: alias בודד (`format.<key>`) ו-`format.apply`
 * מרובה-מפתחות. הממצאים:
 *
 * ### היחידות — ה-API בנקודות, ה-XML ביחידות OOXML
 *
 *     letterSpacing: 2  → <w:spacing w:val="40"/>   (×20, twips)
 *     position: 3       → <w:position w:val="6"/>   (×2, חצאי נקודות)
 *     kerning: 12       → <w:kern w:val="24"/>      (×2)
 *     fontSizeCs: 12.5  → <w:szCs w:val="25"/>      (×2; חצאי נקודות מותרות,
 *                                                    כמו fontSize עצמו)
 *     charScale: 125    → <w:w w:val="125"/>        (אחוזים, כמות שהוא)
 *
 * ### מה שעובר בשקט — ולכן הוולידציה יושבת כאן
 *
 *     charScale: 9999 → success:true ו-<w:w w:val="9999"/>.
 *       Word תחום 1..600 אחוז; מעבר לזה הטקסט יוצג מטורף או לא ייפתח.
 *     kerning: -5 → success:true ו-<w:kern w:val="-10"/>.
 *       ST_HpsMeasure אינו חתום; ערך שלילי פסול.
 *
 * לכן: charScale שלם 1..600, kerning שלם ≥ 0. letterSpacing ו-position
 * דווקא חתומים וחוקיים בשלילי („מכווץ"/„מונמך") — נמדד ונשלח.
 *
 * ### הליבה העברית, והממצא שקובע
 *
 * על טקסט עברי:
 *     rtl/cs/bCs/iCs → <w:rtl/> <w:cs/> <w:bCs/> <w:iCs/>
 *     fontSizeCs: 24 → <w:szCs w:val="48"/>
 *     lang {bidi:'he-IL'} → <w:lang w:bidi="he-IL"/>
 *     rFonts {cs:'David'} → <w:rFonts w:cs="David"/>
 *
 * **אבל `bold: true` כתב רק `<w:b/>`, בלי `<w:bCs/>`.** ב-Word, הדגשת טקסט
 * מורכב (עברית) נקראת מ-`bCs`; ריצה עברית שנושאת `b` בלבד אינה תוצג
 * מודגשת. זה פער במנוע (docs/engine-gaps.md), וזו הסיבה שהדיאלוג מציע
 * „מודגש (מורכב)" דרך `bCs` — הדרך היחידה להבטיח הדגשה שתיראה.
 *
 * ### NO_OP, ואיפה הוא מופיע
 *
 * ה-alias הבודדים החזירו `success:true` גם על חזרה זהה; `format.apply`
 * מחזיר `NO_OP` ("produced no change") כשה-patch לא משנה דבר — למשל
 * `vanish: null` על טקסט שאינו מוסתר. NO_OP הוא הצלחה, כתמיד.
 *
 * ### מה לא נשלח, ולמה
 *
 * - `webHidden` — אין לו משמעות ממשק מחוץ לתשתית הסתרה של Word; אין פקד.
 * - `rStyle` — נוגע בסגנונות תו; גל 13 (סגנונות) הוא בעליו, ושני מסלולים
 *   לאותה כתיבה הם באג. ייבחן שם.
 * - `smallCaps`/`caps` וכל משפחת CJK/East-Asian — אסורים מראש: אין להם
 *   משמעות בעברית (הוראת הגל).
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt } from './document-api';

/** הנוסח שהתכנית קובעת ב-§12 לפקד שאין לו API זמין. */
const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

const LOADING_TEXT = 'המסמך עדיין נטען';

/** גבולות שנקבעו ממדידה + מהתקן. ראו הערת הפתיחה. */
export const CHAR_SCALE_MIN = 1;
export const CHAR_SCALE_MAX = 600;
export const SPACING_PT_LIMIT = 1000;
export const KERNING_PT_MAX = 1000;

/**
 * ה-patch שהדיאלוג בונה. כל שדה אופציונלי — שדה שאינו נוכח אינו נשלח כלל
 * („ללא שינוי"). היחידות: נקודות למספרים; ההמרה ל-OOXML נעשית במנוע
 * (×20 ל-spacing, ×2 ל-position/kerning/szCs — נמדד).
 */
export interface FontAdvancedPatch {
  /** מתיחה אופקית באחוזים, 1..600. */
  charScale?: number;
  /** ריווח תווים בנקודות. שלילי = מכווץ (חוקי — נמדד). */
  letterSpacingPt?: number;
  /** מיקום התו בנקודות. חיובי = מוגבה, שלילי = מונמך (נמדד). */
  positionPt?: number;
  /** רף קרנינג בנקודות, ≥ 0. */
  kerningPt?: number;
  /** קו חוצה כפול. */
  dstrike?: boolean;
  outline?: boolean;
  shadow?: boolean;
  emboss?: boolean;
  imprint?: boolean;
  /**
   * טקסט מוסתר. **מסתיר תוכן מעיני המשתמש** — הפקד מציג אזהרה, וההסרה
   * (`false`) כותבת `w:vanish w:val="0"` בדיוק כמו Word (נמדד).
   */
  vanish?: boolean;
  /** גודל גופן מורכב בנקודות; חצאי נקודות מותרות (נמדד: 12.5 → szCs 25). */
  fontSizeCsPt?: number;
  /** מודגש לטקסט מורכב — מה ש-Word מציג על עברית (ראו הממצא על `bold`). נשלח למנוע כ-`bCs`. */
  boldCs?: boolean;
  /** נשלח למנוע כ-`iCs`. */
  italicCs?: boolean;
  /** „השתמש בגופן מורכב". נשלח למנוע כ-`cs`. */
  complexScript?: boolean;
  /** כיוון הריצה מימין לשמאל. */
  rtl?: boolean;
  /** גופן ה-CS, למשל „David". */
  complexFontName?: string;
  /** שפת הגהה לטקסט מורכב, למשל he-IL. מחרוזת לא-ריקה או היעדרות. */
  proofingLangBidi?: string;
}

export interface FontAdvancedDocumentApi {
  selection?: {
    current?: () => MaybePromise<SelectionInfoLike | undefined>;
  };
  format?: {
    apply?: (input: { target: unknown; inline: Record<string, unknown> }) => MaybePromise<DocReceipt>;
  };
}

export interface FontAdvancedHost {
  activeEditor?: { doc?: FontAdvancedDocumentApi | null } | null;
}

export type FontAdvancedTarget = SuperDoc | FontAdvancedHost | null | undefined;

interface SelectionInfoLike {
  empty?: boolean;
  selectionTarget?: unknown;
}

type MaybePromise<T> = T | Promise<T>;

function docOf(host: FontAdvancedTarget): FontAdvancedDocumentApi | null {
  return (host as FontAdvancedHost | null | undefined)?.activeEditor?.doc ?? null;
}

/**
 * שולפת את ה-`selectionTarget` של הבחירה החיה — אותו מסלול שנמדד ב-
 * vert-align.ts. `format.apply` דורש SelectionTarget ואינו מקבל TextTarget
 * (נמדד: „target must be a SelectionTarget object"), וטווח ריק אינו יעד.
 */
async function readSelectionTarget(
  doc: FontAdvancedDocumentApi,
): Promise<{ target: unknown } | { failure: CommandOutcome }> {
  const current = doc.selection?.current;
  if (typeof current !== 'function') {
    return { failure: { ok: false, message: LOADING_TEXT, reason: 'not-ready' } };
  }

  let info: SelectionInfoLike | undefined;
  try {
    info = await current();
  } catch (error) {
    console.warn('[otzaria-word] קריאת הבחירה לגופן מתקדם נכשלה', error);
    info = undefined;
  }

  if (!info || info.empty === true || !info.selectionTarget) {
    return {
      failure: {
        ok: false,
        message: 'יש לסמן טקסט תחילה',
        reason: 'range-selection-required',
      },
    };
  }

  return { target: info.selectionTarget };
}

/**
 * האם יש **טווח** מסומן — השאלה שהדיאלוג שואל בפתיחה.
 *
 * `applyFontAdvanced` שואל אותה ממילא, אבל רק ברגע ה„אישור” — כלומר אחרי
 * שהמשתמש כבר מילא שבעה-עשר פקדים. הוא היה מקבל אז „יש לסמן טקסט תחילה”
 * בשורת המצב, על דיאלוג שכבר נסגר ולקח איתו את כל מה שהוקלד. הפקד הזה מזיז
 * את אותה תשובה לפתיחה, וזו אותה קריאה בדיוק — `readSelectionTarget` —
 * ולא שאלה שנייה שעלולה לענות אחרת.
 *
 * ארבע תשובות ולא שתיים, מפני שהן מובילות לשלוש הודעות שונות:
 *
 * - `'range'` — יש טווח. הדיאלוג פתוח לעסקים.
 * - `'none'` — הסמן מכווץ. זה מה שהמשתמש יכול לתקן בעצמו, ולכן זו ההודעה
 *   היחידה שאומרת לו מה לעשות.
 * - `'unavailable'` — יש מסמך, ואין לו `format.apply`. הדיאלוג היה נפתח,
 *   מתמלא, ונכשל ב„אינו זמין בגרסה זו”; ההצהרה ב-HomeTab („הפתיחה מסבירה”)
 *   מחייבת שזה ייאמר בפתיחה ולא אחריה.
 * - `'unknown'` — אין לדעת: אין מסמך פעיל (טעינה, או מעבר בין מסמכים), אין
 *   `selection.current`, או שהקריאה לא ענתה. **אינו נועל דבר.** השער האמיתי
 *   נשאר ב-`applyFontAdvanced` ממילא, וזו הקדמה של ההודעה בלבד — לא שער שני.
 *
 * קריאה שזרקה נספרת כ-`'none'` ולא כ-`'unknown'`, וזה מכוון: `applyFontAdvanced`
 * עובר דרך אותה `readSelectionTarget` ויענה בדיוק אותו דבר, ולכן „פתוח לעסקים”
 * היה הבטחה שתיכשל מיד. והדיאלוג שואל שוב בכל שינוי בחירה, כך שתשובה חולפת
 * מתקנת את עצמה.
 */
export type SelectionReadiness = 'range' | 'none' | 'unavailable' | 'unknown';

export async function hasRangeSelection(host: FontAdvancedTarget): Promise<SelectionReadiness> {
  const doc = docOf(host);
  /*
   * „אין מסמך” אינו „אין יכולת”. `docOf` מחזיר `null` גם בין מסמכים וגם בזמן
   * טעינה (App.vue מאפס את המסמך הפעיל ואז פותח), ודיאלוג שהיה אומר שם „אינו
   * זמין בגרסה זו של המנוע” היה מוסר טענה שקרית על הבניין. `unavailable` שמור
   * למשטח שיש לו מסמך ואין לו `format.apply` — הפער שהוא באמת מתאר.
   */
  if (!doc) return 'unknown';
  if (typeof doc.format?.apply !== 'function') return 'unavailable';

  const outcome = await readSelectionTarget(doc);
  if (!('failure' in outcome)) return 'range';
  const { failure } = outcome;
  return !failure.ok && failure.reason === 'range-selection-required' ? 'none' : 'unknown';
}

/** מאמתת מספר שלם בתחום. `null` = פסול. */
function intIn(value: number, min: number, max: number): number | null {
  return typeof value === 'number' && Number.isInteger(value) && value >= min && value <= max
    ? value
    : null;
}

/**
 * בונה את ה-patch למנוע מתוך הקלט של הדיאלוג, עם השערים שהמנוע אינו מפעיל.
 * כל ערך פסול הוא בדיוק אלה שהמנוע היה בולע ב-`success:true` (נמדד).
 */
export function buildInlinePatch(patch: FontAdvancedPatch): { inline: Record<string, unknown> } | { error: CommandOutcome } {
  const inline: Record<string, unknown> = {};

  if (patch.charScale !== undefined) {
    const value = intIn(patch.charScale, CHAR_SCALE_MIN, CHAR_SCALE_MAX);
    if (value === null) {
      return { error: { ok: false, message: `מתיחה אופקית חייבת להיות מספר שלם בין ${CHAR_SCALE_MIN} ל-${CHAR_SCALE_MAX} אחוזים`, reason: 'invalid-char-scale' } };
    }
    inline.charScale = value;
  }

  if (patch.letterSpacingPt !== undefined) {
    const value = intIn(patch.letterSpacingPt, -SPACING_PT_LIMIT, SPACING_PT_LIMIT);
    if (value === null) {
      return { error: { ok: false, message: `ריווח תווים חייב להיות מספר שלם בין -${SPACING_PT_LIMIT} ל-${SPACING_PT_LIMIT} נקודות`, reason: 'invalid-letter-spacing' } };
    }
    inline.letterSpacing = value;
  }

  if (patch.positionPt !== undefined) {
    const value = intIn(patch.positionPt, -SPACING_PT_LIMIT, SPACING_PT_LIMIT);
    if (value === null) {
      return { error: { ok: false, message: `מיקום התו חייב להיות מספר שלם בין -${SPACING_PT_LIMIT} ל-${SPACING_PT_LIMIT} נקודות`, reason: 'invalid-position' } };
    }
    inline.position = value;
  }

  if (patch.kerningPt !== undefined) {
    // המנוע קיבל `-5` בחיוב וכתב `w:kern="-10"` — ST_HpsMeasure אינו חתום.
    const value = intIn(patch.kerningPt, 0, KERNING_PT_MAX);
    if (value === null) {
      return { error: { ok: false, message: `הקרנינג חייב להיות מספר שלם בין 0 ל-${KERNING_PT_MAX} נקודות`, reason: 'invalid-kerning' } };
    }
    inline.kerning = value;
  }

  if (patch.fontSizeCsPt !== undefined) {
    // חצאי נקודות מותרות (12.5 → szCs 25); שאר השברים אינם מובנים.
    const halfPoints = patch.fontSizeCsPt * 2;
    if (!Number.isFinite(halfPoints) || !Number.isInteger(halfPoints) || halfPoints <= 0 || halfPoints > 1600) {
      return { error: { ok: false, message: 'גודל הגופן המורכב חייב להיות מספר חיובי עד 800 נקודות', reason: 'invalid-font-size-cs' } };
    }
    inline.fontSizeCs = patch.fontSizeCsPt;
  }

  // שמות השדה בפאץ' של הדיאלוג אינם שמות המפתחות של המנוע: המנוע מכיר
  // `bCs`/`iCs`/`cs`, לא `boldCs`/`italicCs`/`complexScript` (ראו הערת הפתיחה
  // של המודול — הרשימה נמדדה). מיפוי חסר כאן הוא בדיוק מה שהחזיר
  // `INVALID_INPUT: Unknown inline property` והפיל את כל ה-patch באותה קריאה.
  const ENGINE_KEY: Partial<Record<'boldCs' | 'italicCs' | 'complexScript', string>> = {
    boldCs: 'bCs',
    italicCs: 'iCs',
    complexScript: 'cs',
  };

  for (const key of ['dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish', 'boldCs', 'italicCs', 'complexScript', 'rtl'] as const) {
    const value = patch[key];
    if (value !== undefined) {
      if (typeof value !== 'boolean') {
        return { error: { ok: false, message: 'מאפייני האפקטים מקבלים כן/לא בלבד', reason: 'invalid-flag' } };
      }
      inline[ENGINE_KEY[key as keyof typeof ENGINE_KEY] ?? key] = value;
    }
  }

  if (patch.complexFontName !== undefined) {
    const name = typeof patch.complexFontName === 'string' ? patch.complexFontName.trim() : '';
    if (name === '' || name.length > 100) {
      return { error: { ok: false, message: 'שם הגופן המורכב נדרש וקצר מ-100 תווים', reason: 'invalid-complex-font' } };
    }
    inline.rFonts = { cs: name };
  }

  if (patch.proofingLangBidi !== undefined) {
    const lang = typeof patch.proofingLangBidi === 'string' ? patch.proofingLangBidi.trim() : '';
    if (!/^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$/.test(lang)) {
      return { error: { ok: false, message: 'קוד שפה אינו תקין (למשל he-IL)', reason: 'invalid-lang' } };
    }
    inline.lang = { bidi: lang };
  }

  return { inline };
}

/**
 * מחילת עיצוב גופן מתקדם על **הטקסט המסומן**, בקריאה אחת ל-`format.apply`.
 *
 * למה apply ולא עשרים alias בודדים: קריאה אחת = מוטציה אחת = undo אחד;
 * alias לכל מפתח היה יוצר עשרים צעדי היסטוריה לפעולה אחת של המשתמש.
 *
 * לעולם אינה זורקת. NO_OP הוא הצלחה.
 */
export async function applyFontAdvanced(
  host: FontAdvancedTarget,
  patch: FontAdvancedPatch,
): Promise<CommandOutcome> {
  const failedAction = 'החלת עיצוב הגופן נכשלה';

  const built = buildInlinePatch(patch);
  if ('error' in built) return built.error;
  if (Object.keys(built.inline).length === 0) {
    // דיאלוג שנפתח ואושר בלי שינוי אינו כשל — אין מה לעשות, וזו הצלחה.
    return { ok: true };
  }

  const doc = docOf(host);
  const apply = doc?.format?.apply;
  if (!doc || typeof apply !== 'function') {
    return { ok: false, message: `${failedAction}: ${UNAVAILABLE_TEXT}`, reason: 'command-unsupported' };
  }

  const selection = await readSelectionTarget(doc);
  if ('failure' in selection) return selection.failure;

  let receipt: DocReceipt;
  try {
    receipt = await apply({ target: selection.target, inline: built.inline });
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  if (receipt?.success === false && receipt.failure?.code !== 'NO_OP') {
    return { ok: false, message: receiptFailureText(failedAction, receipt), reason: receipt.failure?.code };
  }

  return { ok: true };
}




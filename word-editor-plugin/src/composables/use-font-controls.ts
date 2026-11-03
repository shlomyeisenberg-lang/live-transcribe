/**
 * בורר הגופן ובורר הגודל — מצב אחד לשני המקומות שמציגים אותם.
 *
 * ## למה זה יצא מ-`HomeTab.vue`
 *
 * הפקדים האלה מופיעים עכשיו גם בתפריט הלחצן הימני, ופקד שמחזיק מצב משלו
 * במקום השני הוא פקד שמשקר באחד מהם. המצב כאן אינו „הערך שהמנוע מדווח” — זה
 * ממילא משותף דרך ה-`CommandAdapter` — אלא **שתי שכבות הזיכרון** שסביבו:
 *
 * 1. „האחרון שידענו”, שנדרש מפני שהמנוע מדווח `undefined` גם על בחירה מעורבת
 *    וגם על רגע שבו הוא עוד לא פתר את הבחירה. בורר שמתרוקן בכל תנועת סמן הוא
 *    גרוע מבורר שמציג את הערך האחרון שכן ידענו.
 * 2. „מה שנבחר וטרם נענה”, שנדרש כדי שבחירה שהמסמך דחה תיעלם מהמסך.
 *
 * שתיהן היו `ref` מקומי בלשונית „בית”. תפריט הקשר שנפתח היה מקבל עותק טרי
 * שלהן, ולכן בדיוק ברגע שהמנוע אינו מדווח ערך — כלומר מיד אחרי שהתפריט הזיז
 * את הסמן — הוא היה מציג „Assistant 12” בזמן שהרצועה מציגה את גופן המסמך.
 * זה גם ההפך הגמור ממה שהמשתמש ביקש: שמה שלמעלה יהיה גם למטה.
 *
 * ## ולמה הזרקה ולא סינגלטון במודול
 *
 * מצב ברמת המודול נשמר גם בין הרכבה להרכבה, כלומר בין בדיקה לבדיקה: לשונית
 * שהוחלפה או בדיקה שהחילה גופן היו מזהמות את מי שבא אחריהן. `FONT_MEMORY`
 * נוצר פעם אחת במעטפת (`App.vue`) ומוזרק, ומי שמורכב בלעדיו — בדיקת רכיב,
 * או רצועה שעולה לפני המעטפת — מקבל זיכרון פרטי חדש. אותה תבנית בדיוק של
 * `useFontOptions`.
 */
import {
  computed,
  inject,
  onUnmounted,
  ref,
  shallowRef,
  watch,
  type ComputedRef,
  type Ref,
} from 'vue';
import { DOCUMENT_GENERATION, FONT_MEMORY, READOUT_SELECTION } from './keys';
import { useCommand } from './useCommand';
import { useFontOptions } from './useFontOptions';
import { createFontPreview } from './font-preview';
import { createFontSample } from './font-sample';
import { familyPickerOptions } from './font-family-options';
import { applyOptimistically, withCurrent, type PickerOption } from './picker-value';
import { ACTIVE_SUPERDOC } from '../engine/document-api';
import { ensureFamilyDrawable } from '../engine/picker-fonts';
import { captureRange, paintFamily, readSelectionText } from '../engine/font-preview';
import { UNSETTLED_SELECTION } from '../engine/readout-hold';
import {
  DEFAULT_FONT_SIZE_PT,
  fontFamilyPayload,
  fontSizePayload,
  grownFontSize,
  parseFontFamily,
  parseFontSizeInput,
  parseFontSizePt,
  shrunkFontSize,
} from '../engine/payloads';

/**
 * מה שהבורר מציג לפני שהמנוע דיווח על גופן. Assistant הוא הגופן הארוז, כלומר
 * היחיד שבטוח קיים בכל פלטפורמה.
 */
export const DEFAULT_FONT_FAMILY = 'Assistant';

/**
 * טווח הגודל שבו פס הדגימה מצייר, בפיקסלים. ראו `sampleSize` — הפס יושב ברשימה
 * בגובה קבוע, וגדלים חריגים (6pt, 72pt) היו קטנים מכדי להראות צורה או גדולים
 * מכדי להיכנס. הגבול העליון תפור לתיבת התוכן של `.ribbon-combo-sample`:
 * `line-height: 30px` בתוך גובה 44px — שינוי שם מחייב שינוי כאן.
 */
export const SAMPLE_SIZE_MIN_PX = 12;
export const SAMPLE_SIZE_MAX_PX = 24;

/**
 * הגודל שבו פס הדגימה מצייר את הטקסט המסומן — מחרוזת CSS מוכנה, מגודל הבחירה.
 *
 * זה מה שהופך את הפס ל„איך הטקסט שלי ייראה” ולא רק „הנה האותיות”: הוא מצייר
 * את הבחירה בגופן שמרחפים עליו **ובגודל שבמסמך**. ההמרה היא `pt × 4/3` —
 * ‎96dpi, כלומר בדיוק היחס שבו הדפדפן עצמו מצייר `pt`.
 *
 * החסימה ל-[`SAMPLE_SIZE_MIN_PX`, `SAMPLE_SIZE_MAX_PX`] אינה קוסמטית: הפס יושב
 * ברשימה נגללת בגובה קבוע. ‎72pt הם 96px והיו שוברים את הרשימה, ו-6pt קטנים
 * מכדי להראות צורת אות. בטווח הנפוץ (12–18pt) הגודל יוצא מדויק, ומעליו נשמרת
 * הצורה והמידה מוקטנת — הצורה היא מה שהמשתמש בא לראות.
 *
 * פונקציה טהורה ולא `computed` בגוף הקומפוזבל, כדי שהחוזה הזה ייבדק ישירות
 * (`tests/unit/font-sample.test.ts`) בלי להרכיב קומפוננטה.
 *
 * גודל שאינו מספר סופי (מנוע ששתק, קלט פגום) נופל ל-`SAMPLE_SIZE_MIN_PX`:
 * `NaN` היה יוצא `'NaNpx'`, כלומר `font-size` פסול שהדפדפן מתעלם ממנו בשקט.
 */
export function sampleSizePx(pt: number): string {
  if (!Number.isFinite(pt)) return `${SAMPLE_SIZE_MIN_PX}px`;
  const px = pt * (4 / 3);
  const clamped = Math.min(SAMPLE_SIZE_MAX_PX, Math.max(SAMPLE_SIZE_MIN_PX, px));
  return `${Math.round(clamped)}px`;
}

/** שתי שכבות הזיכרון של הבוררים. ראו הערת הראש. */
export interface FontMemory {
  /** מה שהמנוע דיווח לאחרונה. */
  family: Ref<string>;
  size: Ref<number>;
  /** מה שהמשתמש בחר וטרם קיבל תשובה. `null` = אין בקשה באוויר. */
  pendingFamily: Ref<string | null>;
  pendingSize: Ref<number | null>;
}

export function createFontMemory(): FontMemory {
  return {
    family: ref(DEFAULT_FONT_FAMILY),
    size: ref(DEFAULT_FONT_SIZE_PT),
    pendingFamily: ref<string | null>(null),
    pendingSize: ref<number | null>(null),
  };
}

/** מה שפקד רואה. הכול לקריאה — הכתיבה עוברת בפקודות. */
export interface FontControls {
  familyOptions: ComputedRef<readonly PickerOption[]>;
  sizeOptions: ComputedRef<readonly PickerOption[]>;
  /** הגופן המוצג. */
  family: ComputedRef<string>;
  /** הגודל המוצג, כמחרוזת של הבורר. „12” ולא „12.0”, אבל „20.5” נשמר. */
  size: ComputedRef<string>;
  familyEnabled: ComputedRef<boolean>;
  sizeEnabled: ComputedRef<boolean>;
  setFamily: (family: string) => void;
  setSize: (size: string) => void;
  /**
   * מה שהקלדה בתיבת הגודל מחילה. `null` = לא מספר, והתיבה חוזרת לגודל שבמסמך.
   *
   * מחרוזת ולא מספר מפני שזה מה שהפקד מדבר בו — הוא אינו יודע שמדובר בגדלים,
   * וההמרה חזרה למספר קורית ב-`setSize` יחד עם כל בחירה אחרת.
   */
  normalizeSize: (typed: string) => string | null;
  /** דרגה אחת על סולם הגדלים של Word, **מהערך של המנוע**. */
  grow: () => void;
  shrink: () => void;
  /**
   * תצוגה חיה: הסימון ברשימת הגופנים עבר לגופן הזה. `null` = אין סימון.
   *
   * הקטע המסומן במסמך נצבע בו אחרי השהיה קצרה, כך שרואים את הגופן בטקסט
   * האמיתי ולא רק בשם. ראו font-preview.ts (העיתוי) ו-engine/font-preview.ts
   * (המגע במסמך), כולל מה שמונע תצוגה על בחירה מעורבת.
   */
  hoverFamily: (family: string | null) => void;
  /** רשימת הגופנים נסגרה. `committed` = נבחר גופן, ולכן אין מה להחזיר. */
  endHoverFamily: (committed: boolean) => void;
  /**
   * מה שפס הדגימה בתחתית רשימת הגופנים מציג — הטקסט המסומן של המשתמש, או
   * משפט ברירת מחדל כשאין בחירה.
   *
   * זו התצוגה **שכן נשלחת**, בשונה מ-`hoverFamily` שמעליה: היא עונה על אותה
   * שאלה („איך זה ייראה בטקסט שלי”) בלי לגעת במסמך. ההנמקה המלאה, כולל מה
   * שנמדד ונשלל בצביעה במסמך, ב-`engine/font-preview.ts`.
   */
  sampleText: ComputedRef<string>;
  /**
   * הגודל שבו פס הדגימה מצייר — מחרוזת CSS (`'18px'`), נגזרת מגודל הבחירה
   * במסמך וחסומה לטווח שנכנס בפס. זה מה שהופך את הדגימה לנאמנה: הטקסט שלך,
   * בגופן הזה, **ובגודל שלו** — ולא בגודל קבוע של הרשימה.
   *
   * `''` כשאין בחירה כלל, ואז הפס מציג את פסוק ברירת המחדל בגודל הקבוע של
   * הרשימה. ראו את ההנמקה במימוש.
   */
  sampleSize: ComputedRef<string>;
}

export function useFontControls(): FontControls {
  // הצורה עם factory (`true`) ולא ערך ישיר: אחרת זיכרון נבנה בכל הרכבה גם
  // כשהמעטפת כן מספקת את המפתח.
  const memory = inject(FONT_MEMORY, createFontMemory, true);

  const familyCmd = useCommand('font-family');
  const sizeCmd = useCommand('font-size');

  /**
   * המופע הפעיל והבחירה — שניהם רק בשביל התצוגה החיה.
   *
   * ברירות מחדל שאינן מציגות דבר: קומפוננטה שמורכבת בלי המעטפת (בדיקת רכיב,
   * רצועה שעולה לפני שנפתח מסמך) מקבלת `null` ובחירה שלא התיישבה, ואז
   * `allowed` למטה מחזיר `false` — כלומר הבורר עובד בדיוק כמו קודם.
   */
  const host = inject(ACTIVE_SUPERDOC, () => shallowRef(null), true);
  const selection = inject(READOUT_SELECTION, () => shallowRef(UNSETTLED_SELECTION), true);

  /**
   * מסמך אחר — והזיכרון של הקודם נמחק.
   *
   * ## למה זה אינו „זהירות” אלא באג שנסגר
   *
   * שתי שכבות הזיכרון קיימות בשביל רגע אחד: המנוע אינו מדווח ערך (בחירה
   * מעורבת, או בחירה שטרם נפתרה), ואז הבורר מציג את **האחרון שידענו** במקום
   * להתרוקן. הצדקה שלהן היא שהערך הזה נמדד באותו מסמך — ובדיוק זה מה שנשבר
   * בהחלפה: „TaameyDavidCLM 20” של הספר שנסגר אינו „האחרון שידענו” על הטאב
   * שנפתח עכשיו, הוא ניחוש על מסמך אחר. וזה בדיוק הרגע שבו המנוע שותק —
   * מסמך טרי, לפני שהבחירה התיישבה — כלומר הזיהום נראה תמיד.
   *
   * `DOCUMENT_GENERATION` ולא זהות `ACTIVE_SUPERDOC`: אותו מסמך שנטען מחדש
   * ברכיב (טאב שהוחלף וחזר) אינו „מסמך אחר”, והזיכרון שם דווקא נכון. ראו
   * composables/keys.ts.
   *
   * גם שכבת ה-`pending` מתאפסת, ומאותו טעם: בקשה שיצאה למסמך שכבר אינו על
   * המסך אינה אמורה להיות מוצגת. התשובה שלה תיפול ממילא על השומר
   * `pending.value !== next` ב-`applyOptimistically`.
   */
  const documentGeneration = inject(DOCUMENT_GENERATION, () => shallowRef(0), true);
  watch(documentGeneration, () => {
    memory.family.value = DEFAULT_FONT_FAMILY;
    memory.size.value = DEFAULT_FONT_SIZE_PT;
    memory.pendingFamily.value = null;
    memory.pendingSize.value = null;
  });

  /**
   * אפשרויות הגופן מהמנוע (`ui.fonts`), ממוזגות עם שלנו. ראו
   * engine/font-options.ts — רשימה קשיחה לא הייתה יודעת מה יש במסמך.
   */
  const { families, sizes } = useFontOptions();

  /**
   * `CommandState.value` הוא המקור לערך שהבורר מציג — לא ref מקומי. שלושת
   * הבוררים היו פעם refs שאותחלו לערך קשיח ולעולם לא התעדכנו: לחיצה על טקסט
   * ב-20pt השאירה „12” בתיבה, ו„הגדל גופן” חישב מהמספר השגוי הזה.
   */
  const engineFamily = computed(() => parseFontFamily(familyCmd.value.value));
  const engineSize = computed(() => parseFontSizePt(sizeCmd.value.value));

  // הזיכרון נכתב רק מדיווח שיש בו ערך — „מעורב” אינו מוחק אותו.
  watch(engineFamily, (value) => {
    if (value) memory.family.value = value;
  });
  watch(engineSize, (value) => {
    if (value) memory.size.value = value;
  });

  /**
   * סדר העדיפויות: מה שנבחר וטרם נענה, אחר כך מה שהמנוע מדווח, ולבסוף האחרון
   * שידענו. בקשה שנדחתה נעלמת מהשכבה הראשונה, ואז המסמך חוזר להיות מה שמוצג.
   */
  const reported = computed(
    () => memory.pendingFamily.value ?? engineFamily.value ?? memory.family.value,
  );

  /**
   * הגופן המוחל, **באיות של הרשימה**.
   *
   * המנוע מדווח את מה שכתוב במסמך, והמסמך אינו מחויב לאיות שלנו: ריצה שמצהירה
   * „APTOS” מול רשימה שמחזיקה „Aptos” היא אותו גופן בדיוק — ההתאמה ב-CSS
   * ובמנוע חסרת רגישות לאותיות, ו-`familyKey` במיזוג מאחד אותם.
   *
   * הפקד, לעומת זאת, משווה **מדויק** בארבעה מקומות. בלי הנרמול כאן היו שתי
   * תוצאות רעות ואין שלישית: או שורה כפולה לאותו גופן, או — וזה מה שנמדד —
   * רשימה בלי שום שורה מסומנת, שגם נפתחת מהראש במקום על הגופן שבמסמך.
   */
  const family = computed(() => {
    const value = reported.value;
    const key = value.trim().toLowerCase();
    return families.value.find((option) => option.value.trim().toLowerCase() === key)?.value ?? value;
  });
  const sizePt = computed(() => memory.pendingSize.value ?? engineSize.value ?? memory.size.value);
  const size = computed(() => String(sizePt.value));

  /*
   * המיפוי עצמו יצא ל-composables/font-family-options.ts: הוא נדרש עכשיו גם
   * בשני דיאלוגים („גופן מתקדם”, „ברירות מחדל למסמך”) שקודם ביקשו להקליד שם
   * גופן ביד, ומיפוי מועתק הוא מיפוי שיתוקן פעם אחת בלבד.
   */
  const familyOptions = computed(() => familyPickerOptions(families.value, family.value));

  const sizeOptions = computed(() =>
    withCurrent(
      sizes.value.map((option) => ({ value: option.value, label: option.label })),
      size.value,
      // `numeric` ובלי `preview`: ראו composables/picker-value.ts — 13pt נכנס
      // בין 12 ל-14 ואינו מכריז על `font-family` בשם „13”.
      { numeric: true },
    ),
  );

  // ה-payloads נבנים ב-engine/payloads.ts ולא כליטרל כאן: מה שנשלח לפקודה הוא
  // חוזה מול ולידטור בתוך המנוע, והוולידטור נכשל **סגור**.
  function setFamily(next: string): void {
    const payload = fontFamilyPayload(next);
    if (payload === null) return;
    void applyOptimistically(memory.pendingFamily, memory.family, payload, () =>
      familyCmd.run(payload),
    );
    /*
     * הגופן שנבחר אינו בהכרח גופן שהדפדפן מצייר, וזה תוקן אחרי שנמדד: המנייה
     * מדווחת שמות GDI, ווריאנט משקל כמו `Guttman Kav-Light` נשאר שם שאינו
     * נפתר — 43 כאלה מ-287 במכונה שנמדדה. בלי השליפה הטקסט נכתב עם השם הנכון
     * ב-OOXML אבל מצויר ב-fallback.
     *
     * **אחרי** ההחלה ולא לפניה, ובלי `await`: ההחלה היא מה שהמשתמש ביקש והיא
     * אינה אמורה להמתין לרשת של המארח. השליפה עומדת בפני עצמה — היא מזריקה
     * `@font-face` ומודיעה, והבורר מתעורר בעצמו (App.vue).
     *
     * אינה זורקת בשום מסלול, ולכן אין כאן `catch`: ראו `ensureFamilyDrawable`.
     */
    void ensureFamilyDrawable(next);
  }

  function applySize(pt: number): void {
    const payload = fontSizePayload(pt);
    if (payload === null) return;
    void applyOptimistically(memory.pendingSize, memory.size, payload, () => sizeCmd.run(payload));
  }

  function setSize(next: string): void {
    const pt = parseFontSizePt(next);
    if (pt !== null) applySize(pt);
  }

  function normalizeSize(typed: string): string | null {
    const pt = parseFontSizeInput(typed);
    return pt === null ? null : String(pt);
  }

  /**
   * הגופן שהמנוע מדווח, **בלי** ההחזקה של readout-hold ובלי הזיכרון.
   *
   * זו ההבחנה שהתצוגה החיה נשענת עליה: `family` למעלה מציג את „האחרון שידענו”
   * מפני שבורר שמתרוקן בכל תזוזת סמן גרוע מבורר מעופש — אבל **להחזיר** אפשר
   * רק גופן שבאמת נמצא במסמך עכשיו. `null` כאן פירושו „מעורב, או טרם נפתר”,
   * ובשני המצבים אין תצוגה חיה: צביעה משטחת בחירה מעורבת לגופן אחד, וההחזרה
   * לא הייתה יודעת להשיב את מה שהיה.
   */
  const rawFamily = computed(() => parseFontFamily(familyCmd.engineState.value.value));

  const preview = createFontPreview({
    allowed: () =>
      familyCmd.enabled.value &&
      selection.value.settled &&
      !selection.value.empty &&
      rawFamily.value !== null,
    origin: () => rawFamily.value,
    capture: () => captureRange(host.value),
    paint: (target, family) => paintFamily(host.value, target, family),
  });

  /**
   * רשת ביטחון: לשונית שהוחלפה או תפריט שנהרס בזמן שהתצוגה החיה על המסך אינם
   * אמורים להשאיר את הגופן שאיש לא בחר. בזרימה רגילה `end` כבר נקרא מהסגירה.
   */
  onUnmounted(() => preview.end(false));

  /**
   * פס הדגימה. שני האותות שמזיזים אותו הם בדיוק אלה שמזיזים את התצוגה החיה
   * הכבויה — הסימון זז, והרשימה נסגרה — ולכן הם עוברים דרך אותן שתי פונקציות.
   *
   * הפרדה מלאה הייתה מחייבת את הפקד לפלוט „נפתחתי” בנפרד, וזו פליטה שלישית
   * לאותו מידע. ראו `composables/font-sample.ts`.
   */
  const sample = createFontSample({ read: () => readSelectionText(host.value) });

  /**
   * הגודל שבו פס הדגימה מצייר. החישוב עצמו ב-[sampleSizePx].
   *
   * `''` כשאין בחירה, וזו אינה קיצור דרך: הגודל שבמסמך הוא ההבטחה „**הטקסט
   * שלך** ייראה כך”, ובלי בחירה אין טקסט כזה — מה שמוצג הוא פסוק שכל תפקידו
   * להראות את כל האותיות. הוא ארוך פי שניים מהבחירה הטיפוסית, ובגודל של
   * המסמך (עד 24px) לא היה נכנס בפס בשום גופן; הפס מקבל אותו בגודל הקבוע של
   * הרשימה, בשתי שורות. מה שריק כאן הוא מה שמדליק את המצב הזה ב-`RibbonCombo`
   * — ולכן אין כאן דגל נוסף לאותה ידיעה.
   */
  const sampleSize = computed(() => (sample.isFallback.value ? '' : sampleSizePx(sizePt.value)));

  function hoverFamily(family: string | null): void {
    if (family !== null) sample.begin();
    preview.hover(family);
  }

  function endHoverFamily(committed: boolean): void {
    sample.end();
    preview.end(committed);
  }

  return {
    familyOptions,
    sizeOptions,
    family,
    size,
    familyEnabled: familyCmd.enabled,
    sizeEnabled: sizeCmd.enabled,
    setFamily,
    setSize,
    normalizeSize,
    grow: () => applySize(grownFontSize(sizePt.value)),
    shrink: () => applySize(shrunkFontSize(sizePt.value)),
    hoverFamily,
    endHoverFamily,
    sampleText: sample.text,
    sampleSize,
  };
}

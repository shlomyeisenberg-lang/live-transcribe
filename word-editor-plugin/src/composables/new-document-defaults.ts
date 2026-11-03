/**
 * ברירות המחדל של מסמך חדש — המצב המשותף לדיאלוגים שכותבים אותן ולפתיחת
 * מסמך חדש שקוראת אותן. העדפה של התוכנה ולא תכונה של מסמך, כמו
 * composables/paste-default.ts; הכתיבה ל-XML ב-engine/new-document-defaults.ts.
 */
import { shallowRef } from 'vue';
import type { StyleCatalogItem } from 'superdoc/ui';
import { styleDisplayLabel } from '../engine/style-gallery';
import { loadNewDocumentDefaults as loadStored, saveNewDocumentDefaults } from '../host/settings';
import { blankDocumentSource } from '../engine/blank-document-source';
import {
  customizeBlankDocx,
  hasNewDocumentDefaults,
  normalizeNewDocumentDefaults,
  newDocumentDefaultsSupported,
  NO_NEW_DOCUMENT_DEFAULTS,
  readTemplateStyles,
  rememberStyle,
  withRunDefaults,
  type NewDocumentDefaults,
  type RunDefaults,
} from '../engine/new-document-defaults';
import { listStyles, type StyleFormat } from '../engine/docx-style-edit';

/**
 * הודעות שורת המצב של „ברירות מחדל”. כאן ולא ב-LayoutTab.vue: כל מחרוזת עברית
 * בקוד של הרצועה נחשבת פריט שמוצג דרך `menuString` וחייבת אנגלית
 * (tests/unit/menu-locale-coverage.test.ts), ושורת המצב אינה מתורגמת — כמו
 * ההודעות של composables/use-style-manager.ts.
 */
export const RUN_DEFAULTS_SAVED_MESSAGE = 'ברירת המחדל תחול גם על כל מסמך חדש';
export const UNSUPPORTED_STYLE_DEFAULTS_MESSAGE = 'עיצוב הסגנונות אינו נתמך במסמכים חדשים בלי הבדל בין העורך ל-Word; ההעדפה לא נשמרה';
export const STORED_STYLE_DEFAULTS_REJECTED_MESSAGE = 'ברירות מחדל שמורות לסגנונות לא הוחלו: העיצוב אינו נתמך בלי הבדל בין העורך ל-Word. ניתן לקבוע אותן מחדש או לבחור „החזר למקור”';
let storedStylesRejected = false;

export const NEW_DOCUMENT_DEFAULTS_RESET_MESSAGE = 'מסמך חדש ייפתח שוב בברירות המחדל המקוריות';

/** מה שחל עכשיו על כל מסמך חדש. */
export const newDocumentDefaults = shallowRef<NewDocumentDefaults>(NO_NEW_DOCUMENT_DEFAULTS);

let loading: Promise<void> | null = null;
let userChoiceVersion = 0;

/**
 * נקראת בעלייה, ושוב — בלי קריאה נוספת לגשר — מכל מי שצריך את הערך לפני שהיא
 * הסתיימה: מסמך חדש שנפתח ברגע העלייה חייב לחכות לה, אחרת הוא נפתח בלי
 * ההגדרות שהמשתמש קבע.
 */
export function loadNewDocumentDefaults(): Promise<void> {
  loading ??= (async () => {
    const version = userChoiceVersion;
    const stored = await loadStored();
    if (version !== userChoiceVersion) return;
    const next = normalizeNewDocumentDefaults(stored);
    const xml = Object.keys(next.styles).length > 0 ? await templateStyles() : null;
    if (version !== userChoiceVersion) return;
    storedStylesRejected = Object.keys(next.styles).length > 0 && (!xml || !newDocumentDefaultsSupported(xml, next));
    newDocumentDefaults.value = storedStylesRejected ? { run: next.run, styles: {} } : next;
  })();
  return loading;
}

function store(next: NewDocumentDefaults): void {
  userChoiceVersion += 1;
  storedStylesRejected = false;
  newDocumentDefaults.value = next;
  void saveNewDocumentDefaults(hasNewDocumentDefaults(next) ? next : null);
}

/** „הגדר כברירת מחדל” של דיאלוג הגופן. */
export async function rememberRunDefaults(run: RunDefaults): Promise<void> {
  await loadNewDocumentDefaults();
  store(withRunDefaults(newDocumentDefaults.value, run));
}

/**
 * „מסמכים חדשים” של דיאלוג הסגנון. `format` הוא העיצוב השלם שהדיאלוג החזיר;
 * מה שנשמר הוא רק ההבדל ממה שמסמך חדש היה מקבל בלי הטלאי של הסגנון הזה —
 * ולכן גם סימון בלי שינוי, על סגנון שכבר תואם לתבנית, אינו שומר דבר.
 */
export async function rememberStyleDefault(id: string, format: StyleFormat): Promise<boolean> {
  await loadNewDocumentDefaults();
  const xml = await templateStyles();
  const next = rememberStyle(newDocumentDefaults.value, xml, id, format);
  if (!next) return false;
  store(next);
  return true;
}

let templateXml: Promise<string | null> | null = null;

/** ‏`styles.xml` של התבנית המקורית — נקרא פעם אחת; התבנית אינה משתנה בריצה. */
function templateStyles(): Promise<string | null> {
  if (!templateXml) {
    const blank = blankDocumentSource();
    templateXml = blank ? readTemplateStyles(blank) : Promise.resolve(null);
  }
  return templateXml;
}

/** האם לסגנון יש מקום בתבנית — רק אז יש משמעות ל„גם לכל מסמך חדש”. */
export async function templateHasStyle(id: string): Promise<boolean> {
  return listStyles(await templateStyles()).some((style) => style.id === id);
}

/** חזרה לתבנית המקורית. */
export function resetNewDocumentDefaults(): void {
  store(NO_NEW_DOCUMENT_DEFAULTS);
}

/** התבנית שמסמך חדש נפתח ממנה, אחרי ההגדרות. `undefined` כשאין תבנית (בדיקות). */
export async function newDocumentSource(onWarning?: (message: string) => void): Promise<Blob | undefined> {
  const blank = blankDocumentSource();
  if (!blank) return undefined;
  await loadNewDocumentDefaults();
  if (storedStylesRejected) onWarning?.(STORED_STYLE_DEFAULTS_REJECTED_MESSAGE);
  return customizeBlankDocx(blank, newDocumentDefaults.value);
}

/** מה שנקבע כרגע, בשורה אחת לדיאלוג. `null` = התבנית המקורית. */
export async function describeNewDocumentDefaults(): Promise<string | null> {
  await loadNewDocumentDefaults();
  const { run, styles } = newDocumentDefaults.value;
  if (!hasNewDocumentDefaults(newDocumentDefaults.value) && !storedStylesRejected) return null;
  const parts: string[] = storedStylesRejected ? [STORED_STYLE_DEFAULTS_REJECTED_MESSAGE] : [];
  if (run.fontFamily) parts.push(run.fontFamily);
  if (run.fontSizePt) parts.push(`${run.fontSizePt} נק'`);
  const ids = Object.keys(styles);
  if (ids.length > 0) {
    const known = listStyles(await templateStyles());
    const labels = ids.map((id) => {
      const info = known.find((style) => style.id === id);
      return info ? styleDisplayLabel({ id, name: info.name, custom: info.custom } as StyleCatalogItem) : id;
    });
    parts.push(`סגנונות: ${labels.join(', ')}`);
  }
  return parts.join('; ');
}

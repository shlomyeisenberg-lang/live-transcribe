/**
 * ברירות המחדל של **מסמך חדש** שהמשתמש קבע — „הגדר כברירת מחדל” של דיאלוג
 * הגופן ו„מסמכים חדשים המבוססים על תבנית זו” של שינוי סגנון ב-Word.
 *
 * ## מה נשמר
 *
 * שני דברים, ושניהם ביחידות של המשתמש (נקודות, ס"מ), כמו `StyleFormat`:
 *
 * - `run` — גופן וגודל של המסמך כולו. נכתבים ל-`w:docDefaults/w:rPrDefault`,
 *   ולא לסגנון „רגיל”: „ברירות מחדל למסמך” במסמך שנפתח כותב לשם, וסגנון
 *   „רגיל” שמצהיר גופן היה מסתיר את מה שהדיאלוג הזה כותב אחר כך.
 * - `styles` — טלאי לכל סגנון (`Normal`, `Heading1`…) **ביחס לתבנית**: רק מה
 *   שהמשתמש קבע אחרת ממה שמסמך חדש היה מקבל בלעדיו. כך שינוי של „רגיל”
 *   ממשיך לזרום לכותרות שמבוססות עליו, ולא קופא בהן כעותק.
 *
 * ## איפה זה חל
 *
 * רק על התבנית של מסמך חדש (`customizeBlankDocx`), לפני שהמנוע פותח אותה —
 * ב-`styles.xml` עצמו, כך שמה שנשמר ל-DOCX הוא סגנונות רגילים ש-Word קורא
 * כמו כל מסמך. קובץ קיים שנפתח אינו עובר כאן לעולם.
 */
import { readZip, rewriteDocxXmlParts, readEntryText, type Bytes } from './docx-parts';
import { editRunDefaults, editStyles, listStyles, liveStyleTogglesSupported, resolveFormat, type StyleAlignment, type StyleFormat } from './docx-style-edit';
import { formatPatch } from './style-from-selection';

export interface RunDefaults {
  fontFamily?: string;
  fontSizePt?: number;
}

export interface NewDocumentDefaults {
  run: RunDefaults;
  /** טלאי לפי מזהה סגנון. סגנון שאינו בתבנית מדולג. */
  styles: Readonly<Record<string, StyleFormat>>;
}

export const NO_NEW_DOCUMENT_DEFAULTS: NewDocumentDefaults = Object.freeze({ run: {}, styles: {} });

const STYLES_PART = 'word/styles.xml';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** תקרות — הרשומה נכתבת ל-storage של אוצריא, ורשומה פגומה אינה אמורה להתנפח. */
const MAX_STYLES = 40;
const STYLE_ID = /^[A-Za-z0-9._-]{1,64}$/;
const ALIGNMENTS: readonly StyleAlignment[] = ['right', 'center', 'left', 'justify'];

const isNumber = (value: unknown, min: number, max: number): value is number =>
  typeof value === 'number' && Number.isFinite(value) && value >= min && value <= max;

function fontName(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() !== '' && value.length <= 100 ? value.trim() : undefined;
}

/** שדה פגום נזרק לבדו; השאר נשמר. `null` (= לרשת) נשמר רק היכן שהוא חוקי. */
function normalizeFormat(raw: unknown): StyleFormat {
  if (!raw || typeof raw !== 'object') return {};
  const source = raw as Record<string, unknown>;
  const out: StyleFormat = {};
  const font = fontName(source.fontFamily);
  if (font !== undefined) out.fontFamily = font;
  if (isNumber(source.fontSizePt, 1, 400)) out.fontSizePt = source.fontSizePt;
  for (const key of ['bold', 'italic', 'underline'] as const) {
    if (typeof source[key] === 'boolean') out[key] = source[key] as boolean;
  }
  if (typeof source.color === 'string' && /^[0-9A-Fa-f]{6}$/.test(source.color)) out.color = source.color.toUpperCase();
  if (typeof source.alignment === 'string' && ALIGNMENTS.includes(source.alignment as StyleAlignment)) {
    out.alignment = source.alignment as StyleAlignment;
  }
  for (const key of ['spaceBeforePt', 'spaceAfterPt'] as const) {
    if (isNumber(source[key], 0, 1584)) out[key] = source[key] as number;
  }
  if (isNumber(source.lineSpacing, 0.25, 10)) out.lineSpacing = source.lineSpacing;
  for (const key of ['indentStartCm', 'indentEndCm'] as const) {
    if (isNumber(source[key], 0, 55)) out[key] = source[key] as number;
  }
  if (isNumber(source.firstLineCm, -55, 55)) out.firstLineCm = source.firstLineCm;
  if (isNumber(source.outlineLevel, 0, 9) && Number.isInteger(source.outlineLevel)) out.outlineLevel = source.outlineLevel;
  return out;
}

/** הרשומה מה-storage. כל מה שאינו נקרא נופל ל„אין ברירות מחדל”. */
export function normalizeNewDocumentDefaults(raw: unknown): NewDocumentDefaults {
  if (!raw || typeof raw !== 'object') return NO_NEW_DOCUMENT_DEFAULTS;
  const source = raw as { run?: unknown; styles?: unknown };
  const run: RunDefaults = {};
  if (source.run && typeof source.run === 'object') {
    const { fontFamily, fontSizePt } = source.run as Record<string, unknown>;
    const font = fontName(fontFamily);
    if (font !== undefined) run.fontFamily = font;
    if (isNumber(fontSizePt, 0.5, 800)) run.fontSizePt = fontSizePt;
  }
  const styles: Record<string, StyleFormat> = {};
  if (source.styles && typeof source.styles === 'object') {
    for (const [id, format] of Object.entries(source.styles as Record<string, unknown>)) {
      if (Object.keys(styles).length >= MAX_STYLES) break;
      if (!STYLE_ID.test(id)) continue;
      const normalized = normalizeFormat(format);
      if (Object.keys(normalized).length > 0) styles[id] = normalized;
    }
  }
  return { run, styles };
}

export function hasNewDocumentDefaults(defaults: NewDocumentDefaults): boolean {
  return Object.keys(defaults.run).length > 0 || Object.keys(defaults.styles).length > 0;
}

/**
 * גופן/גודל חדשים מתמזגים עם מה שנשמר: שדה שלא נשלח נשאר.
 *
 * ואותו שדה יורד מהטלאי של „רגיל”: סגנון „רגיל” שמצהיר גופן או גודל מסתיר את
 * `rPrDefault`, ולכן „גם לכל מסמך חדש” של „ברירות מחדל” היה נבלע בשקט אחרי
 * „שנה סגנון” על „רגיל” — למשל במסמך מהדיסק שבו „רגיל” מצהיר גופן, שנשמר
 * כולו לטלאי. הבחירה האחרונה של המשתמש היא שקובעת.
 */
export function withRunDefaults(defaults: NewDocumentDefaults, run: RunDefaults): NewDocumentDefaults {
  const next: RunDefaults = { ...defaults.run };
  const font = fontName(run.fontFamily);
  if (font !== undefined) next.fontFamily = font;
  if (isNumber(run.fontSizePt, 0.5, 800)) next.fontSizePt = run.fontSizePt;

  const normal = defaults.styles.Normal;
  if (!normal) return { run: next, styles: defaults.styles };
  const patch: StyleFormat = { ...normal };
  if (font !== undefined) delete patch.fontFamily;
  if (isNumber(run.fontSizePt, 0.5, 800)) delete patch.fontSizePt;
  const styles = { ...defaults.styles };
  if (Object.keys(patch).length === 0) delete styles.Normal;
  else styles.Normal = patch;
  return { run: next, styles };
}

/** מחליפה את הטלאי של סגנון אחד. טלאי ריק מוחק אותו — הסגנון חוזר לתבנית. */
export function withStyleDefault(defaults: NewDocumentDefaults, id: string, patch: StyleFormat): NewDocumentDefaults {
  if (!STYLE_ID.test(id)) return defaults;
  const styles = { ...defaults.styles };
  const normalized = normalizeFormat(patch);
  if (Object.keys(normalized).length === 0) delete styles[id];
  else styles[id] = normalized;
  return { run: defaults.run, styles };
}

/** בלי הסגנון הזה — מה שמסמך חדש היה מקבל אילו הטלאי שלו לא היה. */
export function withoutStyleDefault(defaults: NewDocumentDefaults, id: string): NewDocumentDefaults {
  if (!(id in defaults.styles)) return defaults;
  const styles = { ...defaults.styles };
  delete styles[id];
  return { run: defaults.run, styles };
}

/**
 * שומרת את העיצוב השלם שהדיאלוג החזיר כטלאי של הסגנון: רק ההבדל ממה שמסמך
 * חדש היה מקבל בלי הטלאי הקודם של אותו סגנון.
 * כשאין תבנית או שהגרף הסופי אינו נתמך — null, ואין לשמור את ההעדפה.
 */
export function rememberStyle(
  defaults: NewDocumentDefaults,
  templateXml: string | null,
  id: string,
  format: StyleFormat,
): NewDocumentDefaults | null {
  if (!templateXml) return null;
  const without = withoutStyleDefault(defaults, id);
  const baseline = resolveFormat(applyNewDocumentDefaults(templateXml, without), id);
  const next = withStyleDefault(without, id, formatPatch(format, baseline));
  return newDocumentDefaultsSupported(templateXml, next) ? next : null;
}

/**
 * Build the whole graph before validating it: a later parent edit can invalidate
 * an earlier child. Apply ancestors first, independently of storage key order,
 * because editStyles encodes bold/italic relative to the inherited Word value.
 * Reject the style group atomically; safe run defaults can still be used.
 */
function checkedDefaults(stylesXml: string, defaults: NewDocumentDefaults): { xml: string; supported: boolean } {
  const base = Object.keys(defaults.run).length > 0 ? editRunDefaults(stylesXml, defaults.run) ?? stylesXml : stylesXml;
  const known = new Map(listStyles(base).map(style => [style.id, style]));
  const visited = new Set<string>();
  const ordered: string[] = [];
  const visit = (id: string): void => {
    if (visited.has(id)) return;
    visited.add(id);
    const parent = known.get(id)?.basedOn;
    if (parent && known.has(parent)) visit(parent);
    if (known.has(id) && Object.prototype.hasOwnProperty.call(defaults.styles, id)) ordered.push(id);
  };
  Object.keys(defaults.styles).sort().forEach(visit);
  let xml = base;
  for (const id of ordered) {
    const edited = editStyles(xml, { kind: 'modify', id, format: defaults.styles[id] });
    if (!edited.ok) return { xml: base, supported: false };
    xml = edited.xml;
  }
  // Compare the final graph with the original template, including descendants
  // and both script stacks. Requested toggle fields also get the live guard's
  // stricter check for an explicit request in an already inconsistent style.
  if (ordered.length > 0 && !liveStyleTogglesSupported(base, xml, { kind: 'modify', id: '', format: {} }, '', defaults.styles)) {
    return { xml: base, supported: false };
  }
  return { xml, supported: true };
}

export function newDocumentDefaultsSupported(stylesXml: string, defaults: NewDocumentDefaults): boolean {
  return checkedDefaults(stylesXml, defaults).supported;
}

/** Unsafe stored style groups fall back to the template; callers can announce this. */
export function applyNewDocumentDefaults(stylesXml: string, defaults: NewDocumentDefaults): string {
  return checkedDefaults(stylesXml, defaults).xml;
}

async function bytesOf(blob: Blob): Promise<Bytes> {
  return new Uint8Array(await blob.arrayBuffer());
}

/** ‏`styles.xml` של התבנית, או `null`. */
export async function readTemplateStyles(blank: Blob): Promise<string | null> {
  const entry = readZip(await bytesOf(blank))?.find((candidate) => candidate.name === STYLES_PART);
  return entry ? readEntryText(entry) : null;
}

/**
 * התבנית של מסמך חדש אחרי ברירות המחדל של המשתמש. בלי ברירות מחדל, או
 * כשהחבילה אינה נקראת — התבנית כמות שהיא: מסמך חדש נפתח תמיד.
 */
export async function customizeBlankDocx(blank: Blob, defaults: NewDocumentDefaults): Promise<Blob> {
  if (!hasNewDocumentDefaults(defaults)) return blank;
  // ‏`rewriteDocxXmlParts` אינה זורקת: חבילה או חלק שאינם נקראים = `null`.
  const bytes = await rewriteDocxXmlParts(
    await bytesOf(blank),
    (name) => name === STYLES_PART,
    (xml) => {
      const out = applyNewDocumentDefaults(xml, defaults);
      return out === xml ? null : out;
    },
  );
  return bytes ? new Blob([bytes], { type: DOCX_MIME }) : blank;
}

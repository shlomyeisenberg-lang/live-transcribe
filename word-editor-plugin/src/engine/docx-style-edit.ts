/**
 * יצירה, שינוי, שינוי שם ומחיקה של סגנון — ב-`word/styles.xml` עצמו.
 *
 * ## למה ב-XML ולא במנוע
 *
 * `styles.apply` משנה docDefaults בלבד. את הגדרות הסגנונות משנים כאן
 * ומעבירים רק את ההגדרות שהשתנו ל-templates.apply הציבורי (SuperDoc 2.20).
 * המנוע מחיל אותן על המסמך החי כטרנזקציה: בלי פתיחה מחדש, בלי החלפת תוכן,
 * ובלי איפוס היסטוריית הביטול. מחיקת סגנון נשארת כלי XML בלבד ואינה מוצעת
 * בממשק החי, משום שאימוץ תבנית אינו מוחק סגנונות קיימים.
 *
 * ## מה המודול מבטיח
 *
 * - **רק מה שהמשתמש שינה נוגע.** כל שדה ב-`StyleFormat` הוא תלת-מצבי:
 *   חסר = לא לגעת, `null` = להסיר את ההצהרה (לרשת מהבסיס), ערך = לכתוב.
 *   הצהרות שהדיאלוג אינו מכיר (גבולות, טאבים, מספור) נשארות כמות שהן.
 * - **סדר הסכמה.** Word מסרב לפתוח `rPr`/`pPr` שילדיהם אינם בסדר של הסכמה
 *   („תוכן בלתי קריא”), ולכן כל הוספה נכנסת למקומה לפי `RPR_ORDER`/`PPR_ORDER`.
 * - **שתי המחסניות.** טקסט עברי ש-Word מסמן `w:rtl` נקרא מהצד המורכב (`bCs`,
 *   `iCs`, `szCs`, `rFonts@cs`) בכל רמות הסגנון. סגנון שקובע רק את הצד הלטיני
 *   נראה נכון בעורך ושגוי ב-Word, ולכן כל שדה נכתב לשני הצדדים יחד.
 *
 * המודול טהור — מחרוזת נכנסת, מחרוזת יוצאת — ולכן כל ההכרעות כאן נבדקות
 * ב-tests/unit בלי מנוע. ההחלה החיה ב-`live-style-edit.ts` ועריכת חבילה ב-`docx-style-package.ts`.
 */
import { TOKEN_SOURCE, attribute, isOn, skippedCloser, valueOf, wordPrefix } from './docx-parts';

/* ------------------------------------------------------------------ */
/* המודל                                                               */
/* ------------------------------------------------------------------ */

/** יישור כפי שהמשתמש רואה אותו על המסך — ראו `JC_FOR_ALIGNMENT`. */
export type StyleAlignment = 'right' | 'center' | 'left' | 'justify';

/**
 * העיצוב שהדיאלוג מציג וכותב. ביחידות של המשתמש (נקודות, ס"מ), וההמרה
 * ליחידות OOXML כאן. ראו את הכלל התלת-מצבי בהערת הראש.
 */
export interface StyleFormat {
  /** ‏`''` בקריאה = גופן ערכת נושא (ראו `formatOfLevel`). */
  fontFamily?: string | null;
  fontSizePt?: number | null;
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | null;
  /** ‏`RRGGBB`, בלי `#`. */
  color?: string | null;
  alignment?: StyleAlignment | null;
  spaceBeforePt?: number | null;
  spaceAfterPt?: number | null;
  /** כפולה של שורה: 1, 1.15, 1.5, 2. */
  lineSpacing?: number | null;
  /** הזחה מצד תחילת הטקסט (בעברית — מימין). */
  indentStartCm?: number | null;
  indentEndCm?: number | null;
  /** שורה ראשונה: חיובי = הזחה, שלילי = תלויה. */
  firstLineCm?: number | null;
  /**
   * רמת חלוקה (`w:outlineLvl`): 0–8 = כותרת ברמה 1–9, ‏9 = גוף טקסט. זה מה
   * ש-Word קורא כדי לזהות כותרת — בחלונית הניווט ובתוכן העניינים — ולכן
   * סגנון מותאם בלעדיה אינו כותרת שם, גם אם הוא נראה כמו כותרת.
   */
  outlineLevel?: number | null;
}

/** גוף טקסט — הערך של `w:outlineLvl` כשאין רמה. */
export const BODY_TEXT_LEVEL = 9;

export type StyleKind = 'paragraph' | 'character';

/** סגנון כפי שהוא בגיליון, לרשימות של הדיאלוג. */
export interface StyleInfo {
  id: string;
  name: string;
  type: string;
  basedOn: string | null;
  next: string | null;
  link: string | null;
  /** ‏`w:customStyle` — סגנון שהמשתמש יצר. רק הוא ניתן למחיקה ולשינוי שם. */
  custom: boolean;
  /** ‏`w:default` — „רגיל” של הסוג. */
  isDefault: boolean;
  hidden: boolean;
}

export type StyleEdit =
  | { kind: 'modify'; id: string; basedOn?: string | null; next?: string | null; format: StyleFormat }
  | { kind: 'create'; name: string; type: StyleKind; basedOn: string | null; next: string | null; format: StyleFormat }
  | { kind: 'rename'; id: string; name: string }
  | { kind: 'delete'; id: string };

export type StyleEditResult =
  | {
      ok: true;
      xml: string;
      /** הסגנון שנוצר או שונה. במחיקה — הסגנון שנמחק. */
      id: string;
      /** במחיקה בלבד: המזהים שההפניות אליהם יוסרו מגוף המסמך. */
      removed: readonly string[];
    }
  | { ok: false; message: string };

/* ------------------------------------------------------------------ */
/* יחידות                                                              */
/* ------------------------------------------------------------------ */

const TWIPS_PER_PT = 20;
const TWIPS_PER_CM = 1440 / 2.54;
/** ‏`lineRule="auto"` נמדד בעשרים-וארבעים חלקים של שורה. */
const LINE_UNITS = 240;

const round = (value: number): number => Math.round(value);
const twipsToCm = (twips: number): number => Math.round((twips / TWIPS_PER_CM) * 100) / 100;

/**
 * היישור על המסך → `w:jc`.
 *
 * הסגנונות של המסמכים כאן הם פסקאות `w:bidi`, וב-Transitional (מה ש-Word
 * כותב) `left`/`right` של פסקה כזאת הם **לוגיים**: `left` = תחילת השורה,
 * כלומר ימין בעברית. כך Word כותב וכך המנוע מצייר — נמדד שניהם, ראו
 * tests/unit/docx-style-edit.test.ts.
 */
export const JC_FOR_ALIGNMENT: Readonly<Record<StyleAlignment, string>> = {
  right: 'left',
  center: 'center',
  left: 'right',
  justify: 'both',
};

function alignmentOfJc(jc: string, rtl: boolean): StyleAlignment | undefined {
  switch (jc) {
    case 'left':
    case 'start':
      return rtl ? 'right' : 'left';
    case 'right':
    case 'end':
      return rtl ? 'left' : 'right';
    case 'center':
      return 'center';
    case 'both':
    case 'distribute':
      return 'justify';
    default:
      return undefined;
  }
}

/* ------------------------------------------------------------------ */
/* סורק רדוד                                                           */
/* ------------------------------------------------------------------ */

interface Node {
  /** השם המקומי, בלי קידומת. */
  name: string;
  start: number;
  end: number;
  /** המאפיינים של תג הפתיחה, כמחרוזת גולמית. */
  attrs: string;
  /** תחילת התוכן הפנימי וסופו. בתג סוגר-עצמי שניהם `end`. */
  innerStart: number;
  innerEnd: number;
}

/**
 * הילדים הישירים בטווח `[from, to)`. הערות, CDATA והוראות עיבוד נבלעות — אותו
 * כלל של הסורק ב-docx-parts.ts, ומאותו טעם.
 */
function childrenOf(xml: string, from: number, to: number): Node[] {
  const out: Node[] = [];
  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  token.lastIndex = from;
  let depth = 0;
  let open: { name: string; start: number; attrs: string; innerStart: number } | null = null;
  for (let match = token.exec(xml); match && match.index < to; match = token.exec(xml)) {
    const closer = skippedCloser(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      if (end < 0) break;
      token.lastIndex = end + closer.length;
      continue;
    }
    const closing = match[1] === '/';
    const attrs = match[4] ?? '';
    const selfClosing = !closing && attrs.endsWith('/');
    if (closing) {
      depth -= 1;
      if (depth === 0 && open) {
        out.push({ ...open, end: token.lastIndex, innerEnd: match.index });
        open = null;
      }
      if (depth < 0) break;
      continue;
    }
    if (depth === 0) {
      if (selfClosing) {
        const end = token.lastIndex;
        out.push({ name: match[3], start: match.index, end, attrs: attrs.slice(0, -1), innerStart: end, innerEnd: end });
        continue;
      }
      open = { name: match[3], start: match.index, attrs, innerStart: token.lastIndex };
    }
    if (!selfClosing) depth += 1;
  }
  return out;
}

/** אלמנט כמחרוזת, מפורק לתג פתיחה ולתוכן. */
interface Element {
  prefix: string;
  name: string;
  attrs: string;
  /** ‏`null` = סוגר-עצמי. */
  inner: string | null;
}

function parseElement(text: string): Element | null {
  const nodes = childrenOf(text, 0, text.length);
  const node = nodes[0];
  if (!node) return null;
  const head = new RegExp(TOKEN_SOURCE.source).exec(text.slice(node.start, node.end));
  return {
    prefix: head?.[2] ?? 'w',
    name: node.name,
    attrs: node.attrs,
    inner: node.innerStart === node.end ? null : text.slice(node.innerStart, node.innerEnd),
  };
}

function serialize(element: Element): string {
  const tag = `${element.prefix}:${element.name}`;
  const attrs = element.attrs.replace(/\s+$/, '');
  return element.inner === null || element.inner === ''
    ? `<${tag}${attrs}/>`
    : `<${tag}${attrs}>${element.inner}</${tag}>`;
}

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function unescapeXml(text: string): string {
  return text
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** קובעת מאפיין (`null` = מסירה). השם בלי קידומת; הכתיבה עם הקידומת של האלמנט. */
function withAttribute(element: Element, name: string, value: string | null): Element {
  const pattern = new RegExp(`\\s(?:[\\w.-]+:)?${escapeRegExp(name)}\\s*=\\s*(?:"[^"]*"|'[^']*')`);
  let attrs = element.attrs.replace(pattern, '');
  if (value !== null) attrs = `${attrs.replace(/\s+$/, '')} ${element.prefix}:${name}="${escapeXml(value)}"`;
  return { ...element, attrs };
}

/** הילד הראשון בשם הזה, כמחרוזת, או `null`. */
function childText(element: Element, name: string): string | null {
  if (element.inner === null) return null;
  const node = childrenOf(element.inner, 0, element.inner.length).find((child) => child.name === name);
  return node ? element.inner.slice(node.start, node.end) : null;
}

/**
 * מחליפה את הילד בשם הזה (`null` = מוחקת), ומכניסה ילד חדש למקומו לפי סדר
 * הסכמה. ילד ששמו אינו ב-`order` נשאר במקומו, ולידו לא נכנס דבר.
 */
function withChild(element: Element, name: string, replacement: string | null, order: readonly string[]): Element {
  const inner = element.inner ?? '';
  const nodes = childrenOf(inner, 0, inner.length);
  const existing = nodes.filter((node) => node.name === name);

  if (existing.length > 0) {
    let out = inner;
    // מהסוף להתחלה, כדי שההיסטים לא יזוזו. הראשון מוחלף, כפילויות נמחקות.
    for (let index = existing.length - 1; index >= 0; index--) {
      const node = existing[index];
      const text = index === 0 && replacement !== null ? replacement : '';
      out = out.slice(0, node.start) + text + out.slice(node.end);
    }
    return { ...element, inner: out };
  }
  if (replacement === null) return element;

  const rank = order.indexOf(name);
  const after = rank < 0 ? undefined : nodes.find((node) => {
    const other = order.indexOf(node.name);
    return other > rank;
  });
  const at = after ? after.start : inner.length;
  return { ...element, inner: inner.slice(0, at) + replacement + inner.slice(at) };
}

/* ------------------------------------------------------------------ */
/* סדר הסכמה                                                           */
/* ------------------------------------------------------------------ */

/** ‏CT_Style. */
const STYLE_ORDER = [
  'name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority', 'semiHidden',
  'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose', 'personalReply', 'rsid',
  'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr',
] as const;

/** ‏CT_PPrBase, ואחריו `rPr`/`sectPr`/`pPrChange` של CT_PPr. */
const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
  'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
  'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
  'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
] as const;

/** ‏EG_RPrBase, ואחריו `rPrChange`. */
const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
  'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
  'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
  'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange',
] as const;

/* ------------------------------------------------------------------ */
/* הגיליון                                                             */
/* ------------------------------------------------------------------ */

interface Sheet {
  xml: string;
  prefix: string;
  /** תחום התוכן של `w:styles`. */
  innerStart: number;
  innerEnd: number;
  styles: Node[];
}

function readSheet(xml: string): Sheet | null {
  const prefix = wordPrefix(xml);
  if (!prefix) return null;
  const root = childrenOf(xml, 0, xml.length).find((node) => node.name === 'styles');
  if (!root) return null;
  const styles = childrenOf(xml, root.innerStart, root.innerEnd).filter((node) => node.name === 'style');
  return { xml, prefix, innerStart: root.innerStart, innerEnd: root.innerEnd, styles };
}

function styleIdOf(node: Node): string {
  return unescapeXml(attribute(node.attrs, 'styleId') ?? '');
}

function childValue(element: Element, name: string): string | null {
  const text = childText(element, name);
  if (text === null) return null;
  const child = parseElement(text);
  const value = child ? valueOf(child.attrs) : null;
  return value === null ? null : unescapeXml(value);
}

function infoOf(sheet: Sheet, node: Node): StyleInfo {
  const element = parseElement(sheet.xml.slice(node.start, node.end));
  const flag = (name: string): boolean => {
    const value = attribute(node.attrs, name);
    return value !== null && isOn(` w:val="${value}"`);
  };
  const id = styleIdOf(node);
  return {
    id,
    name: (element && childValue(element, 'name')) ?? id,
    type: attribute(node.attrs, 'type') ?? 'paragraph',
    basedOn: element ? childValue(element, 'basedOn') : null,
    next: element ? childValue(element, 'next') : null,
    link: element ? childValue(element, 'link') : null,
    custom: flag('customStyle'),
    isDefault: flag('default'),
    hidden: element !== null && (childText(element, 'hidden') !== null || childText(element, 'semiHidden') !== null),
  };
}

/** כל הסגנונות שבגיליון, בסדר שלהם. גיליון שאינו נקרא — רשימה ריקה. */
export function listStyles(stylesXml: string | null | undefined): StyleInfo[] {
  const sheet = stylesXml ? readSheet(stylesXml) : null;
  return sheet ? sheet.styles.map((node) => infoOf(sheet, node)) : [];
}

function findStyle(sheet: Sheet, id: string): Node | undefined {
  return sheet.styles.find((node) => styleIdOf(node) === id);
}

function replaceRange(sheet: Sheet, start: number, end: number, text: string): string {
  return sheet.xml.slice(0, start) + text + sheet.xml.slice(end);
}

/* ------------------------------------------------------------------ */
/* קריאת העיצוב                                                        */
/* ------------------------------------------------------------------ */

/** ההצהרות של רמה אחת — `rPr` ו-`pPr` — כ-`StyleFormat` חלקי. */
function formatOfLevel(rPr: Element | null, pPr: Element | null, rtl = false, complex = true): StyleFormat {
  const format: StyleFormat = {};
  if (rPr) {
    const fonts = childText(rPr, 'rFonts');
    if (fonts) {
      const attrs = parseElement(fonts)?.attrs ?? '';
      // ערכת נושא גוברת על שם באותה רמה — כך Word פותר `rFonts`. השם שלה אינו
      // ב-styles.xml, ולכן `''`: „גופן ערכת הנושא”, ולא השם של רמה נמוכה יותר.
      const themed = ['cstheme', 'asciiTheme', 'hAnsiTheme'].some((name) => attribute(attrs, name) !== null);
      const name = attribute(attrs, 'cs') ?? attribute(attrs, 'ascii') ?? attribute(attrs, 'hAnsi');
      if (themed) format.fontFamily = '';
      else if (name) format.fontFamily = unescapeXml(name);
    }
    const size = childValue(rPr, 'szCs') ?? childValue(rPr, 'sz');
    if (size !== null && Number.isFinite(Number(size))) format.fontSizePt = Number(size) / 2;
    const toggle = (cs: string, latin: string): boolean | undefined => {
      const text = childText(rPr, complex ? cs : latin);
      return text === null ? undefined : isOn(parseElement(text)?.attrs ?? '');
    };
    const bold = toggle('bCs', 'b');
    if (bold !== undefined) format.bold = bold;
    const italic = toggle('iCs', 'i');
    if (italic !== undefined) format.italic = italic;
    const underline = childValue(rPr, 'u');
    if (underline !== null) format.underline = underline !== 'none';
    const color = childValue(rPr, 'color');
    if (color !== null && /^[0-9a-f]{6}$/i.test(color)) format.color = color.toUpperCase();
  }
  if (pPr) {
    const jc = childValue(pPr, 'jc');
    const alignment = jc === null ? undefined : alignmentOfJc(jc, rtl);
    if (alignment) format.alignment = alignment;
    const spacing = childText(pPr, 'spacing');
    if (spacing) {
      const attrs = parseElement(spacing)?.attrs ?? '';
      const before = attribute(attrs, 'before');
      const after = attribute(attrs, 'after');
      const line = attribute(attrs, 'line');
      const rule = attribute(attrs, 'lineRule') ?? 'auto';
      if (before !== null) format.spaceBeforePt = Number(before) / TWIPS_PER_PT;
      if (after !== null) format.spaceAfterPt = Number(after) / TWIPS_PER_PT;
      if (line !== null && rule === 'auto') format.lineSpacing = Math.round((Number(line) / LINE_UNITS) * 100) / 100;
    }
    const outline = childValue(pPr, 'outlineLvl');
    if (outline !== null && /^\d$/.test(outline)) format.outlineLevel = Number(outline);
    const ind = childText(pPr, 'ind');
    if (ind) {
      const attrs = parseElement(ind)?.attrs ?? '';
      const start = attribute(attrs, 'start') ?? attribute(attrs, 'left');
      const end = attribute(attrs, 'end') ?? attribute(attrs, 'right');
      const firstLine = attribute(attrs, 'firstLine');
      const hanging = attribute(attrs, 'hanging');
      if (start !== null) format.indentStartCm = twipsToCm(Number(start));
      if (end !== null) format.indentEndCm = twipsToCm(Number(end));
      if (hanging !== null) format.firstLineCm = -twipsToCm(Number(hanging));
      else if (firstLine !== null) format.firstLineCm = twipsToCm(Number(firstLine));
    }
  }
  return format;
}

function levelsOf(element: Element | null): { rPr: Element | null; pPr: Element | null } {
  if (!element) return { rPr: null, pPr: null };
  const rPr = childText(element, 'rPr');
  const pPr = childText(element, 'pPr');
  return { rPr: rPr ? parseElement(rPr) : null, pPr: pPr ? parseElement(pPr) : null };
}

/** Effective bidi follows docDefaults and the basedOn chain, including explicit false. */
function styleIsRtl(sheet: Sheet, id: string | null): boolean {
  const chain: Node[] = [];
  const seen = new Set<string>();
  for (let current = id; current && !seen.has(current); ) {
    seen.add(current);
    const node = findStyle(sheet, current);
    if (!node) break;
    chain.unshift(node);
    current = infoOf(sheet, node).basedOn;
  }
  const root = parseElement(sheet.xml);
  const defaults = root && childText(root, 'docDefaults');
  const wrapper = defaults && childText(parseElement(defaults)!, 'pPrDefault');
  const defaultsPr = wrapper && childText(parseElement(wrapper)!, 'pPr');
  let rtl = false;
  for (const pPr of [defaultsPr && parseElement(defaultsPr), ...chain.map(node =>
    levelsOf(parseElement(sheet.xml.slice(node.start, node.end))).pPr)]) {
    const bidi = pPr && childText(pPr, 'bidi');
    if (bidi) rtl = isOn(parseElement(bidi)?.attrs ?? '');
  }
  return rtl;
}

/** Only changed style definitions are imported; defaults and unrelated styles stay live. */
export function changedStyleDefinitions(before: string, after: string): string | null {
  const oldSheet = readSheet(before);
  const sheet = readSheet(after);
  if (!oldSheet || !sheet) return null;
  const originals = new Map(oldSheet.styles.map(node => [styleIdOf(node), before.slice(node.start, node.end)]));
  const changed = sheet.styles.filter(node => originals.get(styleIdOf(node)) !== after.slice(node.start, node.end));
  const root = parseElement(after);
  if (!root) return null;
  return serialize({ ...root, inner: changed.map(node => after.slice(node.start, node.end)).join('') });
}

/** מה שהסגנון עצמו מצהיר, בלי ירושה. */
export function readOwnFormat(stylesXml: string, id: string): StyleFormat | null {
  const sheet = readSheet(stylesXml);
  const node = sheet ? findStyle(sheet, id) : undefined;
  if (!sheet || !node) return null;
  const { rPr, pPr } = levelsOf(parseElement(sheet.xml.slice(node.start, node.end)));
  return formatOfLevel(rPr, pPr, styleIsRtl(sheet, id));
}

/**
 * העיצוב האפקטיבי של הסגנון: ברירות המחדל של המסמך, ומעליהן שרשרת `basedOn`
 * מהשורש אל הסגנון. זה מה שהדיאלוג מציג — „גודל: 16” של כותרת שאינה מצהירה
 * גודל ויורשת אותו מ„רגיל” הוא מה שהמשתמש רואה בעמוד.
 */
function resolveScriptFormat(stylesXml: string, id: string, complex: boolean): StyleFormat {
  const sheet = readSheet(stylesXml);
  if (!sheet) return {};

  const chain: Node[] = [];
  const seen = new Set<string>();
  for (let current: string | null = id; current && !seen.has(current); ) {
    seen.add(current);
    const node = findStyle(sheet, current);
    if (!node) break;
    chain.unshift(node);
    current = infoOf(sheet, node).basedOn;
  }

  let format: StyleFormat = {};
  const rtl = styleIsRtl(sheet, id);
  const defaults = childrenOf(sheet.xml, sheet.innerStart, sheet.innerEnd).find((node) => node.name === 'docDefaults');
  if (defaults) {
    const parts = childrenOf(sheet.xml, defaults.innerStart, defaults.innerEnd);
    const inner = (name: string, child: string): Element | null => {
      const wrapper = parts.find((node) => node.name === name);
      if (!wrapper) return null;
      const element = parseElement(sheet.xml.slice(wrapper.start, wrapper.end));
      const text = element ? childText(element, child) : null;
      return text ? parseElement(text) : null;
    };
    format = formatOfLevel(inner('rPrDefault', 'rPr'), inner('pPrDefault', 'pPr'), rtl, complex);
  }
  for (const node of chain) {
    const { rPr, pPr } = levelsOf(parseElement(sheet.xml.slice(node.start, node.end)));
    const own = formatOfLevel(rPr, pPr, rtl, complex);
    // In named styles, true toggles inherited bold/italic; false leaves them unchanged (§17.7.3).
    for (const key of ['bold', 'italic'] as const) {
      if (own[key] !== undefined) own[key] = own[key] ? !format[key] : (format[key] ?? false);
    }
    format = { ...format, ...own };
  }
  return format;
}

export function resolveFormat(stylesXml: string, id: string): StyleFormat {
  return resolveScriptFormat(stylesXml, id, true);
}

interface InheritedRuns { latin: StyleFormat; complex: StyleFormat }
function inheritedRuns(xml: string, basedOn: string | null): InheritedRuns {
  return { latin: resolveScriptFormat(xml, basedOn ?? '', false), complex: resolveScriptFormat(xml, basedOn ?? '', true) };
}

/** The pinned browser renderer uses last-write-wins for named-style toggles.
 * Refuse NEW Word/renderer disagreements before importing any definitions.
 * Resolve both graphs with memoization: work scales with stylesheet size, not body size.
 */
export function liveStyleTogglesSupported(
  beforeXml: string,
  afterXml: string,
  edit: StyleEdit,
  editedId: string,
  /** Atomic template edits: validate all explicit requests in the same final graph. */
  additionalRequests?: Readonly<Record<string, StyleFormat>>,
): boolean {
  const before = readSheet(beforeXml);
  const after = readSheet(afterXml);
  if (!before || !after) return false;
  type Toggle = 'b' | 'bCs' | 'i' | 'iCs';
  type State = Record<Toggle, { word: boolean; renderer: boolean }>;
  const keys: Toggle[] = ['b', 'bCs', 'i', 'iCs'];
  const raw = (rPr: Element | null, key: Toggle): boolean | undefined => {
    const text = rPr && childText(rPr, key);
    return text ? isOn(parseElement(text)?.attrs ?? '') : undefined;
  };
  const graph = (sheet: Sheet) => {
    const root = parseElement(sheet.xml);
    const defaultsXml = root && childText(root, 'docDefaults');
    const defaults = defaultsXml && parseElement(defaultsXml);
    const wrapperXml = defaults && childText(defaults, 'rPrDefault');
    const wrapper = wrapperXml && parseElement(wrapperXml);
    const rPrXml = wrapper && childText(wrapper, 'rPr');
    const rPr = rPrXml ? parseElement(rPrXml) : null;
    const initial = Object.fromEntries(keys.map(key => {
      const value = raw(rPr, key) ?? false;
      return [key, { word: value, renderer: value }];
    })) as State;
    const definitions = new Map(sheet.styles.map(node => {
      const element = parseElement(sheet.xml.slice(node.start, node.end));
      return [styleIdOf(node), { parent: element ? childValue(element, 'basedOn') : null, rPr: levelsOf(element).rPr }];
    }));
    const cached = new Map<string, State | null>();
    const resolve = (id: string, path = new Set<string>()): State | null => {
      if (cached.has(id)) return cached.get(id)!;
      if (path.has(id)) return null;
      const own = definitions.get(id);
      if (!own) return initial;
      const next = new Set(path).add(id);
      const parent = own.parent ? resolve(own.parent, next) : initial;
      if (!parent) { cached.set(id, null); return null; }
      const state = Object.fromEntries(keys.map(key => {
        const value = raw(own.rPr, key);
        return [key, {
          word: value === true ? !parent[key].word : parent[key].word,
          renderer: value === undefined ? parent[key].renderer : value,
        }];
      })) as State;
      cached.set(id, state);
      return state;
    };
    return { definitions, resolve };
  };
  const old = graph(before);
  const current = graph(after);
  const originals = new Map(before.styles.map(node => [styleIdOf(node), beforeXml.slice(node.start, node.end)]));
  const affected = new Set(after.styles.filter(node => originals.get(styleIdOf(node)) !== afterXml.slice(node.start, node.end)).map(styleIdOf));
  // A base edit can introduce a disagreement in a descendant that wasn't written.
  const children = new Map<string, string[]>();
  for (const [id, definition] of current.definitions) {
    if (definition.parent) {
      const siblings = children.get(definition.parent) ?? [];
      if (!children.has(definition.parent)) children.set(definition.parent, siblings);
      siblings.push(id);
    }
  }
  const pending = [...affected];
  while (pending.length) for (const id of children.get(pending.pop()!) ?? []) {
    if (!affected.has(id)) { affected.add(id); pending.push(id); }
  }
  for (const id of affected) {
    const wanted = current.resolve(id);
    if (!wanted) return false;
    const previous = old.definitions.has(id) ? old.resolve(id) : null;
    for (const key of keys) {
      if (wanted[key].word === wanted[key].renderer) continue;
      const format = id === editedId && (edit.kind === 'create' || edit.kind === 'modify')
        ? edit.format : additionalRequests?.[id];
      const requested = format && has(format, key.startsWith('b') ? 'bold' : 'italic');
      if (requested || !previous || previous[key].word !== wanted[key].word || previous[key].renderer !== wanted[key].renderer) return false;
    }
  }
  return true;
}

/* ------------------------------------------------------------------ */
/* כתיבת העיצוב                                                        */
/* ------------------------------------------------------------------ */

function emptyElement(prefix: string, name: string): Element {
  return { prefix, name, attrs: '', inner: null };
}

function leaf(prefix: string, name: string, value?: string): string {
  return value === undefined ? `<${prefix}:${name}/>` : `<${prefix}:${name} ${prefix}:val="${escapeXml(value)}"/>`;
}

function has(format: StyleFormat, key: keyof StyleFormat): boolean {
  return Object.prototype.hasOwnProperty.call(format, key) && format[key] !== undefined;
}

/** מחילה את שדות הריצה על `rPr`. */
function patchRunProperties(rPr: Element, format: StyleFormat, inherited: InheritedRuns): Element {
  const p = rPr.prefix;
  let out = rPr;
  const set = (name: string, text: string | null): void => {
    out = withChild(out, name, text, RPR_ORDER);
  };

  if (has(format, 'fontFamily')) {
    if (format.fontFamily === null) {
      set('rFonts', null);
    } else {
      const existing = childText(out, 'rFonts');
      let fonts = (existing && parseElement(existing)) || emptyElement(p, 'rFonts');
      // ערכת נושא גוברת על שם גופן באותה רמה — ובלי ההסרה הבחירה לא הייתה חלה.
      for (const theme of ['asciiTheme', 'hAnsiTheme', 'cstheme', 'eastAsiaTheme']) {
        fonts = withAttribute(fonts, theme, null);
      }
      for (const slot of ['ascii', 'hAnsi', 'cs']) fonts = withAttribute(fonts, slot, format.fontFamily as string);
      set('rFonts', serialize(fonts));
    }
  }

  if (has(format, 'fontSizePt')) {
    const half = format.fontSizePt === null ? null : String(round((format.fontSizePt as number) * 2));
    set('sz', half === null ? null : leaf(p, 'sz', half));
    set('szCs', half === null ? null : leaf(p, 'szCs', half));
  }

  const toggles: [keyof StyleFormat, string, string][] = [
    ['bold', 'b', 'bCs'],
    ['italic', 'i', 'iCs'],
  ];
  for (const [key, latin, complex] of toggles) {
    if (!has(format, key)) continue;
    const value = format[key];
    const text = (name: string, baseline: StyleFormat): string | null =>
      value === null || (value === true && baseline[key] === true) ? null
        : value !== (baseline[key] ?? false) ? leaf(p, name) : leaf(p, name, '0');
    set(latin, text(latin, inherited.latin));
    set(complex, text(complex, inherited.complex));
  }

  if (has(format, 'underline')) {
    set('u', format.underline === null ? null : leaf(p, 'u', format.underline ? 'single' : 'none'));
  }

  if (has(format, 'color')) {
    set('color', format.color === null ? null : leaf(p, 'color', (format.color as string).toUpperCase()));
  }

  return out;
}

/** מחילה את שדות הפסקה על `pPr`. */
function patchParagraphProperties(pPr: Element, format: StyleFormat, rtl: boolean): Element {
  const p = pPr.prefix;
  let out = pPr;

  if (has(format, 'alignment')) {
    out = withChild(
      out,
      'jc',
      format.alignment === null ? null : leaf(p, 'jc', rtl ? JC_FOR_ALIGNMENT[format.alignment as StyleAlignment] : ({ right: 'right', left: 'left', center: 'center', justify: 'both' } as const)[format.alignment as StyleAlignment]),
      PPR_ORDER,
    );
  }

  const spacingKeys = has(format, 'spaceBeforePt') || has(format, 'spaceAfterPt') || has(format, 'lineSpacing');
  if (spacingKeys) {
    const existing = childText(out, 'spacing');
    let spacing = (existing && parseElement(existing)) || emptyElement(p, 'spacing');
    if (has(format, 'spaceBeforePt')) {
      const value = format.spaceBeforePt;
      spacing = withAttribute(spacing, 'before', value === null ? null : String(round((value as number) * TWIPS_PER_PT)));
      spacing = withAttribute(spacing, 'beforeAutospacing', null);
    }
    if (has(format, 'spaceAfterPt')) {
      const value = format.spaceAfterPt;
      spacing = withAttribute(spacing, 'after', value === null ? null : String(round((value as number) * TWIPS_PER_PT)));
      spacing = withAttribute(spacing, 'afterAutospacing', null);
    }
    if (has(format, 'lineSpacing')) {
      const value = format.lineSpacing;
      spacing = withAttribute(spacing, 'line', value === null ? null : String(round((value as number) * LINE_UNITS)));
      spacing = withAttribute(spacing, 'lineRule', value === null ? null : 'auto');
    }
    const empty = spacing.attrs.trim() === '';
    out = withChild(out, 'spacing', empty ? null : serialize(spacing), PPR_ORDER);
  }

  if (has(format, 'outlineLevel')) {
    const level = format.outlineLevel;
    const valid = typeof level === 'number' && Number.isInteger(level) && level >= 0 && level <= BODY_TEXT_LEVEL;
    out = withChild(out, 'outlineLvl', valid ? leaf(p, 'outlineLvl', String(level)) : null, PPR_ORDER);
  }

  const indentKeys = has(format, 'indentStartCm') || has(format, 'indentEndCm') || has(format, 'firstLineCm');
  if (indentKeys) {
    const existing = childText(out, 'ind');
    let ind = (existing && parseElement(existing)) || emptyElement(p, 'ind');
    const twips = (cm: number): string => String(round(cm * TWIPS_PER_CM));
    // `start`/`end` — מה שהעורך עצמו כותב בפסקה (ראו ParagraphDialog.vue).
    // שני השמות של אותו צד מוסרים יחד: `left` ישן לצד `start` חדש הוא שתי
    // תשובות סותרות לאותה שאלה, וכל קורא בוחר אחרת.
    if (has(format, 'indentStartCm')) {
      ind = withAttribute(withAttribute(ind, 'start', null), 'left', null);
      if (format.indentStartCm !== null) ind = withAttribute(ind, 'start', twips(format.indentStartCm as number));
    }
    if (has(format, 'indentEndCm')) {
      ind = withAttribute(withAttribute(ind, 'end', null), 'right', null);
      if (format.indentEndCm !== null) ind = withAttribute(ind, 'end', twips(format.indentEndCm as number));
    }
    if (has(format, 'firstLineCm')) {
      ind = withAttribute(withAttribute(ind, 'firstLine', null), 'hanging', null);
      const value = format.firstLineCm;
      if (value !== null && value !== undefined && value !== 0) {
        ind = value > 0 ? withAttribute(ind, 'firstLine', twips(value)) : withAttribute(ind, 'hanging', twips(-value));
      } else if (value === 0) {
        ind = withAttribute(ind, 'firstLine', '0');
      }
    }
    const empty = ind.attrs.trim() === '';
    out = withChild(out, 'ind', empty ? null : serialize(ind), PPR_ORDER);
  }

  return out;
}

function hasParagraphFields(format: StyleFormat): boolean {
  return (['alignment', 'spaceBeforePt', 'spaceAfterPt', 'lineSpacing', 'indentStartCm', 'indentEndCm', 'firstLineCm', 'outlineLevel'] as const)
    .some((key) => has(format, key));
}

function hasRunFields(format: StyleFormat): boolean {
  return (['fontFamily', 'fontSizePt', 'bold', 'italic', 'underline', 'color'] as const).some((key) => has(format, key));
}

/** מחילה עיצוב על אלמנט `w:style` שלם. `rPr`/`pPr` ריקים נמחקים. */
function patchStyleElement(style: Element, format: StyleFormat, kind: string, rtl = false, inherited: InheritedRuns = { latin: {}, complex: {} }): Element {
  let out = style;
  if (hasRunFields(format)) {
    const text = childText(out, 'rPr');
    const rPr = patchRunProperties((text && parseElement(text)) || emptyElement(out.prefix, 'rPr'), format, inherited);
    out = withChild(out, 'rPr', rPr.inner ? serialize(rPr) : null, STYLE_ORDER);
  }
  if (kind === 'paragraph' && hasParagraphFields(format)) {
    const text = childText(out, 'pPr');
    const pPr = patchParagraphProperties((text && parseElement(text)) || emptyElement(out.prefix, 'pPr'), format, rtl);
    out = withChild(out, 'pPr', pPr.inner ? serialize(pPr) : null, STYLE_ORDER);
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* הפעולות                                                             */
/* ------------------------------------------------------------------ */

/**
 * מזהה לסגנון חדש. Word גוזר אותו מהשם ומשמיט את מה שאינו ASCII — בשם עברי
 * לגמרי הוא מגיע ל-`a`, `a0`, `a1`. כאן „Style” במקום, כדי שהמזהה יהיה
 * קריא בקובץ. ייחודי בלי תלות באותיות גדולות: Word משווה מזהים כך.
 */
export function newStyleId(stylesXml: string, name: string): string {
  const taken = new Set(listStyles(stylesXml).map((style) => style.id.toLowerCase()));
  const base = name.replace(/[^A-Za-z0-9]/g, '') || 'Style';
  if (!taken.has(base.toLowerCase())) return base;
  for (let index = 1; ; index++) {
    const candidate = `${base}${index}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}

/** האם השם תפוס — בלי תלות באותיות גדולות, כמו ב-Word. */
export function isStyleNameTaken(stylesXml: string, name: string, exceptId?: string): boolean {
  const key = name.trim().toLowerCase();
  return listStyles(stylesXml).some((style) => style.id !== exceptId && style.name.trim().toLowerCase() === key);
}

function fail(message: string): StyleEditResult {
  return { ok: false, message };
}

function setReference(style: Element, name: 'basedOn' | 'next', value: string | null | undefined): Element {
  if (value === undefined) return style;
  return withChild(style, name, value === null ? null : leaf(style.prefix, name, value), STYLE_ORDER);
}

/** מחילה עריכה על `styles.xml`. אינה זורקת: קלט שאינו נקרא הוא `ok: false`. */
export function editStyles(stylesXml: string, edit: StyleEdit): StyleEditResult {
  const sheet = readSheet(stylesXml);
  if (!sheet) return fail('גיליון הסגנונות של המסמך אינו קריא');
  const p = sheet.prefix;

  switch (edit.kind) {
    case 'create': {
      const name = edit.name.trim();
      if (name === '') return fail('לסגנון חייב להיות שם');
      if (isStyleNameTaken(stylesXml, name)) return fail(`כבר יש סגנון בשם „${name}”`);
      if (edit.basedOn && !findStyle(sheet, edit.basedOn)) return fail('סגנון הבסיס אינו במסמך');

      const id = newStyleId(stylesXml, name);
      let style: Element = {
        prefix: p,
        name: 'style',
        attrs: ` ${p}:type="${edit.type}" ${p}:customStyle="1" ${p}:styleId="${escapeXml(id)}"`,
        inner: leaf(p, 'name', name) + leaf(p, 'qFormat'),
      };
      style = setReference(style, 'basedOn', edit.basedOn);
      if (edit.type === 'paragraph') style = setReference(style, 'next', edit.next ?? id);
      style = patchStyleElement(style, edit.format, edit.type, styleIsRtl(sheet, edit.basedOn), inheritedRuns(stylesXml, edit.basedOn));
      const xml = replaceRange(sheet, sheet.innerEnd, sheet.innerEnd, serialize(style));
      return { ok: true, xml, id, removed: [] };
    }

    case 'modify': {
      const node = findStyle(sheet, edit.id);
      if (!node) return fail('הסגנון אינו במסמך');
      const info = infoOf(sheet, node);
      if (edit.basedOn && edit.basedOn !== info.basedOn) {
        if (!findStyle(sheet, edit.basedOn)) return fail('סגנון הבסיס אינו במסמך');
        // מעגל בשרשרת: Word מסרב לו, ומנוע שמתעלם ממנו נתקע בלולאה.
        const visited = new Set<string>();
        for (let current: string | null = edit.basedOn; current; ) {
          if (visited.has(current)) return fail('שרשרת סגנונות הבסיס מכילה מעגל');
          visited.add(current);
          if (current === edit.id) return fail('סגנון אינו יכול להיות מבוסס על עצמו');
          const parent = findStyle(sheet, current);
          current = parent ? infoOf(sheet, parent).basedOn : null;
        }
      }

      let style = parseElement(sheet.xml.slice(node.start, node.end));
      if (!style) return fail('הסגנון אינו קריא');
      style = setReference(style, 'basedOn', edit.basedOn);
      if (info.type === 'paragraph') style = setReference(style, 'next', edit.next);
      style = patchStyleElement(style, edit.format, info.type, styleIsRtl(readSheet(replaceRange(sheet, node.start, node.end, serialize(style)))!, edit.id), inheritedRuns(stylesXml, edit.basedOn === undefined ? info.basedOn : edit.basedOn));
      let xml = replaceRange(sheet, node.start, node.end, serialize(style));

      // הסגנון המקושר (`Heading1Char` של `Heading1`) נושא את אותם מאפייני
      // ריצה, וזה מה ש-Word מעדכן יחד: טקסט שסומן בו נראה אחרת מהכותרת.
      const link = info.link;
      if (link && hasRunFields(edit.format)) {
        const linkedSheet = readSheet(xml);
        const linkedNode = linkedSheet ? findStyle(linkedSheet, link) : undefined;
        const linked = linkedNode && linkedSheet ? parseElement(xml.slice(linkedNode.start, linkedNode.end)) : null;
        if (linkedSheet && linkedNode && linked) {
          const runOnly: StyleFormat = {};
          for (const key of ['fontFamily', 'fontSizePt', 'bold', 'italic', 'underline', 'color'] as const) {
            if (has(edit.format, key)) (runOnly as Record<string, unknown>)[key] = edit.format[key];
          }
          xml = replaceRange(linkedSheet, linkedNode.start, linkedNode.end, serialize(patchStyleElement(linked, runOnly, 'character', false, inheritedRuns(xml, infoOf(linkedSheet, linkedNode).basedOn))));
        }
      }
      return { ok: true, xml, id: edit.id, removed: [] };
    }

    case 'rename': {
      const node = findStyle(sheet, edit.id);
      if (!node) return fail('הסגנון אינו במסמך');
      if (!infoOf(sheet, node).custom) return fail('אי אפשר לשנות שם של סגנון מובנה');
      const name = edit.name.trim();
      if (name === '') return fail('לסגנון חייב להיות שם');
      if (isStyleNameTaken(stylesXml, name, edit.id)) return fail(`כבר יש סגנון בשם „${name}”`);
      const style = parseElement(sheet.xml.slice(node.start, node.end));
      if (!style) return fail('הסגנון אינו קריא');
      const xml = replaceRange(sheet, node.start, node.end, serialize(withChild(style, 'name', leaf(p, 'name', name), STYLE_ORDER)));
      return { ok: true, xml, id: edit.id, removed: [] };
    }

    case 'delete': {
      const node = findStyle(sheet, edit.id);
      if (!node) return fail('הסגנון אינו במסמך');
      const info = infoOf(sheet, node);
      if (!info.custom || info.isDefault) return fail('אי אפשר למחוק סגנון מובנה');

      const removed = [edit.id];
      if (info.link) {
        const linked = findStyle(sheet, info.link);
        if (linked && infoOf(sheet, linked).custom) removed.push(info.link);
      }
      return { ok: true, xml: removeStyles(stylesXml, removed), id: edit.id, removed };
    }
  }
}

/**
 * מוחקת סגנונות מהגיליון, ומתקנת את מי שהצביע עליהם: `basedOn` עובר לבסיס של
 * הסגנון שנמחק (כך העיצוב שירש לא משתנה), `next` ו-`link` נמחקים.
 */
function removeStyles(stylesXml: string, ids: readonly string[]): string {
  let xml = stylesXml;
  const doomed = new Set(ids);
  const before = readSheet(xml);
  const parentOf = new Map<string, string | null>();
  for (const id of ids) {
    const node = before ? findStyle(before, id) : undefined;
    parentOf.set(id, before && node ? infoOf(before, node).basedOn : null);
  }
  const inheritedBase = (id: string): string | null => {
    let current: string | null = id;
    const seen = new Set<string>();
    while (current && doomed.has(current) && !seen.has(current)) {
      seen.add(current);
      current = parentOf.get(current) ?? null;
    }
    return current;
  };

  {
    const sheet = readSheet(xml);
    if (!sheet) return xml;
    // מהסוף, כדי שההיסטים של מה שלפני לא יזוזו.
    const edits: { start: number; end: number; text: string }[] = [];
    for (const node of sheet.styles) {
      const id = styleIdOf(node);
      if (doomed.has(id)) {
        edits.push({ start: node.start, end: node.end, text: '' });
        continue;
      }
      const info = infoOf(sheet, node);
      let style = parseElement(sheet.xml.slice(node.start, node.end));
      if (!style) continue;
      let changed = false;
      if (info.basedOn && doomed.has(info.basedOn)) {
        style = setReference(style, 'basedOn', inheritedBase(info.basedOn));
        changed = true;
      }
      for (const name of ['next', 'link'] as const) {
        const value = name === 'next' ? info.next : info.link;
        if (value && doomed.has(value)) {
          style = withChild(style, name, null, STYLE_ORDER);
          changed = true;
        }
      }
      if (changed) edits.push({ start: node.start, end: node.end, text: serialize(style) });
    }
    for (const edit of edits.sort((a, b) => b.start - a.start)) {
      xml = xml.slice(0, edit.start) + edit.text + xml.slice(edit.end);
    }
    return xml;
  }
}

/**
 * מסירה מחלק תוכן (`document.xml`, כותרות, הערות שוליים) את ההפניות לסגנונות
 * שנמחקו. פסקה שאיבדה את `pStyle` שלה חוזרת ל„רגיל”, כמו ב-Word. מחזירה
 * `null` כשלא היה מה להסיר.
 */
export function stripStyleReferences(partXml: string, ids: readonly string[]): string | null {
  if (ids.length === 0) return null;
  const prefix = wordPrefix(partXml);
  if (!prefix) return null;
  const values = ids.map((id) => escapeRegExp(escapeXml(id))).join('|');
  const pattern = new RegExp(
    `<${escapeRegExp(prefix)}:(?:pStyle|rStyle)\\s+${escapeRegExp(prefix)}:val\\s*=\\s*(?:"(?:${values})"|'(?:${values})')\\s*/>`,
    'g',
  );
  const out = partXml.replace(pattern, '');
  return out === partXml ? null : out;
}

/** ‏CT_DocDefaults. */
const DOC_DEFAULTS_ORDER = ['rPrDefault', 'pPrDefault'] as const;

/**
 * גופן וגודל ברירת המחדל של המסמך כולו — `w:docDefaults/w:rPrDefault/w:rPr`,
 * המקום שבו Word כותב „הגדר כברירת מחדל” של דיאלוג הגופן. מקבלת רק את שני
 * השדות האלה; השאר מתעלמים. שני צדי הכתב נכתבים יחד (`sz`+`szCs`,
 * `ascii`/`hAnsi`/`cs`), כמו בסגנון. מחזירה `null` כשהגיליון אינו נקרא.
 */
export function editRunDefaults(stylesXml: string, format: Pick<StyleFormat, 'fontFamily' | 'fontSizePt'>): string | null {
  const sheet = readSheet(stylesXml);
  if (!sheet) return null;
  const p = sheet.prefix;
  const run: StyleFormat = {};
  if (has(format, 'fontFamily')) run.fontFamily = format.fontFamily;
  if (has(format, 'fontSizePt')) run.fontSizePt = format.fontSizePt;
  if (Object.keys(run).length === 0) return stylesXml;

  const node = childrenOf(sheet.xml, sheet.innerStart, sheet.innerEnd).find((child) => child.name === 'docDefaults');
  let defaults = node ? parseElement(sheet.xml.slice(node.start, node.end)) : emptyElement(p, 'docDefaults');
  if (!defaults) return null;
  const rPrDefaultText = childText(defaults, 'rPrDefault');
  let rPrDefault = (rPrDefaultText && parseElement(rPrDefaultText)) || emptyElement(p, 'rPrDefault');
  const rPrText = childText(rPrDefault, 'rPr');
  const rPr = patchRunProperties((rPrText && parseElement(rPrText)) || emptyElement(p, 'rPr'), run, { latin: {}, complex: {} });
  rPrDefault = withChild(rPrDefault, 'rPr', rPr.inner ? serialize(rPr) : null, ['rPr']);
  defaults = withChild(defaults, 'rPrDefault', serialize(rPrDefault), DOC_DEFAULTS_ORDER);
  // ‏`docDefaults` הוא הילד הראשון של `w:styles` לפי הסכמה.
  return node
    ? replaceRange(sheet, node.start, node.end, serialize(defaults))
    : replaceRange(sheet, sheet.innerStart, sheet.innerStart, serialize(defaults));
}

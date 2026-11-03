/**
 * הוולידטורים והקטלוג **האמיתיים** של superdoc@2.8.0, נחלצים מהמקור הארוז
 * ומורצים כמו שהם.
 *
 * למה מודול משותף ולא העתק בכל בדיקה: שני צרכנים שונים צריכים בדיוק את אותה
 * תשובה לשאלה „האם המנוע היה מקבל את ה-payload הזה” —
 * tests/contract/command-payloads.test.ts (חוזה ה-payload), וכפיל האדפטר
 * שבדיקות הקומפוננטות מרכיבות מולו (tests/component/harness.ts). כפיל שמחזיר
 * `true` לכל דבר הוא בדיוק מה שאישר בירוק את חמשת ה-payloads השבורים, ולכן
 * הכפיל חייב לשאול את המנוע — ושאלה אחת פירושה מקור אחד.
 *
 * הוולידטורים אינם exports ציבוריים: הם פנימיים ל-chunk של החבילה. הקוד שרץ
 * כאן הוא הקוד שהחבילה שולחת, ולא העתקה של הלוגיקה — שינוי שלו בגרסה עתידית
 * יפיל את הבדיקות בשתי צורות (החילוץ לא ימצא, או ההתנהגות תשתנה), ובשתיהן זה
 * מה שאנחנו רוצים לדעת.
 */
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

const CHUNKS_DIR = join(process.cwd(), 'node_modules/superdoc/dist/chunks');

/** שם הקובץ נושא hash שמשתנה בכל build של החבילה, ולכן הוא נמצא ולא נכתב. */
function readControllerChunk(): string {
  const file = readdirSync(CHUNKS_DIR).find((name) =>
    /^create-super-doc-ui-.*\.es\.js$/.test(name),
  );
  if (!file) throw new Error('לא נמצא ה-chunk של controller ה-UI ב-superdoc');
  return readFileSync(join(CHUNKS_DIR, file), 'utf8');
}

export const CHUNK = readControllerChunk();

/**
 * הפונקציות ברמת המודול. הן מעוצבות עם tab, ולכן `}` בתחילת שורה הוא סוף
 * הפונקציה — מה שהופך את החילוץ לחד-משמעי בלי מנתח JS.
 */
const TOP_LEVEL_FUNCTIONS = [
  'normalizeFontSizePayload',
  'normalizeZoomPayload',
  'normalizeColorPayload',
  'normalizeAlignmentPayload',
  'normalizeLineHeightPayload',
  'normalizeStyleIdPayload',
  'normalizeDocumentModePayload',
  'normalizeMeasurementUnitPayload',
  'unwrapScalar',
  'buildInlineFormatInput',
  'paragraphTarget',
  // 2.20: פקודות מספור העמודים. הקטלוג מפנה אליהן, ובלעדיהן הוא אינו נטען כלל.
  'readPageNumberKind',
  'readPageNumbersRecord',
  'normalizePageNumbersApplyPayload',
  'normalizePageNumbersRemovePayload',
] as const;

/** קבועי `new Set([...])` ברמת המודול שהפונקציות שנחלצו נשענות עליהם. */
const TOP_LEVEL_SETS = ['PAGE_NUMBER_ALIGNMENTS', 'PAGE_NUMBER_FORMATS'] as const;

/** ה-closures שבתוך ה-controller. אינם ברמת המודול, ולכן נחלצים לפי ההזחה. */
const CONTROLLER_CLOSURES = ['instanceCommandPayloadIsValid', 'buildBlockParagraphInput'] as const;

export type InlineKind = 'toggle' | 'value-string' | 'value-number' | 'clear';
export interface InlineSpec {
  key: string;
  kind: InlineKind;
}
export interface BlockParagraphSpec {
  kind: 'alignment' | 'spacing-line' | 'style' | 'direction';
  fixedValue?: string;
}
export type InlineInput = { target: unknown; value?: unknown; inline?: unknown } | null;

/** ה-descriptor של פקודה בקטלוג, בשדות שמכריעים את גורל ה-payload. */
export interface EngineCommandDescriptor {
  id: string;
  family?: string;
  disposition?: 'routed' | 'deferred' | 'unsupported';
  inline?: InlineSpec;
  blockParagraph?: BlockParagraphSpec;
  instanceRoute?: string;
  docRoute?: string;
  normalizePayload?: (payload: unknown) => unknown;
}

export interface EngineValidators {
  normalizeFontSizePayload(payload: unknown): unknown;
  normalizeZoomPayload(payload: unknown): unknown;
  normalizeColorPayload(payload: unknown): unknown;
  normalizeAlignmentPayload(payload: unknown): unknown;
  normalizeLineHeightPayload(payload: unknown): unknown;
  normalizeStyleIdPayload(payload: unknown): unknown;
  normalizeDocumentModePayload(payload: unknown): unknown;
  normalizeMeasurementUnitPayload(payload: unknown): unknown;
  unwrapScalar(payload: unknown, keys: string[]): unknown;
  buildInlineFormatInput(
    spec: InlineSpec,
    target: unknown,
    payload: unknown,
    active: boolean,
  ): InlineInput;
  buildBlockParagraphInput(
    spec: BlockParagraphSpec,
    blockId: string,
    payload: unknown,
    story?: unknown,
  ): Record<string, unknown> | null;
  instanceCommandPayloadIsValid(descriptor: { id: string }, payload: unknown): boolean;
  COMMAND_CATALOG: EngineCommandDescriptor[];
}

function extractTopLevelFunction(name: string): string {
  const match = CHUNK.match(new RegExp(String.raw`^function ${name}\([\s\S]*?\n\}`, 'm'));
  if (!match) {
    throw new Error(`הפונקציה ${name} לא נמצאה ב-chunk — חוזה ה-payload של superdoc השתנה`);
  }
  return match[0];
}

/**
 * closure בתוך ה-controller: `const x = (...) => { ... };` בהזחת tab אחד.
 * ההזחה היא הסימן שמתחם אותו, ולכן `\n\t};` הוא סופו.
 */
function extractControllerClosure(name: string): string {
  const match = CHUNK.match(new RegExp(String.raw`^\tconst ${name} = \([\s\S]*?\n\t\};`, 'm'));
  if (!match) {
    throw new Error(`ה-closure ${name} לא נמצא ב-chunk — חוזה הפקודות של superdoc השתנה`);
  }
  return match[0].replace(/^\t/, '').replace(/^const /, 'var ');
}

/** בלוק ברמת המודול (`var X = {...};` או `var X = [...];`). */
function extractTopLevelBlock(name: string, open: '{' | '['): string {
  const close = open === '{' ? '}' : ']';
  const match = CHUNK.match(
    new RegExp(`^var ${name} = \\${open}[\\s\\S]*?\\n\\${close};`, 'm'),
  );
  if (!match) throw new Error(`הקבוע ${name} לא נמצא ב-chunk — מבנה ה-chunk של superdoc השתנה`);
  return match[0];
}

/** `var X = /* @__PURE__ *\/ new Set([...]);` ברמת המודול. */
function extractTopLevelSet(name: string): string {
  const match = CHUNK.match(new RegExp(String.raw`^var ${name} = [^\n]*new Set\(\[[\s\S]*?\n\]\);`, 'm'));
  if (!match) throw new Error(`הקבוע ${name} לא נמצא ב-chunk — מבנה ה-chunk של superdoc השתנה`);
  return match[0];
}

/** שורה בודדת (`var X = ...;`). */
function extractTopLevelLine(name: string): string {
  const match = CHUNK.match(new RegExp(String.raw`^var ${name} = [^\n]*`, 'm'));
  if (!match) throw new Error(`הקבוע ${name} לא נמצא ב-chunk`);
  return match[0];
}

/**
 * `CLEAR_INLINE_PATCH` הוא הקבוע היחיד שהפונקציות שנחלצו נשענות עליו, והוא
 * נוגע רק ל-`kind: 'clear'` (נקה עיצוב) — פקודה בלי payload, שאין לה מה
 * לאמת.
 */
export const engine: EngineValidators = new Function(
  'CLEAR_INLINE_PATCH',
  [
    extractTopLevelBlock('SUPERDOC_UI_REASONS', '{'),
    extractTopLevelLine('UNSUPPORTED'),
    extractTopLevelLine('TABLE_CONTEXT'),
    ...TOP_LEVEL_SETS.map((name) => extractTopLevelSet(name)),
    ...TOP_LEVEL_FUNCTIONS.map((name) => extractTopLevelFunction(name)),
    ...CONTROLLER_CLOSURES.map((name) => extractControllerClosure(name)),
    extractTopLevelBlock('COMMAND_CATALOG', '['),
    `return { ${[...TOP_LEVEL_FUNCTIONS, ...CONTROLLER_CLOSURES, 'COMMAND_CATALOG'].join(', ')} };`,
  ].join('\n'),
)({}) as EngineValidators;

export const COMMAND_CATALOG: readonly EngineCommandDescriptor[] = engine.COMMAND_CATALOG;

const BY_ID = new Map(COMMAND_CATALOG.map((descriptor) => [descriptor.id, descriptor]));

/** ה-descriptor של פקודה, או `null` אם המנוע אינו מכיר את המזהה. */
export function commandDescriptor(id: string): EngineCommandDescriptor | null {
  return BY_ID.get(id) ?? null;
}

/** גוף ה-entry של פקודה בקטלוג, כטקסט. `id` הוא השדה הראשון בכל entry. */
export function descriptorSource(id: string): string {
  const start = CHUNK.indexOf(`id: "${id}",`);
  if (start === -1) throw new Error(`הפקודה ${id} אינה בקטלוג של superdoc`);
  const end = CHUNK.indexOf('\n\t},', start);
  return CHUNK.slice(start, end === -1 ? undefined : end);
}

/**
 * היעד שהוולידטורים מקבלים. תוכנו אינו נבדק על ידם — הם מכריעים לפי ה-payload
 * בלבד — ולכן הוא מייצג „יש בחירה” ולא יותר מזה.
 */
const INLINE_TARGET = { kind: 'text', segments: [] };
const BLOCK_ID = 'block-1';

export interface PayloadVerdict {
  /** האם המנוע היה מנתב את הפקודה עם ה-payload הזה. */
  accepted: boolean;
  /** האם היה בכלל ולידטור לשאול. `false` = הפקודה אינה בונה קלט מ-payload. */
  checked: boolean;
  /** איזה מסלול הכריע, לצורך הודעת כשל שאפשר לעשות איתה משהו. */
  route: string;
  /** מה שהמנוע היה מעביר ל-Document API. */
  input: unknown;
}

/**
 * מריצה על ה-payload את אותה שרשרת שה-controller מריץ: נרמול לפי ה-descriptor,
 * ואז בניית הקלט לפי סוג הפקודה. `accepted: false` פירושו שה-controller מחזיר
 * `false` בלי לגעת במסמך — כלומר כפתור שנלחץ ולא קרה כלום.
 */
export function checkPayload(id: string, payload?: unknown): PayloadVerdict {
  const descriptor = commandDescriptor(id);
  if (!descriptor) {
    return { accepted: false, checked: true, route: 'unknown-command', input: undefined };
  }

  const normalized = descriptor.normalizePayload ? descriptor.normalizePayload(payload) : payload;

  if (descriptor.inline) {
    const input = engine.buildInlineFormatInput(descriptor.inline, INLINE_TARGET, normalized, false);
    return {
      accepted: input !== null,
      checked: true,
      route: `inline:${descriptor.inline.kind}`,
      input,
    };
  }

  if (descriptor.blockParagraph) {
    const input = engine.buildBlockParagraphInput(descriptor.blockParagraph, BLOCK_ID, normalized);
    return {
      accepted: input !== null,
      checked: true,
      route: `paragraph:${descriptor.blockParagraph.kind}`,
      input,
    };
  }

  if (descriptor.instanceRoute) {
    return {
      accepted: engine.instanceCommandPayloadIsValid(descriptor, normalized),
      checked: true,
      route: `instance:${descriptor.instanceRoute}`,
      input: normalized,
    };
  }

  // פקודות שהקלט שלהן נבנה מהבחירה ולא מ-payload (טבלה, תמונה, קישור, undo).
  // אין להן ולידטור שאפשר להריץ בלי מסמך חי, ולכן הכפיל אינו מתיימר להכריע.
  return { accepted: true, checked: false, route: 'unchecked', input: normalized };
}

/**
 * גלריית הסגנונות. מה שנבדק כאן הוא שני דברים שהיו שבורים ואינם נראים בעין:
 *
 *   1. הרשימה הייתה קשיחה — שישה מזהים אנגליים עם תוויות שתרגמנו — ולכן
 *      הסגנונות של המסמך עצמו, כולל סגנונות מותאמים ושמות עבריים, לא הופיעו
 *      בגלריה כלל. הקטלוג נפתר אסינכרונית, ולכן בדיקה שקוראת פעם אחת אינה
 *      מוכיחה כלום; המסלול האמיתי נבדק דרך ההרשמה.
 *   2. כפתורי הגלילה היו הפוכים ב-RTL: `left: +120` נקרא „ימינה” בזמן שבמכולה
 *      כזאת הוא מחזיר את הרשימה לאחור, ימינה.
 *
 * הכפיל כאן מחזיר `StylesSlice` מלא ומטופס, ולא אובייקט מתירני: כפיל שמאשר כל
 * צורה מאשר גם צורה שהמנוע לא היה מחזיר, ואז הבדיקה ירוקה והתוסף שבור.
 */
import { describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { StyleCatalogItem } from 'superdoc/ui';
import {
  FALLBACK_STYLE_IDS,
  builtInStyleLabel,
  styleDisplayLabel,
  GALLERY_SCROLL_STEP_PX,
  HEADING_LADDER_FLOOR,
  clampPreviewFontSize,
  fallbackStyleGallery,
  galleryScrollAvailability,
  galleryScrollDelta,
  galleryScrollIcon,
  observeStyleGallery,
  previewStyleFor,
  readStyleGallery,
  styleRole,
  toGalleryItems,
  toGalleryState,
  withHeadingLadder,
  type StyleGalleryHost,
  type StyleGalleryState,
  type StylesSlice,
} from '../../src/engine/style-gallery';

/* ------------------------------------------------------------------ */
/* כפילים בצורת החוזה                                                  */
/* ------------------------------------------------------------------ */

/** פריט קטלוג מלא. כל השדות שהחוזה מחייב, כדי שהכפיל לא יהיה מתירני מהמנוע. */
function catalogItem(overrides: Partial<StyleCatalogItem> & { id: string }): StyleCatalogItem {
  return {
    name: overrides.id,
    aliases: [],
    type: 'paragraph',
    custom: false,
    builtin: true,
    default: false,
    basedOn: null,
    next: null,
    link: null,
    priority: null,
    qFormat: true,
    hidden: false,
    semiHidden: false,
    unhideWhenUsed: false,
    locked: false,
    provenance: 'authored',
    visibility: {
      quickGallery: true,
      recommended: true,
      all: true,
      effectivelyHidden: false,
    },
    ...overrides,
  };
}

/** slice מלא. ברירת המחדל היא המצב שלפני התייצבות הקטלוג. */
function slice(overrides: Partial<StylesSlice> = {}): StylesSlice {
  return {
    ready: true,
    status: 'ready',
    catalogRevision: 'rev-1',
    quickGallery: [],
    activeParagraphStyleId: null,
    activeParagraphStyleName: null,
    mixedSelection: false,
    sourceStatus: null,
    diagnostics: [],
    ...overrides,
  };
}

/** גלריה עברית, כמו במסמך שנוצר ב-Word בעברית. */
const HEBREW_GALLERY: readonly StyleCatalogItem[] = [
  catalogItem({ id: 'Normal', name: 'רגיל' }),
  catalogItem({ id: 'Heading1', name: 'כותרת 1' }),
  catalogItem({ id: 'Heading2', name: 'כותרת 2' }),
  catalogItem({ id: 'MyStyle', name: 'סגנון שלי', custom: true, builtin: false }),
];

/**
 * מזהי רשת הביטחון כפי שהגלריה מציגה אותם: `FALLBACK_STYLE_IDS` **אחרי**
 * השלמת סולם הכותרות. הרשימה כתובה במלואה ולא נגזרת מ-`withHeadingLadder`,
 * כדי שהבדיקה לא תאשר את עצמה; `סולם הכותרות` למטה מקבע את הקשר בין השתיים.
 */
const fallbackIds = ['Normal', 'NoSpacing', 'Heading1', 'Heading2', 'Heading3', 'Subtitle', 'Quote'];
const ids = (state: StyleGalleryState) => state.items.map((item) => item.id);

/* ------------------------------------------------------------------ */
/* נפילה סגורה                                                         */
/* ------------------------------------------------------------------ */

describe('נפילה לרשת הביטחון', () => {
  it('אין `ui` בכלל — הרשימה שלנו, לא רשימה ריקה', () => {
    expect(ids(readStyleGallery(null))).toEqual(fallbackIds);
    expect(ids(readStyleGallery(undefined))).toEqual(fallbackIds);
    expect(readStyleGallery(null).fromDocument).toBe(false);
  });

  it('גרסת מנוע בלי `styles` אינה מפילה את הרצועה', () => {
    expect(ids(readStyleGallery({}))).toEqual(fallbackIds);
  });

  it('קטלוג ריק לפני ההתייצבות (`status: pending`) — רשת הביטחון', () => {
    const host: StyleGalleryHost = {
      styles: { getSnapshot: () => slice({ status: 'pending', catalogRevision: null }) },
    };
    const state = readStyleGallery(host);
    expect(ids(state)).toEqual(fallbackIds);
    expect(state.fromDocument).toBe(false);
  });

  it('`getSnapshot` שזורק — נופל למסלול `getQuickGallery`', () => {
    const host: StyleGalleryHost = {
      styles: {
        getSnapshot: () => {
          throw new Error('not mounted');
        },
        getQuickGallery: () => HEBREW_GALLERY,
        getActiveParagraphStyle: () => ({
          styleId: 'Heading1',
          styleName: 'כותרת 1',
          mixed: false,
          diagnostics: [],
        }),
      },
    };
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const state = readStyleGallery(host);
    warn.mockRestore();

    expect(ids(state)).toEqual(['Normal', 'Heading1', 'Heading2', 'Heading3', 'MyStyle']);
    expect(state.activeId).toBe('Heading1');
  });

  it('`getQuickGallery` שזורק — רשת הביטחון ולא קריסה', () => {
    const host: StyleGalleryHost = {
      styles: {
        getQuickGallery: () => {
          throw new Error('worker-backed editor');
        },
      },
    };
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const state = readStyleGallery(host);
    warn.mockRestore();

    expect(ids(state)).toEqual(fallbackIds);
    expect(state.fromDocument).toBe(false);
  });

  it('רשת הביטחון אינה משותפת — קורא אחד אינו משנה אותה לכולם', () => {
    const first = fallbackStyleGallery();
    const second = fallbackStyleGallery();
    expect(first.items).not.toBe(second.items);
    expect(first.items[0].previewStyle).not.toBe(second.items[0].previewStyle);
  });

  it('לכל פריט ברשת הביטחון יש צבע מתפקיד ה-API, בלי hex', () => {
    for (const item of fallbackStyleGallery().items) {
      expect(item.previewStyle.color).toMatch(/^var\(--color-/);
    }
  });
});

/* ------------------------------------------------------------------ */
/* הגלריה של המסמך                                                     */
/* ------------------------------------------------------------------ */

describe('הגלריה מהמסמך', () => {
  it('הפריטים והשמות מהמנוע — כולל סגנון מותאם ושמות עבריים', () => {
    const state = toGalleryState(slice({ quickGallery: HEBREW_GALLERY }));
    // „כותרת 3” אינה במנוע והיא בגלריה: ראו `withHeadingLadder`.
    expect(ids(state)).toEqual(['Normal', 'Heading1', 'Heading2', 'Heading3', 'MyStyle']);
    expect(state.items.map((item) => item.label)).toEqual([
      'רגיל',
      'כותרת 1',
      'כותרת 2',
      'כותרת 3',
      'סגנון שלי',
    ]);
    expect(state.fromDocument).toBe(true);
  });

  it('שם שהמנוע לא מסר נופל למזהה, ולא לכרטיס ריק', () => {
    // מזהה שאינו בטבלת הסגנונות המובנים, כדי שהנפילה לשם ולמזהה תהיה זו שנמדדת.
    const items = toGalleryItems([catalogItem({ id: 'AcmeLead', name: '   ' })]);
    expect(items[0].label).toBe('AcmeLead');
  });

  it('פריט בלי מזהה וכפילות נשמטים', () => {
    const items = toGalleryItems([
      catalogItem({ id: '  ' }),
      catalogItem({ id: 'Normal' }),
      catalogItem({ id: 'Normal', name: 'שוב' }),
    ]);
    expect(items.map((item) => item.id)).toEqual(['Normal', 'Heading1', 'Heading2', 'Heading3']);
  });

  it('סגנון מוסתר אינו מוצג', () => {
    const items = toGalleryItems([
      catalogItem({
        id: 'Hidden',
        visibility: { quickGallery: false, recommended: false, all: true, effectivelyHidden: true },
      }),
      catalogItem({ id: 'Normal' }),
    ]);
    // `Hidden` נשמט; מה שנשאר הוא „רגיל” וסולם הכותרות שמושלם אחריו.
    expect(items.map((item) => item.id)).toEqual(['Normal', 'Heading1', 'Heading2', 'Heading3']);
  });

  it('`activeParagraphStyleId` הוא הכרטיס הפעיל', () => {
    const state = toGalleryState(
      slice({ quickGallery: HEBREW_GALLERY, activeParagraphStyleId: 'Heading2' }),
    );
    expect(state.activeId).toBe('Heading2');
  });

  it('בחירה מעורבת — אין כרטיס פעיל', () => {
    const state = toGalleryState(
      slice({
        quickGallery: HEBREW_GALLERY,
        activeParagraphStyleId: 'Heading2',
        mixedSelection: true,
      }),
    );
    expect(state.activeId).toBeNull();
  });

  it('`mixed` במסלול `getActiveParagraphStyle` מתנהג אותו דבר', () => {
    const host: StyleGalleryHost = {
      styles: {
        getQuickGallery: () => HEBREW_GALLERY,
        getActiveParagraphStyle: () => ({
          styleId: null,
          styleName: null,
          mixed: true,
          diagnostics: [],
        }),
      },
    };
    expect(readStyleGallery(host).activeId).toBeNull();
  });
});

/* ------------------------------------------------------------------ */
/* ההרשמה — המסלול האמיתי                                              */
/* ------------------------------------------------------------------ */

describe('הרשמה לקטלוג', () => {
  /** `observe` של המנוע: יורה מיד, ואז על כל שינוי. */
  function observableHost(): {
    host: StyleGalleryHost;
    settle: (next: StylesSlice) => void;
    disposed: () => boolean;
  } {
    let current = slice({ status: 'pending', catalogRevision: null });
    const listeners = new Set<(value: StylesSlice) => void>();
    let disposed = false;

    return {
      host: {
        styles: {
          getSnapshot: () => current,
          observe: (listener) => {
            listeners.add(listener);
            listener(current);
            return () => {
              disposed = true;
              listeners.delete(listener);
            };
          },
        },
      },
      settle(next) {
        current = next;
        for (const listener of listeners) listener(next);
      },
      disposed: () => disposed,
    };
  }

  it('הגלריה מתעדכנת כשהקטלוג נפתר אחרי הפתיחה', () => {
    const { host, settle } = observableHost();
    const seen: StyleGalleryState[] = [];
    observeStyleGallery(host, (state) => seen.push(state));

    // הירייה המיידית: הקטלוג עוד לא נפתר, ולכן רשת הביטחון.
    expect(seen).toHaveLength(1);
    expect(ids(seen[0])).toEqual(fallbackIds);
    expect(seen[0].fromDocument).toBe(false);

    settle(slice({ quickGallery: HEBREW_GALLERY, activeParagraphStyleId: 'Normal' }));

    expect(seen).toHaveLength(2);
    expect(ids(seen[1])).toEqual(['Normal', 'Heading1', 'Heading2', 'Heading3', 'MyStyle']);
    expect(seen[1].fromDocument).toBe(true);
    expect(seen[1].activeId).toBe('Normal');
  });

  it('ה-disposer של המנוע מוחזר ומשוחרר', () => {
    const { host, disposed } = observableHost();
    const off = observeStyleGallery(host, () => {});
    expect(disposed()).toBe(false);
    off();
    expect(disposed()).toBe(true);
  });

  it('בלי `observe` נשלח המצב הנוכחי, ומוחזר disposer שאין בו כלום', () => {
    const seen: StyleGalleryState[] = [];
    const off = observeStyleGallery(
      { styles: { getQuickGallery: () => HEBREW_GALLERY } },
      (state) => seen.push(state),
    );
    expect(seen).toHaveLength(1);
    expect(seen[0].fromDocument).toBe(true);
    expect(() => off()).not.toThrow();
  });

  it('`observe` שזורק אינו מפיל את פתיחת המסמך', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const seen: StyleGalleryState[] = [];
    const off = observeStyleGallery(
      {
        styles: {
          observe: () => {
            throw new Error('not-ready');
          },
        },
      },
      (state) => seen.push(state),
    );
    warn.mockRestore();

    expect(ids(seen[0])).toEqual(fallbackIds);
    expect(() => off()).not.toThrow();
  });
});

/* ------------------------------------------------------------------ */
/* תצוגה מקדימה                                                        */
/* ------------------------------------------------------------------ */

describe('תצוגה מקדימה', () => {
  it('התפקיד נגזר מהמזהה האנגלי, שאינו מתורגם במסמך עברי', () => {
    expect(styleRole('Heading1')).toBe('heading');
    expect(styleRole('Title')).toBe('heading');
    expect(styleRole('Subtitle')).toBe('subtle');
    expect(styleRole('IntenseQuote')).toBe('subtle');
    expect(styleRole('Normal')).toBe('body');
    expect(styleRole('MyStyle')).toBe('body');
  });

  it('הטיפוגרפיה מהמנוע, והצבע מהתפקיד — כדי שמצב כהה יישאר קריא', () => {
    const style = previewStyleFor(
      catalogItem({
        id: 'Heading1',
        name: 'כותרת 1',
        preview: {
          available: true,
          // צבע הסגנון במסמך. ב-Word ברירת המחדל היא כחול כהה, ועל משטח כהה
          // הוא בלתי קריא — ולכן אינו עובר.
          css: { fontWeight: '600', fontFamily: 'Cambria', color: '#2e74b5', margin: '24pt 0' },
        },
      }),
    );

    expect(style.fontWeight).toBe('600');
    expect(style.fontFamily).toBe('Cambria');
    expect(style.color).toBe('var(--color-primary)');
    expect(style.margin).toBeUndefined();
  });

  it('`available: false` — התפקיד לבדו', () => {
    const style = previewStyleFor(
      catalogItem({
        id: 'Subtitle',
        preview: { available: false, unsupportedReason: 'not-resolved' },
      }),
    );
    expect(style.fontStyle).toBe('italic');
    expect(style.color).toBe('var(--color-on-surface-variant)');
  });

  it('אין אף צבע קשיח בשום פריט שהגלריה מייצרת', () => {
    const items = toGalleryItems([
      ...HEBREW_GALLERY,
      catalogItem({
        id: 'Title',
        preview: { available: true, css: { color: 'rgb(31, 78, 120)', fontSize: '32pt' } },
      }),
    ]);
    for (const item of items) {
      expect(JSON.stringify(item.previewStyle)).not.toMatch(/#[0-9a-f]{3,8}|rgb|hsl/i);
    }
  });

  it('גודל הכותרת מוקטן לגבולות הכרטיס', () => {
    expect(clampPreviewFontSize('32pt')).toBe('16px');
    expect(clampPreviewFontSize('16pt')).toBe('16px');
    expect(clampPreviewFontSize('8px')).toBe('11px');
    expect(clampPreviewFontSize('1em')).toBe('13px');
    expect(clampPreviewFontSize(14)).toBe('14px');
  });

  it('גודל שאינו נקרא אינו נכתב', () => {
    expect(clampPreviewFontSize(undefined)).toBeUndefined();
    expect(clampPreviewFontSize('inherit')).toBeUndefined();
    expect(clampPreviewFontSize('0px')).toBeUndefined();
  });

  it('כרטיס גוף מציג „AaBbCc”, וכרטיס כותרת את שם הסגנון — כמו ב-Word', () => {
    const items = toGalleryItems(HEBREW_GALLERY);
    expect(items.find((item) => item.id === 'Normal')?.previewText).toBe('AaBbCc');
    expect(items.find((item) => item.id === 'Heading1')?.previewText).toBe('כותרת 1');
  });
});

/* ------------------------------------------------------------------ */
/* סולם הכותרות                                                        */
/* ------------------------------------------------------------------ */

/**
 * הבאג שנסגר כאן: המנוע הציע „כותרת 1” בלבד — ב-`word/styles.xml` של התבנית
 * `Heading2`…`Heading9` נושאים `w:semiHidden`, ו-`quickGallery` מסנן אותם —
 * ולכן לא הייתה דרך במסך להחיל רמה 2, ותוכן העניינים נבנה מרמה אחת גם אחרי
 * „התאמה אישית” שמבקשת 1–3. ההנמקה המלאה ב-engine/style-gallery.ts.
 */
describe('סולם הכותרות', () => {
  it('גלריה שמציעה „כותרת 1” בלבד מקבלת את 2 ו-3, מיד אחריה', () => {
    const items = toGalleryItems([
      catalogItem({ id: 'Normal' }),
      catalogItem({ id: 'Heading1' }),
      catalogItem({ id: 'Title' }),
    ]);
    expect(items.map((item) => item.id)).toEqual([
      'Normal',
      'Heading1',
      'Heading2',
      'Heading3',
      'Title',
    ]);
  });

  it('רמה שהמסמך מסתיר מסונתזת עם התווית העברית ובטוקני התפקיד', () => {
    const items = toGalleryItems([catalogItem({ id: 'Heading1' })]);
    const level2 = items.find((item) => item.id === 'Heading2');
    expect(level2?.label).toBe('כותרת 2');
    // הכרטיס של כותרת מציג את שמה, בשונה מ„AaBbCc” של סגנון גוף.
    expect(level2?.previewText).toBe('כותרת 2');
    expect(level2?.previewStyle.fontWeight).toBe('700');
  });

  it('מסמך שמציע רמה עמוקה שומר אותה בלי להמציא את הרמות שבאמצע', () => {
    const items = toGalleryItems([
      catalogItem({ id: 'Normal' }),
      catalogItem({ id: 'Heading1' }),
      catalogItem({ id: 'Heading5' }),
    ]);
    expect(items.map((item) => item.id)).toEqual([
      'Normal',
      'Heading1',
      'Heading2',
      'Heading3',
      'Heading5',
    ]);
  });

  it('רמה 9 קיימת נשארת, אך אינה פותחת כרטיסים חסרים 4–8', () => {
    const items = toGalleryItems([catalogItem({ id: 'Heading9' })]);
    expect(items.map((item) => item.id)).toEqual([
      'Heading1',
      'Heading2',
      'Heading3',
      'Heading9',
    ]);
  });

  it('`Title` הוא תפקיד כותרת ואינו רמה — הסולם נבנה בלעדיו', () => {
    const items = toGalleryItems([catalogItem({ id: 'Title' })]);
    expect(items.map((item) => item.id)).toEqual(['Title', 'Heading1', 'Heading2', 'Heading3']);
  });

  it('רשימה ריקה נשארת ריקה — היא מה שמפעיל את רשת הביטחון', () => {
    expect(withHeadingLadder([])).toEqual([]);
    expect(toGalleryItems([])).toEqual([]);
    expect(toGalleryItems(undefined)).toEqual([]);
  });

  it('כל מזהה שברשת הביטחון נשאר בגלריה שלה', () => {
    const shown = fallbackStyleGallery().items.map((item) => item.id);
    for (const id of FALLBACK_STYLE_IDS) expect(shown).toContain(id);
    expect(HEADING_LADDER_FLOOR).toBe(3);
  });
});

/* ------------------------------------------------------------------ */
/* תוויות עבריות                                                       */
/* ------------------------------------------------------------------ */

describe('תוויות הסגנונות המובנים', () => {
  it('מכסה את המזהים שבתבניות של Word', () => {
    const expected: Record<string, string> = {
      Normal: 'רגיל',
      NoSpacing: 'ללא מרווח',
      Title: 'כותרת',
      Subtitle: 'כותרת משנה',
      Heading1: 'כותרת 1',
      Heading2: 'כותרת 2',
      Heading3: 'כותרת 3',
      Heading4: 'כותרת 4',
      Heading5: 'כותרת 5',
      Heading6: 'כותרת 6',
      Quote: 'ציטוט',
      IntenseQuote: 'ציטוט מודגש',
      ListParagraph: 'פסקת רשימה',
      Caption: 'כתובית',
      Header: 'כותרת עליונה',
      Footer: 'כותרת תחתונה',
    };
    for (const [id, label] of Object.entries(expected)) {
      expect(builtInStyleLabel(id)).toBe(label);
    }
  });

  it('„heading 1” של התבנית — אות קטנה ורווח — מתורגם כמו `Heading1`', () => {
    // זו הרגרסיה מהריצה החיה: השם בקטלוג של המסמך הריק אינו ה-styleId הקנוני.
    expect(builtInStyleLabel('heading 1')).toBe('כותרת 1');
    expect(builtInStyleLabel('No Spacing')).toBe('ללא מרווח');
    expect(builtInStyleLabel('intense_quote')).toBe('ציטוט מודגש');
  });

  it('מזהה שאינו בטבלה מחזיר `undefined`, ולא מחרוזת ריקה', () => {
    expect(builtInStyleLabel('MyStyle')).toBeUndefined();
    expect(builtInStyleLabel('')).toBeUndefined();
    expect(builtInStyleLabel(undefined)).toBeUndefined();
  });

  it('המסמך הריק של המנוע — תבנית אנגלית — מוצג בעברית', () => {
    // בדיוק החמישה שנראו בצילום המסך: Normal, heading 1, Title, Subtitle, Quote.
    const english: readonly StyleCatalogItem[] = [
      catalogItem({ id: 'Normal', name: 'Normal' }),
      catalogItem({ id: 'Heading1', name: 'heading 1' }),
      catalogItem({ id: 'Title', name: 'Title' }),
      catalogItem({ id: 'Subtitle', name: 'Subtitle' }),
      catalogItem({ id: 'Quote', name: 'Quote' }),
    ];
    expect(toGalleryItems(english).map((item) => item.label)).toEqual([
      'רגיל',
      'כותרת 1',
      // שתי אלה מסונתזות: התבנית מסתירה אותן, וראו `withHeadingLadder`.
      'כותרת 2',
      'כותרת 3',
      'כותרת',
      'כותרת משנה',
      'ציטוט',
    ]);
  });

  it('סגנון מותאם אישית שומר את שמו — גם כששמו „Quote”', () => {
    expect(
      styleDisplayLabel(
        catalogItem({ id: 'Quote1', name: 'Quote', custom: true, builtin: false }),
      ),
    ).toBe('Quote');
    expect(
      styleDisplayLabel(catalogItem({ id: 'MyStyle', name: 'סגנון שלי', custom: true })),
    ).toBe('סגנון שלי');
  });

  it('מזהה שאינו מוכר נופל לשם שבמסמך, ושם ריק נופל למזהה', () => {
    expect(styleDisplayLabel(catalogItem({ id: 'AcmeLead', name: 'Acme Lead' }))).toBe('Acme Lead');
    expect(styleDisplayLabel(catalogItem({ id: 'AcmeLead', name: '   ' }))).toBe('AcmeLead');
  });

  it('מסמך עברי אינו משתנה — השם שבמסמך הוא כבר התשובה', () => {
    expect(styleDisplayLabel(catalogItem({ id: 'Heading1', name: 'כותרת 1' }))).toBe('כותרת 1');
    expect(styleDisplayLabel(catalogItem({ id: 'AcmeLead', name: 'פתיח' }))).toBe('פתיח');
  });

  it('`styleId` מקומי מתורגם דרך השם או ה-aliases האנגליים', () => {
    // תבניות לא-אנגליות כותבות לעיתים styleId מקומי ומשאירות את השם הקנוני.
    expect(styleDisplayLabel(catalogItem({ id: 'berschrift1', name: 'Heading 1' }))).toBe(
      'כותרת 1',
    );
    expect(
      styleDisplayLabel(
        catalogItem({ id: 'Standardowy', name: 'Styl domyslny', aliases: ['Normal'] }),
      ),
    ).toBe('רגיל');
  });

  it('רשת הביטחון מקבלת את התוויות מאותה טבלה', () => {
    expect(fallbackStyleGallery().items.map((item) => item.label)).toEqual([
      'רגיל',
      'ללא מרווח',
      'כותרת 1',
      'כותרת 2',
      'כותרת 3',
      'כותרת משנה',
      'ציטוט',
    ]);
  });
});

/* ------------------------------------------------------------------ */
/* גאומטריית הכרטיס                                                    */
/* ------------------------------------------------------------------ */

/**
 * הליקוי שנראה בריצה חיה: התקרה של מיכל הגלילה הייתה 340px — מספר עגול שאינו
 * כפולה של רוחב כרטיס, ולכן הכרטיס החמישי הוצג חצוי. ב-Word הגלריה מציגה
 * כרטיסים שלמים בלבד. מה שנבדק כאן הוא שהתקרה **מחושבת** מרוחב הכרטיס ולא
 * חוזרת להיות מספר קסם.
 */
describe('רוחב הגלריה נצמד לגבול כרטיס', () => {
  const source = readFileSync(
    join(process.cwd(), 'src/ui/ribbon/common/StyleGallery.vue'),
    'utf8',
  );

  it('רוחב הכרטיס קבוע, ולא טווח שאין לו גבול', () => {
    expect(source).toMatch(/width:\s*var\(--style-card-width\)/);
    expect(source).not.toMatch(/\bmin-width:\s*\d+px/);
    expect(source).not.toMatch(/\bmax-width:\s*\d+px/);
  });

  it('התקרה מחושבת מרוחב הכרטיס, מהמרווח ומהריפוד', () => {
    const declaration = /max-width:\s*calc\(([\s\S]*?)\);/.exec(source)?.[1] ?? '';
    expect(declaration).toContain('--style-cards-visible');
    expect(declaration).toContain('--style-card-width');
    expect(declaration).toContain('--style-card-gap');
    expect(declaration).toContain('--style-cards-padding');
  });

  it('הגלילה נצמדת לכרטיס גם כשהיא לא מגיעה מהכפתורים', () => {
    expect(source).toMatch(/scroll-snap-type:\s*inline mandatory/);
    expect(source).toMatch(/scroll-snap-align:\s*start/);
  });

  it('צעד הכפתור הוא כפולה שלמה של פסיעת כרטיס', () => {
    const width = Number(/--style-card-width:\s*(\d+)px/.exec(source)?.[1]);
    const gap = Number(/--style-card-gap:\s*(\d+)px/.exec(source)?.[1]);
    expect(Number.isFinite(width) && Number.isFinite(gap)).toBe(true);
    expect(GALLERY_SCROLL_STEP_PX % (width + gap)).toBe(0);
  });
});

/* ------------------------------------------------------------------ */
/* גלילה                                                               */
/* ------------------------------------------------------------------ */

describe('כיוון הגלילה', () => {
  it('ב-RTL „הסגנונות הבאים” הם delta שלילי — זה ההיפוך שהיה שבור', () => {
    expect(galleryScrollDelta('end', true)).toBe(-GALLERY_SCROLL_STEP_PX);
    expect(galleryScrollDelta('start', true)).toBe(GALLERY_SCROLL_STEP_PX);
  });

  it('ב-LTR הכיוונים מתהפכים', () => {
    expect(galleryScrollDelta('end', false)).toBe(GALLERY_SCROLL_STEP_PX);
    expect(galleryScrollDelta('start', false)).toBe(-GALLERY_SCROLL_STEP_PX);
  });

  it('החץ תואם את הכיוון שעל המסך', () => {
    expect(galleryScrollIcon('end', true)).toBe('chevronLeft');
    expect(galleryScrollIcon('start', true)).toBe('chevronRight');
    expect(galleryScrollIcon('end', false)).toBe('chevronRight');
    expect(galleryScrollIcon('start', false)).toBe('chevronLeft');
  });
});

describe('זמינות הגלילה', () => {
  it('גלריה שנכנסת כולה — אין כפתורים', () => {
    expect(galleryScrollAvailability({ scrollLeft: 0, scrollWidth: 300, clientWidth: 340 })).toEqual({
      canScrollStart: false,
      canScrollEnd: false,
    });
  });

  it('בתחילת הרשימה אפשר רק קדימה', () => {
    expect(galleryScrollAvailability({ scrollLeft: 0, scrollWidth: 800, clientWidth: 340 })).toEqual({
      canScrollStart: false,
      canScrollEnd: true,
    });
  });

  it('ב-RTL ה-scrollLeft שלילי, וזה נמדד לפי המרחק מתחילת הרשימה', () => {
    expect(
      galleryScrollAvailability({ scrollLeft: -200, scrollWidth: 800, clientWidth: 340 }),
    ).toEqual({ canScrollStart: true, canScrollEnd: true });

    // הקצה: 800 - 340 = 460.
    expect(
      galleryScrollAvailability({ scrollLeft: -460, scrollWidth: 800, clientWidth: 340 }),
    ).toEqual({ canScrollStart: true, canScrollEnd: false });
  });

  it('ב-LTR אותו חשבון עם סימן חיובי', () => {
    expect(
      galleryScrollAvailability({ scrollLeft: 460, scrollWidth: 800, clientWidth: 340 }),
    ).toEqual({ canScrollStart: true, canScrollEnd: false });
  });

  it('הפרש תת-פיקסלי אינו כפתור', () => {
    expect(
      galleryScrollAvailability({ scrollLeft: 0.4, scrollWidth: 340.6, clientWidth: 340 }),
    ).toEqual({ canScrollStart: false, canScrollEnd: false });
  });
});

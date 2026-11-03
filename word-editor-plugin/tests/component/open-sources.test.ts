/**
 * העץ והרשימה של „פתח מסמך”, כפי שהמשתמש פוגש אותם.
 *
 * מה שנמדד כאן: שהאירועים נושאים את המקום/הספר/הקובץ הנכונים (החלפה בין
 * שניים היא בדיוק הבאג שסריקת מקור מאשרת בירוק), שהחצים בעץ פועלים בכיוון
 * של ממשק עברי, ושכל מצב — טעינה, שגיאה, ריק, חיפוש — מציג את מה שהוא אמור.
 */
import { describe, expect, it } from 'vitest';
import OpenSourcesNav from '../../src/ui/panels/OpenSourcesNav.vue';
import OpenSourceList from '../../src/ui/panels/OpenSourceList.vue';
import OpenDocumentDialog from '../../src/ui/panels/OpenDocumentDialog.vue';
import type { ListingState } from '../../src/composables/use-open-sources';
import {
  listingKey,
  pruneLibraryTree,
  type DocFolder,
  type OpenPlace,
} from '../../src/sessions/open-sources';
import { autoUnmount, mountUi, settle } from './harness';

autoUnmount();

const LIBRARY = pruneLibraryTree({
  categories: [
    {
      title: 'הלכה',
      path: '/הלכה',
      books: [{ title: 'קונטרס', type: 'docx', bookUid: 'uid:1', author: 'פלוני' }],
      categories: [{ title: 'שו"ת', path: '/הלכה/שו"ת', books: [{ title: 'שו"ת דברי שלום', type: 'docx', bookUid: 'uid:2' }] }],
    },
    { title: 'אישי', path: '/אישי', books: [{ title: 'חידושים', type: 'docx', bookUid: 'uid:3', source: 'user' }] },
  ],
});

const FOLDER: DocFolder = { token: 'ft', name: 'העבודות שלי', path: 'C:\\Users\\x\\העבודות שלי', addedAt: 1 };

const READY: ListingState = {
  state: 'ready',
  refreshing: false,
  listing: {
    truncated: false,
    entries: [
      { name: 'שיעורים', path: 'שיעורים', type: 'dir', size: 0, modified: 0 },
      { name: 'מכתב.docx', path: 'מכתב.docx', type: 'file', size: 2048, modified: 0 },
    ],
  },
};

function nav(props: Record<string, unknown> = {}) {
  return mountUi(OpenSourcesNav, {
    props: {
      place: { kind: 'recent' },
      recentCount: 4,
      library: LIBRARY,
      libraryState: 'ready',
      personalCount: 1,
      folders: [FOLDER],
      listings: new Map([[listingKey('ft', ''), READY]]),
      ...props,
    },
  });
}

function labels(wrapper: ReturnType<typeof nav>['wrapper']): string[] {
  return wrapper.findAll('[role="treeitem"] .src-label').map((el) => el.text());
}

describe('OpenSourcesNav — מה בעץ', () => {
  it('אחרונים, הספרייה פתוחה ברמה הראשונה, אישיים, והתיקיות', () => {
    const { wrapper } = nav();
    expect(labels(wrapper)).toEqual([
      'אחרונים',
      'ספריית אוצריא',
      'הלכה',
      'אישי',
      'ספרים אישיים',
      'התיקיות שלי',
      'העבודות שלי',
    ]);
  });

  it('הצומת הנבחר מוכרז, ורק הוא בסדר ה-Tab', () => {
    const { wrapper } = nav({ place: { kind: 'library', path: '/הלכה' } });
    const items = wrapper.findAll('[role="treeitem"]');
    const selected = items.filter((el) => el.attributes('aria-selected') === 'true');
    expect(selected.map((el) => el.find('.src-label').text())).toEqual(['הלכה']);
    expect(items.filter((el) => el.attributes('tabindex') === '0')).toHaveLength(1);
  });

  it('ספרייה שאוצריא אינה תומכת בה — הענף אינו מרונדר', () => {
    const { wrapper } = nav({ library: null, libraryState: 'unsupported' });
    expect(labels(wrapper)).not.toContain('ספריית אוצריא');
  });

  it('ספרייה בטעינה ראשונה — הענף מופיע, מעומעם, עם „טוען…” בטולטיפ', () => {
    const { wrapper } = nav({ library: null, libraryState: 'loading' });
    const root = wrapper.findAll('[role="treeitem"]')[1]!;
    expect(root.find('.src-label').text()).toBe('ספריית אוצריא');
    expect(root.classes()).toContain('src-item--muted');
    expect(root.attributes('data-tip-title')).toBe('טוען…');
  });

  it('אוצריא בלי תיקיות: אין כפתור הוספה, ויש הסבר', () => {
    const { wrapper } = nav({ folders: [], foldersSupported: false });
    expect(wrapper.find('.src-add').exists()).toBe(false);
    expect(wrapper.find('.src-note').text()).toContain('גרסה חדשה');
  });
});

describe('OpenSourcesNav — לחיצה ומקלדת', () => {
  it('לחיצה על מדף פולטת את המקום שלו', async () => {
    const { wrapper } = nav();
    await wrapper.findAll('[role="treeitem"]')[2]!.trigger('click');
    expect(wrapper.emitted('select')?.[0]).toEqual([{ kind: 'library', path: '/הלכה' }]);
  });

  it('„התיקיות שלי” נסגר ונפתח ואינו נבחר', async () => {
    const { wrapper } = nav();
    const group = () => wrapper.findAll('[role="treeitem"]').find((el) => el.text().includes('התיקיות שלי'))!;
    await group().trigger('click');
    expect(wrapper.emitted('select')).toBeUndefined();
    expect(labels(wrapper)).not.toContain('העבודות שלי');
    expect(group().attributes('aria-expanded')).toBe('false');
  });

  it('ArrowDown בוחר את הצומת הבא — הבחירה עוקבת אחרי המיקוד', async () => {
    const { wrapper } = nav();
    await wrapper.find('[role="tree"]').trigger('keydown', { key: 'ArrowDown' });
    expect(wrapper.emitted('select')?.[0]).toEqual([{ kind: 'library', path: '/' }]);
  });

  it('ב-RTL: → סוגר, ← פותח ופולט טעינת תת-תיקיות לתיקייה', async () => {
    const { wrapper } = nav({ place: { kind: 'library', path: '/' } });
    const tree = wrapper.find('[role="tree"]');
    await tree.trigger('keydown', { key: 'ArrowRight' });
    expect(labels(wrapper)).not.toContain('הלכה');
    await tree.trigger('keydown', { key: 'ArrowLeft' });
    expect(labels(wrapper)).toContain('הלכה');

    const { wrapper: w2 } = nav({ place: { kind: 'folder', token: 'ft', path: '' } });
    await w2.find('[role="tree"]').trigger('keydown', { key: 'ArrowLeft' });
    const expands = w2.emitted('expand-folder') ?? [];
    expect(expands[expands.length - 1]).toEqual(['ft', '']);
    expect(labels(w2)).toContain('שיעורים');
  });

  it('„הוסף תיקייה…” פולט add-folder', async () => {
    const { wrapper } = nav();
    await wrapper.find('.src-add').trigger('click');
    expect(wrapper.emitted('add-folder')).toHaveLength(1);
  });
});

function list(place: OpenPlace, props: Record<string, unknown> = {}) {
  return mountUi(OpenSourceList, {
    props: { place, library: LIBRARY, libraryState: 'ready', folder: FOLDER, listing: READY, ...props },
  });
}

describe('OpenSourceList', () => {
  it('מדף: תתי-מדפים לפני ספרים, ולחיצה על ספר פולטת אותו', async () => {
    const { wrapper } = list({ kind: 'library', path: '/הלכה' });
    const names = wrapper.findAll('.sl-name').map((el) => el.text());
    expect(names).toEqual(['שו"ת', 'קונטרס']);
    await wrapper.findAll('.sl-open')[1]!.trigger('click');
    expect(wrapper.emitted('open-book')?.[0]?.[0]).toMatchObject({ title: 'קונטרס', bookUid: 'uid:1' });
  });

  it('לחיצה על מדף נכנסת אליו', async () => {
    const { wrapper } = list({ kind: 'library', path: '/הלכה' });
    await wrapper.findAll('.sl-open')[0]!.trigger('click');
    expect(wrapper.emitted('navigate')?.[0]).toEqual([{ kind: 'library', path: '/הלכה/שו"ת' }]);
  });

  it('חיפוש בספרייה מוצא בכל העץ, לא רק במדף, עם המיקום כשורה שנייה', () => {
    const { wrapper } = list({ kind: 'library', path: '/אישי' }, { query: 'שו״ת' });
    expect(wrapper.findAll('.sl-name').map((el) => el.text())).toEqual(['שו"ת דברי שלום']);
    expect(wrapper.find('.sl-sub').text()).toBe('הלכה / שו"ת');
  });

  it('שביל הלחם: כל מקטע חוץ מהאחרון חוזר אליו', async () => {
    const { wrapper } = list({ kind: 'library', path: '/הלכה/שו"ת' });
    const crumbs = wrapper.findAll('button.sl-crumb');
    expect(crumbs.map((el) => el.text())).toEqual(['ספריית אוצריא', 'הלכה']);
    await crumbs[1]!.trigger('click');
    expect(wrapper.emitted('navigate')?.[0]).toEqual([{ kind: 'library', path: '/הלכה' }]);
  });

  it('שם קובץ מוצג בלי .docx, ו-.docm נשאר', () => {
    const { wrapper } = list({ kind: 'folder', token: 'ft', path: '' }, {
      listing: {
        state: 'ready',
        refreshing: false,
        listing: {
          truncated: false,
          entries: [
            { name: 'א.docx', path: 'א.docx', type: 'file', size: 1, modified: 0 },
            { name: 'ב.docm', path: 'ב.docm', type: 'file', size: 1, modified: 0 },
          ],
        },
      },
    });
    expect(wrapper.findAll('.sl-name').map((el) => el.text())).toEqual(['א', 'ב.docm']);
  });

  it('תיקייה: לחיצה על קובץ פולטת token, נתיב ושם; על תת-תיקייה — ניווט', async () => {
    const { wrapper } = list({ kind: 'folder', token: 'ft', path: '' });
    const rows = wrapper.findAll('.sl-open');
    await rows[1]!.trigger('click');
    expect(wrapper.emitted('open-file')?.[0]).toEqual(['ft', 'מכתב.docx', 'מכתב.docx']);
    await rows[0]!.trigger('click');
    expect(wrapper.emitted('navigate')?.[0]).toEqual([{ kind: 'folder', token: 'ft', path: 'שיעורים' }]);
  });

  it('תיקייה שנעלמה: הסבר, „נסה שוב” ו„הסר מהרשימה”', async () => {
    const { wrapper } = list(
      { kind: 'folder', token: 'ft', path: '' },
      { listing: { state: 'error', reason: 'not-found', message: 'x' } },
    );
    expect(wrapper.find('.sl-state__title').text()).toBe('התיקייה לא נמצאה');
    await wrapper.find('.sl-link--danger').trigger('click');
    expect(wrapper.emitted('remove-folder')?.[0]).toEqual(['ft']);
  });

  it('טעינה ראשונה מציגה „טוען…”, ורענון של רשימה קיימת אינו מסתיר אותה', () => {
    expect(list({ kind: 'folder', token: 'ft', path: '' }, { listing: { state: 'loading' } }).wrapper.text()).toContain('טוען…');
    const { wrapper } = list({ kind: 'folder', token: 'ft', path: '' }, { listing: { ...READY, refreshing: true } });
    expect(wrapper.findAll('.sl-open')).toHaveLength(2);
    expect(wrapper.find('.sl-iconbtn--spin').exists()).toBe(true);
  });

  it('חיפוש בלי תוצאות אומר מה חיפשו', () => {
    const { wrapper } = list({ kind: 'folder', token: 'ft', path: '' }, { query: 'אין כזה' });
    expect(wrapper.find('.sl-state__title').text()).toBe('לא נמצא דבר שתואם ל„אין כזה”');
  });

  it('ArrowRight עולה רמה — אחורה ב-RTL', async () => {
    const { wrapper } = list({ kind: 'folder', token: 'ft', path: 'שיעורים' }, {
      listing: { state: 'ready', refreshing: false, listing: { truncated: false, entries: [{ name: 'א.docx', path: 'שיעורים/א.docx', type: 'file', size: 1, modified: 0 }] } },
    });
    await wrapper.find('.sl-list').trigger('keydown', { key: 'ArrowRight' });
    expect(wrapper.emitted('navigate')?.[0]).toEqual([{ kind: 'folder', token: 'ft', path: '' }]);
  });
});

describe('OpenDocumentDialog — עם העץ', () => {
  it('בלי props של העץ: „אחרונים” נבחר, והרשימה הקיימת היא מה שמוצג', () => {
    const { wrapper } = mountUi(OpenDocumentDialog, { props: { isOpen: true, recents: [] } });
    expect(wrapper.find('.rec-section').exists()).toBe(true);
    expect(wrapper.find('.sl').exists()).toBe(false);
    expect(wrapper.find('[role="treeitem"][aria-selected="true"]').text()).toContain('אחרונים');
  });

  it('מקום אחר מחליף את הרשימה, והאירועים עוברים הלאה', async () => {
    const { wrapper } = mountUi(OpenDocumentDialog, {
      props: { isOpen: true, place: { kind: 'library', path: '/הלכה' }, library: LIBRARY, libraryState: 'ready' },
    });
    await settle();
    expect(wrapper.find('.rec-section').exists()).toBe(false);
    await wrapper.findAll('.sl-open')[1]!.trigger('click');
    expect(wrapper.emitted('open-library-book')?.[0]?.[0]).toMatchObject({ title: 'קונטרס' });
  });

  it('חיפוש ב„אחרונים” מציע גם ספרים מהספרייה, ו„הצג הכול” פולט את השאילתה', async () => {
    const { wrapper } = mountUi(OpenDocumentDialog, {
      props: { isOpen: true, recents: [], searchQuery: 'קונ', library: LIBRARY, libraryState: 'ready' },
    });
    expect(wrapper.findAll('.rec-also__name').map((el) => el.text())).toEqual(['קונטרס']);
    await wrapper.find('.rec-also__all').trigger('click');
    expect(wrapper.emitted('search-library')?.[0]).toEqual(['קונ']);
  });

  it('המדף של הספרים האישיים אינו מופיע פעמיים — רק כצומת „ספרים אישיים”', () => {
    const { wrapper } = mountUi(OpenDocumentDialog, {
      props: { isOpen: true, library: LIBRARY, libraryState: 'ready' },
    });
    const labels = wrapper.findAll('[role="treeitem"] .src-label').map((el) => el.text());
    expect(labels).toEqual(['אחרונים', 'ספריית אוצריא', 'הלכה', 'ספרים אישיים']);
  });

  it('הודעה מוצגת בתחתית כשאין פתיחה בדרך', () => {
    const { wrapper } = mountUi(OpenDocumentDialog, { props: { isOpen: true, notice: 'הספר לא נפתח' } });
    expect(wrapper.find('.open-status').text()).toBe('הספר לא נפתח');
  });
});

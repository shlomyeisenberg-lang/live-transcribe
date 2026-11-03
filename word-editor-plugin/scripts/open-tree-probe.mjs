/**
 * שער העץ של „פתח מסמך” — בדפדפן אמיתי, על ה-dist הארוז.
 *
 * מה שבדיקת רכיב אינה יכולה לראות: שהעמודה והרשימה יושבות זו לצד זו (העץ
 * מימין), שהשורות אינן נחתכות ואינן גולשות, שבחלון צר העץ עובר מעל הרשימה,
 * ושהדיאלוג נפתח **מלא מהמטמון** — לפני שאוצריא ענתה על `library.getTree`.
 * המדידה האחרונה היא טענת המהירות, ולכן היא נבדקת: ה-stub כאן מעכב את
 * התשובה בשתי שניות, והעץ חייב להופיע לפניהן.
 *
 *   npm run build && node scripts/open-tree-probe.mjs
 *
 * צילומים ב-tmp/open-tree-*.png.
 */
import { readFileSync, writeFileSync, existsSync, mkdirSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { openPage, requireChrome, sleep } from './cdp.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const DIST = join(ROOT, 'dist');
const TMP = join(ROOT, 'tmp');
const READY_MS = 40_000;

if (!existsSync(join(DIST, 'index.html'))) {
  console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}
requireChrome();
mkdirSync(TMP, { recursive: true });

const TREE = {
  title: 'ספריית אוצריא',
  path: '/',
  books: [],
  categories: [
    {
      title: 'הלכה',
      path: '/הלכה',
      books: [
        { bookId: 'קונטרס בהלכות שבת', title: 'קונטרס בהלכות שבת', type: 'docx', bookUid: 'uid:1', source: 'library', author: 'ר׳ פלוני אלמוני' },
        { bookId: 'משנה ברורה', title: 'משנה ברורה', type: 'text' },
      ],
      categories: [
        {
          title: 'שו״ת',
          path: '/הלכה/שו״ת',
          books: [
            { bookId: 'שו״ת דברי שלום', title: 'שו״ת דברי שלום', type: 'docx', bookUid: 'uid:2', source: 'library' },
            { bookId: 'שו״ת אמרי יושר — חלק שני, עם הגהות והערות מכתב יד', title: 'שו״ת אמרי יושר — חלק שני, עם הגהות והערות מכתב יד', type: 'docx', bookUid: 'uid:4', source: 'library' },
          ],
          categories: [],
        },
      ],
    },
    {
      title: 'ספרים אישיים',
      path: '/ספרים אישיים',
      books: [{ bookId: 'חידושי תורה שלי', title: 'חידושי תורה שלי', type: 'docx', bookUid: 'uid:3', source: 'user' }],
      categories: [],
    },
    { title: 'תנך', path: '/תנך', books: [{ bookId: 'בראשית', title: 'בראשית', type: 'text' }], categories: [] },
  ],
};

const now = Date.now();
const LISTING = {
  entries: [
    { name: 'שיעורים', path: 'שיעורים', type: 'dir', size: 0, modified: null },
    { name: 'חידושים על הדף — מסכת ברכות.docx', path: 'חידושים על הדף — מסכת ברכות.docx', type: 'file', size: 48_200, modified: new Date(now - 3_600_000).toISOString() },
    { name: 'Letter_to_the_printer_final_v2.docx', path: 'Letter_to_the_printer_final_v2.docx', type: 'file', size: 18_004, modified: new Date(now - 12 * 86_400_000).toISOString() },
  ],
  truncated: false,
};

const RECENTS = [
  { token: 't1', name: 'שולחן ערוך — אורח חיים.docx', size: 1_482_000, openedAt: now - 3_600_000, pinned: true, writable: true },
  { token: 't2', name: 'בראשית — הגהות.docx', size: 24_500, openedAt: now - 7_200_000, pinned: false, writable: true },
];

/**
 * `cachedTree` — מה שנשמר בפעם הקודמת. `getTree` מעוכב בשתי שניות, כדי
 * שאפשר יהיה לראות אם העץ הגיע מהמטמון או מהתשובה.
 */
const stubFor = ({ place, cachedTree }) => `<script>
  window.__otzariaStorage = {
    'recent-documents': ${JSON.stringify(RECENTS)},
    'open-dialog-folders': [{ token: 'ft', name: 'העבודות שלי', path: 'C:\\\\Users\\\\me\\\\Documents\\\\העבודות שלי', addedAt: 1 }],
    'open-dialog-place': ${JSON.stringify(place)},
    'open-dialog-library': ${cachedTree ? JSON.stringify({ version: 1, savedAt: 1, tree: cachedTree }) : 'null'}
  };
  window.__treeCalls = 0;
  const ok = (data) => Promise.resolve({ success: true, data, error: null });
  window.Otzaria = {
    call: function (method, params) {
      if (method === 'app.getInfo') return ok({ version: '0.9.98', platform: 'p' });
      if (method === 'app.getTheme') return ok({ mode: 'light', colorScheme: {}, typography: {} });
      if (method === 'storage.get') return ok(window.__otzariaStorage[params && params.key] || null);
      if (method === 'storage.set') return ok(true);
      if (method === 'library.getTree') {
        window.__treeCalls += 1;
        return new Promise((r) => setTimeout(() => r({ success: true, data: ${JSON.stringify(TREE)}, error: null }), 2000));
      }
      if (method === 'fs.listUserFolder') return ok(${JSON.stringify(LISTING)});
      return Promise.resolve({ success: false, data: null, error: { code: 'error.unknown_method', message: 'no' } });
    },
    on: function () {},
    off: function () {}
  };
</script>`;

const html = readFileSync(join(DIST, 'index.html'), 'utf8');
const latchEnd = html.indexOf('</script>') + '</script>'.length;

function writePage(name, options) {
  const path = join(DIST, name);
  writeFileSync(path, html.slice(0, latchEnd) + stubFor(options) + html.slice(latchEnd));
  return path;
}

const failures = [];
function check(ok, message) {
  console.log(`${ok ? '  ✓' : '  ✗'} ${message}`);
  if (!ok) failures.push(message);
}

// המטמון הוא העץ הגזום — בדיוק מה ש-`saveLibraryCache` כותב. נגזר מ-TREE
// בלי הספרים שאינם Word, כמו ש-`pruneLibraryTree` היה עושה.
const CACHED = {
  title: 'ספריית אוצריא', path: '/', total: 4, books: [],
  shelves: [
    {
      title: 'הלכה', path: '/הלכה', total: 3,
      books: [{ key: 'uid:1', title: 'קונטרס בהלכות שבת', type: 'docx', bookId: 'קונטרס בהלכות שבת', bookUid: 'uid:1', id: null, source: 'library', author: 'ר׳ פלוני אלמוני' }],
      shelves: [{
        title: 'שו״ת', path: '/הלכה/שו״ת', total: 2, shelves: [],
        books: [
          { key: 'uid:2', title: 'שו״ת דברי שלום', type: 'docx', bookId: 'שו״ת דברי שלום', bookUid: 'uid:2', id: null, source: 'library', author: '' },
          { key: 'uid:4', title: 'שו״ת אמרי יושר — חלק שני, עם הגהות והערות מכתב יד', type: 'docx', bookId: 'x', bookUid: 'uid:4', id: null, source: 'library', author: '' },
        ],
      }],
    },
    { title: 'ספרים אישיים', path: '/ספרים אישיים', total: 1, shelves: [], books: [{ key: 'uid:3', title: 'חידושי תורה שלי', type: 'docx', bookId: 'חידושי תורה שלי', bookUid: 'uid:3', id: null, source: 'user', author: '' }] },
  ],
};

const SCENARIOS = [
  { key: 'recent', label: 'אחרונים, עם מטמון', place: 'recent', cachedTree: CACHED, width: 1280, height: 860 },
  { key: 'shelf', label: 'מדף בספרייה', place: 'library\u0000/הלכה', cachedTree: CACHED, width: 1280, height: 860 },
  { key: 'folder', label: 'תיקייה', place: 'folder\u0000ft\u0000', cachedTree: CACHED, width: 1280, height: 860 },
  { key: 'nocache', label: 'בלי מטמון — טעינה ראשונה', place: 'library\u0000/', cachedTree: null, width: 1280, height: 860 },
  { key: 'narrow', label: 'חלון צר', place: 'library\u0000/הלכה', cachedTree: CACHED, width: 560, height: 860 },
];

const written = [];
for (const scenario of SCENARIOS) {
  console.log(`\n════ ${scenario.label} ════`);
  const path = writePage(`open-tree-${scenario.key}-tmp.html`, scenario);
  written.push(path);
  const page = await openPage(pathToFileURL(path).href, { label: `open-tree-${scenario.key}` });
  try {
    await page.cdp.send('Emulation.setDeviceMetricsOverride', {
      width: scenario.width, height: scenario.height, deviceScaleFactor: 1, mobile: false,
    });
    const deadline = Date.now() + READY_MS;
    while (!(await page.cdp.evaluate('!!window.__otzariaEditor && !document.getElementById("otzaria-splash")'))) {
      if (Date.now() > deadline) throw new Error('העלייה לא הסתיימה');
      await sleep(250);
    }
    // Ctrl+O עד שהדיאלוג נפתח — ראו open-dialog-probe.mjs על המשמר של הפתיחה.
    const openedAt = Date.now();
    for (;;) {
      await page.cdp.evaluate("window.dispatchEvent(new KeyboardEvent('keydown', { key: 'o', code: 'KeyO', ctrlKey: true, bubbles: true }))");
      await sleep(150);
      if (await page.cdp.evaluate("!!document.querySelector('.open-dialog')")) break;
      if (Date.now() - openedAt > READY_MS) throw new Error('הדיאלוג לא נפתח');
    }
    const shownAt = Date.now();

    const early = await page.cdp.evaluate(`(() => {
      const labels = [...document.querySelectorAll('.src-item .src-label')].map((e) => e.textContent.trim());
      return { labels, calls: window.__treeCalls };
    })()`);
    if (scenario.cachedTree) {
      check(early.labels.includes('הלכה'), `העץ מלא מיד מהמטמון, לפני תשובת אוצריא (${early.labels.join(' | ')})`);
    } else {
      check(await page.cdp.evaluate("!!document.querySelector('.sl-spinner')"), 'בלי מטמון: „טוען…” עד שהתשובה מגיעה');
    }

    await sleep(Math.max(0, 2300 - (Date.now() - shownAt)));

    const m = await page.cdp.evaluate(`(() => {
      const r = (sel) => { const e = document.querySelector(sel); if (!e) return null; const b = e.getBoundingClientRect(); return { x: b.x, y: b.y, w: b.width, h: b.height, r: b.right, b: b.bottom }; };
      const main = r('.rec-section') || r('.sl');
      const overflowing = [...document.querySelectorAll('.src-label, .sl-name, .rec-name')].filter((e) => e.scrollWidth > e.clientWidth + 1 && getComputedStyle(e).textOverflow !== 'ellipsis').length;
      const dialog = document.querySelector('.open-dialog');
      return {
        nav: r('.src-nav'), main, dialog: r('.open-dialog'),
        pageScrollX: document.documentElement.scrollWidth > window.innerWidth,
        dialogOverflowX: dialog.scrollWidth > dialog.clientWidth + 1,
        overflowing,
        labels: [...document.querySelectorAll('.src-item .src-label')].map((e) => e.textContent.trim()),
        rows: [...document.querySelectorAll('.sl-name, .rec-name')].map((e) => e.textContent.trim()),
        selected: document.querySelector('.src-item--selected .src-label')?.textContent.trim() ?? null,
      };
    })()`);

    check(m.nav && m.main, 'יש עמודת עץ ורשימה');
    if (scenario.width >= 700) {
      check(m.nav.x > m.main.x && Math.abs(m.nav.y - m.main.y) < 2, `העץ מימין לרשימה, באותו גובה (עץ x=${Math.round(m.nav.x)}, רשימה x=${Math.round(m.main.x)})`);
      check(Math.abs(m.nav.b - m.main.b) < 2, `העץ והרשימה נגמרים באותו קו (${Math.round(m.nav.b)} / ${Math.round(m.main.b)})`);
    } else {
      check(m.nav.b <= m.main.y + 1, `בחלון צר העץ מעל הרשימה (${Math.round(m.nav.b)} ≤ ${Math.round(m.main.y)})`);
    }
    check(!m.pageScrollX && !m.dialogOverflowX, 'אין גלילה אופקית בדף ובדיאלוג');
    check(m.overflowing === 0, `אין שם שגולש בלי קיצוץ (${m.overflowing})`);
    console.log(`  נבחר: ${m.selected}; שורות: ${m.rows.slice(0, 5).join(' | ')}`);

    const shot = await page.cdp.send('Page.captureScreenshot', { format: 'png' });
    writeFileSync(join(TMP, `open-tree-${scenario.key}.png`), Buffer.from(shot.result.data, 'base64'));
    console.log(`  📸 tmp/open-tree-${scenario.key}.png`);
  } finally {
    await page.close();
  }
}

for (const path of written) rmSync(path, { force: true });
if (failures.length > 0) {
  console.error(`\n✗ ${failures.length} כשלים`);
  process.exit(1);
}
console.log('\n✓ העץ של „פתח מסמך” תקין');

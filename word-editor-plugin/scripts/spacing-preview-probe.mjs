/**
 * גשש: האם אפשר תצוגה מקדימה **אנכית** של ריווח פסקה בתוך הדף, בלי לגעת במסמך.
 *
 * הרקע: docs/engine-gaps.md, „תצוגה חיה של גופן”. שם נמדד שספק הדקורציות של
 * superdoc 2 הוא render-only ובטוח לחלוטין (undoDepth 0→0, ה-docx זהה
 * בייט-בבייט), ובכל זאת נשלל — כי המנוע מצייר במיקום מוחלט ואינו שובר שורות
 * מחדש, והטקסט נחתך. אותה מדידה הייתה **אופקית**. כאן נמדד הכיוון האנכי, שאינו
 * נגזר ממנה: מרווח שורות בתוך פסקה, וריווח לפני/אחרי פסקה.
 *
 * מה שנמדד כאן הוא **המכניקה של הפריסה**, ולא ידית המסירה. הידית כבר נמדדה
 * (ראו למעלה) והיא „מדביקה class”, כלומר מפל CSS. הגשש מזריק את אותו מפל דרך
 * גיליון סגנון גלובלי — המסלול שנמדד שם כשקול — ומודד פיקסלים. ה-dist הארוז
 * אינו רושם `extensions` בבנאי (src/engine/create-editor.ts), ולכן רישום
 * דקורציה אמיתי היה מחייב שינוי מקור ובנייה מחדש; זה לא היה משנה את הגאומטריה,
 * שהיא כל מה שנשאל כאן.
 *
 * הסלקטור הוא `[data-source-node-id]` — תכונה פנימית של המנוע. זה מותר בגשש
 * מדידה ואסור במוצר (ראו ההערה על `.superdoc-text-run` ב-engine-gaps).
 *
 * **מדידת החפיפה היא לפי תחתית השורה האחרונה ולא לפי תחתית הפסקה**, וזה לא
 * פרט: תיבת הפסקה היא `position: absolute` עם `height` קשיח, והיא **אינה
 * גדלה** כשהשורות שבתוכה גדלות. מדידה מול `frag.bottom` מדווחת „אין חפיפה” על
 * חפיפה של עשרות פיקסלים. הגרסה הראשונה של הגשש הזה נפלה בדיוק כך.
 *
 *   npm run build && node scripts/spacing-preview-probe.mjs
 */
import { readFileSync, writeFileSync, existsSync, rmSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, requireChrome, sleep } from './cdp.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const DIST = join(ROOT, 'dist');
const TMP = join(ROOT, 'tmp');
const OBSERVE_MS = 60_000;
const VIEWPORT = { width: 1440, height: 900 };

/** הפסקה שנמדדת: מספיק פנימה כדי שיהיו שכנות משני הצדדים. */
const TARGET_INDEX = 5;

/** 480 twips = 24pt = 32px בדיוק. גדול מספיק כדי שהפרש ייראה, ועגול. */
const EXACT_TWIPS = 480;
const EXACT_PX = 32;

if (!existsSync(join(DIST, 'index.html'))) {
  console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}
requireChrome();
mkdirSync(TMP, { recursive: true });

const STUB = `<script>
  window.Otzaria = {
    call: function (method) {
      if (method === 'app.getInfo') return Promise.resolve({ success: true, data: { version: '9', platform: 'p' }, error: null });
      if (method === 'app.getTheme') return Promise.resolve({ success: true, data: { mode: 'light', colorScheme: {}, typography: {} }, error: null });
      return Promise.resolve({ success: false, data: null, error: { message: 'no' } });
    },
    on: function () {},
    off: function () {}
  };
</script>`;

const html = readFileSync(join(DIST, 'index.html'), 'utf8');
const latchEnd = html.indexOf('</script>') + '</script>'.length;
const path = join(DIST, 'spacing-preview-tmp.html');
writeFileSync(path, html.slice(0, latchEnd) + STUB + html.slice(latchEnd));

const READY = `(async () => {
  const settle = (ms) => new Promise((r) => setTimeout(r, ms));
  const deadline = Date.now() + ${OBSERVE_MS};
  while (Date.now() < deadline) {
    if (document.querySelector('.word-tab-strip') && window.__otzariaEditor) break;
    await settle(250);
  }
  const sd = window.__otzariaEditor && window.__otzariaEditor.superdoc;
  if (!sd) return false;
  while (Date.now() < deadline) {
    if (sd.ui && sd.ui.commands.get('zoom').getState().enabled) break;
    await settle(250);
  }
  await settle(2000);
  return true;
})()`;

/** 40 פסקאות בנות שלוש שורות — קצת יותר משני עמודים, כדי שיהיה גם קצה עמוד. */
const FILL = `(async () => {
  const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
  const body = 'טקסט מילוי ארוך כדי שהפסקה תתפרס על כמה שורות ונוכל למדוד את המרחק בין השורות בתוך הפסקה עצמה ולא רק בין פסקאות. ';
  const ps = [];
  for (let i = 0; i < 40; i++) ps.push('<p>פסקה ' + (i + 1) + ' — ' + body + body + '</p>');
  await doc.insert({ value: ps.join(''), type: 'html' });
  await new Promise((r) => setTimeout(r, 6000));
  return document.querySelectorAll('.superdoc-page').length;
})()`;

const TOOLS = `(() => {
  const q = (s, r) => Array.from((r || document).querySelectorAll(s));
  const n1 = (v) => Math.round(v * 10) / 10;
  const rect = (el) => { const r = el.getBoundingClientRect(); return {
    top: n1(r.top), bottom: n1(r.bottom), left: n1(r.left), right: n1(r.right), h: n1(r.height), w: n1(r.width) }; };

  window.__sp = {
    fragments: () => q('.superdoc-fragment').map((f, i) => ({
      i, id: f.dataset.sourceNodeId,
      page: f.closest('.superdoc-page') ? Number(f.closest('.superdoc-page').dataset.pageIndex) : -1,
      cssTop: getComputedStyle(f).top, pos: getComputedStyle(f).position,
      lines: f.querySelectorAll('.superdoc-line').length,
      text: (f.textContent || '').slice(0, 18),
    })),
    snap: (id) => {
      const f = document.querySelector('.superdoc-fragment[data-source-node-id="' + id + '"]');
      if (!f) return null;
      const page = f.closest('.superdoc-page');
      const cs = getComputedStyle(f);
      const lines = q('.superdoc-line', f).map((l) => {
        const lcs = getComputedStyle(l);
        return { ...rect(l), lh: lcs.lineHeight, cssH: lcs.height, pos: lcs.position, inline: l.getAttribute('style') };
      });
      return {
        id, epoch: f.dataset.layoutEpoch,
        frag: { ...rect(f), cssTop: cs.top, pos: cs.position, mt: cs.marginTop, pt: cs.paddingTop, cssH: cs.height, overflow: cs.overflow },
        lines,
        /** מה שהעין רואה: הקצה התחתון של הטקסט, ולא של התיבה. */
        inkTop: lines.length ? Math.min(...lines.map((l) => l.top)) : null,
        inkBottom: lines.length ? Math.max(...lines.map((l) => l.bottom)) : null,
        page: page ? { ...rect(page), index: Number(page.dataset.pageIndex), overflow: getComputedStyle(page).overflow } : null,
        pageCount: q('.superdoc-page').length,
      };
    },
    /** גבול תיבת הטקסט: הראשונה והאחרונה מבין הפסקאות שבאותו עמוד, בבסיס. */
    textBox: (pageIndex) => {
      const page = q('.superdoc-page')[pageIndex];
      if (!page) return null;
      const frs = q('.superdoc-fragment', page);
      if (!frs.length) return null;
      const pr = page.getBoundingClientRect();
      const tops = frs.map((f) => f.getBoundingClientRect().top);
      const bottoms = frs.map((f) => f.getBoundingClientRect().bottom);
      return { pageTop: n1(pr.top), pageBottom: n1(pr.bottom), first: n1(Math.min(...tops)), last: n1(Math.max(...bottoms)) };
    },
    css: (text) => {
      let el = document.getElementById('__sp_style');
      if (!el) { el = document.createElement('style'); el.id = '__sp_style'; document.head.appendChild(el); }
      el.textContent = text || '';
      return el.textContent.length;
    },
    /** כל הצאצאים של העמוד, לפי tag.class — כדי לגלות מה נולד עם הבחירה. */
    census: () => {
      const map = {};
      /* כל הדף ולא רק תחת .superdoc-page: נמדד ששכבת הבחירה יושבת **מחוץ**
       * לעמוד, כ-div בלי class, במיכל משלה. מפקד שמוגבל לעמוד מחמיץ אותה
       * לגמרי ומדווח „לא נמצאו מלבני בחירה”. */
      for (const el of q('body *')) {
        const c = typeof el.className === 'string' ? el.className.trim().split(/\\s+/).join('.') : '';
        const key = el.tagName.toLowerCase() + (c ? '.' + c : '');
        map[key] = (map[key] || 0) + 1;
      }
      return map;
    },
    /** מלבנים של אלמנטים שנוספו בבחירה, לפי רשימת מפתחות. */
    rectsOf: (keys) => {
      const out = [];
      /* כל הדף ולא רק תחת .superdoc-page: נמדד ששכבת הבחירה יושבת **מחוץ**
       * לעמוד, כ-div בלי class, במיכל משלה. מפקד שמוגבל לעמוד מחמיץ אותה
       * לגמרי ומדווח „לא נמצאו מלבני בחירה”. */
      for (const el of q('body *')) {
        const c = typeof el.className === 'string' ? el.className.trim().split(/\\s+/).join('.') : '';
        const key = el.tagName.toLowerCase() + (c ? '.' + c : '');
        if (!keys.includes(key)) continue;
        const r = el.getBoundingClientRect();
        if (r.width < 0.5 && r.height < 0.5) continue;
        out.push({ key, ...rect(el) });
      }
      return out;
    },
  };
  return true;
})()`;

const round = (v) => Math.round(v * 10) / 10;

function steps(snap) {
  const tops = snap.lines.map((l) => l.top);
  return tops.slice(1).map((t, i) => round(t - tops[i]));
}

async function screenshot(cdp, file) {
  const response = await cdp.send('Page.captureScreenshot', { format: 'png' });
  const data = response?.result?.data;
  if (!data) return;
  writeFileSync(file, Buffer.from(data, 'base64'));
  console.log(`📷 ${file}`);
}

const report = [];
const say = (line) => { console.log(line); report.push(line); };

let page;
try {
  page = await openPage(`file://${path}`, { label: 'spacing-preview' });
  const { cdp } = page;
  const js = (e) => cdp.evaluate(e);
  await cdp.send('Emulation.setDeviceMetricsOverride', { ...VIEWPORT, deviceScaleFactor: 1, mobile: false });
  if (!(await js(READY))) throw new Error('האתחול לא הושלם');

  const pageCount = await js(FILL);
  await js(TOOLS);
  const frags = await js('window.__sp.fragments()');
  say(`\n== המסמך: ${pageCount} עמודים, ${frags.length} פסקאות מצוירות, חלון ${VIEWPORT.width}×${VIEWPORT.height}, זום 100% ==`);

  const target = frags[TARGET_INDEX];
  const next = frags[TARGET_INDEX + 1];
  const base = await js(`window.__sp.snap("${target.id}")`);
  const baseNext = await js(`window.__sp.snap("${next.id}")`);
  const box = await js('window.__sp.textBox(0)');

  say(`\n-- בסיס --`);
  say(`  פסקת היעד #${target.i} id=${target.id} "${target.text}" — ${target.lines} שורות`);
  say(`  תיבת הפסקה: position=${base.frag.pos} cssTop=${base.frag.cssTop} height=${base.frag.cssH} overflow=${base.frag.overflow}`);
  say(`  תיבת השורה (inline): ${base.lines[0].inline}`);
  say(`  ראשי שורות: ${base.lines.map((l) => l.top).join(' / ')}  → פסיעה ${steps(base).join(' / ')}px`);
  say(`  דיו הפסקה: ${base.inkTop}..${base.inkBottom}   תיבת הפסקה: ${base.frag.top}..${base.frag.bottom}`);
  say(`  ראש הפסקה הבאה: ${baseNext.frag.top} (דיו ${baseNext.inkTop})`);
  say(`  העמוד: ${base.page.top}..${base.page.bottom} overflow=${base.page.overflow}`);
  say(`  תיבת הטקסט בעמוד 1: ${box.first}..${box.last} (שוליים תחתונים ≈ ${round(box.pageBottom - box.last)}px)`);
  await screenshot(cdp, join(TMP, 'spacing-preview-base.png'));

  /* ---------------------------------------------------------------- */
  const sel = `.superdoc-fragment[data-source-node-id="${target.id}"]`;
  const cases = [
    { name: `א. line-height ${EXACT_PX}px בלבד על השורות`,
      css: `${sel} .superdoc-line { line-height: ${EXACT_PX}px !important; }` },
    { name: `ב. line-height ${EXACT_PX}px + height:auto`,
      css: `${sel} .superdoc-line { line-height: ${EXACT_PX}px !important; height: auto !important; }` },
    { name: `ג. line-height + height ${EXACT_PX}px (שליטה מדויקת)`,
      css: `${sel} .superdoc-line { line-height: ${EXACT_PX}px !important; height: ${EXACT_PX}px !important; }` },
    { name: `ד. margin-top 40px על הפסקה (ריווח „לפני”)`,
      css: `${sel} { margin-top: 40px !important; }` },
    { name: `ה. padding-top 40px על הפסקה`,
      css: `${sel} { padding-top: 40px !important; }` },
    { name: `ו. transform: translateY(40px) על הפסקה`,
      css: `${sel} { transform: translateY(40px) !important; }` },
    { name: `ז. margin-bottom 40px על הפסקה (ריווח „אחרי”)`,
      css: `${sel} { margin-bottom: 40px !important; }` },
  ];

  for (const c of cases) {
    await js(`window.__sp.css(${JSON.stringify(c.css)})`);
    await sleep(700);
    const s = await js(`window.__sp.snap("${target.id}")`);
    const n = await js(`window.__sp.snap("${next.id}")`);
    /* החפיפה האמיתית: תחתית הדיו שלנו מול ראש הדיו של השכנה. */
    const overlap = round(s.inkBottom - n.inkTop);
    say(`\n-- ${c.name} --`);
    say(`  פסיעה בין שורות: ${steps(s).join(' / ')}px   (בסיס ${steps(base).join(' / ')})`);
    say(`  דיו הפסקה: ${s.inkTop}..${s.inkBottom}   (בסיס ${base.inkTop}..${base.inkBottom}, גדל ב-${round(s.inkBottom - base.inkBottom)}px)`);
    say(`  תיבת הפסקה: ${s.frag.top}..${s.frag.bottom} h=${s.frag.h} (בסיס h=${base.frag.h}) — גדלה ב-${round(s.frag.h - base.frag.h)}px`);
    say(`  הפסקה הבאה זזה ב-${round(n.frag.top - baseNext.frag.top)}px (דיו ${n.inkTop}, בסיס ${baseNext.inkTop})`);
    say(`  ${overlap > 0 ? `**חפיפה ${overlap}px** — הטקסט שלנו נוחת על השכנה` : 'אין חפיפה'}`);
    say(`  עמודים: ${s.pageCount} (בסיס ${base.pageCount}), layout-epoch=${s.epoch} (בסיס ${base.epoch})`);
  }
  await js('window.__sp.css("")');
  await sleep(400);

  /* ---------------------------------------------------------------- */
  /* גלישה מהעמוד: אותו כלל על הפסקה האחרונה שבעמוד הראשון. */
  const lastOnFirst = [...frags].reverse().find((f) => f.page === 0);
  if (lastOnFirst) {
    const lsel = `.superdoc-fragment[data-source-node-id="${lastOnFirst.id}"]`;
    const bBefore = await js(`window.__sp.snap("${lastOnFirst.id}")`);
    say(`\n-- גלישה מקצה העמוד (הפסקה האחרונה בעמוד 1, #${lastOnFirst.i}, ${bBefore.lines.length} שורות) --`);
    say(`  לפני: דיו ${bBefore.inkTop}..${bBefore.inkBottom}; תחתית תיבת הטקסט ≈ ${box.last}; תחתית הנייר ${bBefore.page.bottom}`);
    for (const px of [EXACT_PX, 48, 96]) {
      await js(`window.__sp.css(${JSON.stringify(`${lsel} .superdoc-line { line-height: ${px}px !important; height: ${px}px !important; }`)})`);
      await sleep(900);
      const a = await js(`window.__sp.snap("${lastOnFirst.id}")`);
      say(`  שורה ${px}px: דיו עד ${a.inkBottom} — ` +
          `${round(Math.max(0, a.inkBottom - box.last))}px מעבר לתיבת הטקסט, ` +
          `${round(Math.max(0, a.inkBottom - a.page.bottom))}px מעבר לקצה הנייר; ` +
          `עמודים ${a.pageCount}`);
      if (px === 96) await screenshot(cdp, join(TMP, 'spacing-preview-overflow.png'));
    }
    await js('window.__sp.css("")');
    await sleep(400);
  }

  /* ---------------------------------------------------------------- */
  /* הבחירה: גוררים על שתי שורות בפסקת היעד, ואז מחילים את הכלל. */
  const censusBefore = await js('window.__sp.census()');
  const l0 = base.lines[0];
  const l1 = base.lines[1];
  for (const [type, x, y] of [
    ['mousePressed', l0.right - 8, l0.top + l0.h / 2],
    ['mouseMoved', (l0.right + l0.left) / 2, l0.top + l0.h / 2],
    ['mouseMoved', l1.left + 40, l1.top + l1.h / 2],
    ['mouseReleased', l1.left + 40, l1.top + l1.h / 2],
  ]) {
    await cdp.send('Input.dispatchMouseEvent', {
      type, x, y, button: 'left', clickCount: 1, buttons: type === 'mouseReleased' ? 0 : 1,
    });
    await sleep(150);
  }
  await sleep(1200);
  const censusAfter = await js('window.__sp.census()');
  const born = Object.keys(censusAfter).filter((k) => (censusAfter[k] || 0) > (censusBefore[k] || 0));
  const selSnapshot = await js(
    `(function(){try{var s=window.__otzariaEditor.superdoc.ui.selection.getSnapshot();return s.status+'/'+(s.empty?'empty':'range');}catch(e){return 'ERR:'+e.message;}})()`,
  );
  say(`\n-- פס הבחירה --`);
  say(`  מצב הבחירה במנוע: ${selSnapshot}`);
  say(`  אלמנטים שנולדו עם הבחירה: ${born.length ? born.join(', ') : 'אף אחד'}`);
  if (born.length) {
    const rBefore = await js(`window.__sp.rectsOf(${JSON.stringify(born)})`);
    await js(`window.__sp.css(${JSON.stringify(`${sel} .superdoc-line { line-height: ${EXACT_PX}px !important; height: ${EXACT_PX}px !important; }`)})`);
    await sleep(900);
    const rAfter = await js(`window.__sp.rectsOf(${JSON.stringify(born)})`);
    const sAfter = await js(`window.__sp.snap("${target.id}")`);
    say(`  מלבני בחירה לפני: ${rBefore.map((r) => `${r.key}@${r.top}..${r.bottom}`).join(' | ')}`);
    say(`  מלבני בחירה אחרי: ${rAfter.map((r) => `${r.key}@${r.top}..${r.bottom}`).join(' | ')}`);
    say(`  ראשי השורות אחרי: ${sAfter.lines.map((l) => l.top).join(' / ')} (בסיס ${base.lines.map((l) => l.top).join(' / ')})`);
    await screenshot(cdp, join(TMP, 'spacing-preview-selection.png'));
    await js('window.__sp.css("")');
    await sleep(500);
  }

  /* ---------------------------------------------------------------- */
  /*
   * פיצוי בקסקייד: אם השכנות אינן זזות מעצמן — נזיז אותן בכלל CSS.
   *
   * הפסקאות הן אחיות בתוך `.superdoc-page`, ולכן `~` מגיע בדיוק לכל מה
   * שאחרי היעד **באותו עמוד** ולא מעבר לו. זה המבחן שמפריד „לא ניתן” מ„ניתן
   * עם חור”: ריווח אנכי, בשונה מהחלפת גופן, אינו משנה את **שבירת השורות** —
   * רוחב השורה זהה — ולכן בתוך הפסקה התצוגה יכולה להיות מדויקת לפיקסל.
   */
  /*
   * הקבוע: `שורות × Δ`, ולא `(שורות−1) × Δ`.
   *
   * תיבת הפסקה גדלה בגובה של **כל** שורה ולא במרווחים שביניהן — נמדד באמת:
   * h 55.2 → 96, כלומר 40.8 = 3 × 13.6. הריצה הראשונה של הגשש חישבה 27.2
   * ונשארה עם 13.6px חפיפה; זו שגיאת חשבון, לא גבול של המנגנון.
   */
  const shift = round((EXACT_PX - 18.4) * base.lines.length);
  await js(
    `window.__sp.css(${JSON.stringify(
      `${sel} .superdoc-line { line-height: ${EXACT_PX}px !important; height: ${EXACT_PX}px !important; }` +
        `${sel} ~ .superdoc-fragment { transform: translateY(${shift}px) !important; }`,
    )})`,
  );
  await sleep(900);
  {
    const s = await js(`window.__sp.snap("${target.id}")`);
    const n = await js(`window.__sp.snap("${next.id}")`);
    const lastId = [...frags].reverse().find((f) => f.page === 0).id;
    const lastAfter = await js(`window.__sp.snap("${lastId}")`);
    const firstPage2 = frags.find((f) => f.page === 1);
    const p2 = firstPage2 ? await js(`window.__sp.snap("${firstPage2.id}")`) : null;
    say(`\n-- פיצוי בקסקייד: translateY(${shift}px) על כל האחיות שאחרי היעד --`);
    say(`  פסיעה בין שורות ביעד: ${steps(s).join(' / ')}px; דיו ${s.inkTop}..${s.inkBottom}`);
    say(`  הפסקה הבאה זזה ב-${round(n.inkTop - baseNext.inkTop)}px → דיו ${n.inkTop}`);
    say(`  חפיפה: ${round(s.inkBottom - n.inkTop) > 0 ? round(s.inkBottom - n.inkTop) + 'px' : 'אין'}`);
    say(`  הפסקה האחרונה בעמוד 1: דיו עד ${lastAfter.inkBottom} — ` +
        `${round(Math.max(0, lastAfter.inkBottom - box.last))}px מעבר לתיבת הטקסט, ` +
        `${round(Math.max(0, lastAfter.inkBottom - lastAfter.page.bottom))}px מעבר לקצה הנייר`);
    if (p2) say(`  הפסקה הראשונה בעמוד 2: דיו ${p2.inkTop} (הצירוף ~ אינו חוצה עמודים)`);
    say(`  עמודים: ${s.pageCount} (בסיס ${base.pageCount}), layout-epoch=${s.epoch} (בסיס ${base.epoch})`);
    await screenshot(cdp, join(TMP, 'spacing-preview-cascade.png'));
  }
  await js('window.__sp.css("")');
  await sleep(400);

  /* ---------------------------------------------------------------- */
  /* הרכב העמודים לפני האמת — כדי לראות אם האמת מעמדת מחדש. */
  const composeBefore = frags.reduce((m, f) => ((m[f.page] = (m[f.page] || 0) + 1), m), {});

  /* האמת: אותו ריווח דרך ה-API האמיתי. זה מה שהתצוגה אמורה לחזות. */
  const TRUTH = `(async () => {
    const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    const before = await doc.history.get();
    const setSpacing = doc.format && doc.format.paragraph && doc.format.paragraph.setSpacing;
    if (typeof setSpacing !== 'function') return { error: 'אין setSpacing' };
    const target = { kind: 'block', nodeType: 'paragraph', nodeId: "${target.id}" };
    let receipt;
    try { receipt = await setSpacing({ target, before: 0, after: 0, line: ${EXACT_TWIPS}, lineRule: 'exact' }); }
    catch (e) { receipt = { success: false, error: String(e && e.message) }; }
    await new Promise((r) => setTimeout(r, 4000));
    const after = await doc.history.get();
    return { receipt, undoBefore: before && before.undoDepth, undoAfter: after && after.undoDepth };
  })()`;
  const truth = await js(TRUTH);
  say(`\n-- האמת: setSpacing(line=${EXACT_TWIPS} twips = ${EXACT_PX}px, exact) דרך ה-API --`);
  say(`  קבלה: ${JSON.stringify(truth.receipt)}   undoDepth ${truth.undoBefore} → ${truth.undoAfter}`);
  const tFrags = await js('window.__sp.fragments()');
  const tIdx = tFrags.findIndex((f) => f.id === target.id);
  const tSnap = await js(`window.__sp.snap("${target.id}")`);
  const tNext = tIdx >= 0 && tFrags[tIdx + 1] ? await js(`window.__sp.snap("${tFrags[tIdx + 1].id}")`) : null;
  if (tSnap) {
    say(`  פסיעה בין שורות: ${steps(tSnap).join(' / ')}px   (בסיס ${steps(base).join(' / ')})`);
    say(`  דיו הפסקה: ${tSnap.inkTop}..${tSnap.inkBottom}   תיבת הפסקה: ${tSnap.frag.top}..${tSnap.frag.bottom} h=${tSnap.frag.h} (בסיס ${base.frag.h})`);
    say(`  שורה inline style: ${tSnap.lines[0].inline}`);
    if (tNext) say(`  הפסקה הבאה: ${tNext.frag.top} (בסיס ${baseNext.frag.top}) — זזה ב-${round(tNext.frag.top - baseNext.frag.top)}px`);
    say(`  עמודים: ${tSnap.pageCount} (בסיס ${base.pageCount}), layout-epoch=${tSnap.epoch} (בסיס ${base.epoch})`);
    const composeAfter = tFrags.reduce((m, f) => ((m[f.page] = (m[f.page] || 0) + 1), m), {});
    say(`  הרכב העמודים: ${JSON.stringify(composeBefore)} → ${JSON.stringify(composeAfter)}`);
  }
  await screenshot(cdp, join(TMP, 'spacing-preview-truth.png'));

  writeFileSync(join(TMP, 'spacing-preview-report.txt'), report.join('\n'), 'utf8');
  console.log(`\n📝 ${join(TMP, 'spacing-preview-report.txt')}`);
} catch (error) {
  console.error(`spacing-preview-probe נכשל: ${error.message}`);
  console.error(error.stack);
} finally {
  page?.close();
  rmSync(path, { force: true });
}

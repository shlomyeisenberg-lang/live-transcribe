/**
 * שער סקר: ניהול סגנונות מהגלריה — „שנה”, „צור”, „עדכן כך שיתאים לבחירה”,
 * שינוי שם (composables/use-style-manager.ts, engine/docx-style-edit.ts).
 *
 * המסלול שנבדק הוא המסלול של המשתמש: לחצן ימני על כרטיס בגלריה, הדיאלוג,
 * „אישור”, ואז עריכת הגדרות סגנון דרך templates.apply באותו טאב (`applyStyleEdit` ב-App.vue).
 * ההוכחות הן שלוש, ואף אחת אינה „הפקודה דיווחה הצלחה”:
 *  - הגאומטריה המצוירת (גודל מחושב, מיקום הטקסט בשורה ביחס לעמוד);
 *  - ה-OOXML: `styles.xml` ו-`document.xml` מהייצוא;
 *  - מצב המעטפת: זום, סמן, „לא נשמר”, היסטוריית ביטול.
 *
 * כל שלב רץ בדפדפן משלו: המנוע מתדרדר אחרי עשרות ייצואים באותו דף
 * (home-paragraph-qa.mjs), אין פתיחה מחדש בעת עריכת סגנון.
 *
 * `QA_SAVE_DIR` — כששמור, קבצי ה-docx שנשמרו בשמירה האמיתית (Ctrl+S) נכתבים
 * לשם, לבדיקה ב-Word.
 *
 * הרצה:  npm run build && node scripts/qa/style-manager-qa.mjs   (QA_PORT דורס 9731)
 * ‏`QA_PHASES=A,B` מריץ רק את השלבים האלה.
 */
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, createReport, unzip, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9731);
const SAVE_DIR = process.env.QA_SAVE_DIR ?? null;
const ONLY = process.env.QA_PHASES ? new Set(process.env.QA_PHASES.split(',')) : null;
const report = createReport('ניהול סגנונות מהגלריה');

const P1 = 'פסקה ראשונה של בדיקה';
const P2 = 'כותרת הבדיקה השנייה';
const P3 = 'פסקה שלישית של בדיקה';
const P2_LONG =
  'כותרת ארוכה מאוד שנועדה להיגלש לשתי שורות לפחות כדי שאפשר יהיה למדוד הזחה של שורה ראשונה מול השורות שאחריה בעמוד';

/** סדר הסכמה — העתק של RPR_ORDER/PPR_ORDER/STYLE_ORDER, כדי שהשער לא יסמוך על הקוד שהוא בודק. */
const STYLE_ORDER = ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority', 'semiHidden',
  'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose', 'personalReply', 'rsid',
  'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr'];
const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
  'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
  'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
  'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
  'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
  'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
  'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];

/* ------------------------------------------------------------------ */
/* OOXML                                                               */
/* ------------------------------------------------------------------ */

function styleXml(styles, id) {
  const re = new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`);
  return styles?.match(re)?.[0] ?? null;
}
function styleByName(styles, name) {
  return (styles?.match(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g) ?? []).find((s) => s.includes(`<w:name w:val="${name}"/>`)) ?? null;
}
const idOf = (style) => style?.match(/w:styleId="([^"]+)"/)?.[1] ?? null;

/** שמות הילדים הישירים של אלמנט (סורק רדוד, בלי תלות בקוד המוצר). */
function childNames(xml) {
  if (!xml) return [];
  const inner = xml.replace(/^<[^>]+>/, '').replace(/<\/[^>]+>$/, '');
  const out = [];
  let depth = 0;
  for (const m of inner.matchAll(/<(\/?)w:([A-Za-z]+)([^>]*?)(\/?)>/g)) {
    if (m[1] === '/') { depth--; continue; }
    if (depth === 0) out.push(m[2]);
    if (m[4] !== '/') depth++;
  }
  return out;
}
const sub = (xml, name) => xml?.match(new RegExp(`<w:${name}>[\\s\\S]*?</w:${name}>|<w:${name}/>`))?.[0] ?? null;

/** הילדים בסדר הסכמה? מחזירה null כשתקין, אחרת תיאור. */
function orderProblem(xml, order) {
  const names = childNames(xml);
  let last = -1;
  for (const n of names) {
    const rank = order.indexOf(n);
    if (rank < 0) return `ילד לא מוכר ${n} ב-[${names.join(',')}]`;
    if (rank < last) return `סדר שגוי: [${names.join(',')}]`;
    last = rank;
  }
  return null;
}
function schemaCheck(style) {
  if (!style) return 'אין סגנון';
  return orderProblem(style, STYLE_ORDER) ?? orderProblem(sub(style, 'pPr'), PPR_ORDER) ?? orderProblem(sub(style, 'rPr'), RPR_ORDER);
}

function paragraphsOf(files) {
  const xml = files?.['word/document.xml'] ?? '';
  return (xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []).map((p) => ({
    text: (p.match(/<w:t[^>]*>[^<]*<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, '')).join('').replace(/‏/g, ''),
    pPr: p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? '',
    pStyle: p.match(/<w:pStyle w:val="([^"]+)"/)?.[1] ?? null,
    jc: p.match(/<w:pPr>[\s\S]*?<w:jc w:val="([^"]+)"/)?.[1] ?? null,
    bidi: /<w:bidi\s*\/>|<w:bidi w:val="(1|true|on)"/.test(p),
    xml: p,
  }));
}
const paraBy = (files, text) => paragraphsOf(files).find((p) => p.text.includes(text));

/* ------------------------------------------------------------------ */
/* מדידה בדף                                                           */
/* ------------------------------------------------------------------ */

/**
 * הפסקאות המצוירות **הנראות**: החלה חיה עלולה להשאיר את המסמך הקודם ב-DOM
 * (נמדד בשערים אחרים), ולכן הסינון על `offsetParent`.
 */
const MEASURE = `(() => {
  const stack = document.querySelector('.editor-stack') || document;
  const seen = new Set();
  const out = [];
  for (const el of stack.querySelectorAll('.superdoc-fragment[data-source-node-id]')) {
    if (el.offsetParent === null || !el.getClientRects().length) continue;
    const id = el.getAttribute('data-source-node-id');
    if (seen.has(id)) continue;
    seen.add(id);
    const page = el.closest('.superdoc-page') || el.parentElement;
    const pr = page.getBoundingClientRect();
    const lines = [...el.querySelectorAll('.superdoc-line')].map((line) => {
      const lr = line.getBoundingClientRect();
      const runs = [...line.querySelectorAll('.superdoc-text-run')];
      let tl = Infinity, tr = -Infinity;
      for (const run of runs) {
        const range = document.createRange(); range.selectNodeContents(run);
        for (const r of range.getClientRects()) { if (r.width) { tl = Math.min(tl, r.left); tr = Math.max(tr, r.right); } }
      }
      return { lineL: Math.round(lr.left - pr.left), lineR: Math.round(pr.right - lr.right), lineW: Math.round(lr.width),
        textL: Math.round(tl - pr.left), textR: Math.round(pr.right - tr), gapL: Math.round(tl - lr.left), gapR: Math.round(lr.right - tr) };
    });
    const run = el.querySelector('.superdoc-text-run');
    const cs = run ? getComputedStyle(run) : null;
    const text = [...el.querySelectorAll('.superdoc-text-run')].map((r) => r.textContent || '').join('').replace(/\\u200f/g, '').trim();
    out.push({ id, text: text.slice(0, 30), fontSize: cs && cs.fontSize, fontWeight: cs && cs.fontWeight, color: cs && cs.color,
      fontFamily: cs && cs.fontFamily, dir: cs && cs.direction, pageW: Math.round(pr.width), scale: +(pr.width / page.offsetWidth).toFixed(3), lines });
  }
  return JSON.stringify(out);
})()`;

function classify(line) {
  if (!line) return null;
  const d = line.textL - line.textR;
  // מרחק מצדי העמוד; עמוד A4 ברוחב ~794px בזום 100%.
  if (Math.abs(d) <= 6) return 'center';
  return d > 0 ? 'right' : 'left';
}

/* ------------------------------------------------------------------ */
/* העזרים שמעל `app`                                                   */
/* ------------------------------------------------------------------ */

function ctxOf(app) {
  const js = app.js;
  const J = async (expr) => { const raw = await js(expr); return raw == null ? null : JSON.parse(raw); };

  const measure = () => J(MEASURE);
  const para = async (text) => (await measure()).find((p) => p.text.includes(text.slice(0, 12))) ?? null;

  /** מחכה עד ש-`fn()` מחזיר אמת, או עד הזמן. */
  async function until(fn, ms, every = 400) {
    const t0 = Date.now();
    while (Date.now() - t0 < ms) {
      const v = await fn();
      if (v) return { ok: true, ms: Date.now() - t0, value: v };
      await sleep(every);
    }
    return { ok: false, ms: Date.now() - t0 };
  }

  async function engineExport(label) {
    const timer = sleep(90_000).then(() => { throw new Error(`הייצוא נתקע (${label})`); });
    return Promise.race([app.docx(), timer]);
  }

  /** מסמן את מופע המנוע הנוכחי — מופע חדש = המסמך נפתח מחדש. */
  async function markInstance() {
    const tag = `g${Date.now()}`;
    await js(`(window.__otzariaEditor && window.__otzariaEditor.superdoc) ? (window.__otzariaEditor.superdoc.__qaTag = ${JSON.stringify(tag)}, true) : false`);
    return tag;
  }
  const sameInstance = (tag) => js(`!!(window.__otzariaEditor && window.__otzariaEditor.superdoc && window.__otzariaEditor.superdoc.__qaTag === ${JSON.stringify(tag)})`);

  async function waitApply(tag, ms = 120_000) {
    const r = await until(async () => {
      const same = await sameInstance(tag);
      if (!same) throw new Error("style edit unexpectedly replaced the engine instance");
      const ready = await js(`!!(window.__qa && window.__qa.ready() && document.querySelector('.editor-stack .superdoc-fragment'))`);
      const dialog = await app.exists('[data-testid="style-dialog"]');
      return ready && !dialog;
    }, ms, 500);
    if (r.ok) await sleep(2500); // גלריה, קטלוג, השבת הסמן
    return r;
  }

  /** לחיצה אמיתית באלמנט לפי בורר. */
  async function clickSel(selector, { button = 'left', after = 500, index = 0 } = {}) {
    const r = await J(`(() => { const e = document.querySelectorAll(${JSON.stringify(selector)})[${index}];
      if (!e) return 'null'; e.scrollIntoView({ block: 'nearest', inline: 'center' }); const r = e.getBoundingClientRect();
      return JSON.stringify({ x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2), disabled: !!e.disabled }); })()`);
    if (!r) return null;
    if (button === 'left') {
      await app.clickAt(r.x, r.y);
    } else {
      await app.cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: r.x, y: r.y, button: 'none', buttons: 0 });
      await app.cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: r.x, y: r.y, button: 'right', buttons: 2, clickCount: 1 });
      await sleep(40);
      await app.cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: r.x, y: r.y, button: 'right', buttons: 0, clickCount: 1 });
    }
    await sleep(after);
    return r;
  }

  const cardSel = (label) => `.style-card[data-tip-title=${JSON.stringify(label)}]`;
  const gallery = () => app.galleryItems();
  const activeCard = async () => (await gallery()).find((g) => g.active)?.label ?? null;

  async function clickCard(label) {
    const r = await clickSel(cardSel(label), { after: 1500 });
    return r !== null;
  }

  /** לחצן ימני על כרטיס → פריטי התפריט. */
  async function cardMenu(label) {
    const r = await clickSel(cardSel(label), { button: 'right', after: 600 });
    if (!r) return null;
    return J(`JSON.stringify([...document.querySelectorAll('[data-testid="style-menu"] [data-action]')].map(b => ({ action: b.dataset.action, label: b.textContent.trim(), disabled: b.disabled })))`);
  }
  async function moreMenu() {
    await clickSel('[data-testid="style-gallery-more"]', { after: 600 });
    return J(`JSON.stringify([...document.querySelectorAll('[data-testid="style-menu"] [data-action]')].map(b => ({ action: b.dataset.action, label: b.textContent.trim(), disabled: b.disabled })))`);
  }
  const menuOpen = () => app.exists('[data-testid="style-menu"]');
  const pickMenu = (action) => clickSel(`[data-testid="style-menu"] [data-action="${action}"]`, { after: 300 });

  /** מחכה לדיאלוג (הפתיחה קוראת את הגיליון בייצוא — לא מיידית). */
  async function waitDialog(ms = 60_000) {
    return until(() => app.exists('[data-testid="style-dialog"]'), ms, 300);
  }

  /** הערכים שהדיאלוג מציג. */
  const dialogValues = () => J(`(() => { const d = document.querySelector('[data-testid="style-dialog"]'); if (!d) return 'null';
    const v = (s) => d.querySelector(s)?.value ?? null;
    const pressed = (l) => d.querySelector('button[aria-label="' + l + '"]')?.getAttribute('aria-pressed');
    return JSON.stringify({ title: d.getAttribute('aria-label'), name: v('#sd-name'), based: v('#sd-based'), next: v('#sd-next'),
      font: d.querySelector('input[role="combobox"]')?.value ?? null, size: v('input[aria-label="גודל, בנקודות"]'),
      bold: pressed('מודגש'), italic: pressed('נטוי'), underline: pressed('קו תחתון'), color: v('input[type="color"]'),
      align: ['יישור לימין','מרכז','יישור לשמאל','יישור לשני הצדדים'].find((l) => pressed(l) === 'true') ?? null,
      before: v('#sd-before'), after: v('#sd-after'), line: v('#sd-line'), start: v('#sd-start'), end: v('#sd-end'), first: v('#sd-first'),
      error: d.querySelector('.sd-error')?.textContent?.trim() ?? null, primary: d.querySelector('.sd-btn-primary')?.textContent?.trim() }); })()`);

  /** כתיבה לשדה בדיאלוג (input+change, כמו הקלדה). */
  const fill = (selector, value) => js(`(() => { const e = document.querySelector('[data-testid="style-dialog"] ' + ${JSON.stringify(selector)});
    if (!e) return 'not-found'; e.value = ${JSON.stringify(String(value))};
    e.dispatchEvent(new Event('input', { bubbles: true })); e.dispatchEvent(new Event('change', { bubbles: true })); return 'ok'; })()`);
  const dialogButton = (label) => clickSel(`[data-testid="style-dialog"] button[aria-label="${label}"]`, { after: 250 });
  const ok = () => clickSel('[data-testid="style-dialog"] .sd-btn-primary', { after: 200 });

  /** פותחת „שנה” על כרטיס, ממלאת, ומאשרת. מחזירה את מה שנמדד בדרך. */
  async function modify(card, edits) {
    const menu = await cardMenu(card);
    await pickMenu('modify');
    const opened = await waitDialog();
    if (!opened.ok) return { menu, opened: false };
    const shown = await dialogValues();
    for (const [sel, value] of edits.fields ?? []) await fill(sel, value);
    for (const label of edits.buttons ?? []) await dialogButton(label);
    const tag = await markInstance();
    await ok();
    const reload = await waitApply(tag);
    return { menu, opened: true, openMs: opened.ms, shown, reload };
  }

  /** ממקמת סמן בפסקה (לפי טקסט), ומאמתת מול המנוע. */
  const caretIn = (text) => app.caretPara(text, { verifyMs: 30_000 });

  async function selection() {
    return J(`(async () => { try { const d = window.__otzariaEditor.superdoc.activeEditor.doc;
      const s = await Promise.race([d.selection.current(), new Promise(r => setTimeout(() => r(null), 15000))]);
      return JSON.stringify(s); } catch (e) { return JSON.stringify({ error: String(e && e.message) }); } })()`);
  }

  /** בוחרת את כל הטקסט בפסקה בגרירה מקצה לקצה של השורה. */
  async function selectPara(text) {
    const r = await J(`JSON.stringify(window.__qa.paraRectByText(${JSON.stringify(text)}))`);
    if (!r) throw new Error(`אין פסקה ${text}`);
    const m = (type, x, buttons) => app.cdp.send('Input.dispatchMouseEvent', { type, x, y: r.y, button: buttons || type !== 'mouseMoved' ? 'left' : 'none', buttons, clickCount: 1 });
    await m('mouseMoved', r.right + 2, 0);
    await m('mousePressed', r.right + 2, 1);
    await sleep(80);
    await m('mouseMoved', r.left - 2, 1);
    await sleep(120);
    await m('mouseReleased', r.left - 2, 0);
    await sleep(700);
    return app.selection();
  }

  const shell = () => J(`JSON.stringify({ dirty: !!document.querySelector('.dirty-indicator'),
    zoom: document.querySelector('.zoom-pct-btn')?.textContent?.trim() ?? null,
    undo: !(document.querySelector('.quick-access-tools button[data-tip-title="בטל"]')?.disabled ?? true),
    title: document.querySelector('.doc-title-input')?.value ?? null,
    status: document.querySelector('.status-message')?.textContent?.trim() ?? null })`);

  async function type(lines) {
    for (let i = 0; i < lines.length; i++) {
      await app.type(lines[i], 60);
      if (i < lines.length - 1) { await app.press('Enter', 'Enter', 13, 0, '\r'); await sleep(400); }
    }
    await sleep(2000);
  }

  /** שמירה אמיתית (Ctrl+S) עם יירוט ההעלאה — הבייטים שהמשתמש מקבל. */
  async function installSaveCapture() {
    await js(`(function(){if(window.__origFetch)return;window.__origFetch=window.fetch.bind(window);
      window.fetch=function(url,opts){ if(opts&&opts.method==='PUT'){ var body=opts.body;
        var read=body&&body.arrayBuffer?body.arrayBuffer():Promise.resolve(null);
        return read.then(function(buf){ if(buf){var b=new Uint8Array(buf),s='';for(var i=0;i<b.length;i++)s+=String.fromCharCode(b[i]); window.__savedDocx=btoa(s);}
          return new Response('',{status:200});}); }
        return window.__origFetch(url,opts);};})()`);
    await js(`window.__qaHost.replies['fs.beginBinaryWrite']=function(){return Promise.resolve({success:true,error:null,data:{writeToken:'wt-sm',uploadUrl:'https://qa-upload.local/sm',maxBytes:999999999}})}`);
    await js(`window.__qaHost.replies['fs.commitUserFileWrite']=function(){return Promise.resolve({success:true,error:null,data:{cancelled:false,token:'tok-sm',name:'style-manager.docx',size:1234}})}`);
  }
  async function save(label) {
    await js('window.__savedDocx = ""');
    await app.press('s', 'KeyS', 83, 2, 's');
    for (let i = 0; i < 60; i++) {
      await sleep(500);
      const saved = await js('window.__savedDocx || ""');
      if (saved) {
        const bytes = Buffer.from(saved, 'base64');
        if (SAVE_DIR) writeFileSync(join(SAVE_DIR, `${label}.docx`), bytes);
        return unzip(bytes);
      }
    }
    return null;
  }

  return { js, J, measure, para, until, engineExport, markInstance, sameInstance, waitApply, clickSel, cardSel, gallery, activeCard,
    clickCard, cardMenu, moreMenu, menuOpen, pickMenu, waitDialog, dialogValues, fill, dialogButton, ok, modify, caretIn, selection,
    selectPara, shell, type, installSaveCapture, save };
}

/* ------------------------------------------------------------------ */
/* שלבים                                                               */
/* ------------------------------------------------------------------ */

const PHASE_MS = 900_000;

async function phase(key, label, body, { extra = '' } = {}) {
  if (ONLY && !ONLY.has(key)) return;
  console.log(`\n########## ${key}: ${label} ##########`);
  let app;
  try {
    app = await openApp({ name: `style-mgr-${key}`, port: PORT, extra });
    await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });
    await sleep(2000);
    await app.tab('בית');
    const ctx = ctxOf(app);
    await ctx.installSaveCapture();
    const done = await Promise.race([
      body(app, ctx).then(() => true),
      sleep(PHASE_MS).then(() => false),
    ]);
    if (!done) report.stuck(`שלב ${key}`, `לא הסתיים בתוך ${PHASE_MS / 1000}s`);
    const log = await Promise.race([app.log(), sleep(8000).then(() => 'timeout')]);
    console.log('לוג הדף:', JSON.stringify(log).slice(0, 1500));
  } catch (error) {
    report.fail(`שלב ${key} — ${label}`, `השער זרק: ${error?.message ?? error}`);
    console.log(error?.stack ?? error);
  } finally {
    try { app?.close(); } catch { /* הדפדפן ממילא נסגר */ }
  }
  await sleep(1500);
}

const log = (...a) => console.log(...a);

/* ================================================================== */
/* A — שינוי כותרת 1, „עדכן”, מקרי קצה                                 */
/* ================================================================== */

await phase('A', 'שינוי סגנון מובנה, עדכון לפי הבחירה, קצוות', async (app, c) => {
  await c.caretIn(0);
  await c.type([P1, P2, P3]);
  await c.caretIn(P2);
  await c.clickCard('כותרת 1');
  log('אחרי הקלקה על כותרת 1 — כרטיס פעיל:', await c.activeCard());

  // זום 120% — כדי לראות שהוא נשמר בהחלה חיה.
  await app.click('הגדל תצוגה', { after: 600 });
  await app.click('הגדל תצוגה', { after: 600 });

  // סמן באמצע הפסקה השנייה: לחיצה, Home, ושלושה צעדים.
  await c.caretIn(P2);
  await app.press('Home', 'Home', 36);
  await sleep(200);
  for (let i = 0; i < 3; i++) { await app.press('ArrowLeft', 'ArrowLeft', 37); await sleep(120); }
  await sleep(600);

  const before = { paras: await c.measure(), shell: await c.shell(), sel: await c.selection(), styles: (await c.engineExport('לפני'))['word/styles.xml'] };
  log('לפני:', JSON.stringify({ paras: before.paras, shell: before.shell, sel: before.sel }));
  log('Heading1 לפני:', styleXml(before.styles, 'Heading1'));
  log('Heading1Char לפני:', styleXml(before.styles, 'Heading1Char'));

  /* ---- 7: תפריט — פריטים, ביטול ב-Escape, לחיצה בחוץ ---- */
  const menu = await c.cardMenu('כותרת 1');
  log('תפריט כותרת 1:', JSON.stringify(menu));
  const disabled = Object.fromEntries((menu ?? []).map((m) => [m.action, m.disabled]));
  if (menu && menu.length === 4 && disabled.rename && !("delete" in disabled) && !disabled.modify && !disabled.update && !disabled.apply)
    report.pass('תפריט כרטיס מובנה', 'apply/update/modify זמינים, rename מושבת; מחיקה אינה מוצגת');
  else report.fail('תפריט כרטיס מובנה', JSON.stringify(menu));
  await app.press('Escape', 'Escape', 27);
  await sleep(400);
  const afterEsc = await c.menuOpen();
  await c.cardMenu('כותרת 1');
  // לחיצה בשורת המצב, רחוק מהתפריט ומהמסמך.
  const sb = await c.J(`(() => { const r = document.querySelector('.word-statusbar')?.getBoundingClientRect(); return r ? JSON.stringify({ x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2) }) : 'null'; })()`);
  if (sb) await app.clickAt(sb.x, sb.y);
  await sleep(400);
  const afterOutside = await c.menuOpen();
  log(`תפריט: אחרי Escape פתוח=${afterEsc}, אחרי לחיצה בחוץ פתוח=${afterOutside}`);
  !afterEsc ? report.pass('Escape סוגר את התפריט') : report.fail('Escape סוגר את התפריט', 'התפריט נשאר פתוח');
  !afterOutside ? report.pass('לחיצה בחוץ סוגרת את התפריט') : report.fail('לחיצה בחוץ סוגרת את התפריט', 'התפריט נשאר פתוח');

  /* ---- 7: דיאלוג — Escape, ואישור בלי שינוי ---- */
  await c.cardMenu('כותרת 1');
  await c.pickMenu('modify');
  let d = await c.waitDialog();
  log('דיאלוג נפתח אחרי', d.ms, 'ms; ערכים:', JSON.stringify(await c.dialogValues()));
  await app.press('Escape', 'Escape', 27);
  await sleep(500);
  const dialogAfterEsc = await app.exists('[data-testid="style-dialog"]');
  !dialogAfterEsc ? report.pass('Escape סוגר את הדיאלוג') : report.fail('Escape סוגר את הדיאלוג', 'הדיאלוג נשאר פתוח');
  if (dialogAfterEsc) await c.clickSel('[data-testid="style-dialog"] .sd-close-btn');

  // הסמן חוזר לפסקה השנייה לפני האישור-בלי-שינוי, כדי שהיסטוריית הביטול תהיה ידועה.
  await c.cardMenu('כותרת 1');
  await c.pickMenu('modify');
  d = await c.waitDialog();
  const undoBefore = (await c.shell()).undo;
  const tag0 = await c.markInstance();
  await c.ok();
  await sleep(8000);
  const noop = { same: await c.sameInstance(tag0), dialog: await app.exists('[data-testid="style-dialog"]'), shell: await c.shell(), msgs: await app.messages() };
  log('אישור בלי שינוי:', JSON.stringify(noop), 'undo לפני:', undoBefore);
  noop.same && !noop.dialog && noop.shell.undo === undoBefore
    ? report.pass('אישור בלי שינוי — אין החלה חיה', `אותו מופע מנוע אחרי 8s, הדיאלוג נסגר, undo=${noop.shell.undo}`)
    : report.fail('אישור בלי שינוי — אין החלה חיה', JSON.stringify(noop));

  /* ---- 1: שינוי כותרת 1 — 22pt, מרכז ---- */
  // הסמן הוזז ע"י הלחיצות בתפריט? מחזירים אותו למקום הידוע לפני השינוי.
  await c.caretIn(P2);
  await app.press('Home', 'Home', 36);
  await sleep(200);
  for (let i = 0; i < 3; i++) { await app.press('ArrowLeft', 'ArrowLeft', 37); await sleep(120); }
  await sleep(600);
  const selBefore = await c.selection();
  const shellBefore = await c.shell();
  log('לפני השינוי — בחירה:', JSON.stringify(selBefore), 'מעטפת:', JSON.stringify(shellBefore));

  const t0 = Date.now();
  const r1 = await c.modify('כותרת 1', { fields: [['input[aria-label="גודל, בנקודות"]', '22']], buttons: ['מרכז'] });
  log('שינוי כותרת 1:', JSON.stringify(r1), 'סה"כ', Date.now() - t0, 'ms');
  if (!r1.reload?.ok) {
    report.fail('שינוי כותרת 1 — החלה חיה', `המסמך לא הוחל בתוך ${r1.reload?.ms}ms; הודעות: ${JSON.stringify(await app.messages())}`);
    return;
  }
  report.pass('שינוי כותרת 1 — החלה חיה', `אותו מופע מנוע אחרי ${r1.reload.ms}ms`);

  const after = { paras: await c.measure(), shell: await c.shell(), sel: await c.selection(), active: await c.activeCard() };
  log('אחרי:', JSON.stringify(after));
  const b2 = before.paras.find((p) => p.text.startsWith(P2.slice(0, 10)));
  const a2 = after.paras.find((p) => p.text.startsWith(P2.slice(0, 10)));
  const a1 = after.paras.find((p) => p.text.startsWith(P1.slice(0, 10)));
  const b1 = before.paras.find((p) => p.text.startsWith(P1.slice(0, 10)));
  const a3 = after.paras.find((p) => p.text.startsWith(P3.slice(0, 10)));
  const b3 = before.paras.find((p) => p.text.startsWith(P3.slice(0, 10)));
  const px22 = (22 * 96) / 72;
  const size2 = parseFloat(a2?.fontSize);
  const align2 = classify(a2?.lines[0]);
  log(`פסקה 2: גודל ${b2?.fontSize} → ${a2?.fontSize} (22pt = ${px22.toFixed(2)}px), יישור ${classify(b2?.lines[0])} → ${align2}, שורה=${JSON.stringify(a2?.lines[0])}`);
  Math.abs(size2 - px22) < 0.6 && align2 === 'center'
    ? report.pass('שינוי כותרת 1 — הפסקה בסגנון מצוירת מחדש', `font-size ${a2.fontSize} (22pt), טקסט ממורכז: textL=${a2.lines[0].textL} textR=${a2.lines[0].textR}`)
    : report.fail('שינוי כותרת 1 — הפסקה בסגנון מצוירת מחדש', `font-size ${a2?.fontSize}, יישור ${align2}, ${JSON.stringify(a2?.lines)}`);
  const same = (x, y) => x && y && x.fontSize === y.fontSize && classify(x.lines[0]) === classify(y.lines[0]) && Math.abs(x.lines[0].textR - y.lines[0].textR) <= 2;
  same(a1, b1) && same(a3, b3)
    ? report.pass('שינוי כותרת 1 — שאר הפסקאות לא השתנו', `P1 ${a1.fontSize} textR=${a1.lines[0].textR}; P3 ${a3.fontSize} textR=${a3.lines[0].textR}`)
    : report.fail('שינוי כותרת 1 — שאר הפסקאות לא השתנו', JSON.stringify({ b1, a1, b3, a3 }));

  const zoomOk = after.shell.zoom === shellBefore.zoom;
  zoomOk ? report.pass('שינוי כותרת 1 — זום נשמר', after.shell.zoom) : report.fail('שינוי כותרת 1 — זום נשמר', `${shellBefore.zoom} → ${after.shell.zoom}`);
  after.shell.dirty ? report.pass('שינוי כותרת 1 — מסומן „לא נשמר”', `dirty-indicator קיים; status=${after.shell.status}`) : report.fail('שינוי כותרת 1 — מסומן „לא נשמר”', JSON.stringify(after.shell));
  log(`undo: לפני=${shellBefore.undo} אחרי=${after.shell.undo} (נשמר)`);

  // הסמן: אותה פסקה (לפי טקסט) ואותו היסט.
  const seg = (s) => (s?.target?.segments ?? [])[0] ?? null;
  const caretText = async (s) => {
    const id = seg(s)?.blockId;
    return id ? (await c.measure()).find((p) => p.id === id)?.text ?? `(בלוק ${id} לא מצויר)` : null;
  };
  const sb4 = seg(selBefore); const sa4 = seg(after.sel);
  const tb = before.paras.find((p) => p.id === sb4?.blockId)?.text ?? null;
  const ta = await caretText(after.sel);
  log('סמן לפני:', JSON.stringify(sb4), tb, '| אחרי:', JSON.stringify(sa4), ta);
  const offs = (s) => JSON.stringify(s?.range ?? s?.start ?? s?.offset ?? s);
  tb && ta && tb === ta && offs({ ...sb4, blockId: 0 }) === offs({ ...sa4, blockId: 0 })
    ? report.pass('שינוי כותרת 1 — הסמן נשמר', `„${ta}”, ${offs({ ...sa4, blockId: 0 })}`)
    : report.fail('שינוי כותרת 1 — הסמן נשמר', `לפני ${JSON.stringify(sb4)} „${tb}” / אחרי ${JSON.stringify(sa4)} „${ta}”`);

  // ההוכחה בקובץ: שמירה אמיתית.
  const saved = await c.save('A-heading1-22-center');
  const h1 = styleXml(saved?.['word/styles.xml'], 'Heading1');
  const h1c = styleXml(saved?.['word/styles.xml'], 'Heading1Char');
  log('Heading1 בשמירה:', h1);
  log('Heading1Char בשמירה:', h1c);
  const p2x = paraBy(saved, P2);
  log('פסקה 2 בשמירה:', p2x?.pPr, 'jc ישיר:', p2x?.jc);
  const szOk = /<w:sz w:val="44"\/>/.test(h1 ?? '') && /<w:szCs w:val="44"\/>/.test(h1 ?? '');
  const jcOk = /<w:jc w:val="center"\/>/.test(h1 ?? '');
  const schema = schemaCheck(h1);
  szOk && jcOk && !schema
    ? report.pass('שינוי כותרת 1 — styles.xml', `sz/szCs=44, jc=center, סדר הסכמה תקין: style[${childNames(h1)}] pPr[${childNames(sub(h1, 'pPr'))}] rPr[${childNames(sub(h1, 'rPr'))}]`)
    : report.fail('שינוי כותרת 1 — styles.xml', `sz=${szOk} jc=${jcOk} schema=${schema}: ${h1}`);
  const linkedOk = /<w:sz w:val="44"\/>/.test(h1c ?? '');
  linkedOk ? report.pass('שינוי כותרת 1 — הסגנון המקושר עודכן', 'Heading1Char sz=44') : report.partial('שינוי כותרת 1 — הסגנון המקושר עודכן', h1c ?? 'אין Heading1Char');
  const sc = schemaCheck(h1c);
  if (h1c && sc) report.fail('Heading1Char — סדר הסכמה', sc);

  /* ---- 5: עדכן כך שיתאים לבחירה ---- */
  await c.caretIn(P3);
  await c.clickCard('כותרת 1');
  await c.selectPara(P2);
  const sv = await app.selectValue('גודל גופן', '28');
  await sleep(1200);
  log('גודל 28 לפסקה 2 בבורר:', sv, 'בחירה:', JSON.stringify(await app.selection()));
  await c.caretIn(P2);
  const mid = await c.measure();
  log('לפני „עדכן”:', JSON.stringify(mid.map((p) => [p.text, p.fontSize])));
  const stylesBeforeUpdate = (await c.engineExport('לפני עדכן'))['word/styles.xml'];
  const menu5 = await c.cardMenu('כותרת 1');
  const tag5 = await c.markInstance();
  await c.pickMenu('update');
  const r5 = await c.waitApply(tag5, 120_000);
  const after5 = await c.measure();
  log('„עדכן”: החלה חיה', JSON.stringify(r5), 'מדידה:', JSON.stringify(after5.map((p) => [p.text, p.fontSize, classify(p.lines[0])])), 'הודעות:', JSON.stringify(await app.messages()));
  const s5 = (await c.engineExport('אחרי עדכן'))['word/styles.xml'];
  const h15 = styleXml(s5, 'Heading1');
  log('Heading1 לפני „עדכן”:', styleXml(stylesBeforeUpdate, 'Heading1'));
  log('Heading1 אחרי „עדכן”:', h15);
  const p35 = after5.find((p) => p.text.startsWith(P3.slice(0, 10)));
  const px28 = (28 * 96) / 72;
  r5.ok && /<w:sz w:val="56"\/>/.test(h15 ?? '') && Math.abs(parseFloat(p35?.fontSize) - px28) < 0.6
    ? report.pass('עדכן כך שיתאים לבחירה', `Heading1 sz=56; P3 (כותרת 1 בלי עיצוב ישיר) מצוירת ${p35.fontSize}; תפריט=${JSON.stringify(menu5?.map((m) => m.action))}`)
    : report.fail('עדכן כך שיתאים לבחירה', `reload=${r5.ok} P3=${p35?.fontSize} Heading1=${h15}`);
  const extra = childNames(sub(h15, 'pPr')).filter((n) => !childNames(sub(styleXml(stylesBeforeUpdate, 'Heading1'), 'pPr')).includes(n));
  const extraR = childNames(sub(h15, 'rPr')).filter((n) => !childNames(sub(styleXml(stylesBeforeUpdate, 'Heading1'), 'rPr')).includes(n));
  log('„עדכן” הוסיף לסגנון: pPr', JSON.stringify(extra), 'rPr', JSON.stringify(extraR));
  const spBefore = styleXml(stylesBeforeUpdate, 'Heading1')?.match(/<w:spacing [^>]*\/>/)?.[0] ?? null;
  const spAfter = h15?.match(/<w:spacing [^>]*\/>/)?.[0] ?? null;
  const fBefore = styleXml(stylesBeforeUpdate, 'Heading1')?.match(/<w:rFonts [^>]*\/>/)?.[0] ?? null;
  const fAfter = h15?.match(/<w:rFonts [^>]*\/>/)?.[0] ?? null;
  spBefore === spAfter
    ? report.pass('„עדכן” אינו נוגע במה שלא שונה — ריווח', spAfter)
    : report.fail('„עדכן” אינו נוגע במה שלא שונה — ריווח', `שונה רק הגודל, אבל ${spBefore} → ${spAfter}`);
  fBefore === fAfter
    ? report.pass('„עדכן” אינו נוגע במה שלא שונה — גופן', fAfter)
    : report.partial('„עדכן” אינו נוגע במה שלא שונה — גופן', `${fBefore} → ${fAfter}`);
  const sc5 = schemaCheck(h15);
  sc5 ? report.fail('„עדכן” — סדר הסכמה', sc5) : report.pass('„עדכן” — סדר הסכמה', `pPr[${childNames(sub(h15, 'pPr'))}] rPr[${childNames(sub(h15, 'rPr'))}]`);
});

/* ================================================================== */
/* B — יישור: מה הסגנון מצייר, ומה הרצועה כותבת                         */
/* ================================================================== */

const ALIGNS = ['יישור לימין', 'מרכז', 'יישור לשמאל', 'יישור לשני הצדדים'];
/** הסגנון נבדק מימין אחרון: „ימין” על סגנון בלי jc הוא ברירת המחדל שהדיאלוג מציג, ולכן אינו שינוי. */
const STYLE_ALIGNS = ['מרכז', 'יישור לשמאל', 'יישור לשני הצדדים', 'יישור לימין'];
const EXPECT = { 'יישור לימין': 'right', 'מרכז': 'center', 'יישור לשמאל': 'left', 'יישור לשני הצדדים': 'right' /* שורה יחידה: לא נמתחת */ };

await phase('B', 'יישור בסגנון מול יישור ברצועה', async (app, c) => {
  await c.caretIn(0);
  await c.type([P1, P2, P3]);
  await c.caretIn(P2);
  await c.clickCard('כותרת 1');

  /* (b) קודם הרצועה, על פסקה רגילה — מה המנוע עצמו כותב ואיפה הוא מצייר. */
  for (const label of ALIGNS) {
    await c.caretIn(P1);
    await app.click(label);
    await sleep(900);
    const p = await c.para(P1);
    const files = await c.engineExport(`רצועה ${label}`);
    const x = paraBy(files, P1);
    log(`רצועה „${label}”: jc=${x?.jc} bidi=${x?.bidi} cmd=${JSON.stringify(await app.cmd('text-align'))} נראה=${classify(p?.lines[0])} ${JSON.stringify(p?.lines[0])}`);
    report.pass(`רצועה „${label}” על פסקה עברית`, `המנוע כותב w:jc="${x?.jc}" (bidi=${x?.bidi}), מצויר ${classify(p?.lines[0])}`);
  }
  // מחזירים את P1 לימין כדי שלא תפריע.
  await c.caretIn(P1);
  await app.click('יישור לימין');

  /* (a) הסגנון. */
  for (const label of STYLE_ALIGNS) {
    await c.caretIn(P3);
    const r = await c.modify('כותרת 1', { buttons: [label] });
    if (!r.reload?.ok) {
      report.fail(`סגנון „${label}”`, `לא הוחל (${JSON.stringify(r)}); הודעות ${JSON.stringify(await app.messages())}`);
      continue;
    }
    const p = await c.para(P2);
    const seen = classify(p?.lines[0]);
    const saved = await c.save(`B-style-${['right', 'center', 'left', 'justify'][ALIGNS.indexOf(label)]}`);
    const jc = styleXml(saved?.['word/styles.xml'], 'Heading1')?.match(/<w:jc w:val="([^"]+)"/)?.[1] ?? null;
    const x = paraBy(saved, P2);
    log(`סגנון „${label}”: styles jc=${jc}; פסקה: pStyle=${x?.pStyle} bidi=${x?.bidi} jc ישיר=${x?.jc}; מצויר ${seen} ${JSON.stringify(p?.lines[0])}`);
    seen === EXPECT[label]
      ? report.pass(`סגנון „${label}” — מצויר במקום`, `Heading1 jc="${jc}", הטקסט ${seen} (textL=${p.lines[0].textL}, textR=${p.lines[0].textR})`)
      : report.fail(`סגנון „${label}” — מצויר במקום`, `Heading1 jc="${jc}", אבל הטקסט ${seen} (textL=${p?.lines[0]?.textL}, textR=${p?.lines[0]?.textR})`);
  }

  // הדיאלוג קורא את היישור בחזרה כמו שנכתב?
  await c.cardMenu('כותרת 1');
  await c.pickMenu('modify');
  await c.waitDialog();
  const shown = await c.dialogValues();
  log('הדיאלוג אחרי „ימין” (jc=left):', JSON.stringify(shown));
  shown?.align === 'יישור לימין'
    ? report.pass('הדיאלוג קורא את היישור בחזרה', shown.align)
    : report.fail('הדיאלוג קורא את היישור בחזרה', JSON.stringify(shown));
  await app.press('Escape', 'Escape', 27);
});

/* ================================================================== */
/* C — הזחות                                                           */
/* ================================================================== */

await phase('C', 'הזחות בסגנון (start / שורה ראשונה / תלויה)', async (app, c) => {
  await c.caretIn(0);
  await c.type([P1, P2_LONG, P3]);
  await c.caretIn(P2_LONG.slice(0, 20));
  await c.clickCard('כותרת 1');
  const base = await c.para(P2_LONG);
  log('לפני:', JSON.stringify(base));
  const cases = [
    ['start 1 / first 0.5', [['#sd-start', '1'], ['#sd-first', '0.5']]],
    ['start 1 / hanging 0.5', [['#sd-start', '1'], ['#sd-first', '-0.5']]],
  ];
  for (const [name, fields] of cases) {
    await c.caretIn(P3);
    const r = await c.modify('כותרת 1', { fields });
    if (!r.reload?.ok) { report.fail(`הזחה ${name}`, `לא הוחל: ${JSON.stringify(r)}`); continue; }
    const p = await c.para(P2_LONG);
    const p1 = await c.para(P1);
    const files = await c.engineExport(name);
    const ind = styleXml(files['word/styles.xml'], 'Heading1')?.match(/<w:ind [^>]*\/>/)?.[0] ?? null;
    const scale = p?.scale ?? 1;
    const cmPx = (96 / 2.54) * scale;
    const l0 = p?.lines[0]; const l1 = p?.lines[1];
    // הקצה הימני של הטקסט ביחס לפסקה רגילה (P1) — בעברית תחילת השורה בימין.
    const d0 = l0 && p1 ? (l0.textR - p1.lines[0].textR) / cmPx : null;
    const d1 = l1 && p1 ? (l1.textR - p1.lines[0].textR) / cmPx : null;
    const dl = l1 && p1 ? (l1.textL - p1.lines[0].textL) : null;
    log(`הזחה ${name}: ${ind}; שורה 1 ימין +${d0?.toFixed(2)}cm, שורה 2 ימין +${d1?.toFixed(2)}cm; שורות=${JSON.stringify(p?.lines)} P1=${JSON.stringify(p1?.lines[0])}; dL=${dl}`);
    const want = name.includes('hanging') ? [0.5, 1] : [1.5, 1];
    l0 && l1 && Math.abs(d0 - want[0]) < 0.08 && Math.abs(d1 - want[1]) < 0.08
      ? report.pass(`הזחה ${name}`, `${ind}; מצויר מימין: שורה ראשונה ${d0.toFixed(2)}cm, שאר השורות ${d1.toFixed(2)}cm`)
      : report.fail(`הזחה ${name}`, `${ind}; מצויר: שורה ראשונה ${d0?.toFixed(2)}cm מימין, שנייה ${d1?.toFixed(2)}cm (צפוי ${want}); שמאל שורה 2 זז ${dl}px; ${JSON.stringify(p?.lines)}`);
  }
  await c.save('C-indent-hanging');
});

/* ================================================================== */
/* D — צור מהבחירה, שינוי שם                                    */
/* ================================================================== */

await phase('D', 'צור סגנון מהבחירה, שינוי שם', async (app, c) => {
  await c.caretIn(0);
  await c.type([P1, P2, P3]);

  // עיצוב ישיר דרך הרצועה: גודל 18, מודגש, אדום.
  const sel = await c.selectPara(P3);
  log('בחירה:', JSON.stringify(sel));
  const sz = await app.selectValue('גודל גופן', '18');
  await sleep(900);
  await app.click('מודגש', { after: 900 });
  const opened = await app.click('בחירת צבע', { index: 1, after: 700 });
  const swatches = (await app.paletteSwatches()) ?? [];
  const red = await c.J(`JSON.stringify([...document.querySelectorAll('.color-palette-popover button')].findIndex((b) => {
    const bg = [b, ...b.querySelectorAll('*')].map((e) => getComputedStyle(e).backgroundColor);
    return bg.includes('rgb(255, 0, 0)') || /ff0000|אדום/i.test((b.getAttribute('data-tip-title') || '') + (b.getAttribute('aria-label') || '') + (b.title || ''));
  }))`);
  log('משבצות:', JSON.stringify(swatches.slice(0, 12)));
  log('גודל:', sz, 'פופאובר צבע נפתח:', opened, 'אדום באינדקס', red, 'מתוך', swatches.length);
  if (red >= 0) await app.clickPalette(red, { after: 1000 });
  else await app.escape();
  const direct = paraBy(await c.engineExport('עיצוב ישיר'), P3);
  log('P3 אחרי עיצוב ישיר:', direct?.xml);

  await c.caretIn(P3);
  await app.press('End', 'End', 35);
  await sleep(600);
  const colorCmd = await app.cmd('text-color');
  // משבצת „אדום” של לוח Word היא C0504D ולא FF0000 — מה שהוחל הוא מה שנבדק.
  const applied = String(colorCmd?.value ?? '').replace(/^#/, '').toUpperCase();
  log('מצב פקודות בסמן:', JSON.stringify({ size: await app.cmd('font-size'), bold: await app.cmd('bold'), color: colorCmd }));

  const more = await c.moreMenu();
  log('תפריט „עוד”:', JSON.stringify(more));
  await c.pickMenu('create');
  const d = await c.waitDialog();
  const shown = await c.dialogValues();
  log('דיאלוג „צור”:', d.ms, 'ms', JSON.stringify(shown));
  shown && applied && parseFloat(shown.size) === 18 && shown.bold === 'true' && (shown.color ?? '').replace(/^#/, '').toUpperCase() === applied
    ? report.pass('צור — הדיאלוג לוכד את הבחירה', `גודל ${shown.size}, מודגש, צבע ${shown.color}, שם מוצע „${shown.name}”`)
    : report.fail('צור — הדיאלוג לוכד את הבחירה', JSON.stringify(shown));
  await c.fill('#sd-name', 'ציטוט מקור');
  const tag = await c.markInstance();
  await c.ok();
  const r = await c.waitApply(tag);
  if (!r.ok) {
    report.fail('צור — החלה חיה', `לא הוחל בתוך ${r.ms}ms; דיאלוג=${JSON.stringify(await c.dialogValues())} הודעות=${JSON.stringify(await app.messages())}`);
    return;
  }
  await sleep(1500);
  const cards = await c.gallery();
  const active = await c.activeCard();
  log('גלריה אחרי „צור”:', JSON.stringify(cards), 'פעיל:', active, 'מעטפת:', JSON.stringify(await c.shell()));
  const files = await c.engineExport('אחרי צור');
  const created = styleByName(files['word/styles.xml'], 'ציטוט מקור');
  const id = idOf(created);
  const p3 = paraBy(files, P3);
  log('הסגנון שנוצר:', created);
  log('P3 אחרי צור:', p3?.pPr, 'pStyle=', p3?.pStyle);
  cards.some((g) => g.label === 'ציטוט מקור')
    ? report.pass('צור — כרטיס חדש בגלריה', `${cards.length} כרטיסים`)
    : report.fail('צור — כרטיס חדש בגלריה', `הכרטיסים: ${cards.map((g) => g.label).join(', ')}`);
  p3?.pStyle === id && active === 'ציטוט מקור'
    ? report.pass('צור — הסגנון הוחל על פסקת הסמן', `pStyle=${id}, הכרטיס הפעיל „${active}”`)
    : report.fail('צור — הסגנון הוחל על פסקת הסמן', `pStyle=${p3?.pStyle} (id=${id}), פעיל=${active}`);
  const okXml = created && /w:customStyle="1"/.test(created) && /<w:qFormat\/>/.test(created) &&
    /<w:sz w:val="36"\/>/.test(created) && /<w:szCs w:val="36"\/>/.test(created) && /<w:b\/>/.test(created) && /<w:bCs\/>/.test(created) && created.includes(`<w:color w:val="${applied}"/>`);
  const sc = schemaCheck(created);
  okXml && !sc
    ? report.pass('צור — styles.xml', `id=${id}, customStyle, qFormat, sz/szCs=36, b/bCs, color ${applied}; סדר תקין`)
    : report.fail('צור — styles.xml', `schema=${sc}; ${created}`);

  await c.save('D-after-create');

  /* ---- שינוי שם ---- */
  const m = await c.cardMenu('ציטוט מקור');
  log('תפריט הסגנון החדש:', JSON.stringify(m));
  (m ?? []).every((e) => !e.disabled)
    ? report.pass('תפריט כרטיס מותאם', 'rename זמין; מחיקה אינה מוצגת')
    : report.fail('תפריט כרטיס מותאם', JSON.stringify(m));
  await c.pickMenu('rename');
  await c.waitDialog();
  log('דיאלוג שינוי שם:', JSON.stringify(await c.dialogValues()));
  await c.fill('#sd-name', 'ציטוט מקור חדש');
  const tagR = await c.markInstance();
  await c.ok();
  const rr = await c.waitApply(tagR);
  const cardsR = await c.gallery();
  const filesR = await c.engineExport('אחרי שינוי שם');
  const renamed = styleXml(filesR['word/styles.xml'], id);
  log('אחרי שינוי שם:', JSON.stringify(rr), JSON.stringify(cardsR.map((g) => [g.label, g.active])), renamed);
  rr.ok && renamed?.includes('<w:name w:val="ציטוט מקור חדש"/>') && cardsR.some((g) => g.label === 'ציטוט מקור חדש') && !cardsR.some((g) => g.label === 'ציטוט מקור') && paraBy(filesR, P3)?.pStyle === id
    ? report.pass('שינוי שם', `אותו id (${id}), שם חדש בגיליון ובגלריה, הפסקה עדיין בסגנון`)
    : report.fail('שינוי שם', `reload=${rr.ok}; ${renamed}; כרטיסים ${cardsR.map((g) => g.label).join(', ')}; pStyle=${paraBy(filesR, P3)?.pStyle}`);


});

/* ================================================================== */
/* E — בקרה להזחות: אותן הזחות כעיצוב ישיר וכסגנון, בקובץ שלא עבר בתוסף */
/* ================================================================== */

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const LONG = `${P2_LONG} ${P2_LONG}`;
const run = (t) => `<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const INDENTS = {
  plain: '',
  'direct-first': '<w:ind w:start="567" w:firstLine="283"/>',
  'direct-hanging': '<w:ind w:start="567" w:hanging="283"/>',
  'style-first': null,
  'style-hanging': null,
};
export const INDENT_FIXTURE_STYLES =
  `<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="IndFirst"><w:name w:val="IndFirst"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:start="567" w:firstLine="283"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="IndHanging"><w:name w:val="IndHanging"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:start="567" w:hanging="283"/></w:pPr></w:style></w:styles>';
export function indentFixtureBody() {
  return Object.entries(INDENTS).map(([key, ind]) => {
    const pPr = key === 'style-first' ? '<w:pStyle w:val="IndFirst"/><w:bidi/>'
      : key === 'style-hanging' ? '<w:pStyle w:val="IndHanging"/><w:bidi/>'
        : `<w:bidi/>${ind}`;
    return `<w:p><w:pPr>${pPr}</w:pPr>${run(`${key} ${LONG}`)}</w:p>`;
  }).join('');
}

await phase('E', 'בקרה: הזחת שורה ראשונה/תלויה בפסקה עברית — ישיר מול סגנון', async (app, c) => {
  const { buildDocx } = await import('./docx-fixtures.mjs');
  const bytes = buildDocx({ body: indentFixtureBody(), styles: INDENT_FIXTURE_STYLES });
  if (SAVE_DIR) writeFileSync(join(SAVE_DIR, 'E-indent-fixture.docx'), bytes);
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + Buffer.from(bytes).toString('base64');
  await c.js(`window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,data:{token:'tok-ind',url:${JSON.stringify(dataUrl)},name:'indent-fixture.docx',size:${bytes.length},access:'readwrite'}})}`);
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await c.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 12000 });
  const paras = await c.measure();
  const cmPx = (p) => (96 / 2.54) * (p?.scale ?? 1);
  const plain = paras.find((p) => p.text.startsWith('plain'));
  for (const key of Object.keys(INDENTS).filter((k) => k !== 'plain')) {
    const p = paras.find((x) => x.text.startsWith(key));
    const d0 = p && plain ? (p.lines[0].textR - plain.lines[1].textR) / cmPx(p) : null;
    const d1 = p && plain ? (p.lines[1].textR - plain.lines[1].textR) / cmPx(p) : null;
    const want = key.endsWith('first') ? [1.5, 1] : [0.5, 1];
    log(`${key}: שורה ראשונה ${d0?.toFixed(2)}cm מימין, שנייה ${d1?.toFixed(2)}cm (צפוי ${want}); ${JSON.stringify(p?.lines?.slice(0, 2))}`);
    p && Math.abs(d0 - want[0]) < 0.08 && Math.abs(d1 - want[1]) < 0.08
      ? report.pass(`בקרה — ${key}`, `שורה ראשונה ${d0.toFixed(2)}cm, שאר ${d1.toFixed(2)}cm`)
      : report.fail(`בקרה — ${key}`, `שורה ראשונה ${d0?.toFixed(2)}cm, שאר ${d1?.toFixed(2)}cm (צפוי ${want[0]}, ${want[1]})`);
  }
});

/* ================================================================== */
/* F — „צור” במסמך Word שה„רגיל” שלו נושא ריווח (after=160, line=259)   */
/* ================================================================== */

await phase('F', '„צור” על „רגיל” עם ריווח של Word — האם הריווח נשמר בירושה', async (app, c) => {
  const { buildDocx } = await import('./docx-fixtures.mjs');
  const styles = `<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>` +
    '<w:pPrDefault><w:pPr><w:bidi/><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>';
  const body = [P1, P2, P3].map((t) => `<w:p><w:pPr><w:bidi/></w:pPr>${run(t)}</w:p>`).join('');
  const bytes = buildDocx({ body, styles });
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + Buffer.from(bytes).toString('base64');
  await c.js(`window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,data:{token:'tok-sp',url:${JSON.stringify(dataUrl)},name:'spacing-fixture.docx',size:${bytes.length},access:'readwrite'}})}`);
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await c.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 12000 });
  await app.tab('בית');
  await c.caretIn(P2);
  const before = await c.measure();
  await c.moreMenu();
  await c.pickMenu('create');
  await c.waitDialog();
  const shown = await c.dialogValues();
  log('דיאלוג „צור” על „רגיל” עם ריווח:', JSON.stringify(shown));
  await c.fill('#sd-name', 'בדיקת ריווח');
  const tag = await c.markInstance();
  await c.ok();
  const r = await c.waitApply(tag);
  const files = await c.engineExport('F');
  const created = styleByName(files['word/styles.xml'], 'בדיקת ריווח');
  const after = await c.measure();
  const gap = (ps) => { const a = ps.find((p) => p.text.startsWith(P2.slice(0, 10))); const b = ps.find((p) => p.text.startsWith(P3.slice(0, 10))); return a && b ? { p2: a, p3: b } : null; };
  log('הסגנון שנוצר:', created, 'reload', JSON.stringify(r));
  log('לפני:', JSON.stringify(gap(before)), 'אחרי:', JSON.stringify(gap(after)));
  const sp = created?.match(/<w:spacing [^>]*\/>/)?.[0] ?? null;
  !sp
    ? report.pass('„צור” אינו כותב ריווח שלא שונה', 'אין w:spacing בסגנון החדש — יורש מ„רגיל”')
    : report.fail('„צור” אינו כותב ריווח שלא שונה', `המשתמש לא נגע בריווח, אבל נכתב ${sp} (הבסיס: after=160 line=259); הדיאלוג הציג לפני=${shown?.before} אחרי=${shown?.after} שורות=${shown?.line}`);
});

/* ================================================================== */
/* עזרים לקיצורים                                                      */
/* ================================================================== */

/** Ctrl=2, Alt=1, Shift=8 ב-CDP. */
const MOD = { alt: 1, ctrl: 2, shift: 8 };
const COMBO_Q = { key: 'Q', code: 'KeyQ', vk: 81, mods: MOD.ctrl | MOD.alt | MOD.shift, label: 'Ctrl+Shift+Alt+Q' };
const COMBO_W = { key: 'W', code: 'KeyW', vk: 87, mods: MOD.ctrl | MOD.alt | MOD.shift, label: 'Ctrl+Shift+Alt+W' };
const COMBO_B = { key: 'b', code: 'KeyB', vk: 66, mods: MOD.ctrl, label: 'Ctrl+B' };
const pressCombo = (app, k) => app.press(k.key, k.code, k.vk, k.mods);

/**
 * לחיצה אמיתית על הלוכד ואז הצירוף. מחזירה מה שנמדד: מי במיקוד אחרי הלחיצה,
 * ומה הלוכד מציג אחרי הצירוף. בלי עקיפה — אם הלחיצה לא נותנת מיקוד, זה ממצא.
 */
async function record(app, c, k) {
  const r = await c.clickSel('[data-testid="style-shortcut-recorder"]', { after: 300 });
  const focus = await c.J(`JSON.stringify({ active: document.activeElement?.getAttribute('data-testid') || document.activeElement?.className || document.activeElement?.tagName,
    label: document.querySelector('[data-testid="style-shortcut-recorder"]')?.getAttribute('aria-label') })`);
  await pressCombo(app, k);
  await sleep(400);
  const read = () => c.J(`JSON.stringify(document.querySelector('[data-testid="style-shortcut-recorder"]')?.textContent?.trim() ?? null)`);
  const shown = await read();
  if (shown === k.label) return { found: !!r, focus, shown, mouseOnly: true };
  // עקיפה כדי למדוד את ההמשך: מיקוד כמו Tab, ואז אותה לחיצה. המסלול בעכבר בלבד כבר נמדד ונכשל.
  await c.js(`document.querySelector('[data-testid="style-shortcut-recorder"]')?.focus()`);
  await c.clickSel('[data-testid="style-shortcut-recorder"]', { after: 300 });
  await pressCombo(app, k);
  await sleep(400);
  return { found: !!r, focus, shown, mouseOnly: false, withFocus: await read() };
}

const dialogState = (c) => c.J(`(() => { const d = document.querySelector('[data-testid="style-dialog"]'); if (!d) return 'null';
  return JSON.stringify({ okDisabled: !!d.querySelector('.sd-btn-primary')?.disabled, error: d.querySelector('.sd-error')?.textContent?.trim() ?? null,
    recorder: d.querySelector('[data-testid="style-shortcut-recorder"]')?.textContent?.trim() ?? null }); })()`);

/** File → „ניהול קיצורים” → הרשימה. סוגר אחרי הקריאה. */
async function managerList(app, c) {
  await app.tab('קובץ');
  await app.click('ניהול קיצורים', { after: 1200 });
  const list = await c.J(`JSON.stringify([...document.querySelectorAll('.sm-item')].map((b) => ({
    name: b.querySelector('.sm-item-name')?.textContent?.trim(), combo: b.querySelector('.sm-item-combo')?.textContent?.trim(),
    summary: b.querySelector('.sm-item-preset')?.textContent?.trim() })))`);
  await app.escape();
  await sleep(400);
  await app.tab('בית');
  return list;
}

/** File → „קיצורים” → השורות שמזכירות סגנון. */
async function helpRows(app, c) {
  await app.tab('קובץ');
  await app.click('קיצורים', { after: 1200, exact: true });
  const rows = await c.J(`JSON.stringify([...document.querySelectorAll('.shortcuts-dialog .shortcut-desc')]
    .map((e) => ((e.previousElementSibling?.textContent || '') + ' | ' + (e.textContent || '')).replace(/\\s+/g, ' ').trim()).filter((t) => t.includes('סגנון')))`);
  const open = await app.exists('.shortcuts-dialog');
  await app.escape();
  await sleep(400);
  await app.tab('בית');
  return { open, rows: rows ?? [] };
}

const storedShortcuts = (c) => c.J(`JSON.stringify(window.__qaHost.storage['custom-shortcuts'] ?? null)`);

/* ================================================================== */
/* G — רמת חלוקה בסגנון מותאם, ותוכן עניינים                           */
/* ================================================================== */

const PG = ['פתיחה רגילה', 'פרק בכותרת אחת', 'פרק בסגנון מותאם', 'סיום רגיל'];

await phase('G', 'רמת חלוקה בסגנון מותאם — styles.xml ותוכן עניינים', async (app, c) => {
  await c.caretIn(0);
  await c.type(PG);
  await c.caretIn(PG[1]);
  await c.clickCard('כותרת 1');
  await c.caretIn(PG[2]);
  await c.moreMenu();
  await c.pickMenu('create');
  await c.waitDialog();
  const opts = await c.J(`JSON.stringify([...document.querySelectorAll('#sd-outline option')].map((o) => [o.value, o.textContent.trim()]))`);
  log('אפשרויות רמת חלוקה:', JSON.stringify(opts), 'ערך בפתיחה:', await c.js(`document.querySelector('#sd-outline')?.value`));
  await c.fill('#sd-name', 'כותרת מקור');
  await c.fill('#sd-outline', '0');
  const tag = await c.markInstance();
  await c.ok();
  const r = await c.waitApply(tag);
  if (!r.ok) { report.fail('רמת חלוקה — יצירה', `לא הוחל: ${JSON.stringify(await c.dialogValues())}`); return; }
  const files = await c.engineExport('G');
  const st = styleByName(files['word/styles.xml'], 'כותרת מקור');
  const id = idOf(st);
  const p = paraBy(files, PG[2]);
  log('הסגנון:', st, 'פסקה:', p?.pPr);
  const sc = schemaCheck(st);
  /<w:outlineLvl w:val="0"\/>/.test(st ?? '') && !sc && p?.pStyle === id
    ? report.pass('רמת חלוקה — styles.xml', `<w:outlineLvl w:val="0"/>, pPr[${childNames(sub(st, 'pPr'))}], הפסקה בסגנון ${id}`)
    : report.fail('רמת חלוקה — styles.xml', `schema=${sc} pStyle=${p?.pStyle}: ${st}`);

  // הדיאלוג קורא את הרמה בחזרה.
  await c.cardMenu('כותרת מקור');
  await c.pickMenu('modify');
  await c.waitDialog();
  const back = await c.js(`document.querySelector('#sd-outline')?.value`);
  await app.escape();
  back === '0' ? report.pass('רמת חלוקה — הדיאלוג קורא בחזרה', 'רמה 1') : report.fail('רמת חלוקה — הדיאלוג קורא בחזרה', `value=${back}`);

  // תוכן עניינים בעורך — מראש המסמך.
  await c.caretIn(PG[0]);
  await app.press('Home', 'Home', 36);
  await sleep(300);
  await app.tab('הפניות');
  const clicked = await app.click('תוכן עניינים', { after: 4000, exact: true });
  await sleep(2000);
  const text = await app.screenText();
  const count = (t) => (text.match(new RegExp(t, 'g')) ?? []).length;
  log('תוכן עניינים: נלחץ', clicked, '| מופעים על המסך:', JSON.stringify({ h1: count(PG[1]), custom: count(PG[2]), normal: count(PG[0]) }));
  log('טקסט המסך:', text.slice(0, 400));
  const saved = await c.save('G-outline-toc');
  const doc = saved?.['word/document.xml'] ?? '';
  const instr = doc.match(/TOC [^<]*/)?.[0] ?? null;
  const tocTexts = (doc.match(/<w:hyperlink\b[\s\S]*?<\/w:hyperlink>/g) ?? []).map((h) => h.replace(/<[^>]+>/g, ''));
  log('instrText:', instr, '| ערכי TOC בקובץ:', JSON.stringify(tocTexts));
  const h1In = count(PG[1]) >= 2;
  const customIn = count(PG[2]) >= 2;
  if (!h1In) report.fail('רמת חלוקה — תוכן עניינים בעורך', `גם כותרת 1 לא נאספה (מופעים ${count(PG[1])}) — המדידה אינה תקפה`);
  else customIn
    ? report.pass('רמת חלוקה — תוכן עניינים בעורך', `הפסקה בסגנון המותאם נאספה (${count(PG[2])} מופעים על המסך, כותרת 1: ${count(PG[1])}); instr=${instr}`)
    : report.fail('רמת חלוקה — תוכן עניינים בעורך', `כותרת 1 נאספה (${count(PG[1])} מופעים), הפסקה בסגנון המותאם לא (${count(PG[2])}); instr=${instr}`);
});

/* ================================================================== */
/* H — קיצור מקשים לסגנון                                               */
/* ================================================================== */

let savedShortcutStorage = null;

await phase('H', 'קיצור מקשים לסגנון מובנה — הקלטה, החלה, רשימות, התנגשות', async (app, c) => {
  await c.caretIn(0);
  await c.type([P1, P2, P3]);
  await c.caretIn(P2);
  await c.cardMenu('כותרת 1');
  await c.pickMenu('modify');
  await c.waitDialog();
  const rec = await record(app, c, COMBO_Q);
  const ds = await dialogState(c);
  log('הקלטה:', JSON.stringify(rec), JSON.stringify(ds));
  rec.shown === COMBO_Q.label
    ? report.pass('קיצור — הלוכד מקליט בלחיצה אמיתית', `מוצג „${rec.shown}”; מיקוד אחרי הלחיצה: ${rec.focus?.active}`)
    : report.fail('קיצור — הלוכד מקליט בלחיצה אמיתית', `אחרי לחיצה ו-${COMBO_Q.label} הלוכד מציג „${rec.shown}”; מיקוד: ${JSON.stringify(rec.focus)}`);
  if (!rec.mouseOnly) {
    rec.withFocus === COMBO_Q.label
      ? report.pass('קיצור — הלוכד מקליט כשהוא במיקוד (Tab ואז לחיצה)', `מוצג „${rec.withFocus}” — ההמשך נמדד דרך המסלול הזה`)
      : report.fail('קיצור — הלוכד מקליט כשהוא במיקוד (Tab ואז לחיצה)', `מוצג „${rec.withFocus}”`);
  }
  const undoBefore = (await c.shell()).undo;
  const tag = await c.markInstance();
  await c.ok();
  await sleep(6000);
  const after = { same: await c.sameInstance(tag), dialog: await app.exists('[data-testid="style-dialog"]'), shell: await c.shell() };
  log('אחרי אישור (רק קיצור):', JSON.stringify(after), 'אחסון:', JSON.stringify(await storedShortcuts(c)));
  after.same && !after.dialog && after.shell.undo === undoBefore
    ? report.pass('קיצור — אין טעינה מחדש כשרק הקיצור השתנה', `אותו מופע אחרי 6s, undo נשאר ${after.shell.undo}; סטטוס „${after.shell.status}”`)
    : report.fail('קיצור — אין טעינה מחדש כשרק הקיצור השתנה', JSON.stringify(after));
  /נשמר/.test(after.shell.status ?? '')
    ? report.pass('קיצור — הודעת סטטוס', after.shell.status)
    : report.fail('קיצור — הודעת סטטוס', `status=${after.shell.status}`);

  // החלה בצירוף אמיתי על פסקה רגילה.
  await c.caretIn(P3);
  const beforeCard = await c.activeCard();
  await pressCombo(app, COMBO_Q);
  await sleep(2500);
  const card = await c.activeCard();
  const p3 = paraBy(await c.engineExport('H החלה'), P3);
  const st = await c.shell();
  log(`צירוף על P3: כרטיס ${beforeCard} → ${card}; pStyle=${p3?.pStyle}; status=${st.status}`);
  p3?.pStyle === 'Heading1' && card === 'כותרת 1'
    ? report.pass('קיצור — הצירוף מחיל את הסגנון', `P3: ${beforeCard} → ${card}, pStyle=Heading1; סטטוס „${st.status}”`)
    : report.fail('קיצור — הצירוף מחיל את הסגנון', `pStyle=${p3?.pStyle}, כרטיס ${card}, סטטוס ${st.status}`);

  const list = await managerList(app, c);
  log('ניהול קיצורים:', JSON.stringify(list));
  const entry = (list ?? []).find((e) => e.summary === 'החלת הסגנון „כותרת 1”');
  entry && entry.combo === COMBO_Q.label
    ? report.pass('קיצור — ב„ניהול קיצורים”', `${entry.name} | ${entry.combo} | ${entry.summary}`)
    : report.fail('קיצור — ב„ניהול קיצורים”', JSON.stringify(list));
  const help = await helpRows(app, c);
  log('דיאלוג הקיצורים:', JSON.stringify(help));
  help.rows.some((t) => t.includes('החלת הסגנון „כותרת 1”'))
    ? report.pass('קיצור — בדיאלוג רשימת הקיצורים', help.rows.find((t) => t.includes('החלת הסגנון „כותרת 1”')))
    : report.fail('קיצור — בדיאלוג רשימת הקיצורים', `נפתח=${help.open}; ${JSON.stringify(help.rows)}`);

  // התנגשות: מובנה (Ctrl+B), וקיצור של סגנון אחר.
  for (const [what, k] of [['פעולה מובנית', COMBO_B], ['קיצור של סגנון אחר', COMBO_Q]]) {
    await c.cardMenu('כותרת 2');
    await c.pickMenu('modify');
    await c.waitDialog();
    const rr = await record(app, c, k);
    const d = await dialogState(c);
    log(`התנגשות (${what}):`, JSON.stringify(rr), JSON.stringify(d));
    d?.okDisabled && /[א-ת]/.test(d.error ?? '')
      ? report.pass(`קיצור — התנגשות עם ${what}`, `„אישור” מושבת: ${d.error}`)
      : report.fail(`קיצור — התנגשות עם ${what}`, `הלוכד „${rr.withFocus ?? rr.shown}”, אישור מושבת=${d?.okDisabled}, הודעה=${d?.error}`);
    await app.press('Escape', 'Escape', 27);
    await sleep(300);
    if (await app.exists('[data-testid="style-dialog"]')) await c.clickSel('[data-testid="style-dialog"] .sd-close-btn');
  }

  savedShortcutStorage = await storedShortcuts(c);
  log('נשמר באחסון המאחז:', JSON.stringify(savedShortcutStorage));
});

/* H2 — אותו אחסון בטעינה חדשה של התוסף (המאחז מחזיק אותו, כמו אוצריא). */
if (savedShortcutStorage) {
  const seed = `<script>window.__qaHost.storage['custom-shortcuts']=${JSON.stringify(savedShortcutStorage)};</script>`;
  await phase('H', 'קיצור מקשים — אחרי טעינה מחדש של התוסף', async (app, c) => {
    const list = await managerList(app, c);
    await c.caretIn(0);
    await c.type([P1]);
    await c.caretIn(P1);
    await pressCombo(app, COMBO_Q);
    await sleep(2500);
    const p1 = paraBy(await c.engineExport('H2'), P1);
    log('אחרי טעינה:', JSON.stringify(list), 'pStyle=', p1?.pStyle);
    (list ?? []).some((e) => e.summary === 'החלת הסגנון „כותרת 1”') && p1?.pStyle === 'Heading1'
      ? report.pass('קיצור — שורד טעינה מחדש', 'נטען מהאחסון, מופיע ברשימה, והצירוף מחיל כותרת 1')
      : report.fail('קיצור — שורד טעינה מחדש', `רשימה ${JSON.stringify(list)}, pStyle=${p1?.pStyle}`);
  }, { extra: seed });
} else if (!ONLY || ONLY.has('H')) {
  report.skip('קיצור — שורד טעינה מחדש', 'לא נשמר דבר באחסון המאחז בשלב הקודם');
}

/* ================================================================== */
/* I — שינוי שם מסנכרנים את הקיצור                               */
/* ================================================================== */

await phase('I', 'קיצור לסגנון מותאם — שינוי שם', async (app, c) => {
  await c.caretIn(0);
  await c.type([P1, P2, P3]);
  await c.caretIn(P3);
  await c.moreMenu();
  await c.pickMenu('create');
  await c.waitDialog();
  await c.fill('#sd-name', 'ציטוט מקור');
  const rec = await record(app, c, COMBO_W);
  log('הקלטה ביצירה:', JSON.stringify(rec));
  let tag = await c.markInstance();
  await c.ok();
  let r = await c.waitApply(tag);
  let list = await managerList(app, c);
  log('אחרי יצירה:', JSON.stringify(r), JSON.stringify(list), JSON.stringify(await storedShortcuts(c)));
  const created = (list ?? []).find((e) => e.combo === COMBO_W.label);
  created?.summary === 'החלת הסגנון „ציטוט מקור”'
    ? report.pass('קיצור — נוצר עם הסגנון', `${created.name} | ${created.combo} | ${created.summary}`)
    : report.fail('קיצור — נוצר עם הסגנון', JSON.stringify(list));

  // הצירוף מחיל את הסגנון המותאם על פסקה אחרת.
  await c.caretIn(P1);
  await pressCombo(app, COMBO_W);
  await sleep(2500);
  const files = await c.engineExport('I החלה');
  const id = idOf(styleByName(files['word/styles.xml'], 'ציטוט מקור'));
  const p1 = paraBy(files, P1);
  p1?.pStyle && p1.pStyle === id
    ? report.pass('קיצור — מחיל סגנון מותאם', `P1 pStyle=${id}`)
    : report.fail('קיצור — מחיל סגנון מותאם', `P1 pStyle=${p1?.pStyle}, id=${id}`);

  await c.cardMenu('ציטוט מקור');
  await c.pickMenu('rename');
  await c.waitDialog();
  await c.fill('#sd-name', 'ציטוט מקור חדש');
  tag = await c.markInstance();
  await c.ok();
  r = await c.waitApply(tag);
  list = await managerList(app, c);
  log('אחרי שינוי שם:', JSON.stringify(r), JSON.stringify(list));
  const renamed = (list ?? []).find((e) => e.combo === COMBO_W.label);
  renamed?.summary === 'החלת הסגנון „ציטוט מקור חדש”' && renamed?.name === 'ציטוט מקור חדש'
    ? report.pass('קיצור — עוקב אחרי שינוי שם', `${renamed.name} | ${renamed.summary}`)
    : report.fail('קיצור — עוקב אחרי שינוי שם', JSON.stringify(renamed ?? list));


});

process.exit(report.print() > 0 ? 1 : 0);

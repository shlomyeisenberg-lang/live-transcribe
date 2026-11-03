/**
 * ברירות המחדל של מסמך חדש — מסלול המשתמש מקצה לקצה, בשתי הפעלות:
 *
 * 1. הפעלה בלי הגדרות: המסמך החדש מיושר לשני הצדדים (`w:jc="both"` ב„רגיל”).
 *    „ברירות מחדל” עם „גם לכל מסמך חדש”, ו„שינוי סגנון” של כותרת 1 עם אותו
 *    סימון — בלחיצות אמיתיות — נשמרים ל-storage של המאחז.
 * 2. הפעלה חדשה עם מה שנשמר: המסמך החדש נפתח איתם, ב-`styles.xml` שיוצא.
 *    „החזר למקור” מוחק את הרשומה.
 *
 * `SHOTS=<תיקייה>` שומר צילומי מסך של המסמך ושל הדיאלוגים, ו-`DOCX_OUT=<קובץ>`
 * את המסמך החדש של ההפעלה השנייה — לפתיחה ב-Word.
 */
import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, sleep, createReport } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9743);
const SHOTS = process.env.SHOTS ?? '';
const KEY = 'new-document-defaults';
const report = createReport('ברירות מחדל למסמך חדש', { strict: true });

const TEXT =
  'בראשית ברא אלהים את השמים ואת הארץ והארץ היתה תהו ובהו וחשך על פני תהום ורוח אלהים מרחפת על פני המים ויאמר אלהים יהי אור ויהי אור וירא אלהים את האור כי טוב ויבדל אלהים בין האור ובין החשך';

function seed(storage) {
  return `<script>(function(){var H=window.__qaHost;var v=${JSON.stringify(storage)};Object.keys(v).forEach(function(k){H.storage[k]=v[k];});})();</script>`;
}

function helpers(app) {
  const ask = async (code) => JSON.parse(await app.js(`(async()=>JSON.stringify(await (${code})))()`));
  const until = async (predicate, what) => {
    for (let i = 0; i < 60; i++) {
      if (await predicate()) return;
      await sleep(250);
    }
    throw new Error(`לא הגיע: ${what}`);
  };
  const fill = (selector, value) =>
    app.js(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)throw Error('missing ${selector}');e.value=${JSON.stringify(value)};e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}))})()`);
  const click = async (selector) => {
    await app.js(`document.querySelector(${JSON.stringify(selector)})?.scrollIntoView({block:'nearest',inline:'nearest'})`);
    return app.clickSel(selector, 0, { after: 150 });
  };
  const rightClick = async (selector) => {
    const r = await ask(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'nearest',inline:'nearest'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2}})()`);
    for (const type of ['mouseMoved', 'mousePressed', 'mouseReleased']) {
      await app.cdp.send('Input.dispatchMouseEvent', { type, x: Math.round(r.x), y: Math.round(r.y), button: type === 'mouseMoved' ? 'none' : 'right', buttons: type === 'mousePressed' ? 2 : 0, clickCount: 1 });
      await sleep(50);
    }
    await sleep(250);
  };
  const shot = async (name) => {
    if (!SHOTS) return;
    const data = (await app.cdp.send('Page.captureScreenshot', { format: 'png' }))?.result?.data;
    writeFileSync(join(SHOTS, `${name}.png`), Buffer.from(data, 'base64'));
  };
  const stored = () => ask(`Promise.resolve(window.__qaHost.storage[${JSON.stringify(KEY)}] ?? null)`);
  return { ask, until, fill, click, rightClick, shot, stored };
}

function style(xml, id) {
  return new RegExp(`<w:style [^>]*w:styleId="${id}">[\\s\\S]*?</w:style>`).exec(xml)?.[0] ?? '';
}

/** קצוות השורות של הפסקה הראשונה ביחס לעמוד — שורה מיושרת לשני הצדדים נוגעת בשניהם. */
async function lineEdges(app) {
  return JSON.parse(await app.js(`JSON.stringify((()=>{
    const lines=[...document.querySelectorAll('.editor-stack .superdoc-line')];
    return lines.slice(0,4).map(l=>{const runs=[...l.querySelectorAll('.superdoc-text-run, span')].map(s=>s.getBoundingClientRect()).filter(r=>r.width>0);
      if(!runs.length) return null; return {left:Math.min(...runs.map(r=>r.left)),right:Math.max(...runs.map(r=>r.right))};});
  })())`));
}

async function stage(name, storage, body) {
  console.log(`\n────── ${name} ──────`);
  let app = null;
  try {
    app = await openApp({ name: `ndd-${name}`, port: PORT, extra: seed(storage), settle: 1500 });
    await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 900, deviceScaleFactor: 1, mobile: false });
    await sleep(500);
    return await body(app, helpers(app));
  } catch (error) {
    report.fail(`${name} — השלב נכשל`, error.stack ?? error.message);
    return null;
  } finally {
    if (app) app.close();
  }
}

/* 1. בלי הגדרות — יישור לשני הצדדים, ושמירה מהדיאלוגים. */
const saved = await stage('pristine', {}, async (app, h) => {
  const files = await app.docx();
  const normal = style(files['word/styles.xml'], 'Normal');
  normal.includes('<w:jc w:val="both"/>')
    ? report.pass('מסמך חדש: „רגיל” מיושר לשני הצדדים', 'w:jc="both" בסגנון')
    : report.fail('מסמך חדש: „רגיל” מיושר לשני הצדדים', normal);
  /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*<w:jc /.test(files['word/document.xml'])
    ? report.fail('בלי עיצוב ישיר על הפסקה', 'נמצא w:jc בפסקה')
    : report.pass('בלי עיצוב ישיר על הפסקה');

  await app.tab('בית');
  await app.caretPara('', { verifyMs: 10000 });
  await app.type(TEXT, 5);
  await sleep(800);
  const edges = (await lineEdges(app)).filter(Boolean);
  const full = edges.slice(0, -1);
  const spread = full.length ? Math.max(...full.map((e) => Math.abs(e.left - edges[0].left) + Math.abs(e.right - edges[0].right))) : NaN;
  full.length >= 1 && spread < 3
    ? report.pass('השורות המלאות נוגעות בשני השוליים', JSON.stringify(edges))
    : report.fail('השורות המלאות נוגעות בשני השוליים', JSON.stringify(edges));
  await h.shot('after-new-document');

  // „ברירות מחדל” — גודל 14, וגם לכל מסמך חדש.
  await app.tab('פריסה');
  assert.equal(await app.click('ברירות מחדל'), true);
  await h.until(() => app.exists('.docdef-dialog'), 'דיאלוג ברירות מחדל');
  await h.fill('#dd-size', '14');
  assert.equal(await h.click('[data-testid="docdef-new-documents"]'), true);
  await sleep(200);
  await h.shot('docdefaults-dialog');
  assert.equal(await h.click('.docdef-dialog .dd-btn-primary'), true);
  await h.until(async () => (await h.stored())?.run?.fontSizePt === 14, 'run.fontSizePt=14 ב-storage');
  report.pass('„ברירות מחדל” + „גם לכל מסמך חדש” נשמר', JSON.stringify(await h.stored()));

  // „שינוי סגנון” של כותרת 1 — גודל 18 וצבע, וגם לכל מסמך חדש.
  await app.tab('בית');
  const card = await h.ask(`Promise.resolve([...document.querySelectorAll('.style-card')].map(c=>c.getAttribute('data-tip-title')))`);
  const heading = card.find((label) => /כותרת 1|Heading 1/i.test(label ?? ''));
  assert.ok(heading, `אין כרטיס כותרת 1 בגלריה: ${card}`);
  await h.rightClick(`.style-card[data-tip-title="${heading}"]`);
  assert.equal(await h.click('[data-testid="style-menu"] [data-action="modify"]'), true);
  await h.until(() => app.exists('[data-testid="style-dialog"]'), 'דיאלוג סגנון');
  assert.equal(await app.exists('[data-testid="style-new-documents"]'), true, 'אין „גם לכל מסמך חדש” בדיאלוג הסגנון');
  await h.fill('[data-testid="style-dialog"] input[aria-label="גודל, בנקודות"]', '18');
  await h.fill('[data-testid="style-dialog"] input[aria-label="צבע גופן"]', '#1f3864');
  assert.equal(await h.click('[data-testid="style-new-documents"]'), true);
  await sleep(200);
  await h.shot('style-dialog');
  assert.equal(await h.click('[data-testid="style-dialog"] .sd-btn-primary'), true);
  await h.until(async () => (await h.stored())?.styles?.Heading1?.fontSizePt === 18, 'Heading1 ב-storage');
  const record = await h.stored();
  JSON.stringify(record.styles) === JSON.stringify({ Heading1: { fontSizePt: 18, color: '1F3864' } })
    ? report.pass('כותרת 1 נשמרה כטלאי ביחס לתבנית', JSON.stringify(record.styles))
    : report.fail('כותרת 1 נשמרה כטלאי ביחס לתבנית', JSON.stringify(record.styles));
  return record;
});

/* 2. הפעלה חדשה עם מה שנשמר. */
if (saved) {
  await stage('seeded', { [KEY]: saved }, async (app, h) => {
    const files = await app.docx();
    const xml = files['word/styles.xml'];
    const run = /<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/.exec(xml)?.[0] ?? '';
    run.includes('<w:sz w:val="28"/><w:szCs w:val="28"/>')
      ? report.pass('מסמך חדש: גודל ברירת המחדל 14 בשני צדי הכתב')
      : report.fail('מסמך חדש: גודל ברירת המחדל 14 בשני צדי הכתב', run);
    const heading = style(xml, 'Heading1');
    heading.includes('<w:sz w:val="36"/>') && heading.includes('<w:szCs w:val="36"/>') && heading.includes('w:val="1F3864"')
      ? report.pass('מסמך חדש: כותרת 1 בגודל 18 ובצבע שנבחר')
      : report.fail('מסמך חדש: כותרת 1 בגודל 18 ובצבע שנבחר', heading);
    style(xml, 'Normal').includes('<w:jc w:val="both"/>')
      ? report.pass('היישור לשני הצדדים נשאר')
      : report.fail('היישור לשני הצדדים נשאר', style(xml, 'Normal'));

    await app.tab('בית');
    await app.caretPara('', { verifyMs: 10000 });
    await app.type('כותרת', 20);
    const card = await h.ask(`Promise.resolve([...document.querySelectorAll('.style-card')].map(c=>c.getAttribute('data-tip-title')))`);
    const label = card.find((l) => /כותרת 1|Heading 1/i.test(l ?? ''));
    await h.click(`.style-card[data-tip-title="${label}"]`);
    await app.press('End', 'End', 35);
    await app.press('Enter', 'Enter', 13, 0, '\r');
    await app.type(TEXT, 5);
    await sleep(800);
    await h.shot('seeded-new-document');
    if (process.env.DOCX_OUT) {
      const base64 = JSON.parse(await app.js(`window.__qa.exportBase64().then(JSON.stringify)`));
      writeFileSync(process.env.DOCX_OUT, Buffer.from(base64, 'base64'));
    }

    await app.tab('פריסה');
    assert.equal(await app.click('ברירות מחדל'), true);
    await h.until(() => app.exists('[data-testid="docdef-reset-new-documents"]'), 'סיכום ההגדרות בדיאלוג');
    await h.shot('docdefaults-dialog-saved');
    assert.equal(await h.click('[data-testid="docdef-reset-new-documents"]'), true);
    await h.until(async () => (await h.stored()) === null, 'הרשומה נמחקה');
    report.pass('„החזר למקור” מוחק את ההגדרות');
  });
}

/* 3. Old or hand-edited storage cannot introduce Word/renderer disagreement. */
await stage('unsupported-stored', { [KEY]: { run: { fontSizePt: 15 }, styles: { Normal: { italic: true } } } }, async (app, h) => {
  const xml = (await app.docx())['word/styles.xml'];
  assert.ok(!/<w:i(?:Cs)?\b/.test(style(xml, 'Normal')), 'unsafe Normal italic reached the new document');
  assert.ok(style(xml, 'Heading4').includes('<w:iCs/>'), 'original Heading4 formatting changed');
  assert.ok(/<w:rPrDefault>[\s\S]*?<w:szCs w:val="30"\/>/.test(xml), 'safe run size was discarded');
  report.pass('העדפת סגנון לא נתמכת אינה מוחלת; גודל בטוח נשמר');
  const status = await h.ask(`Promise.resolve(document.querySelector('#status .status-message')?.textContent ?? '')`);
  assert.ok(status.includes('לא הוחלו'), `missing visible warning: ${status}`);
  report.pass('אזהרה גלויה על סגנונות שמורים שלא הוחלו');
  await app.tab('פריסה');
  assert.equal(await app.click('ברירות מחדל'), true);
  await h.until(() => app.exists('[data-testid="docdef-reset-new-documents"]'), 'סיכום ההעדפות שנדחו');
  const summary = await h.ask(`Promise.resolve(document.querySelector('.docdef-dialog')?.textContent ?? '')`);
  assert.ok(summary.includes('לא הוחלו'), 'dialog claims incompatible styles were applied');
  assert.equal(await h.click('[data-testid="docdef-reset-new-documents"]'), true);
  await h.until(async () => (await h.stored()) === null, 'איפוס ההעדפות שנדחו');
  report.pass('סיכום ההעדפות מסביר את הסירוב ואפשר לאפס אותן');
});

process.exit(report.print() > 0 ? 1 : 0);

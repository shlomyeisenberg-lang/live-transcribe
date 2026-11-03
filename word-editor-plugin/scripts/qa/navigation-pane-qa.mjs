/**
 * שער: „חלונית ניווט” — הכותרות של המסמך בצד, ולחיצה עוברת אליהן.
 *
 * על ה-dist הארוז ובלחיצות אמיתיות:
 *
 *  1. הכפתור בלשונית „תצוגה” פותח חלונית שמציגה את הכותרות לפי הסדר, עם הרמות.
 *  2. לחיצה על כותרת שמחוץ למסך מעבירה אליה את הסמן **ואת התצוגה**, והפרק
 *     הנוכחי מסומן.
 *  3. אחרי הלחיצה אפשר להקליד מיד — המיקוד חזר למסמך — והחלונית מתעדכנת
 *     מהעריכה.
 *  4. קיפול מסתיר את תתי-הכותרות, והחלת סגנון כותרת על פסקה מוסיפה אותה.
 *  5. החלונית נזכרת: הכתיבה ל-storage נעשית.
 *
 *   npm run build && node scripts/qa/navigation-pane-qa.mjs
 */
import { writeFileSync } from 'node:fs';
import { openApp, createReport } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9398);
const SHOT = process.env.QA_SHOT ?? '';
const report = createReport('חלונית ניווט', { strict: true });
const log = (...a) => console.log('   ', ...a);

async function paneItems(app) {
  return JSON.parse(
    await app.js(
      `JSON.stringify([...document.querySelectorAll('.nav-pane .nav-item')].map(function(el){` +
        `return {text:el.querySelector('.nav-label').textContent.trim(),level:Number(el.getAttribute('aria-level')),` +
        `current:el.classList.contains('nav-item--current'),expanded:el.getAttribute('aria-expanded')};}))`,
    ),
  );
}

/** האם השורה של הפסקה גלויה בתוך אזור המסמך. */
async function lineVisible(app, text) {
  return JSON.parse(
    await app.js(
      `(function(){var stack=document.querySelector('.editor-stack').getBoundingClientRect();` +
        `var L=document.querySelectorAll('.superdoc-line');` +
        `for(var i=0;i<L.length;i++){if((L[i].textContent||'').indexOf(${JSON.stringify(text)})<0)continue;` +
        `var r=L[i].getBoundingClientRect();` +
        `return JSON.stringify({top:Math.round(r.top-stack.top),visible:r.bottom>stack.top&&r.top<stack.bottom});}` +
        `return JSON.stringify(null);})()`,
    ),
  );
}

async function itemIndex(app, text) {
  return (await paneItems(app)).findIndex((item) => item.text.includes(text));
}

async function waitFor(app, predicate, ms = 8000) {
  const end = Date.now() + ms;
  let last;
  while (Date.now() < end) {
    last = await paneItems(app);
    if (predicate(last)) return last;
    await app.sleep(250);
  }
  return last;
}

const app = await openApp({ name: 'navigation-pane', port: PORT });
try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1500,
    height: 900,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await app.sleep(2000);

  // פרק שני רחוק מספיק כדי שלא ייראה בלי גלילה.
  await app.caret(0);
  // הכותרות העליונות מוחלות לפני שהמסמך מתארך: `caretPara` לוחץ על שורה
  // מצוירת, ופסקה שמחוץ למסך אינה כזאת.
  for (const line of ['פרק ראשון', 'סעיף אלף', 'גוף הטקסט']) {
    await app.type(line);
    await app.press('Enter', 'Enter', 13);
    await app.sleep(300);
  }
  await app.sleep(600);
  await app.tab('בית');
  for (const [para, style] of [
    ['פרק ראשון', 'כותרת 1'],
    ['סעיף אלף', 'כותרת 2'],
  ]) {
    await app.caretPara(para);
    await app.sleep(250);
    if (!(await app.clickGallery(style, { after: 900 }))) report.fail(`„${style}”`, 'הכרטיס לא נמצא');
  }

  // ואז פרק שני, רחוק מספיק כדי שלא ייראה בלי גלילה. הסמן נשאר בו.
  await app.caretPara('גוף הטקסט');
  await app.press('End', 'End', 35);
  for (let i = 0; i < 45; i++) {
    await app.press('Enter', 'Enter', 13);
    await app.sleep(120);
  }
  await app.type('פרק שני');
  await app.sleep(900);
  if (!(await app.clickGallery('כותרת 1', { after: 900 }))) report.fail('„כותרת 1”', 'הכרטיס לא נמצא');
  // חזרה לראש המסמך (Ctrl+Home), כדי שפרק שני יהיה מחוץ למסך כשקופצים אליו.
  await app.press('Home', 'Home', 36, 2);
  await app.sleep(800);

  /* 1 — פתיחה */
  await app.tab('תצוגה');
  if (!(await app.exists('.doc-ruler__page'))) await app.click('סרגל', { after: 400 });
  const rulerAligned = async () => JSON.parse(await app.js(`JSON.stringify((() => {
    const ruler = document.querySelector('.doc-ruler__page')?.getBoundingClientRect();
    const page = document.querySelector('.superdoc-page')?.getBoundingClientRect();
    return { left: ruler && page ? Math.abs(ruler.left - page.left) : null, width: ruler && page ? Math.abs(ruler.width - page.width) : null };
  })())`));
  const rulerBefore = await rulerAligned();
  const opened = await app.click('חלונית ניווט', { after: 1500 });
  const rulerAfter = await rulerAligned();
  [rulerBefore, rulerAfter].every((r) => r.left !== null && r.left < 1 && r.width < 1)
    ? report.pass('הסרגל נשאר מיושר לעמוד בפתיחת החלונית', JSON.stringify({ rulerBefore, rulerAfter }))
    : report.fail('יישור הסרגל', JSON.stringify({ rulerBefore, rulerAfter }));
  if (!opened) report.fail('הכפתור', '„חלונית ניווט” לא נמצא בלשונית „תצוגה”');
  const first = await waitFor(app, (items) => items.length >= 3);
  log('פריטים:', JSON.stringify(first));
  const shape = first.map((item) => `${item.level}:${item.text}`).join(' | ');
  shape === '1:פרק ראשון | 2:סעיף אלף | 1:פרק שני'
    ? report.pass('החלונית מציגה את הכותרות לפי הסדר והרמה', shape)
    : report.fail('רשימת הכותרות שגויה', shape);

  // F6 enters the tree. RTL arrows change focus and collapse without moving
  // the document caret; Enter/Space return to editing through the public API.
  await app.js(`document.querySelector('.word-ribbon-container button').focus(); true`);
  await app.press('F6', 'F6', 117);
  const focusedHeading = async () => await app.js(`document.activeElement?.matches('.nav-item') ? document.activeElement.querySelector('.nav-label').textContent.trim() : ''`);
  const entered = await focusedHeading();
  await app.press('ArrowRight', 'ArrowRight', 39);
  const collapsedByKey = (await paneItems(app)).length === 2;
  await app.press('ArrowDown', 'ArrowDown', 40);
  const down = await focusedHeading();
  await app.press('Home', 'Home', 36);
  await app.press('ArrowLeft', 'ArrowLeft', 37);
  await app.press('ArrowLeft', 'ArrowLeft', 37);
  const child = await focusedHeading();
  await app.press('ArrowRight', 'ArrowRight', 39);
  const parent = await focusedHeading();
  await app.press('End', 'End', 35);
  const end = await focusedHeading();
  await app.press('ArrowUp', 'ArrowUp', 38);
  await app.press('Enter', 'Enter', 13);
  await app.sleep(400);
  const keyCaret = await app.caretBlock();
  const keyboard = entered === 'פרק ראשון' && collapsedByKey && down === 'פרק שני' && child === 'סעיף אלף' && parent === 'פרק ראשון' && end === 'פרק שני' && keyCaret?.text?.includes('סעיף אלף');
  keyboard
    ? report.pass('F6 ועץ הכותרות נגישים במקלדת, כולל קיפול, הורה/ילד והפעלה', '')
    : report.fail('ניווט מקלדת', JSON.stringify({ entered, collapsedByKey, down, child, parent, end, keyCaret }));
  await app.press('F6', 'F6', 117, 8);
  (await focusedHeading()) === 'סעיף אלף'
    ? report.pass('Shift+F6 מחזיר מהמסמך לכותרת הממוקדת', '')
    : report.fail('Shift+F6 לחלונית', await focusedHeading());
  await app.press(' ', 'Space', 32, 0, ' ');
  await app.sleep(300);
  const spaceCaret = await app.caretBlock();
  spaceCaret?.text?.includes('סעיף אלף')
    ? report.pass('Space מפעיל כותרת בלי להוסיף רווח למסמך', spaceCaret.text)
    : report.fail('הפעלת כותרת ב-Space', JSON.stringify(spaceCaret));

  const paneRect = JSON.parse(
    await app.js(
      `(function(){var p=document.querySelector('.nav-pane').getBoundingClientRect();` +
        `var s=document.querySelector('.editor-stack').getBoundingClientRect();` +
        `return JSON.stringify({paneLeft:Math.round(p.left),stackRight:Math.round(s.right),width:Math.round(p.width)});})()`,
    ),
  );
  paneRect.paneLeft >= paneRect.stackRight - 1
    ? report.pass('החלונית בצד ימין של המסמך', JSON.stringify(paneRect))
    : report.fail('החלונית אינה בצד ימין', JSON.stringify(paneRect));

  /* 2 — קפיצה לכותרת מחוץ למסך */
  const before = await lineVisible(app, 'פרק שני');
  log('פרק שני לפני:', JSON.stringify(before));
  await app.clickSel('.nav-pane .nav-item', await itemIndex(app, 'פרק שני'), { after: 1200 });
  const landed = await app.caretBlock();
  const after = await lineVisible(app, 'פרק שני');
  log('סמן:', JSON.stringify(landed), 'שורה:', JSON.stringify(after));
  landed?.text?.includes('פרק שני')
    ? report.pass('הסמן עבר לכותרת', landed.text)
    : report.fail('הסמן לא עבר לכותרת', JSON.stringify(landed));
  after?.visible
    ? report.pass('התצוגה נגללה אל הכותרת', `לפני: ${JSON.stringify(before)}, אחרי: ${JSON.stringify(after)}`)
    : report.fail('הכותרת אינה גלויה אחרי הלחיצה', JSON.stringify(after));
  const marked = await waitFor(app, (items) => items.find((i) => i.current)?.text === 'פרק שני', 4000);
  marked.find((i) => i.current)?.text === 'פרק שני'
    ? report.pass('הפרק הנוכחי מסומן', 'פרק שני')
    : report.fail('הסימון אינו על הפרק הנוכחי', JSON.stringify(marked));

  /* 3 — הקלדה מיד אחרי הלחיצה, והחלונית מתעדכנת */
  await app.type('חדש ');
  const typed = await waitFor(app, (items) => items.some((i) => i.text === 'חדש פרק שני'));
  typed.some((i) => i.text === 'חדש פרק שני')
    ? report.pass('אפשר להקליד מיד, והחלונית מתעדכנת', JSON.stringify(typed.map((i) => i.text)))
    : report.fail('ההקלדה לא הגיעה לכותרת או לחלונית', JSON.stringify(typed.map((i) => i.text)));

  /* תזוזת סמן בלי עריכה מזיזה את הסימון */
  // „גוף הטקסט” מחוץ למסך עכשיו: מגיעים דרך החלונית לכותרת שמעליו, ומשם בחץ.
  await app.clickSel('.nav-pane .nav-item', await itemIndex(app, 'סעיף אלף'), { after: 900 });
  await app.press('ArrowDown', 'ArrowDown', 40);
  await app.sleep(400);
  const moved = await waitFor(app, (items) => items.find((i) => i.current)?.text === 'סעיף אלף', 4000);
  moved.find((i) => i.current)?.text === 'סעיף אלף'
    ? report.pass('הסימון עוקב אחרי הסמן', 'גוף הטקסט ← סעיף אלף')
    : report.fail('הסימון לא עקב אחרי הסמן', JSON.stringify(moved));

  /* 4 — קיפול, והוספת כותרת */
  await app.clickSel('.nav-pane .nav-twisty', 0, { after: 500 });
  const folded = await paneItems(app);
  folded.map((i) => i.text).join('|') === 'פרק ראשון|חדש פרק שני' && folded[0].expanded === 'false'
    ? report.pass('קיפול מסתיר את תת-הכותרת', JSON.stringify(folded.map((i) => i.text)))
    : report.fail('הקיפול לא עבד', JSON.stringify(folded));
  await app.clickSel('.nav-pane .nav-twisty', 0, { after: 500 });

  // הסמן עדיין ב„גוף הטקסט” — החצים והקיפול אינם מזיזים אותו.
  const still = await app.caretBlock();
  if (!still?.text?.includes('גוף הטקסט')) report.fail('הסמן זז בקיפול', JSON.stringify(still));
  await app.tab('בית');
  const gallery = (await app.galleryItems()).map((item) => item.label ?? item);
  const h3 = await app.clickGallery('כותרת 3', { after: 900 });
  const xml = (await app.docx())['word/document.xml'] ?? '';
  const bodyPara = (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).find((p) => p.includes('גוף הטקסט')) ?? '';
  log('גלריה:', JSON.stringify(gallery), 'נלחץ:', h3, 'סגנון:', bodyPara.match(/<w:pStyle w:val="([^"]+)"/)?.[1] ?? '(אין)');
  const grown = await waitFor(app, (items) => items.some((i) => i.text === 'גוף הטקסט'));
  const grownShape = grown.map((i) => `${i.level}:${i.text}`).join(' | ');
  grownShape === '1:פרק ראשון | 2:סעיף אלף | 3:גוף הטקסט | 1:חדש פרק שני'
    ? report.pass('כותרת חדשה נכנסת במקומה ובעומק הנכון', grownShape)
    : report.fail('הכותרת החדשה לא נכנסה כראוי', grownShape);

  if (SHOT) {
    const shot = await app.cdp.send('Page.captureScreenshot', { format: 'png' });
    if (shot?.result?.data) writeFileSync(SHOT, Buffer.from(shot.result.data, 'base64'));
    log('צילום:', SHOT);
  }

  /* 5 — נזכרת */
  const calls = await app.hostCalls();
  const saved = calls.filter(
    (c) => c.method === 'storage.set' && c.payload?.key === 'navigation-pane-visible',
  );
  saved.some((c) => c.payload.value === true)
    ? report.pass('הפתיחה נשמרת להפעלה הבאה', JSON.stringify(saved.map((c) => c.payload.value)))
    : report.fail('הפתיחה לא נשמרה', JSON.stringify(calls.map((c) => c.method).slice(-20)));

  /* סגירה */
  await app.js(`document.querySelector('.nav-pane .nav-close').focus(); true`);
  // Native button activation needs the Enter character phase, not only a raw key.
  await app.press('Enter', 'Enter', 13, 0, '\r');
  await app.sleep(600);
  (await app.js(`!!document.activeElement?.closest('.editor-stack')`))
    ? report.pass('סגירת החלונית במקלדת מחזירה מיקוד למסמך', '')
    : report.fail('מיקוד אחרי סגירה', await app.js(`document.activeElement?.outerHTML.slice(0, 200)`));
  (await app.exists('.nav-pane'))
    ? report.fail('סגירה', 'החלונית עדיין פתוחה')
    : report.pass('כפתור הסגירה סוגר את החלונית', '');
} catch (error) {
  report.fail('חריגה', String(error?.stack ?? error));
} finally {
  app.close();
}

report.print();

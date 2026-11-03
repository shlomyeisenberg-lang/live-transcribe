/**
 * קובץ Word שתוסף אחר שולח (`plugin.page_opened`), בדפדפן אמיתי ועל מנוע אמיתי.
 *
 * קבלה, שחזור והמתנה לפעולות שכבר רצות:
 *
 * 1. **לפני הבאנדל.** „סייר קבצים” פותח את הלשונית, ואוצריא מוסרת את האירוע
 *    מיד בסוף ה-boot — לפני שהמאזין ב-App.vue קיים. האירוע נורה כאן מתוך הדף
 *    עצמו, אחרי `plugin.boot` ולפני שהבאנדל רץ, וה-latch שב-`index.html` הוא
 *    היחיד שיכול לתפוס אותו. מה שנבדק הוא שהמסמך **נפתח עם התוכן שלו**, ולא
 *    רק שהופיע טאב.
 * 2. **כשהעורך כבר חי** — דרך המאזין החי (`__qaHost.emit`, כמו ה-SDK).
 *    המסמך הראשון מסומן כלא-שמור, ולכן השני נפתח בטאב נוסף ולא דורס אותו.
 * 3. **בלי תוכן** — `{ name, path }`, מה שהסייר שולח היום. הודעה בשורת
 *    המצב, ושום טאב לא נפתח.
 * 4. **בלי קובץ בכלל** — כפתור „לשונית חדשה” של אוצריא. שום דבר לא משתנה.
 * וגם: הפעלה מחדש ללא עריכה, אירוע תוך כדי יציאה/סגירה/שמירה, כשל כתיבה
 * וביטול פתיחה בזמן שקריאת הבייטים עוד לא הסתיימה.
 *
 * הרצה: npm run build && node scripts/qa/received-file-qa.mjs
 */
import { openApp, createReport, sleep } from './harness.mjs';
import { buildDocx, p } from './docx-fixtures.mjs';

const PORT = Number(process.env.QA_PORT ?? 9671);
const EARLY_TEXT = 'נשלח מהסייר לפני שהעורך עלה';
const LIVE_TEXT = 'נשלח מהסייר כשהעורך כבר פתוח';

const dataUrl = (text) =>
  `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${buildDocx({ body: p(text) }).toString('base64')}`;

/*
 * אחרי הדמה של אוצריא (שיורה `plugin.boot` בעצמו) ולפני הבאנדל — כמו
 * `PluginPageLauncher.markPageReady`, שמוסר את האירועים הממתינים מיד אחרי
 * הזרקת ה-boot. ההזרקה יושבת **לפני** ה-latch בדף, ולכן הירייה ממתינה שהוא
 * ייפרס; הבאנדל הוא מודול, כלומר רץ רק אחרי סוף הפריסה. `__qaEarly` מתעד אם
 * המאזין החי כבר היה שם ברגע הירייה — אם כן, השער בודק את המסלול הלא נכון.
 */
// המרחב הפרטי ורשומת ההפעלה שורדים reload, כמו במארח. כך נבדק שחזור
// מהבייטים שנכתבו בפועל, ולא רק דגל dirty בזיכרון.
const workspace = `<script>
(function () {
  var H = window.__qaHost;
  var stored = JSON.parse(sessionStorage.getItem('received-file-state') || '{}');
  H.storage = stored.storage || {};
  var files = stored.files || {};
  window.__receivedFiles = files;
  function ok(data) { return Promise.resolve({ success: true, data: data, error: null }); }
  function persist() { sessionStorage.setItem('received-file-state', JSON.stringify({ storage: H.storage, files: files })); }
  H.replies['storage.set'] = function (p) { H.storage[p.key] = p.value; persist(); return ok({}); };
  H.replies['fs.writeFile'] = function (p) { files[p.path] = p.content; persist(); return ok({ path: p.path }); };
  H.replies['fs.readFile'] = function (p) { return ok({ content: files[p.path] || null }); };
  H.replies['fs.deleteEntry'] = function (p) { delete files[p.path]; persist(); return ok({}); };
})();
</script>`;
const early = `<script>
(function fire() {
  if (window.__qaHost.storage.session) return;
  var latch = window.__otzariaPageOpened;
  if (!latch) { setTimeout(fire, 0); return; }
  window.__qaEarly = { liveAtDispatch: latch.live };
  window.dispatchEvent(new CustomEvent('plugin.page_opened', { detail: ${JSON.stringify({
    param: { name: 'מהסייר.docx', dataUrl: dataUrl(EARLY_TEXT) },
    openedBy: 'com.example.explorer',
  })} }));
})();
</script>`;

const report = createReport('קובץ מתוסף אחר', { strict: true });
const app = await openApp({ name: 'received-file', port: PORT, extra: workspace + early });

const tabs = () => app.js('JSON.stringify([...document.querySelectorAll(".word-doctab")].map(e => e.textContent.trim()))').then(JSON.parse);
const activeTab = () => app.js('document.querySelector(".word-doctab.active")?.textContent.trim() ?? null');

async function until(predicate, ms = 20_000) {
  for (let waited = 0; waited < ms; waited += 250) {
    if (await predicate()) return true;
    await sleep(250);
  }
  return false;
}
const settled = () => until(() => app.js('!document.querySelector(".status-load__skip") && !document.querySelector(".word-doctabs-new")?.disabled'));

try {
  // 1 — לפני הבאנדל
  const earlyProbe = await app.js('JSON.stringify(window.__qaEarly ?? null)').then(JSON.parse);
  earlyProbe && earlyProbe.liveAtDispatch === false
    ? report.pass('האירוע נורה לפני שהמאזין החי נרשם — כלומר רק ה-latch יכול היה לתפוס אותו')
    : report.fail('האירוע נורה לפני שהמאזין החי נרשם — כלומר רק ה-latch יכול היה לתפוס אותו', JSON.stringify(earlyProbe));

  const earlyOpened = await until(async () => (await app.screenText()).includes(EARLY_TEXT));
  earlyOpened
    ? report.pass('אירוע שנורה לפני הבאנדל נפתח, עם התוכן שלו')
    : report.fail('אירוע שנורה לפני הבאנדל נפתח, עם התוכן שלו', `על המסך: ${(await app.screenText()).slice(0, 80)}`);

  const afterEarly = await tabs();
  afterEarly.length === 1 && afterEarly[0].includes('מהסייר')
    ? report.pass('הוא החליף את המסמך הריק של העלייה, בשם שנשלח')
    : report.fail('הוא החליף את המסמך הריק של העלייה, בשם שנשלח', JSON.stringify(afterEarly));

  const earlyStatus = await app.status();
  /נפתח מתוסף אחר/.test(earlyStatus.text ?? '') && !earlyStatus.error
    ? report.pass('שורת המצב אומרת מאיפה הוא הגיע ומה „שמור” יעשה')
    : report.fail('שורת המצב אומרת מאיפה הוא הגיע ומה „שמור” יעשה', JSON.stringify(earlyStatus));

  (await app.exists('.word-doctab.active .word-doctab-dirty'))
    ? report.pass('הוא מסומן כלא-שמור — אין לו קובץ לכתוב אליו')
    : report.fail('הוא מסומן כלא-שמור — אין לו קובץ לכתוב אליו', 'אין סימון שינויים בטאב');

  const drafted = await until(async () => app.js('!!window.__qaHost.storage.session?.documents[0]?.draft'));
  const draft = await app.js('JSON.stringify(window.__qaHost.storage.session?.documents[0]?.draft)').then(JSON.parse);
  const exactBytes = draft && await app.js(`window.__receivedFiles[${JSON.stringify(draft.path)}] === ${JSON.stringify(dataUrl(EARLY_TEXT).split(',')[1])}`);
  drafted && draft.name === 'מהסייר.docx' && exactBytes
    ? report.pass('המקור נכתב לשחזור מיד, עם שמו, בלי עריכה או ייצוא')
    : report.fail('המקור נכתב לשחזור מיד, עם שמו, בלי עריכה או ייצוא', JSON.stringify(draft));

  await app.cdp.send('Page.reload');
  const reopened = await until(async () => ((await app.screenText()) ?? '').includes(EARLY_TEXT));
  await settled();
  reopened && (await activeTab())?.includes('מהסייר') && await app.exists('.word-doctab.active .word-doctab-dirty')
    ? report.pass('הפעלה מחדש בלי עריכה או flush משחזרת את התוכן והשם כלא-שמורים')
    : report.fail('הפעלה מחדש בלי עריכה או flush משחזרת את התוכן והשם כלא-שמורים');

  // האירוע מגיע אחרי ש-onExit כבר לכדה את רשימת המסמכים. הוא חייב להמתין
  // עד אחרי סיום השאלה והסגירה, ואז להיפתח ולא להימחק יחד עם המסמך הישן.
  await app.tab('קובץ');
  await app.click('יציאה');
  await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({
    param: { name: 'ממתין.docx', dataUrl: dataUrl('הגיע בזמן שאלת יציאה') },
  })})`);
  await sleep(400);
  (await tabs()).length === 1 && await app.exists('.unsaved-dialog')
    ? report.pass('אירוע בזמן שאלת יציאה ממתין ואינו מוסיף טאב לרשימת הסגירה')
    : report.fail('אירוע בזמן שאלת יציאה ממתין ואינו מוסיף טאב לרשימת הסגירה');
  await app.clickSel('.unsaved-dialog [data-choice="discard"]');
  const afterExit = await until(async () => (await app.screenText()).includes('הגיע בזמן שאלת יציאה'));
  afterExit && (await tabs()).length === 1 && (await activeTab())?.includes('ממתין')
    ? report.pass('אחרי סיום היציאה הקובץ הממתין נפתח ונשאר חי')
    : report.fail('אחרי סיום היציאה הקובץ הממתין נפתח ונשאר חי');

  // 2 — העורך חי
  await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({
    param: { name: 'שני.docx', dataUrl: dataUrl(LIVE_TEXT) },
    openedBy: 'com.example.explorer',
  })})`);
  const liveOpened = await until(async () => (await app.screenText()).includes(LIVE_TEXT));
  await settled();
  await until(async () => /שני.docx נפתח מתוסף אחר/.test((await app.status()).text ?? ''));
  const afterLive = await tabs();
  liveOpened && afterLive.length === 2 && (await activeTab())?.includes('שני')
    ? report.pass('אירוע שמגיע לעורך חי נפתח בטאב חדש ופעיל')
    : report.fail('אירוע שמגיע לעורך חי נפתח בטאב חדש ופעיל', `נפתח=${liveOpened} טאבים=${JSON.stringify(afterLive)}`);

  await app.clickSel('.word-doctab.active .word-doctab-close');
  await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({
    param: { name: 'אחרי ביטול.docx', dataUrl: dataUrl('הגיע בזמן סגירת טאב') },
  })})`);
  await sleep(400);
  (await tabs()).length === 2 && await app.exists('.unsaved-dialog')
    ? report.pass('אירוע בזמן סגירת טאב אינו מחליף את המסמך שעליו נשאלים')
    : report.fail('אירוע בזמן סגירת טאב אינו מחליף את המסמך שעליו נשאלים');
  await app.clickSel('.unsaved-dialog [data-choice="cancel"]');
  const afterCancel = await until(async () => (await app.screenText()).includes('הגיע בזמן סגירת טאב'));
  await settled();
  await until(async () => /אחרי ביטול.docx נפתח מתוסף אחר/.test((await app.status()).text ?? ''));
  afterCancel && (await tabs()).length === 3
    ? report.pass('ביטול סגירת טאב משאיר את המסמכים ומאפשר לפתוח את האירוע הממתין')
    : report.fail('ביטול סגירת טאב משאיר את המסמכים ומאפשר לפתוח את האירוע הממתין');

  await app.js(`(() => {
    window.__qaHost.replies['fs.beginBinaryWrite'] = { writeToken: 'qa-save', uploadUrl: 'https://qa.local/upload' };
    window.__qaHost.replies['fs.abortBinaryWrite'] = {};
    const fetch = window.fetch;
    window.fetch = (input, options) => input === 'https://qa.local/upload' ? Promise.resolve(new Response('', { status: 200 })) : fetch(input, options);
    window.__qaHost.replies['fs.commitUserFileWrite'] = () => new Promise(resolve => {
      window.__finishSave = () => resolve({ success: true, data: { cancelled: true }, error: null });
    });
    const session = [...window.__otzariaEditors.values()].pop();
    void session.save.saveNow({ forceSaveAs: true });
  })()`);
  await until(async () => app.js('!!window.__finishSave'));
  await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({
    param: { name: 'אחרי שמירה.docx', dataUrl: dataUrl('הגיע בזמן שמור בשם') },
  })})`);
  await sleep(400);
  (await tabs()).length === 3
    ? report.pass('אירוע בזמן „שמור בשם” ממתין לסיום השמירה')
    : report.fail('אירוע בזמן „שמור בשם” ממתין לסיום השמירה');
  await app.js('window.__finishSave()');
  const afterSave = await until(async () => (await app.screenText()).includes('הגיע בזמן שמור בשם'));
  afterSave && (await tabs()).length === 4
    ? report.pass('ביטול „שמור בשם” משחרר את האירוע הממתין')
    : report.fail('ביטול „שמור בשם” משחרר את האירוע הממתין');

  // 3 — בלי תוכן
  await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({
    param: { name: 'נתיב.docx', path: 'C:\\docs\\נתיב.docx' },
    openedBy: 'com.example.explorer',
  })})`);
  const explained = await until(async () => /נתיב\.docx.*לא צירף את תוכן הקובץ/.test((await app.status()).text ?? ''), 5_000);
  const rejectedStatus = await app.status();
  explained && rejectedStatus.error && (await tabs()).length === 4
    ? report.pass('נתיב בלי תוכן — הודעת שגיאה שמסבירה, ובלי טאב')
    : report.fail('נתיב בלי תוכן — הודעת שגיאה שמסבירה, ובלי טאב', `${JSON.stringify(rejectedStatus)} טאבים=${(await tabs()).length}`);

  // 4 — בלי קובץ
  await app.js(`window.__qaHost.emit('plugin.page_opened', { source: 'newTabButton' })`);
  await sleep(1_500);
  const quietStatus = await app.status();
  (await tabs()).length === 4 && quietStatus.text === rejectedStatus.text
    ? report.pass('אירוע בלי קובץ אינו פותח דבר ואינו מציג הודעה')
    : report.fail('אירוע בלי קובץ אינו פותח דבר ואינו מציג הודעה', JSON.stringify(quietStatus));

  await app.js(`window.__qaHost.replies['fs.writeFile'] = () => Promise.resolve({ success: false, data: null, error: { message: 'מכסה מלאה' } })`);
  await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({
    param: { name: 'בלי גיבוי.docx', dataUrl: dataUrl('המקור פתוח למרות כשל כתיבה') },
  })})`);
  const writeFailure = await until(async () => /לא נוצר עותק לשחזור/.test((await app.status()).text ?? ''));
  writeFailure && (await app.status()).error && (await app.screenText()).includes('המקור פתוח למרות כשל כתיבה')
    ? report.pass('כשל בכתיבת המקור מוצג כשגיאה שמבקשת לשמור, והמסמך נשאר פתוח')
    : report.fail('כשל בכתיבת המקור מוצג כשגיאה שמבקשת לשמור, והמסמך נשאר פתוח');

  // קריאת Blob שעוד לא חזרה אחרי „דלג” אינה יכולה לנעול פתיחה חדשה.
  await app.js(`(() => {
    const read = Blob.prototype.arrayBuffer;
    Blob.prototype.arrayBuffer = function () {
      if (!window.__blockedRead) {
        window.__blockedRead = true;
        return new Promise(resolve => { window.__releaseRead = () => read.call(this).then(resolve); });
      }
      return read.call(this);
    };
  })()`);
  await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({
    param: { name: 'קריאה ממתינה.docx', dataUrl: dataUrl('לא אמור להיפתח אחרי דלג') },
  })})`);
  await until(async () => app.js('!!window.__releaseRead'));
  await app.clickSel('.status-load__skip');
  const beforeNew = (await tabs()).length;
  await app.clickSel('.word-doctabs-new');
  const skipReleased = await until(async () => (await tabs()).length === beforeNew + 1 && await app.js('window.__qa.ready()'));
  skipReleased
    ? report.pass('„דלג” משחרר פתיחה חדשה מיד, גם כשקריאת הבייטים הישנה עדיין ממתינה')
    : report.fail('„דלג” משחרר פתיחה חדשה מיד, גם כשקריאת הבייטים הישנה עדיין ממתינה');
  await app.js('window.__releaseRead()');
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);

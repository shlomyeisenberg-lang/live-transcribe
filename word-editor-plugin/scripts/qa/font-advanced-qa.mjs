/**
 * שער QA ל„גופן מתקדם” — ובמפורש למה שאף שער אחר לא מדד: **מה מצויר על המסך.**
 *
 * ## למה שער נוסף, כששלב 5 של `home-font-qa` כבר בודק את הדיאלוג
 *
 * הוא בודק את ה-OOXML. וזה מספיק לרוב הפקדים, אבל לא כאן: „לחצתי אישור ולא
 * קרה כלום” הוא **דיווח על המסך**, ומסמך שנכתב נכון אינו מפריך אותו. נמדד:
 * `<w:outline/>`, `<w:shadow/>`, `<w:emboss/>` ו-`<w:imprint/>` יוצאים ל-docx
 * כמו שצריך, ו-`.superdoc-text-run` על המסך אינו נושא שום סימן לאף אחד מהם.
 * שער שמסתכל בקובץ בלבד היה עובר על התקלה הזאת בירוק.
 *
 * ## חמש השורות שהשער מחזיק, וההנמקה של השלוש האחרונות
 *
 * 1. **בלי בחירה הדיאלוג אומר את זה בפתיחה** — ולא אחרי שהמשתמש מילא שבעה-עשר
 *    פקדים ולחץ. „אישור” נעול, והמונה אומר למה.
 * 2. **מה שהוחל נכתב ומצויר** — ריווח, מתיחה ומיקום, בקובץ ועל המסך.
 * 3. **הרשימה שהדיאלוג מכריז עליה כ„לא מצוירת” היא בדיוק מה שאינו מצויר** —
 *    אפקט אחד לכל ריצה, כדי שאפשר יהיה לייחס פיקסל לאפקט.
 * 4. **„קו חוצה כפול” מצויר, וכקו בודד** — הטענה השנייה של אותה הודעה.
 * 5. **גם „קרנינג” ו„גודל הגופן המורכב” ברשימה** — ההודעה אינה מונה אפקטים
 *    בלבד, ושער שמאמת ארבעה מתוך שישה משאיר שני חלקים ממנה בלי שומר.
 *
 * שלוש האחרונות מחזיקות את היושר של ההודעה למשתמש לאורך זמן: ברגע שהמנוע
 * יתחיל לצייר צל, או קו כפול (התיקון נשלח למעלה — ראו docs/engine-gaps.md),
 * השורה תיפול ותאמר מה להוריד מההודעה. הודעה שאיש אינו מודד היא הודעה
 * שתשקר בשקט ברגע שהעולם מתחתיה ישתנה.
 *
 * יציאה 9366 בלבד — שערים אחרים רצים במקביל על יציאות אחרות.
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('גופן מתקדם — מה נכתב ומה מצויר', { strict: true });
const app = await openApp({ name: 'font-adv', port: Number(process.env.QA_PORT ?? 9366) });

const notes = [];
function note(...p) {
  const l = p.map((x) => (typeof x === 'string' ? x : JSON.stringify(x))).join(' ');
  notes.push(l);
  console.log(l);
}

/** כל קריאה ל-CDP חוסמת ללא סוף כשהדף תקוע; בלי זה השער נתקע בלי לומר איפה. */
const T = (p, label, ms = 40_000) =>
  Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error(`תקיעה ב-${label}`)), ms))]);

const body = (d) => d.slice(d.indexOf('<w:body'));

/** ה-rPr של הריצה שמכילה בדיוק את הטקסט הזה. '' כשאין rPr, null כשאין ריצה. */
function rPrOf(doc, text) {
  const runs = body(doc).match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g) || [];
  const hit = runs.find((r) => new RegExp(`<w:t[^>]*>${text}</w:t>`).test(r));
  if (!hit) return null;
  const m = hit.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

const docx = () => T(app.docx(), 'ייצוא docx', 90_000);

/**
 * הסגנון המחושב של הריצה שעל המסך.
 *
 * `.superdoc-text-run` הוא מה שהצייר של המנוע מייצר, והוא **אינו** נושא את
 * מאפייני הריצה כתכונות — ולכן אין דרך לשאול „האם יש כאן `w:outline`”. מה
 * שאפשר למדוד הוא הפיקסלים: מסגרת היא `-webkit-text-stroke`, צל הוא
 * `text-shadow`, וקו כפול הוא `text-decoration-style: double`. זו הסיבה
 * שהשער מודד CSS מחושב ולא DOM.
 */
const runStyle = (needle) =>
  app.js(`(function(){
    var runs = document.querySelectorAll('.superdoc-text-run'), hit = null;
    for (var i = 0; i < runs.length; i++) { if ((runs[i].textContent || '') === ${JSON.stringify(needle)}) { hit = runs[i]; break; } }
    if (!hit) for (var j = 0; j < runs.length; j++) { if ((runs[j].textContent || '').indexOf(${JSON.stringify(needle)}) >= 0) { hit = runs[j]; break; } }
    if (!hit) return JSON.stringify({ found: false });
    var cs = getComputedStyle(hit);
    return JSON.stringify({
      found: true,
      fontFamily: cs.fontFamily,
      fontSize: cs.fontSize,
      fontKerning: cs.fontKerning,
      letterSpacing: cs.letterSpacing,
      transform: cs.transform,
      decoLine: cs.textDecorationLine,
      decoStyle: cs.textDecorationStyle,
      stroke: cs.webkitTextStrokeWidth,
      color: cs.color,
      fill: cs.webkitTextFillColor,
      shadow: cs.textShadow
    });
  })()`).then(JSON.parse);

/** מה שהדיאלוג מציג כרגע: ההודעה החוסמת, הערת „לא מצויר”, המונה, ומצב „אישור”. */
const dialogFace = () =>
  app.js(`(function(){
    var d = document.querySelector('[role="dialog"]');
    if (!d) return JSON.stringify({ open: false });
    var blocking = d.querySelector('.fa-notice-blocking');
    var notice = d.querySelector('.fa-notice:not(.fa-notice-blocking)');
    var count = d.querySelector('.fa-count');
    var ok = d.querySelector('.fa-btn-primary');
    return JSON.stringify({
      open: true,
      blocking: blocking ? blocking.textContent.trim() : null,
      notice: notice ? notice.textContent.trim() : null,
      count: count ? count.textContent.trim() : null,
      okDisabled: ok ? !!ok.disabled : null
    });
  })()`).then(JSON.parse);

/** לוחצת TriToggle לפי התווית עד שהוא במצב המבוקש. */
async function tri(label, want = 'yes') {
  for (let i = 0; i < 3; i++) {
    const state = await app.js(`(function(){
      var d = document.querySelector('[role="dialog"]'); if (!d) return 'no-dialog';
      var all = d.querySelectorAll('.tri-toggle');
      for (var i = 0; i < all.length; i++) if ((all[i].getAttribute('data-tip-title') || '') === ${JSON.stringify(label)}) return all[i].getAttribute('data-state');
      return 'not-found';
    })()`);
    if (state === want) return true;
    if (state === 'not-found' || state === 'no-dialog') return false;
    if (!(await app.clickDialog(label, { after: 200 }))) return false;
  }
  return false;
}

async function selectRange(from, to) {
  await app.press('Home', 'Home', 36);
  await app.sleep(140);
  for (let i = 0; i < from; i++) { await app.press('ArrowRight', 'ArrowRight', 39); await app.sleep(22); }
  for (let i = from; i < to; i++) { await app.press('ArrowRight', 'ArrowRight', 39, 8); await app.sleep(22); }
  await app.sleep(500);
}

/**
 * מה שנבחר **לפי המנוע** — בלוק והיסטים. `ui.selection.get().text` אינו
 * מוחזר ב-2.15.0-next.15 (נמדד: חסר), ומה שנראה מסומן על המסך אינו ראיה.
 */
const engineSelection = () =>
  app.js(`(async function(){
    var info = await window.__otzariaEditor.superdoc.activeEditor.doc.selection.current();
    var t = (info && info.selectionTarget) || {};
    return JSON.stringify({
      empty: !!(info && info.empty),
      startBlock: t.start ? t.start.blockId : null, start: t.start ? t.start.offset : null,
      endBlock: t.end ? t.end.blockId : null, end: t.end ? t.end.offset : null
    });
  })()`).then(JSON.parse);

/**
 * בוחרת מילה שעומדת לבדה בפסקה. מחזירה null כשהמנוע מאשר שנבחרה המילה כולה
 * ורק היא, ואחרת את מה שכן נבחר.
 *
 * ## למה האימות הוא חלק מהבחירה
 *
 * שלב 5 שואל „נכתב ל-docx?” על הריצה שהטקסט שלה הוא בדיוק המילה. בחירה שנחתה
 * שלוש אותיות ליד נענתה שם „לא נכתב” — כלומר נקראה ככשל של המנוע — כשבפועל
 * העיצוב נכתב, רק על אותיות אחרות. `caretPara` מאמת את הפסקה, וההיסטים
 * 0..אורך המילה באותו בלוק מאמתים את הטווח; מה שאינו כזה נופל כאן, בשמו.
 */
async function selectAlone(word) {
  let at;
  try {
    at = await T(app.caretPara(word), `סמן בפסקה „${word}”`, 90_000);
  } catch (error) {
    return `הסמן לא הגיע לפסקה „${word}”: ${error && error.message}`;
  }
  /*
   * ההרחבה היא לכיוון **החזותי** של המילה, ולא `ArrowRight` תמיד: המנוע מזיז
   * את הסמן בחצים לפי מה שנראה על המסך, ולכן בפסקה שכולה עברית `ArrowRight`
   * הולך אחורה לוגית. נמדד כאן: הבחירה על „אבגד” יצאה מהפסקה אל זו שלפניה
   * (בלוק אחר, היסט 3). מילה עברית נבחרת בחץ שמאלה, לועזית בחץ ימינה.
   */
  const [key, code] = /[\u0590-\u05FF]/.test(word) ? ['ArrowLeft', 37] : ['ArrowRight', 39];
  await app.press('Home', 'Home', 36);
  await app.sleep(140);
  for (let i = 0; i < word.length; i++) {
    await app.press(key, key, code, 8);
    await app.sleep(22);
  }
  await app.sleep(500);
  const sel = await T(engineSelection(), 'selection.current');
  const exact =
    !sel.empty && sel.startBlock === at.nodeId && sel.endBlock === at.nodeId && sel.start === 0 && sel.end === word.length;
  return exact ? null : `הבחירה אינה „${word}” (בלוק ${at.nodeId}, 0..${word.length}): ${JSON.stringify(sel)}`;
}

/**
 * האפקטים שהדיאלוג מכריז עליהם כ„נכתבים ולא מצוירים”, והחתימה שתעיד שהם
 * **כן** צוירו.
 *
 * ## פקד אחד לכל ריצה, וזה לא ניואנס
 *
 * גרסה ראשונה החילה את כל הארבעה על אותה ריצה. `shadow`, `emboss` ו-`imprint`
 * מצוירים כולם ב-`text-shadow`, ולכן ברגע שהמנוע היה מתחיל לצייר **אחד** מהם
 * שלושת הבדיקות היו מדווחות „כן מצויר” — והשער, שכל תפקידו לשמור על יושר
 * ההודעה, היה מורה להוריד ממנה שניים שהמנוע עדיין בולע. כלומר בדיוק השקר
 * שהוא נועד למנוע, בחתימתו שלו. מילת עוגן משלה לכל אפקט מפרידה ביניהם.
 *
 * `probe` מקבל את הסגנון המחושב ומחזיר את מה שאפשר לייחס לאפקט הזה בלבד.
 */
const EFFECT_PROBES = [
  {
    label: 'מסגרת לתו',
    tag: 'outline',
    word: 'outl',
    /*
     * שתי הדרכים שבהן אפשר לצייר מסגרת: קו מתאר, או מילוי שרוקן כדי שהקו
     * ייראה. הצורה המחושבת של „ריק” היא `rgba(r, g, b, 0)` ולעולם לא
     * המילה `transparent` — השוואה למילה היא ענף מת שלעולם אינו נבדק.
     */
    drawn: (s) => Number.parseFloat(s.stroke) > 0 || /rgba\([^)]*,\s*0\)\s*$/.test(s.fill ?? ''),
  },
  { label: 'צל', tag: 'shadow', word: 'shdw', drawn: (s) => s.shadow !== 'none' && s.shadow !== '' },
  { label: 'חרוט', tag: 'emboss', word: 'embs', drawn: (s) => s.shadow !== 'none' && s.shadow !== '' },
  { label: 'שקוע', tag: 'imprint', word: 'impr', drawn: (s) => s.shadow !== 'none' && s.shadow !== '' },
];

// `fontSizeCs` חייב להימדד על כתב מורכב אמיתי. באנגלית מותר למנוע להתעלם
// מ-`w:szCs`; עברית מבדילה בין „לא צויר” בפועל לבין מסלול שלא הופעל כלל.
//
// שלוש השורות האחרונות הן מילה אחת לפסקה, וזה תיקון שנמדד ולא סגנון. שלוש
// המילים ישבו פעם בשורה אחת, `kern dstx אבגד`, והבחירה של שלב 5 הייתה Home
// ואחריו חצים. שלב 4 מפצל את השורה לשלוש ריצות, ומאותו רגע Home בפסקה
// המעורבת נחת בהיסט 9 — תחילת ` אבגד` — ולא ב-0 (נמדד על 2.15.0-next.15):
// הקרנינג נכתב על ` אבג`, העברית נשארה בלי בחירה, והשלב היה אדום גם על 2.14
// בלי שהמנוע שינה דבר. בפסקה של מילה אחת אין גבול בין ריצות שהסמן נופל עליו.
const LINES = ['numx spac', 'outl shdw embs impr', 'dstx', 'kern', 'אבגד'];

try {
  /* חלון רחב — ברירת המחדל ב-headless צרה, והרצועה גולשת ממנה. */
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false });
  await app.sleep(800);
  await T(app.tab('בית'), 'מעבר ללשונית בית');

  await app.caret(0);
  for (let i = 0; i < LINES.length; i++) {
    await app.type(LINES[i], 28);
    if (i < LINES.length - 1) { await app.press('Enter', 'Enter', 13); await app.sleep(260); }
  }
  await app.sleep(1500);
  note('פיקסטורה:', await T(app.screenText(), 'screenText'));

  /* ============ 1. בלי בחירה — הדיאלוג אומר את זה בפתיחה ============ */
  await app.reset();
  await app.caret(0);
  await app.sleep(400);
  if (!(await app.click('מתקדם', { after: 1200 }))) {
    report.fail('„מתקדם” נפתח', 'הכפתור לא נמצא או אינו מוצג');
  } else {
    /*
     * מצב ההתחלה הוא „ללא שינוי” בכל שדה, ולכן „אישור” נעול ממילא. הבדיקה
     * היא שהוא נשאר נעול **אחרי** שנבחר משהו, ושנאמר למה — ולכן הבחירה
     * חייבת להצליח. בלי האימות הזה, תווית שהשתנתה הייתה משאירה את המונה על
     * אפס, את הכפתור נעול מסיבה אחרת לגמרי, ואת השורה ירוקה על לא כלום.
     */
    if (!(await tri('קו חוצה כפול', 'yes'))) {
      report.fail('בלי בחירה — נאמר בפתיחה', 'כפתור האפקט „קו חוצה כפול” לא נמצא בדיאלוג');
    }
    await app.sleep(300);
    const face = await T(dialogFace(), 'dialogFace');
    note('בלי בחירה:', JSON.stringify(face));
    const said = typeof face.blocking === 'string' && face.blocking.includes('מסומן');
    if (said && face.okDisabled === true) {
      report.pass('בלי בחירה — נאמר בפתיחה', `„${face.count}”, „אישור” נעול`);
    } else {
      report.fail('בלי בחירה — נאמר בפתיחה', `blocking=${JSON.stringify(face.blocking)} okDisabled=${face.okDisabled}`);
    }
    await app.escape();
  }

  /* ============ 2. מספרים — נכתבים ומצוירים ============ */
  await app.reset();
  await app.caret(0);
  await selectRange(0, 4);
  if (!(await app.click('מתקדם', { after: 1200 }))) {
    report.fail('מספרים — נכתבים ומצוירים', 'הדיאלוג לא נפתח');
  } else {
    await T(app.dialogFill('fa-spacing', '2'), 'fill spacing');
    await T(app.dialogFill('fa-scale', '125'), 'fill scale');
    await T(app.dialogFill('fa-position', '3'), 'fill position');
    await app.sleep(400);
    const before = await T(dialogFace(), 'dialogFace');
    note('לפני אישור:', JSON.stringify(before));
    await T(app.clickDialog('אישור', { after: 2000 }), 'לחיצה אישור');
    const rpr = rPrOf((await docx())['word/document.xml'], 'numx');
    const seen = await T(runStyle('numx'), 'runStyle');
    note('rPr(numx) =', JSON.stringify(rpr), '| מסך =', JSON.stringify(seen));
    const written = {
      spacing: !!rpr && /<w:spacing w:val="40"/.test(rpr),
      scale: !!rpr && /<w:w w:val="125"/.test(rpr),
      position: !!rpr && /<w:position w:val="6"/.test(rpr),
    };
    const drawn = {
      spacing: seen.found && Number.parseFloat(seen.letterSpacing) > 0,
      scale: seen.found && seen.transform !== 'none',
    };
    if (Object.values(written).every(Boolean) && Object.values(drawn).every(Boolean)) {
      report.pass('מספרים — נכתבים ומצוירים', `ריווח ${seen.letterSpacing}, מתיחה ${seen.transform}`);
    } else {
      report.fail('מספרים — נכתבים ומצוירים', `נכתב=${JSON.stringify(written)} צויר=${JSON.stringify(drawn)}`);
    }
  }

  /* ============ 3. ההודעה על „לא מצויר” מתארת את מה שנמדד ============ */
  const mismatch = [];
  const declaredFor = {};
  for (const [index, probe] of EFFECT_PROBES.entries()) {
    // אפקט אחד לכל ריצה, ולכל אפקט מילת עוגן משלו — ראו ההנמקה ליד
    // `EFFECT_PROBES`. ארבע פתיחות במקום אחת, וזה המחיר של מדידה שאפשר לייחס.
    await app.reset();
    await app.caret(2);
    await selectRange(index * 5, index * 5 + 4);
    if (!(await app.click('מתקדם', { after: 1200 }))) {
      mismatch.push(`${probe.label}: הדיאלוג לא נפתח`);
      continue;
    }
    if (!(await tri(probe.label, 'yes'))) {
      mismatch.push(`${probe.label}: כפתור האפקט לא נמצא בדיאלוג`);
      await app.escape();
      continue;
    }
    await app.sleep(400);
    const face = await T(dialogFace(), 'dialogFace');
    declaredFor[probe.label] = face.notice ?? '';
    await T(app.clickDialog('אישור', { after: 2200 }), 'לחיצה אישור');

    const rpr = rPrOf((await docx())['word/document.xml'], probe.word);
    const seen = await T(runStyle(probe.word), 'runStyle');
    note(`${probe.label} (${probe.word}): rPr=${JSON.stringify(rpr)} | מסך=${JSON.stringify(seen)} | הודעה=${JSON.stringify(face.notice)}`);

    const inFile = !!rpr && new RegExp(`<w:${probe.tag}\\s*/?>`).test(rpr);
    /*
     * הריצה חייבת להימצא על המסך לפני שמודדים עליה משהו. בלי השורה הזאת
     * `seen.found === false` היה נקרא כ„אינו מצויר”, וכל ארבע השורות היו
     * ירוקות אחרי שינוי מבנה שהעיף את הסלקטור — כלומר השער היה מאשר את
     * ההודעה בדיוק כשאין לו במה למדוד אותה.
     */
    if (!seen.found) {
      mismatch.push(`${probe.label}: הריצה לא נמצאה על המסך — אין מה למדוד`);
      continue;
    }
    const isDrawn = probe.drawn(seen);
    const named = (face.notice ?? '').includes(probe.label);
    // שלוש שאלות נפרדות: נכתב לקובץ, מצויר על המסך, והוכרז בדיאלוג.
    if (!inFile) mismatch.push(`${probe.label}: לא נכתב ל-docx`);
    else if (isDrawn && named) mismatch.push(`${probe.label}: **כן** מצויר — יש להסיר אותו מההודעה`);
    else if (!isDrawn && !named) mismatch.push(`${probe.label}: אינו מצויר ואינו מוכרז בהודעה`);
  }
  if (mismatch.length === 0) {
    report.pass('הערת „לא מצויר” — מדויקת', `${EFFECT_PROBES.length} אפקטים, כל אחד בריצה משלו`);
  } else {
    report.fail('הערת „לא מצויר” — מדויקת', mismatch.join(' | '));
  }

  /* ============ 4. „קו חוצה כפול” — הטענה שההודעה עצמה מוסרת ============ */
  await app.reset();
  await app.caret(4);
  await selectRange(0, 4); // 'dstx' — לבדה בשורה השלישית
  if (!(await app.click('מתקדם', { after: 1200 }))) {
    report.fail('„קו חוצה כפול” — מצויר, וכקו בודד', 'הדיאלוג לא נפתח');
  } else if (!(await tri('קו חוצה כפול', 'yes'))) {
    report.fail('„קו חוצה כפול” — מצויר, וכקו בודד', 'כפתור האפקט לא נמצא בדיאלוג');
  } else {
    await app.sleep(400);
    const face = await T(dialogFace(), 'dialogFace');
    await T(app.clickDialog('אישור', { after: 2200 }), 'לחיצה אישור');
    const rpr = rPrOf((await docx())['word/document.xml'], 'dstx');
    const seen = await T(runStyle('dstx'), 'runStyle');
    note('dstx:', JSON.stringify(rpr), JSON.stringify(seen), JSON.stringify(face.notice));

    const written = !!rpr && /<w:dstrike\s*\/?>/.test(rpr);
    const drawn = seen.found && seen.decoLine.includes('line-through');
    const doubled = seen.found && seen.decoStyle === 'double';
    const said = (face.notice ?? '').includes('בודד');
    /*
     * ההודעה טוענת „מצויר, אך כקו בודד”. שלוש השאלות הן בדיוק שלוש המילים
     * שלה, ואם המנוע יתחיל לצייר כפול — התיקון נשלח למעלה — השורה תיפול
     * ותאמר להוריד את המשפט.
     */
    if (written && drawn && !doubled && said) {
      report.pass('„קו חוצה כפול” — מצויר, וכקו בודד', `deco=${seen.decoLine}/${seen.decoStyle}, וההודעה אומרת זאת`);
    } else if (written && doubled) {
      report.fail('„קו חוצה כפול” — מצויר, וכקו בודד', 'המנוע מצייר כפול — יש להסיר את המשפט מההודעה');
    } else {
      report.fail(
        '„קו חוצה כפול” — מצויר, וכקו בודד',
        `נכתב=${written} מצויר=${drawn} כפול=${doubled} הוכרז=${said} deco=${seen.decoLine}/${seen.decoStyle}`,
      );
    }
  }
  /* ============ 5. קרנינג וגודל הכתב המורכב — גם הם ברשימה ============ */
  /*
   * ההודעה מונה יותר מארבעה אפקטים — היא מונה גם את „קרנינג” ואת „גודל הגופן
   * המורכב”, שנמדדו באותה מדידה בדיוק כנכתבים-ואינם-מצוירים. שער שמאמת ארבעה
   * מתוך שישה משאיר שני חלקים מההודעה בלי מי שיתפוס אותם כשישתנו.
   *
   * שני אלה אינם כפתורי מיתוג אלא שדות מספר, ולכן הם כאן ולא ב-`EFFECT_PROBES`.
   */
  const FIELD_PROBES = [
    { label: 'קרנינג', id: 'fa-kerning', value: '12', xml: /<w:kern w:val="24"/, word: 'kern' },
    { label: 'גודל הגופן המורכב', id: 'fa-sizecs', value: '28', xml: /<w:szCs w:val="56"/, word: 'אבגד' },
  ];
  const fieldMismatch = [];
  for (const probe of FIELD_PROBES) {
    await app.reset();
    // הבחירה מאומתת מול המנוע לפני שהדיאלוג נפתח — ראו `selectAlone`. בחירה
    // שנחתה ליד המילה נענתה כאן „לא נכתב ל-docx”, כלומר כשל שנתלה במנוע.
    const missed = await selectAlone(probe.word);
    if (missed) {
      fieldMismatch.push(`${probe.label}: ${missed}`);
      continue;
    }
    if (!(await app.click('מתקדם', { after: 1200 }))) {
      fieldMismatch.push(`${probe.label}: הדיאלוג לא נפתח`);
      continue;
    }
    if ((await T(app.dialogFill(probe.id, probe.value), `fill ${probe.id}`)) !== 'ok') {
      fieldMismatch.push(`${probe.label}: השדה ${probe.id} לא נמצא בדיאלוג`);
      await app.escape();
      continue;
    }
    await app.sleep(400);
    const face = await T(dialogFace(), 'dialogFace');
    const before = await T(runStyle(probe.word), 'runStyle לפני');
    await T(app.clickDialog('אישור', { after: 2200 }), 'לחיצה אישור');

    const rpr = rPrOf((await docx())['word/document.xml'], probe.word);
    const after = await T(runStyle(probe.word), 'runStyle אחרי');
    note(`${probe.label} (${probe.word}): rPr=${JSON.stringify(rpr)} | לפני=${JSON.stringify(before)} | אחרי=${JSON.stringify(after)} | הודעה=${JSON.stringify(face.notice)}`);

    const inFile = !!rpr && probe.xml.test(rpr);
    // שתי התמונות, ולא רק זו שאחרי: „לפני” שלא נמצאה הייתה נקראת כ„אינו
    // מצויר” — כלומר אישור שקט להודעה, בדיוק כשאין במה למדוד אותה.
    if (!before.found || !after.found) {
      fieldMismatch.push(
        `${probe.label}: הריצה לא נמצאה על המסך — אין מה למדוד (לפני=${before.found}, אחרי=${after.found})`,
      );
      continue;
    }
    // „מצויר” כאן הוא **כל** שינוי בסגנון המחושב, ולא תכונה אחת: אין מראש
    // מועמד יחיד לקרנינג או לגודל CS, וכל שינוי היה מפריך את ההודעה.
    const isDrawn = JSON.stringify(before) !== JSON.stringify(after);
    const named = (face.notice ?? '').includes(probe.label);
    if (!inFile) fieldMismatch.push(`${probe.label}: לא נכתב ל-docx`);
    else if (isDrawn && named) fieldMismatch.push(`${probe.label}: **כן** מצויר — יש להסיר אותו מההודעה`);
    else if (!isDrawn && !named) fieldMismatch.push(`${probe.label}: אינו מצויר ואינו מוכרז בהודעה`);
  }
  if (fieldMismatch.length === 0) {
    report.pass('הערת „לא מצויר” — גם על מה שאינו אפקט', `${FIELD_PROBES.length} פקדים, כל אחד בריצה משלו`);
  } else {
    report.fail('הערת „לא מצויר” — גם על מה שאינו אפקט', fieldMismatch.join(' | '));
  }
} catch (error) {
  note('!! השער נפל:', String(error && error.stack ? error.stack : error));
  report.stuck('השער', String(error && error.message ? error.message : error));
} finally {
  app.close();
}

report.print();

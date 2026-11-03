/**
 * שער הגופן: האם ה-@font-face שהתוסף מזריק בזמן ריצה נטען בפועל מ-`file://` —
 * ה-origin שממנו אוצריא טוענת תוסף ארוז.
 *
 * רץ על `dist/index.html` עצמו, בלי לשנות בו כלום: `installBundledFonts()` הוא
 * הדבר הראשון שהמעטפת עושה, לפני ההמתנה ל-Host שלא יגיע כאן.
 *
 * מה שנמדד, ולמה דווקא זה: `document.fonts.check()` **אינו עדות** — נמדד שהוא
 * מחזיר `true` גם בדף בלי שום `@font-face`, על מכונה שאין בה את הגופן. מה
 * שמפריד הוא רוחב טקסט על canvas מול serif — בלטינית וגם בעברית, מפני
 * ש-Assistant מכסה את שניהם ועברית היא כמעט כל מה שייכתב בתוסף הזה.
 *
 *   npm run build && npm run check:fonts
 */
import { existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, requireChrome, sleep } from './cdp.mjs';

const DIST = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist');
const INDEX = join(DIST, 'index.html');

if (!existsSync(INDEX)) {
  console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}
requireChrome();

/** ארבעת המשקלים שנארזים. אין נטוי: האריזה כוללת רק פנים זקופות. */
const SPECS = [
  '400 16px "Assistant"',
  '500 16px "Assistant"',
  '600 16px "Assistant"',
  '700 16px "Assistant"',
];

const PROBE = `(async function () {
  for (var i = 0; i < 200 && !document.getElementById('bundled-fonts'); i++) {
    await new Promise(function (r) { setTimeout(r, 25); });
  }
  if (!document.getElementById('bundled-fonts')) return { injected: false };

  var specs = ${JSON.stringify(SPECS)};
  var results = [];
  for (var s = 0; s < specs.length; s++) {
    try {
      // לטינית ועברית יחד: הגופן מכסה את שניהם, ובקשה למתו שאינו בגופן
      // לא הייתה טוענת אותו.
      var faces = await document.fonts.load(specs[s], 'ABC \u05d0\u05d1\u05d2');
      results.push({ spec: specs[s], loaded: faces.length });
    } catch (e) {
      results.push({ spec: specs[s], error: String(e) });
    }
  }

  function width(family, text) {
    var c = document.createElement('canvas').getContext('2d');
    c.font = '40px ' + family;
    return c.measureText(text).width;
  }
  var LATIN = 'Handgloves Wgm';
  var HEBREW = '\u05d0\u05d5\u05e6\u05e8\u05d9\u05d0 \u05e2\u05d5\u05e8\u05da';

  return {
    injected: true,
    origin: location.protocol,
    results: results,
    ours: (function () {
      var n = 0;
      document.fonts.forEach(function (face) { if (face.family === 'Assistant') n++; });
      return n;
    })(),
    widthLatin: width('"Assistant", serif', LATIN),
    widthLatinSerif: width('serif', LATIN),
    widthHebrew: width('"Assistant", serif', HEBREW),
    widthHebrewSerif: width('serif', HEBREW)
  };
})()`;

const page = await openPage(`file://${INDEX}`, { label: 'fonts' });
let report;
try {
  // המעטפת מזריקה את ההצהרות בשורה הראשונה של main(); הדף עדיין עולה.
  await sleep(1500);
  report = await page.cdp.evaluate(PROBE);
} finally {
  page.close();
}

const errors = [];
if (!report?.injected) {
  errors.push('ההצהרות לא הוזרקו — installBundledFonts() לא רץ');
} else {
  for (const result of report.results) {
    if (result.loaded !== 1) {
      errors.push(`הפנים ${result.spec} לא נטענה (${result.error ?? `loaded=${result.loaded}`})`);
    }
  }
  // נספרות רק הפנים שלנו: SuperDoc רושם פנים משלו
  // (`__superdoc_core_symbols__`), והן אינן ענייננו. פנים שלנו בכפל פירושה
  // הזרקה שרצה פעמיים.
  if (report.ours !== SPECS.length) {
    errors.push(`document.fonts מכיל ${report.ours} פנים של Assistant, צפוי ${SPECS.length}`);
  }
  if (report.widthLatin === report.widthLatinSerif) {
    errors.push('רוחב הטקסט הלטיני ב-Assistant זהה ל-serif — הקובץ לא נטען בפועל');
  }
  if (report.widthHebrew === report.widthHebrewSerif) {
    errors.push('רוחב הטקסט העברי ב-Assistant זהה ל-serif — העברית לא באה מהגופן הארוז');
  }
}

console.log(
  `origin=${report?.origin ?? '?'} faces=${report?.ours ?? 0} ` +
    `latin=${report?.widthLatin}px (serif ${report?.widthLatinSerif}px) ` +
    `hebrew=${report?.widthHebrew}px (serif ${report?.widthHebrewSerif}px)`,
);

if (errors.length) {
  for (const error of errors) console.error(`שגיאה: ${error}`);
  process.exit(1);
}
console.log('שער הגופן עבר: הגופן הארוז נטען מ-file:// בכל המשקלים, בלטינית ובעברית.');

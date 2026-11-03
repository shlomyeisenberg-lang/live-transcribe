/**
 * מי חוסם את ה-thread כשדיאלוג נסגר ב-Escape בפעם השנייה.
 *
 * המדידה שקדמה (scripts/qa/file-freeze-qa.mjs) קבעה את הגבולות: סגירה
 * ראשונה נקייה, סגירה בכפתור ✕ נקייה גם פעמיים, שני Escape בלי דיאלוג
 * נקיים — ורק Escape שסוגר דיאלוג **בפעם השנייה** תוקע את הדף ל-6–11
 * שניות, בכמחצית מהניסיונות. השאלה שנשארה היא מי עושה את העבודה הזאת.
 *
 * פרופיילר דגימה ולא ניחוש: הוא ממשיך לדגום גם כשה-thread חסום, וזה בדיוק
 * החלון שמעניין. הבאנדל ממוזער, ולכן השמות חלקית מעוותים — אבל ה-URL של
 * ה-frame מפריד בין הקוד שלנו לבין המנוע, וזו ההבחנה שקובעת אצל מי לתקן.
 *
 *   node scripts/qa/dialog-stall-profile.mjs
 */
import { openApp, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9372);
/** כמה ניסיונות לפני שמוותרים: התקלה מופיעה בכמחצית מהפעמים. */
const ATTEMPTS = Number(process.env.ATTEMPTS ?? 4);

function withTimeout(promise, ms, label) {
  let timer;
  return Promise.race([
    promise.finally(() => clearTimeout(timer)),
    new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(`פג הזמן: ${label}`)), ms);
    }),
  ]);
}

async function alive(app, ms = 2_500) {
  try {
    return (await withTimeout(app.js('1+1'), ms, 'ping')) === 2;
  } catch {
    return false;
  }
}

/** זמן-עצמי לכל פונקציה, מתוך פרופיל הדגימה. */
function selfTime(profile) {
  const byId = new Map(profile.nodes.map((n) => [n.id, n]));
  const totals = new Map();

  // deltas[i] הוא הזמן שחלף עד הדגימה samples[i] — כלומר הזמן שנזקף לה.
  for (let i = 0; i < profile.samples.length; i++) {
    const node = byId.get(profile.samples[i]);
    if (!node) continue;
    const frame = node.callFrame;
    const key = `${frame.functionName || '(אנונימי)'} @ ${(frame.url || '').split('/').pop()}:${frame.lineNumber}`;
    totals.set(key, (totals.get(key) ?? 0) + (profile.timeDeltas[i] ?? 0));
  }
  return [...totals.entries()].sort((a, b) => b[1] - a[1]);
}

const app = await openApp({ name: 'stall', port: PORT });
try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false,
  });
  await sleep(400);
  await app.tab('קובץ');

  let caught = false;

  for (let attempt = 1; attempt <= ATTEMPTS && !caught; attempt++) {
    console.log(`\n────── ניסיון ${attempt} ──────`);

    // סגירה ראשונה — היא לעולם אינה תוקעת, ולכן היא מחוץ לפרופיל.
    await app.click('קיצורים');
    await app.escape();
    await sleep(1_000);
    console.log(`  אחרי הסגירה הראשונה חי? ${await alive(app)}`);

    await app.cdp.send('Profiler.enable');
    await app.cdp.send('Profiler.setSamplingInterval', { interval: 200 });
    await app.cdp.send('Profiler.start');

    await app.click('קיצורים');
    await app.escape();

    const t0 = Date.now();
    const immediate = await alive(app, 2_500);
    let blockedMs = 0;
    if (!immediate) {
      while (Date.now() - t0 < 40_000) {
        if (await alive(app, 2_000)) break;
        await sleep(500);
      }
      blockedMs = Date.now() - t0;
    }

    // הפרופיילר עצמו נעצר דרך CDP, וזה עובד גם כשה-thread חסום.
    const stopped = await withTimeout(app.cdp.send('Profiler.stop'), 30_000, 'Profiler.stop');
    const profile = stopped?.result?.profile;

    console.log(`  נחסם? ${!immediate} | משך: ${blockedMs}ms | דגימות: ${profile?.samples?.length ?? 0}`);

    if (!immediate && profile) {
      caught = true;
      const rows = selfTime(profile);
      const total = rows.reduce((sum, [, us]) => sum + us, 0);
      console.log(`\n  זמן-עצמי, 15 הראשונים (סה"כ ${Math.round(total / 1000)}ms):`);
      for (const [name, us] of rows.slice(0, 15)) {
        const ms = Math.round(us / 1000);
        if (ms < 1) continue;
        console.log(`    ${String(ms).padStart(6)}ms  ${name}`);
      }

      const log = await app.log().catch(() => []);
      console.log(`\n  מה שהדף אמר: ${JSON.stringify(log).slice(0, 600)}`);
    }

    await app.cdp.send('Profiler.disable');
    await sleep(500);
  }

  if (!caught) console.log('\nהתקלה לא נתפסה בניסיונות האלה. הריצו שוב או העלו את ATTEMPTS.');
} finally {
  app.close();
}

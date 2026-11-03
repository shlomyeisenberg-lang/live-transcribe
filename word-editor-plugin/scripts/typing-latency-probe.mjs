/**
 * בדיקת אבחון (לא שער): זמן הקלדה מלחיצת המקש עד שהתו מצויר, long tasks על
 * החוט הראשי, והודעות קונסולה בזמן ההקלדה — על ה-dist הארוז.
 *
 * זה הכלי שמצא את חישוב הסגנון של כל המסמך בכל תו (ראו README, „איטיות
 * בהקלדה”); השער הקבוע שנגזר ממנו הוא scripts/qa/typing-style-recalc-qa.mjs.
 * הכלי נשאר מפני שהוא עונה על שאלות שהשער אינו שואל: כמה זמן באמת עובר עד
 * הציור, האם התוסף שקט בקונסולה, ומה קורה **בתוך אוצריא** ולא רק ב-Chrome.
 *
 * שני מצבים:
 *   - ברירת מחדל: פותח את ה-dist ב-Chrome headless (file://) עם דמה של המאחז.
 *   - --attach <port>: מתחבר לדף התוסף שכבר רץ ב-WebView2 עם יציאת CDP. את
 *     אוצריא מריצים לפני כן עם
 *     WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=<port>.
 *     **לבדוק את שורת „פוקוס” בפלט**: `vis: hidden` פירושו חלון ממוזער — rAF
 *     לא רץ וטיימרים מוגבלים, וזה נראה כמו הקלדה איטית ואינו כזה.
 *     המקשים כאן סינתטיים (CDP); למסלול האמיתי — הקשת מערכת ההפעלה ועד הפיקסלים
 *     על המסך — יש scripts/typing-latency-inapp.mjs.
 *
 * אפשרויות:
 *   --docx <path>        פותח מסמך דרך דמה של בורר הקבצים (ברירת המחדל: המסמך הריק)
 *   --tab <label>        לשונית ברצועה לפני ההקלדה (למשל „קובץ” — בלי מנויי פקודות)
 *   --drop-css <regex>   מוחק בזמן ריצה, לפני ההקלדה, כל כלל CSS שהסלקטור שלו
 *                        תואם. למדידה בלבד — כך הוכרע איזה כלל גרם לחישוב המלא.
 *   --trace <out.json>   שומר trace של ה-renderer (devtools.timeline + invalidationTracking)
 *   --trace-light        עם --trace: בלי invalidationTracking/stack (קל יותר; למסמכים גדולים)
 *   --profile <out.json> שומר פרופיל CPU   --profile-interval <µs> (250)   --settle <ms> שהות אחרי פתיחה (3000)
 *   --open-wait <ms>     כמה להמתין לפתיחת --docx (60000)
 *   --gap <ms>           הפער בין הקשות (80)   --n1/--n2/--n3  מספר הקשות בכל שלב (60/40/40)
 *   --port <n> --label <x> --dist <dir>
 *
 * הרצה: node scripts/typing-latency-probe.mjs [--docx c:/path/doc.docx]
 */
import { readFileSync, writeFileSync, rmSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, sleep } from './cdp.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const opt = (name, dflt) => {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : dflt;
};
const DIST = opt('--dist', join(ROOT, 'dist'));
const PROFILE_OUT = opt('--profile', null);
const TRACE_OUT = opt('--trace', null);
/** trace קל: בלי invalidationTracking ובלי stack — אלה מכבידים על הדף עצמו ומסיטים את המדידה במסמכים גדולים. */
const TRACE_LIGHT = args.includes('--trace-light');
const DROP_CSS = opt('--drop-css', null);
const PORT = Number(opt('--port', 9361));
const LABEL = opt('--label', 'lat');
const GAP_MS = Number(opt('--gap', 80));
const DOCX = opt('--docx', null);
const TAB = opt('--tab', null);
const ATTACH = opt('--attach', null);
const COUNTS = [Number(opt('--n1', 60)), Number(opt('--n2', 40)), Number(opt('--n3', 40))];
/** כמה להמתין לפתיחת המסמך שנמסר ב---docx. מסמך של מאות עמודים צריך יותר מדקה. */
const OPEN_WAIT_MS = Number(opt('--open-wait', 60_000));
/** שהות אחרי שהמסמך נפתח ולפני ההקלדה. מסמך גדול „מסדר את התצוגה” כמה שניות אחרי שהטקסט כבר בעץ, והקלדה בזמן הזה מודדת את הפתיחה. */
const SETTLE_MS = Number(opt('--settle', 3000));
const PROFILE_INTERVAL_US = Number(opt('--profile-interval', 250));

const HOST_STUB = readFileSync(join(ROOT, 'scripts', 'qa', 'host-stub.js'), 'utf8');
const QA_API = readFileSync(join(ROOT, 'scripts', 'qa', 'qa-api.js'), 'utf8');

/** המדידה בתוך הדף. מותקנת דרך `install()` אחרי שהמסמך פתוח. */
const INSTRUMENT_JS = `
(function () {
  if (window.__lat) return;
  var L = (window.__lat = { keys: [], long: [], phases: [], consoleCounts: {}, visChanges: [], installed: false });
  L.mark = function (name) { L.phases.push({ name: name, t: performance.now() }); try { performance.mark('lat:' + name); } catch (e) {} return name; };
  document.addEventListener('visibilitychange', function () { L.visChanges.push({ t: performance.now(), vis: document.visibilityState }); });
  ['log', 'info', 'warn', 'error', 'debug'].forEach(function (level) {
    var orig = console[level];
    console[level] = function () {
      try {
        var key = level + ': ' + Array.prototype.slice.call(arguments, 0, 2).map(function (a) { return typeof a === 'string' ? a : (a && a.message) || typeof a; }).join(' ').slice(0, 70);
        L.consoleCounts[key] = (L.consoleCounts[key] || 0) + 1;
      } catch (e) {}
      return orig.apply(console, arguments);
    };
  });
  L.install = function () {
    if (L.installed) return 'already';
    var session = window.__otzariaEditor;
    if (!session) return 'no-editor';
    var container = session.container;
    var pending = null;
    var lenOf = function () { return container.textContent.length; };
    window.addEventListener('keydown', function (e) {
      if (e.key.length !== 1) return;
      pending = { key: e.key, down: performance.now(), mutated: null, painted: null, lenAtDown: lenOf(), vis: document.visibilityState };
      L.keys.push(pending);
    }, true);
    // „צויר” = המוטציה הראשונה שהגדילה את הטקסט, ואחריה פריים אחד.
    new MutationObserver(function () {
      if (!pending || pending.mutated !== null) return;
      if (lenOf() <= pending.lenAtDown) return;
      var p = pending;
      p.mutated = performance.now();
      requestAnimationFrame(function () { p.painted = performance.now(); });
    }).observe(container, { childList: true, subtree: true, characterData: true });
    try {
      new PerformanceObserver(function (list) {
        list.getEntries().forEach(function (e) { L.long.push({ start: Math.round(e.startTime), dur: Math.round(e.duration) }); });
      }).observe({ type: 'longtask', buffered: true });
    } catch (err) { L.longError = String(err); }
    L.installed = true;
    return 'ok';
  };
})();
`;

/** דמה של „פתח קובץ”: הבורר מחזיר blob: של הקובץ שנמסר. */
function docxStub(path) {
  const b64 = readFileSync(path).toString('base64');
  const name = path.split(/[\\/]/).pop();
  return `
<script>
(function () {
  function install() {
    if (!window.__qaHost) return setTimeout(install, 20);
    var bytes = Uint8Array.from(atob(${JSON.stringify(b64)}), function (c) { return c.charCodeAt(0); });
    var url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    window.__qaHost.replies['fs.pickUserFile'] = function () {
      return Promise.resolve({ success: true, data: { cancelled: false, token: 'probe-doc', url: url, name: ${JSON.stringify(name)}, size: bytes.length, access: 'read' }, error: null });
    };
  }
  install();
})();
</script>`;
}

const TEXT = 'בראשית ברא אלהים את השמים ואת הארץ והארץ היתה תהו ובהו וחשך על פני תהום ורוח אלהים מרחפת על פני המים ';
const vkFor = (ch) => (ch === ' ' ? { code: 'Space', vk: 32 } : { code: 'KeyA', vk: 65 });

/**
 * חיבור CDP לדף קיים. יודע לקבל אירועים — `Tracing.dataCollected` מגיע כאירוע,
 * ו-cdp.mjs מטפל רק בתשובות לפי id. CDP מרשה כמה לקוחות לאותו יעד.
 */
async function attachTo(port) {
  const list = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
  const page = list.find((t) => t.type === 'page' && /otzaria-word-editor\/dist\/index\.html|__lat-/.test(t.url));
  if (!page) throw new Error('לא נמצא דף של התוסף. דפים: ' + list.map((t) => t.type + ' ' + t.url).join(' | '));
  const socket = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true });
    socket.addEventListener('error', () => reject(new Error('CDP attach נכשל')), { once: true });
  });
  let nextId = 0;
  const pending = new Map();
  const events = [];
  const eventWaiters = [];
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data);
    if (message.id === undefined) {
      events.push(message);
      for (const w of eventWaiters.splice(0)) w(message);
      return;
    }
    const settle = pending.get(message.id);
    if (!settle) return;
    pending.delete(message.id);
    settle(message);
  });
  const cdp = {
    events,
    waitForEvent(method, timeoutMs = 60000) {
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('timeout waiting ' + method)), timeoutMs);
        const check = (m) => {
          if (m.method === method) {
            clearTimeout(timer);
            resolve(m);
          } else eventWaiters.push(check);
        };
        eventWaiters.push(check);
      });
    },
    send(method, params) {
      const id = ++nextId;
      return new Promise((resolve) => {
        pending.set(id, resolve);
        socket.send(JSON.stringify({ id, method, params }));
      });
    },
    async evaluate(expression) {
      const response = await this.send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
      const result = response.result?.result;
      if (response.result?.exceptionDetails) throw new Error(`CDP: הביטוי זרק — ${result?.description ?? 'ללא פירוט'}`);
      return result?.value;
    },
    close: () => socket.close(),
  };
  return { cdp, close: () => socket.close(), url: page.url };
}

async function click(cdp, x, y) {
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y, button: 'none', buttons: 0 });
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', buttons: 1, clickCount: 1 });
  await sleep(60);
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', buttons: 0, clickCount: 1 });
}

let probe = null;
let opened;
if (ATTACH) {
  opened = await attachTo(Number(ATTACH));
  console.log('מחובר ל-' + opened.url);
} else {
  const index = join(DIST, 'index.html');
  if (!existsSync(index)) {
    console.error('אין index.html ב-' + DIST + ' — הריצו npm run build');
    process.exit(1);
  }
  const html = readFileSync(index, 'utf8');
  const cut = html.indexOf('</script>') + '</script>'.length;
  probe = join(DIST, `__lat-${LABEL}.html`);
  const inject = `\n<script>${HOST_STUB}</script>\n<script>${QA_API}</script>\n<script>${INSTRUMENT_JS}</script>\n${DOCX ? docxStub(DOCX) : ''}\n`;
  writeFileSync(probe, html.slice(0, cut) + inject + html.slice(cut));
  opened = await openPage(`file:///${probe.split('\\').join('/')}`, { port: PORT, label: LABEL });
}

let result = null;
try {
  const { cdp, close } = opened;
  try {
    for (let waited = 0; ; waited += 500) {
      if (waited >= 90_000) throw new Error('התוסף לא הגיע למסמך פתוח בזמן');
      await sleep(500);
      if (await cdp.evaluate('!!window.__otzariaEditor && !document.getElementById("otzaria-splash") && !!document.querySelector(".superdoc-line, .superdoc-fragment")')) break;
    }
    await sleep(2500);
    if (ATTACH) {
      await cdp.evaluate('delete window.__lat');
      await cdp.evaluate(INSTRUMENT_JS);
    }

    if (DOCX) {
      // „פתח קובץ” דרך הרצועה, כמו המשתמש: לשונית „קובץ” → הפקד → „עיון בקבצים…”.
      const tabRect = JSON.parse(await cdp.evaluate(`JSON.stringify(window.__qa.tabRect('קובץ'))`));
      if (!tabRect) throw new Error('לשונית „קובץ” לא נמצאה');
      await click(cdp, tabRect.x, tabRect.y);
      await sleep(400);
      const ctrl = JSON.parse(await cdp.evaluate(`JSON.stringify(window.__qa.rect('פתח קובץ', {}))`));
      if (!ctrl) throw new Error('הפקד „פתח קובץ” לא נמצא');
      await click(cdp, ctrl.x, ctrl.y);
      await sleep(800);
      const browse = JSON.parse(
        await cdp.evaluate(
          `JSON.stringify((function(){var d=document.querySelector('[role="dialog"]');if(!d)return null;var b=Array.prototype.find.call(d.querySelectorAll('button'),function(x){return /עיון בקבצים/.test(x.textContent||'')});return b?window.__qa.rectOf(b):null;})())`,
        ),
      );
      if (browse) await click(cdp, browse.x, browse.y);
      let openedLen = 0;
      for (let waited = 0; waited < OPEN_WAIT_MS; waited += 500) {
        await sleep(500);
        openedLen = await cdp.evaluate('(window.__otzariaEditor && window.__otzariaEditor.container.textContent.length) || 0');
        if (waited % 10_000 === 0 && waited > 0) console.log(`ממתין לפתיחה… ${waited / 1000}s, טקסט ${openedLen}, עמודים ${await cdp.evaluate('document.querySelectorAll("[data-page-index]").length')}`);
        if (openedLen > 2000) break;
      }
      if (openedLen <= 2000) console.log('אזהרה: המסמך לא נפתח בזמן ההמתנה — הקונסולה עד כאן: ' + (await cdp.evaluate('JSON.stringify(window.__lat.consoleCounts)')));
      await sleep(SETTLE_MS);
    }

    if (TAB) {
      const rect = JSON.parse(await cdp.evaluate(`JSON.stringify(window.__qa.tabRect(${JSON.stringify(TAB)}))`));
      if (!rect) throw new Error(`לשונית „${TAB}” לא נמצאה`);
      await click(cdp, rect.x, rect.y);
      await sleep(500);
    }

    // שורת טקסט גלויה (השנייה, כשיש) — לחיצה על שטח ריק ממקדת בלי לפתור יעד.
    const line = await cdp.evaluate(
      '(function(){var ls=[...document.querySelectorAll(".superdoc-line")].filter(function(l){var r=l.getBoundingClientRect();return r.width>40&&r.height>8&&r.top>0&&r.top<window.innerHeight;});var l=ls[Math.min(ls.length-1,1)]||document.querySelector(".superdoc-fragment");if(!l)return "none";var r=l.getBoundingClientRect();return JSON.stringify({x:Math.round(r.x+Math.min(20,r.width/2)),y:Math.round(r.y+r.height/2)});})()',
    );
    if (line === 'none') throw new Error('לא נמצאה שורת טקסט');
    const at = JSON.parse(line);
    // ב---attach לא לוחצים: בחלון צר הדף רחב מהחלון והנקודה עלולה להיות מחוץ לו (x שלילי), ואז
    // הלחיצה נוחתת ברצועה — „מסמך חדש” בלשונית „קובץ” — והמדידה נעשית במסמך אחר. ממקדים בלי ללחוץ (למטה).
    if (!ATTACH) await click(cdp, at.x, at.y);
    await sleep(1500);
    // בתוך WebView2 הלחיצה הסינתטית לא תמיד ממקדת את משטח ההקלדה — ממקדים במפורש.
    console.log(
      'פוקוס: ' +
        (await cdp.evaluate(`(function(){
      var c = window.__otzariaEditor.container; var a = document.activeElement; var inside = a && c.contains(a);
      if (!inside) { var ta = c.querySelector('textarea, [contenteditable="true"]'); if (ta) { ta.focus(); a = document.activeElement; inside = c.contains(a); } }
      return JSON.stringify({ vis: document.visibilityState, hasFocus: document.hasFocus(), active: a ? a.tagName : null, inside: inside });
    })()`)),
    );

    if (DROP_CSS) {
      const dropped = await cdp.evaluate(`(function(){
        var re = new RegExp(${JSON.stringify(DROP_CSS)}); var out = [];
        for (var i = 0; i < document.styleSheets.length; i++) {
          var sheet = document.styleSheets[i]; var rules;
          try { rules = sheet.cssRules; } catch (e) { continue; }
          for (var j = rules.length - 1; j >= 0; j--) {
            var r = rules[j];
            if (r.selectorText && re.test(r.selectorText)) { out.push(r.selectorText.slice(0, 160)); sheet.deleteRule(j); }
          }
        }
        return JSON.stringify(out);
      })()`);
      console.log('כללי CSS שהוסרו לצורך המדידה: ' + dropped);
      await sleep(500);
    }

    const installed = await cdp.evaluate('window.__lat.install()');
    if (installed !== 'ok') throw new Error('התקנת המדידה נכשלה: ' + installed);
    // שגיאות שנצברו עד כאן — בפתיחה — נדפסות לפני האיפוס, כדי שפתיחה שנכשלה לא תיעלם.
    const beforeTyping = JSON.parse(await cdp.evaluate('JSON.stringify(window.__lat.consoleCounts)'));
    const loud = Object.entries(beforeTyping).filter(([k]) => /^(error|warn)/.test(k));
    if (loud.length) console.log('קונסולה לפני ההקלדה: ' + loud.map(([k, v]) => `${v}× ${k}`).join(' | '));
    await cdp.evaluate('window.__lat.consoleCounts = {}');

    let tracer = null;
    if (TRACE_OUT) {
      tracer = await attachTo(ATTACH ? Number(ATTACH) : PORT);
      await tracer.cdp.send('Tracing.start', {
        categories: TRACE_LIGHT
          ? 'devtools.timeline,disabled-by-default-devtools.timeline,blink.user_timing'
          : 'devtools.timeline,disabled-by-default-devtools.timeline,disabled-by-default-devtools.timeline.invalidationTracking,blink.user_timing,disabled-by-default-devtools.timeline.stack',
        transferMode: 'ReportEvents',
      });
    }
    if (PROFILE_OUT) {
      await cdp.send('Profiler.enable');
      await cdp.send('Profiler.setSamplingInterval', { interval: PROFILE_INTERVAL_US });
      await cdp.send('Profiler.start');
    }

    const cdpTimes = [];
    const typeChars = async (count, phase) => {
      await cdp.evaluate(`window.__lat.mark(${JSON.stringify(phase)})`);
      for (let i = 0; i < count; i++) {
        const ch = TEXT[i % TEXT.length];
        const { code, vk } = vkFor(ch);
        const t0 = Date.now();
        await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: ch, code, text: ch, unmodifiedText: ch, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk });
        await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: ch, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk });
        cdpTimes.push({ phase, ms: Date.now() - t0 });
        await sleep(GAP_MS);
      }
    };

    // שלושה שלבים: רצוף; אחרי הפסקה קצרה (autosave ב-2.5s); אחרי 12s (כתיבת טיוטה ב-10s).
    await typeChars(COUNTS[0], 'A-רצוף');
    await cdp.evaluate('window.__lat.mark("B-הפסקה 4s")');
    await sleep(4000);
    await typeChars(COUNTS[1], 'C-אחרי הפסקה');
    await cdp.evaluate('window.__lat.mark("D-הפסקה 12s")');
    await sleep(12_000);
    await typeChars(COUNTS[2], 'E-אחרי טיוטה');
    await cdp.evaluate('window.__lat.mark("F-סוף")');
    await sleep(3000);

    if (PROFILE_OUT) {
      const { result: prof } = await cdp.send('Profiler.stop');
      writeFileSync(PROFILE_OUT, JSON.stringify(prof.profile));
    }
    if (tracer) {
      await tracer.cdp.send('Tracing.end');
      await tracer.cdp.waitForEvent('Tracing.tracingComplete', 120000);
      const all = [];
      for (const ev of tracer.cdp.events) if (ev.method === 'Tracing.dataCollected') all.push(...ev.params.value);
      writeFileSync(TRACE_OUT, JSON.stringify(all));
      console.log('trace: ' + all.length + ' אירועים → ' + TRACE_OUT);
      tracer.close();
    }

    result = JSON.parse(
      await cdp.evaluate(
        'JSON.stringify({ lat: window.__lat, calls: (window.__qaHost && window.__qaHost.calls) || [], pages: document.querySelectorAll("[data-page-index]").length, textLen: window.__otzariaEditor.container.textContent.length })',
      ),
    );
    result.cdpTimes = cdpTimes;
  } finally {
    close();
  }
} finally {
  if (probe) rmSync(probe, { force: true });
}

/* ------------------------------ החשבון ------------------------------ */
const { lat, calls } = result;
const t0 = lat.keys[0]?.down ?? 0;
const phaseOf = (t) => {
  let name = '?';
  for (const p of lat.phases) if (p.t <= t) name = p.name;
  return name;
};
const stat = (xs) => {
  if (!xs.length) return 'אין';
  const s = [...xs].sort((a, b) => a - b);
  const q = (p) => s[Math.min(s.length - 1, Math.floor(p * s.length))];
  return `n=${s.length} חציון=${Math.round(q(0.5))} p90=${Math.round(q(0.9))} מקס=${Math.round(s[s.length - 1])}`;
};

console.log(`עמודים: ${result.pages}, אורך טקסט: ${result.textLen}`);
console.log('\n== זמן מלחיצה עד צביעה (ms), לפי שלב ==');
const byPhase = new Map();
for (const k of lat.keys) {
  const ph = phaseOf(k.down) + (k.vis ? ` (${k.vis})` : '');
  if (!byPhase.has(ph)) byPhase.set(ph, []);
  if (k.painted !== null) byPhase.get(ph).push(k.painted - k.down);
}
for (const [ph, xs] of byPhase) console.log(`${ph}: ${stat(xs)}`);
console.log(`תווים שלא נצפו כמצוירים בנפרד (התמזגו עם הבא): ${lat.keys.filter((k) => k.painted === null).length} מתוך ${lat.keys.length}`);

console.log('\n== CDP: זמן קריאת dispatchKeyEvent (ms) ==');
const cdpBy = new Map();
for (const c of result.cdpTimes) {
  if (!cdpBy.has(c.phase)) cdpBy.set(c.phase, []);
  cdpBy.get(c.phase).push(c.ms);
}
for (const [ph, xs] of cdpBy) console.log(`${ph}: ${stat(xs)}`);

console.log('\n== הקשות איטיות (>150ms עד צביעה) ==');
for (const k of lat.keys) {
  if (k.painted === null || k.painted - k.down <= 150) continue;
  console.log(`t=${Math.round(k.down - t0)}ms ‚${k.key}‘ צביעה+${Math.round(k.painted - k.down)} [${phaseOf(k.down)}]`);
}

console.log(`\n== long tasks (>50ms): ${lat.long.filter((l) => l.start >= t0 - 500).length} ==`);
if (lat.longError) console.log('longtask לא נתמך: ' + lat.longError);
for (const l of lat.long) {
  if (l.start < t0 - 500) continue;
  console.log(`t=${Math.round(l.start - t0)}ms משך ${l.dur}ms [${phaseOf(l.start)}]`);
}

console.log('\n== הודעות קונסולה בזמן ההקלדה ==');
for (const [k, v] of Object.entries(lat.consoleCounts).sort((a, b) => b[1] - a[1]).slice(0, 12)) console.log(`${v}× ${k}`);

console.log('\n== קריאות למאחז מאז תחילת ההקלדה ==');
for (const c of calls) {
  if (c.t < t0 - 500) continue;
  console.log(`t=${Math.round(c.t - t0)}ms ${c.method} [${phaseOf(c.t)}]`);
}
console.log('\n== שינויי נראות בזמן הריצה ==');
for (const v of lat.visChanges) console.log(`t=${Math.round(v.t - t0)}ms ${v.vis}`);
console.log('\n== שלבים ==');
for (const p of lat.phases) console.log(`t=${Math.round(p.t - t0)}ms ${p.name}`);

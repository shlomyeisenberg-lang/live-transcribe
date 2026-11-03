/**
 * מי בולע את `Alt+<אות>` כשהסמן במסמך.
 *
 * כל מאזין `keydown` שנרשם בדף נעטף **לפני** שהבאנדל עולה, והעטיפה מודדת את
 * `defaultPrevented` לפני ואחרי. מי שהפך אותו — הוא הבעלים. הערימה של
 * `addEventListener` נשמרת, ולכן גם ידוע מאיזה מודול הוא בא.
 *
 *   node scripts/qa/custom-shortcut-owner-probe.mjs
 */
import { openApp, sleep } from './harness.mjs';

const ALT = 1;
const SHIFT = 8;

const EXTRA = `<script>
(function () {
  var seq = 0;
  window.__owners = [];
  var raw = EventTarget.prototype.addEventListener;
  EventTarget.prototype.addEventListener = function (type, fn, opts) {
    if (type !== 'keydown' || typeof fn !== 'function') return raw.call(this, type, fn, opts);
    var id = ++seq;
    var where = '';
    try { throw new Error('x'); } catch (e) { where = String(e.stack || '').split('\\n').slice(2, 5).join(' | '); }
    var capture = opts === true || (opts && opts.capture === true);
    var tag = this === window ? 'window' : this === document ? 'document' :
      (this.tagName ? this.tagName + '.' + String(this.className || '').slice(0, 30) : String(this));
    var wrapped = function (e) {
      var before = e.defaultPrevented;
      var out = fn.apply(this, arguments);
      if (!before && e.defaultPrevented && window.__catch) {
        window.__catch.push({ id: id, tag: tag, capture: capture, where: where });
      }
      return out;
    };
    fn.__wrapped = wrapped;
    window.__owners.push({ id: id, tag: tag, capture: capture, where: where });
    return raw.call(this, type, wrapped, opts);
  };
  var rawRemove = EventTarget.prototype.removeEventListener;
  EventTarget.prototype.removeEventListener = function (type, fn, opts) {
    if (type === 'keydown' && typeof fn === 'function' && fn.__wrapped) {
      return rawRemove.call(this, type, fn.__wrapped, opts);
    }
    return rawRemove.call(this, type, fn, opts);
  };
})();
</script>`;

const app = await openApp({ name: 'csowner', port: 9664, extra: EXTRA });
try {
  await app.caret(0);
  await app.type('abcd', 60);
  await sleep(600);
  await app.press('Home', 'Home', 36);
  await sleep(200);
  for (let i = 0; i < 4; i += 1) {
    await app.press('ArrowRight', 'ArrowRight', 39, SHIFT);
    await sleep(60);
  }
  await sleep(600);

  console.log('מאזיני keydown רשומים:');
  const owners = JSON.parse(await app.js('JSON.stringify(window.__owners)'));
  for (const o of owners) console.log(`  #${o.id} ${o.tag} capture=${o.capture}\n      ${o.where}`);

  await app.js('window.__catch = []');
  await app.press('x', 'KeyX', 88, ALT);
  await sleep(1_000);
  console.log('\nמי בלע את Alt+X:');
  console.log(await app.js('JSON.stringify(window.__catch, null, 1)'));
  console.log('טקסט:', await app.js("(document.querySelector('.editor-stack').innerText||'').replace(/\\s+/g,' ').slice(0,80)"));

  await app.js('window.__catch = []');
  await app.press('7', 'Digit7', 55, ALT);
  await sleep(1_000);
  console.log('\nמי בלע את Alt+7:');
  console.log(await app.js('JSON.stringify(window.__catch, null, 1)'));
  console.log('טקסט:', await app.js("(document.querySelector('.editor-stack').innerText||'').replace(/\\s+/g,' ').slice(0,80)"));
} finally {
  app.close();
}

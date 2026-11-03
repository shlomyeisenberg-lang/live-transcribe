/** Ambiguous and edited corrections must never offer another word's original. */
import { openApp, createReport } from './harness.mjs';
const report = createReport('זהות מחוון תיקון אוטומטי', { strict: true });
let app = await openApp({ name: 'acorr-identity', port: Number(process.env.QA_PORT ?? 9636) });

async function hoverCopyright(nth, blockId = null) {
  const box = JSON.parse(await app.js(`(() => {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let node, count = 0;
    while ((node = walker.nextNode())) {
      const fragment = node.parentElement?.closest('[data-source-node-id][data-pm-start]');
      if (!fragment || (${JSON.stringify(blockId)} !== null && fragment.getAttribute('data-source-node-id') !== ${JSON.stringify(blockId)})) continue;
      for (let i = 0; i < node.length; i++) {
        if (node.nodeValue[i] !== '©' || count++ !== ${nth}) continue;
        const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1);
        const r = range.getBoundingClientRect();
        return JSON.stringify({ x:(r.left+r.right)/2, y:(r.top+r.bottom)/2 });
      }
    }
    return 'null';
  })()`));
  if (!box) throw new Error('Copyright word was not found');
  await app.js(`document.elementFromPoint(${box.x},${box.y}).dispatchEvent(new PointerEvent('pointermove',{bubbles:true,clientX:${box.x},clientY:${box.y}}))`);
  await app.sleep(150);
  return app.js("document.querySelector('.otzaria-acorr-chip').style.display === 'block'");
}

try {
  await app.caret(0);
  await app.type('(c) (c) (c) (C) ');
  await app.sleep(900);
  console.log('בודק מחוונים במופעים זהים');
  if (!(await hoverCopyright(2))) throw new Error('Uniquely identified third correction was not marked');
  report.pass('המופע המזוהה חד משמעית מציג מחוון');
  if (await hoverCopyright(3)) report.fail('מופע עמום אינו מציע ביטול');
  else report.pass('מופע עמום אינו מציע ביטול');
  console.log('בודק סגירת תפריט במקלדת');
  await hoverCopyright(2);
  await app.clickSel('.otzaria-acorr-chip');
  await app.press('Shift', 'ShiftLeft', 16);
  if (await app.js("document.querySelector('.otzaria-acorr-menu').style.display === 'block'")) report.fail('מקש בעורך סוגר את התפריט שנמדד קודם');
  else report.pass('מקש בעורך סוגר את התפריט שנמדד קודם');

  console.log('בודק פסילת מחוון אחרי מחיקה');
  app.close();
  app = await openApp({ name: 'acorr-identity-delete', port: Number(process.env.QA_PORT ?? 9636) });
  await app.caret(0);
  await app.type('(c) (C) (C) ');
  await app.sleep(900);
  await app.press('Home', 'Home', 36);
  await app.press('Delete', 'Delete', 46);
  await app.press('Delete', 'Delete', 46);
  await app.sleep(700);
  const remaining = await app.js("[...document.querySelectorAll('[data-source-node-id][data-pm-start]')].map(el => el.textContent).join('')");
  if (remaining !== '© © ') throw new Error(`The first correction was not deleted: ${JSON.stringify(remaining)}`);
  if (await hoverCopyright(0)) report.fail('מחיקה אינה משייכת את המקור של המילה שנמחקה למילה ששרדה');
  else report.pass('מחיקה אינה משייכת את המקור של המילה שנמחקה למילה ששרדה');
} finally {
  app.close();
  report.print();
}

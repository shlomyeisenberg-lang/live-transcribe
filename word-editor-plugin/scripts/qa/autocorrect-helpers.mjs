/** The personal Hebrew preset is enabled explicitly through the real dialog. */
export async function loadHebrewPreset(app) {
  await app.tab('סקירה');
  if (!(await app.click('תיקון אוטומטי'))) throw new Error('AutoCorrect dialog did not open');
  await app.js("document.querySelector('.acorr-load-hebrew').scrollIntoView({ block:'nearest' })");
  if (!(await app.clickSel('.acorr-load-hebrew'))) throw new Error('Hebrew preset button was unavailable');
  if (!(await app.js("Array.from(document.querySelectorAll('.acorr-row')).some(row => row.textContent.includes('אאכ'))"))) throw new Error('Hebrew preset was not loaded');
  if (!(await app.clickSel('.acorr-btn-primary'))) throw new Error('AutoCorrect options were not applied');
  if (await app.js("!!document.querySelector('.acorr-dialog')")) throw new Error('AutoCorrect dialog did not close');
}

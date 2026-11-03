/**
 * ה-DOM שמנוע ה-DOCX מצייר בשכבת הכותרת העליונה והתחתונה, כפי שהוא.
 *
 * למה מודול משותף: שני צרכנים צריכים בדיוק את אותו עץ — tests/unit/hf-chrome
 * (העברות עצמה) ו-tests/unit/create-editor (החיווט, כלומר שהעברות מותקנת על
 * ה-container של המנוע ומפורקת איתו). שני העתקים היו נבדלים בשקט, וברגע כזה
 * אחת מהבדיקות מודדת מבנה שאינו קיים.
 *
 * המבנה כאן אינו המצאה: הוא הועתק מהאריזה של @superdoc/docx-engine v0.7.0 —
 * סדר הילדים, המחלקות, התכונות והמחרוזות האנגליות. שדרוג מנוע שישנה אותו
 * ייתפס ב-tests/contract/engine-hf-chrome.test.ts, שקורא את האריזה עצמה;
 * כאן אין קריאה של האריזה בכוונה, מפני שבדיקת יחידה צריכה עץ ודאי ומהיר.
 */

export interface ChromeOptions {
  /** `data-sd-hf-region` על העוטף. ערך שאינו header/footer בודק fail-soft. */
  region?: string;
  /** `data-sd-hf-variant` על העוטף. */
  variant?: string;
  /** הטקסט שהמנוע כותב בתג הצף, כולל סיומת מקטע כשיש. */
  label?: string;
  /** היחידה שהמנוע כותב בשורות המרחק: `in` או `cm`. */
  unit?: string;
  /**
   * תג ההמשך שבעמודים האחרים. `null` — אין אזור המשך.
   *
   * המנוע קורא שם ל-`u_e(kind, 'default')` בלי אפשרויות (נמדד באריזה), ולכן
   * בפועל אין בתג הזה סיומת מקטע ואין עליו region/variant. בדיקה שמוסרת כאן
   * „Footer -Section 2-” מודדת את הפרסר ולא את המנוע — וזה בכוונה: הדקדוק
   * הוא מקור אחד, והפרסר צריך להיות נכון עליו כולו.
   */
  continuation?: string | null;
}

/**
 * העוטף הפעיל, התג הצף, פאנל האפשרויות ותג ההמשך.
 *
 * הפאנל נבנה תמיד: העברות שלו היא CSS, ובדיקות היחידה נשענות על כך שהוא
 * נמצא בעץ **ולא** מתורגם ב-JS — זה מה שמוכיח שאין תווית בשני מקומות.
 */
export function headerFooterChrome(options: ChromeOptions = {}): HTMLElement {
  const {
    region = 'header',
    variant = 'default',
    label = 'Header',
    unit = 'cm',
    continuation = 'Footer',
  } = options;

  const overlay = document.createElement('div');
  overlay.className = 'v2-hf-overlay';
  overlay.innerHTML = `
    <div class="v2-hf-active"
         data-sd-header-footer-active="true"
         data-sd-hf-region="${region}"
         data-sd-hf-variant="${variant}"
         data-sd-hf-ref-id="rId7"
         role="group"
         aria-label="Header and footer controls">
      <div class="v2-hf-label" data-sd-hf-label="">
        <span>${label}</span>
        <button class="v2-hf-options-btn" data-sd-hf-options="" type="button"
                title="Header and footer options">Options ▾</button>
        <button class="v2-hf-exit" data-sd-hf-exit="" type="button"
                title="Close header and footer">×</button>
      </div>
      <div class="v2-hf-options-panel" data-sd-hf-options-panel="">
        <label class="v2-hf-option" data-sd-hf-option="different-first-page">
          <input type="checkbox"><span>Different First Page</span>
        </label>
        <label class="v2-hf-option" data-sd-hf-option="different-odd-even">
          <input type="checkbox"><span>Different Odd &amp; Even Pages</span>
        </label>
        <div class="v2-hf-options-divider"></div>
        <label class="v2-hf-distance" data-sd-hf-option="header-from-top">
          <span>Header from Top</span><input type="number"><span class="v2-hf-distance-unit">${unit}</span>
        </label>
        <label class="v2-hf-distance" data-sd-hf-option="footer-from-bottom">
          <span>Footer from Bottom</span><input type="number"><span class="v2-hf-distance-unit">${unit}</span>
        </label>
      </div>
    </div>
    ${
      continuation === null
        ? ''
        : `<div class="v2-hf-active">
             <div class="v2-hf-line v2-hf-line--footer"></div>
             <div class="v2-hf-label v2-hf-label--continuation" data-sd-hf-continuation-label="">
               <span>${continuation}</span>
             </div>
           </div>`
    }
  `;
  return overlay;
}

/** הטקסט של אלמנט, בלי רווחי ה-HTML שסביבו. */
export function textOf(root: ParentNode, selector: string): string {
  const element = root.querySelector(selector);
  if (element === null) throw new Error(`אין אלמנט שמתאים ל-${selector}`);
  return (element.textContent ?? '').trim();
}

/**
 * הגופן שנארז עם התוסף.
 */

/** קובץ גופן אחד ומשקלו. */
export interface BundledFile {
  weight: string;
  /** נתיב יחסי ל-document. חייב להישאר יחסי — `/` מוחלט נשבר ב-file://. */
  url: string;
}

/**
 * ארבעת המשקלים שנארזים — בדיוק אלה שה-CSS של הממשק מבקש.
 */
export const BUNDLED_FILES: readonly BundledFile[] = [
  { weight: '400', url: './fonts/Assistant-Regular.ttf' },
  { weight: '500', url: './fonts/Assistant-Medium.ttf' },
  { weight: '600', url: './fonts/Assistant-SemiBold.ttf' },
  { weight: '700', url: './fonts/Assistant-Bold.ttf' },
];

/**
 * השם שבו ה-CSS והמסמכים קוראים לגופן.
 */
export const BUNDLED_FAMILIES: readonly string[] = ['Assistant'];

const STYLE_ID = 'bundled-fonts';

/** בלוק ה-`@font-face` — פנים לכל משקל, לכל שם. */
export function bundledFontFaceCss(
  files: readonly BundledFile[] = BUNDLED_FILES,
  families: readonly string[] = BUNDLED_FAMILIES,
): string {
  return families
    .flatMap((family) =>
      files.map(
        (file) =>
          `@font-face{font-family:'${family}';font-style:normal;` +
          `font-weight:${file.weight};src:url('${file.url}') format('truetype');` +
          `font-display:swap}`,
      ),
    )
    .join('\n');
}

/**
 * מזריקה את ההצהרות ל-`<head>`. אידמפוטנטית: קריאה שנייה אינה מכפילה אותן.
 */
export function installBundledFonts(doc: Document = document): void {
  const parent = doc.head ?? doc.documentElement;
  if (!parent || doc.getElementById(STYLE_ID)) return;

  const style = doc.createElement('style');
  style.id = STYLE_ID;
  style.textContent = bundledFontFaceCss();
  parent.appendChild(style);
}

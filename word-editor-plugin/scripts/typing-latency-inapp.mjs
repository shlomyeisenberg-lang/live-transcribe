/**
 * מדידת הקלדה **בתוך אוצריא**, מקצה לקצה: הקשה אמיתית של מערכת ההפעלה → keydown בדף →
 * ציור ב-renderer → פיקסלים על המסך. משלים את typing-latency-probe.mjs (--attach), שמודד
 * רק מה שקורה בתוך הדף ובמקשים סינתטיים של CDP.
 *
 * למה צריך את זה: הדף יכול לצייר תוך 30ms ובכל זאת התו יופיע למשתמש מאוחר — כי המארח
 * (WebView2 → Windows.Graphics.Capture → טקסטורה של Flutter → raster) מוסיף שלב משלו,
 * וכי חלון ממוזער או WebView מושהה הופכים את הדף ל-`hidden` (rAF לא רץ). שני אלה
 * אינם נראים במדידה דרך CDP בלבד.
 *
 * הכנה: להריץ את אוצריא עם WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9444,
 * לפתוח את התוסף עם מסמך, ולא לגעת במחשב בזמן הריצה (הכלי מביא את החלון לחזית).
 * ההקשות נכנסות למסמך הפעיל — לא להריץ על מסמך שאין לו גיבוי.
 *
 * אפשרויות:
 *   --port <n>       יציאת CDP של WebView2 (9444)
 *   --keys <n>       מספר הקשות (40); 0 = רק חיבור, מכשור ובדיקת מצב (בלי קלט)
 *   --gap <ms>       מרווח בין הקשות (80; עם --screen לפחות 350)
 *   --screen         גם דגימת המסך: כמה זמן עד שהפיקסלים בשורת הסמן משתנים
 *   --no-click       בלי לחיצת עכבר אמיתית לפני ההקלדה (ברירת המחדל לוחצת על שורת
 *                    טקסט גלויה — פוקוס JS בלבד אינו מעביר את פוקוס המקלדת של
 *                    המערכת אל חלון ה-WebView, וההקשות נופלות ב-Flutter)
 *   --target <מה>    איזו שורה ללחוץ: auto (השורה הגלויה השלישית), wide (שורה ברוחב
 *                    העמוד — כותרת או מקטע של טור אחד), left / right (שורה בטור
 *                    השמאלי/הימני של מקטע טורים; ב-SuperDoc השמאלי הוא הראשון בסדר
 *                    המסמך), text:<קטע> (שורה שמכילה את הטקסט). שורה שאינה גלויה
 *                    מגוללת אל המרכז. הדוח מציין מה נלחץ.
 *
 * הרצה: node scripts/typing-latency-inapp.mjs --screen
 */
import { readFileSync, writeFileSync, mkdtempSync, rmSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const args = process.argv.slice(2);
const opt = (name, dflt) => {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : dflt;
};
const PORT = Number(opt('--port', 9444));
const KEYS = Number(opt('--keys', 40));
const SCREEN = args.includes('--screen');
const GAP_MS = Math.max(Number(opt('--gap', 80)), SCREEN ? 350 : 0);
const CLICK = !args.includes('--no-click');
const TARGET = opt('--target', 'auto');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/* ------------------------------ CDP ------------------------------ */
const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`).catch(() => null))?.json?.() ?? null;
if (!list) {
  console.error(`אין CDP ב-127.0.0.1:${PORT}. להריץ את אוצריא עם WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=${PORT}`);
  process.exit(1);
}
const page = list.find((t) => t.type === 'page' && /otzaria-word-editor\/dist\/index\.html/.test(t.url));
if (!page) {
  console.error('דף התוסף לא נמצא. דפים: ' + list.map((t) => `${t.type} ${t.url}`).join(' | '));
  process.exit(1);
}
const ws = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((resolve, reject) => {
  ws.addEventListener('open', resolve, { once: true });
  ws.addEventListener('error', () => reject(new Error('חיבור CDP נכשל')), { once: true });
});
let nextId = 0;
const pending = new Map();
ws.addEventListener('message', (e) => {
  const m = JSON.parse(e.data);
  if (m.id && pending.has(m.id)) {
    pending.get(m.id)(m);
    pending.delete(m.id);
  }
});
const send = (method, params) =>
  new Promise((resolve) => {
    const id = ++nextId;
    pending.set(id, resolve);
    ws.send(JSON.stringify({ id, method, params }));
  });
const js = async (expression) => {
  const m = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
  if (m.result?.exceptionDetails) throw new Error('CDP: ' + (m.result.exceptionDetails.exception?.description ?? m.result.exceptionDetails.text));
  return m.result?.result?.value;
};

/* ------------------------------ החלון ------------------------------ */
// חלון ממוזער = דף `hidden`; משחזרים אותו לפני בדיקת המצב, כמו שהמשתמש היה עושה.
const hwndOf = () => {
  const r = spawnSync('powershell', ['-NoProfile', '-Command', '(Get-Process otzaria -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1).MainWindowHandle'], { encoding: 'utf8' });
  const n = Number((r.stdout || '').trim());
  return Number.isFinite(n) && n > 0 ? n : 0;
};
const HWND = Number(opt('--hwnd', 0)) || hwndOf();
if (!HWND) {
  console.error('לא נמצא חלון של otzaria.exe');
  ws.close();
  process.exit(2);
}
{
  const r = spawnSync('powershell', ['-NoProfile', '-Command', `Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class Wn { [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h); }'; $h=[IntPtr]${HWND}; if ([Wn]::IsIconic($h)) { [void][Wn]::ShowWindow($h, 9); Start-Sleep -Milliseconds 400; 'restored' } else { 'shown' }`], { encoding: 'utf8' });
  const word = (r.stdout || '').trim();
  if (word === 'restored') {
    console.log('החלון היה ממוזער — שוחזר.');
    await sleep(800);
  }
}

/* ------------------------------ מצב הדף ------------------------------ */
const state = JSON.parse(
  await js(`JSON.stringify({
    vis: document.visibilityState, focus: document.hasFocus(),
    active: document.activeElement ? document.activeElement.tagName : null,
    tab: ((document.querySelector('.word-doctab.active') || {}).textContent || '').trim(),
    inner: [innerWidth, innerHeight], screen: [screenX, screenY],
    text: window.__otzariaEditor ? window.__otzariaEditor.container.textContent.length : -1,
    pages: document.querySelectorAll('[data-page-index]').length,
  })`),
);
console.log(`דף: ${page.url}`);
console.log(`מצב: נראות ${state.vis}, פוקוס ${state.focus}, לשונית „${state.tab}”, ${state.pages} עמודים, ${state.text} תווים ב-DOM, חלון ${state.inner.join('×')} במסך (${state.screen.join(',')})`);
if (state.vis !== 'visible') {
  console.error('הדף `hidden` — חלון ממוזער או WebView מושהה. במצב הזה rAF לא רץ והמדידה תראה „הקלדה איטית” שאינה כזו. לשחזר את החלון ולהריץ שוב.');
  ws.close();
  process.exit(2);
}
if (state.text < 0) {
  console.error('אין עורך פתוח (window.__otzariaEditor). לפתוח מסמך בתוסף ולהריץ שוב.');
  ws.close();
  process.exit(2);
}

/* ------------------------------ מכשור בדף ------------------------------ */
await js(`(function () {
  delete window.__inapp;
  var O = (window.__inapp = { keys: [], text: [], long: [], vis: [] });
  var epoch = function (t) { return performance.timeOrigin + t; };
  var c = window.__otzariaEditor.container;
  var last = c.textContent;
  document.addEventListener('keydown', function (e) {
    O.keys.push({ key: e.key, down: epoch(performance.now()) });
  }, true);
  // המוטציה הראשונה אחרי הקשה היא בדרך כלל שכבת הסמן/הבחירה, לא התו (נמדד: היא
  // מגיעה ~10ms אחרי keydown, והתו עצמו ~50–150ms אחריו). לכן נרשם רק שינוי בטקסט
  // של המסמך, וה-rAF הראשון אחריו הוא הציור. הרישום הוא סדרה, לא „ההקשה האחרונה”:
  // בהקלדה מהירה התו של הקשה N מגיע אחרי keydown של N+1.
  new MutationObserver(function () {
    var t = c.textContent; if (t === last) return;
    var rec = { t: epoch(performance.now()), d: t.length - last.length, painted: null };
    last = t; O.text.push(rec);
    requestAnimationFrame(function () { rec.painted = epoch(performance.now()); });
  }).observe(c, { childList: true, subtree: true, characterData: true, attributes: true });
  document.addEventListener('visibilitychange', function () { O.vis.push({ t: Date.now(), vis: document.visibilityState }); });
  try {
    new PerformanceObserver(function (l) { l.getEntries().forEach(function (en) { O.long.push({ start: epoch(en.startTime), dur: Math.round(en.duration) }); }); }).observe({ entryTypes: ['longtask'] });
  } catch (e) { O.longError = String(e); }
  return 'ok';
})()`);

/* ------------------------------ PowerShell: קלט אמיתי ומסך ------------------------------ */
// שני שלבים: (1) לחיצת עכבר על שורת טקסט גלויה, כדי שפוקוס המקלדת של המערכת יעבור לחלון
// ה-WebView (Chrome_WidgetWin_0); (2) הקשות SendInput ב-KEYEVENTF_UNICODE עם חותמת זמן
// לכל הקשה, ועם --screen גם דגימת המסך עד שינוי בשורת הסמן (סף 60 פיקסלים: אות עברית
// קטנה משנה כ-100, הבהוב הסמן כ-30 והוא מסונן בבסיס הכפול). INPUT הוא 40 בתים ב-x64
// (union של 32) — גודל שגוי מחזיר 0 עם ERROR_INVALID_PARAMETER ואין קלט בכלל.
const PS = String.raw`
param([string]$Mode, [int]$Hwnd, [int]$X, [int]$Y, [int]$Count, [int]$GapMs, [int]$RX, [int]$RY, [int]$RW, [int]$RH, [string]$Out)
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System; using System.Runtime.InteropServices; using System.Text;
public static class N {
  [StructLayout(LayoutKind.Sequential)] public struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
  [StructLayout(LayoutKind.Sequential)] public struct MOUSEINPUT { public int dx; public int dy; public uint mouseData; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
  [StructLayout(LayoutKind.Explicit)] public struct INPUTUNION { [FieldOffset(0)] public KEYBDINPUT ki; [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public long p0; [FieldOffset(8)] public long p1; [FieldOffset(16)] public long p2; [FieldOffset(24)] public long p3; }
  [StructLayout(LayoutKind.Sequential)] public struct INPUT { public uint type; public INPUTUNION u; }
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [StructLayout(LayoutKind.Sequential)] public struct GUITHREADINFO { public uint cbSize; public uint flags; public IntPtr hwndActive; public IntPtr hwndFocus; public IntPtr hwndCapture; public IntPtr hwndMenuOwner; public IntPtr hwndMoveSize; public IntPtr hwndCaret; public RECT rcCaret; }
  [DllImport("user32.dll", SetLastError=true)] public static extern uint SendInput(uint n, INPUT[] inputs, int size);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool GetGUIThreadInfo(uint tid, ref GUITHREADINFO info);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  static int Size() { return Marshal.SizeOf(typeof(INPUT)); }
  public static string SendChar(char c) {
    var d = new INPUT(); d.type = 1; d.u.ki.wScan = c; d.u.ki.dwFlags = 0x0004;
    var u = new INPUT(); u.type = 1; u.u.ki.wScan = c; u.u.ki.dwFlags = 0x0006;
    uint n = SendInput(2, new INPUT[] { d, u }, Size());
    return n == 2 ? "ok" : ("SendInput=" + n + " err=" + Marshal.GetLastWin32Error());
  }
  public static string Click(int x, int y) {
    SetCursorPos(x, y); System.Threading.Thread.Sleep(80);
    var d = new INPUT(); d.type = 0; d.u.mi.dwFlags = 0x0002; var u = new INPUT(); u.type = 0; u.u.mi.dwFlags = 0x0004;
    uint n = SendInput(1, new INPUT[] { d }, Size()); System.Threading.Thread.Sleep(40); n += SendInput(1, new INPUT[] { u }, Size());
    return n == 2 ? "ok" : ("SendInput=" + n + " err=" + Marshal.GetLastWin32Error());
  }
  public static string FocusClass() {
    IntPtr fg = GetForegroundWindow(); uint pid; uint tid = GetWindowThreadProcessId(fg, out pid);
    var g = new GUITHREADINFO(); g.cbSize = (uint)Marshal.SizeOf(typeof(GUITHREADINFO));
    if (!GetGUIThreadInfo(tid, ref g)) return "?";
    var sb = new StringBuilder(256); GetClassNameW(g.hwndFocus, sb, 256); return sb.ToString();
  }
  public static int Diff(System.Drawing.Bitmap a, System.Drawing.Bitmap b) {
    var rect = new System.Drawing.Rectangle(0, 0, a.Width, a.Height);
    var ra = a.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
    var rb = b.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
    int n = a.Width * a.Height; int[] pa = new int[n]; int[] pb = new int[n];
    Marshal.Copy(ra.Scan0, pa, 0, n); Marshal.Copy(rb.Scan0, pb, 0, n);
    a.UnlockBits(ra); b.UnlockBits(rb);
    int d = 0; for (int i = 0; i < n; i++) if (pa[i] != pb[i]) d++; return d;
  }
}
"@
function Grab { $bmp = New-Object System.Drawing.Bitmap $RW, $RH, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb); $g = [System.Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen($RX, $RY, 0, 0, (New-Object System.Drawing.Size $RW, $RH)); $g.Dispose(); return $bmp }
$h = [IntPtr]$Hwnd
if ([N]::IsIconic($h)) { [void][N]::ShowWindow($h, 9); Start-Sleep -Milliseconds 300 }
[void][N]::SetForegroundWindow($h); Start-Sleep -Milliseconds 400
$result = @{ fg = [int64][N]::GetForegroundWindow(); focusBefore = [N]::FocusClass() }
if ($Mode -eq 'click') {
  $result.click = [N]::Click($X, $Y); Start-Sleep -Milliseconds 400
  $result.focusAfter = [N]::FocusClass()
} else {
  $letters = -join (0x05D0..0x05EA | ForEach-Object { [char]$_ })
  $keys = New-Object System.Collections.ArrayList
  # CPU של כל התהליכים במחשב לפני ואחרי הפרץ: בנייה או חבילת בדיקות של סשן אחר על שתי ליבות
  # מכפילה את עלות ההקשה, ובלי השורה הזאת המדידה נראית תקינה. מודדים לפי TotalProcessorTime.
  function CpuSnap { $h = @{}; Get-Process -ErrorAction SilentlyContinue | ForEach-Object { try { $h[$_.Id] = @($_.Name, $_.TotalProcessorTime.TotalMilliseconds) } catch {} }; return $h }
  $cpu0 = CpuSnap; $cpuT0 = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  for ($i = 0; $i -lt $Count; $i++) {
    $c = $letters[$i % $letters.Length]
    $base = $null
    if ($RW -gt 0) { $base = Grab; Start-Sleep -Milliseconds 25; $b2 = Grab; if ([N]::Diff($base, $b2) -gt 0) { $base.Dispose(); $base = $b2 } else { $b2.Dispose() } }
    $t0 = $sw.Elapsed.TotalMilliseconds
    $epoch = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    $sent = [N]::SendChar($c)
    $seen = -1; $d = 0
    if ($base -ne $null) {
      while (($sw.Elapsed.TotalMilliseconds - $t0) -lt 4000) { $cur = Grab; $d = [N]::Diff($base, $cur); $cur.Dispose(); if ($d -ge 60) { $seen = [Math]::Round($sw.Elapsed.TotalMilliseconds - $t0); break } }
      $base.Dispose()
    }
    [void]$keys.Add(@{ i = $i; ch = [string]$c; t = $epoch; sent = $sent; screenMs = $seen; diff = $d })
    Start-Sleep -Milliseconds $GapMs
  }
  $result.keys = $keys
  $result.focusAfter = [N]::FocusClass()
  $cpu1 = CpuSnap; $cpuDt = [Math]::Max(1, [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() - $cpuT0)
  $mine = @{}; $others = @{}
  foreach ($id in $cpu1.Keys) {
    if (-not $cpu0.ContainsKey($id)) { continue }
    $nm = $cpu1[$id][0]; $pct = 100 * ($cpu1[$id][1] - $cpu0[$id][1]) / $cpuDt
    if ($pct -le 0) { continue }
    if ($nm -eq 'otzaria' -or $nm -eq 'msedgewebview2' -or $id -eq $PID) { $mine[$nm] = [double]($mine[$nm]) + $pct } else { $others[$nm] = [double]($others[$nm]) + $pct }
  }
  $top = $others.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 3 | ForEach-Object { @{ name = $_.Key; pct = [Math]::Round($_.Value) } }
  $result.cpu = @{ logical = [Environment]::ProcessorCount; measured = [Math]::Round(($mine.Values | Measure-Object -Sum).Sum); others = [Math]::Round(($others.Values | Measure-Object -Sum).Sum); top = @($top) }
}
$result | ConvertTo-Json -Compress -Depth 4 | Out-File -Encoding utf8 $Out
`;

const tmp = mkdtempSync(join(tmpdir(), 'owe-inapp-'));
const psPath = join(tmp, 'inapp.ps1');
writeFileSync(psPath, '﻿' + PS, 'utf8');
const runPs = (params) => {
  const out = join(tmp, `out-${Date.now()}.json`);
  const argv = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', psPath, '-Out', out];
  for (const [k, v] of Object.entries(params)) argv.push(`-${k}`, String(v));
  const r = spawnSync('powershell', argv, { encoding: 'utf8', timeout: 600_000 });
  if (r.status !== 0) throw new Error('PowerShell נכשל: ' + (r.stderr || r.stdout).trim().slice(0, 600));
  return JSON.parse(readFileSync(out, 'utf8').replace(/^﻿/, ''));
};
let result = null;
try {
  const hwnd = HWND;
  if (KEYS === 0) {
    console.log('--keys 0: מכשור הותקן, אין קלט. (hwnd ' + hwnd + ')');
  } else {
    if (CLICK) {
      // שורת טקסט גלויה, בתוך החלון (בחלון צר הדף רחב מהחלון והשורות חורגות ממנו).
      // שורה בלי קישורים: לחיצה על קישור במסמך פותחת ספר באוצריא ומחליפה לשונית.
      // סיווג השורה לפי רוחבה מול העמוד שלה: wide = מעל 60% מרוחב העמוד; אחרת טור,
      // ימני או שמאלי לפי מרכז השורה מול מרכז העמוד.
      const pt = JSON.parse(
        await js(`(async function(target){var W=innerWidth,H=innerHeight;var all=[...document.querySelectorAll('.superdoc-line')].filter(function(l){return !l.querySelector('a,[href],[data-href],[data-link],[class*=link]');});var info=function(l){var r=l.getBoundingClientRect();var pg=l.closest('[data-page-index]');var pr=pg?pg.getBoundingClientRect():{left:0,right:W};var kind=r.width>(pr.right-pr.left)*0.6?'wide':((r.left+r.right)/2<(pr.left+pr.right)/2?'left':'right');var full=(l.textContent||'').trim();return {el:l,r:r,kind:kind,full:full,txt:full.slice(0,40)};};var vis=function(o){var r=o.r;return r.height>8&&r.top>120&&r.bottom<H-40&&Math.min(r.right,W)-Math.max(r.left,0)>80;};var match=function(o){return target==='auto'||(target.indexOf('text:')===0?o.full.indexOf(target.slice(5))>=0:o.kind===target);};var os=all.map(info);var cands=os.filter(match);var seen=cands.filter(vis);var pick=target==='auto'?seen[Math.min(seen.length-1,2)]:seen[0];if(!pick&&cands.length){cands[0].el.scrollIntoView({block:'center'});await new Promise(function(r){setTimeout(r,700)});var o=info(cands[0].el);if(vis(o))pick=o;}if(!pick)return JSON.stringify({error:cands.length?'not-visible':'no-match',kinds:os.reduce(function(m,o){m[o.kind]=(m[o.kind]||0)+1;return m;},{})});var r=pick.r;var left=Math.max(r.left,0),right=Math.min(r.right,W);return JSON.stringify({x:Math.round((left+right)/2),y:Math.round(r.top+r.height/2),kind:pick.kind,w:Math.round(r.width),txt:pick.txt});})(${JSON.stringify(TARGET)})`),
      );
      if (!pt || pt.error) throw new Error(pt && pt.error === 'no-match' ? `אין שורה מסוג „${TARGET}” במסמך (יש: ${JSON.stringify(pt.kinds)})` : 'לא נמצאה שורת טקסט גלויה ללחיצה');
      console.log(`שורת היעד: ${pt.kind}, רוחב ${pt.w}px, „${pt.txt}”`);
      const tabBefore = state.tab;
      // מיקום החלון נקרא מחדש: אוצריא עשויה להזיז את החלון (שחזור מיקום) גם שניות אחרי הפתיחה.
      state.screen = JSON.parse(await js('JSON.stringify([screenX, screenY])'));
      const c = runPs({ Mode: 'click', Hwnd: hwnd, X: state.screen[0] + pt.x, Y: state.screen[1] + pt.y, Count: 0, GapMs: 0, RX: 0, RY: 0, RW: 0, RH: 0 });
      console.log(`לחיצה ב-(${state.screen[0] + pt.x},${state.screen[1] + pt.y}): ${c.click}; פוקוס המקלדת: ${c.focusBefore} → ${c.focusAfter}`);
      // אחרי הלחיצה: העורך חייב להיות קיים ובאותה לשונית (החלפת לשונית מאפסת את __otzariaEditor לרגע).
      let settled = false;
      for (let i = 0; i < 20 && !settled; i++) {
        await sleep(150);
        settled = await js(`!!(window.__otzariaEditor && window.__otzariaEditor.container) && ((document.querySelector('.word-doctab.active') || {}).textContent || '').trim() === ${JSON.stringify(tabBefore)}`);
      }
      if (!settled) {
        // הלחיצה החליפה לשונית (או שהעורך נעלם לרגע): חוזרים ללשונית המקורית דרך הדף. פוקוס
        // המקלדת של המערכת כבר על חלון ה-WebView אחרי הלחיצה, ולכן די בפוקוס JS למשטח ההקלדה.
        console.log('אחרי הלחיצה הלשונית/העורך לא במקום — חוזרים ללשונית „' + tabBefore + '”');
        await js(`(function(){var t=Array.prototype.find.call(document.querySelectorAll('.word-doctab'),function(x){return (x.textContent||'').trim()===${JSON.stringify(tabBefore)}}); if(t) t.click(); return !!t;})()`);
        for (let i = 0; i < 30 && !settled; i++) {
          await sleep(200);
          settled = await js(`!!(window.__otzariaEditor && window.__otzariaEditor.container)`);
        }
        if (!settled) throw new Error('העורך לא חזר אחרי הלחיצה');
        await js(`(function(){var c=window.__otzariaEditor.container;var ta=c.querySelector('textarea, [contenteditable="true"]'); if(ta) ta.focus(); return document.activeElement&&document.activeElement.tagName;})()`);
      }
      await sleep(250);
    }
    let region = { RX: 0, RY: 0, RW: 0, RH: 0 };
    if (SCREEN) {
      const caret = JSON.parse(
        await js(`(function(){var c=window.__otzariaEditor.container;var e=c.querySelector('.sd-v2-local-selection-caret')||c.querySelector('textarea');if(!e)return 'null';var r=e.getBoundingClientRect();return JSON.stringify({x:Math.round(r.x),y:Math.round(r.y),h:Math.round(r.height)});})()`),
      );
      if (!caret || caret.h <= 0) throw new Error('הסמן לא נמצא — ללחוץ במסמך קודם (בלי --no-click)');
      // שלוש שורות מגובה הסמן, לרוחב החלון: הסמן יורד שורה בזמן ההקלדה.
      state.screen = JSON.parse(await js('JSON.stringify([screenX, screenY])'));
      state.inner = JSON.parse(await js('JSON.stringify([innerWidth, innerHeight])'));
      region = { RX: state.screen[0], RY: state.screen[1] + caret.y - 8, RW: state.inner[0], RH: caret.h * 3 + 16 };
      console.log(`אזור המסך הנדגם: (${region.RX},${region.RY}) ${region.RW}×${region.RH}`);
    }
    const textBefore = await js('window.__otzariaEditor.container.textContent.length');
    const r = runPs({ Mode: 'type', Hwnd: hwnd, X: 0, Y: 0, Count: KEYS, GapMs: GAP_MS, ...region });
    await sleep(2500);
    const O = JSON.parse(await js('JSON.stringify(window.__inapp)'));
    const textAfter = await js('window.__otzariaEditor.container.textContent.length');
    result = { sent: r.keys, O, textBefore, textAfter, focus: `${r.focusBefore} → ${r.focusAfter}`, cpu: r.cpu };
  }
} finally {
  ws.close();
  rmSync(tmp, { recursive: true, force: true });
}
if (!result) process.exit(0);

/* ------------------------------ החשבון ------------------------------ */
const { sent, O } = result;
const stat = (xs) => {
  if (!xs.length) return 'אין';
  const s = [...xs].sort((a, b) => a - b);
  const q = (p) => s[Math.min(s.length - 1, Math.floor(p * s.length))];
  return `n=${s.length} חציון=${Math.round(q(0.5))} p90=${Math.round(q(0.9))} מקס=${Math.round(s[s.length - 1])}`;
};
const failed = sent.filter((k) => k.sent !== 'ok');
console.log(`\nנשלחו ${sent.length} הקשות${failed.length ? ` (${failed.length} נכשלו: ${failed[0].sent})` : ''}; keydown בדף: ${O.keys.length}; טקסט ב-DOM ${result.textBefore} → ${result.textAfter}; פוקוס המקלדת: ${result.focus}`);
if (result.cpu) {
  // עומס זר בזמן הפרץ (בנייה, vitest, Chrome של שער QA מסשן אחר): על שתי ליבות הוא מכפיל את
  // עלות ההקשה, והתוצאה נראית כמו „המארח איטי”. מדידה עם עומס זר ניכר אינה מדידה.
  const c = result.cpu;
  const top = (c.top || []).map((t) => `${t.name} ${t.pct}%`).join(', ');
  const foreign = c.others > 50 * (c.logical / 4);
  console.log(`CPU בזמן הפרץ (אחוז מליבה לוגית אחת, ${c.logical} ליבות לוגיות): אוצריא+WebView2+הכלי ${c.measured}%, תהליכים אחרים ${c.others}%${top ? ` (${top})` : ''}${foreign ? ' — ⚠ עומס זר במחשב: המדידה מזוהמת, להריץ שוב כשהמחשב שקט' : ''}`);
}
if (O.keys.length === 0) {
  console.log('אף הקשה לא הגיעה לדף: פוקוס המקלדת של המערכת אינו על חלון ה-WebView (Chrome_WidgetWin_0). להריץ בלי --no-click, או ללחוץ במסמך ידנית.');
  process.exit(1);
}
const n = Math.min(sent.length, O.keys.length);
// שיוך תו↔הקשה: כשמספר עדכוני הטקסט שווה למספר ההקשות — לפי סדר (המנוע מעבד הקשה
// אחת בכל פעם ואינו מאחד); אחרת כל עדכון משויך להקשה האחרונה שקדמה לו.
const inOrder = O.text.length === O.keys.length;
const textOf = (i) => {
  if (inOrder) return O.text[i] ?? null;
  const k = O.keys[i], next = O.keys[i + 1];
  return O.text.find((x) => x.t > k.down && (!next || x.t <= next.down)) ?? null;
};
const host = [], dom = [], paint = [], total = [], present = [], screen = [];
for (let i = 0; i < n; i++) {
  const s = sent[i], k = O.keys[i], x = textOf(i);
  host.push(k.down - s.t);
  if (x) {
    dom.push(x.t - k.down);
    if (x.painted !== null) {
      paint.push(x.painted - x.t);
      total.push(x.painted - s.t);
      if (s.screenMs >= 0) {
        screen.push(s.screenMs);
        present.push(s.t + s.screenMs - x.painted);
      }
    }
  }
}
console.log(`עדכוני טקסט ב-DOM: ${O.text.length}${inOrder ? '' : ` (≠ ${O.keys.length} הקשות — השיוך לפי „ההקשה האחרונה שלפני”)`}`);
console.log('\n== לכל הקשה (ms) ==');
console.log('מערכת ההפעלה → keydown בדף:   ' + stat(host));
console.log('keydown → התו ב-DOM:           ' + stat(dom) + '  (המנוע: עסקה → worker → ציור ל-DOM)');
console.log('התו ב-DOM → ציור (rAF):        ' + stat(paint));
console.log('הקשה → ציור בדף, יחד:          ' + stat(total));
if (SCREEN) {
  console.log('ציור בדף → פיקסלים על המסך:    ' + stat(present) + '  (המארח: WebView2 → Capture → Flutter)');
  console.log('הקשה → פיקסלים על המסך, יחד:   ' + stat(screen));
  const unseen = sent.filter((s) => s.screenMs < 0).length;
  if (unseen) console.log(`הקשות בלי שינוי נראה באזור הנדגם: ${unseen} (הסמן יצא מהאזור, או שהתו התמזג עם הבא)`);
}
if (O.text.length && O.keys.length) {
  // המדד שהמשתמש מרגיש בהקלדה מהירה: כמה אחרי ההקשה האחרונה הטקסט מתיישב.
  const lastKey = O.keys[O.keys.length - 1].down, lastText = O.text[O.text.length - 1].t;
  console.log(`ההקשה האחרונה → התו האחרון ב-DOM: ${Math.round(lastText - lastKey)}ms (מרווח ההקלדה ${GAP_MS}ms; ערך שגדל בהרבה מעל הקשה בודדת = המנוע לא עומד בקצב, ההקשות נערמות)`);
}
console.log(`תווים שלא נצפו כמצוירים בנפרד: ${O.text.filter((x) => x.painted === null).length}`);
const slow = [];
for (let i = 0; i < n; i++) {
  const s = sent[i], k = O.keys[i], x = textOf(i);
  // התו מ-SendInput (מקש unicode מגיע כ-`Unidentified` ב-e.key).
  if (x && x.t - s.t > 150) slow.push(`#${i} ‚${s.ch || k.key}‘ ${Math.round(x.t - s.t)}`);
}
if (slow.length) console.log('הקשות איטיות (>150ms עד התו ב-DOM): ' + slow.join(', '));
const t0 = sent[0]?.t ?? 0;
const long = O.long.filter((l) => l.start >= t0 - 200);
console.log(`\nlong tasks בזמן ההקלדה: ${long.length}${long.length ? ' — ' + long.map((l) => l.dur + 'ms').join(', ') : ''}`);
if (O.vis.length) console.log('שינויי נראות בזמן הריצה: ' + O.vis.map((v) => v.vis).join(', '));

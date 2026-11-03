# תוכנית בנייה — הגנה מפני קריסה בזמן שמירה אוטומטית

מממשת את **חלק ג׳ בלבד** של `docs/crash-safety-design.md`: שישה שינויים, בסדר
מימוש. מה שחלק ב׳ פסל (דגל סגירה נקייה, צ׳קפוינט סינכרוני, `localStorage`,
קובץ צל) אינו כאן.

## 0. שלושה דברים בחלק ג׳ שהקוד אינו תואם להם

### 0.1 „`await writeDraftNow()` לפני ההרצה” (ג׳5) — לא מספיק כפי שנוסח

`writeDraftNow` פותח ב-`if (!deps.isDirty() || revision === draftedRevision) return;`.
בתרחיש היעד — „שמור” → כלי → autosave — המסמך **נקי** ברגע שהכלי מתחיל,
ולכן הקריאה חוזרת בלי לכתוב דבר. בנוסף `writeDraftNow` אינו חלק מהממשק
`SessionKeeper` ואינו נגיש למעטפת. לכן נדרש **`checkpoint()`** חדש שכותב גם
מסמך נקי — ועל מסמך נקי כותב ישירות ל-`previousDraft`, כי „טיוטה” של מסמך
נקי אינה עבודה לא-שמורה ושלושת הצרכנים של `draft` אסור שיראו אותה.

### 0.2 `flush` בזמן hold היה **דורס** את הצ׳קפוינט

עם משבצות, `draft.path` הוא משבצת קבועה עד השמירה הבאה. `flush` שרץ באמצע
כלי — `plugin.suspended` / `visibilitychange` — היה מייצא מסמך חצי-מוחל
וכותב אותו **לאותה משבצת** שבה הצ׳קפוינט. לכן ה-hold חייב לחסום גם את מסלול
ה-`final`, והצ׳קפוינט הוא היוצא מן הכלל היחיד.

### 0.3 ג׳3 אינו בר-מימוש מעל `markDirty` הקיים

`markDirty` נקרא משלושה מקומות ב-App.vue — עריכה במנוע, שחזור מטיוטה ושינוי
כותרת — ואינו יכול להבחין בין „המסמך שנפתח אינו מה שבדיסק” ל„המשתמש ערך”.
לכן ג׳3 ממומש כמתודה **נפרדת** `markRestored()`.

## 1. סדר מימוש ותלויות

| שלב | מה | קבצים | תלוי ב- |
|---|---|---|---|
| א1 | ג׳1 `hold()` + ג׳3 `markRestored()` | `save-coordinator.ts` | — |
| א2 | ג׳4 `previousDraft` + משבצות | `session-state.ts` | — |
| ב | ג׳4 בזוכר + ג׳2 `isHeld` + ג׳5 `checkpoint()` | `session-keeper.ts` | א2 |
| ג | חיווט: `guardHeavyAction`, שלוש נקודות ה-hold | `App.vue`, `ShulchanTab.vue`, `tools-registration.ts`, `keys.ts` | א1, ב |
| ד | ממשק „גרסה קודמת” | `OpenDocumentDialog.vue`, `App.vue` | ג |
| ה | `npm run typecheck && npx vitest run` | — | הכול |

א1 ו-א2 בלתי תלויים. משלב ג ואילך שלושת כפילי הקואורדינטור בבדיקות הרכיב
חייבים לקבל את החברים החדשים, אחרת המעטפת נופלת ב-`TypeError` בעלייה.

## 2. שינוי 1 — `hold()` ברכז השמירה

```ts
export const HEAVY_ACTION_HOLD_MAX_MS = 5 * 60_000;

hold(options?: { maxMs?: number }): () => void;
readonly isHeld: boolean;
markRestored(): void;
```

מונה, לא דגל: הקפאות מקוננות משתחררות רק עם האחרונה. נפרד לחלוטין
מ-`autosaveEnabled` — המתג הוא העדפה נראית ונשמרת, ושימוש בו כהקפאה היה
**מדליק** אותו בסוף הכלי למי שכיבה אותו.

- מצב: `holdCount`, `restoreRelease`.
- `hold`: מצלם `mine = epoch`; `holdCount += 1`; `cancelAutosave()`; `maxMs`
  → watchdog עם `console.warn`. `release` אידמפוטנטי; אם `mine !== epoch`
  יוצא בלי לגעת במונה; אחרת מפחית, ובאפס — `scheduleAutosave()`.
- `scheduleAutosave`: `|| holdCount > 0` בשער.
- `reset` / `dispose`: `holdCount = 0; restoreRelease = null;`.
- **לא בסנאפשוט** — הוספת שדה הייתה שוברת `toEqual` מדויק בשלושה מקומות.
- **חוזה ה-epoch אינו נפתח:** `release` רק קורא את `epoch` להשוואה.

### בדיקות (fake timers)

עריכה בזמן הקפאה אינה מריצה שמירה · שחרור מתזמן את מה שהמתין ולא את העריכה
הבאה · הקפאה מבטלת סבב ממתין · מקוננות משתחררות עם האחרונה · שחרור כפול
אינו משחרר של אחר · השחרור אינו מדליק מתג כבוי · `reset` מאפס · ידית מלפני
`reset` היא no-op · `dispose` מאפס ואינו מתזמן · הקפאה אינה עוצרת סבב באוויר
(מתעדת את המגבלה) · תקרת זמן משחררת הקפאה תקועה · `isHeld` מדווח נכון.

## 3. שינוי 3 — `markRestored()`

`save.reset(...)` מגדיר יעד ואז `save.markDirty()` בשחזור טיוטה מפעיל
`scheduleAutosave` — הטיוטה נכתבת לקובץ 2.5 שניות אחרי העלייה, **בכל עלייה**.

- `markRestored()`: `restoreRelease?.(); restoreRelease = hold(); dirtyRevision += 1; publish();`
  — בלי `scheduleAutosave`.
- `markDirty()`: משחרר את `restoreRelease` בשורה הראשונה.
- `saveNow()`: משחרר אחרי בדיקת `disposed` ולפני `cancelAutosave()` — הסדר
  חשוב, כי השחרור מתזמן טיימר וה-`cancelAutosave` שאחריו מבטל אותו.
- App.vue: `save?.markDirty()` בשחזור → `save?.markRestored()`.

### בדיקות

מסמך ששוחזר מסומן מלוכלך ואינו נשמר מעצמו · עריכה ראשונה מחדשת · „שמור”
משחרר · שחזור אינו מדליק מתג כבוי · שחזור חוזר (טאב נרדם) אינו מצטבר ·
ברמת הרכיב: המעטפת קוראת `markRestored` ולא `markDirty`.

**סיכון שדורש שער חי:** אם SuperDoc מפעיל `onUpdate` כבר בטעינה, ההקפאה
תשתחרר מיד. לא ניתן לאמת ב-jsdom.

## 4. שינוי 4 — משבצות `-a`/`-b` ו-`previousDraft`

### session-state.ts

```ts
export type DraftSlot = 'a' | 'b';
export function draftPathFor(id: DocumentSessionId, slot: DraftSlot): string;
export function draftSlotPaths(id: DocumentSessionId): readonly [string, string];
export function nextDraftPath(paths, rewrite, keep): string;
```

כללי `nextDraftPath`: (1) אם `rewrite?.path` הוא משבצת ושונה מ-`keep?.path`
→ הוא; (2) אחרת הראשון מ-`paths` ששונה מ-`keep?.path`; (3) אחרת `paths[0]`.
נתיב ישן (`session-draft-<id>.docx`) אינו נדרס — הכתיבה עוברת למשבצת.

`SessionDocumentEntry.previousDraft: SessionDraft | null` — **שדה נפרד**, כי
שלושה צרכנים קוראים „יש טיוטה” כ„יש עבודה לא-שמורה”: `isDirty` של טאב ממתין,
`hasDraft` ב-open-flow, והגיבוי של „לא לשמור” — האחרון היה מגבה את הגרסה
הישנה במקום את מה שעל המסך.

`readDocumentEntry`: `previousDraft: readDraft(entry.previousDraft)` —
מחזיר `null` על חסר או פגום, ולכן רשומה ישנה נקראת בלי לזרוק.
**`SESSION_VERSION` נשאר 2** — תוספת של שדה אופציונלי אינה שינוי שובר,
והעלאה הייתה מוחקת לכל המשתמשים את הטאבים והטיוטות.

`decideDraftRecovery` — ללא שינוי; מקבל `draft` בלבד.

### session-keeper.ts

deps: `writeDraft(content, path)`, `removeDraft(path)`, `draftPaths`, `isHeld`.
`draftPath: string` יורד.

`noteSaved`: במקום `await deps.removeDraft()` —
`state = withActiveEntry(state, { draft: null, previousDraft: entry.draft })`.
אפס בייטים בגשר.

`discardDraft`: מוחק את שתי המשבצות ואת נתיב הרשומה, כשיש מה למחוק.

`writeDraftNow(mode: 'timer' | 'final' | 'checkpoint')` — היעד:
`checkpoint` על מסמך נקי → `previousDraft`; אחרת `draft`. הנתיב מ-`nextDraftPath`.
`draftedRevision` מתקדם רק על `draft`.

### תאימות לאחור

רשומה ישנה: `previousDraft: null`, `draft` כפי שהוא, הכתיבה הבאה עוברת
למשבצת `-a`, `discardDraft` מוחק את הישן ואת שתי המשבצות. קובץ יתום אחד
לטאב שהיה קיים ברגע העדכון — חד-פעמי.

## 5. שינוי 2 — ההקפאה שולטת גם בטיוטה

ב-`writeDraftNow`, אחרי השער של `isSaving`:

```ts
if (mode !== 'checkpoint' && deps.isHeld()) {
  if (mode === 'timer') { draftDeadline = Date.now() + DRAFT_DELAY_MS; scheduleDraft(); }
  return;
}
```

`final` (flush) יוצא **בלי לכתוב** — הטיוטה שבמשבצת היא הצ׳קפוינט שלפני
הכלי, וייצוא של מסמך חצי-מוחל היה דורס אותו (0.2).

## 6. שינוי 5 — `checkpoint()`

```ts
checkpoint(): Promise<void>;
```

כותבת עותק של המסמך כפי שהוא, **גם כשהוא נקי** — זה ההבדל מ-`flush`.
מסמך מלוכלך → `draft`; מסמך נקי → `previousDraft` ישירות (אילו נרשם
כ-`draft` היה נפתח בעלייה כ„שינויים שלא נשמרו” שאינם קיימים). מתעלמת
מ-`isHeld` (הקורא מקפיא לפניה בכוונה), ממתינה לשמירה שרצה, וכשל ייצוא נרשם
ללוג ואינו מפיל את הפעולה.

העלות: ייצוא DOCX מלא אחד לכל הרצת כלי. לא רציפה, ולא על הקלדה.

## 7. שינוי 6 — שלוש נקודות ה-hold

`keys.ts`: `HEAVY_ACTION_GUARD` — `InjectionKey<<T>(action: () => Promise<T>) => Promise<T>>`,
ברירת מחדל: הפעולה כמות שהיא.

App.vue:

```ts
async function guardHeavyAction<T>(session, action): Promise<T> {
  const release = session.save.hold({ maxMs: HEAVY_ACTION_HOLD_MAX_MS });
  try {
    await session.keeper.checkpoint();
    return await action();
  } finally { release(); }
}
```

הסדר: hold **לפני** הצ׳קפוינט, כדי שסבב autosave לא יתחיל תוך כדי הייצוא.

שלוש הנקודות: `onReplaceAllText` (App.vue), `runTool` (ShulchanTab.vue, דרך
`inject`), ו-`registerShulchanTools` (פרמטר `guard` אופציונלי שעוטף כל `run`).

## 8. ממשק „גרסה קודמת”

כפתור שני ב-`OpenDocumentDialog.vue`, מוצג רק כשיש `previousDraft` ששייך
למסמך הפתוח. נפתח בטאב חדש כמסמך **בלי יעד כתיבה** — „שמור” פותח „שמור
בשם”, ולכן אינו יכול לדרוס דבר.

## 9. מלכודות

### 9.1 jsdom אינו מממש `Blob.prototype.arrayBuffer`

מסלול הטיוטה קורא לו. בבדיקות רכיב, `createSuperdocDouble` אינו מגדיר
`host.export`, ולכן הייצוא זורק, `writeDraftNow` בולע ו-`checkpoint` נפתר
בשקט — **בדיקה שרק מוודאת „לא נזרק” עוברת על קוד שבור**. הכלל: להזריק
`host.export` שמחזיר Blob עם `arrayBuffer`; לטעון חיובית על
`stub.workspaceWrites`; ולהריץ פעם אחת עם המוטציה כדי לראות אדום.

### 9.2 שלושת כפילי הקואורדינטור

`app-shell.test.ts`, `shortcuts-core.test.ts`, `shortcuts-help.test.ts`
מחזירים אובייקט ידני. בלי `hold`/`isHeld`/`markRestored` זה `TypeError`
שמפיל את **כל** בדיקות המעטפת. הטיפוס אינו מגן — המפעל מוחזר לא-מוקלד.

### 9.3 `markDirty` בזמן hold — המצב נשמר, התזמון אובד

`scheduleAutosave` נקרא משני מקומות בלבד. בלי `scheduleAutosave()` ב-`release`,
כלי שהיה הפעולה האחרונה לא נכתב לעולם.

### 9.4 hold שלא משוחרר

`maxMs = 5 דקות`. לולאה סינכרונית אינה הבעיה (כשהיא נגמרת ה-`finally` רץ);
הבעיה היא הבטחה שאינה נפתרת. גם `reset` מרפא, ו-`saveNow` עובד תמיד בלי קשר
ל-hold. הקפאת השחזור בלי תקרה, בכוונה.

## 10. אימות ברמת החבילה

`npm run typecheck && npx vitest run`. מה נשבר בכוונה: ה-harness של
session-keeper (`draftPath` → `draftPaths`, חתימות `writeDraft`/`removeDraft`),
בדיקת „שמירה מוצלחת מוחקת” שהחוזה שלה השתנה, `toEqual` מדויק ב-session-state,
ושלושת כפילי הקואורדינטור.

## 11. מה לא נעשה

אין שינוי ב-`runOnce`/`saveLoop`/`runChain`/`epoch`. אין `move` (אינו ב-SDK).
אין base64 נוסף על מסלול חם. `autosaveEnabled` אינו נקרא ואינו נכתב על ידי
שום שינוי כאן. `SESSION_VERSION` נשאר 2.

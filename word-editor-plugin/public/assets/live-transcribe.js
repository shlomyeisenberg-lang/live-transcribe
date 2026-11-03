/* Live transcription for "Word for Otzaria".
 *
 * Otzaria does not let plugins use the microphone, so listening and speech
 * recognition happen in the Live Transcribe desktop app, which serves a
 * small local API on 127.0.0.1:47321 (the same API the standalone "תמלול חי"
 * plugin talks to). This script is a front end for it inside the Word editor:
 *
 *  - a microphone button in the title bar opens a docked panel;
 *  - the panel starts/stops live dictation or transcribes an audio file;
 *  - every finished sentence is typed into the open document at the cursor
 *    ("dictate into the document"), or can be inserted line by line;
 *  - each line can be searched in the Otzaria library or copied.
 *
 * Self-contained and outside the compiled bundle, like english-mode.js. It
 * reaches the editor only through window.__otzariaEditors / __otzariaEditor,
 * and inserts with the same engine call "Send to document" uses
 * (activeEditor.doc.insert at the current selection).
 *
 * Classic script on purpose: WebView2 cannot load ES modules from file://.
 */
(function () {
  'use strict';

  var API = 'http://127.0.0.1:47321';
  var AUDIO_EXT = ['mp3', 'wav', 'm4a', 'ogg', 'opus', 'oga', 'flac', 'aac', 'wma', 'webm', 'mp4', 'amr', '3gp'];
  var PREF_DICTATE = 'otzaria-word.transcribe.dictate';
  var PREF_NEWLINE = 'otzaria-word.transcribe.newParagraph';
  var OV_API = 'http://127.0.0.1:47811';
  var PREF_HE_MODEL = 'otzaria-word.transcribe.heModel';
  var PREF_TS = 'otzaria-word.transcribe.timestamps';

  // Languages and models offered by Live Transcribe 1.2+ (keep in sync
  // with the app's web/engines/engines.js).
  var W = ['whistle', 'parakeet-redux'], P = ['parakeet-redux'];
  var LANGS = [
    { code: 'he', he: 'עברית', en: 'Hebrew', models: ['gigaam-he', 'zipformer-he', 'otzaria-voice'] },
    { code: 'auto', he: 'זיהוי אוטומטי', en: 'Auto-detect', models: W },
    { code: 'en', he: 'אנגלית', en: 'English', models: W },
    { code: 'de', he: 'גרמנית', en: 'German', models: W },
    { code: 'fr', he: 'צרפתית', en: 'French', models: W },
    { code: 'es', he: 'ספרדית', en: 'Spanish', models: W },
    { code: 'it', he: 'איטלקית', en: 'Italian', models: W },
    { code: 'nl', he: 'הולנדית', en: 'Dutch', models: W },
    { code: 'pl', he: 'פולנית', en: 'Polish', models: W },
    { code: 'pt', he: 'פורטוגזית', en: 'Portuguese', models: P },
    { code: 'ru', he: 'רוסית', en: 'Russian', models: P },
    { code: 'uk', he: 'אוקראינית', en: 'Ukrainian', models: P },
    { code: 'cs', he: 'צ׳כית', en: 'Czech', models: P },
    { code: 'sk', he: 'סלובקית', en: 'Slovak', models: P },
    { code: 'hu', he: 'הונגרית', en: 'Hungarian', models: P },
    { code: 'ro', he: 'רומנית', en: 'Romanian', models: P },
    { code: 'bg', he: 'בולגרית', en: 'Bulgarian', models: P },
    { code: 'el', he: 'יוונית', en: 'Greek', models: P },
    { code: 'hr', he: 'קרואטית', en: 'Croatian', models: P },
    { code: 'sl', he: 'סלובנית', en: 'Slovenian', models: P },
    { code: 'sv', he: 'שוודית', en: 'Swedish', models: P },
    { code: 'da', he: 'דנית', en: 'Danish', models: P },
    { code: 'fi', he: 'פינית', en: 'Finnish', models: P },
    { code: 'et', he: 'אסטונית', en: 'Estonian', models: P },
    { code: 'lv', he: 'לטבית', en: 'Latvian', models: P },
    { code: 'lt', he: 'ליטאית', en: 'Lithuanian', models: P },
    { code: 'mt', he: 'מלטית', en: 'Maltese', models: P },
  ];
  var MODEL_NAMES = {
    'zipformer-he': { he: 'Zipformer (עברית)', en: 'Zipformer (Hebrew)' },
    'gigaam-he': { he: 'GigaAM עברית + מילון הספרייה', en: 'GigaAM Hebrew + library words' },
    'otzaria-voice': { he: 'שירות הדיבור של אוצריא (GigaAM)', en: 'Otzaria Voice service (GigaAM)' },
    whistle: { he: 'Whistle', en: 'Whistle' },
    'parakeet-redux': { he: 'Parakeet Redux', en: 'Parakeet Redux' },
  };
  function langInfo(code) {
    for (var i = 0; i < LANGS.length; i++) if (LANGS[i].code === code) return LANGS[i];
    return LANGS[0];
  }


  /* ------------------------------------------------------------ strings */

  var STR = {
    he: {
      button: 'תמלול',
      buttonTip: 'תמלול חי — הכתבה בעברית לתוך המסמך (Alt+V)',
      title: 'תמלול חי',
      close: 'סגירה',
      connecting: 'מתחבר…',
      offline: 'האפליקציה אינה פועלת',
      starting: 'האפליקציה עולה…',
      recording: 'מקליט',
      fileBusy: 'מתמלל קובץ',
      loadingModel: 'טוען מודל…',
      connected: 'מחובר',
      offlineTitle: 'אפליקציית התמלול אינה פועלת',
      offlineBody: 'אוצריא אינה מאפשרת לתוספים גישה למיקרופון, ולכן התמלול עצמו נעשה באפליקציה Live Transcribe שרצה ברקע. הפעילו את LiveTranscribe.exe פעם אחת — מגרסה 1.3 היא עולה מעצמה עם Windows ונשארת ברקע גם כשסוגרים את חלונה, והעורך יתחבר אליה מעצמו.',
      noSdk: 'התמלול זמין רק כשהעורך פועל בתוך אוצריא.',
      retry: 'נסו שוב',
      startRec: 'התחלת הכתבה',
      stopRec: 'מקליט… לחצו לעצירה',
      liveSub: 'דברו בעברית — כל משפט שהסתיים נכתב במסמך במקום הסמן.',
      liveSubOff: 'דברו בעברית — הטקסט יופיע כאן, וממנו אפשר להכניס למסמך.',
      pickFile: 'תמלול קובץ שמע…',
      pickTitle: 'בחירת קובץ שמע לתמלול',
      cancel: 'ביטול',
      dictate: 'כתיבה ישירה למסמך',
      dictateTip: 'כשפעיל, כל משפט שתומלל נכתב מיד במסמך במקום הסמן',
      newPara: 'פסקה לכל משפט',
      newParaTip: 'כל משפט מתחיל פסקה חדשה במקום להמשיך את השורה',
      empty: 'התמלול יופיע כאן. כל שורה אפשר להכניס למסמך, לחפש באוצריא או להעתיק.',
      fromFile: 'מתוך הקובץ: ',
      recordingName: 'הקלטה',
      insertLine: 'הכנסה למסמך במקום הסמן',
      searchLine: 'חיפוש באוצריא',
      copyLine: 'העתקת השורה',
      inserted: 'נכתב במסמך',
      insertAll: 'הכנסת הכול למסמך',
      copyAll: 'העתקת הכול',
      clear: 'ניקוי',
      clearTitle: 'ניקוי התמלול',
      clearBody: 'למחוק את כל הטקסט שתומלל? הפעולה מנקה גם את חלון האפליקציה. מה שכבר נכתב במסמך נשאר.',
      stopFirst: 'עצרו את התמלול לפני הניקוי',
      copied: 'הועתק',
      nothing: 'אין עדיין טקסט',
      noDoc: 'אין מסמך פתוח שאפשר לכתוב בו',
      insertFailed: 'הכתיבה למסמך נכשלה',
      noConn: 'אין חיבור לאפליקציית התמלול',
      notReady: 'אפליקציית התמלול עדיין עולה — נסו שוב בעוד רגע',
      busy: 'כבר מתבצע תמלול — עצרו אותו קודם',
      badUrl: 'לא ניתן להעביר את הקובץ לאפליקציה',
      failed: 'הפעולה נכשלה',
      fileOpenFailed: 'לא ניתן לפתוח את הקובץ: ',
      searchFailed: 'פתיחת החיפוש באוצריא נכשלה',
      noPermission: 'לתוסף אין הרשאת רשת מקומית. אפשר להעניק אותה בהגדרות התוספים של אוצריא.',
      fileTitle: 'תמלול הקלטה',
      file: 'קובץ',
      jobDownloading: 'מעביר את "%s" לאפליקציה…',
      jobQueued: 'ממתין לתמלול "%s"…',
      jobDecoding: 'קורא את "%s"…',
      jobTranscribing: 'מתמלל את "%s" — %p%',
      jobDone: 'התמלול של "%s" הסתיים',
      jobCancelled: 'התמלול של "%s" בוטל',
      jobError: 'שגיאה בתמלול "%s": ',
      stCatchUp: 'מקליט… (משלים פיגור של %n שניות)',
      stRecording: 'מקליט… דברו בעברית',
      stDownloading: 'הפעלה ראשונה: מוריד את המודל… %n%',
      stLoading: 'טוען את מודל הזיהוי…',
      stNoMic: 'המיקרופון אינו זמין — בדקו את הגדרות הפרטיות של Windows',
      stModelFailed: 'טעינת המודל נכשלה — פתחו את חלון האפליקציה לפרטים',
      timestamps: 'זמנים',
      timestampsTip: 'הצגת זמן ליד כל שורה, וגם בטקסט שנכתב למסמך ובהעתקה',
      language: 'שפה',
      model: 'מודל',
      builtIn: 'מובנה באפליקציה',
      installed: 'מותקן',
      downloads: 'יורד בשימוש הראשון (כ-%n MB)',
      autoLang: 'מזהה את השפה בעצמו',
      speakHe: 'דברו בעברית',
      speakAny: 'דברו',
      speakIn: 'דברו ב%s',
      liveSubTail: ' — כל משפט שהסתיים נכתב במסמך במקום הסמן.',
      liveSubTailOff: ' — הטקסט יופיע כאן, וממנו אפשר להכניס למסמך.',
      editTip: 'לחצו כדי לתקן את השורה',
      editFailed: 'אין חיבור לאפליקציה — התיקון לא נשמר',
      showApp: 'פתיחת חלון האפליקציה',
      pickerOffline: 'הבחירה תתאפשר כשאפליקציית התמלול תפעל',
      pickerOld: 'כדי לבחור שפה ומודל צריך את אפליקציית Live Transcribe בגרסה 1.3 ומעלה',
      ovConnected: 'מחובר (שירות הדיבור)',
      ovListening: 'מקשיב — דברו בעברית',
      ovReady: 'מודל GigaAM עם מילון הספרייה, דרך שירות הדיבור של אוצריא — בלי אינטרנט',
      ovMissing: 'שירות הדיבור של אוצריא אינו פועל. אפשר להתקין אותו מ-github.com/yakov100/otzaria-voice, או לבחור מודל של Live Transcribe',
      ovNoPack: 'חבילת הקול (כ-330MB) עדיין לא מותקנת — לחיצה על ההכתבה תוריד אותה',
      ovPackBusy: 'מוריד את חבילת הקול של שירות הדיבור… %p%',
      ovNoRuntime: 'שירות הדיבור לא מצא את ONNX Runtime של אוצריא — ודאו שאוצריא מותקנת כרגיל',
      ovNoFile: 'תמלול קבצים עובד עם Live Transcribe בלבד',
      ovFailed: 'זיהוי הדיבור נכשל: ',
      ovBusy: 'שירות הדיבור כבר מקליט בחלון אחר',
      ovNoMic: 'שירות הדיבור אינו יכול לגשת למיקרופון',
      ovOfflineBody: 'שירות הדיבור של אוצריא אינו פועל. הוא מתקין את עצמו פעם אחת ועולה עם Windows. אפשר גם להפעיל את Live Transcribe ולבחור מודל שלה.',
    },
    en: {
      button: 'Dictate',
      buttonTip: 'Live transcription — dictate Hebrew into the document (Alt+V)',
      title: 'Live Transcription',
      close: 'Close',
      connecting: 'Connecting…',
      offline: 'App not running',
      starting: 'App starting…',
      recording: 'Recording',
      fileBusy: 'Transcribing file',
      loadingModel: 'Loading model…',
      connected: 'Connected',
      offlineTitle: 'The transcription app is not running',
      offlineBody: 'Otzaria does not give plugins access to the microphone, so transcription runs in the Live Transcribe app in the background. Start LiveTranscribe.exe once — from version 1.3 it starts with Windows and keeps running in the background when its window is closed — and the editor will connect by itself.',
      noSdk: 'Transcription is only available when the editor runs inside Otzaria.',
      retry: 'Try again',
      startRec: 'Start dictation',
      stopRec: 'Recording… click to stop',
      liveSub: 'Speak Hebrew — each finished sentence is typed into the document at the cursor.',
      liveSubOff: 'Speak Hebrew — the text appears here and you can insert it into the document.',
      pickFile: 'Transcribe audio file…',
      pickTitle: 'Choose an audio file to transcribe',
      cancel: 'Cancel',
      dictate: 'Write into document',
      dictateTip: 'When on, every transcribed sentence is typed into the document at the cursor right away',
      newPara: 'New paragraph per sentence',
      newParaTip: 'Start each sentence on a new paragraph instead of continuing the line',
      empty: 'The transcript will appear here. Each line can be inserted into the document, searched in Otzaria, or copied.',
      fromFile: 'From file: ',
      recordingName: 'recording',
      insertLine: 'Insert into the document at the cursor',
      searchLine: 'Search in Otzaria',
      copyLine: 'Copy line',
      inserted: 'Written to the document',
      insertAll: 'Insert all into document',
      copyAll: 'Copy all',
      clear: 'Clear',
      clearTitle: 'Clear transcript',
      clearBody: 'Delete all transcribed text? This also clears the app window. Text already written to the document stays.',
      stopFirst: 'Stop transcribing before clearing',
      copied: 'Copied',
      nothing: 'No text yet',
      noDoc: 'No open document to write into',
      insertFailed: 'Writing to the document failed',
      noConn: 'No connection to the transcription app',
      notReady: 'The transcription app is still starting — try again in a moment',
      busy: 'A transcription is already running — stop it first',
      badUrl: 'Could not hand the file to the app',
      failed: 'The action failed',
      fileOpenFailed: 'Could not open the file: ',
      searchFailed: 'Could not open the Otzaria search',
      noPermission: 'The plugin lacks local network permission. You can grant it in Otzaria plugin settings.',
      fileTitle: 'Transcribe a recording',
      file: 'file',
      jobDownloading: 'Handing "%s" to the app…',
      jobQueued: 'Waiting to transcribe "%s"…',
      jobDecoding: 'Reading "%s"…',
      jobTranscribing: 'Transcribing "%s" — %p%',
      jobDone: 'Finished transcribing "%s"',
      jobCancelled: 'Transcription of "%s" cancelled',
      jobError: 'Error transcribing "%s": ',
      stCatchUp: 'Recording… (catching up, %ns behind)',
      stRecording: 'Recording… speak Hebrew',
      stDownloading: 'First run: downloading the model… %n%',
      stLoading: 'Loading the recognition model…',
      stNoMic: 'Microphone unavailable — check Windows privacy settings',
      stModelFailed: 'Could not load the model — open the app window for details',
      timestamps: 'Timestamps',
      timestampsTip: 'Show the time next to each line, and include it in text written to the document and in copies',
      language: 'Language',
      model: 'Model',
      builtIn: 'built into the app',
      installed: 'installed',
      downloads: 'downloads on first use (~%n MB)',
      autoLang: 'detects the language by itself',
      speakHe: 'Speak Hebrew',
      speakAny: 'Speak',
      speakIn: 'Speak %s',
      liveSubTail: ' — each finished sentence is typed into the document at the cursor.',
      liveSubTailOff: ' — the text appears here and you can insert it into the document.',
      editTip: 'Click to correct this line',
      editFailed: 'No connection to the app — the correction was not saved',
      showApp: 'Open the app window',
      pickerOffline: 'You can choose once the transcription app is running',
      pickerOld: 'Choosing a language and model needs Live Transcribe 1.3 or later',
      ovConnected: 'Connected (Otzaria Voice)',
      ovListening: 'Listening — speak Hebrew',
      ovReady: 'GigaAM model with the library word list, via the Otzaria Voice service — no internet',
      ovMissing: 'The Otzaria Voice service is not running. Install it from github.com/yakov100/otzaria-voice, or choose a Live Transcribe model',
      ovNoPack: 'The voice pack (~330 MB) is not installed yet — starting dictation downloads it',
      ovPackBusy: 'Downloading the Otzaria Voice pack… %p%',
      ovNoRuntime: "The voice service could not find Otzaria's ONNX Runtime — make sure Otzaria is installed normally",
      ovNoFile: 'File transcription works with Live Transcribe only',
      ovFailed: 'Speech recognition failed: ',
      ovBusy: 'The voice service is already recording in another window',
      ovNoMic: 'The voice service cannot access the microphone',
      ovOfflineBody: 'The Otzaria Voice service is not running. It installs once and starts with Windows. You can also start Live Transcribe and choose one of its models.',
    },
  };

  function lang() {
    return document.documentElement.getAttribute('data-ui-language') === 'en' ? 'en' : 'he';
  }
  function T(key) {
    return STR[lang()][key] || STR.he[key] || key;
  }

  /* ------------------------------------------------------------ state */

  var S = {
    started: false,     // sync loop has been started at least once
    paused: false,      // Otzaria suspended the plugin tab
    loopRunning: false,
    panelOpen: false,
    connected: false,
    everConnected: false,
    pageReady: false,
    seq: 0,
    engine: 'idle',
    recording: false,
    busy: '',
    status: '',
    job: null,
    lines: [],          // {id, text, at, offset, source, job}
    lineEls: {},        // id -> row element
    known: {},          // id -> true for every line seen (never auto-insert twice)
    inserted: {},       // id -> true once written into the document
    jobNames: {},
    lastJob: null,
    liveEl: null,
    dictate: readBool(PREF_DICTATE, true),
    newPara: readBool(PREF_NEWLINE, false),
    showTs: readBool(PREF_TS, false),
    features: {},
    models: {},
    model: '',
    language: '',
    helloAt: 0,
    // Otzaria Voice (GigaAM-He) — the second engine, on 127.0.0.1:47811.
    heModel: readStr(PREF_HE_MODEL, ''), // '' = automatic, 'otzaria-voice', or an app model ('gigaam-he', 'zipformer-he')
    ov: null,           // last /status of the service, or null when unreachable
    ovChecked: false,
    ovLive: false,
    ovListening: false,
    ovN: 0,
  };

  function readStr(key, dflt) {
    try {
      var v = localStorage.getItem(key);
      return v === null ? dflt : v;
    } catch (e) {
      return dflt;
    }
  }
  function writeStr(key, v) {
    try {
      localStorage.setItem(key, v);
    } catch (e) {
      /* convenience only */
    }
  }

  function readBool(key, dflt) {
    try {
      var v = localStorage.getItem(key);
      return v === null ? dflt : v === '1';
    } catch (e) {
      return dflt;
    }
  }
  function writeBool(key, v) {
    try {
      localStorage.setItem(key, v ? '1' : '0');
    } catch (e) {
      /* storage is a convenience */
    }
  }

  var $ = function (id) {
    return document.getElementById(id);
  };
  function sdk() {
    return window.Otzaria || null;
  }
  function sleep(ms) {
    return new Promise(function (r) {
      setTimeout(r, ms);
    });
  }
  function pad(n, w) {
    n = String(n);
    while (n.length < w) n = '0' + n;
    return n;
  }
  function fmtOffset(sec) {
    sec = Math.max(0, Math.floor(sec || 0));
    var h = Math.floor(sec / 3600),
      m = Math.floor((sec % 3600) / 60),
      s = sec % 60;
    return h ? h + ':' + pad(m, 2) + ':' + pad(s, 2) : m + ':' + pad(s, 2);
  }
  function fmtClock(iso) {
    var d = new Date(iso);
    if (isNaN(d)) return '';
    return pad(d.getHours(), 2) + ':' + pad(d.getMinutes(), 2);
  }
  function lineLabel(l) {
    return l.source === 'file' ? fmtOffset(l.offset) : fmtClock(l.at);
  }
  // The text of a line as it goes into the document or a copy.
  function lineOut(l) {
    return (S.showTs ? '[' + lineLabel(l) + '] ' : '') + l.text;
  }
  function langName(code) {
    var li = langInfo(code);
    return lang() === 'en' ? li.en : li.he;
  }

  var toastTimer = null;
  function toast(msg) {
    var t = $('otz-hlt-toast');
    if (!t) return;
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      t.classList.remove('show');
    }, 2600);
  }

  /* ------------------------------------------------------------ desktop app API */

  // One HTTP request through Otzaria's network bridge. Returns {status, json, text}.
  async function api(path, opts) {
    opts = opts || {};
    var O = sdk();
    if (!O) throw new Error('no-sdk');
    var req = {
      url: API + path,
      method: opts.method || 'GET',
      headers: { 'X-HLT-Client': 'otzaria-word-plugin', 'Content-Type': 'application/json' },
      timeoutMs: opts.timeoutMs || 15000,
    };
    if (opts.body !== undefined) req.body = JSON.stringify(opts.body);
    var status = 0,
      text = '';
    var chunks = O.call('network.fetchStream', req);
    // Some SDK builds hand back a promise of {success:false} instead of a stream.
    if (chunks && typeof chunks.then === 'function' && !chunks[Symbol.asyncIterator]) {
      var maybe = await chunks;
      if (maybe && maybe.success === false) {
        var err = new Error((maybe.error && maybe.error.message) || 'network');
        err.code = maybe.error && maybe.error.code;
        throw err;
      }
      chunks = maybe;
    }
    for await (var c of chunks) {
      if (c.type === 'response') status = c.status;
      else if (c.type === 'data') text += c.body;
      else if (c.type === 'error') {
        var e2 = new Error(c.message || 'network');
        e2.code = c.code;
        throw e2;
      }
    }
    var json = null;
    try {
      json = JSON.parse(text);
    } catch (e) {
      /* NDJSON or empty */
    }
    return { status: status, json: json, text: text };
  }

  function isPermissionError(e) {
    return !!(e && e.code && /permission/i.test(String(e.code)));
  }

  function errorMessage(res) {
    var code = res && res.json && res.json.error;
    if (code === 'page_not_ready') return T('notReady');
    if (code === 'busy') return T('busy');
    if (code === 'bad_url') return T('badUrl');
    return (res && res.json && res.json.message) || T('failed');
  }

  async function command(action) {
    try {
      var res = await api('/api/plugin/command', { method: 'POST', body: { action: action } });
      if (res.status !== 200) toast(errorMessage(res));
      return res.status === 200;
    } catch (e) {
      setConnected(false);
      toast(isPermissionError(e) ? T('noPermission') : T('noConn'));
      return false;
    }
  }


  /* ------------------------------------------------------------ Otzaria Voice engine */

  // The Otzaria Voice service (github.com/yakov100/otzaria-voice) runs GigaAM-He
  // CTC on ONNX Runtime with a word model of the library. It records from the
  // microphone itself and streams newline-delimited JSON back.
  async function* ovLines(path, opts) {
    opts = opts || {};
    var O = sdk();
    if (!O) throw new Error('no-sdk');
    var chunks = O.call('network.fetchStream', {
      url: OV_API + path,
      method: opts.method || 'POST',
      headers: { 'X-Otzaria-Voice': '1', 'Content-Type': 'application/json' },
      body: opts.body ? JSON.stringify(opts.body) : '',
      timeoutMs: opts.timeoutMs || 120000,
    });
    if (chunks && typeof chunks.then === 'function' && !chunks[Symbol.asyncIterator]) {
      var maybe = await chunks;
      if (maybe && maybe.success === false) throw new Error((maybe.error && maybe.error.message) || 'network');
      chunks = maybe;
    }
    var status = 200, pending = '', errBody = '';
    for await (var c of chunks) {
      if (c.type === 'response') { status = c.status; continue; }
      if (c.type === 'error') { var e = new Error(c.message || 'network'); e.code = c.code; throw e; }
      if (typeof c.body !== 'string') continue;
      if (status !== 200) { errBody += c.body; continue; }
      pending += c.body;
      var parts = pending.split('\n');
      pending = parts.pop() || '';
      for (var i = 0; i < parts.length; i++) if (parts[i].trim()) yield JSON.parse(parts[i]);
    }
    if (status !== 200) {
      var j = {};
      try { j = JSON.parse(errBody); } catch (e3) { /* not JSON */ }
      var err = new Error(j.message || 'HTTP ' + status);
      err.code = j.error || 'http_' + status;
      throw err;
    }
    if (pending.trim()) yield JSON.parse(pending);
  }

  async function ovCall(path, method) {
    var last = null;
    for await (var o of ovLines(path, { method: method || 'POST', timeoutMs: 5000 })) last = o;
    return last;
  }

  async function ovRefresh() {
    if (!sdk()) return;
    try {
      var st = await ovCall('/status', 'GET');
      S.ov = st && st.service === 'otzaria-voice' ? st : null;
    } catch (e) {
      S.ov = null;
    }
    S.ovChecked = true;
    render();
  }

  function ovUsable() {
    return !!(S.ov && S.ov.runtime);
  }

  // Which engine the Hebrew dictation uses. Explicit choice wins; otherwise the
  // desktop app when it runs, and the voice service when only it is there.
  function useOv() {
    if (S.ovLive) return true;
    var lc = S.language || 'he';
    if (lc !== 'he') return false;
    if (S.heModel === 'otzaria-voice') return true;
    if (S.heModel) return false; // one of the desktop app's models
    return !S.connected && !!S.ov;
  }

  function ovCommit(text) {
    text = String(text || '').replace(/\s+/g, ' ').trim();
    if (!text) return;
    var now = new Date();
    addLine({ id: 'ov-' + now.getTime() + '-' + ++S.ovN, text: text, at: now.toISOString(), source: 'live' });
  }

  async function ovStart() {
    if (S.ovLive) return;
    await ovRefresh();
    if (!S.ov) { toast(T('ovMissing')); return; }
    if (!S.ov.runtime) { toast(T('ovNoRuntime')); return; }
    if (!S.ov.pack) {
      var jb = S.ov.packJob && S.ov.packJob.state;
      if (!jb || jb === 'error' || jb === 'idle') { try { await ovCall('/install-pack', 'POST'); } catch (e) { /* reported below */ } }
      toast(T('ovNoPack'));
      ovWatchPack();
      return;
    }
    S.ovLive = true;
    S.armed = true;
    S.wroteSomething = false;
    S.paraAnchor = null;
    render();
    // The service ends a dictation after a sentence and a short pause; keep
    // listening until the user stops, like the desktop app does.
    while (S.ovLive) {
      var segs = {}, done = {}, reason = 'error';
      try {
        for await (var ev of ovLines('/dictate', { body: { context: [], endAfterSilenceMs: 0, noSpeechMs: 0 }, timeoutMs: 130000 })) {
          if (!S.ovLive) break;
          if (ev.type === 'ready') { S.ovListening = true; render(); }
          else if (ev.type === 'update') {
            if (done[ev.segment]) continue;
            if (ev.final) { done[ev.segment] = true; delete segs[ev.segment]; ovCommit(ev.text); }
            else segs[ev.segment] = ev.text;
            setPartial(Object.keys(segs).sort(function (a, b) { return a - b; }).map(function (k) { return segs[k]; }).join(' '));
          } else if (ev.type === 'error') toast(T('ovFailed') + (ev.message || ''));
          else if (ev.type === 'end') reason = ev.reason;
        }
      } catch (e) {
        var code = String((e && e.code) || '');
        toast(code === 'busy' ? T('ovBusy') : /microphone/.test(code) ? T('ovNoMic') : isPermissionError(e) ? T('noPermission') : T('ovFailed') + ((e && e.message) || ''));
        S.ovLive = false;
      }
      Object.keys(segs).sort(function (a, b) { return a - b; }).forEach(function (k) { ovCommit(segs[k]); });
      setPartial('');
      S.ovListening = false;
      if (reason === 'stopped' || reason === 'cancelled' || reason === 'error') S.ovLive = false;
    }
    S.ovLive = false;
    render();
  }

  function ovStop() {
    if (!S.ovLive) return;
    S.ovLive = false;
    render();
    ovCall('/stop', 'POST').catch(function () {});
  }

  var packTimer = null;
  function ovWatchPack() {
    if (packTimer) return;
    packTimer = setInterval(async function () {
      await ovRefresh();
      var jb = S.ov && S.ov.packJob;
      if (S.ov && jb && jb.state && /downloading|verifying|extracting/.test(jb.state)) {
        var pct = jb.total ? Math.round((100 * (jb.received || 0)) / jb.total) : 0;
        var note = $('otz-hlt-engine-note');
        if (note) note.textContent = T('ovPackBusy').replace('%p', pct);
        return;
      }
      clearInterval(packTimer);
      packTimer = null;
    }, 2000);
  }

  /* ------------------------------------------------------------ sync */

  function shouldPoll() {
    // Keep polling while the panel is open or while the app is reachable (so
    // dictation keeps flowing into the document with the panel closed). An app
    // that isn't running isn't retried in the background forever.
    return !S.paused && (S.panelOpen || S.connected);
  }

  async function syncLoop() {
    if (S.loopRunning || !sdk()) return;
    S.loopRunning = true;
    S.started = true;
    var failures = 0;
    while (shouldPoll()) {
      try {
        var res = await api('/api/plugin/sync?after=' + S.seq + '&wait=20', { timeoutMs: 45000 });
        if (res.status !== 200) throw new Error('HTTP ' + res.status);
        res.text.split('\n').forEach(function (ln) {
          if (!ln.trim()) return;
          try {
            handleEvent(JSON.parse(ln));
          } catch (e) {
            console.error('[otzaria-word] transcribe: bad event', e, ln);
          }
        });
        failures = 0;
        setConnected(true);
        if (Date.now() - S.helloAt > 30000) hello();
      } catch (e) {
        failures++;
        if (isPermissionError(e)) S.permissionDenied = true;
        setConnected(false);
        await sleep(failures < 3 ? 1500 : 4000);
      }
    }
    S.loopRunning = false;
  }

  // What the app supports (1.2+: model/language switching, editing, show).
  async function hello() {
    S.helloAt = Date.now();
    try {
      var r = await api('/api/plugin/hello');
      if (r.status !== 200 || !r.json) return;
      S.features = {};
      (r.json.features || []).forEach(function (f) {
        S.features[f] = true;
      });
      S.models = {};
      (r.json.models || []).forEach(function (m) {
        S.models[m.id] = m;
      });
      if (S.features.edit) {
        S.lines.forEach(function (l) {
          var row = S.lineEls[l.id],
            t = row && row.querySelector('.text');
          if (t && !t.isContentEditable) makeEditable(row, t, l);
        });
      }
      if (r.json.state) applyState(r.json.state);
      else render();
    } catch (e) {
      /* the sync loop reports connection problems */
    }
  }

  function handleEvent(ev) {
    if (typeof ev.seq === 'number') S.seq = ev.seq;
    switch (ev.type) {
      case 'snapshot':
        S.pageReady = !!ev.pageReady;
        applyState(ev.state);
        clearView();
        (ev.lines || []).forEach(function (l) {
          addLine(l);
        });
        setPartial(ev.partial || '');
        break;
      case 'line':
        addLine(ev.line);
        break;
      case 'partial':
        setPartial(ev.partial || '');
        break;
      case 'state':
        // Recording started from the desktop app's own window also writes
        // into the document (only new sentences, never the backlog).
        if (ev.state && ev.state.recording && !S.recording) {
          S.armed = !ev.state.universal;
          S.wroteSomething = false;
      S.paraAnchor = null;
        }
        applyState(ev.state);
        break;
      case 'clear':
        clearView();
        break;
      case 'edit':
        applyEdit(ev.line);
        break;
      case 'tick':
        S.pageReady = !!ev.pageReady;
        render();
        break;
    }
  }

  function applyState(st) {
    if (!st) return;
    S.engine = st.engine || 'idle';
    S.recording = !!st.recording;
    S.busy = st.busy || '';
    S.status = st.status || '';
    S.job = st.job || null;
    if (S.job && S.job.id && S.job.name) S.jobNames[S.job.id] = S.job.name;
    if (st.model) S.model = st.model;
    if (st.language) S.language = st.language;
    S.universal = !!st.universal;
    render();
  }

  function setConnected(on) {
    if (on) {
      S.everConnected = true;
      S.permissionDenied = false;
    }
    if (S.connected === on) return;
    S.connected = on;
    render();
  }

  /* ------------------------------------------------------------ the editor */

  // The superdoc of the document tab the user is looking at.
  function activeSuperdoc() {
    var entries = window.__otzariaEditors;
    if (entries && typeof entries.forEach === 'function') {
      var found = null;
      entries.forEach(function (entry) {
        if (found || !entry || !entry.pane || !entry.swap) return;
        if (entry.pane.style.display === 'none') return;
        var cur = entry.swap.current;
        if (cur && cur.superdoc) found = cur.superdoc;
      });
      if (found) return found;
    }
    var h = window.__otzariaEditor;
    return (h && (h.superdoc || h)) || null;
  }

  function caretAt(selTarget, edge) {
    var p = selTarget && selTarget[edge];
    if (!p) return null;
    var t = { kind: 'selection', start: p, end: p };
    if (selTarget.story) t.story = selTarget.story;
    if (selTarget.coordinateSpace) t.coordinateSpace = selTarget.coordinateSpace;
    return t;
  }

  async function writeToDoc(text) {
    var sd = activeSuperdoc();
    var doc = sd && sd.activeEditor && sd.activeEditor.doc;
    if (!doc || typeof doc.insert !== 'function') return { ok: false, message: T('noDoc') };

    var target = null,
      blockId = null;
    try {
      var cur = doc.selection && typeof doc.selection.current === 'function' ? await doc.selection.current() : null;
      if (cur && cur.selectionTarget && typeof cur.selectionTarget === 'object') target = cur.selectionTarget;
      var segs = cur && cur.target && Array.isArray(cur.target.segments) ? cur.target.segments : [];
      for (var i = segs.length - 1; i >= 0; i--) {
        if (segs[i] && typeof segs[i].blockId === 'string') {
          blockId = segs[i].blockId;
          break;
        }
      }
      if (!blockId && target && target.end && typeof target.end.blockId === 'string') blockId = target.end.blockId;
    } catch (e) {
      console.warn('[otzaria-word] transcribe: reading the selection failed', e);
    }

    async function insert(value) {
      var payload = { value: value, type: 'text' };
      if (target) payload.target = target;
      var receipt = await doc.insert(payload);
      if (receipt && receipt.success === false) {
        var m = receipt.failure && (receipt.failure.message || receipt.failure.code);
        throw new Error(m || 'insert');
      }
      // Move the caret to the end of what was just written, so the next
      // sentence lands after it (same thing paste does).
      var ins = receipt && receipt.effects && receipt.effects.insertedText;
      var last = Array.isArray(ins) && ins.length ? ins[ins.length - 1] : null;
      var caret = last ? caretAt(last.selectionTarget, 'end') : null;
      var surface = null;
      try {
        surface = sd.ui && sd.ui.selection;
      } catch (e) {
        surface = null;
      }
      if (caret && surface && typeof surface.apply === 'function') {
        try {
          await surface.apply(caret);
        } catch (e) {
          /* the caret simply stays where the engine left it */
        }
        target = caret;
      } else {
        target = null; // let the engine use its own caret next time
      }
    }

    try {
      // "New paragraph per sentence": from the second sentence on, add a
      // paragraph after the caret's block — the same engine call the editor
      // uses to add a paragraph after a table — and put the caret at its end.
      // If the caret hasn't moved since the last paragraph we added, chain
      // after that paragraph (the engine may not have moved the caret into it).
      if (S.newPara && S.paraAnchor && blockId === S.paraFrom) blockId = S.paraAnchor;
      if (S.newPara && S.wroteSomething && blockId && doc.create && typeof doc.create.paragraph === 'function') {
        var made = null;
        var fromBlock = blockId;
        try {
          made = await doc.create.paragraph({
            at: { kind: 'after', target: { kind: 'block', nodeType: 'paragraph', nodeId: blockId } },
            text: text,
          });
        } catch (e) {
          made = null;
        }
        var newId = made && made.success === true && made.insertionPoint && made.insertionPoint.blockId;
        if (typeof newId === 'string') {
          var surf = sd.ui && sd.ui.selection;
          var moved = false;
          if (surf && typeof surf.apply === 'function') {
            var pos = { kind: 'text', blockId: newId, offset: text.length };
            // The new block may not be painted yet; retry briefly, as the
            // editor does when it moves the caret after a table.
            for (var k = 0; k < 4 && !moved; k++) {
              if (k) await sleep(80);
              try {
                var res = await surf.apply({ kind: 'selection', start: pos, end: pos });
                moved = !res || res.ok !== false;
              } catch (e) {
                break;
              }
            }
          }
          S.paraAnchor = newId;
          S.paraFrom = moved ? newId : fromBlock;
          return { ok: true };
        }
        // Not a plain paragraph (heading, list, table cell…): continue the line.
      }
      await insert(text + (S.newPara ? '' : ' '));
      S.wroteSomething = true;
      return { ok: true };
    } catch (e) {
      return { ok: false, message: T('insertFailed') + (e && e.message ? ': ' + e.message : '') };
    }
  }

  // Inserts run strictly one after another, so sentences keep their order.
  var chain = Promise.resolve();
  function queueInsert(line, quiet) {
    chain = chain.then(async function () {
      var r = await writeToDoc(lineOut(line));
      if (r.ok) {
        S.inserted[line.id] = true;
        markInserted(line.id);
      } else if (!quiet || !S.lastInsertError) {
        toast(r.message);
      }
      S.lastInsertError = !r.ok;
      return r;
    });
    return chain;
  }

  /* ------------------------------------------------------------ view */

  var JOB_ACTIVE = { downloading: 1, queued: 1, decoding: 1, transcribing: 1 };

  function translateStatus(s) {
    if (lang() === 'en') return s;
    var m;
    if (/^Recording/.test(s)) {
      m = s.match(/catching up, (\d+)s behind/);
      return m ? T('stCatchUp').replace('%n', m[1]) : T('stRecording');
    }
    if ((m = s.match(/downloading model… (\d+)%/))) return T('stDownloading').replace('%n', m[1]);
    if (/Loading model|Warming up/.test(s)) return T('stLoading');
    if (/Microphone unavailable/.test(s)) return T('stNoMic');
    if (/Could not load the model/.test(s)) return T('stModelFailed');
    return s;
  }

  function render() {
    renderButton();
    if (!$('otz-hlt-panel')) return;
    var chip = $('otz-hlt-chip'),
      txt = $('otz-hlt-chip-text');
    chip.className = 'otz-hlt-chip';
    var ready = S.connected && S.pageReady;
    var jobActive = !!(S.job && JOB_ACTIVE[S.job.status]);
    var ov = useOv();
    if (ov) ready = ovUsable();
    if (!sdk()) txt.textContent = T('offline');
    else if (ov) {
      if (!S.ov) { chip.classList.add('bad'); txt.textContent = S.ovChecked ? T('offline') : T('connecting'); }
      else if (S.ovLive) { chip.classList.add('rec'); txt.textContent = T('recording'); }
      else { chip.classList.add('ok'); txt.textContent = T('ovConnected'); }
    } else if (!S.connected) {
      chip.classList.add('bad');
      txt.textContent = S.loopRunning && !S.everConnected ? T('connecting') : T('offline');
    } else if (!S.pageReady) txt.textContent = T('starting');
    else if (S.recording) {
      chip.classList.add('rec');
      txt.textContent = T('recording');
    } else if (S.busy === 'file' || jobActive) {
      chip.classList.add('rec');
      txt.textContent = T('fileBusy');
    } else if (S.engine === 'loading') txt.textContent = T('loadingModel');
    else {
      chip.classList.add('ok');
      txt.textContent = T('connected');
    }

    var notice = $('otz-hlt-offline');
    notice.hidden = ov ? !!S.ov || !S.ovChecked : S.connected || (!S.connected && !!S.ov && (!S.heModel || S.heModel === 'otzaria-voice'));
    $('otz-hlt-offline-body').textContent = !sdk() ? T('noSdk') : S.permissionDenied ? T('noPermission') : ov ? T('ovOfflineBody') : T('offlineBody');
    $('otz-hlt-retry').hidden = !sdk();

    var mic = $('otz-hlt-mic');
    var recNow = ov ? S.ovLive : S.recording;
    mic.disabled = ov ? !ready : !ready || S.busy === 'file' || jobActive;
    mic.classList.toggle('recording', recNow);
    mic.setAttribute('aria-label', recNow ? T('stopRec') : T('startRec'));
    $('otz-hlt-live-title').textContent = recNow ? T('stopRec') : T('startRec');
    var lc = S.language || 'he';
    var speak = lc === 'he' ? T('speakHe') : lc === 'auto' ? T('speakAny') : T('speakIn').replace('%s', langName(lc));
    var sub = speak + (S.dictate ? T('liveSubTail') : T('liveSubTailOff'));
    if (S.connected && S.status && (S.recording || S.engine === 'loading' || S.engine === 'error')) sub = translateStatus(S.status);
    $('otz-hlt-live-sub').textContent = sub;

    if (ov && S.ovListening) sub = T('ovListening') + (S.dictate ? T('liveSubTail') : T('liveSubTailOff'));
    $('otz-hlt-live-sub').textContent = sub;
    $('otz-hlt-pick').disabled = !(S.connected && S.pageReady) || !!S.busy || jobActive || S.ovLive;
    $('otz-hlt-pick').title = ov ? T('ovNoFile') : '';
    $('otz-hlt-cancel').hidden = !jobActive;
    var fileRow = $('otz-hlt-file-status'),
      prog = $('otz-hlt-progress');
    if (S.job) {
      var j = S.job,
        name = j.name || T('file');
      var pct = Math.round((j.progress || 0) * 100);
      var key = {
        downloading: 'jobDownloading',
        queued: 'jobQueued',
        decoding: 'jobDecoding',
        transcribing: 'jobTranscribing',
        done: 'jobDone',
        cancelled: 'jobCancelled',
        error: 'jobError',
      }[j.status];
      var line = key ? T(key).replace('%s', name).replace('%p', pct) : name;
      if (j.status === 'error') line += j.message || '';
      fileRow.textContent = line;
      fileRow.hidden = false;
      prog.hidden = !jobActive;
      $('otz-hlt-progress-bar').style.width = (j.status === 'transcribing' ? pct : j.status === 'decoding' ? 2 : 0) + '%';
    } else {
      fileRow.hidden = true;
      prog.hidden = true;
    }

    setToggle($('otz-hlt-dictate'), S.dictate);
    setToggle($('otz-hlt-newpara'), S.newPara);
    setToggle($('otz-hlt-ts'), S.showTs);
    $('otz-hlt-panel').classList.toggle('no-ts', !S.showTs);
    $('otz-hlt-show').hidden = !(S.connected && S.features.show);
    renderEngineBar(ready);
    var pending = S.lines.some(function (l) {
      return !S.inserted[l.id];
    });
    $('otz-hlt-insert-all').disabled = !pending;
  }

  var engineBarBuilt = false;
  function renderEngineBar(ready) {
    var bar = $('otz-hlt-engine');
    bar.hidden = false; // always visible; usable once the app (1.3+) is connected
    var supported = S.connected && S.features.config;
    var ov = useOv();
    var ls = $('otz-hlt-lang'),
      ms = $('otz-hlt-model');
    var en = lang() === 'en';
    if (!engineBarBuilt || ls.dataset.ui !== lang()) {
      engineBarBuilt = true;
      ls.dataset.ui = lang();
      ls.textContent = '';
      LANGS.forEach(function (l) {
        var o = document.createElement('option');
        o.value = l.code;
        o.textContent = en ? l.en : l.he + (l.code !== 'auto' && l.code !== 'he' ? ' · ' + l.en : '');
        ls.appendChild(o);
      });
      ms.dataset.models = '';
    }
    var lc = S.language || 'he';
    if (ls.value !== lc && document.activeElement !== ls) ls.value = lc;
    var li = langInfo(ls.value);
    var want = li.models.join(',') + lang();
    if (ms.dataset.models !== want) {
      ms.dataset.models = want;
      ms.textContent = '';
      li.models.forEach(function (id) {
        var o = document.createElement('option');
        o.value = id;
        o.textContent = MODEL_NAMES[id] ? MODEL_NAMES[id][en ? 'en' : 'he'] : id;
        ms.appendChild(o);
      });
    }
    var shown = ov ? 'otzaria-voice' : S.model;
    if (!ov && !S.connected && ls.value === 'he') shown = S.heModel || 'gigaam-he';
    if (shown && li.models.indexOf(shown) >= 0 && document.activeElement !== ms) ms.value = shown;
    var busy = S.recording || S.ovLive || !!S.busy || !!(S.job && JOB_ACTIVE[S.job.status]);
    ls.disabled = !(S.connected && S.pageReady) || busy || !supported;
    // Hebrew can always switch between the two engines.
    ms.disabled = busy || (ls.value === 'he' ? false : !ready || !supported);
    if (ms.value === 'otzaria-voice') {
      if (!packTimer) $('otz-hlt-engine-note').textContent = !S.ov ? T('ovMissing') : !S.ov.runtime ? T('ovNoRuntime') : !S.ov.pack ? T('ovNoPack') : T('ovReady');
      return;
    }
    if (!supported) {
      $('otz-hlt-engine-note').textContent = !S.connected ? T('pickerOffline') : T('pickerOld');
      return;
    }
    var m = S.models[ms.value],
      note = '';
    if (m) note = m.source === 'built-in' ? T('builtIn') : m.source === 'downloaded' ? T('installed') : T('downloads').replace('%n', m.approxMB);
    if (ms.value === 'parakeet-redux' && lc !== 'auto') note += (note ? ' · ' : '') + T('autoLang');
    $('otz-hlt-engine-note').textContent = note;
  }

  async function setEngine(language, model) {
    try {
      var res = await api('/api/plugin/command', { method: 'POST', body: { action: 'config', language: language, model: model } });
      if (res.status !== 200) {
        toast(errorMessage(res));
        render();
        return;
      }
      S.language = language;
      S.model = model;
      render();
    } catch (e) {
      setConnected(false);
      toast(isPermissionError(e) ? T('noPermission') : T('noConn'));
    }
  }

  function setToggle(btn, on) {
    btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    btn.classList.toggle('on', on);
  }

  function renderButton() {
    var b = $('otz-hlt-toggle');
    if (!b) return;
    var active = S.recording || !!(S.job && JOB_ACTIVE[S.job.status]);
    b.classList.toggle('rec', active && S.connected);
    b.classList.toggle('open', S.panelOpen);
    b.setAttribute('aria-expanded', S.panelOpen ? 'true' : 'false');
  }

  function clearView() {
    S.lines = [];
    S.lineEls = {};
    S.lastJob = null;
    var tr = $('otz-hlt-transcript');
    if (tr) {
      Array.prototype.slice.call(tr.querySelectorAll('.otz-hlt-line, .otz-hlt-sep')).forEach(function (n) {
        n.remove();
      });
      $('otz-hlt-empty').hidden = false;
    }
    S.liveEl = null;
    render();
  }

  var ICON_MIC =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" y1="19" x2="12" y2="23"/><line x1="8" y1="23" x2="16" y2="23"/></svg>';
  var ICON_INSERT =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><polyline points="14 3 14 9 20 9"/><line x1="12" y1="12" x2="12" y2="18"/><line x1="9" y1="15" x2="15" y2="15"/></svg>';
  var ICON_CHECK =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>';
  var ICON_SEARCH =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>';
  var ICON_COPY =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';
  var ICON_WINDOW =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/></svg>';
  var ICON_CLOSE =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';

  function iconButton(icon, tip, onClick) {
    var b = document.createElement('button');
    b.type = 'button';
    b.innerHTML = icon;
    b.setAttribute('data-tip-title', tip);
    b.setAttribute('aria-label', tip);
    b.addEventListener('click', onClick);
    return b;
  }

  function makeRow(l, live) {
    var row = document.createElement('div');
    row.className = 'otz-hlt-line' + (live ? ' live' : '');
    var when = document.createElement('div');
    when.className = 'when';
    when.textContent = live ? '…' : lineLabel(l);
    var text = document.createElement('div');
    text.className = 'text';
    text.textContent = l.text;
    text.dir = 'auto';
    row.appendChild(when);
    row.appendChild(text);
    if (!live && l.edited) row.classList.add('edited');
    if (!live && l.id) makeEditable(row, text, l);
    if (!live) {
      var acts = document.createElement('div');
      acts.className = 'acts';
      var ins = iconButton(ICON_INSERT, T('insertLine'), function () {
        queueInsert(l, false);
      });
      ins.className = 'ins';
      var sb = iconButton(ICON_SEARCH, T('searchLine'), function () {
        openSearch(sb._sel || defaultQuery(l.text));
        sb._sel = '';
      });
      // Read a partial selection in the line before the click collapses it.
      sb.addEventListener('mousedown', function () {
        sb._sel = selectedTextIn(text);
      });
      var cb = iconButton(ICON_COPY, T('copyLine'), function () {
        copyText(lineOut(l));
      });
      acts.appendChild(ins);
      acts.appendChild(sb);
      acts.appendChild(cb);
      row.appendChild(acts);
    }
    return row;
  }

  // Transcribed lines can be corrected in place; the app keeps the corrected
  // text so the standalone plugin and the app window show it too. A line not
  // yet written to the document is written with the correction.
  function makeEditable(row, el, l) {
    if (!S.features.edit) return;
    el.contentEditable = 'plaintext-only';
    el.spellcheck = false;
    el.setAttribute('data-tip-title', T('editTip'));
    el.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        el.blur();
      }
      if (e.key === 'Escape') {
        el.textContent = l.text;
        el.blur();
      }
      e.stopPropagation();
    });
    el.addEventListener('blur', async function () {
      var v = el.textContent.replace(/\s+/g, ' ').trim();
      if (!v) {
        el.textContent = l.text;
        return;
      }
      if (v === l.text) return;
      var old = l.text;
      l.text = v;
      l.edited = true;
      row.classList.add('edited');
      try {
        var r = await api('/api/plugin/edit', { method: 'POST', body: { id: l.id, text: v } });
        if (r.status !== 200) toast(errorMessage(r));
      } catch (e) {
        l.text = old;
        el.textContent = old;
        toast(T('editFailed'));
      }
    });
  }

  function applyEdit(nl) {
    if (!nl) return;
    for (var i = 0; i < S.lines.length; i++) {
      if (S.lines[i].id !== nl.id) continue;
      S.lines[i].text = nl.text;
      S.lines[i].edited = true;
      var row = S.lineEls[nl.id];
      if (row) {
        var t = row.querySelector('.text');
        if (document.activeElement !== t) t.textContent = nl.text;
        row.classList.add('edited');
      }
      return;
    }
  }

  function markInserted(id) {
    var row = S.lineEls[id];
    if (!row) return;
    row.classList.add('done');
    var b = row.querySelector('.ins');
    if (b) {
      b.innerHTML = ICON_CHECK;
      b.setAttribute('data-tip-title', T('inserted') + ' — ' + T('insertLine'));
      b.setAttribute('aria-label', T('inserted'));
    }
    render();
  }

  function nearBottom() {
    var c = $('otz-hlt-transcript');
    return !c || c.scrollHeight - c.scrollTop - c.clientHeight < 120;
  }
  function scrollToEnd() {
    var c = $('otz-hlt-transcript');
    if (c) c.scrollTop = c.scrollHeight;
  }

  function addLine(l) {
    if (!l || !l.id || S.lineEls[l.id]) return;
    var isNew = !S.known[l.id];
    S.known[l.id] = true;
    S.lines.push(l);
    var tr = $('otz-hlt-transcript');
    if (tr) {
      var stick = nearBottom();
      var jobKey = l.source === 'file' ? l.job || 'file' : null;
      if (jobKey && jobKey !== S.lastJob) {
        var sep = document.createElement('div');
        sep.className = 'otz-hlt-sep';
        sep.textContent = T('fromFile') + (S.jobNames[l.job] || T('recordingName'));
        tr.insertBefore(sep, S.liveEl);
      }
      S.lastJob = jobKey;
      var el = makeRow(l, false);
      S.lineEls[l.id] = el;
      tr.insertBefore(el, S.liveEl);
      $('otz-hlt-empty').hidden = true;
      if (S.inserted[l.id]) markInserted(l.id);
      if (stick) scrollToEnd();
    }
    // Only sentences that arrive while this editor is listening are written
    // automatically — never the backlog found on first connect.
    if (isNew && S.dictate && S.armed && !S.universal) queueInsert(l, true);
    render();
  }

  function setPartial(text) {
    var tr = $('otz-hlt-transcript');
    if (!tr) return;
    var stick = nearBottom();
    if (!text) {
      if (S.liveEl) {
        S.liveEl.remove();
        S.liveEl = null;
      }
      return;
    }
    if (!S.liveEl) {
      S.liveEl = makeRow({ text: '' }, true);
      tr.appendChild(S.liveEl);
    }
    S.liveEl.querySelector('.text').textContent = text;
    $('otz-hlt-empty').hidden = true;
    if (stick) scrollToEnd();
  }

  function transcriptText() {
    var out = [],
      lastJob = null;
    S.lines.forEach(function (l) {
      var jobKey = l.source === 'file' ? l.job || 'file' : null;
      if (jobKey && jobKey !== lastJob) out.push('', '— ' + (S.jobNames[l.job] || T('recordingName')) + ' —');
      lastJob = jobKey;
      out.push(lineOut(l));
    });
    return out.join('\n').replace(/^\n+/, '');
  }

  /* ------------------------------------------------------------ actions */

  async function copyText(text) {
    if (!text) {
      toast(T('nothing'));
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
    } catch (e) {
      var ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
    }
    toast(T('copied'));
  }

  function insertAll() {
    var pending = S.lines.filter(function (l) {
      return !S.inserted[l.id];
    });
    if (!pending.length) {
      toast(T('nothing'));
      return;
    }
    pending.forEach(function (l) {
      queueInsert(l, true);
    });
  }

  async function clearAll() {
    if (!S.lines.length && !S.liveEl) return;
    if (S.recording || S.busy) {
      toast(T('stopFirst'));
      return;
    }
    var ok = false;
    var O = sdk();
    if (O) {
      try {
        var c = await O.call('ui.showConfirm', { title: T('clearTitle'), content: T('clearBody') });
        ok = !!(c && c.success && c.data && c.data.confirmed);
      } catch (e) {
        ok = window.confirm(T('clearBody'));
      }
    } else ok = window.confirm(T('clearBody'));
    if (!ok) return;
    if (S.connected) await command('clear');
    else clearView();
  }

  async function toggleLive() {
    if (useOv()) {
      if (S.ovLive) ovStop();
      else ovStart();
      return;
    }
    var mic = $('otz-hlt-mic');
    if (mic) mic.disabled = true;
    var starting = !S.recording;
    if (starting) {
      S.armed = true;
      S.wroteSomething = false;
      S.paraAnchor = null;
    }
    await command(starting ? 'live-start' : 'live-stop');
    setTimeout(render, 400);
  }

  async function pickFile() {
    var O = sdk();
    if (!O) return;
    var res = await O.call('fs.pickUserFile', { title: T('pickTitle'), extensions: AUDIO_EXT });
    if (!res.success) {
      toast(T('fileOpenFailed') + (res.error ? res.error.message : ''));
      return;
    }
    if (res.data.cancelled) return;
    $('otz-hlt-pick').disabled = true;
    S.armed = true;
    S.wroteSomething = false;
      S.paraAnchor = null;
    try {
      var r = await api('/api/plugin/file', { method: 'POST', body: { url: res.data.url, name: res.data.name } });
      if (r.status !== 200) toast(errorMessage(r));
      else if (r.json && r.json.job) S.jobNames[r.json.job] = res.data.name;
    } catch (e) {
      setConnected(false);
      toast(isPermissionError(e) ? T('noPermission') : T('noConn'));
    }
    render();
  }

  function selectedTextIn(node) {
    var sel = window.getSelection();
    if (!sel || sel.isCollapsed) return '';
    if (!node.contains(sel.anchorNode) || !node.contains(sel.focusNode)) return '';
    return sel.toString().trim();
  }
  function defaultQuery(text) {
    var words = text
      .replace(/[.,:;!?"״׳()]/g, ' ')
      .split(/\s+/)
      .filter(Boolean);
    return words.slice(0, 6).join(' ');
  }

  // The editor already holds reader.open, which covers Otzaria's search tab.
  async function openSearch(query) {
    var O = sdk();
    if (!O || !query) return;
    try {
      var r = await O.call('reader.openSearchTab', { query: query });
      if (r && r.success === false) toast(T('searchFailed'));
    } catch (e) {
      toast(T('searchFailed'));
    }
  }

  /* ------------------------------------------------------------ panel */

  function openPanel() {
    ensurePanel();
    S.panelOpen = true;
    $('otz-hlt-panel').hidden = false;
    if (!S.loopRunning) syncLoop();
    render();
    scrollToEnd();
  }
  function closePanel() {
    S.panelOpen = false;
    var p = $('otz-hlt-panel');
    if (p) p.hidden = true;
    render();
  }
  function togglePanel() {
    if (S.panelOpen) closePanel();
    else openPanel();
  }

  function ensurePanel() {
    if ($('otz-hlt-panel')) return;
    var p = document.createElement('aside');
    p.id = 'otz-hlt-panel';
    // Our own Hebrew/English strings; keep english-mode.js from rewriting them.
    p.className = 'otz-no-translate';
    p.setAttribute('role', 'dialog');
    p.hidden = true;
    p.innerHTML =
      '<div class="otz-hlt-head">' +
      '  <div class="otz-hlt-title"></div>' +
      '  <span class="otz-hlt-chip" id="otz-hlt-chip"><span class="dot"></span><span id="otz-hlt-chip-text"></span></span>' +
      '  <button type="button" class="otz-hlt-icon" id="otz-hlt-show" hidden>' + ICON_WINDOW + '</button>' +
      '  <button type="button" class="otz-hlt-icon" id="otz-hlt-close">' + ICON_CLOSE + '</button>' +
      '</div>' +
      '<div class="otz-hlt-notice" id="otz-hlt-offline" hidden>' +
      '  <div class="h" id="otz-hlt-offline-title"></div>' +
      '  <div id="otz-hlt-offline-body"></div>' +
      '  <button type="button" class="otz-hlt-btn tonal" id="otz-hlt-retry"></button>' +
      '</div>' +
      '<div class="otz-hlt-engine" id="otz-hlt-engine" hidden>' +
      '  <label for="otz-hlt-lang" id="otz-hlt-lang-l"></label><select id="otz-hlt-lang"></select>' +
      '  <label for="otz-hlt-model" id="otz-hlt-model-l"></label><select id="otz-hlt-model"></select>' +
      '  <span class="note" id="otz-hlt-engine-note"></span>' +
      '</div>' +
      '<div class="otz-hlt-controls">' +
      '  <button type="button" class="otz-hlt-mic" id="otz-hlt-mic" disabled>' + ICON_MIC + '</button>' +
      '  <div class="otz-hlt-ctext">' +
      '    <div class="t" id="otz-hlt-live-title"></div>' +
      '    <div class="s" id="otz-hlt-live-sub"></div>' +
      '  </div>' +
      '</div>' +
      '<div class="otz-hlt-opts">' +
      '  <button type="button" class="otz-hlt-switch" id="otz-hlt-dictate"><i></i><span></span></button>' +
      '  <button type="button" class="otz-hlt-switch" id="otz-hlt-newpara"><i></i><span></span></button>' +
      '  <button type="button" class="otz-hlt-switch" id="otz-hlt-ts"><i></i><span></span></button>' +
      '</div>' +
      '<div class="otz-hlt-file">' +
      '  <button type="button" class="otz-hlt-btn tonal" id="otz-hlt-pick" disabled></button>' +
      '  <button type="button" class="otz-hlt-btn" id="otz-hlt-cancel" hidden></button>' +
      '  <div class="otz-hlt-fstatus" id="otz-hlt-file-status" hidden></div>' +
      '  <div class="otz-hlt-progress" id="otz-hlt-progress" hidden><i id="otz-hlt-progress-bar"></i></div>' +
      '</div>' +
      '<div class="otz-hlt-transcript" id="otz-hlt-transcript" aria-live="polite">' +
      '  <div class="otz-hlt-empty" id="otz-hlt-empty"></div>' +
      '</div>' +
      '<div class="otz-hlt-foot">' +
      '  <button type="button" class="otz-hlt-btn filled" id="otz-hlt-insert-all"></button>' +
      '  <span class="grow"></span>' +
      '  <button type="button" class="otz-hlt-btn" id="otz-hlt-copy-all"></button>' +
      '  <button type="button" class="otz-hlt-btn" id="otz-hlt-clear"></button>' +
      '</div>' +
      '<div class="otz-hlt-toast" id="otz-hlt-toast" role="status"></div>';
    document.body.appendChild(p);

    // Clicking the panel must not take the caret out of the document:
    // dictation writes where the user left it.
    p.addEventListener('mousedown', function (e) {
      if (e.target.closest('button')) e.preventDefault();
    });

    $('otz-hlt-close').addEventListener('click', closePanel);
    $('otz-hlt-retry').addEventListener('click', function () {
      if (!S.loopRunning) syncLoop();
      render();
    });
    $('otz-hlt-mic').addEventListener('click', toggleLive);
    $('otz-hlt-pick').addEventListener('click', function () {
      pickFile().catch(function (e) {
        console.error(e);
        toast(T('failed'));
      });
    });
    $('otz-hlt-cancel').addEventListener('click', function () {
      command('cancel');
    });
    $('otz-hlt-dictate').addEventListener('click', function () {
      S.dictate = !S.dictate;
      writeBool(PREF_DICTATE, S.dictate);
      render();
    });
    $('otz-hlt-newpara').addEventListener('click', function () {
      S.newPara = !S.newPara;
      writeBool(PREF_NEWLINE, S.newPara);
      render();
    });
    $('otz-hlt-ts').addEventListener('click', function () {
      S.showTs = !S.showTs;
      writeBool(PREF_TS, S.showTs);
      render();
    });
    $('otz-hlt-show').addEventListener('click', function () {
      command('show');
    });
    $('otz-hlt-lang').addEventListener('change', function () {
      var li = langInfo($('otz-hlt-lang').value);
      var m = li.models.indexOf(S.model) >= 0 ? S.model : li.models[0];
      if (m === 'otzaria-voice') m = 'gigaam-he';
      setEngine(li.code, m);
    });
    $('otz-hlt-model').addEventListener('change', function () {
      var m = $('otz-hlt-model').value;
      if ($('otz-hlt-lang').value === 'he' || !S.connected) {
        S.heModel = m;
        writeStr(PREF_HE_MODEL, m);
        if (m === 'otzaria-voice') { ovRefresh(); render(); return; }
        if (!S.connected) { render(); return; }
      }
      setEngine($('otz-hlt-lang').value, m);
    });
    $('otz-hlt-insert-all').addEventListener('click', insertAll);
    $('otz-hlt-copy-all').addEventListener('click', function () {
      copyText(transcriptText());
    });
    $('otz-hlt-clear').addEventListener('click', clearAll);

    // Rows added before the panel existed (dictation with the panel closed).
    var backlog = S.lines.slice();
    S.lines = [];
    S.lineEls = {};
    S.lastJob = null;
    backlog.forEach(addLine);
    applyLabels();
  }

  function applyLabels() {
    var p = $('otz-hlt-panel');
    if (p) {
      p.setAttribute('dir', 'rtl');
      p.setAttribute('aria-label', T('title'));
      p.querySelector('.otz-hlt-title').textContent = T('title');
      var close = $('otz-hlt-close');
      close.setAttribute('aria-label', T('close'));
      close.setAttribute('data-tip-title', T('close'));
      $('otz-hlt-offline-title').textContent = T('offlineTitle');
      $('otz-hlt-retry').textContent = T('retry');
      $('otz-hlt-pick').textContent = T('pickFile');
      $('otz-hlt-cancel').textContent = T('cancel');
      $('otz-hlt-dictate').querySelector('span').textContent = T('dictate');
      $('otz-hlt-dictate').setAttribute('data-tip-title', T('dictateTip'));
      $('otz-hlt-newpara').querySelector('span').textContent = T('newPara');
      $('otz-hlt-newpara').setAttribute('data-tip-title', T('newParaTip'));
      $('otz-hlt-ts').querySelector('span').textContent = T('timestamps');
      $('otz-hlt-ts').setAttribute('data-tip-title', T('timestampsTip'));
      $('otz-hlt-lang-l').textContent = T('language');
      $('otz-hlt-model-l').textContent = T('model');
      $('otz-hlt-show').setAttribute('aria-label', T('showApp'));
      $('otz-hlt-show').setAttribute('data-tip-title', T('showApp'));
      $('otz-hlt-empty').textContent = T('empty');
      $('otz-hlt-insert-all').textContent = T('insertAll');
      $('otz-hlt-copy-all').textContent = T('copyAll');
      $('otz-hlt-clear').textContent = T('clear');
      Object.keys(S.lineEls).forEach(function (id) {
        var row = S.lineEls[id];
        var bs = row.querySelectorAll('.acts button');
        if (bs[0]) {
          var t0 = S.inserted[id] ? T('inserted') + ' — ' + T('insertLine') : T('insertLine');
          bs[0].setAttribute('data-tip-title', t0);
          bs[0].setAttribute('aria-label', t0);
        }
        if (bs[1]) {
          bs[1].setAttribute('data-tip-title', T('searchLine'));
          bs[1].setAttribute('aria-label', T('searchLine'));
        }
        if (bs[2]) {
          bs[2].setAttribute('data-tip-title', T('copyLine'));
          bs[2].setAttribute('aria-label', T('copyLine'));
        }
      });
    }
    var b = $('otz-hlt-toggle');
    if (b) {
      b.querySelector('.otz-hlt-label').textContent = T('button');
      b.setAttribute('data-tip-title', T('buttonTip'));
      b.setAttribute('aria-label', T('buttonTip'));
    }
    render();
  }

  /* ------------------------------------------------------------ title-bar button */

  var BTN_ID = 'otz-hlt-toggle';

  function ensureButton() {
    if ($(BTN_ID)) return;
    var host = document.querySelector('.titlebar-end');
    if (!host) return;
    var b = document.createElement('button');
    b.type = 'button';
    b.id = BTN_ID;
    b.className = 'otz-no-translate';
    b.setAttribute('aria-haspopup', 'dialog');
    b.innerHTML = ICON_MIC + '<span class="otz-hlt-label"></span><span class="otz-hlt-recdot" aria-hidden="true"></span>';
    b.addEventListener('mousedown', function (e) {
      e.preventDefault();
    });
    b.addEventListener('click', togglePanel);
    // Sit right before the language button (english-mode.js puts it first).
    var lang = document.getElementById('otz-lang-toggle');
    host.insertBefore(b, lang ? lang.nextSibling : host.firstChild);
    applyLabels();
  }

  /* ------------------------------------------------------------ styles */

  var css = [
    '#otz-hlt-toggle{display:inline-flex;align-items:center;gap:5px;height:24px;padding:0 10px;margin-inline-end:8px;border-radius:999px;cursor:pointer;font:inherit;font-size:12px;font-weight:600;line-height:1;color:var(--color-primary,#1565c0);background:transparent;border:1px solid var(--color-primary,#1565c0);white-space:nowrap;transition:background-color .15s,color .15s;position:relative}',
    '#otz-hlt-toggle svg{width:14px;height:14px;flex-shrink:0}',
    '#otz-hlt-toggle:hover,#otz-hlt-toggle.open{background:var(--color-primary,#1565c0);color:var(--color-on-primary,#fff)}',
    '#otz-hlt-toggle:focus-visible{outline:2px solid var(--color-primary,#1565c0);outline-offset:2px}',
    '#otz-hlt-toggle .otz-hlt-recdot{display:none;width:7px;height:7px;border-radius:50%;background:var(--color-error,#b00020);animation:otzHltPulse 1.2s ease-in-out infinite}',
    '#otz-hlt-toggle.rec .otz-hlt-recdot{display:inline-block}',
    '#otz-hlt-toggle.rec{border-color:var(--color-error,#b00020)}',
    '@keyframes otzHltPulse{0%,100%{opacity:1}50%{opacity:.25}}',

    '#otz-hlt-panel{position:fixed;inset-inline-end:16px;bottom:40px;z-index:2000;width:min(380px,calc(100vw - 32px));max-height:min(560px,calc(100vh - 140px));display:flex;flex-direction:column;background:var(--color-surface,#fff);color:var(--color-on-surface,#1a1a2e);border:1px solid var(--color-outline-variant,var(--color-outline,#c4c4c4));border-radius:12px;box-shadow:0 10px 32px rgba(0,0,0,.18);font-family:var(--font-main,"Assistant","Segoe UI",sans-serif);font-size:13px;overflow:hidden}',
    '#otz-hlt-panel[hidden]{display:none}',
    '#otz-hlt-panel button{font:inherit;color:inherit}',
    '.otz-hlt-head{display:flex;align-items:center;gap:8px;padding:10px 12px 8px;border-bottom:1px solid var(--color-outline-variant,#e1dfdd)}',
    '.otz-hlt-title{font-weight:700;font-size:14px}',
    '.otz-hlt-chip{display:inline-flex;align-items:center;gap:5px;padding:2px 8px;border-radius:999px;font-size:11px;background:var(--color-surface-container-high,#f0f0f0);color:var(--color-on-surface-variant,#49454f)}',
    '.otz-hlt-chip .dot{width:7px;height:7px;border-radius:50%;background:currentColor;opacity:.6}',
    '.otz-hlt-chip.ok .dot{background:#2e7d32;opacity:1}',
    '.otz-hlt-chip.bad .dot{background:var(--color-error,#b00020);opacity:1}',
    '.otz-hlt-chip.rec{color:var(--color-error,#b00020)}',
    '.otz-hlt-chip.rec .dot{background:var(--color-error,#b00020);opacity:1;animation:otzHltPulse 1.2s ease-in-out infinite}',
    '#otz-hlt-panel .otz-hlt-icon[hidden]{display:none}',
    '#otz-hlt-panel .otz-hlt-chip{margin-inline-end:auto}',
    '#otz-hlt-panel .otz-hlt-icon{width:28px;height:28px;padding:0;display:inline-flex;align-items:center;justify-content:center;border:0;background:transparent;border-radius:6px;cursor:pointer}',
    '.otz-hlt-icon:hover{background:var(--color-surface-container-high,#f0f0f0)}',
    '#otz-hlt-panel .otz-hlt-icon svg{width:16px;min-width:16px;height:16px;flex-shrink:0}',
    '.otz-hlt-notice{margin:10px 12px 0;padding:10px 12px;border-radius:8px;background:var(--color-error-subtle,rgba(176,0,32,.08));line-height:1.5}',
    '.otz-hlt-notice .h{font-weight:700;margin-bottom:4px}',
    '.otz-hlt-notice .otz-hlt-btn{margin-top:8px}',
    '.otz-hlt-controls{display:flex;align-items:center;gap:12px;padding:12px}',
    '#otz-hlt-panel .otz-hlt-mic{flex-shrink:0;width:48px;height:48px;padding:0;border-radius:50%;border:0;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;background:var(--color-primary,#1565c0);color:var(--color-on-primary,#fff)!important;transition:transform .1s,background-color .15s}',
    '#otz-hlt-panel .otz-hlt-mic svg{width:22px;min-width:22px;height:22px;flex-shrink:0}',
    '#otz-hlt-panel .otz-hlt-mic:hover:not(:disabled){transform:scale(1.05)}',
    '#otz-hlt-panel .otz-hlt-mic:disabled{opacity:.4;cursor:default}',
    '#otz-hlt-panel .otz-hlt-mic.recording{background:var(--color-error,#b00020);box-shadow:0 0 0 0 rgba(176,0,32,.5);animation:otzHltRing 1.4s ease-out infinite}',
    '@keyframes otzHltRing{0%{box-shadow:0 0 0 0 rgba(176,0,32,.45)}100%{box-shadow:0 0 0 12px rgba(176,0,32,0)}}',
    '.otz-hlt-ctext{min-width:0}',
    '.otz-hlt-ctext .t{font-weight:700}',
    '.otz-hlt-ctext .s{color:var(--color-on-surface-variant,#49454f);font-size:12px;line-height:1.4;margin-top:2px}',
    '.otz-hlt-opts{display:flex;flex-wrap:wrap;gap:6px 14px;padding:0 12px 10px}',
    '.otz-hlt-switch{display:inline-flex;align-items:center;gap:6px;border:0;background:transparent;cursor:pointer;padding:2px 0;font-size:12px}',
    '.otz-hlt-switch i{position:relative;width:28px;height:16px;border-radius:999px;background:var(--color-outline,#79747e);transition:background-color .15s;flex-shrink:0}',
    '.otz-hlt-switch i::after{content:"";position:absolute;top:2px;inset-inline-start:2px;width:12px;height:12px;border-radius:50%;background:#fff;transition:inset-inline-start .15s}',
    '.otz-hlt-switch.on i{background:var(--color-primary,#1565c0)}',
    '.otz-hlt-switch.on i::after{inset-inline-start:14px}',
    '.otz-hlt-engine{display:flex;flex-wrap:wrap;align-items:center;gap:4px 8px;padding:8px 12px 0;font-size:12px}',
    '.otz-hlt-engine[hidden]{display:none}',
    '.otz-hlt-engine label{font-weight:600;color:var(--color-on-surface-variant,#49454f)}',
    '#otz-hlt-panel .otz-hlt-engine select{font:inherit;font-size:12px;padding:3px 6px;border-radius:6px;border:1px solid var(--color-outline,#79747e);background:var(--color-surface,#fff);color:var(--color-on-surface,#1a1a2e);max-width:150px}',
    '.otz-hlt-engine select:disabled{opacity:.55}',
    '.otz-hlt-engine .note{flex-basis:100%;color:var(--color-on-surface-variant,#49454f);font-size:11px}',
    '#otz-hlt-panel.no-ts .otz-hlt-line .when{display:none}',
    '.otz-hlt-line .text[contenteditable]{outline:none;border-radius:4px;cursor:text}',
    '.otz-hlt-line .text[contenteditable]:focus{box-shadow:0 0 0 2px var(--color-primary,#1565c0);background:var(--color-surface,#fff)}',
    '.otz-hlt-line.edited .when::after{content:" ✎"}',
    '.otz-hlt-file{display:flex;flex-wrap:wrap;align-items:center;gap:6px 8px;padding:0 12px 10px}',
    '.otz-hlt-fstatus{flex-basis:100%;font-size:12px;color:var(--color-on-surface-variant,#49454f)}',
    '.otz-hlt-progress{flex-basis:100%;height:4px;border-radius:999px;background:var(--color-surface-container-highest,#e6e0e9);overflow:hidden}',
    '.otz-hlt-progress i{display:block;height:100%;width:0;background:var(--color-primary,#1565c0);transition:width .3s}',
    '.otz-hlt-btn{height:28px;padding:0 12px;border-radius:999px;border:1px solid var(--color-outline-variant,#c4c4c4);background:transparent;cursor:pointer;font-size:12px;font-weight:600;white-space:nowrap}',
    '.otz-hlt-btn:hover:not(:disabled){background:var(--color-surface-container-high,#f0f0f0)}',
    '.otz-hlt-btn:disabled{opacity:.45;cursor:default}',
    '.otz-hlt-btn.tonal{border-color:transparent;background:var(--color-secondary-container,#e8def8);color:var(--color-on-secondary-container,#1d192b)}',
    '.otz-hlt-btn.filled{border-color:transparent;background:var(--color-primary,#1565c0);color:var(--color-on-primary,#fff)}',
    '.otz-hlt-btn.filled:hover:not(:disabled){background:var(--color-primary,#1565c0);filter:brightness(1.08)}',
    '.otz-hlt-transcript{flex:1;min-height:90px;overflow-y:auto;border-top:1px solid var(--color-outline-variant,#e1dfdd);padding:6px 0}',
    '.otz-hlt-empty{padding:16px;text-align:center;color:var(--color-on-surface-variant,#49454f);font-size:12px;line-height:1.5}',
    '.otz-hlt-sep{margin:8px 12px 2px;font-size:11px;font-weight:700;color:var(--color-primary,#1565c0)}',
    '.otz-hlt-line{display:flex;align-items:flex-start;gap:8px;padding:5px 12px}',
    '.otz-hlt-line:hover{background:var(--color-surface-container-high,#f5f5f5)}',
    '.otz-hlt-line .when{flex-shrink:0;width:38px;font-size:11px;color:var(--color-on-surface-variant,#49454f);padding-top:2px;direction:ltr;text-align:end}',
    '.otz-hlt-line .text{flex:1;min-width:0;unicode-bidi:plaintext;line-height:1.55;font-size:14px;user-select:text;word-wrap:break-word}',
    '.otz-hlt-line.live .text{color:var(--color-on-surface-variant,#49454f);font-style:italic}',
    '.otz-hlt-line.done .text{color:var(--color-on-surface-variant,#49454f)}',
    '.otz-hlt-line .acts{display:flex;gap:2px;flex-shrink:0;opacity:0;transition:opacity .12s}',
    '.otz-hlt-line:hover .acts,.otz-hlt-line .acts:focus-within{opacity:1}',
    '#otz-hlt-panel .otz-hlt-line.done .acts .ins{color:#2e7d32}',
    '.otz-hlt-line.done .acts{opacity:1}',
    '.otz-hlt-line.done:not(:hover) .acts button:not(.ins){visibility:hidden}',
    '#otz-hlt-panel .otz-hlt-line .acts button{width:24px;min-width:24px;height:24px;padding:0;flex-shrink:0;border:0;background:transparent;border-radius:5px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;color:var(--color-on-surface-variant,#49454f)}',
    '#otz-hlt-panel .otz-hlt-line .acts button:hover{background:var(--color-surface-container-highest,#e6e0e9);color:var(--color-primary,#1565c0)}',
    '#otz-hlt-panel .otz-hlt-line .acts svg{width:15px;min-width:15px;height:15px;flex-shrink:0}',
    '.otz-hlt-foot{display:flex;align-items:center;gap:6px;padding:8px 12px;border-top:1px solid var(--color-outline-variant,#e1dfdd)}',
    '.otz-hlt-foot .grow{flex:1}',
    '.otz-hlt-toast{position:absolute;inset-inline:12px;bottom:52px;padding:8px 12px;border-radius:8px;background:var(--color-inverse-surface,#313033);color:var(--color-inverse-on-surface,#f4eff4);font-size:12px;text-align:center;opacity:0;transform:translateY(6px);transition:opacity .18s,transform .18s;pointer-events:none}',
    '.otz-hlt-toast.show{opacity:1;transform:none}',
    ':root[data-ui-language=en] #otz-hlt-panel{direction:ltr}',
    '@media (prefers-reduced-motion:reduce){.otz-hlt-mic.recording,#otz-hlt-toggle .otz-hlt-recdot,.otz-hlt-chip.rec .dot{animation:none}}',
  ].join('\n');

  /* ------------------------------------------------------------ boot */

  function onSdk(fn) {
    if (sdk()) {
      fn(sdk());
      return;
    }
    var tries = 0;
    var t = setInterval(function () {
      if (sdk() || ++tries > 100) {
        clearInterval(t);
        if (sdk()) fn(sdk());
      }
    }, 200);
  }

  function boot() {
    var style = document.createElement('style');
    style.textContent = css;
    document.head.appendChild(style);

    // Keep the button present even if the title bar is re-rendered.
    new MutationObserver(function () {
      if (!$(BTN_ID)) ensureButton();
    }).observe(document.body, { childList: true, subtree: true });
    ensureButton();

    // Follow the interface language chosen in english-mode.js.
    new MutationObserver(function () {
      applyLabels();
      var p = $('otz-hlt-panel');
      if (p) p.setAttribute('dir', lang() === 'en' ? 'ltr' : 'rtl');
    }).observe(document.documentElement, { attributes: true, attributeFilter: ['data-ui-language'] });

    document.addEventListener(
      'keydown',
      function (e) {
        // Alt+V: start/stop dictation (opens the panel the first time).
        if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && e.code === 'KeyV') {
          e.preventDefault();
          e.stopPropagation();
          if (!S.panelOpen) openPanel();
          var mic = $('otz-hlt-mic');
          if (mic && !mic.disabled) toggleLive();
          return;
        }
        if (e.key === 'Escape' && S.panelOpen && document.activeElement && $('otz-hlt-panel').contains(document.activeElement)) closePanel();
      },
      true,
    );

    onSdk(function (O) {
      if (typeof O.on !== 'function') return;
      // Leaving the tab: stop polling (transcription carries on in the
      // desktop app); on return, catch up from where we were.
      O.on('plugin.suspended', function () {
        S.paused = true;
      });
      setInterval(function () {
        if (!S.paused && (S.panelOpen || S.ovLive) && !packTimer) ovRefresh();
      }, 6000);
      ovRefresh();
      O.on('plugin.resumed', function () {
        S.paused = false;
        if (S.started && !S.loopRunning) syncLoop();
      });
    });
  }

  if (document.body) boot();
  else document.addEventListener('DOMContentLoaded', boot, { once: true });
})();

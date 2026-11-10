# KolScribe (קול סופר)

*Formerly "Live Transcribe" and before that "Hebrew Live Transcribe".*

Free, on-device speech-to-text for Windows and Mac. Speak Hebrew, English and 25+ other languages and watch the text appear live. Everything runs on your own computer: your audio never leaves it.

There are two apps, built from the same code:

* **KolScribe (קול סופר)**: for Hebrew speakers, Torah study and Otzaria users. English and Hebrew interface.
* **KolScribe English**: for English-speaking users. English only (Hebrew can be switched on in Settings), without the Torah-study and Otzaria features. Installs and runs side by side with KolScribe.

## Download

**[⬇ Download the latest version](https://github.com/shlomyeisenberg-lang/live-transcribe/releases/latest)**

| KolScribe | KolScribe English | What it is |
|---|---|---|
| `KolScribe-Setup.exe` | `KolScribeEnglish-Setup.exe` | **Windows installer (recommended).** Per-user, no administrator needed, updates in place. |
| `KolScribe.exe` | `KolScribeEnglish.exe` | Portable Windows app (~65 MB). Extra models download once, the first time you use them. |
| `KolScribe-Offline.exe` | `KolScribeEnglish-Offline.exe` | The same app with the models inside, for computers without internet (large). |
| `KolScribe-mac.dmg` | `KolScribeEnglish-mac.dmg` | **Mac** (macOS 12+, Apple silicon and Intel). Open it and drag the app to Applications. |
| `KolScribe-Offline-mac.dmg` | `KolScribeEnglish-Offline-mac.dmg` | The Mac app with the models inside. |

Also in each release: `*.otzplugin` (plugins for [Otzaria](https://github.com/Otzaria): live transcription, and a custom Word editor with dictation; install in Otzaria via **Settings → Plugins → Install from file**), `SHA256SUMS.txt` to verify your download, and the files the apps use to update themselves (`latest-*.json`, `catalog.json`, the Mac `.zip` files).

Windows may warn that the app is unsigned: click **More info → Run anyway**.

Mac: the first time, macOS says it cannot check the app (it is free and not registered with Apple). On macOS 15 or later press **Done**, then go to **System Settings → Privacy & Security** and click **Open Anyway**. On macOS 12–14, right-click the app and choose **Open**. For Type anywhere (**Control+Option+D**), allow the app under **Privacy & Security → Accessibility**.

## What it does

* **Live transcription** as you speak, or transcribe an audio/video file.
* **Type anywhere:** dictate straight into Word, a browser, WhatsApp, any app, with a keyboard shortcut you choose.
* **Voice commands:** say "period", "new line", "scratch that", "delete last word", "stop dictation" and more, in English and Hebrew. Edit, add or import your own commands in Settings.
* **Hands-free start** (off by default): the app listens for "start dictation" (or "התחל הכתבה") and starts transcribing. Listening audio is never saved or sent.
* **Get more models:** add speech-to-text models from Hugging Face (catalog, search or paste a link), with a check of whether each model runs on your computer.
* **Rabbinic Hebrew boost** (KolScribe): better recognition of Torah and Talmud vocabulary in any Hebrew model.
* **Library:** every recording is saved with its audio; play, correct, re-transcribe and export (TXT, SRT, VTT, Word).
* **AI tools** (optional, your choice of local or cloud model): clean up, polish, summarize and translate. A **Summary card** appears above the original text, which is kept. You can set a time limit for an AI answer.
* **Automatic updates** for the installer and portable versions (you can turn the check off). The Microsoft Store versions are updated by the Store.
* Models include GigaAM with the Otzaria library word model, Qwen3-ASR Hebrew, Whistle, Parakeet and Moonshine.
* English and Hebrew interface, light and dark themes.

## Privacy

Your voice and transcripts stay on your computer; the apps have no telemetry, analytics or account. See the full [privacy policy](PRIVACY.md) ([עברית](PRIVACY.he.md)).

## Feedback, bugs and feature requests

* **[Open an issue](https://github.com/shlomyeisenberg-lang/live-transcribe/issues/new)** for bugs and feature requests.
* Anything private, complaints or questions: **[ceshlomo@gmail.com](mailto:ceshlomo@gmail.com)**

## Need software or AI for your business?

If you like this app, I build custom software and set up AI and automation for businesses. **[ceshlomo@gmail.com](mailto:ceshlomo@gmail.com)** · [github.com/shlomyeisenberg-lang](https://github.com/shlomyeisenberg-lang)

## License

KolScribe and KolScribe English are free to download and use. It is not open source: the app's source code is private and all rights are reserved by Shlomo Eisenberg. You may not copy, modify, redistribute or sell the app or its code without written permission.

Third-party parts keep their own licenses (see the notices inside the app and its releases). The **Custom Word editor plugin** is based on otzaria-word-editor and SuperDoc, which are licensed under the GNU AGPL v3; its full source is published here in [`word-editor-plugin/`](word-editor-plugin/) under that license.

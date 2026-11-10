# Privacy policy — KolScribe and KolScribe English

*Last updated: 2026-11. Publisher: CESH.* עברית: [PRIVACY.he.md](PRIVACY.he.md)

KolScribe (קול סופר) and KolScribe English are speech-to-text apps for Windows and macOS.
**Your voice and your transcripts stay on your computer.**

## What the app does with your data

- **Speech recognition runs on your device.** Audio from the microphone (or from a file you open)
  is processed locally by the speech models. It is not sent to us or to anyone else.
- **No telemetry, no analytics, no advertising, no account.** We run no server that the app reports to.
  We do not collect, receive or sell any personal data.
- **Your files stay with you.** Transcripts, the library, settings and downloaded models are stored in
  the app's data folder on your computer (`%LOCALAPPDATA%\HebrewLiveTranscribe` for KolScribe,
  `%LOCALAPPDATA%\KolScribeEnglish` for KolScribe English; `~/Library/Application Support/…` on a Mac).
  Uninstalling the app does not delete them; you can delete the folder yourself.
- **Microphone:** used only while you record or dictate. Windows and macOS ask for your permission.
- **Typing into other apps and global shortcuts** (Type anywhere, voice commands) only send the text you
  dictated to the app that has focus, on your computer. The app's local server listens on `127.0.0.1`
  only (not reachable from other computers); other programs on your computer, such as the Otzaria plugins
  in KolScribe, can connect to it.

## When the app uses the network

The app connects to the internet only for things you ask for or switch on:

1. **Downloading models and runtimes** (from Hugging Face, or from the public KolScribe downloads page on
   GitHub), and the list of recommended models. These requests contain what any download does (your IP
   address, the file requested); the app sends no information about you.
2. **Checking for updates** (installer and portable versions, on GitHub; you can turn it off in
   Settings → About & updates). The Microsoft Store version does not check: the Store updates it.
3. **AI features, only if you set up a provider** (Settings → AI). Then the text you choose to polish,
   summarize or translate is sent to **that provider** (for example OpenAI, Anthropic, Google, or a model
   running on your own computer such as Ollama or LM Studio), under that provider's terms and privacy
   policy. Your API keys are kept in the system key store (Windows Credential/DPAPI, macOS Keychain).
   Nothing is sent to an AI provider unless you press the button.
4. **Links you click** (for example Help or the project page) open in your browser.

## Children

The app is a general-purpose tool and is not directed at children; it collects no data from anyone.

## Changes and contact

If this policy changes, the new version is published at the same address and the date above changes.
Questions: open an issue at https://github.com/shlomyeisenberg-lang/live-transcribe/issues

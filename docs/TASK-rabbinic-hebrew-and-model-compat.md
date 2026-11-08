# Proposed new task: Rabbinic-Hebrew data for any Hebrew model + model compatibility check

Written as a standalone task so it does not clash with task 5 or task 6.
Suggested handling: run it AFTER task 6 finishes, as a separate task (task 7).
It touches only the Settings > Models UI and model-loading code. Check what
tasks 5 and 6 touch before starting, and rebase if they changed the same files.

## Part A - Rabbinic Hebrew data usable with any Hebrew model

Today one Hebrew model (GigaAM with the Otzaria library word model) uses extra
data (the Otzaria word list / language model) to recognize rabbinic Hebrew.

Goal: let the user apply that data to ANY Hebrew model, not only the default.

1. In Settings, add an option "Use rabbinic Hebrew vocabulary (Otzaria library data)".
   - Offered for every installed model whose language is Hebrew.
   - Stored per model, defaulting to on for the current default model so
     existing behavior does not change.
2. Decouple the extra data from the model definition: load the word data as an
   independent resource and attach it to whichever model is selected.
   - Models that can use it directly (decoder biasing / hotwords / LM
     rescoring) get it as is.
   - Models with no such hook (e.g. Whistle, Qwen3-ASR) get it as a
     post-processing step instead (word-level correction against the list).
     Show which method is used next to the option.
3. Download the data once on first enable. Show size and progress, and make it
   work offline when the Offline build ships the data.
4. Tests: output with the option on and off, for at least one model of each type.

## Part B - Hugging Face model downloader: compatibility check

When the user downloads a model from Hugging Face inside the app:

1. Before downloading, read the model's files and config and mark it as one of:
   - Compatible with this app version
   - Needs a newer runtime (e.g. a newer ONNX Runtime or other runtime)
   - Not supported
   Show the reason in plain language.
2. If it needs a newer runtime, offer "Install runtime" inside the app: download
   the required runtime component into the app's own data folder, verify its
   checksum, restart the engine, and re-check. No manual steps and no admin
   rights. Roll back to the old runtime if it fails.
3. Ask the user before downloading anything, and show sizes.
4. Tests: compatible, needs-runtime and unsupported models, and a failed install.

## Open questions for the owner
- Which runtime(s) does the app use, and where is the app source (private repo)?
- Should the rabbinic data apply to non-Hebrew multilingual models when
  Hebrew is selected?

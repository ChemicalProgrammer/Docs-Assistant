# Project structure

- `Code.gs` — menu, sidebar/dialog bootstrap, document context.
- `Formatting.gs` — styles, paragraph formatting, lists, breaks.
- `GeminiService.gs` — Gemini calls, quick actions, insertion/replacement.
- `SettingsService.gs` — per-user API key storage.
- `Sidebar.html` — main UI.
- `Settings.html` — API-key settings dialog.
- `Styles.html` — Material/Workspace-inspired visual layer.
- `Scripts.html` — client-side sidebar behavior.
- `appsscript.json` — manifest/scopes.
- `README.md` — installation and current limitations.

The separation is deliberate: deterministic document manipulation stays independent from Gemini.

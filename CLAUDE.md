# Word Document Assistant (NVDA add-on) — guide for Claude

## About the project and the owner

- NVDA add-on that lets blind users read, navigate, manage and format Microsoft Word documents from the keyboard.
- Owner and author: **Peera Pilarit** (`addon_author="Peera Pilarit <9peera@gmail.com>"`). Do not add a company name anywhere.
- The owner writes in Thai. Reply in Thai, keep code, identifiers and code comments in English.
- Public on GitHub (9Peera/word-document-assistant) and submitted to the NVDA Add-on Store, stable channel. Releases are built by GitHub Actions when a tag such as `v1.0.0` is pushed.
- Built from the NV Access AddonTemplate (new template, uv/ruff/pyright/scons).

## Targets

- NVDA 2026.1 and later (`minimumNVDAVersion = 2026.1`, `lastTestedNVDAVersion = 2026.1.1`). NVDA 2026.1 is **64-bit, Python 3.13**: never bundle 32-bit DLLs or compiled libraries.
- Microsoft Word: Microsoft 365, Office 2021, Office 2019 (32-bit and 64-bit Office, UIA on and off).
- Word is driven through its **object model over COM** (`wordCom.py`), not through UIA.

## Commands

```bash
python3 -m pytest -q                      # unit tests (fake Word object model, no NVDA needed)
ruff check addon && ruff format addon     # lint and format (tabs, line length 110)
python3 devtools/updateTranslations.py    # refresh the Thai .po after changing strings
scons                                     # build wordDocAssistant-<version>.nvda-addon
```

Real testing needs Windows with NVDA and Word. Say clearly when something was only tested with the fake Word.

## Layout (`addon/globalPlugins/wordDocAssistant/`)

| File | Purpose |
|---|---|
| `__init__.py` | GlobalPlugin: navigation mode, key handling, host methods (dialogs, AI flows, tools menus) |
| `registry.py` | All element types in left/right arrow order; each letter must be unique (asserted) |
| `elements.py` | Base `ElementType`, `Item`, `Action`, `Host` protocol; spelling, grammar, comments, notes, revisions |
| `structure.py` | Headings, tables, links, bookmarks, fields, cross-references, citations |
| `a11y.py` | Accessibility checker, pictures and alternative text, language problems |
| `thaidict.py` / `thaicheck.py` / `wordListUpdates.py` | Offline Thai word check, word lists, update checks |
| `ai.py` | AI providers (Gemini, Anthropic Claude, OpenAI, OpenAI-compatible, Ollama), model choice, retries, all AI tasks |
| `keyStore.py` / `keySetup.py` | API keys encrypted with Windows DPAPI; guided key setup dialog |
| `presets.py` / `formatDialog.py` | Document format presets and AI formatting |
| `tools.py` | Thai digit and era tools, result and compare dialogs |
| `dialogs.py`, `settings.py`, `sounds.py`, `wordCom.py` | Shared dialogs, settings panel and config spec, earcons, COM helpers |
| `data/` | Bundled word lists (PyThaiNLP CC0) and the misspelling list; see `data/README.md` for licences |

Tests live in `tests/`: `nvdaStubs.py` stands in for NVDA modules, `fakeword.py` is a small fake Word object model, and `fakeword.Host` records what actions ask of the add-on.

## Keys (keep this table in sync with the help text in `__init__.py` and both readmes)

- `NVDA+Shift+F8` toggles navigation mode (only captured in Word). `NVDA+Control+Shift+F8` opens AI setup.
- Element letters (Shift = previous): S spelling, G grammar, C comments, F footnotes, E endnotes, R revisions, O headings, U tables, M pictures, K links, X cross-references, I citations, B bookmarks, D fields, L language, A accessibility, Z Thai word check.
- Command letters: T track changes (Shift+T toggles), N note at caret, P position, Q summary, W AI proofread, J AI formatting, V AI tools, Y Thai tools, H help.
- In the mode: arrows move and change type, Home/End, Space list, Applications/Shift+F10 actions, Ctrl+C / Ctrl+Shift+C copy, Enter leaves ready to edit, Escape leaves, Backspace returns to the start.
- All 26 letters are now used in the mode. A new type or command needs a key with a modifier, a key outside the mode, or a design change agreed with the owner.

## Rules that matter

1. **Every document change is one undo step**: wrap it in `with undoRecord(ctx, _("..."))`, call `host.modified()` afterwards, and speak the result with `host.speak`.
2. **AI never changes the document by itself.** It suggests; the user checks and chooses (choice dialog, editable prompt, compare dialog, format preview). Keep it that way.
3. Catch Word failures with `COM_ERRORS`; never let a COM error reach NVDA.
4. All user-facing text goes through `_()` with a `# Translators:` comment, then run `devtools/updateTranslations.py` and translate the new entries in `addon/locale/th/LC_MESSAGES/nvda.po` (natural Thai, keep `{placeholders}` and `&` accelerators).
5. Network work runs in background threads (`ai.runInBackground`, `wordListUpdates._background`); UI work on the main thread (`wx.CallAfter`).
6. Respect `globalVars.appArgs.secure`: no settings panel, no network, no writes on secure screens.
7. Only use word lists and data whose licence allows reuse. The Royal Institute (ราชบัณฑิตยสถาน) dictionary is copyrighted and not open data: do not scrape or bundle it.
8. Never put API keys or other secrets in the repository or in tests.

## Releasing a version

1. Bump `addon_version` in `buildVars.py`; update `addon_changelog` and, when features change, `addon_description` (keep the trailing commas — `tests/test_build.py` checks the file parses).
2. Add the Thai translations of the new description and changelog to `nvda.po`.
3. Update `changelog.md`, `readme.md` (English) and `addon/doc/th/readme.md` (Thai).
4. `python3 -m pytest -q`, `ruff check addon`, `scons`, then check the package contents and both manifests.

## Status and next ideas

- Headings are collected with Find by heading style (built-in Heading 1-9 plus custom styles with an outline level), with GoTo only as a time-limited fallback. Do not go back to GoTo heading by heading: it made Word stop responding in long documents.
- Current version: 1.0.1 (1.0.0 was the first public release; 1.0.1 makes heading collection fast, the owner found it hanging in real Word). Add-on Store stable channel. Everything is unit tested with the fake Word only; testing with real Word/NVDA (test matrix above) is the top open task.
- Ideas the owner has agreed are worth doing: ask the Royal Institute for permission to use their data; an organization misspelling list hosted on GitHub (supported via the `misspellingListUrl` setting); word-frequency ranking of suggestions (TNC frequency list, CC0); AI filtering of names in unknown words; checking while typing; a preset editor.

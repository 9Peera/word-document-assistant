# Changelog

## 1.0.2

* Long documents no longer freeze NVDA. Every search for items stops at a time limit (setting "Longest search time in seconds", 10 by default) and reports how many it found so far.
* Spelling and grammar errors are asked from Word part by part (about 20,000 characters at a time) instead of for the whole document at once.
* Language problems, empty paragraphs used for spacing and skipped heading levels are found with Word's own search instead of reading every paragraph.
* The Thai word check reads the text in large parts instead of paragraph by paragraph.
* Comments, notes, tracked changes, tables, links, bookmarks, fields and pictures are read with the collection's enumerator, which stays fast in long collections.
* Left and right arrows only say the number of items when Word can count them quickly; spelling, grammar and other slow types are searched when you move to an item.

## 1.0.1

* Headings (O) are collected with Word's search by heading style, which is much faster in long documents. Before, Word could stop responding and the add-on reported "Word did not respond, please try again".
* Custom paragraph styles with an outline level are found as headings too. Looking for headings never takes more than a few seconds.

## 1.0.0

* First public release, for the NVDA Add-on Store.
* The weekly check for new Thai word lists is off until the user allows it; the add-on asks once, the first time the Thai word check (Z) is used.
* Newest PyThaiNLP word list; project and source addresses; privacy section in the documentation.

## 0.10.0 (internal)

* Offline Thai word check (Z): commonly misspelled words with the right form, and words not in the word list with suggestions or the completed word.
* Open word lists: PyThaiNLP (CC0, included) and the LibreOffice NECTEC dictionary (LGPL, optional download); my words, my misspellings and an organization list from a web address.
* Weekly check for new word lists with a notice, NVDA Tools menu item, optional automatic download; downloads are checked before use.
* Look up a word on the Royal Institute dictionary website.

## 0.9.0 (internal)

Stability (0.6)
* API keys are encrypted with Windows DPAPI for the current user; keys saved by earlier versions are encrypted on start.
* Modern comments in Microsoft 365: when Word refuses a reply or an edit, the text is copied and the comment is opened in Word.
* "Collecting..." is spoken while long lists such as spelling errors are gathered.
* Unit tests (pytest, 50+ tests with a fake Word object model) run on GitHub Actions.

Structure (0.7)
* New element types: headings (O), tables (U), links (K), bookmarks (B), fields (D), cross-references (X) and citations (I), each with actions.
* Insert cross-references and citations, add sources, insert or update the bibliography, change the citation style.

Accessibility (0.8)
* Accessibility checker (A) with a fix for each problem.
* Pictures (M): edit alternative text, let AI describe the picture, mark as decorative.
* Language problems (L): find and fix Thai and English text marked with the wrong language.
* AI services accept pictures (Gemini, Claude, OpenAI, OpenAI compatible and Ollama vision models).

Tools (0.9)
* AI tools (V): summaries, rewriting, translation, comment summary, consistency check, shown next to the original.
* Thai tools (Y): Thai and Arabic digits, Buddhist and Christian era years.

## 0.5.1 (internal)

* Author changed to Peera Pilarit.
* Full list of main features in the add-on description shown in the Add-on Store and the installation dialog.

## 0.5.0 (internal)

* Format the document with AI (J in navigation mode): 10 formats including Thai official letter, memo,
  announcement, business and general letters, report, article, academic paper, meeting minutes and an accessible format.
  AI recommends a format and decides the role of each paragraph; the words are not changed.
  A preview lets you correct every role before formatting, and Control+Z undoes it in one step.

## 0.4.0 (internal)

* Paid AI services: Anthropic Claude and OpenAI, besides Google Gemini (free or paid), OpenAI compatible services and Ollama.
* Every service can choose its model automatically, with a Find models button to pick one, and moves to another model when one is busy.
* Each service keeps its own API key.

## 0.3.3 (internal)

* When the AI service is busy (error 500, 502, 503, 504) the request is tried again after 2 and 5 seconds.
* With automatic model choice, a busy or out-of-quota Gemini model is replaced by the next best model the key can use.

## 0.3.2 (internal)

* Fixed error 404 with Google Gemini: when the model field is empty, the add-on asks Google which models the key can use and picks the newest Flash model. It chooses again if that model is retired.
* Works with the new Google keys that start with AQ. A clear message is given if Google does not accept that kind of key for the account.

## 0.3.1 (internal)

* Fixed: "Paste key from clipboard and test" now tests the key typed in the field when the clipboard has no key.
* API keys are no longer required to start with AIza; spaces, line breaks and quotes copied with a key are removed.

## 0.3.0 (internal)

* AI setup dialog (NVDA+Control+Shift+F8, also in the settings): opens the free Google AI Studio key page,
  picks up the copied key from the clipboard, tests and saves it. Opens by itself when the key is missing or rejected.
* Ollama discovery: finds Ollama on this computer and lists its installed models.

## 0.2.1 (internal)

* Fixed: NVDA+Shift+F8 reported "Move the focus to the document text first" in Word documents.
  The Word document window is now found reliably.

## 0.2.0 (internal)

* Navigation mode key changed to NVDA+Shift+F8.
* Optional AI suggestions (Google Gemini free tier, OpenAI compatible services, or Ollama offline):
  W checks the current paragraph, Shift+W the text selected before the mode started,
  AI suggestions for a misspelled word, and AI drafts of comment replies.
  Nothing changes until the user chooses a suggestion.

## 0.1.0 (internal)

* Document navigation mode on NVDA+Shift+F8 for Microsoft Word.
* Spelling errors, grammar issues, comments, footnotes, endnotes and tracked changes:
  move with letters (Shift for previous), arrows, Home and End.
* Item list with separate detail fields and copy buttons.
* Actions: replace misspelled words, reply to, edit, resolve and delete comments,
  edit, insert, delete and convert notes, accept and reject revisions.
* Track Changes status and toggle, note at caret, position and document summary.
* Thai translation.

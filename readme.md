# Word Document Assistant

* Author: Peera Pilarit
* NVDA compatibility: 2026.1 and later
* Microsoft Word: Microsoft 365, Office 2021, Office 2019

Word Document Assistant lets you move through, read and manage review items in Microsoft Word without leaving the document:
spelling errors, grammar issues, comments, footnotes, endnotes and tracked changes.

## Language support

The add-on works with Word documents in any language, and its interface can be translated.

| Works with any language | Extras for Thai users |
|---|---|
| Navigation mode, item lists and actions | Thai word list check: words not in the Thai word list (part of Z) |
| Spelling and grammar (S, G) with Word's own checker for the document's language | Buddhist era (พ.ศ.) and Christian era (ค.ศ.) years (part of Y) |
| Comments, notes, tracked changes, headings, tables, links, fields, cross-references, citations | Thai official document formats: official letter, memo, announcement, meeting minutes (part of J) |
| Accessibility checker (A) and language check (L) for many writing systems | Thai justification and Thai fonts when the text has Thai |
| Digits of many scripts (Y) and lists of commonly misspelled words in any language (Z) | |
| AI proofreading, summaries, rewriting and translation into more than 50 languages | |
| Formats such as APA, MLA, Harvard, report, article, letters and an accessible format (J) | |

The interface is in English and Thai. To translate it into your language, see "Translating the add-on" at the end.

## Navigation mode

Press NVDA+Shift+F8 in a Word document to enter navigation mode. You hear a rising sound and the current element type.
Press NVDA+Shift+F8 again, Escape or Enter to leave.

While you navigate, the real Word caret moves to each item.
When you leave the mode the caret stays at the last item you worked on, so you do not lose your place.

### Moving to an element type

Press the letter to move to the next item of that type. Add Shift to move to the previous one.

| Key | Element type |
|---|---|
| S | Spelling errors |
| G | Grammar issues |
| C | Comments |
| F | Footnotes |
| E | Endnotes |
| R | Revisions (tracked changes) |
| O | Headings |
| U | Tables |
| M | Pictures and their alternative text |
| K | Links |
| X | Cross-references |
| I | Citations |
| B | Bookmarks |
| D | Fields |
| L | Text marked with the language of another writing system |
| A | Accessibility problems |
| Z | Word list check: commonly misspelled words in any language, and for Thai words not in the word list |

### Keys inside the mode

| Key | Action |
|---|---|
| Down arrow / Up arrow | Next or previous item of the current type |
| Left arrow / Right arrow | Change the element type. The number of items is said when Word can count them quickly; types that need a search of the whole document, such as spelling, are searched when you move to an item |
| Home / End | First or last item |
| Space | List of all items of the current type |
| Applications or Shift+F10 | Actions for the current item |
| Control+C | Copy the text of the current item |
| Control+Shift+C | Copy all details of the current item |
| T | Report Track Changes status. Shift+T turns Track Changes on or off |
| N | Read the footnote or endnote at the caret |
| P | Report page, section, line and heading |
| Q | Document summary |
| W | AI check of the current paragraph. Shift+W checks the text that was selected before the mode started |
| J | Format the document with AI |
| V | AI tools: summarize, rewrite, translate, summarize comments, check consistency |
| Y | Numbers and dates: digits of many scripts, Buddhist and Christian era years for Thai |
| H | Help |

### Leaving the mode

| Key | Result |
|---|---|
| Enter | Leave at the current item, ready to edit. A misspelled word is selected so typing replaces it. For footnotes and endnotes the caret moves into the note text. |
| Escape | Leave. The caret stays at the current item. |
| Backspace | Go back to where the caret was when the mode started, and leave. |
| Any other key | Leave the mode. The key does its normal job in Word. |

Letters and digits that have no command are not typed into the document while the mode is on; you hear an error sound instead.

The mode reads keys by their position, so it works the same with any keyboard layout, for example Thai, Russian or Arabic.

## The item list

Press Space in navigation mode to open the list of all items of the current type.

* Type in the Filter field to show only matching items.
* Each detail has its own read-only field. For a comment these are Author, Date, Status, Commented text, Comment and Replies,
  so each part can be read and copied on its own.
* Go to moves the caret to the item. You stay in navigation mode.
* Actions opens the commands for the item. The Applications key on the list does the same.
* Copy text copies the main text. Copy all details copies every field with its label.

## Actions

| Element type | Actions |
|---|---|
| Spelling errors | Replace with a suggestion, type a replacement, copy |
| Grammar issues | Copy |
| Comments | Reply, edit, resolve or reopen, delete, copy the comment or the commented text, new comment at the starting position |
| Footnotes and endnotes | Edit note text, delete, convert between footnote and endnote, insert a new note at the starting position, copy |
| Revisions | Accept, reject, accept or reject all by the same author, accept or reject all in the document, copy |

"Starting position" is where the caret was, or the text that was selected, when you pressed NVDA+Shift+F8.
Each change is one Word undo step, so Control+Z undoes it.

## AI suggestions (optional)

AI can suggest spelling and wording corrections, correct spellings for a misspelled word, and drafts of replies to comments.
**AI only suggests. Nothing in the document changes until you choose a suggestion yourself.**

* W in navigation mode checks the current paragraph. Shift+W checks the text you selected before pressing NVDA+Shift+F8.
  A dialog lists each correction as "original to suggestion", with the original, the suggestion and the reason in separate fields.
  Check the ones you want with the Space bar and press Apply checked. All applied corrections are one undo step.
* On a misspelled word, the actions menu has "AI suggestions...". Pick one to replace the word.
* On a comment, the actions menu has "Draft a reply with AI...". Pick a draft, edit it, then confirm to post it.

Turn AI on in the settings and choose a service:

| Service | Cost | Where the text goes |
|---|---|---|
| Google Gemini (the model is chosen automatically when the model field is empty) | Free tier with a free API key from Google AI Studio | Google. On the free tier Google may use the content to improve its products |
| OpenAI compatible service | Depends on the service (several have free tiers) | That service |
| Anthropic Claude | Paid, per use. Create the key and set up billing at console.anthropic.com | Anthropic |
| OpenAI | Paid, per use. Create the key and set up billing at platform.openai.com | OpenAI |
| Ollama on this computer | Free | Stays on this computer. Recommended for confidential documents |

### Getting a free API key

Press NVDA+Control+Shift+F8 anywhere (or Get or update the API key in the settings) to open the AI setup.
The same dialog opens by itself the first time you use AI, and whenever the key is missing or rejected, so the key can be updated at any time.

1. Press Get a free API key. Google AI Studio opens in your browser.
2. Sign in with your own Google account, create a key and copy it.
3. Come back to the dialog. The key is found on the clipboard, tested and filled in. (Paste key from clipboard and test does the same by hand.)
4. Press Save and turn on AI.

The add-on cannot get a key for you: a key belongs to your own account, and keys found on the internet belong to other people.

For Ollama, choose it in the dialog, press Get Ollama for free to download it, then Find Ollama on this computer. The installed models are listed and one is selected. No key is needed.

Before text is first sent to an online service, the add-on asks for your agreement.
The API key is stored in the NVDA configuration on this computer, encrypted for your Windows user account.

## Accessibility of the document

Press A in navigation mode to move through accessibility problems; Space lists them all, and the Applications key offers the fix for each.

| Problem | Fix |
|---|---|
| Picture without alternative text | Type it, let AI write a short or detailed description to check, or mark the picture as decorative |
| Table without a header row | Make the first row a header row |
| Heading level skipped | Change to the expected level |
| Empty paragraphs used for spacing | Remove them and add space after the paragraph instead |
| Link text such as "click here" | Change the link text |
| Document has no title | Set the title (the first heading is suggested) |
| Long document without headings | Format the document with AI |
| Text marked with the wrong language | Give every word the language of its writing system |

M lists every picture with its alternative text. AI reads the picture itself (PNG, JPEG, GIF, BMP, TIFF; drawings such as EMF must be saved as a picture first) together with the nearby text, and you edit the description before it is saved.

### Language check (L)

Word stores a language on every piece of text. NVDA speaks text with the voice of that language, and Word checks its spelling with that language's dictionary. Text copied from elsewhere often keeps the wrong language: Thai text marked as English, Russian text marked as German, English text marked as Arabic.

L finds paragraphs where text of one writing system is marked with the language of another writing system, or with no language. It knows Latin, Thai, Lao, Khmer, Burmese, Arabic (Arabic, Persian, Urdu), Hebrew, Cyrillic, Greek, Devanagari, Bengali, Gurmukhi, Gujarati, Tamil, Telugu, Kannada, Malayalam, Sinhala, Georgian, Armenian and Ethiopic. Chinese, Japanese and Korean are not checked, because Word marks their language reliably by itself. A language the add-on does not know is never reported, so unusual but correct choices stay.

The actions fix one paragraph, mark a whole paragraph with one language, or fix the whole document. When fixing, each word gets the language of its writing system: in the settings, check the languages you write in (for example Ukrainian rather than Russian for Cyrillic text, or Persian rather than Arabic). Without a choice, Thai text becomes Thai, Latin-script text gets the language of the Normal style, and other scripts get their main language.

## Document structure

| Key | Actions |
|---|---|
| O headings | Promote, demote, change to any level or to body text |
| U tables | Header row on or off, alternative text typed or written by AI; Enter puts the caret in the first cell |
| K links | Change text or address, open, remove the link but keep the text; internal links to missing bookmarks are reported as broken |
| B bookmarks | Add at the starting position, delete |
| D fields | Update one or all fields and tables of contents, copy the code, convert to plain text; fields showing an error are reported |
| X cross-references | Shows what each points to and reports broken ones; insert a new one: choose heading, bookmark, footnote, endnote, numbered item or caption, then the item and what to show |
| I citations | Insert a citation from the document's sources or a new source, add sources, insert or update the bibliography, change the citation style |

## AI tools

Press V in navigation mode. Summaries and the consistency check use the text selected before NVDA+Shift+F8, or the whole document. Rewriting and translation use the selected text or the current paragraph.

* Summaries: short, section by section, or key points with dates and actions requested.
* Rewrite in the formal official style of the text's language, more politely, shorter and clearer, or in easy-to-read language; or correct spelling and grammar only.
* Translate to NVDA's language, to English, or to any of more than 50 other languages.
* Summarize all comments, with the requests that are still open.
* Check consistency: spellings, local digits mixed with 0-9, years of different eras, dates and numbering.

AI answers in the language of the text. To always get answers in one language, choose it in the settings ("AI answers in language").

Rewrites and translations open next to the original. You can edit the new text, then replace the original, insert the new text after it (translations), copy it, or close without changes.

## Numbers and dates

Press Y in navigation mode, for the selected text or the whole document. The menu only offers what the text needs and says how many digits will change:

* Digits of another script to 0-9: Thai, Lao, Khmer, Burmese, Arabic-Indic, Persian and Urdu, Devanagari, Bengali, Gurmukhi, Gujarati, Tamil, Telugu, Kannada and Malayalam digits.
* 0-9 to the digits of a script the text is written in (for example Thai digits for Thai text, Arabic-Indic digits for Arabic text), or to any other digits with "0-9 to other digits".
* For Thai: years written with ค.ศ. to พ.ศ. (adds 543), and back. Only years written with the era are changed, so other numbers stay as they are.

Each tool is one undo step.

## Word list check (offline)

Press Z in navigation mode. It works without AI or internet:

* **Commonly misspelled words in any language**, with the right form: a starter English list (recieve, seperate, definately...) and a Thai list (อนุญาติ, อีเมล์, นะค่ะ...) written for the add-on, your own list, and an organization list. For languages written with spaces only whole words match, and a word at the start of a sentence is found too. Use the action "Edit my list of commonly misspelled words" to add pairs for your language: wrong form, Tab, right form, Tab, optional note.
* **For Thai: words not in the Thai word list**, with suggestions one letter or mark away, or the complete word when it was typed only partly (คอมพิวเตอ: คอมพิวเตอร์). Thai is written without spaces between words and Word's own Thai checking is weak, which is why this part exists for Thai. Names after นาย, นาง, คุณ, ดร. and similar are not reported. This part can be turned off in the settings.

Spelling of other languages is checked by Word itself (S), with the dictionary of the document's language.

Actions: replace with a suggestion, replace every occurrence, type a replacement, add the word to my words, ignore it until NVDA restarts, edit my list, look up a Thai word on the Royal Institute dictionary website (the word is copied for you), and update the word lists.

### Word lists

| List | Licence | How it is used |
|---|---|---|
| PyThaiNLP Thai word list (62,000 words) | CC0 1.0 | Included in the add-on; newer versions are downloaded |
| LibreOffice Thai dictionary (NECTEC LEXiTRON, 50,000 words) | LGPL | Downloaded when turned on in the settings |
| Commonly misspelled English and Thai words of the add-on | GPL v2 | Included |
| My words, my misspellings | Yours | Files in the NVDA configuration folder; open them from the settings |
| Organization list of commonly misspelled words | Yours | Downloaded from the address set in the settings, for example a file on your own web site or GitHub |

The Royal Institute dictionary itself is not published as open data and is protected by copyright, so the add-on does not download it; the look-up action opens its website instead.
The PyThaiNLP list is collected from real text and also contains common misspellings; the add-on removes every wrong form in the misspelling lists before using it.

### New word lists

If you allow it (you are asked the first time you press Z, or turn it on in the settings), once a week, a minute after NVDA starts, the add-on asks the list websites whether there is a newer version (only a small header request).
When there is, it tells you, and you update with NVDA menu, Tools, Update Thai word lists, with the actions of Z, or with Update word lists now in the settings. You can also let it download new lists automatically.
A downloaded file is only used when it really is a Thai word list.

## Formatting a document with AI

Press J in navigation mode.

1. A dialog lists the formats. AI reads the beginning of the document and recommends one; you can choose another.
2. Choose the whole document or only the text selected before NVDA+Shift+F8, and whether to set paper and margins, fonts, and remove empty paragraphs used for spacing.
3. Press Analyse and preview. AI decides the role of every paragraph (subject, recipient, body, closing, signature, heading and so on). The words are never changed.
4. The preview lists every paragraph with its role. Change any role with the Role box.
5. Press Format now. The whole formatting is one undo step: Control+Z undoes it.

| Format | Used for |
|---|---|
| APA 7th edition paper | Papers in APA style: double spacing, bold headings by level, references with a hanging indent |
| MLA 9th edition paper | Papers in MLA style: heading block at the top left, plain centred title, Works Cited |
| Harvard style paper | Papers with Harvard referencing: 1.5 line spacing, space between paragraphs |
| Thai official letter | หนังสือภายนอก: ที่, ส่วนราชการ, วันที่, เรื่อง, เรียน, อ้างถึง, สิ่งที่ส่งมาด้วย, body, closing, signature, contact |
| Internal memo | บันทึกข้อความ |
| Announcement or order | ประกาศ, คำสั่ง |
| Formal or business letter | Letters to companies and organizations |
| General letter | Personal or informal letters |
| Formal report | Reports with real Word headings |
| Article | Articles and blog posts |
| Academic paper | Abstract, keywords, headings, captions, bibliography |
| Meeting minutes | รายงานการประชุม with agenda items as headings |
| Easy to read and accessible | Left aligned, larger text, real headings, no empty lines |

Lists keep their numbering and paragraphs inside tables are not re-indented.
Thai fonts and Thai justification are only used when the text has Thai; in other documents the font for Arabic, Hebrew, Hindi and other scripts is left as it is. The APA, MLA and Harvard formats keep your paper size (Letter or A4).
The official measurements follow common practice; check them against your organization's rules. They can be changed in presets.py.

## Privacy

The add-on only uses the internet in these cases:

* **AI features, only when you turn AI on.** The text you check, rewrite, translate or summarize, and pictures you ask to describe, are sent to the AI service you chose (Google Gemini, Anthropic Claude, OpenAI, an OpenAI compatible service, or Ollama, which stays on your computer). You are asked before text is first sent to an online service. Free tiers may use the content to improve the service.
* **Thai word lists.** Only if you allow it, once a week the add-on asks GitHub (and the organization address, if you set one) whether a newer word list exists; this sends no document text. You are asked once, the first time the word list check (Z) runs on a document with Thai text. Lists are only downloaded when you agree, or automatically if you turned that on. Turn the weekly check off in the settings.
* **Links you open** and the Royal Institute dictionary website, only when you choose them.

Your API keys are stored in the NVDA configuration, encrypted for your Windows user account. The add-on collects no usage data.

## Settings

NVDA menu, Preferences, Settings, Word Document Assistant:

* Play sounds for navigation mode and wrapping
* Report item position, for example 3 of 12
* Report the author, and the date, of comments and revisions
* Speak the text around misspelled words, and how many characters on each side
* Maximum items collected for each type (for very long documents)
* Longest search time in seconds (default 10). In a very long document a search stops at this limit and reports how many items it found so far, so NVDA never waits a long time for Word
* Languages I write in: used by the language check (L) to choose the language of each writing system
* AI suggestions: on or off, service, API key, service address, model, the language AI answers in (the language of the text, or a fixed language), and a connection test
* Word list check: for Thai also report words not in the word list, weekly check for new Thai word lists, the LibreOffice Thai dictionary, and an organization list of commonly misspelled words

The NVDA+Shift+F8 command can be changed in Input Gestures, category Word Document Assistant.

## Notes

* Navigation covers the main text of the document.
* In recent Microsoft 365 builds, some grammar suggestions are only shown by the Editor pane and are not reported to add-ons.
* Documents opened in Protected View must have editing enabled first.

## Translating the add-on

The interface and this documentation can be translated into any language.

* Every release on GitHub has the file `wordDocAssistant.pot` with all the messages. Translate it with Poedit or any gettext editor and send the `.po` file (and a translated readme if you like) as a GitHub issue or pull request at https://github.com/9Peera/word-document-assistant.
* Keep the `{placeholders}` and the `&` shortcut letters, and test that the shortcut letters of each dialog are different.

## Working with other add-ons

Word Document Assistant only acts while its navigation mode is on, and its command keys are NVDA+Shift+F8 and NVDA+Control+Shift+F8, so it can be used next to other add-ons for Microsoft Office. Authors of other Office add-ons are welcome to reuse ideas or code (GPL v2), for example the accessibility checker or the time-limited collection of items in long documents.

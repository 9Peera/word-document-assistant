# Word Document Assistant

* Author: Peera Pilarit
* NVDA compatibility: 2026.1 and later
* Microsoft Word: Microsoft 365, Office 2021, Office 2019

Word Document Assistant lets you move through, read and manage review items in Microsoft Word without leaving the document:
spelling errors, grammar issues, comments, footnotes, endnotes and tracked changes.

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
| L | Text marked with the wrong language |
| A | Accessibility problems |
| Z | Thai word check with the offline word lists |

### Keys inside the mode

| Key | Action |
|---|---|
| Down arrow / Up arrow | Next or previous item of the current type |
| Left arrow / Right arrow | Change the element type |
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
| Y | Thai tools: Thai and Arabic digits, Buddhist and Christian era years |
| H | Help |

### Leaving the mode

| Key | Result |
|---|---|
| Enter | Leave at the current item, ready to edit. A misspelled word is selected so typing replaces it. For footnotes and endnotes the caret moves into the note text. |
| Escape | Leave. The caret stays at the current item. |
| Backspace | Go back to where the caret was when the mode started, and leave. |
| Any other key | Leave the mode. The key does its normal job in Word. |

Letters and digits that have no command are not typed into the document while the mode is on; you hear an error sound instead.

The mode reads keys by their position, so it works the same with the Thai and English keyboard layouts.

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
| Text marked with the wrong language | Mark Thai words as Thai and English words as English |

M lists every picture with its alternative text. AI reads the picture itself (PNG, JPEG, GIF, BMP, TIFF; drawings such as EMF must be saved as a picture first) together with the nearby text, and you edit the description before it is saved.
L lists paragraphs whose language is wrong, which stops Thai spell checking and makes NVDA use the wrong voice; one action fixes the whole document.

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
* Rewrite in official Thai style, more politely, shorter and clearer, or in easy-to-read language; or correct spelling and grammar only.
* Translate to English or to Thai.
* Summarize all comments, with the requests that are still open.
* Check consistency: spellings, Thai and Arabic digits, era years, dates and numbering.

Rewrites and translations open next to the original. You can edit the new text, then replace the original, insert the new text after it (translations), copy it, or close without changes.

## Thai tools

Press Y in navigation mode, for the selected text or the whole document:

* Arabic digits to Thai digits, and back. The menu says how many digits will change.
* Years written with ค.ศ. to พ.ศ. (adds 543), and back. Only years written with the era are changed, so other numbers stay as they are.

Each tool is one undo step.

## Thai word check (offline)

Press Z in navigation mode. It works without AI or internet, with open word lists:

* **Commonly misspelled words** such as อนุญาติ, อีเมล์ or นะค่ะ, with the right form. The add-on comes with a starter list of about 65 words, written for it from well known Royal Institute spellings.
* **Words not in the word list**, with suggestions one letter or mark away, or the complete word when it was typed only partly (คอมพิวเตอ: คอมพิวเตอร์). Names after นาย, นาง, คุณ, ดร. and similar are not reported. This part can be turned off in the settings.

Actions: replace with a suggestion, replace every occurrence, type a replacement, add the word to my words, ignore it until NVDA restarts, look it up on the Royal Institute dictionary website (the word is copied for you), and update the word lists.

### Word lists

| List | Licence | How it is used |
|---|---|---|
| PyThaiNLP Thai word list (62,000 words) | CC0 1.0 | Included in the add-on; newer versions are downloaded |
| LibreOffice Thai dictionary (NECTEC LEXiTRON, 50,000 words) | LGPL | Downloaded when turned on in the settings |
| Commonly misspelled words of the add-on | GPL v2 | Included |
| My words, my misspellings | Yours | Files in the NVDA configuration folder; open them from the settings |
| Organization list of commonly misspelled words | Yours | Downloaded from the address set in the settings, for example a file on your own web site or GitHub |

The Royal Institute dictionary itself is not published as open data and is protected by copyright, so the add-on does not download it; the look-up action opens its website instead.
The PyThaiNLP list is collected from real text and also contains common misspellings; the add-on removes every wrong form in the misspelling lists before using it.

### New word lists

Once a week, a minute after NVDA starts, the add-on asks the list websites whether there is a newer version (only a small header request).
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
The official measurements follow common practice; check them against your organization's rules. They can be changed in presets.py.

## Settings

NVDA menu, Preferences, Settings, Word Document Assistant:

* Play sounds for navigation mode and wrapping
* Report item position, for example 3 of 12
* Report the author, and the date, of comments and revisions
* Speak the text around misspelled words, and how many characters on each side
* Maximum items collected for each type (for very long documents)
* AI suggestions: on or off, service, API key, service address, model, document language, and a connection test

The NVDA+Shift+F8 command can be changed in Input Gestures, category Word Document Assistant.

## Notes

* Navigation covers the main text of the document.
* In recent Microsoft 365 builds, some grammar suggestions are only shown by the Editor pane and are not reported to add-ons.
* Documents opened in Protected View must have editing enabled first.

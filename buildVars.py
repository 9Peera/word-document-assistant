# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.

from site_scons.site_tools.NVDATool.typings import AddonInfo, BrailleTables, SymbolDictionaries, SpeechDictionaries

# Since some strings in `addon_info` are translatable,
# we need to include them in the .po files.
# Gettext recognizes only strings given as parameters to the `_` function.
# To avoid initializing translations in this module we simply import a "fake" `_` function
# which returns whatever is given to it as an argument.
from site_scons.site_tools.NVDATool.utils import _

# Add-on information variables
addon_info = AddonInfo(
	# add-on Name/identifier, internal for NVDA
	addon_name="wordDocAssistant",
	# Add-on summary/title, usually the user visible name of the add-on
	# Translators: Summary/title for this add-on
	# to be shown on installation and add-on information found in add-on store
	addon_summary=_("Word Document Assistant"),
	# Add-on description
	# Translators: Long description to be shown for this add-on on add-on information from add-on store
	addon_description=_("""Read, manage and format Microsoft Word documents from the keyboard with NVDA. Works with Microsoft 365, Office 2021 and Office 2019.
Main features:
- Document navigation mode (NVDA+Shift+F8): move to spelling errors (S), grammar issues (G), comments (C), footnotes (F), endnotes (E), tracked changes (R), headings (O), tables (U), pictures (M), links (K), cross-references (X), citations (I), bookmarks (B), fields (D), language problems (L) and accessibility problems (A); Shift with the letter moves back, arrow keys move and change the type, Home and End go to the first and last item.
- When you leave the mode the caret stays at the last item, ready to edit; Enter leaves and lets you type straight away, Backspace returns to where you started.
- Item lists (Space) with each detail in its own field: author, date, status, commented text, comment and replies; copy the text or all details.
- Actions (Applications key): replace misspelled words, reply to, edit, resolve and delete comments, insert, edit, delete and convert footnotes and endnotes, accept and reject tracked changes one by one, by author or all.
- Track Changes status and toggle (T, Shift+T), note at the caret (N), page, section and heading (P), document summary (Q).
- Accessibility checker (A): pictures without alternative text, tables without header rows, skipped heading levels, empty paragraphs used for spacing, unclear link text, missing title and wrong language, each with a fix. AI writes alternative text for pictures and tables in Thai or English for you to check.
- Structure: promote and demote headings, header rows and alternative text of tables, edit and check links, add bookmarks, update fields, insert cross-references, citations and the bibliography.
- Optional AI suggestions: proofread a paragraph or the selection (W), suggestions for a misspelled word, drafts of comment replies. AI tools (V): summaries, rewriting in official, polite, concise or easy-to-read style, Thai and English translation, comment summary and consistency check, always shown next to the original before anything changes.
- Thai tools (Y): Thai and Arabic digits, Buddhist and Christian era years.
- Offline Thai word check (Z): commonly misspelled words with the right form, and words not in open word lists (PyThaiNLP CC0, optional LibreOffice NECTEC dictionary), with suggestions, my words, an organization list, and a weekly check for new word lists.
- Format the document with AI (J): Thai official letter, memo, announcement or order, meeting minutes, business and general letters, report, article, academic paper and an accessible easy-to-read format. AI only decides the role of each paragraph; the words are not changed, you check a preview and Control+Z undoes it.
- AI services: Google Gemini (free tier), Anthropic Claude and OpenAI (paid), OpenAI compatible services, or Ollama on this computer for confidential documents. Guided API key setup (NVDA+Control+Shift+F8) with automatic model choice.
- Every change is one Word undo step. API keys are encrypted for the Windows user. Keys work with Thai and English keyboard layouts. Thai and English interface."""),
	# version
	addon_version="0.10.0",
	# Brief changelog for this version
	# Translators: what's new content for the add-on version to be shown in the add-on store
	addon_changelog=_("""Offline Thai word check (Z) with open word lists, a list of commonly misspelled words, my words, an organization list and notices when new word lists are available."""),
	# Author(s)
	addon_author="Peera Pilarit <9peera@gmail.com>",
	# URL for the add-on documentation support
	addon_url="https://github.com/9Peera/word-document-assistant",
	# URL for the add-on repository where the source code can be found
	addon_sourceURL="https://github.com/9Peera/word-document-assistant",
	# Documentation file name
	addon_docFileName="readme.html",
	# Minimum NVDA version supported (e.g. "2019.3.0", minor version is optional)
	addon_minimumNVDAVersion="2026.1",
	# Last NVDA version supported/tested (e.g. "2024.4.0", ideally more recent than minimum version)
	addon_lastTestedNVDAVersion="2026.1.1",
	# Add-on update channel (default is None, denoting stable releases,
	# and for development releases, use "dev".)
	# Do not change unless you know what you are doing!
	addon_updateChannel=None,
	# Add-on license such as GPL 2
	addon_license="GPL v2",
	# URL for the license document the ad-on is licensed under
	addon_licenseURL="https://www.gnu.org/licenses/old-licenses/gpl-2.0.html",
)

# Define the python files that are the sources of your add-on.
# You can either list every file (using ""/") as a path separator,
# or use glob expressions.
# For example to include all files with a ".py" extension from the "globalPlugins" dir of your add-on
# the list can be written as follows:
# pythonSources = ["addon/globalPlugins/*.py"]
# For more information on SCons Glob expressions please take a look at:
# https://scons.org/doc/production/HTML/scons-user/apd.html
pythonSources: list[str] = ["addon/globalPlugins/wordDocAssistant/*.py"]

# Files that contain strings for translation. Usually your python sources
i18nSources: list[str] = pythonSources + ["buildVars.py"]

# Files that will be ignored when building the nvda-addon file
# Paths are relative to the addon directory, not to the root directory of your addon sources.
# You can either list every file (using ""/") as a path separator,
# or use glob expressions.
excludedFiles: list[str] = ["**/__pycache__/**", "*.pyc"]

# Base language for the NVDA add-on
# If your add-on is written in a language other than english, modify this variable.
# For example, set baseLanguage to "es" if your add-on is primarily written in spanish.
# You must also edit .gitignore file to specify base language files to be ignored.
baseLanguage: str = "en"

# Markdown extensions for add-on documentation
# Most add-ons do not require additional Markdown extensions.
# If you need to add support for markup such as tables, fill out the below list.
# Extensions string must be of the form "markdown.extensions.extensionName"
# e.g. "markdown.extensions.tables" to add tables.
markdownExtensions: list[str] = ["markdown.extensions.tables"]

# Custom braille translation tables
# If your add-on includes custom braille tables (most will not), fill out this dictionary.
# Each key is a dictionary named according to braille table file name,
# with keys inside recording the following attributes:
# displayName (name of the table shown to users and translatable),
# contracted (contracted (True) or uncontracted (False) braille code),
# output (shown in output table list),
# input (shown in input table list).
brailleTables: BrailleTables = {}

# Custom speech symbol dictionaries
# Symbol dictionary files reside in the locale folder, e.g. `locale\en`, and are named `symbols-<name>.dic`.
# If your add-on includes custom speech symbol dictionaries (most will not), fill out this dictionary.
# Each key is the name of the dictionary,
# with keys inside recording the following attributes:
# displayName (name of the speech dictionary shown to users and translatable),
# mandatory (True when always enabled, False when not).
symbolDictionaries: SymbolDictionaries = {}

# Custom speech dictionaries (distinct from symbol dictionaries above)
# Speech dictionary files reside in the speechDicts folder and are named `name.dic`.
# If your add-on includes custom speech (pronunciation) dictionaries (most will not), fill out this dictionary.
# Each key is the name of the dictionary,
# with keys inside recording the following attributes:
# displayName (name of the speech dictionary shown to users and translatable),
# mandatory (True when always enabled, False when not).
speechDictionaries: SpeechDictionaries = {}

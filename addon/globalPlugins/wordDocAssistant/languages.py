# Word Document Assistant for NVDA
# Languages, writing systems (scripts) and digit systems shared by the language check, the number tools
# and the AI tools.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

"""Languages and scripts.

Word stores a language number (LCID) on every piece of text. A screen reader speaks text with the
voice of that language and Word checks its spelling with that language's dictionary, so text written
in one script but marked with the language of another script (Thai text marked as English, Russian
text marked as German...) is read and checked wrongly. This module knows which languages belong to
which script, so the add-on can find and fix such text for many languages, not only Thai.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from functools import cache

import addonHandler

addonHandler.initTranslation()

#: Word reports this when a range has more than one language.
WD_UNDEFINED = 9999999
LANG_NONE = 0
LANG_NO_PROOFING = 1024
LANG_THAI = 1054
LANG_ENGLISH_US = 1033
LANG_ENGLISH_UK = 2057


@dataclass(frozen=True)
class Language:
	code: str  # short code used in settings and for the AI, such as "th" or "en"
	lcid: int  # Word's language number
	script: str  # id of the script the language is written in
	englishName: str  # name sent to the AI


#: Languages the add-on knows by name, roughly by how widely they are used with Word.
LANGUAGES: tuple[Language, ...] = (
	Language("th", LANG_THAI, "thai", "Thai"),
	Language("en", LANG_ENGLISH_US, "latin", "English"),
	Language("en-GB", LANG_ENGLISH_UK, "latin", "British English"),
	Language("fr", 1036, "latin", "French"),
	Language("de", 1031, "latin", "German"),
	Language("es", 3082, "latin", "Spanish"),
	Language("it", 1040, "latin", "Italian"),
	Language("pt", 1046, "latin", "Portuguese"),
	Language("pt-PT", 2070, "latin", "European Portuguese"),
	Language("nl", 1043, "latin", "Dutch"),
	Language("sv", 1053, "latin", "Swedish"),
	Language("da", 1030, "latin", "Danish"),
	Language("nb", 1044, "latin", "Norwegian"),
	Language("fi", 1035, "latin", "Finnish"),
	Language("pl", 1045, "latin", "Polish"),
	Language("cs", 1029, "latin", "Czech"),
	Language("sk", 1051, "latin", "Slovak"),
	Language("hu", 1038, "latin", "Hungarian"),
	Language("ro", 1048, "latin", "Romanian"),
	Language("hr", 1050, "latin", "Croatian"),
	Language("sl", 1060, "latin", "Slovenian"),
	Language("tr", 1055, "latin", "Turkish"),
	Language("vi", 1066, "latin", "Vietnamese"),
	Language("id", 1057, "latin", "Indonesian"),
	Language("ms", 1086, "latin", "Malay"),
	Language("fil", 1124, "latin", "Filipino"),
	Language("ru", 1049, "cyrillic", "Russian"),
	Language("uk", 1058, "cyrillic", "Ukrainian"),
	Language("bg", 1026, "cyrillic", "Bulgarian"),
	Language("sr", 3098, "cyrillic", "Serbian"),
	Language("el", 1032, "greek", "Greek"),
	Language("ar", 1025, "arabic", "Arabic"),
	Language("fa", 1065, "arabic", "Persian"),
	Language("ur", 1056, "arabic", "Urdu"),
	Language("he", 1037, "hebrew", "Hebrew"),
	Language("hi", 1081, "devanagari", "Hindi"),
	Language("mr", 1102, "devanagari", "Marathi"),
	Language("ne", 1121, "devanagari", "Nepali"),
	Language("bn", 1093, "bengali", "Bengali"),
	Language("pa", 1094, "gurmukhi", "Punjabi"),
	Language("gu", 1095, "gujarati", "Gujarati"),
	Language("ta", 1097, "tamil", "Tamil"),
	Language("te", 1098, "telugu", "Telugu"),
	Language("kn", 1099, "kannada", "Kannada"),
	Language("ml", 1100, "malayalam", "Malayalam"),
	Language("si", 1115, "sinhala", "Sinhala"),
	Language("lo", 1108, "lao", "Lao"),
	Language("km", 1107, "khmer", "Khmer"),
	Language("my", 1109, "myanmar", "Burmese"),
	Language("ka", 1079, "georgian", "Georgian"),
	Language("hy", 1067, "armenian", "Armenian"),
	Language("am", 1118, "ethiopic", "Amharic"),
	Language("zh-CN", 2052, "cjk", "Simplified Chinese"),
	Language("zh-TW", 1028, "cjk", "Traditional Chinese"),
	Language("ja", 1041, "cjk", "Japanese"),
	Language("ko", 1042, "cjk", "Korean"),
)

BY_CODE = {language.code: language for language in LANGUAGES}
BY_LCID = {language.lcid: language for language in LANGUAGES}


def languageNames() -> dict[str, str]:
	"""Names of the languages in the user's language, by code."""
	return {
		# Translators: Name of a language.
		"th": _("Thai"),
		# Translators: Name of a language.
		"en": _("English (United States)"),
		# Translators: Name of a language.
		"en-GB": _("English (United Kingdom)"),
		# Translators: Name of a language.
		"fr": _("French"),
		# Translators: Name of a language.
		"de": _("German"),
		# Translators: Name of a language.
		"es": _("Spanish"),
		# Translators: Name of a language.
		"it": _("Italian"),
		# Translators: Name of a language.
		"pt": _("Portuguese (Brazil)"),
		# Translators: Name of a language.
		"pt-PT": _("Portuguese (Portugal)"),
		# Translators: Name of a language.
		"nl": _("Dutch"),
		# Translators: Name of a language.
		"sv": _("Swedish"),
		# Translators: Name of a language.
		"da": _("Danish"),
		# Translators: Name of a language.
		"nb": _("Norwegian"),
		# Translators: Name of a language.
		"fi": _("Finnish"),
		# Translators: Name of a language.
		"pl": _("Polish"),
		# Translators: Name of a language.
		"cs": _("Czech"),
		# Translators: Name of a language.
		"sk": _("Slovak"),
		# Translators: Name of a language.
		"hu": _("Hungarian"),
		# Translators: Name of a language.
		"ro": _("Romanian"),
		# Translators: Name of a language.
		"hr": _("Croatian"),
		# Translators: Name of a language.
		"sl": _("Slovenian"),
		# Translators: Name of a language.
		"tr": _("Turkish"),
		# Translators: Name of a language.
		"vi": _("Vietnamese"),
		# Translators: Name of a language.
		"id": _("Indonesian"),
		# Translators: Name of a language.
		"ms": _("Malay"),
		# Translators: Name of a language.
		"fil": _("Filipino"),
		# Translators: Name of a language.
		"ru": _("Russian"),
		# Translators: Name of a language.
		"uk": _("Ukrainian"),
		# Translators: Name of a language.
		"bg": _("Bulgarian"),
		# Translators: Name of a language.
		"sr": _("Serbian (Cyrillic)"),
		# Translators: Name of a language.
		"el": _("Greek"),
		# Translators: Name of a language.
		"ar": _("Arabic"),
		# Translators: Name of a language.
		"fa": _("Persian"),
		# Translators: Name of a language.
		"ur": _("Urdu"),
		# Translators: Name of a language.
		"he": _("Hebrew"),
		# Translators: Name of a language.
		"hi": _("Hindi"),
		# Translators: Name of a language.
		"mr": _("Marathi"),
		# Translators: Name of a language.
		"ne": _("Nepali"),
		# Translators: Name of a language.
		"bn": _("Bengali"),
		# Translators: Name of a language.
		"pa": _("Punjabi"),
		# Translators: Name of a language.
		"gu": _("Gujarati"),
		# Translators: Name of a language.
		"ta": _("Tamil"),
		# Translators: Name of a language.
		"te": _("Telugu"),
		# Translators: Name of a language.
		"kn": _("Kannada"),
		# Translators: Name of a language.
		"ml": _("Malayalam"),
		# Translators: Name of a language.
		"si": _("Sinhala"),
		# Translators: Name of a language.
		"lo": _("Lao"),
		# Translators: Name of a language.
		"km": _("Khmer"),
		# Translators: Name of a language.
		"my": _("Burmese"),
		# Translators: Name of a language.
		"ka": _("Georgian"),
		# Translators: Name of a language.
		"hy": _("Armenian"),
		# Translators: Name of a language.
		"am": _("Amharic"),
		# Translators: Name of a language.
		"zh-CN": _("Chinese (Simplified)"),
		# Translators: Name of a language.
		"zh-TW": _("Chinese (Traditional)"),
		# Translators: Name of a language.
		"ja": _("Japanese"),
		# Translators: Name of a language.
		"ko": _("Korean"),
	}


def languageName(lcid: int) -> str:
	"""The name of a Word language number, in the user's language."""
	language = BY_LCID.get(lcid)
	if language is not None:
		return languageNames()[language.code]
	special = {
		# Translators: Name of a proofing language.
		LANG_NO_PROOFING: _("no proofing"),
		# Translators: Name of a proofing language.
		LANG_NONE: _("no language"),
		# Translators: Name of the language of text that mixes several languages.
		WD_UNDEFINED: _("mixed languages"),
	}
	if lcid in special:
		return special[lcid]
	# Translators: Name of a language Word reports by number.
	return _("language {number}").format(number=lcid)


# Scripts ---------------------------------------------------------------------------------------


@dataclass(frozen=True)
class Script:
	id: str
	letters: str  # character ranges of the script's letters, usable in a regex class and a Word wildcard


#: Scripts the language check knows. Chinese, Japanese and Korean are not checked: Word keeps their
#: language separately (East Asian language) and marks it reliably on its own.
SCRIPTS: tuple[Script, ...] = (
	Script("latin", "A-Za-zÀ-ÖØ-öø-ɏ"),
	Script("thai", "ก-๎"),
	Script("lao", "ກ-໎"),
	Script("khmer", "ក-៓"),
	Script("myanmar", "က-ဿၐ-ႏ"),
	Script("arabic", "ؠ-يٮ-ۓۺ-ۿݐ-ݿ"),
	Script("hebrew", "א-ת"),
	Script("cyrillic", "Ѐ-ӿ"),
	Script("greek", "Ά-ώ"),
	Script("devanagari", "ऄ-हक़-ॡ"),
	Script("bengali", "অ-হ"),
	Script("gurmukhi", "ਅ-ਹ"),
	Script("gujarati", "અ-હ"),
	Script("tamil", "அ-ஹ"),
	Script("telugu", "అ-హ"),
	Script("kannada", "ಅ-ಹ"),
	Script("malayalam", "അ-ഺ"),
	Script("sinhala", "අ-ෆ"),
	Script("georgian", "ა-ჺ"),
	Script("armenian", "Ա-և"),
	Script("ethiopic", "ሀ-ፚ"),
)

SCRIPTS_BY_ID = {script.id: script for script in SCRIPTS}


def scriptNames() -> dict[str, str]:
	return {
		# Translators: Name of a writing system (script).
		"latin": _("Latin-script"),
		# Translators: Name of a writing system (script).
		"thai": _("Thai"),
		# Translators: Name of a writing system (script).
		"lao": _("Lao"),
		# Translators: Name of a writing system (script).
		"khmer": _("Khmer"),
		# Translators: Name of a writing system (script).
		"myanmar": _("Burmese"),
		# Translators: Name of a writing system (script).
		"arabic": _("Arabic-script"),
		# Translators: Name of a writing system (script).
		"hebrew": _("Hebrew"),
		# Translators: Name of a writing system (script).
		"cyrillic": _("Cyrillic"),
		# Translators: Name of a writing system (script).
		"greek": _("Greek"),
		# Translators: Name of a writing system (script).
		"devanagari": _("Devanagari"),
		# Translators: Name of a writing system (script).
		"bengali": _("Bengali"),
		# Translators: Name of a writing system (script).
		"gurmukhi": _("Gurmukhi"),
		# Translators: Name of a writing system (script).
		"gujarati": _("Gujarati"),
		# Translators: Name of a writing system (script).
		"tamil": _("Tamil"),
		# Translators: Name of a writing system (script).
		"telugu": _("Telugu"),
		# Translators: Name of a writing system (script).
		"kannada": _("Kannada"),
		# Translators: Name of a writing system (script).
		"malayalam": _("Malayalam"),
		# Translators: Name of a writing system (script).
		"sinhala": _("Sinhala"),
		# Translators: Name of a writing system (script).
		"georgian": _("Georgian"),
		# Translators: Name of a writing system (script).
		"armenian": _("Armenian"),
		# Translators: Name of a writing system (script).
		"ethiopic": _("Ethiopic"),
	}


def scriptName(scriptId: str) -> str:
	return scriptNames().get(scriptId, scriptId)


@cache
def _letterPattern(scriptId: str) -> re.Pattern[str]:
	return re.compile(f"[{SCRIPTS_BY_ID[scriptId].letters}]")


def wordWildcard(scriptId: str) -> str:
	"""A Word wildcard pattern for a run of the script's letters."""
	return f"[{SCRIPTS_BY_ID[scriptId].letters}]@"


def scriptsIn(text: str) -> list[str]:
	"""Ids of the checked scripts whose letters occur in the text, in SCRIPTS order."""
	return [script.id for script in SCRIPTS if _letterPattern(script.id).search(text)]


def scriptOf(text: str) -> str | None:
	"""The script with the most letters in the text, for example of one word."""
	best, bestCount = None, 0
	for script in SCRIPTS:
		count = len(_letterPattern(script.id).findall(text))
		if count > bestCount:
			best, bestCount = script.id, count
	return best


def languagesOf(scriptId: str) -> set[int]:
	return {language.lcid for language in LANGUAGES if language.script == scriptId}


def isWrongLanguage(scriptId: str, lcid: int) -> bool:
	"""True when text of the script is marked with a language of another script, or with no language.

	Languages the add-on does not know are never called wrong, so unusual but correct choices stay.
	"""
	if lcid in (LANG_NO_PROOFING, WD_UNDEFINED):
		return False
	if lcid == LANG_NONE:
		return True
	known = BY_LCID.get(lcid)
	return known is not None and known.script != scriptId


def preferredLanguage(scriptId: str, preferred: list[int], latinDefault: int) -> int | None:
	"""The language to give text of a script: the user's first language in that script, else a default."""
	for lcid in preferred:
		known = BY_LCID.get(lcid)
		if known is not None and known.script == scriptId:
			return lcid
	if scriptId == "latin":
		return latinDefault
	ordered = [language.lcid for language in LANGUAGES if language.script == scriptId]
	return ordered[0] if ordered else None


def parseLanguageList(value: str) -> list[int]:
	"""Language numbers from the setting, a comma separated list of language codes."""
	result = []
	for code in (value or "").split(","):
		language = BY_CODE.get(code.strip())
		if language is not None and language.lcid not in result:
			result.append(language.lcid)
	return result


# Digits ------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class DigitSystem:
	id: str
	digits: str  # zero to nine
	script: str | None  # script whose text usually uses these digits


DIGIT_SYSTEMS: tuple[DigitSystem, ...] = (
	DigitSystem("thai", "๐๑๒๓๔๕๖๗๘๙", "thai"),
	DigitSystem("lao", "".join(chr(0x0ED0 + i) for i in range(10)), "lao"),
	DigitSystem("khmer", "".join(chr(0x17E0 + i) for i in range(10)), "khmer"),
	DigitSystem("myanmar", "".join(chr(0x1040 + i) for i in range(10)), "myanmar"),
	DigitSystem("arabicIndic", "".join(chr(0x0660 + i) for i in range(10)), "arabic"),
	DigitSystem("persian", "".join(chr(0x06F0 + i) for i in range(10)), "arabic"),
	DigitSystem("devanagari", "".join(chr(0x0966 + i) for i in range(10)), "devanagari"),
	DigitSystem("bengali", "".join(chr(0x09E6 + i) for i in range(10)), "bengali"),
	DigitSystem("gurmukhi", "".join(chr(0x0A66 + i) for i in range(10)), "gurmukhi"),
	DigitSystem("gujarati", "".join(chr(0x0AE6 + i) for i in range(10)), "gujarati"),
	DigitSystem("tamil", "".join(chr(0x0BE6 + i) for i in range(10)), "tamil"),
	DigitSystem("telugu", "".join(chr(0x0C66 + i) for i in range(10)), "telugu"),
	DigitSystem("kannada", "".join(chr(0x0CE6 + i) for i in range(10)), "kannada"),
	DigitSystem("malayalam", "".join(chr(0x0D66 + i) for i in range(10)), "malayalam"),
)
WESTERN_DIGITS = "0123456789"
DIGIT_SYSTEMS_BY_ID = {system.id: system for system in DIGIT_SYSTEMS}


def digitSystemNames() -> dict[str, str]:
	return {
		# Translators: Name of a set of digits.
		"thai": _("Thai digits"),
		# Translators: Name of a set of digits.
		"lao": _("Lao digits"),
		# Translators: Name of a set of digits.
		"khmer": _("Khmer digits"),
		# Translators: Name of a set of digits.
		"myanmar": _("Burmese digits"),
		# Translators: Name of a set of digits.
		"arabicIndic": _("Arabic-Indic digits"),
		# Translators: Name of a set of digits.
		"persian": _("Persian and Urdu digits"),
		# Translators: Name of a set of digits.
		"devanagari": _("Devanagari digits"),
		# Translators: Name of a set of digits.
		"bengali": _("Bengali digits"),
		# Translators: Name of a set of digits.
		"gurmukhi": _("Gurmukhi digits"),
		# Translators: Name of a set of digits.
		"gujarati": _("Gujarati digits"),
		# Translators: Name of a set of digits.
		"tamil": _("Tamil digits"),
		# Translators: Name of a set of digits.
		"telugu": _("Telugu digits"),
		# Translators: Name of a set of digits.
		"kannada": _("Kannada digits"),
		# Translators: Name of a set of digits.
		"malayalam": _("Malayalam digits"),
	}


def countDigitSystems(text: str) -> dict[str, int]:
	"""How many digits of each system the text has; "western" for 0 to 9. Only systems that occur."""
	counts = {"western": sum(text.count(d) for d in WESTERN_DIGITS)}
	for system in DIGIT_SYSTEMS:
		counts[system.id] = sum(text.count(d) for d in system.digits)
	return {key: value for key, value in counts.items() if value}


def digitsOf(systemId: str) -> str:
	return WESTERN_DIGITS if systemId == "western" else DIGIT_SYSTEMS_BY_ID[systemId].digits

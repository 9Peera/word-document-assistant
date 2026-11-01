# Word Document Assistant for NVDA
# Accessibility of the document: checker, images and alternative text, language of the text.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import base64
import re
from dataclasses import dataclass
from typing import Any

import addonHandler

from . import ai, wordCom
from .elements import Action, ElementType, Host, Item, Section, conf
from .wordCom import COM_ERRORS, WordContext, cleanText, flatText, undoRecord

addonHandler.initTranslation()

# Word constants
WD_OUTLINE_BODY_TEXT = 10
WD_WITHIN_TABLE = 12
WD_UNDEFINED = 9999999
LANG_THAI = 1054
LANG_ENGLISH_US = 1033
LANG_NO_PROOFING = 1024
STYLE_HEADING1 = -2  # Heading n is -1 - n
STYLE_NORMAL = -1
INLINE_PICTURE_TYPES = {3, 4, 12, 13, 15}  # picture, linked picture, chart, chart (linked), SmartArt
SHAPE_PICTURE_TYPES = {
	3,
	6,
	11,
	13,
	15,
	24,
	28,
}  # chart, group, linked picture, picture, SmartArt, graphic, 3D model

_THAI = re.compile(r"[฀-๿]")
_LATIN = re.compile(r"[A-Za-z]")

#: Link texts that do not say where the link goes.
UNCLEAR_LINK_TEXTS = {
	"click here",
	"click",
	"here",
	"link",
	"this link",
	"more",
	"read more",
	"learn more",
	"details",
	"คลิก",
	"คลิกที่นี่",
	"คลิกที่นี่เลย",
	"ที่นี่",
	"ลิงก์",
	"ลิงค์",
	"อ่านต่อ",
	"เพิ่มเติม",
	"ดูเพิ่มเติม",
	"รายละเอียด",
}


def languageName(languageId: int) -> str:
	names = {
		# Translators: Name of a proofing language.
		LANG_THAI: _("Thai"),
		# Translators: Name of a proofing language.
		LANG_ENGLISH_US: _("English (United States)"),
		# Translators: Name of a proofing language.
		2057: _("English (United Kingdom)"),
		# Translators: Name of a proofing language.
		LANG_NO_PROOFING: _("no proofing"),
		# Translators: Name of a proofing language.
		0: _("no language"),
		# Translators: Name of the language of text that mixes several languages.
		WD_UNDEFINED: _("mixed languages"),
	}
	# Translators: Name of a language Word reports by number.
	return names.get(languageId, _("language {number}").format(number=languageId))


def _latinLanguage(ctx: WordContext) -> int:
	"""The language for Latin text: the Normal style's language unless that is Thai."""
	try:
		language = int(ctx.document.Styles.Item(STYLE_NORMAL).LanguageID)
		if language not in (LANG_THAI, WD_UNDEFINED, 0, LANG_NO_PROOFING):
			return language
	except (*COM_ERRORS, ValueError, TypeError):
		pass
	return LANG_ENGLISH_US


def _paragraphs(ctx: WordContext, limit: int):
	"""Yield (paragraph, text) for the main text, walking with Next for speed."""
	paragraph = ctx.document.Content.Paragraphs.First
	count = 0
	while paragraph is not None and count < limit:
		try:
			text = str(paragraph.Range.Text or "")
		except COM_ERRORS:
			break
		yield paragraph, text
		count += 1
		try:
			paragraph = paragraph.Next()
		except COM_ERRORS:
			break


def _isEmptyText(text: str) -> bool:
	return not text.strip("\r\n\x07\x0b \t\xa0")


def _inTable(rangeObj: Any) -> bool:
	try:
		return bool(rangeObj.Information(WD_WITHIN_TABLE))
	except COM_ERRORS:
		return False


def _near(ctx: WordContext, rangeObj: Any, length: int = 80) -> str:
	"""Text around a range, used to show where an item is and to help AI describe it."""
	doc = ctx.document
	start = max(0, rangeObj.Start - length)
	end = min(doc.Content.End, rangeObj.End + length)
	return flatText(doc.Range(start, end).Text)


# Images ------------------------------------------------------------------------------


@dataclass
class Picture:
	obj: Any
	floating: bool
	index: int  # position in its collection, used to find its image data


def _altText(obj: Any) -> str:
	try:
		return cleanText(obj.AlternativeText)
	except COM_ERRORS:
		return ""


def _anchor(picture: Picture) -> Any:
	return picture.obj.Anchor if picture.floating else picture.obj.Range


def collectPictures(ctx: WordContext, limit: int) -> list[Picture]:
	doc = ctx.document
	pictures: list[Picture] = []
	try:
		inline = doc.InlineShapes
		for index in range(1, inline.Count + 1):
			shape = inline.Item(index)
			if shape.Type in INLINE_PICTURE_TYPES:
				pictures.append(Picture(shape, False, index))
			if len(pictures) >= limit:
				break
	except COM_ERRORS:
		pass
	try:
		shapes = doc.Shapes
		for index in range(1, shapes.Count + 1):
			shape = shapes.Item(index)
			if shape.Type in SHAPE_PICTURE_TYPES:
				pictures.append(Picture(shape, True, index))
			if len(pictures) >= limit:
				break
	except COM_ERRORS:
		pass
	return pictures


_IMAGE_PART = re.compile(
	r'<pkg:part[^>]*pkg:contentType="(image/[^"]+)"[^>]*>\s*<pkg:binaryData>([^<]+)</pkg:binaryData>',
	re.S,
)
#: Image types the AI services accept directly.
DIRECT_TYPES = {"image/png", "image/jpeg", "image/gif", "image/webp"}
#: Image types that can be converted to PNG with wx.
CONVERTIBLE_TYPES = {"image/bmp", "image/tiff", "image/x-tiff", "image/gif", "image/jpeg", "image/png"}
MAX_IMAGE_SIDE = 1600


class ImageUnavailable(Exception):
	"""The picture's image data cannot be sent to AI. The message is for the user."""


def imageData(ctx: WordContext, picture: Picture) -> tuple[bytes, str]:
	"""The picture as (bytes, mime type), taken from the document's Open XML."""
	rangeObj = _anchor(picture)
	if picture.floating:
		rangeObj = rangeObj.Paragraphs.Item(1).Range
	try:
		xml = str(rangeObj.WordOpenXML)
	except COM_ERRORS as error:
		# Translators: Reported when the add-on cannot read the picture.
		raise ImageUnavailable(_("Cannot read this picture from the document")) from error
	parts = _IMAGE_PART.findall(xml)
	if not parts:
		raise ImageUnavailable(_("Cannot read this picture from the document"))
	usable = [(mime, data) for mime, data in parts if mime in CONVERTIBLE_TYPES or mime in DIRECT_TYPES]
	if not usable:
		raise ImageUnavailable(
			# Translators: Reported when the picture is a drawing type AI cannot read, such as EMF.
			_(
				"This picture is a drawing format AI cannot read ({type}). Save it as PNG or JPEG first"
			).format(type=parts[0][0])
		)
	mime, data = usable[0]
	if picture.floating:
		# Several floating pictures may share one paragraph; take the matching one when possible.
		position = _floatingPositionInParagraph(ctx, picture)
		if 0 <= position < len(usable):
			mime, data = usable[position]
	raw = base64.b64decode(re.sub(r"\s+", "", data))
	return _prepare(raw, mime)


def _floatingPositionInParagraph(ctx: WordContext, picture: Picture) -> int:
	try:
		paragraphStart = picture.obj.Anchor.Paragraphs.Item(1).Range.Start
		position = 0
		shapes = ctx.document.Shapes
		for index in range(1, picture.index):
			if shapes.Item(index).Anchor.Paragraphs.Item(1).Range.Start == paragraphStart:
				position += 1
		return position
	except COM_ERRORS:
		return -1


def _prepare(raw: bytes, mime: str) -> tuple[bytes, str]:
	"""Convert to PNG when needed and make large pictures smaller, to keep requests fast and cheap."""
	try:
		import io

		import wx

		image = wx.Image(io.BytesIO(raw), wx.BITMAP_TYPE_ANY)
		if not image.IsOk():
			raise ValueError("not an image")
		width, height = image.GetWidth(), image.GetHeight()
		largest = max(width, height)
		if mime in DIRECT_TYPES and largest <= MAX_IMAGE_SIDE and len(raw) < 4_000_000:
			return raw, mime
		if largest > MAX_IMAGE_SIDE:
			scale = MAX_IMAGE_SIDE / largest
			image = image.Scale(
				max(1, int(width * scale)), max(1, int(height * scale)), wx.IMAGE_QUALITY_HIGH
			)
		stream = io.BytesIO()
		image.SaveFile(stream, wx.BITMAP_TYPE_PNG)
		return stream.getvalue(), "image/png"
	except Exception:  # noqa: BLE001
		if mime in DIRECT_TYPES:
			return raw, mime
		raise ImageUnavailable(_("Cannot read this picture from the document")) from None


def setAltText(ctx: WordContext, picture: Picture, text: str, host: Host) -> None:
	# Translators: Name of the undo step for changing alternative text.
	with undoRecord(ctx, _("Alternative text")):
		picture.obj.AlternativeText = text
	host.modified()
	# Translators: Reported after alternative text was saved.
	host.speak(_("Alternative text saved"))


def pictureActions(ctx: WordContext, picture: Picture, host: Host) -> list[Action]:
	def edit() -> None:
		text = host.prompt(
			# Translators: Title of the dialog for alternative text.
			_("Alternative text"),
			# Translators: Label of the alternative text field.
			_("Describe the picture for people who cannot see it:"),
			_altText(picture.obj),
		)
		if text is not None:
			setAltText(ctx, picture, text, host)

	def describe(detailed: bool) -> None:
		try:
			image, mime = imageData(ctx, picture)
		except ImageUnavailable as error:
			host.speak(str(error))
			return
		context = _near(ctx, _anchor(picture), 150)
		host.aiDraft(
			_("Alternative text"),
			# Translators: Label of the field with the AI description, which the user can edit.
			_("AI description, edit it before saving:"),
			lambda: ai.describeImage(image, mime, context, detailed),
			lambda text: setAltText(ctx, picture, text, host),
		)

	def decorative() -> None:
		try:
			picture.obj.Decorative = True
		except COM_ERRORS:
			# Translators: Reported when this version of Word cannot mark pictures as decorative.
			host.speak(_("This version of Word cannot mark pictures as decorative"))
			return
		host.modified()
		# Translators: Reported after a picture was marked as decorative.
		host.speak(_("Marked as decorative"))

	actions = [
		# Translators: Action that edits the alternative text of a picture.
		Action(_("Edit alternative text..."), edit),
	]
	if host.aiAvailable():
		actions += [
			# Translators: Action that asks AI for short alternative text of a picture.
			Action(_("Describe with AI (short alternative text)..."), lambda: describe(False)),
			# Translators: Action that asks AI for a detailed description of a picture.
			Action(_("Describe with AI (detailed description)..."), lambda: describe(True)),
		]
	# Translators: Action that marks a picture as decorative, so screen readers skip it.
	actions.append(Action(_("Mark as decorative"), decorative))
	return actions


def _pictureTypeName(picture: Picture) -> str:
	try:
		shapeType = picture.obj.Type
	except COM_ERRORS:
		shapeType = 0
	if (not picture.floating and shapeType in (12, 13)) or (picture.floating and shapeType == 3):
		# Translators: Kind of picture.
		return _("Chart")
	if shapeType == 15 and not picture.floating or shapeType == 24 and picture.floating:
		# Translators: Kind of picture.
		return _("SmartArt or graphic")
	if picture.floating and shapeType == 6:
		# Translators: Kind of picture.
		return _("Group of shapes")
	if picture.floating:
		# Translators: Kind of picture.
		return _("Floating picture")
	# Translators: Kind of picture.
	return _("Picture")


class Pictures(ElementType):
	id = "pictures"
	vk = ord("M")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Pictures")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the picture list.
		return [_("Kind"), _("Alternative text"), _("Nearby text")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the picture details.
			Section(_("Kind")),
			# Translators: Section label in the picture details.
			Section(_("Alternative text"), multiline=True),
			# Translators: Section label in the picture details.
			Section(_("Nearby text"), multiline=True),
		]

	def collect(self, ctx, limit):
		items = []
		for picture in collectPictures(ctx, limit):
			try:
				anchor = _anchor(picture)
				items.append(Item(obj=picture, navRange=anchor, start=anchor.Start))
			except COM_ERRORS:
				continue
		items.sort(key=lambda i: i.start)
		return items, len(items)

	def _alt(self, item: Item) -> str:
		# Translators: Shown when a picture has no alternative text.
		return _altText(item.obj.obj) or _("no alternative text")

	def summary(self, ctx, item):
		return f"{_pictureTypeName(item.obj)}: {self._alt(item)}"

	def row(self, ctx, item):
		return [_pictureTypeName(item.obj), flatText(self._alt(item), 120), _near(ctx, item.navRange, 40)]

	def sectionTexts(self, ctx, item):
		return [_pictureTypeName(item.obj), self._alt(item), _near(ctx, item.navRange, 150)]

	def copyText(self, ctx, item):
		return _altText(item.obj.obj)

	def moveTo(self, ctx, item):
		if item.obj.floating:
			super().moveTo(ctx, item)
		else:
			# Selecting an inline picture lets the user press Shift+F10 on it or delete it.
			item.navRange.Select()

	def actions(self, ctx, item, anchor, host):
		actions = pictureActions(ctx, item.obj, host) if item is not None else []
		return actions + super().actions(ctx, item, anchor, host)


# Language -------------------------------------------------------------------------------


@dataclass
class LanguageProblem:
	paragraph: Any
	text: str
	languageId: int
	kind: str  # "thai" Thai text not marked Thai, "latin" Latin text marked Thai, "mixed" some words wrong


def _problemFor(paragraph: Any, text: str, checkWords: bool) -> LanguageProblem | None:
	hasThai = bool(_THAI.search(text))
	hasLatin = bool(_LATIN.search(text))
	if not hasThai and not hasLatin:
		return None
	try:
		languageId = int(paragraph.Range.LanguageID)
	except (*COM_ERRORS, ValueError, TypeError):
		return None
	if languageId == LANG_NO_PROOFING:
		return None
	if hasThai and not hasLatin and languageId != LANG_THAI:
		return LanguageProblem(paragraph, text, languageId, "thai")
	if hasLatin and not hasThai and languageId == LANG_THAI:
		return LanguageProblem(paragraph, text, languageId, "latin")
	if hasThai and hasLatin:
		if languageId != WD_UNDEFINED:
			# The whole paragraph has one language although it mixes Thai and Latin text.
			return LanguageProblem(paragraph, text, languageId, "mixed")
		if checkWords and _wrongWords(paragraph):
			return LanguageProblem(paragraph, text, languageId, "mixed")
	return None


def _wrongWords(paragraph: Any) -> int:
	count = 0
	try:
		words = paragraph.Range.Words
		for index in range(1, min(words.Count, 400) + 1):
			word = words.Item(index)
			text = str(word.Text or "")
			if _THAI.search(text) and int(word.LanguageID) != LANG_THAI:
				count += 1
			elif _LATIN.search(text) and not _THAI.search(text) and int(word.LanguageID) == LANG_THAI:
				count += 1
	except (*COM_ERRORS, ValueError, TypeError):
		pass
	return count


def fixParagraphLanguage(ctx: WordContext, paragraph: Any) -> int:
	"""Mark Thai words as Thai and Latin words with the document's Latin language. Returns words changed."""
	latin = _latinLanguage(ctx)
	changed = 0
	words = paragraph.Range.Words
	for index in range(1, words.Count + 1):
		word = words.Item(index)
		text = str(word.Text or "")
		wanted = LANG_THAI if _THAI.search(text) else latin if _LATIN.search(text) else None
		if wanted is None:
			continue
		try:
			if int(word.LanguageID) != wanted:
				word.LanguageID = wanted
				changed += 1
		except (*COM_ERRORS, ValueError, TypeError):
			continue
	return changed


#: Above this number of mixed paragraphs, words are not checked one by one (it would be slow).
WORD_CHECK_LIMIT = 150


class LanguageProblems(ElementType):
	id = "language"
	vk = ord("L")
	slow = True

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Language problems")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the language problem list.
		return [_("Problem"), _("Marked as"), _("Text")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the language problem details.
			Section(_("Problem")),
			# Translators: Section label in the language problem details.
			Section(_("Marked as")),
			# Translators: Section label in the language problem details.
			Section(_("Text"), multiline=True),
		]

	def collect(self, ctx, limit):
		paragraphs = list(_paragraphs(ctx, 20000))
		mixed = sum(1 for _p, text in paragraphs if _THAI.search(text) and _LATIN.search(text))
		checkWords = mixed <= WORD_CHECK_LIMIT
		items = []
		for paragraph, text in paragraphs:
			problem = _problemFor(paragraph, text, checkWords)
			if problem is None:
				continue
			rangeObj = paragraph.Range
			items.append(Item(obj=problem, navRange=rangeObj, start=rangeObj.Start))
			if len(items) >= limit:
				break
		return items, len(items)

	@staticmethod
	def _problemText(problem: LanguageProblem) -> str:
		if problem.kind == "thai":
			# Translators: Language problem.
			return _("Thai text not marked as Thai")
		if problem.kind == "latin":
			# Translators: Language problem.
			return _("English text marked as Thai")
		# Translators: Language problem.
		return _("Some words marked with the wrong language")

	def summary(self, ctx, item):
		problem = item.obj
		return (
			f"{self._problemText(problem)}, {languageName(problem.languageId)}: {flatText(problem.text, 120)}"
		)

	def row(self, ctx, item):
		problem = item.obj
		return [self._problemText(problem), languageName(problem.languageId), flatText(problem.text, 120)]

	def sectionTexts(self, ctx, item):
		problem = item.obj
		return [self._problemText(problem), languageName(problem.languageId), cleanText(problem.text)]

	def copyText(self, ctx, item):
		return cleanText(item.obj.text)

	def actions(self, ctx, item, anchor, host):
		def fixOne() -> None:
			# Translators: Name of the undo step for fixing the language of text.
			with undoRecord(ctx, _("Fix language")):
				changed = fixParagraphLanguage(ctx, item.obj.paragraph)
			host.modified()
			# Translators: Reported after the language was fixed; {count} is a number of words.
			host.speak(_("Language fixed for {count} words").format(count=changed))

		def fixAll() -> None:
			items, _total = self.collect(ctx, 100000)
			changed = 0
			with undoRecord(ctx, _("Fix language")):
				for problem in items:
					try:
						changed += fixParagraphLanguage(ctx, problem.obj.paragraph)
					except COM_ERRORS:
						continue
			host.modified()
			# Translators: Reported after all language problems were fixed.
			host.speak(
				_("Language fixed in {paragraphs} paragraphs, {count} words").format(
					paragraphs=len(items), count=changed
				)
			)

		def setLanguage(languageId: int) -> None:
			with undoRecord(ctx, _("Fix language")):
				item.obj.paragraph.Range.LanguageID = languageId
			host.modified()
			# Translators: Reported after the language of a paragraph was set.
			host.speak(_("Paragraph marked as {language}").format(language=languageName(languageId)))

		actions = []
		if item is not None:
			actions += [
				# Translators: Action that marks Thai words as Thai and English words as English in one paragraph.
				Action(_("Fix this paragraph (Thai words Thai, English words English)"), fixOne),
				# Translators: Action that marks a whole paragraph as Thai.
				Action(_("Mark whole paragraph as Thai"), lambda: setLanguage(LANG_THAI)),
				Action(
					# Translators: Action that marks a whole paragraph as English.
					_("Mark whole paragraph as {language}").format(
						language=languageName(_latinLanguage(ctx))
					),
					lambda: setLanguage(_latinLanguage(ctx)),
				),
			]
		# Translators: Action that fixes every language problem in the document.
		actions.append(Action(_("Fix all language problems in the document"), fixAll))
		return actions + super().actions(ctx, item, anchor, host)


# Accessibility checker ---------------------------------------------------------------------


@dataclass
class Issue:
	kind: str
	rangeObj: Any
	detail: str
	target: Any = None
	extra: Any = None


def _issueName(kind: str) -> str:
	names = {
		# Translators: Accessibility problem.
		"altText": _("Picture without alternative text"),
		# Translators: Accessibility problem.
		"tableHeader": _("Table without a header row"),
		# Translators: Accessibility problem.
		"headingSkip": _("Heading level skipped"),
		# Translators: Accessibility problem.
		"emptyParagraphs": _("Empty paragraphs used for spacing"),
		# Translators: Accessibility problem.
		"linkText": _("Link text does not say where it goes"),
		# Translators: Accessibility problem.
		"title": _("Document has no title"),
		# Translators: Accessibility problem.
		"noHeadings": _("Long document without headings"),
		# Translators: Accessibility problem.
		"language": _("Text marked with the wrong language"),
	}
	return names.get(kind, kind)


def _issueHelp(kind: str) -> str:
	help = {
		"altText": _(
			# Translators: How to fix an accessibility problem.
			"Screen reader users only hear that there is a picture. Add alternative text that says what it shows, "
			"or mark it as decorative."
		),
		"tableHeader": _(
			# Translators: How to fix an accessibility problem.
			"Mark the first row as a header row so screen readers announce the column names."
		),
		"headingSkip": _(
			# Translators: How to fix an accessibility problem.
			"Headings should go down one level at a time, for example from level 1 to level 2."
		),
		"emptyParagraphs": _(
			# Translators: How to fix an accessibility problem.
			"Screen readers read every empty line as blank. Use space after paragraphs instead."
		),
		"linkText": _(
			# Translators: How to fix an accessibility problem.
			"Link text such as click here makes no sense in a list of links. Say where the link goes."
		),
		"title": _(
			# Translators: How to fix an accessibility problem.
			"The title is read first when the document or its PDF is opened."
		),
		"noHeadings": _(
			# Translators: How to fix an accessibility problem.
			"Real Word headings let screen reader users jump through the document. Format it with J."
		),
		"language": _(
			# Translators: How to fix an accessibility problem.
			"Wrong language stops spell checking and makes screen readers use the wrong voice. Use L to fix it."
		),
	}
	return help.get(kind, "")


def _isUnclearLink(text: str, address: str) -> bool:
	clean = " ".join(text.lower().split()).strip(" .:>»")
	if not clean:
		return True
	if clean in UNCLEAR_LINK_TEXTS:
		return True
	return clean.startswith(("http://", "https://", "www.")) or (address and clean == address.lower())


#: Paragraph count above which a document without headings is reported.
LONG_DOCUMENT = 25


def checkDocument(ctx: WordContext, limit: int) -> list[Issue]:
	doc = ctx.document
	issues: list[Issue] = []
	# Pictures
	for picture in collectPictures(ctx, 2000):
		try:
			decorative = bool(getattr(picture.obj, "Decorative", False))
		except COM_ERRORS:
			decorative = False
		if not decorative and not _altText(picture.obj):
			try:
				issues.append(Issue("altText", _anchor(picture), _pictureTypeName(picture), target=picture))
			except COM_ERRORS:
				continue
	# Tables
	try:
		tables = doc.Tables
		for index in range(1, tables.Count + 1):
			table = tables.Item(index)
			try:
				if table.Rows.Count > 1 and not table.Rows.Item(1).HeadingFormat:
					# Translators: Describes a table, for example "Table 2".
					issues.append(
						Issue(
							"tableHeader", table.Range, _("Table {number}").format(number=index), target=table
						)
					)
			except COM_ERRORS:
				continue
	except COM_ERRORS:
		pass
	# Paragraph based checks
	previousLevel = 0
	headings = 0
	emptyRun: list[Any] = []
	paragraphCount = 0
	languageChecked = 0
	for paragraph, text in _paragraphs(ctx, 20000):
		paragraphCount += 1
		rangeObj = paragraph.Range
		empty = _isEmptyText(text) and "\x0c" not in text
		if empty and not _inTable(rangeObj):
			try:
				hasPicture = rangeObj.InlineShapes.Count > 0
			except COM_ERRORS:
				hasPicture = True
			if not hasPicture:
				emptyRun.append(paragraph)
				continue
		if len(emptyRun) >= 2:
			issues.append(
				Issue(
					"emptyParagraphs",
					emptyRun[0].Range,
					# Translators: Describes a run of empty paragraphs; {count} is a number.
					_("{count} empty paragraphs in a row").format(count=len(emptyRun)),
					target=list(emptyRun),
				)
			)
		emptyRun = []
		try:
			level = int(paragraph.OutlineLevel)
		except (*COM_ERRORS, ValueError, TypeError):
			level = WD_OUTLINE_BODY_TEXT
		if level < WD_OUTLINE_BODY_TEXT:
			headings += 1
			if previousLevel and level > previousLevel + 1:
				issues.append(
					Issue(
						"headingSkip",
						rangeObj,
						# Translators: Describes a skipped heading level.
						_("Level {level} after level {previous}: {text}").format(
							level=level, previous=previousLevel, text=flatText(text, 80)
						),
						target=paragraph,
						extra=previousLevel + 1,
					)
				)
			previousLevel = level
		if languageChecked < 5000:
			languageChecked += 1
			problem = _problemFor(paragraph, text, checkWords=False)
			if problem is not None:
				issues.append(Issue("language", rangeObj, flatText(text, 80), target=problem))
		if len(issues) >= limit:
			break
	if len(emptyRun) >= 2:
		issues.append(
			Issue(
				"emptyParagraphs",
				emptyRun[0].Range,
				_("{count} empty paragraphs in a row").format(count=len(emptyRun)),
				target=list(emptyRun),
			)
		)
	if paragraphCount > LONG_DOCUMENT and headings == 0:
		issues.append(
			Issue(
				"noHeadings",
				doc.Range(0, 0),
				# Translators: Describes a document without headings; {count} is a number.
				_("{count} paragraphs, no headings").format(count=paragraphCount),
			)
		)
	# Links
	try:
		links = doc.Hyperlinks
		for index in range(1, links.Count + 1):
			link = links.Item(index)
			try:
				text = str(link.TextToDisplay or "")
				address = str(link.Address or "")
			except COM_ERRORS:
				continue
			if _isUnclearLink(text, address):
				issues.append(Issue("linkText", link.Range, flatText(text, 80) or address, target=link))
	except COM_ERRORS:
		pass
	# Title
	try:
		title = str(doc.BuiltInDocumentProperties("Title").Value or "").strip()
	except COM_ERRORS:
		title = "?"
	if not title:
		# Translators: Detail of the missing title problem.
		issues.append(Issue("title", doc.Range(0, 0), _("File, Info, Properties, Title is empty")))
	return issues


def _firstHeadingText(ctx: WordContext) -> str:
	for paragraph, text in _paragraphs(ctx, 200):
		try:
			if int(paragraph.OutlineLevel) < WD_OUTLINE_BODY_TEXT and text.strip():
				return flatText(text, 120)
		except (*COM_ERRORS, ValueError, TypeError):
			continue
	for _paragraph, text in _paragraphs(ctx, 20):
		if text.strip():
			return flatText(text, 120)
	return ""


class AccessibilityIssues(ElementType):
	id = "accessibility"
	vk = ord("A")
	slow = True

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Accessibility problems")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the accessibility problem list.
		return [_("Problem"), _("Where")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the accessibility problem details.
			Section(_("Problem")),
			# Translators: Section label in the accessibility problem details.
			Section(_("Where")),
			# Translators: Section label in the accessibility problem details.
			Section(_("Why and how to fix it"), multiline=True),
		]

	def collect(self, ctx, limit):
		items = []
		for issue in checkDocument(ctx, limit):
			try:
				items.append(Item(obj=issue, navRange=issue.rangeObj, start=issue.rangeObj.Start))
			except COM_ERRORS:
				continue
		items.sort(key=lambda i: i.start)
		return items, len(items)

	def summary(self, ctx, item):
		issue = item.obj
		return f"{_issueName(issue.kind)}: {issue.detail}"

	def row(self, ctx, item):
		return [_issueName(item.obj.kind), item.obj.detail]

	def sectionTexts(self, ctx, item):
		issue = item.obj
		return [_issueName(issue.kind), issue.detail, _issueHelp(issue.kind)]

	def copyText(self, ctx, item):
		return self.summary(ctx, item)

	def actions(self, ctx, item, anchor, host):
		actions: list[Action] = []
		if item is not None:
			actions += self._fixes(ctx, item.obj, host)
		return actions + super().actions(ctx, item, anchor, host)

	def _fixes(self, ctx: WordContext, issue: Issue, host: Host) -> list[Action]:
		kind = issue.kind
		if kind == "altText":
			return pictureActions(ctx, issue.target, host)
		if kind == "tableHeader":

			def header() -> None:
				# Translators: Name of the undo step for setting a table header row.
				with undoRecord(ctx, _("Header row")):
					issue.target.Rows.Item(1).HeadingFormat = True
				host.modified()
				# Translators: Reported after the first row became a header row.
				host.speak(_("The first row is now a header row"))

			# Translators: Action that makes the first row of a table a header row.
			return [Action(_("Make the first row a header row"), header)]
		if kind == "headingSkip":
			level = int(issue.extra)

			def setLevel() -> None:
				# Translators: Name of the undo step for changing a heading level.
				with undoRecord(ctx, _("Heading level")):
					issue.target.Style = STYLE_HEADING1 + 1 - level
				host.modified()
				# Translators: Reported after a heading level changed.
				host.speak(_("Changed to heading level {level}").format(level=level))

			# Translators: Action that changes a heading to the expected level.
			return [Action(_("Change to heading level {level}").format(level=level), setLevel)]
		if kind == "emptyParagraphs":

			def remove() -> None:
				paragraphs = issue.target
				# Translators: Name of the undo step for removing empty paragraphs.
				with undoRecord(ctx, _("Remove empty paragraphs")):
					try:
						previous = paragraphs[0].Previous()
						if previous is not None:
							previous.SpaceAfter = max(float(previous.SpaceAfter), 12.0)
					except (*COM_ERRORS, ValueError, TypeError):
						pass
					for paragraph in reversed(paragraphs):
						try:
							paragraph.Range.Delete()
						except COM_ERRORS:
							continue
				host.modified()
				# Translators: Reported after empty paragraphs were removed; {count} is a number.
				host.speak(
					_("{count} empty paragraphs removed, space after the paragraph added").format(
						count=len(paragraphs)
					)
				)

			# Translators: Action that removes empty paragraphs and adds space after the previous paragraph.
			return [Action(_("Remove them and use space after the paragraph"), remove)]
		if kind == "linkText":

			def editLink() -> None:
				link = issue.target
				try:
					address = str(link.Address or link.SubAddress or "")
				except COM_ERRORS:
					address = ""
				text = host.prompt(
					# Translators: Title of the dialog for changing link text.
					_("Link text"),
					# Translators: Label of the link text field; {address} is the link address.
					_("Text that says where the link goes ({address}):").format(address=address),
					str(link.TextToDisplay or ""),
				)
				if text:
					# Translators: Name of the undo step for changing link text.
					with undoRecord(ctx, _("Link text")):
						link.TextToDisplay = text
					host.modified()
					# Translators: Reported after the link text changed.
					host.speak(_("Link text changed"))

			# Translators: Action that changes the text of a link.
			return [Action(_("Change link text..."), editLink)]
		if kind == "title":

			def setTitle() -> None:
				text = host.prompt(
					# Translators: Title of the dialog for the document title.
					_("Document title"),
					# Translators: Label of the document title field.
					_("Title:"),
					_firstHeadingText(ctx),
				)
				if text:
					ctx.document.BuiltInDocumentProperties("Title").Value = text
					host.modified()
					# Translators: Reported after the document title was set.
					host.speak(_("Document title set"))

			# Translators: Action that sets the document title.
			return [Action(_("Set the document title..."), setTitle)]
		if kind == "noHeadings":
			# Translators: Action that opens formatting with AI.
			return [Action(_("Format the document with AI..."), host.aiFormat, enabled=host.aiAvailable())]
		if kind == "language":

			def fix() -> None:
				with undoRecord(ctx, _("Fix language")):
					changed = fixParagraphLanguage(ctx, issue.target.paragraph)
				host.modified()
				host.speak(_("Language fixed for {count} words").format(count=changed))

			return [Action(_("Fix this paragraph (Thai words Thai, English words English)"), fix)]
		return []


def summaryCounts(items: list[Item]) -> str:
	counts: dict[str, int] = {}
	for item in items:
		counts[item.obj.kind] = counts.get(item.obj.kind, 0) + 1
	return ", ".join(f"{_issueName(kind)} {count}" for kind, count in counts.items())


__all__ = ["AccessibilityIssues", "LanguageProblems", "Pictures", "conf", "wordCom"]

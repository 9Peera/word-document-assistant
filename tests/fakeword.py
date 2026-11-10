"""A small fake of the Word object model, enough for the add-on's logic.

The document is one string with a paragraph mark (\\r) after each paragraph.
Objects that point into the text keep their offsets; edits do not move other objects.
"""

from __future__ import annotations

import base64
import re
import types
from typing import Any

MIXED = 9999999
THAI = 1054
ENGLISH = 1033


class Coll:
	def __init__(self, items=None):
		self._items = list(items or [])

	@property
	def Count(self):
		return len(self._items)

	def Item(self, index):
		return self._items[index - 1]

	@property
	def First(self):
		return self._items[0] if self._items else None

	def __iter__(self):
		return iter(self._items)


class Font:
	def __init__(self):
		self.Name = self.NameBi = ""
		self.Size = self.SizeBi = 0
		self.Bold = self.BoldBi = False
		self.Color = 0


class Find:
	def __init__(self, rangeObj):
		self.range = rangeObj
		self.Replacement = types.SimpleNamespace(ClearFormatting=lambda: None)

	def ClearFormatting(self):
		self.__dict__.pop("Style", None)
		self.__dict__.pop("LanguageID", None)

	def Execute(self, findText, *args):
		style = getattr(self, "Style", None)
		if findText == "" and style is not None:
			return self._findStyle(style)
		if len(args) < 9:
			return self._search(findText, wildcards=len(args) > 2 and bool(args[2]))
		replaceWith = args[8]
		doc, start, end = self.range.doc, self.range.Start, self.range.End
		part = doc.text[start:end].replace(findText, replaceWith)
		doc.setText(start, end, part)
		self.range.End = start + len(part)
		return True

	def _search(self, findText, wildcards):
		"""Plain or wildcard search (only "@" for one or more), optionally limited to a language."""
		self.searches = getattr(self, "searches", 0) + 1
		doc, rangeObj = self.range.doc, self.range
		if wildcards:
			pattern = findText.replace("@", "+")
		else:
			pattern = re.escape(findText).replace(r"\^p", "\r").replace("^p", "\r")
		language = getattr(self, "LanguageID", None)
		text = doc.text
		if language is not None:
			text = "".join(ch if doc.lang[i] == language else "\x00" for i, ch in enumerate(text))
		end = rangeObj.End if rangeObj.End > rangeObj.Start else len(text)
		match = re.compile(pattern).search(text, rangeObj.Start, end)
		if match is None:
			return False
		rangeObj.Start, rangeObj.End = match.start(), match.end()
		return True

	def _findStyle(self, style):
		"""Like Word: the match is the next run of paragraphs in the style, after the current range."""
		self.searches = getattr(self, "searches", 0) + 1
		run = []
		for paragraph in self.range.doc.paragraphs():
			if paragraph.start < self.range.Start or (run and paragraph.start != run[-1].end):
				if run:
					break
				continue
			if paragraph.styleKey == style:
				run.append(paragraph)
			elif run:
				break
		if not run:
			return False
		self.range.Start, self.range.End = run[0].start, run[-1].end
		return True


class Range:
	StoryType = 1

	def __init__(self, doc, start, end):
		self.doc, self.Start, self.End = doc, start, end
		self.Font = Font()
		self.ListFormat = types.SimpleNamespace(ListType=0)

	# Text
	@property
	def Text(self):
		return self.doc.text[self.Start : self.End]

	@Text.setter
	def Text(self, value):
		self.doc.setText(self.Start, self.End, value)
		self.End = self.Start + len(value)

	@property
	def Duplicate(self):
		return Range(self.doc, self.Start, self.End)

	def Collapse(self, direction):
		if direction == 1:
			self.End = self.Start
		else:
			self.Start = self.End

	def Select(self):
		self.doc.selection.Start, self.doc.selection.End = self.Start, self.End

	def MoveEnd(self, unit, count):
		if unit == 4:
			# By paragraph: to the end of the paragraph the range ends in.
			self.End = next((p.end for p in self.doc.paragraphs() if p.end > self.End), len(self.doc.text))
			return
		self.End = max(self.Start, self.End + count)

	def SetRange(self, start, end):
		self.Start, self.End = start, end

	@property
	def SpellingErrors(self):
		self.doc.proofingCalls.append((self.Start, self.End))
		return Coll(Range(self.doc, s, e) for s, e in self.doc.spelling if self.Start <= s < self.End)

	@property
	def GrammaticalErrors(self):
		self.doc.proofingCalls.append((self.Start, self.End))
		return Coll(Range(self.doc, s, e) for s, e in self.doc.grammar if self.Start <= s < self.End)

	def InsertAfter(self, text):
		self.doc.setText(self.End, self.End, text)
		self.End += len(text)

	def Delete(self):
		self.doc.setText(self.Start, self.End, "")
		self.End = self.Start

	# Structure
	@property
	def Paragraphs(self):
		return Coll(
			p for p in self.doc.paragraphs() if p.end > self.Start and p.start < max(self.End, self.Start + 1)
		)

	@property
	def Words(self):
		words = []
		for match in re.finditer(r"\S+\s*", self.Text):
			words.append(Range(self.doc, self.Start + match.start(), self.Start + match.end()))
		return Coll(words)

	@property
	def Cells(self):
		return Coll(self.doc.cellsIn(self))

	@property
	def InlineShapes(self):
		return Coll(
			s for s in self.doc.inlineShapes if self.Start <= s.position < max(self.End, self.Start + 1)
		)

	def Information(self, what):
		if what == 12:
			return any(t.start <= self.Start < t.end for t in self.doc.tables)
		return 1

	# Language
	@property
	def LanguageID(self):
		values = set(self.doc.lang[self.Start : self.End]) or {
			self.doc.lang[self.Start] if self.Start < len(self.doc.lang) else ENGLISH,
		}
		return values.pop() if len(values) == 1 else MIXED

	@LanguageID.setter
	def LanguageID(self, value):
		for index in range(self.Start, self.End):
			self.doc.lang[index] = value

	@property
	def Find(self):
		return Find(self)

	def GoTo(self, what, which):
		for paragraph in self.doc.paragraphs():
			if paragraph.start > self.Start and paragraph.OutlineLevel < 10:
				return Range(self.doc, paragraph.start, paragraph.start)
		return Range(self.doc, self.Start, self.Start)

	@property
	def WordOpenXML(self):
		parts = []
		for shape in self.InlineShapes:
			data = base64.b64encode(shape.imageBytes).decode()
			parts.append(
				f'<pkg:part pkg:name="/word/media/i.png" pkg:contentType="{shape.mime}"><pkg:binaryData>{data}</pkg:binaryData></pkg:part>',
			)
		return "<pkg:package>" + "".join(parts) + "</pkg:package>"

	@property
	def Footnotes(self):
		return Coll()

	@property
	def Endnotes(self):
		return Coll()

	def InsertCrossReference(self, *args):
		self.doc.inserted.append(args)

	def GetSpellingSuggestions(self):
		return Coll([types.SimpleNamespace(Name="hello")])


class Paragraph:
	def __init__(self, doc, index, start, end):
		self.doc, self.index, self.start, self.end = doc, index, start, end

	@property
	def props(self):
		return self.doc.paraProps.setdefault(self.index, {})

	@property
	def Range(self):
		return Range(self.doc, self.start, self.end)

	def Next(self):
		paragraphs = self.doc.paragraphs()
		return paragraphs[self.index + 1] if self.index + 1 < len(paragraphs) else None

	def Previous(self):
		return self.doc.paragraphs()[self.index - 1] if self.index > 0 else None

	@property
	def OutlineLevel(self):
		return self.props.get("level", 10)

	@property
	def styleKey(self):
		"""The style Find matches: the one set, else Heading n for a heading level, else Normal."""
		if "Style" in self.props:
			return self.props["Style"]
		level = self.OutlineLevel
		return -1 - level if level < 10 else -1

	def __getattr__(self, name):
		if name in (
			"Style",
			"Alignment",
			"LeftIndent",
			"FirstLineIndent",
			"SpaceBefore",
			"SpaceAfter",
			"KeepWithNext",
		):
			return self.props.get(name, 0)
		raise AttributeError(name)

	def __setattr__(self, name, value):
		if name in ("doc", "index", "start", "end"):
			object.__setattr__(self, name, value)
			return
		self.props[name] = value
		if name == "Style" and isinstance(value, int):
			# Heading n is -1 - n; Normal is -1.
			self.props["level"] = -1 - value if -10 <= value <= -2 else 10


class InlineShape:
	def __init__(self, doc, position, alt="", shapeType=3, mime="image/png", imageBytes=b"PNGDATA"):
		self.doc, self.position = doc, position
		self.AlternativeText = alt
		self.Type = shapeType
		self.mime = mime
		self.imageBytes = imageBytes

	@property
	def Range(self):
		return Range(self.doc, self.position, self.position + 1)


class Row:
	def __init__(self, table, index):
		self.table, self.index = table, index
		self.HeadingFormat = False

	@property
	def Range(self):
		return Range(self.table.doc, self.table.start, self.table.end)


class Table:
	def __init__(self, doc, start, end, rows, columns):
		self.doc, self.start, self.end = doc, start, end
		self._rows = [Row(self, i) for i in range(rows)]
		self.Columns = types.SimpleNamespace(Count=columns)
		self.Descr = ""
		self.cells = []

	@property
	def Range(self):
		return Range(self.doc, self.start, self.end)

	@property
	def Rows(self):
		return Coll(self._rows)

	def Cell(self, row, column):
		return types.SimpleNamespace(Range=Range(self.doc, self.start, self.start))


class Link:
	def __init__(self, doc, start, end, address="", subAddress=""):
		self.doc, self.start, self.end = doc, start, end
		self.Address, self.SubAddress = address, subAddress
		self.deleted = False

	@property
	def Range(self):
		return Range(self.doc, self.start, self.end)

	@property
	def TextToDisplay(self):
		return self.doc.text[self.start : self.end]

	@TextToDisplay.setter
	def TextToDisplay(self, value):
		self.doc.setText(self.start, self.end, value)
		self.end = self.start + len(value)

	def Delete(self):
		self.deleted = True


class Bookmark:
	def __init__(self, doc, name, start, end):
		self.doc, self.Name, self.start, self.end = doc, name, start, end

	@property
	def Range(self):
		return Range(self.doc, self.start, self.end)

	def Delete(self):
		self.doc.bookmarks.remove(self)


class Bookmarks:
	def __init__(self, doc):
		self.doc = doc
		self.ShowHidden = False

	@property
	def Count(self):
		return len(self.doc.bookmarks)

	def Item(self, key):
		if isinstance(key, int):
			return self.doc.bookmarks[key - 1]
		return next(b for b in self.doc.bookmarks if b.Name == key)

	def Exists(self, name):
		return any(b.Name == name for b in self.doc.bookmarks)

	def Add(self, name, rangeObj):
		self.doc.bookmarks.append(Bookmark(self.doc, name, rangeObj.Start, rangeObj.End))


class Field:
	def __init__(self, doc, code, start, end):
		self.doc, self.start, self.end = doc, start, end
		self.Code = types.SimpleNamespace(Text=f" {code} ", Start=start)
		self.updated = 0

	@property
	def Result(self):
		return Range(self.doc, self.start, self.end)

	def Update(self):
		self.updated += 1

	def Unlink(self):
		self.doc.fields.remove(self)


class Fields(Coll):
	def __init__(self, doc):
		self.doc = doc

	@property
	def _items(self):
		return self.doc.fields

	def Add(self, rangeObj, fieldType, text, preserve):
		self.doc.fields.append(Field(self.doc, text, rangeObj.Start, rangeObj.Start))

	def Update(self):
		for field in self.doc.fields:
			field.Update()


class Source:
	def __init__(self, xml):
		self.XML = xml
		self.Tag = re.search(r"<b:Tag>([^<]*)</b:Tag>", xml).group(1)

	def Field(self, name):
		if name == "Author":
			found = re.search(r"<b:Author>.*</b:Author>", self.XML)
			return found.group(0) if found else ""
		found = re.search(rf"<b:{name}>([^<]*)</b:{name}>", self.XML)
		return found.group(1) if found else ""


class Selection(Range):
	@property
	def Range(self):
		return Range(self.doc, self.Start, self.End)


class Doc:
	def __init__(self, paragraphs: list[str]):
		self.text = "".join(p + "\r" for p in paragraphs)
		self.lang = [THAI if re.match(r"[฀-๿]", ch) else ENGLISH for ch in self.text]
		self.paraProps: dict[int, dict[str, Any]] = {}
		self.selection = Selection(self, 0, 0)
		self.inlineShapes: list[InlineShape] = []
		self.tables: list[Table] = []
		self.links: list[Link] = []
		self.bookmarks: list[Bookmark] = []
		self.fields: list[Field] = []
		self.sources: list[Source] = []
		self.spelling: list[tuple[int, int]] = []
		self.grammar: list[tuple[int, int]] = []
		#: Ranges whose spelling or grammar errors were asked for.
		self.proofingCalls: list[tuple[int, int]] = []
		self.properties = {"Title": ""}
		self.crossReferenceItems: dict[Any, list[str]] = {}
		self.inserted: list[tuple] = []
		self.Comments = Coll()
		self.Revisions = Coll()
		self.Footnotes = Coll()
		self.Endnotes = Coll()
		self.Shapes = Coll()
		self.TablesOfContents = Coll()
		self.styles: dict[Any, Any] = {}
		self.Styles = types.SimpleNamespace(Item=self._style)
		self.PageSetup = types.SimpleNamespace()
		self.Bookmarks = Bookmarks(self)
		self.Fields = Fields(self)
		self.Bibliography = types.SimpleNamespace(Sources=self._SourcesView(self), BibliographyStyle="APA")
		self.TrackRevisions = False

	class _SourcesView(Coll):
		def __init__(self, doc):
			self.doc = doc

		@property
		def _items(self):
			return self.doc.sources

		def Add(self, xml):
			self.doc.sources.append(Source(xml))

	def _style(self, key):
		if key not in self.styles:
			self.styles[key] = types.SimpleNamespace(
				LanguageID=ENGLISH,
				Font=Font(),
				ParagraphFormat=types.SimpleNamespace(),
			)
		return self.styles[key]

	# Helpers for tests
	def find(self, text: str) -> tuple[int, int]:
		start = self.text.index(text)
		return start, start + len(text)

	def setLevel(self, index: int, level: int) -> None:
		self.paraProps.setdefault(index, {})["level"] = level

	def setText(self, start: int, end: int, value: str) -> None:
		language = self.lang[start] if start < len(self.lang) else ENGLISH
		self.text = self.text[:start] + value + self.text[end:]
		self.lang[start:end] = [THAI if re.match(r"[฀-๿]", ch) else language for ch in value]

	def paragraphs(self) -> list[Paragraph]:
		result, start = [], 0
		for index, part in enumerate(self.text.split("\r")[:-1]):
			end = start + len(part) + 1
			result.append(Paragraph(self, index, start, end))
			start = end
		return result

	def cellsIn(self, rangeObj):
		for table in self.tables:
			if table.start == rangeObj.Start:
				return table.cells
		return []

	# Object model
	@property
	def Content(self):
		return Range(self, 0, len(self.text))

	def Range(self, start, end):
		return Range(self, start, end)

	@property
	def InlineShapes(self):
		return Coll(self.inlineShapes)

	@property
	def Tables(self):
		return Coll(self.tables)

	@property
	def Hyperlinks(self):
		return Coll(link for link in self.links if not link.deleted)

	def BuiltInDocumentProperties(self, name):
		doc = self

		class Prop:
			@property
			def Value(self):
				return doc.properties.get(name, "")

			@Value.setter
			def Value(self, value):
				doc.properties[name] = value

		return Prop()

	def GetCrossReferenceItems(self, referenceType):
		return self.crossReferenceItems.get(referenceType, [])


class Application:
	ScreenUpdating = True
	Version = "16.0"
	FontNames = Coll(["TH SarabunIT๙", "TH Sarabun New", "Calibri"])
	CaptionLabels = Coll([types.SimpleNamespace(Name="Table"), types.SimpleNamespace(Name="Figure")])

	def __init__(self):
		self.undo: list[str] = []
		app = self
		self.UndoRecord = types.SimpleNamespace(
			StartCustomRecord=lambda name: app.undo.append(name),
			EndCustomRecord=lambda: None,
		)

	def LinesToPoints(self, lines):
		return lines * 12


class Context:
	"""Stands in for wordCom.WordContext."""

	def __init__(self, doc: Doc):
		self.document = doc
		self.application = Application()
		self.window = types.SimpleNamespace(
			View=types.SimpleNamespace(RevisionsFilter=types.SimpleNamespace(Markup=2)),
		)

	@property
	def selection(self):
		return self.document.selection


class Host:
	"""Records what element actions ask of the add-on."""

	def __init__(self, prompts=None, choices=None, confirm=True, ai=True):
		self.spoken: list[str] = []
		self.copied: list[str] = []
		self.prompts = list(prompts or [])
		self.choices = list(choices or [])
		self.confirmAnswer = confirm
		self.ai = ai
		self.changes = 0
		self.formatted = False

	def speak(self, message):
		self.spoken.append(message)

	def prompt(self, title, label, value=""):
		return self.prompts.pop(0) if self.prompts else value

	def confirm(self, title, message):
		return self.confirmAnswer

	def copy(self, text):
		self.copied.append(text)

	def modified(self):
		self.changes += 1

	def aiAvailable(self):
		return self.ai

	def choose(self, title, label, options):
		return self.choices.pop(0) if self.choices else None

	def aiDraft(self, title, label, work, apply):
		apply(work())

	def aiFormat(self):
		self.formatted = True

	def updateWordLists(self):
		self.spoken.append("update")

	def aiWordSuggestions(self, word, context, apply):
		apply("fixed")

	def aiReplyDrafts(self, author, comment, commentedText, apply):
		apply("reply")

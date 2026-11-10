# Word Document Assistant for NVDA
# Document structure: headings, tables, links, bookmarks, fields, cross-references and citations.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import re
from typing import Any
from xml.sax.saxutils import escape

import addonHandler

from . import ai, wordCom
from .elements import Action, ElementType, Host, Item, Section
from .wordCom import COM_ERRORS, WordContext, cleanText, flatText, undoRecord

addonHandler.initTranslation()

# Word constants
WD_GOTO_HEADING = 11
WD_GOTO_NEXT = 2
WD_OUTLINE_BODY_TEXT = 10
STYLE_NORMAL = -1
STYLE_HEADING1 = -2
WD_FIELD_EMPTY = -1
WD_COLLAPSE_END = 0
WD_STYLE_TYPE_PARAGRAPH = 1
#: Built-in styles Heading 1 to Heading 9 (wdStyleHeading1 is -2, wdStyleHeading9 is -10).
BUILTIN_HEADING_STYLES = tuple(range(STYLE_HEADING1, STYLE_HEADING1 - 9, -1))


def _headingStyles(doc: Any) -> list[Any]:
	"""Heading 1 to 9 plus the document's own paragraph styles that have an outline level."""
	styles: list[Any] = list(BUILTIN_HEADING_STYLES)
	try:
		count = int(doc.Styles.Count)
	except (*COM_ERRORS, AttributeError, TypeError, ValueError):
		return styles
	for index in range(1, count + 1):
		try:
			style = doc.Styles.Item(index)
			if style.BuiltIn or style.Type != WD_STYLE_TYPE_PARAGRAPH:
				continue
			if int(style.ParagraphFormat.OutlineLevel) < WD_OUTLINE_BODY_TEXT:
				styles.append(str(style.NameLocal))
		except (*COM_ERRORS, AttributeError, TypeError, ValueError):
			continue
	return styles


_ERROR_RESULT = re.compile(r"^\s*(error!|ข้อผิดพลาด!|ข้อผิดพลาด)", re.I)


def _fieldCode(field: Any) -> str:
	try:
		return " ".join(str(field.Code.Text or "").split())
	except COM_ERRORS:
		return ""


def _fieldResult(field: Any) -> str:
	try:
		return cleanText(field.Result.Text)
	except COM_ERRORS:
		return ""


def _fieldKind(field: Any) -> str:
	code = _fieldCode(field)
	return code.split(" ", 1)[0].upper() if code else ""


def _isError(field: Any) -> bool:
	return bool(_ERROR_RESULT.match(_fieldResult(field)))


def _okOrError(isError: bool) -> str:
	# Translators: Status of a field or link.
	return _("Error") if isError else _("OK")


# Headings ---------------------------------------------------------------------------------


class Headings(ElementType):
	id = "headings"
	vk = ord("O")
	slow = True

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Headings")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the heading list.
		return [_("Level"), _("Heading")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the heading details.
			Section(_("Level")),
			# Translators: Section label in the heading details.
			Section(_("Heading"), multiline=True),
			# Translators: Section label in the heading details.
			Section(_("Style")),
		]

	def collect(self, ctx, limit):
		doc = ctx.document
		found: dict[int, Item] = {}
		# Word's Find jumps straight to paragraphs in a heading style, which stays fast in long documents.
		# Searching heading by heading with GoTo can take so long that Word stops answering.
		budget = wordCom.budgetOf(ctx)
		for style in _headingStyles(doc):
			self._findStyle(doc, style, found, limit, budget)
			if len(found) >= limit or budget.over():
				break
		if not found:
			# Headings made with a paragraph outline level instead of a heading style.
			self._collectWithGoTo(doc, found, limit, budget)
		items = [found[start] for start in sorted(found)][:limit]
		return items, len(items)

	@staticmethod
	def _add(found: dict[int, Item], paragraph: Any) -> None:
		try:
			rangeObj = paragraph.Range
			start = int(rangeObj.Start)
			if start in found:
				return
			level = int(paragraph.OutlineLevel)
			if level < WD_OUTLINE_BODY_TEXT and cleanText(rangeObj.Text):
				found[start] = Item(obj=paragraph, navRange=rangeObj, start=start, cache={"level": level})
		except (*COM_ERRORS, ValueError, TypeError, AttributeError):
			pass

	def _findStyle(self, doc: Any, style: Any, found: dict[int, Item], limit: int, budget: Any) -> None:
		for match in wordCom.findAll(doc.Content, style=style, budget=budget):
			# One match covers a run of paragraphs with the same style.
			for paragraph in wordCom.iterate(match.Paragraphs, budget):
				self._add(found, paragraph)
				if len(found) >= limit:
					return

	def _collectWithGoTo(self, doc: Any, found: dict[int, Item], limit: int, budget: Any) -> None:
		try:
			self._add(found, doc.Content.Paragraphs.First)
			docEnd = int(doc.Content.End)
		except COM_ERRORS:
			return
		position = doc.Range(0, 0)
		lastStart = -1
		while len(found) < limit and not budget.over():
			try:
				foundRange = position.GoTo(WD_GOTO_HEADING, WD_GOTO_NEXT)
				if foundRange is None or foundRange.Start <= lastStart:
					return
				lastStart = foundRange.Start
				paragraph = foundRange.Paragraphs.Item(1)
				self._add(found, paragraph)
				end = int(paragraph.Range.End)
			except COM_ERRORS:
				return
			if end >= docEnd - 1:
				return
			position = doc.Range(end, end)

	@staticmethod
	def _level(item: Item) -> int:
		return int(item.cache.get("level", 0))

	def summary(self, ctx, item):
		# Translators: Spoken for a heading, for example "Level 2: Budget".
		return _("Level {level}: {text}").format(
			level=self._level(item),
			text=flatText(item.obj.Range.Text, 150),
		)

	def row(self, ctx, item):
		indent = "  " * (self._level(item) - 1)
		return [str(self._level(item)), indent + flatText(item.obj.Range.Text, 150)]

	def sectionTexts(self, ctx, item):
		try:
			style = str(item.obj.Style.NameLocal)
		except COM_ERRORS:
			style = ""
		return [str(self._level(item)), cleanText(item.obj.Range.Text), style]

	def copyText(self, ctx, item):
		return cleanText(item.obj.Range.Text)

	def actions(self, ctx, item, anchor, host):
		actions: list[Action] = []
		if item is not None:
			paragraph = item.obj
			level = self._level(item)

			def setLevel(newLevel: int) -> None:
				# Translators: Name of the undo step for changing a heading level.
				with undoRecord(ctx, _("Heading level")):
					paragraph.Style = STYLE_HEADING1 + 1 - newLevel if newLevel else STYLE_NORMAL
				host.modified()
				if newLevel:
					host.speak(_("Changed to heading level {level}").format(level=newLevel))
				else:
					# Translators: Reported after a heading became body text.
					host.speak(_("Changed to body text"))

			if level > 1:
				# Translators: Action that raises a heading one level, for example from 2 to 1.
				actions.append(
					Action(
						_("Promote to level {level}").format(level=level - 1),
						lambda: setLevel(level - 1),
					),
				)
			if level < 9:
				# Translators: Action that lowers a heading one level, for example from 1 to 2.
				actions.append(
					Action(_("Demote to level {level}").format(level=level + 1), lambda: setLevel(level + 1)),
				)
			for newLevel in range(1, 7):
				if newLevel not in (level, level - 1, level + 1):
					actions.append(
						# Translators: Action that changes a heading to a given level.
						Action(
							_("Change to heading level {level}").format(level=newLevel),
							lambda n=newLevel: setLevel(n),
						),
					)
			# Translators: Action that changes a heading to body text.
			actions.append(Action(_("Change to body text"), lambda: setLevel(0)))
		return actions + super().actions(ctx, item, anchor, host)


# Tables -------------------------------------------------------------------------------------


def tableText(table: Any, maxCells: int = 600) -> str:
	"""The table as lines of tab separated cells."""
	rows: dict[int, list[str]] = {}
	try:
		cells = table.Range.Cells
		for index in range(1, min(cells.Count, maxCells) + 1):
			cell = cells.Item(index)
			rows.setdefault(int(cell.RowIndex), []).append(flatText(cell.Range.Text))
	except COM_ERRORS:
		return flatText(table.Range.Text)
	return "\n".join("\t".join(rows[r]) for r in sorted(rows))


class Tables(ElementType):
	id = "tables"
	vk = ord("U")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Tables")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the table list.
		return [_("Number"), _("Size"), _("Header row"), _("First row")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the table details.
			Section(_("Size")),
			# Translators: Section label in the table details.
			Section(_("Header row")),
			# Translators: Section label in the table details.
			Section(_("Alternative text"), multiline=True),
			# Translators: Section label in the table details.
			Section(_("Content"), multiline=True),
		]

	def collect(self, ctx, limit):
		items = []
		tables = ctx.document.Tables
		for index, table in enumerate(wordCom.iterate(tables, wordCom.budgetOf(ctx)), start=1):
			if len(items) >= limit:
				break
			items.append(
				Item(obj=table, navRange=table.Range, start=table.Range.Start, cache={"number": index}),
			)
		return items, tables.Count

	@staticmethod
	def _size(table: Any) -> str:
		try:
			# Translators: Size of a table, for example "5 rows, 3 columns".
			return _("{rows} rows, {columns} columns").format(
				rows=table.Rows.Count,
				columns=table.Columns.Count,
			)
		except COM_ERRORS:
			# Translators: Size of a table with merged cells that Word cannot count by columns.
			return _("{rows} rows").format(rows=table.Rows.Count)

	@staticmethod
	def _hasHeader(table: Any) -> bool:
		try:
			return bool(table.Rows.Item(1).HeadingFormat)
		except COM_ERRORS:
			return False

	@staticmethod
	def _firstRow(table: Any) -> str:
		try:
			return flatText(table.Rows.Item(1).Range.Text, 120)
		except COM_ERRORS:
			return ""

	def _yesNo(self, value: bool) -> str:
		# Translators: Yes or no.
		return _("yes") if value else _("no")

	def summary(self, ctx, item):
		table = item.obj
		# Translators: Spoken for a table, for example "Table 2: 5 rows, 3 columns, header row yes, Name Age".
		return _("Table {number}: {size}, header row {header}, {first}").format(
			number=item.cache["number"],
			size=self._size(table),
			header=self._yesNo(self._hasHeader(table)),
			first=self._firstRow(table),
		)

	def row(self, ctx, item):
		table = item.obj
		return [
			str(item.cache["number"]),
			self._size(table),
			self._yesNo(self._hasHeader(table)),
			self._firstRow(table),
		]

	def sectionTexts(self, ctx, item):
		table = item.obj
		try:
			descr = cleanText(table.Descr)
		except COM_ERRORS:
			descr = ""
		return [self._size(table), self._yesNo(self._hasHeader(table)), descr, tableText(table, 200)]

	def copyText(self, ctx, item):
		return tableText(item.obj)

	def edit(self, ctx, item):
		# Put the caret in the first cell, so the NVDA table commands work straight away.
		item.obj.Cell(1, 1).Range.Select()
		ctx.selection.Collapse(wordCom.WD_COLLAPSE_START)
		# Translators: Reported after Enter moves into a table.
		return _("In table {number}, first cell").format(number=item.cache["number"])

	def actions(self, ctx, item, anchor, host):
		actions: list[Action] = []
		if item is not None:
			table = item.obj
			header = self._hasHeader(table)

			def toggleHeader() -> None:
				with undoRecord(ctx, _("Header row")):
					table.Rows.Item(1).HeadingFormat = not header
				host.modified()
				if header:
					# Translators: Reported after the header row was removed.
					host.speak(_("The first row is no longer a header row"))
				else:
					host.speak(_("The first row is now a header row"))

			def setDescription(text: str) -> None:
				# Translators: Name of the undo step for changing alternative text.
				with undoRecord(ctx, _("Alternative text")):
					table.Descr = text
				host.modified()
				host.speak(_("Alternative text saved"))

			def editDescription() -> None:
				try:
					current = cleanText(table.Descr)
				except COM_ERRORS:
					current = ""
				# Translators: Label of the field with the alternative text of a table.
				text = host.prompt(
					_("Alternative text"),
					_("Describe the table for people who cannot see it:"),
					current,
				)
				if text is not None:
					setDescription(text)

			def describe() -> None:
				content = tableText(table)
				try:
					before = flatText(
						ctx.document.Range(max(0, table.Range.Start - 200), table.Range.Start).Text,
					)
				except COM_ERRORS:
					before = ""
				host.aiDraft(
					_("Alternative text"),
					_("AI description, edit it before saving:"),
					lambda: ai.describeTable(content, before),
					setDescription,
				)

			actions.append(
				Action(
					# Translators: Action that removes the header row of a table.
					_("Remove header row") if header else _("Make the first row a header row"),
					toggleHeader,
				),
			)
			# Translators: Action that edits the alternative text of a table.
			actions.append(Action(_("Edit alternative text..."), editDescription))
			if host.aiAvailable():
				# Translators: Action that asks AI to describe a table.
				actions.append(Action(_("Describe the table with AI..."), describe))
		return actions + super().actions(ctx, item, anchor, host)


# Links -----------------------------------------------------------------------------------------


class Links(ElementType):
	id = "links"
	vk = ord("K")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Links")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the link list.
		return [_("Text"), _("Address"), _("Status")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the link details.
			Section(_("Text"), multiline=True),
			# Translators: Section label in the link details.
			Section(_("Address"), multiline=True),
			# Translators: Section label in the link details.
			Section(_("Status")),
		]

	def collect(self, ctx, limit):
		items = []
		links = ctx.document.Hyperlinks
		for link in wordCom.iterate(links, wordCom.budgetOf(ctx)):
			if len(items) >= limit:
				break
			try:
				rangeObj = link.Range
			except COM_ERRORS:
				continue
			items.append(Item(obj=link, navRange=rangeObj, start=rangeObj.Start))
		items.sort(key=lambda i: i.start)
		return items, links.Count

	@staticmethod
	def _address(link: Any) -> str:
		try:
			address = str(link.Address or "")
			sub = str(link.SubAddress or "")
		except COM_ERRORS:
			return ""
		return f"{address}#{sub}" if address and sub else address or f"#{sub}" if sub else ""

	@staticmethod
	def _broken(ctx: WordContext, link: Any) -> bool:
		"""An internal link to a bookmark that no longer exists."""
		try:
			if link.Address or not link.SubAddress:
				return False
			bookmarks = ctx.document.Bookmarks
			hidden = bookmarks.ShowHidden
			bookmarks.ShowHidden = True
			try:
				return not bookmarks.Exists(str(link.SubAddress))
			finally:
				bookmarks.ShowHidden = hidden
		except COM_ERRORS:
			return False

	def summary(self, ctx, item):
		link = item.obj
		text = f"{flatText(link.TextToDisplay, 120)}, {self._address(link)}"
		if self._broken(ctx, link):
			# Translators: Spoken after a link whose target no longer exists.
			text += ", " + _("broken")
		return text

	def row(self, ctx, item):
		link = item.obj
		return [flatText(link.TextToDisplay, 100), self._address(link), _okOrError(self._broken(ctx, link))]

	def sectionTexts(self, ctx, item):
		link = item.obj
		return [cleanText(link.TextToDisplay), self._address(link), _okOrError(self._broken(ctx, link))]

	def copyText(self, ctx, item):
		return self._address(item.obj)

	def actions(self, ctx, item, anchor, host):
		actions: list[Action] = []
		if item is not None:
			link = item.obj

			def editText() -> None:
				text = host.prompt(_("Link text"), _("Text:"), cleanText(link.TextToDisplay))
				if text:
					with undoRecord(ctx, _("Link text")):
						link.TextToDisplay = text
					host.modified()
					host.speak(_("Link text changed"))

			def editAddress() -> None:
				# Translators: Title and label of the dialog for changing a link address.
				text = host.prompt(_("Link address"), _("Address:"), str(link.Address or ""))
				if text:
					# Translators: Name of the undo step for changing a link address.
					with undoRecord(ctx, _("Link address")):
						link.Address = text
					host.modified()
					# Translators: Reported after a link address changed.
					host.speak(_("Link address changed"))

			def remove() -> None:
				# Translators: Name of the undo step for removing a link.
				with undoRecord(ctx, _("Remove link")):
					link.Delete()
				host.modified()
				# Translators: Reported after a link was removed; its text stays.
				host.speak(_("Link removed, the text stays"))

			def follow() -> None:
				try:
					link.Follow()
				except COM_ERRORS:
					# Translators: Reported when a link cannot be opened.
					host.speak(_("Cannot open this link"))

			actions += [
				# Translators: Action that changes the text of a link.
				Action(_("Change link text..."), editText),
				# Translators: Action that changes the address of a link.
				Action(_("Change address..."), editAddress),
				# Translators: Action that opens a link.
				Action(_("Open the link"), follow),
				# Translators: Action that removes a link but keeps its text.
				Action(_("Remove link, keep the text"), remove),
			]
		return actions + super().actions(ctx, item, anchor, host)


# Bookmarks ------------------------------------------------------------------------------------

_BOOKMARK_NAME = re.compile(r"^[^\W\d_]\w{0,39}$")


class Bookmarks(ElementType):
	id = "bookmarks"
	vk = ord("B")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Bookmarks")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the bookmark list.
		return [_("Name"), _("Text")]

	@property
	def sections(self) -> list[Section]:
		# Translators: Section labels in the bookmark details.
		return [Section(_("Name")), Section(_("Text"), multiline=True)]

	def collect(self, ctx, limit):
		items = []
		bookmarks = ctx.document.Bookmarks
		for bookmark in wordCom.iterate(bookmarks, wordCom.budgetOf(ctx)):
			if len(items) >= limit:
				break
			items.append(Item(obj=bookmark, navRange=bookmark.Range, start=bookmark.Range.Start))
		items.sort(key=lambda i: i.start)
		return items, bookmarks.Count

	def summary(self, ctx, item):
		text = flatText(item.obj.Range.Text, 120)
		return f"{item.obj.Name}: {text}" if text else str(item.obj.Name)

	def row(self, ctx, item):
		return [str(item.obj.Name), flatText(item.obj.Range.Text, 120)]

	def sectionTexts(self, ctx, item):
		return [str(item.obj.Name), cleanText(item.obj.Range.Text)]

	def copyText(self, ctx, item):
		return str(item.obj.Name)

	def actions(self, ctx, item, anchor, host):
		actions: list[Action] = []
		if item is not None:
			bookmark = item.obj

			def delete() -> None:
				# Translators: Confirmation before deleting a bookmark.
				if host.confirm(
					_("Delete bookmark"),
					_("Delete bookmark {name}? The text stays.").format(name=bookmark.Name),
				):
					bookmark.Delete()
					host.modified()
					# Translators: Reported after a bookmark was deleted.
					host.speak(_("Bookmark deleted"))

			# Translators: Action that deletes a bookmark.
			actions.append(Action(_("Delete bookmark"), delete))

		def add() -> None:
			if anchor is None or not wordCom.isMainStory(anchor):
				# Translators: Reported when a bookmark cannot be added at the starting position.
				host.speak(_("A bookmark can only be added in the main text"))
				return
			name = host.prompt(
				# Translators: Title of the dialog for a new bookmark.
				_("New bookmark"),
				# Translators: Label of the bookmark name field.
				_("Name (letters and numbers, no spaces, starting with a letter):"),
			)
			if not name:
				return
			name = name.strip().replace(" ", "_")
			if not _BOOKMARK_NAME.match(name):
				# Translators: Reported when a bookmark name is not valid.
				host.speak(
					_(
						"Not a valid bookmark name. Use letters and numbers without spaces, starting with a letter",
					),
				)
				return
			ctx.document.Bookmarks.Add(name, anchor)
			host.modified()
			# Translators: Reported after a bookmark was added.
			host.speak(_("Bookmark {name} added").format(name=name))

		# Translators: Action that adds a bookmark where the caret was before navigation mode.
		actions.append(Action(_("Add bookmark at the starting position..."), add, enabled=anchor is not None))
		return actions + super().actions(ctx, item, anchor, host)


# Fields -------------------------------------------------------------------------------------------


class Fields(ElementType):
	id = "fields"
	vk = ord("D")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Fields")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the field list.
		return [_("Type"), _("Result"), _("Status"), _("Code")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the field details.
			Section(_("Type")),
			# Translators: Section label in the field details.
			Section(_("Code"), multiline=True),
			# Translators: Section label in the field details.
			Section(_("Result"), multiline=True),
			# Translators: Section label in the field details.
			Section(_("Status")),
		]

	def wanted(self, field: Any) -> bool:
		return True

	def collect(self, ctx, limit):
		items = []
		total = 0
		for field in wordCom.iterate(ctx.document.Fields, wordCom.budgetOf(ctx)):
			if not self.wanted(field):
				continue
			total += 1
			if len(items) < limit:
				try:
					rangeObj = field.Result if field.Result is not None else field.Code
					items.append(Item(obj=field, navRange=rangeObj, start=rangeObj.Start))
				except COM_ERRORS:
					continue
		items.sort(key=lambda i: i.start)
		return items, total

	def summary(self, ctx, item):
		field = item.obj
		text = f"{_fieldKind(field)}: {flatText(_fieldResult(field), 120)}"
		if _isError(field):
			text += ", " + _("Error")
		return text

	def row(self, ctx, item):
		field = item.obj
		return [
			_fieldKind(field),
			flatText(_fieldResult(field), 100),
			_okOrError(_isError(field)),
			_fieldCode(field),
		]

	def sectionTexts(self, ctx, item):
		field = item.obj
		return [_fieldKind(field), _fieldCode(field), _fieldResult(field), _okOrError(_isError(field))]

	def copyText(self, ctx, item):
		return _fieldResult(item.obj)

	def actions(self, ctx, item, anchor, host):
		actions: list[Action] = []
		if item is not None:
			field = item.obj

			def update() -> None:
				field.Update()
				host.modified()
				# Translators: Reported after a field was updated.
				host.speak(_("Field updated: {result}").format(result=flatText(_fieldResult(field), 120)))

			def unlink() -> None:
				# Translators: Confirmation before turning a field into plain text.
				if host.confirm(
					_("Convert to text"),
					_("Replace this field with its current result as plain text?"),
				):
					# Translators: Name of the undo step for turning a field into text.
					with undoRecord(ctx, _("Convert to text")):
						field.Unlink()
					host.modified()
					# Translators: Reported after a field became plain text.
					host.speak(_("Converted to text"))

			# Translators: Action that updates one field.
			actions.append(Action(_("Update this field"), update))
			# Translators: Action that copies the code of a field.
			actions.append(Action(_("Copy field code"), lambda: host.copy(_fieldCode(field))))
			# Translators: Action that turns a field into plain text.
			actions.append(Action(_("Convert to plain text..."), unlink))

		def updateAll() -> None:
			doc = ctx.document
			doc.Fields.Update()
			try:
				for index in range(1, doc.TablesOfContents.Count + 1):
					doc.TablesOfContents.Item(index).Update()
			except COM_ERRORS:
				pass
			host.modified()
			# Translators: Reported after all fields were updated.
			host.speak(_("All fields and tables of contents updated"))

		# Translators: Action that updates every field in the document.
		actions.append(Action(_("Update all fields and tables of contents"), updateAll))
		return actions + super().actions(ctx, item, anchor, host)


# Cross-references -----------------------------------------------------------------------------

REF_KINDS = ("REF", "PAGEREF", "NOTEREF")

# Word reference types and kinds
REF_TYPE_NUMBERED = 0
REF_TYPE_HEADING = 1
REF_TYPE_BOOKMARK = 2
REF_TYPE_FOOTNOTE = 3
REF_TYPE_ENDNOTE = 4
KIND_TEXT = -1
KIND_NUMBER_RELATIVE = -2
KIND_ENTIRE_CAPTION = 2
KIND_LABEL_AND_NUMBER = 3
KIND_CAPTION_TEXT = 4
KIND_FOOTNOTE_NUMBER = 5
KIND_ENDNOTE_NUMBER = 6
KIND_PAGE = 7
KIND_POSITION = 15


def _refTarget(field: Any) -> str:
	parts = _fieldCode(field).split()
	return parts[1] if len(parts) > 1 else ""


class CrossReferences(Fields):
	id = "crossReferences"
	vk = ord("X")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Cross-references")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the cross-reference list.
		return [_("Shows"), _("Points to"), _("Status")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the cross-reference details.
			Section(_("Shows"), multiline=True),
			# Translators: Section label in the cross-reference details.
			Section(_("Points to"), multiline=True),
			# Translators: Section label in the cross-reference details.
			Section(_("Status")),
			# Translators: Section label in the cross-reference details.
			Section(_("Code")),
		]

	def wanted(self, field):
		return _fieldKind(field) in REF_KINDS

	def _targetText(self, ctx: WordContext, field: Any) -> str:
		name = _refTarget(field)
		if not name:
			return ""
		try:
			bookmarks = ctx.document.Bookmarks
			hidden = bookmarks.ShowHidden
			bookmarks.ShowHidden = True
			try:
				if not bookmarks.Exists(name):
					# Translators: Shown when the target of a cross-reference no longer exists.
					return _("(target not found)")
				return flatText(bookmarks.Item(name).Range.Paragraphs.Item(1).Range.Text, 150)
			finally:
				bookmarks.ShowHidden = hidden
		except COM_ERRORS:
			return name

	def summary(self, ctx, item):
		field = item.obj
		# Translators: Spoken for a cross-reference, for example "3, points to Results".
		text = _("{shows}, points to {target}").format(
			shows=flatText(_fieldResult(field), 80),
			target=self._targetText(ctx, field),
		)
		if _isError(field):
			text += ", " + _("Error")
		return text

	def row(self, ctx, item):
		field = item.obj
		return [flatText(_fieldResult(field), 80), self._targetText(ctx, field), _okOrError(_isError(field))]

	def sectionTexts(self, ctx, item):
		field = item.obj
		return [
			_fieldResult(field),
			self._targetText(ctx, field),
			_okOrError(_isError(field)),
			_fieldCode(field),
		]

	def actions(self, ctx, item, anchor, host):
		actions: list[Action] = []
		if item is not None:
			field = item.obj

			def gotoTarget() -> None:
				name = _refTarget(field)
				bookmarks = ctx.document.Bookmarks
				bookmarks.ShowHidden = True
				if not name or not bookmarks.Exists(name):
					host.speak(_("(target not found)"))
					return
				target = bookmarks.Item(name).Range
				target.Select()
				# Translators: Reported after moving to the target of a cross-reference.
				host.speak(
					_("Target: {text}").format(text=flatText(target.Paragraphs.Item(1).Range.Text, 150)),
				)

			# Translators: Action that moves to the target of a cross-reference.
			actions.append(Action(_("Go to the target"), gotoTarget))
		actions.append(
			Action(
				# Translators: Action that inserts a cross-reference where the caret was.
				_("Insert cross-reference at the starting position..."),
				lambda: insertCrossReference(ctx, anchor, host),
				enabled=anchor is not None,
			),
		)
		return actions + super().actions(ctx, item, anchor, host)


def _captionLabels(ctx: WordContext) -> list[str]:
	labels = []
	try:
		captionLabels = ctx.application.CaptionLabels
		for index in range(1, captionLabels.Count + 1):
			labels.append(str(captionLabels.Item(index).Name))
	except COM_ERRORS:
		pass
	return labels


def insertCrossReference(ctx: WordContext, anchor: Any, host: Host) -> None:
	if anchor is None or not wordCom.isMainStory(anchor):
		# Translators: Reported when a cross-reference cannot be inserted at the starting position.
		host.speak(_("A cross-reference can only be inserted in the main text"))
		return
	captions = _captionLabels(ctx)
	types: list[tuple[str, Any, list[tuple[str, int]]]] = [
		(
			# Translators: Kind of item a cross-reference can point to.
			_("Heading"),
			REF_TYPE_HEADING,
			[
				# Translators: What a cross-reference shows.
				(_("Heading text"), KIND_TEXT),
				# Translators: What a cross-reference shows.
				(_("Page number"), KIND_PAGE),
				# Translators: What a cross-reference shows.
				(_("Heading number"), KIND_NUMBER_RELATIVE),
				# Translators: What a cross-reference shows.
				(_("Above or below"), KIND_POSITION),
			],
		),
		(
			# Translators: Kind of item a cross-reference can point to.
			_("Bookmark"),
			REF_TYPE_BOOKMARK,
			[
				# Translators: What a cross-reference shows.
				(_("Bookmark text"), KIND_TEXT),
				(_("Page number"), KIND_PAGE),
				(_("Above or below"), KIND_POSITION),
			],
		),
		(
			# Translators: Kind of item a cross-reference can point to.
			_("Footnote"),
			REF_TYPE_FOOTNOTE,
			[
				# Translators: What a cross-reference shows.
				(_("Footnote number"), KIND_FOOTNOTE_NUMBER),
				(_("Page number"), KIND_PAGE),
			],
		),
		(
			# Translators: Kind of item a cross-reference can point to.
			_("Endnote"),
			REF_TYPE_ENDNOTE,
			[
				# Translators: What a cross-reference shows.
				(_("Endnote number"), KIND_ENDNOTE_NUMBER),
				(_("Page number"), KIND_PAGE),
			],
		),
		(
			# Translators: Kind of item a cross-reference can point to.
			_("Numbered item"),
			REF_TYPE_NUMBERED,
			[
				# Translators: What a cross-reference shows.
				(_("Paragraph number"), KIND_NUMBER_RELATIVE),
				# Translators: What a cross-reference shows.
				(_("Paragraph text"), KIND_TEXT),
				(_("Page number"), KIND_PAGE),
			],
		),
	]
	for label in captions:
		types.append(
			(
				# Translators: Kind of item a cross-reference can point to, for example "Caption: Table".
				_("Caption: {label}").format(label=label),
				label,
				[
					# Translators: What a cross-reference shows.
					(_("Entire caption"), KIND_ENTIRE_CAPTION),
					# Translators: What a cross-reference shows.
					(_("Only label and number"), KIND_LABEL_AND_NUMBER),
					# Translators: What a cross-reference shows.
					(_("Only caption text"), KIND_CAPTION_TEXT),
					(_("Page number"), KIND_PAGE),
				],
			),
		)
	# Translators: Title of the cross-reference dialogs.
	title = _("Insert cross-reference")
	# Translators: Label of the choice of what a cross-reference points to.
	typeIndex = host.choose(title, _("Point to:"), [name for name, _t, _k in types])
	if typeIndex is None:
		return
	typeName, referenceType, kinds = types[typeIndex]
	try:
		targets = list(ctx.document.GetCrossReferenceItems(referenceType) or [])
	except COM_ERRORS:
		targets = []
	targets = [flatText(t, 150) for t in targets]
	if not targets:
		# Translators: Reported when the document has nothing of the chosen kind to point to.
		host.speak(_("The document has no {kind} to point to").format(kind=typeName))
		return
	# Translators: Label of the list of targets for a cross-reference.
	targetIndex = host.choose(title, _("Which item:"), targets)
	if targetIndex is None:
		return
	# Translators: Label of the choice of what the cross-reference shows.
	kindIndex = host.choose(title, _("Show:"), [name for name, _value in kinds])
	if kindIndex is None:
		return
	position = anchor.Duplicate
	position.Collapse(WD_COLLAPSE_END)
	# Translators: Name of the undo step for inserting a cross-reference.
	with undoRecord(ctx, _("Insert cross-reference")):
		position.InsertCrossReference(referenceType, kinds[kindIndex][1], targetIndex + 1, True)
	host.modified()
	# Translators: Reported after a cross-reference was inserted.
	host.speak(_("Cross-reference inserted"))


# Citations -----------------------------------------------------------------------------------------

BIB_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
CITATION_STYLES = ("APA", "MLA", "Chicago", "IEEE", "Harvard - Anglia", "ISO 690 - Numerical Reference")
SOURCE_TYPES = (
	("Book", "Book"),
	("JournalArticle", "Journal article"),
	("InternetSite", "Web site"),
	("Report", "Report"),
)


def _sourceField(source: Any, name: str) -> str:
	try:
		value = str(source.Field(name) or "")
	except COM_ERRORS:
		return ""
	# Author comes back as XML; keep the readable names.
	if value.startswith("<"):
		names = re.findall(r"<b:(?:Last|First|Corporate)>([^<]*)</b:", value)
		return " ".join(names)
	return value


def _sources(ctx: WordContext) -> list[Any]:
	try:
		sources = ctx.document.Bibliography.Sources
		return [sources.Item(i) for i in range(1, sources.Count + 1)]
	except COM_ERRORS:
		return []


def _sourceLabel(source: Any) -> str:
	parts = [_sourceField(source, "Author"), _sourceField(source, "Year"), _sourceField(source, "Title")]
	return ", ".join(p for p in parts if p) or str(getattr(source, "Tag", ""))


def sourceXml(tag: str, sourceType: str, author: str, title: str, year: str, publisher: str, url: str) -> str:
	people = []
	for name in [n.strip() for n in re.split(r";|\band\b|และ", author) if n.strip()]:
		pieces = name.split()
		last = pieces[-1] if len(pieces) > 1 else name
		first = " ".join(pieces[:-1]) if len(pieces) > 1 else ""
		people.append(
			f"<b:Person><b:Last>{escape(last)}</b:Last><b:First>{escape(first)}</b:First></b:Person>",
		)
	authorXml = (
		f"<b:Author><b:Author><b:NameList>{''.join(people)}</b:NameList></b:Author></b:Author>"
		if people
		else ""
	)
	extra = ""
	if publisher:
		extra += f"<b:Publisher>{escape(publisher)}</b:Publisher>"
	if url:
		extra += f"<b:URL>{escape(url)}</b:URL>"
	return (
		f'<b:Source xmlns:b="{BIB_NS}"><b:Tag>{escape(tag)}</b:Tag><b:SourceType>{sourceType}</b:SourceType>'
		f"{authorXml}<b:Title>{escape(title)}</b:Title><b:Year>{escape(year)}</b:Year>{extra}</b:Source>"
	)


def _newTag(author: str, year: str, existing: set[str]) -> str:
	base = re.sub(r"\W", "", (author.split()[-1] if author.split() else "Source"))[:10] or "Source"
	base += re.sub(r"\D", "", year)[-2:]
	tag, number = base, 1
	while tag in existing:
		number += 1
		tag = f"{base}_{number}"
	return tag


def _ignoreResult(_value: object) -> None:
	"""Run an action for its effect only; actions return nothing."""


def addSource(ctx: WordContext, host: Host) -> str | None:
	"""Ask for the details of a new source and add it. Returns its tag."""
	# Translators: Title of the dialogs for a new source.
	title = _("New source")
	typeNames = [
		# Translators: Kind of source.
		_("Book"),
		# Translators: Kind of source.
		_("Journal article"),
		# Translators: Kind of source.
		_("Web site"),
		# Translators: Kind of source.
		_("Report"),
	]
	# Translators: Label of the choice of source kind.
	typeIndex = host.choose(title, _("Kind of source:"), typeNames)
	if typeIndex is None:
		return None
	# Translators: Label of the source author field.
	author = host.prompt(title, _("Author or authors (separate with ;):"))
	if author is None:
		return None
	# Translators: Label of the source title field.
	sourceTitle = host.prompt(title, _("Title:"))
	if not sourceTitle:
		return None
	# Translators: Label of the source year field.
	year = host.prompt(title, _("Year:")) or ""
	if SOURCE_TYPES[typeIndex][0] == "InternetSite":
		# Translators: Label of the source web address field.
		url = host.prompt(title, _("Web address:")) or ""
		publisher = ""
	else:
		# Translators: Label of the source publisher field.
		publisher = host.prompt(title, _("Publisher or journal:")) or ""
		url = ""
	existing = {str(getattr(s, "Tag", "")) for s in _sources(ctx)}
	tag = _newTag(author, year, existing)
	xml = sourceXml(tag, SOURCE_TYPES[typeIndex][0], author, sourceTitle, year, publisher, url)
	ctx.document.Bibliography.Sources.Add(xml)
	# Translators: Reported after a source was added.
	host.speak(_("Source added: {tag}").format(tag=tag))
	return tag


def insertCitation(ctx: WordContext, anchor: Any, host: Host) -> None:
	if anchor is None or not wordCom.isMainStory(anchor):
		# Translators: Reported when a citation cannot be inserted at the starting position.
		host.speak(_("A citation can only be inserted in the main text"))
		return
	sources = _sources(ctx)
	# Translators: Choice in the citation dialog that adds a new source.
	options = [_sourceLabel(s) for s in sources] + [_("New source...")]
	# Translators: Title and label of the dialog for inserting a citation.
	index = host.choose(_("Insert citation"), _("Source:"), options)
	if index is None:
		return
	tag = str(sources[index].Tag) if index < len(sources) else addSource(ctx, host)
	if not tag:
		return
	position = anchor.Duplicate
	position.Collapse(WD_COLLAPSE_END)
	# Translators: Name of the undo step for inserting a citation.
	with undoRecord(ctx, _("Insert citation")):
		ctx.document.Fields.Add(position, WD_FIELD_EMPTY, f"CITATION {tag}", False)
	host.modified()
	# Translators: Reported after a citation was inserted.
	host.speak(_("Citation inserted"))


def updateBibliography(ctx: WordContext, anchor: Any, host: Host) -> None:
	doc = ctx.document
	for field in wordCom.iterate(doc.Fields):
		if _fieldKind(field) == "BIBLIOGRAPHY":
			field.Update()
			host.modified()
			# Translators: Reported after the bibliography was updated.
			host.speak(_("Bibliography updated"))
			return
	if anchor is None or not wordCom.isMainStory(anchor):
		# Translators: Reported when the bibliography cannot be inserted at the starting position.
		host.speak(_("Place the caret where the bibliography should go, then try again"))
		return
	position = anchor.Duplicate
	position.Collapse(WD_COLLAPSE_END)
	# Translators: Name of the undo step for inserting a bibliography.
	with undoRecord(ctx, _("Insert bibliography")):
		doc.Fields.Add(position, WD_FIELD_EMPTY, "BIBLIOGRAPHY", False)
	host.modified()
	# Translators: Reported after the bibliography was inserted.
	host.speak(_("Bibliography inserted at the starting position"))


def setCitationStyle(ctx: WordContext, host: Host) -> None:
	# Translators: Title and label of the dialog for the citation style.
	index = host.choose(_("Citation style"), _("Style:"), list(CITATION_STYLES))
	if index is None:
		return
	style = CITATION_STYLES[index]
	try:
		ctx.document.Bibliography.BibliographyStyle = style
	except COM_ERRORS:
		# Translators: Reported when Word does not have the chosen citation style.
		host.speak(_("Word does not have the style {style}").format(style=style))
		return
	updateAllCitations(ctx)
	host.modified()
	# Translators: Reported after the citation style changed.
	host.speak(_("Citation style {style}").format(style=style))


def updateAllCitations(ctx: WordContext) -> None:
	for field in wordCom.iterate(ctx.document.Fields):
		if _fieldKind(field) in ("CITATION", "BIBLIOGRAPHY"):
			try:
				field.Update()
			except COM_ERRORS:
				continue


class Citations(Fields):
	id = "citations"
	vk = ord("I")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Citations")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the citation list.
		return [_("Shows"), _("Source")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the citation details.
			Section(_("Shows")),
			# Translators: Section label in the citation details.
			Section(_("Source"), multiline=True),
			# Translators: Section label in the citation details.
			Section(_("Code")),
		]

	def wanted(self, field):
		return _fieldKind(field) == "CITATION"

	def _source(self, ctx: WordContext, field: Any) -> str:
		parts = _fieldCode(field).split()
		tag = parts[1] if len(parts) > 1 else ""
		for source in _sources(ctx):
			if str(getattr(source, "Tag", "")) == tag:
				return _sourceLabel(source)
		return tag

	def summary(self, ctx, item):
		return f"{flatText(_fieldResult(item.obj), 80)}, {self._source(ctx, item.obj)}"

	def row(self, ctx, item):
		return [flatText(_fieldResult(item.obj), 80), self._source(ctx, item.obj)]

	def sectionTexts(self, ctx, item):
		return [_fieldResult(item.obj), self._source(ctx, item.obj), _fieldCode(item.obj)]

	def actions(self, ctx, item, anchor, host):
		actions = [
			Action(
				# Translators: Action that inserts a citation where the caret was.
				_("Insert citation at the starting position..."),
				lambda: insertCitation(ctx, anchor, host),
				enabled=anchor is not None,
			),
			# Translators: Action that adds a new source to the document.
			Action(_("Add a new source..."), lambda: _ignoreResult(addSource(ctx, host))),
			# Translators: Action that inserts or updates the bibliography.
			Action(_("Insert or update the bibliography"), lambda: updateBibliography(ctx, anchor, host)),
			# Translators: Action that changes the citation style.
			Action(_("Change citation style..."), lambda: setCitationStyle(ctx, host)),
		]
		return actions + super().actions(ctx, item, anchor, host)

# Word Document Assistant for NVDA
# Access to the Microsoft Word object model (COM).
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import re
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime
from typing import Any

import addonHandler
import api
import comtypes.client
import oleacc
import winUser
from comtypes import COMError
from comtypes.automation import IDispatch
from logHandler import log

addonHandler.initTranslation()

OBJID_NATIVEOM = getattr(winUser, "OBJID_NATIVEOM", -16)
# Word names its document window "_WwG". Compared without case to be safe.
WORD_DOCUMENT_WINDOW_CLASS = "_wwg"
GA_PARENT = 1

# Word constants
WD_COLLAPSE_START = 1
WD_COLLAPSE_END = 0
WD_MAIN_TEXT_STORY = 1
WD_FOOTNOTES_STORY = 2
WD_ENDNOTES_STORY = 3
WD_COMMENTS_STORY = 4
WD_INFO_SECTION = 2
WD_INFO_PAGE = 3
WD_INFO_PAGE_COUNT = 4
WD_INFO_LINE = 10
WD_GOTO_HEADING = 11
WD_GOTO_PREVIOUS = 3
WD_OUTLINE_BODY_TEXT = 10
WD_STAT_WORDS = 0
WD_STAT_PAGES = 2
WD_STAT_PARAGRAPHS = 4

# Errors raised by the Word object model.
COM_ERRORS = (COMError, OSError, AttributeError)


class WordUnavailable(Exception):
	"""Raised when the Word document cannot be reached through COM."""


def isWordFocused() -> bool:
	focus = api.getFocusObject()
	appModule = getattr(focus, "appModule", None)
	return bool(appModule and appModule.appName == "winword")


def isDocumentFocused() -> bool:
	"""True when the focus is in the text of a Word document window."""
	if not isWordFocused():
		return False
	focus = api.getFocusObject()
	return _isDocumentWindow(getattr(focus, "windowClassName", "")) or bool(
		_findDocumentWindow(focus, searchForeground=False)
	)


def _isDocumentWindow(className: str) -> bool:
	return (className or "").lower() == WORD_DOCUMENT_WINDOW_CLASS


def _findDocumentWindow(focus: Any, searchForeground: bool = True) -> int:
	"""Handle of the Word document window for the focus, or 0."""
	hwnd = getattr(focus, "windowHandle", 0) or 0
	# The focus may be a child of the document window, so walk up the parents.
	for _step in range(10):
		if not hwnd:
			break
		try:
			if _isDocumentWindow(winUser.getClassName(hwnd)):
				return hwnd
			hwnd = winUser.getAncestor(hwnd, GA_PARENT)
		except Exception:  # noqa: BLE001
			break
	if searchForeground:
		try:
			import windowUtils

			foreground = api.getForegroundObject()
			return windowUtils.findDescendantWindow(foreground.windowHandle, className="_WwG")
		except Exception:  # noqa: BLE001
			log.debugWarning("No Word document window found", exc_info=True)
	return 0


def _getWindowObject(hwnd: int) -> Any:
	pDispatch = oleacc.AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, interface=IDispatch)
	return comtypes.client.dynamic.Dispatch(pDispatch)


class WordContext:
	"""The Word window, document and application behind the focused document."""

	def __init__(self, window: Any):
		self.window = window
		self.document = window.Document
		self.application = window.Application

	@classmethod
	def fromFocus(cls) -> WordContext:
		if not isWordFocused():
			# Translators: Reported when a command is used outside Microsoft Word.
			raise WordUnavailable(_("This command is only available in Microsoft Word"))
		focus = api.getFocusObject()
		hwnd = _findDocumentWindow(focus)
		if not hwnd:
			log.debugWarning(
				f"Focus is not in a Word document: class {getattr(focus, 'windowClassName', '')!r}, "
				f"role {getattr(focus, 'role', '')!r}",
			)
			# Translators: Reported when the focus is in Word but not in the document text.
			raise WordUnavailable(_("Move the focus to the document text first"))
		try:
			return cls(_getWindowObject(hwnd))
		except COM_ERRORS as error:
			log.debugWarning("Cannot reach the Word object model", exc_info=True)
			raise WordUnavailable(
				# Translators: Reported when the Word object model cannot be reached.
				_(
					"Cannot access this document. If it is open in Protected View, "
					"enable editing first and try again",
				),
			) from error

	@property
	def selection(self) -> Any:
		return self.window.Selection

	@property
	def version(self) -> float:
		try:
			return float(self.application.Version)
		except (*COM_ERRORS, ValueError):
			return 0.0


@contextmanager
def undoRecord(ctx: WordContext, name: str) -> Iterator[None]:
	"""Group every change made inside the block into one Word undo step."""
	record = None
	try:
		record = ctx.application.UndoRecord
		record.StartCustomRecord(name)
	except COM_ERRORS:
		record = None
	try:
		yield
	finally:
		if record is not None:
			try:
				record.EndCustomRecord()
			except COM_ERRORS:
				pass


_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0c\x0e-\x1f]")


def cleanText(text: Any) -> str:
	"""Turn Word story text into plain multi-line text."""
	if not text:
		return ""
	text = str(text)
	text = text.replace("\r\x07", " ").replace("\x07", " ")
	text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\x0b", "\n")
	text = _CONTROL_CHARS.sub("", text)
	return text.strip()


def sameLengthText(text: Any) -> str:
	"""Readable copy of Word range text with exactly the same length, so offsets still match."""
	text = str(text or "")
	return "".join("\n" if ch in "\r\x0b" else (" " if ord(ch) < 32 and ch != "\n" else ch) for ch in text)


def flatText(text: Any, maxLength: int = 0) -> str:
	"""Plain single-line text suitable for speech."""
	result = " ".join(cleanText(text).split())
	if maxLength and len(result) > maxLength:
		result = result[:maxLength].rstrip() + "…"
	return result


def formatDate(value: Any) -> str:
	if isinstance(value, datetime):
		return value.strftime("%d/%m/%Y %H:%M")
	return str(value) if value else ""


def mainStoryPosition(ctx: WordContext) -> int:
	"""Position of the caret in the main text, also when the caret is inside a note or comment."""
	sel = ctx.selection
	try:
		story = sel.StoryType
		if story == WD_FOOTNOTES_STORY:
			return sel.Footnotes.Item(1).Reference.Start
		if story == WD_ENDNOTES_STORY:
			return sel.Endnotes.Item(1).Reference.Start
		if story == WD_COMMENTS_STORY:
			return sel.Comments.Item(1).Scope.Start
	except COM_ERRORS:
		return 0
	return sel.Start


def isMainStory(rangeObj: Any) -> bool:
	try:
		return rangeObj.StoryType == WD_MAIN_TEXT_STORY
	except COM_ERRORS:
		return False


def noteAtCaret(ctx: WordContext) -> str | None:
	"""Describe the footnote or endnote at or next to the caret."""
	sel = ctx.selection
	story = sel.StoryType
	note = None
	# Translators: Name of a footnote when it is reported.
	footnoteLabel = _("Footnote")
	# Translators: Name of an endnote when it is reported.
	endnoteLabel = _("Endnote")
	label = footnoteLabel
	if story == WD_FOOTNOTES_STORY:
		note = sel.Footnotes.Item(1)
	elif story == WD_ENDNOTES_STORY:
		note, label = sel.Endnotes.Item(1), endnoteLabel
	elif story == WD_MAIN_TEXT_STORY:
		doc = ctx.document
		start = max(0, sel.Start - 1)
		end = min(sel.End + 1, doc.Content.End)
		around = doc.Range(start, end)
		if around.Footnotes.Count:
			note = around.Footnotes.Item(1)
		elif around.Endnotes.Count:
			note, label = around.Endnotes.Item(1), endnoteLabel
	if note is None:
		return None
	return f"{label} {note.Index}: {flatText(note.Range.Text)}"


def currentHeading(ctx: WordContext) -> str:
	doc = ctx.document
	pos = ctx.selection.Start
	here = doc.Range(pos, pos)
	para = here.Paragraphs.Item(1)
	if para.OutlineLevel < WD_OUTLINE_BODY_TEXT:
		return flatText(para.Range.Text, 120)
	previous = here.GoTo(WD_GOTO_HEADING, WD_GOTO_PREVIOUS)
	if previous is None or previous.Start >= pos:
		return ""
	para = previous.Paragraphs.Item(1)
	if para.OutlineLevel < WD_OUTLINE_BODY_TEXT:
		return flatText(para.Range.Text, 120)
	return ""


def positionInfo(ctx: WordContext) -> str:
	sel = ctx.selection
	parts = []
	try:
		# Translators: Reports the page of the caret, for example "Page 2 of 10".
		parts.append(
			_("Page {page} of {pages}").format(
				page=sel.Information(WD_INFO_PAGE),
				pages=sel.Information(WD_INFO_PAGE_COUNT),
			)
		)
		# Translators: Reports the section number of the caret.
		parts.append(_("section {number}").format(number=sel.Information(WD_INFO_SECTION)))
		line = sel.Information(WD_INFO_LINE)
		if isinstance(line, int) and line > 0:
			# Translators: Reports the line number of the caret.
			parts.append(_("line {number}").format(number=line))
	except COM_ERRORS:
		log.debugWarning("Cannot read position information", exc_info=True)
	try:
		heading = currentHeading(ctx)
	except COM_ERRORS:
		heading = ""
	if heading:
		# Translators: Reports the heading the caret is under.
		parts.append(_("under heading {heading}").format(heading=heading))
	return ", ".join(str(p) for p in parts)


def _markupName(markup: int) -> str:
	names = {
		# Translators: Word markup view without any change marks.
		0: _("no markup"),
		# Translators: Word "Simple Markup" view.
		1: _("simple markup"),
		# Translators: Word "All Markup" view.
		2: _("all markup"),
	}
	return names.get(markup, "")


def trackChangesStatus(ctx: WordContext) -> str:
	doc = ctx.document
	if doc.TrackRevisions:
		# Translators: Reported when Track Changes is on.
		parts = [_("Track changes on")]
	else:
		# Translators: Reported when Track Changes is off.
		parts = [_("Track changes off")]
	try:
		markup = _markupName(ctx.window.View.RevisionsFilter.Markup)
		if markup:
			# Translators: Reports the markup view, for example "showing all markup".
			parts.append(_("showing {markup}").format(markup=markup))
	except COM_ERRORS:
		pass
	# Translators: Reports how many tracked changes the document has.
	parts.append(_("{count} revisions").format(count=doc.Revisions.Count))
	return ", ".join(parts)


def toggleTrackChanges(ctx: WordContext) -> str:
	doc = ctx.document
	doc.TrackRevisions = not doc.TrackRevisions
	return trackChangesStatus(ctx)


def documentSummary(ctx: WordContext) -> str:
	doc = ctx.document
	content = doc.Content
	# Translators: Summary of the document. Each {name} is a number.
	stats = _("{words} words, {pages} pages, {paragraphs} paragraphs.").format(
		words=doc.ComputeStatistics(WD_STAT_WORDS),
		pages=doc.ComputeStatistics(WD_STAT_PAGES),
		paragraphs=doc.ComputeStatistics(WD_STAT_PARAGRAPHS),
	)
	# Translators: Counts of review items in the document. Each {name} is a number.
	counts = _(
		"Spelling errors {spelling}, grammar issues {grammar}, comments {comments}, "
		"footnotes {footnotes}, endnotes {endnotes}, revisions {revisions}.",
	).format(
		spelling=content.SpellingErrors.Count,
		grammar=content.GrammaticalErrors.Count,
		comments=doc.Comments.Count,
		footnotes=doc.Footnotes.Count,
		endnotes=doc.Endnotes.Count,
		revisions=doc.Revisions.Count,
	)
	if doc.TrackRevisions:
		# Translators: Reported when Track Changes is on.
		track = _("Track changes on")
	else:
		# Translators: Reported when Track Changes is off.
		track = _("Track changes off")
	return f"{stats} {counts} {track}."

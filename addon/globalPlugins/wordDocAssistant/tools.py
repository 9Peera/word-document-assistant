# Word Document Assistant for NVDA
# Thai text tools and AI writing tools.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import re
from typing import Any

import addonHandler
import wx
from gui import guiHelper

from .wordCom import COM_ERRORS, WordContext, cleanText, undoRecord

addonHandler.initTranslation()

THAI_DIGITS = "๐๑๒๓๔๕๖๗๘๙"
ARABIC_DIGITS = "0123456789"
TO_THAI = str.maketrans(ARABIC_DIGITS, THAI_DIGITS)
TO_ARABIC = str.maketrans(THAI_DIGITS, ARABIC_DIGITS)
BE_OFFSET = 543
WD_REPLACE_ALL = 2
WD_FIND_STOP = 0

# Years with an era: "ค.ศ. 2026", "คศ 2026", "พ.ศ. 2569", in Arabic or Thai digits.
_ERA_YEAR = re.compile(r"(ค\.\s?ศ\.|คศ\.?|พ\.\s?ศ\.|พศ\.?)(\s*)([0-9๐-๙]{4})")


def countDigits(text: str) -> tuple[int, int]:
	thai = sum(text.count(d) for d in THAI_DIGITS)
	arabic = sum(text.count(d) for d in ARABIC_DIGITS)
	return thai, arabic


def _replaceAll(rangeObj: Any, find: str, replace: str) -> None:
	finder = rangeObj.Duplicate.Find
	finder.ClearFormatting()
	finder.Replacement.ClearFormatting()
	finder.Execute(find, True, False, False, False, False, True, WD_FIND_STOP, False, replace, WD_REPLACE_ALL)


def convertDigits(ctx: WordContext, rangeObj: Any, toThai: bool) -> int:
	"""Change every digit in the range. Returns the number of digits changed. One undo step."""
	before = countDigits(str(rangeObj.Text or ""))
	changed = before[1] if toThai else before[0]
	if not changed:
		return 0
	source, target = (ARABIC_DIGITS, THAI_DIGITS) if toThai else (THAI_DIGITS, ARABIC_DIGITS)
	# Translators: Name of the undo step for changing digits.
	with undoRecord(ctx, _("Change digits")):
		for old, new in zip(source, target, strict=True):
			_replaceAll(rangeObj, old, new)
	return changed


def eraReplacements(text: str, toBuddhist: bool) -> list[tuple[str, str]]:
	"""Distinct (old, new) pairs that change years between Christian and Buddhist eras."""
	pairs: dict[str, str] = {}
	for match in _ERA_YEAR.finditer(text):
		era, space, digits = match.groups()
		isChristian = era.startswith("ค")
		if isChristian != toBuddhist:
			continue
		thaiDigits = any(d in THAI_DIGITS for d in digits)
		year = int(digits.translate(TO_ARABIC))
		newYear = year + BE_OFFSET if toBuddhist else year - BE_OFFSET
		newDigits = str(newYear).translate(TO_THAI) if thaiDigits else str(newYear)
		newEra = ("พ.ศ." if "." in era else "พศ") if toBuddhist else ("ค.ศ." if "." in era else "คศ")
		pairs[match.group(0)] = f"{newEra}{space}{newDigits}"
	return list(pairs.items())


def convertEra(ctx: WordContext, rangeObj: Any, toBuddhist: bool) -> int:
	pairs = eraReplacements(str(rangeObj.Text or ""), toBuddhist)
	if not pairs:
		return 0
	# Translators: Name of the undo step for changing years between eras.
	with undoRecord(ctx, _("Change years")):
		for old, new in pairs:
			_replaceAll(rangeObj, old, new)
	return len(pairs)


# Dialogs ----------------------------------------------------------------------------------------


class ResultDialog(wx.Dialog):
	"""Read-only result, for example a summary, with a copy button."""

	def __init__(self, parent: wx.Window, title: str, text: str):
		super().__init__(parent, title=title)
		sizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		self.text = helper.addLabeledControl(
			# Translators: Label of the read-only field with an AI result.
			_("&Result:"),
			wx.TextCtrl,
			value=text,
			style=wx.TE_READONLY | wx.TE_MULTILINE | wx.TE_RICH2,
			size=(640, 300),
		)
		buttons = guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: Button that copies an AI result.
		copyButton = buttons.addButton(self, label=_("&Copy"))
		# Translators: Button that closes a dialog.
		buttons.addButton(self, wx.ID_CLOSE, label=_("Clo&se"))
		helper.addItem(buttons)
		sizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizerAndFit(sizer)
		self.SetEscapeId(wx.ID_CLOSE)
		copyButton.Bind(wx.EVT_BUTTON, lambda evt: self._copy())
		self.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(wx.ID_CLOSE), id=wx.ID_CLOSE)
		self.text.SetFocus()
		self.CentreOnScreen()

	def _copy(self) -> None:
		import api
		import ui

		if api.copyToClip(self.text.GetValue()):
			ui.message(_("Copied"))


#: Results of the compare dialog.
REPLACE = wx.ID_HIGHEST + 1
INSERT_AFTER = wx.ID_HIGHEST + 2


class CompareDialog(wx.Dialog):
	"""Original text and the AI text side by side; the AI text can be edited before it is used."""

	def __init__(self, parent: wx.Window, title: str, original: str, proposed: str, canInsert: bool):
		super().__init__(parent, title=title)
		sizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		self.original = helper.addLabeledControl(
			# Translators: Label of the read-only field with the original text.
			_("&Original:"),
			wx.TextCtrl,
			value=original,
			style=wx.TE_READONLY | wx.TE_MULTILINE | wx.TE_RICH2,
			size=(640, 140),
		)
		self.proposed = helper.addLabeledControl(
			# Translators: Label of the editable field with the AI text.
			_("&New text (you can edit it):"),
			wx.TextCtrl,
			value=proposed,
			style=wx.TE_MULTILINE | wx.TE_RICH2,
			size=(640, 160),
		)
		buttons = guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: Button that replaces the original text with the new text.
		replaceButton = buttons.addButton(self, REPLACE, label=_("&Replace the original"))
		if canInsert:
			# Translators: Button that inserts the new text after the original.
			buttons.addButton(self, INSERT_AFTER, label=_("&Insert after the original"))
		# Translators: Button that copies the new text.
		copyButton = buttons.addButton(self, label=_("&Copy new text"))
		buttons.addButton(self, wx.ID_CANCEL, label=_("Close without changes"))
		helper.addItem(buttons)
		sizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizerAndFit(sizer)
		replaceButton.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(REPLACE))
		self.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(INSERT_AFTER), id=INSERT_AFTER)
		copyButton.Bind(wx.EVT_BUTTON, lambda evt: self._copy())
		self.proposed.SetFocus()
		self.CentreOnScreen()

	@property
	def value(self) -> str:
		return self.proposed.GetValue().strip()

	def _copy(self) -> None:
		import api
		import ui

		if api.copyToClip(self.value):
			ui.message(_("Copied"))


def textRange(rangeObj: Any) -> Any:
	"""The range without its final paragraph mark, so replacing text keeps the paragraph formatting."""
	rangeObj = rangeObj.Duplicate
	try:
		if str(rangeObj.Text or "").endswith("\r") and rangeObj.End > rangeObj.Start:
			rangeObj.MoveEnd(1, -1)
	except COM_ERRORS:
		pass
	return rangeObj


def replaceText(ctx: WordContext, rangeObj: Any, text: str, after: bool) -> None:
	target = textRange(rangeObj)
	# Paragraph breaks from the dialog become Word paragraph marks.
	text = text.replace("\r\n", "\n").replace("\n", "\r")
	# Translators: Name of the undo step for using AI text.
	with undoRecord(ctx, _("AI text")):
		if after:
			# A paragraph break, then the new text, before the original paragraph mark.
			target.InsertAfter("\r" + text)
		else:
			target.Text = text


def commentsText(ctx: WordContext, limit: int = 500) -> str:
	"""All comments as text, with author, status and replies, for AI."""
	lines = []
	try:
		comments = ctx.document.Comments
		for index in range(1, min(comments.Count, limit) + 1):
			comment = comments.Item(index)
			try:
				if comment.Ancestor is not None:
					continue
			except COM_ERRORS:
				pass
			try:
				done = bool(comment.Done)
			except COM_ERRORS:
				done = False
			# Translators: Status of a comment sent to AI.
			status = _("resolved") if done else _("open")
			lines.append(
				f'[{status}] {comment.Author} on "{cleanText(comment.Scope.Text)[:200]}": {cleanText(comment.Range.Text)}'
			)
			try:
				replies = comment.Replies
				for replyIndex in range(1, replies.Count + 1):
					reply = replies.Item(replyIndex)
					lines.append(f"    reply by {reply.Author}: {cleanText(reply.Range.Text)}")
			except COM_ERRORS:
				pass
	except COM_ERRORS:
		pass
	return "\n".join(lines)

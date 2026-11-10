# Word Document Assistant for NVDA
# Element types that can be navigated and managed in a Word document.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any, Protocol

import addonHandler
import config

from . import wordCom
from .wordCom import COM_ERRORS, WordContext, flatText, formatDate, undoRecord

addonHandler.initTranslation()

CONFIG_SECTION = "wordDocAssistant"


def conf(key: str) -> Any:
	return config.conf[CONFIG_SECTION][key]


class Host(Protocol):
	"""What element types need from the add-on while running an action."""

	def speak(self, message: str) -> None: ...
	def prompt(self, title: str, label: str, value: str = "") -> str | None: ...
	def confirm(self, title: str, message: str) -> bool: ...
	def copy(self, text: str) -> None: ...
	def modified(self) -> None: ...
	def aiAvailable(self) -> bool: ...
	def aiWordSuggestions(self, word: str, context: str, apply: Callable[[str], None]) -> None: ...
	def aiReplyDrafts(
		self,
		author: str,
		comment: str,
		commentedText: str,
		apply: Callable[[str], None],
	) -> None: ...
	def choose(self, title: str, label: str, options: list[str]) -> int | None: ...
	def aiDraft(
		self,
		title: str,
		label: str,
		work: Callable[[], str],
		apply: Callable[[str], None],
	) -> None: ...
	def aiFormat(self) -> None: ...
	def updateWordLists(self) -> None: ...


@dataclass
class Item:
	obj: Any
	navRange: Any
	start: int
	cache: dict = field(default_factory=dict)


@dataclass
class Action:
	label: str
	run: Callable[[], None]
	enabled: bool = True


@dataclass
class Section:
	label: str
	multiline: bool = False


class ElementType:
	"""Base class. Subclasses describe one kind of document element."""

	id = ""
	#: Collecting may take a while in long documents; the user is told to wait.
	slow = False
	#: Virtual key of the letter that moves to this element in navigation mode.
	vk = 0

	@property
	def label(self) -> str:
		raise NotImplementedError

	@property
	def columns(self) -> list[str]:
		raise NotImplementedError

	@property
	def sections(self) -> list[Section]:
		raise NotImplementedError

	def collect(self, ctx: WordContext, limit: int) -> tuple[list[Item], int]:
		"""Return up to limit items sorted by position, and the total count."""
		raise NotImplementedError

	def summary(self, ctx: WordContext, item: Item) -> str:
		raise NotImplementedError

	def row(self, ctx: WordContext, item: Item) -> list[str]:
		raise NotImplementedError

	def sectionTexts(self, ctx: WordContext, item: Item) -> list[str]:
		raise NotImplementedError

	def copyText(self, ctx: WordContext, item: Item) -> str:
		"""The main text of the item, used by the copy command."""
		raise NotImplementedError

	def detailsText(self, ctx: WordContext, item: Item) -> str:
		"""Every section with its label, used by the copy all command."""
		parts = []
		for section, text in zip(self.sections, self.sectionTexts(ctx, item), strict=False):
			if text:
				parts.append(f"{section.label}:\n{text}")
		return "\n\n".join(parts)

	def moveTo(self, ctx: WordContext, item: Item) -> None:
		rangeObj = item.navRange.Duplicate
		rangeObj.Collapse(wordCom.WD_COLLAPSE_START)
		rangeObj.Select()

	def edit(self, ctx: WordContext, item: Item) -> str | None:
		"""Put the caret where the user can edit the item. Returns a message to report."""
		self.moveTo(ctx, item)
		return None

	def actions(self, ctx: WordContext, item: Item | None, anchor: Any, host: Host) -> list[Action]:
		actions = []
		if item is not None:
			actions.append(
				Action(
					# Translators: Action that copies the main text of the current item.
					_("Copy text"),
					lambda: host.copy(self.copyText(ctx, item)),
				),
			)
			actions.append(
				Action(
					# Translators: Action that copies all details of the current item.
					_("Copy all details"),
					lambda: host.copy(self.detailsText(ctx, item)),
				),
			)
		return actions


def _collectRanges(
	collection: Any,
	limit: int,
	budget: wordCom.Budget | None = None,
) -> tuple[list[Item], int]:
	items = []
	for rangeObj in wordCom.iterate(collection, budget):
		if len(items) >= limit:
			break
		items.append(Item(obj=rangeObj, navRange=rangeObj, start=rangeObj.Start))
	return items, len(items)


def _collectProofing(ctx: WordContext, limit: int, collectionName: str) -> tuple[list[Item], int]:
	"""Spelling or grammar errors, read part by part so a long document can stop at the time limit.

	Asking the whole document for its errors makes Word check all of it in one call, which can take
	minutes in a long document and cannot be stopped.
	"""
	budget = wordCom.budgetOf(ctx)
	items: list[Item] = []
	seen: set[int] = set()
	for part in wordCom.chunks(ctx, budget):
		try:
			errors = getattr(part, collectionName)
		except COM_ERRORS:
			continue
		found, _total = _collectRanges(errors, limit - len(items), budget)
		for item in found:
			if item.start not in seen:
				seen.add(item.start)
				items.append(item)
		if len(items) >= limit:
			break
	return items, len(items)


def _context(ctx: WordContext, start: int, end: int) -> str:
	length = conf("contextLength")
	doc = ctx.document
	contentEnd = doc.Content.End
	return flatText(doc.Range(max(0, start - length), min(contentEnd, end + length)).Text)


# Spelling errors ------------------------------------------------------------


class SpellingErrors(ElementType):
	id = "spelling"
	slow = True
	vk = ord("S")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Spelling errors")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the spelling error list.
		return [_("Word"), _("Context")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the spelling error details.
			Section(_("Misspelled word")),
			# Translators: Section label in the spelling error details.
			Section(_("Context"), multiline=True),
			# Translators: Section label in the spelling error details.
			Section(_("Suggestions"), multiline=True),
		]

	def collect(self, ctx, limit):
		return _collectProofing(ctx, limit, "SpellingErrors")

	def _word(self, item: Item) -> str:
		if "word" not in item.cache:
			item.cache["word"] = flatText(item.obj.Text)
		return item.cache["word"]

	def _suggestions(self, item: Item) -> list[str]:
		if "suggestions" not in item.cache:
			names = []
			try:
				suggestions = item.obj.GetSpellingSuggestions()
				for index in range(1, min(suggestions.Count, 10) + 1):
					names.append(suggestions.Item(index).Name)
			except COM_ERRORS:
				pass
			item.cache["suggestions"] = names
		return item.cache["suggestions"]

	def summary(self, ctx, item):
		text = self._word(item)
		if conf("speakSpellingContext"):
			# Translators: Spoken after a misspelled word, followed by the text around it.
			text += ", " + _("in context: {context}").format(
				context=_context(ctx, item.obj.Start, item.obj.End),
			)
		return text

	def row(self, ctx, item):
		return [self._word(item), _context(ctx, item.obj.Start, item.obj.End)]

	def sectionTexts(self, ctx, item):
		return [
			self._word(item),
			_context(ctx, item.obj.Start, item.obj.End),
			# Translators: Shown when Word has no suggestions for a misspelled word.
			", ".join(self._suggestions(item)) or _("No suggestions"),
		]

	def copyText(self, ctx, item):
		return self._word(item)

	def moveTo(self, ctx, item):
		# Select the word so it can be retyped straight away after Enter.
		item.obj.Select()

	def edit(self, ctx, item):
		item.obj.Select()
		# Translators: Reported after Enter on a misspelled word; the word is selected.
		return _("{word} selected, type to replace it").format(word=self._word(item))

	def actions(self, ctx, item, anchor, host):
		actions = []
		if item is not None:
			suggestions = self._suggestions(item)
			for suggestion in suggestions:
				actions.append(
					Action(
						# Translators: Action that replaces a misspelled word with a suggestion.
						_("Replace with {suggestion}").format(suggestion=suggestion),
						lambda s=suggestion: self._replace(ctx, item, s, host),
					),
				)
			if not suggestions:
				# Translators: Disabled menu item shown when Word has no suggestions.
				actions.append(Action(_("No suggestions"), lambda: None, enabled=False))
			actions.append(
				Action(
					# Translators: Action that lets the user type a replacement for a misspelled word.
					_("Type a replacement..."),
					lambda: self._replaceTyped(ctx, item, host),
				),
			)
			if host.aiAvailable():
				actions.append(
					Action(
						# Translators: Action that asks AI for corrections of a misspelled word.
						_("AI suggestions..."),
						lambda: host.aiWordSuggestions(
							self._word(item),
							_context(ctx, item.obj.Start, item.obj.End),
							lambda text: self._replace(ctx, item, text, host),
						),
					),
				)
		return actions + super().actions(ctx, item, anchor, host)

	def _replace(self, ctx, item, text: str, host: Host) -> None:
		rangeObj = item.obj
		# Translators: Name of the undo step for replacing a word.
		with undoRecord(ctx, _("Replace word")):
			rangeObj.Text = text
		rangeObj.Collapse(wordCom.WD_COLLAPSE_END)
		rangeObj.Select()
		host.modified()
		# Translators: Reported after a misspelled word was replaced.
		host.speak(_("Replaced with {text}").format(text=text))

	def _replaceTyped(self, ctx, item, host: Host) -> None:
		text = host.prompt(
			# Translators: Title of the dialog for typing a replacement word.
			_("Replace word"),
			# Translators: Label of the edit field for a replacement word.
			_("Replace {word} with:").format(word=self._word(item)),
			self._word(item),
		)
		if text:
			self._replace(ctx, item, text, host)


# Grammar issues --------------------------------------------------------------


class GrammarErrors(ElementType):
	id = "grammar"
	slow = True
	vk = ord("G")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Grammar issues")

	@property
	def columns(self) -> list[str]:
		# Translators: Column header in the grammar issue list.
		return [_("Text")]

	@property
	def sections(self) -> list[Section]:
		# Translators: Section label in the grammar issue details.
		return [Section(_("Text with the issue"), multiline=True)]

	def collect(self, ctx, limit):
		return _collectProofing(ctx, limit, "GrammaticalErrors")

	def summary(self, ctx, item):
		return flatText(item.obj.Text, 200)

	def row(self, ctx, item):
		return [flatText(item.obj.Text, 150)]

	def sectionTexts(self, ctx, item):
		return [wordCom.cleanText(item.obj.Text)]

	def copyText(self, ctx, item):
		return wordCom.cleanText(item.obj.Text)


# Comments -----------------------------------------------------------------------


def _openCommentForEditing(comment: Any) -> None:
	try:
		comment.Edit()
	except COM_ERRORS:
		pass


class Comments(ElementType):
	id = "comments"
	vk = ord("C")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Comments")

	@property
	def columns(self) -> list[str]:
		return [
			# Translators: Column header in the comment list.
			_("Author"),
			# Translators: Column header in the comment list.
			_("Comment"),
			# Translators: Column header in the comment list.
			_("Commented text"),
			# Translators: Column header in the comment list.
			_("Status"),
			# Translators: Column header in the comment list.
			_("Replies"),
		]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the comment details.
			Section(_("Author")),
			# Translators: Section label in the comment details.
			Section(_("Date")),
			# Translators: Section label in the comment details.
			Section(_("Status")),
			# Translators: Section label in the comment details.
			Section(_("Commented text"), multiline=True),
			# Translators: Section label in the comment details.
			Section(_("Comment"), multiline=True),
			# Translators: Section label in the comment details.
			Section(_("Replies"), multiline=True),
		]

	def collect(self, ctx, limit):
		items = []
		truncated = False
		for comment in wordCom.iterate(ctx.document.Comments, wordCom.budgetOf(ctx)):
			try:
				# Replies are shown inside their parent comment.
				if comment.Ancestor is not None:
					continue
			except COM_ERRORS:
				pass
			if len(items) >= limit:
				truncated = True
				break
			scope = comment.Scope
			items.append(Item(obj=comment, navRange=scope, start=scope.Start))
		items.sort(key=lambda i: i.start)
		# The exact number of top level comments is unknown when the list was cut short.
		return items, len(items) + 1 if truncated else len(items)

	@staticmethod
	def _isDone(comment: Any) -> bool:
		try:
			return bool(comment.Done)
		except COM_ERRORS:
			return False

	def _status(self, comment: Any) -> str:
		if self._isDone(comment):
			# Translators: Status of a resolved comment.
			return _("Resolved")
		# Translators: Status of an open comment.
		return _("Open")

	@staticmethod
	def _replies(comment: Any) -> list[Any]:
		try:
			replies = comment.Replies
			return [replies.Item(i) for i in range(1, replies.Count + 1)]
		except COM_ERRORS:
			return []

	def summary(self, ctx, item):
		comment = item.obj
		text = flatText(comment.Range.Text)
		parts = []
		if conf("reportAuthor"):
			parts.append(f"{comment.Author}: {text}")
		else:
			parts.append(text)
		if conf("reportDate"):
			parts.append(formatDate(comment.Date))
		if self._isDone(comment):
			parts.append(self._status(comment))
		replyCount = len(self._replies(comment))
		if replyCount:
			# Translators: Spoken after a comment that has replies.
			parts.append(_("{count} replies").format(count=replyCount))
		return ", ".join(parts)

	def row(self, ctx, item):
		comment = item.obj
		return [
			str(comment.Author),
			flatText(comment.Range.Text, 150),
			flatText(comment.Scope.Text, 80),
			self._status(comment),
			str(len(self._replies(comment))),
		]

	def sectionTexts(self, ctx, item):
		comment = item.obj
		replies = []
		for number, reply in enumerate(self._replies(comment), start=1):
			replies.append(
				f"{number}. {reply.Author} ({formatDate(reply.Date)}):\n{wordCom.cleanText(reply.Range.Text)}",
			)
		initial = ""
		try:
			if comment.Initial:
				initial = f" ({comment.Initial})"
		except COM_ERRORS:
			pass
		return [
			f"{comment.Author}{initial}",
			formatDate(comment.Date),
			self._status(comment),
			# Translators: Shown when a comment is not attached to any text.
			wordCom.cleanText(comment.Scope.Text) or _("(no text selected)"),
			wordCom.cleanText(comment.Range.Text),
			# Translators: Shown when a comment has no replies.
			"\n\n".join(replies) or _("No replies"),
		]

	def copyText(self, ctx, item):
		return wordCom.cleanText(item.obj.Range.Text)

	def actions(self, ctx, item, anchor, host):
		actions = []
		if item is not None:
			comment = item.obj
			actions += [
				# Translators: Action that replies to a comment.
				Action(_("Reply..."), lambda: self._reply(ctx, comment, host)),
			]
			if host.aiAvailable():
				actions.append(
					Action(
						# Translators: Action that asks AI to draft replies to a comment.
						_("Draft a reply with AI..."),
						lambda: host.aiReplyDrafts(
							str(comment.Author),
							wordCom.cleanText(comment.Range.Text),
							wordCom.cleanText(comment.Scope.Text),
							lambda text: self._postReply(ctx, comment, text, host),
						),
					),
				)
			actions += [
				# Translators: Action that edits the text of a comment.
				Action(_("Edit comment..."), lambda: self._editText(ctx, comment, host)),
				Action(
					# Translators: Action that reopens a resolved comment.
					_("Reopen")
					if self._isDone(comment)
					# Translators: Action that marks a comment as resolved.
					else _("Resolve"),
					lambda: self._toggleDone(ctx, comment, host),
				),
				# Translators: Action that deletes a comment.
				Action(_("Delete comment"), lambda: self._delete(ctx, comment, host)),
				Action(
					# Translators: Action that copies the text the comment is attached to.
					_("Copy commented text"),
					lambda: host.copy(wordCom.cleanText(comment.Scope.Text)),
				),
			]
		actions.append(
			Action(
				# Translators: Action that adds a new comment where the caret was before navigation.
				_("New comment at the starting position..."),
				lambda: self._add(ctx, anchor, host),
				enabled=anchor is not None,
			),
		)
		return actions + super().actions(ctx, item, anchor, host)

	def _reply(self, ctx, comment, host: Host) -> None:
		# Translators: Title and label of the dialog for replying to a comment.
		text = host.prompt(_("Reply to comment"), _("Reply to {author}:").format(author=comment.Author))
		if text:
			self._postReply(ctx, comment, text, host)

	def _postReply(self, ctx, comment, text: str, host: Host) -> None:
		# Translators: Name of the undo step for replying to a comment.
		try:
			with undoRecord(ctx, _("Reply to comment")):
				comment.Replies.Add(comment.Scope, text)
		except COM_ERRORS:
			# Some Word versions do not allow replies through the object model.
			host.copy(text)
			_openCommentForEditing(comment)
			host.speak(
				# Translators: Reported when Word does not let the add-on post a reply.
				_(
					"Word did not accept the reply. It was copied; paste it as a reply in the comment Word opened",
				),
			)
			return
		host.modified()
		# Translators: Reported after a reply was added.
		host.speak(_("Reply added"))

	def _editText(self, ctx, comment, host: Host) -> None:
		current = wordCom.cleanText(comment.Range.Text)
		# Translators: Title and label of the dialog for editing a comment.
		text = host.prompt(_("Edit comment"), _("Comment text:"), current)
		if text is None or text == current:
			return
		# Translators: Name of the undo step for editing a comment.
		try:
			with undoRecord(ctx, _("Edit comment")):
				comment.Range.Text = text
		except COM_ERRORS:
			# Modern comments in Microsoft 365 can refuse changes through the object model.
			host.copy(text)
			_openCommentForEditing(comment)
			host.speak(
				# Translators: Reported when Word does not let the add-on change a comment.
				_(
					"Word did not accept the change. The new text was copied; paste it into the comment Word opened",
				),
			)
			return
		host.modified()
		# Translators: Reported after a comment was changed.
		host.speak(_("Comment updated"))

	def _toggleDone(self, ctx, comment, host: Host) -> None:
		done = not self._isDone(comment)
		comment.Done = done
		host.modified()
		# Translators: Reported after a comment was resolved or reopened.
		host.speak(_("Comment resolved") if done else _("Comment reopened"))

	def _delete(self, ctx, comment, host: Host) -> None:
		# Translators: Confirmation before deleting a comment.
		if not host.confirm(_("Delete comment"), _("Delete this comment and its replies?")):
			return
		# Translators: Name of the undo step for deleting a comment.
		with undoRecord(ctx, _("Delete comment")):
			comment.Delete()
		host.modified()
		# Translators: Reported after a comment was deleted.
		host.speak(_("Comment deleted"))

	def _add(self, ctx, anchor, host: Host) -> None:
		if anchor is None or not wordCom.isMainStory(anchor):
			# Translators: Reported when a comment cannot be added at the starting position.
			host.speak(_("A comment can only be added in the main text"))
			return
		# Translators: Title and label of the dialog for a new comment.
		text = host.prompt(_("New comment"), _("Comment text:"))
		if not text:
			return
		# Translators: Name of the undo step for adding a comment.
		with undoRecord(ctx, _("New comment")):
			ctx.document.Comments.Add(anchor, text)
		host.modified()
		# Translators: Reported after a comment was added.
		host.speak(_("Comment added"))


# Footnotes and endnotes -------------------------------------------------------


class _Notes(ElementType):
	"""Shared behaviour of footnotes and endnotes."""

	def _collection(self, ctxOrRange: Any) -> Any:
		raise NotImplementedError

	@property
	def noteName(self) -> str:
		raise NotImplementedError

	@property
	def columns(self) -> list[str]:
		return [
			# Translators: Column header in the footnote and endnote lists.
			_("Number"),
			# Translators: Column header in the footnote and endnote lists.
			_("Note text"),
			# Translators: Column header in the footnote and endnote lists.
			_("Text before the reference"),
		]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the footnote and endnote details.
			Section(_("Number")),
			# Translators: Section label in the footnote and endnote details.
			Section(_("Note text"), multiline=True),
			# Translators: Section label in the footnote and endnote details.
			Section(_("Text before the reference"), multiline=True),
		]

	def collect(self, ctx, limit):
		notes = self._collection(ctx.document)
		total = notes.Count
		items = []
		for note in wordCom.iterate(notes, wordCom.budgetOf(ctx)):
			if len(items) >= limit:
				break
			reference = note.Reference
			items.append(Item(obj=note, navRange=reference, start=reference.Start))
		return items, total

	def _before(self, ctx, note) -> str:
		start = note.Reference.Start
		return flatText(ctx.document.Range(max(0, start - 60), start).Text)

	def summary(self, ctx, item):
		return f"{self.noteName} {item.obj.Index}: {flatText(item.obj.Range.Text)}"

	def row(self, ctx, item):
		return [str(item.obj.Index), flatText(item.obj.Range.Text, 150), self._before(ctx, item.obj)]

	def sectionTexts(self, ctx, item):
		note = item.obj
		return [str(note.Index), wordCom.cleanText(note.Range.Text), self._before(ctx, note)]

	def copyText(self, ctx, item):
		return wordCom.cleanText(item.obj.Range.Text)

	def moveTo(self, ctx, item):
		rangeObj = item.obj.Reference.Duplicate
		rangeObj.Collapse(wordCom.WD_COLLAPSE_END)
		rangeObj.Select()

	def edit(self, ctx, item):
		rangeObj = item.obj.Range.Duplicate
		rangeObj.Collapse(wordCom.WD_COLLAPSE_END)
		rangeObj.Select()
		# Translators: Reported after Enter moves the caret into a footnote or endnote.
		return _("In {note} {number}, ready to edit").format(note=self.noteName, number=item.obj.Index)

	def actions(self, ctx, item, anchor, host):
		actions = []
		if item is not None:
			note = item.obj
			actions += [
				# Translators: Action that edits the text of a footnote or endnote.
				Action(_("Edit note text..."), lambda: self._editText(ctx, note, host)),
				# Translators: Action that deletes a footnote or endnote.
				Action(_("Delete note"), lambda: self._delete(ctx, note, host)),
				Action(self.convertLabel, lambda: self._convert(ctx, note, host)),
			]
		actions.append(
			Action(
				self.insertLabel,
				lambda: self._insert(ctx, anchor, host),
				enabled=anchor is not None,
			),
		)
		return actions + super().actions(ctx, item, anchor, host)

	@property
	def convertLabel(self) -> str:
		raise NotImplementedError

	@property
	def insertLabel(self) -> str:
		raise NotImplementedError

	def _editText(self, ctx, note, host: Host) -> None:
		current = wordCom.cleanText(note.Range.Text)
		# Translators: Title and label of the dialog for editing a footnote or endnote.
		text = host.prompt(_("Edit note"), _("Note text:"), current)
		if text is None or text == current:
			return
		# Translators: Name of the undo step for editing a note.
		with undoRecord(ctx, _("Edit note")):
			note.Range.Text = text
		host.modified()
		# Translators: Reported after a footnote or endnote was changed.
		host.speak(_("Note updated"))

	def _delete(self, ctx, note, host: Host) -> None:
		# Translators: Confirmation before deleting a footnote or endnote.
		if not host.confirm(
			_("Delete note"),
			_("Delete {note} {number}?").format(
				note=self.noteName,
				number=note.Index,
			),
		):
			return
		# Translators: Name of the undo step for deleting a note.
		with undoRecord(ctx, _("Delete note")):
			note.Delete()
		host.modified()
		# Translators: Reported after a footnote or endnote was deleted.
		host.speak(_("Note deleted"))

	def _convert(self, ctx, note, host: Host) -> None:
		# Translators: Name of the undo step for converting a note.
		with undoRecord(ctx, _("Convert note")):
			self._collection(note.Reference).Convert()
		host.modified()
		# Translators: Reported after a footnote or endnote was converted.
		host.speak(_("Converted"))

	def _insert(self, ctx, anchor, host: Host) -> None:
		if anchor is None or not wordCom.isMainStory(anchor):
			# Translators: Reported when a note cannot be inserted at the starting position.
			host.speak(_("A note can only be inserted in the main text"))
			return
		text = host.prompt(self.insertLabel.rstrip("."), _("Note text:"))
		if text is None:
			return
		position = anchor.Duplicate
		position.Collapse(wordCom.WD_COLLAPSE_END)
		with undoRecord(ctx, self.insertLabel.rstrip(".")):
			note = self._collection(ctx.document).Add(position)
			if text:
				note.Range.Text = text
		host.modified()
		# Translators: Reported after a footnote or endnote was inserted.
		host.speak(_("{note} {number} inserted").format(note=self.noteName, number=note.Index))


class Footnotes(_Notes):
	id = "footnotes"
	vk = ord("F")

	def _collection(self, ctxOrRange):
		return ctxOrRange.Footnotes

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Footnotes")

	@property
	def noteName(self) -> str:
		# Translators: Name of a footnote when it is reported.
		return _("Footnote")

	@property
	def convertLabel(self) -> str:
		# Translators: Action that converts a footnote to an endnote.
		return _("Convert to endnote")

	@property
	def insertLabel(self) -> str:
		# Translators: Action that inserts a new footnote.
		return _("Insert footnote at the starting position...")


class Endnotes(_Notes):
	id = "endnotes"
	vk = ord("E")

	def _collection(self, ctxOrRange):
		return ctxOrRange.Endnotes

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Endnotes")

	@property
	def noteName(self) -> str:
		# Translators: Name of an endnote when it is reported.
		return _("Endnote")

	@property
	def convertLabel(self) -> str:
		# Translators: Action that converts an endnote to a footnote.
		return _("Convert to footnote")

	@property
	def insertLabel(self) -> str:
		# Translators: Action that inserts a new endnote.
		return _("Insert endnote at the starting position...")


# Tracked changes ----------------------------------------------------------------


def _revisionTypeName(revisionType: int) -> str:
	names = {
		# Translators: Type of a tracked change.
		1: _("Insertion"),
		# Translators: Type of a tracked change.
		2: _("Deletion"),
		# Translators: Type of a tracked change.
		3: _("Formatting"),
		# Translators: Type of a tracked change.
		4: _("Paragraph numbering"),
		# Translators: Type of a tracked change.
		5: _("Field display"),
		# Translators: Type of a tracked change.
		6: _("Reconciled"),
		# Translators: Type of a tracked change.
		7: _("Conflict"),
		# Translators: Type of a tracked change.
		8: _("Style"),
		# Translators: Type of a tracked change.
		9: _("Replacement"),
		# Translators: Type of a tracked change.
		10: _("Paragraph formatting"),
		# Translators: Type of a tracked change.
		11: _("Table formatting"),
		# Translators: Type of a tracked change.
		12: _("Section formatting"),
		# Translators: Type of a tracked change.
		13: _("Style definition"),
		# Translators: Type of a tracked change.
		14: _("Moved from"),
		# Translators: Type of a tracked change.
		15: _("Moved to"),
		# Translators: Type of a tracked change.
		16: _("Table cell inserted"),
		# Translators: Type of a tracked change.
		17: _("Table cell deleted"),
		# Translators: Type of a tracked change.
		18: _("Table cells merged"),
		# Translators: Type of a tracked change.
		19: _("Table cell split"),
	}
	# Translators: Type of a tracked change that the add-on does not know.
	return names.get(revisionType, _("Change"))


class Revisions(ElementType):
	id = "revisions"
	vk = ord("R")

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Revisions")

	@property
	def columns(self) -> list[str]:
		return [
			# Translators: Column header in the revision list.
			_("Type"),
			# Translators: Column header in the revision list.
			_("Text"),
			# Translators: Column header in the revision list.
			_("Author"),
			# Translators: Column header in the revision list.
			_("Date"),
		]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the revision details.
			Section(_("Type")),
			# Translators: Section label in the revision details.
			Section(_("Author")),
			# Translators: Section label in the revision details.
			Section(_("Date")),
			# Translators: Section label in the revision details.
			Section(_("Changed text"), multiline=True),
			# Translators: Section label in the revision details.
			Section(_("Formatting change"), multiline=True),
		]

	def collect(self, ctx, limit):
		revisions = ctx.document.Revisions
		total = revisions.Count
		items = []
		for revision in wordCom.iterate(revisions, wordCom.budgetOf(ctx)):
			if len(items) >= limit:
				break
			rangeObj = revision.Range
			items.append(Item(obj=revision, navRange=rangeObj, start=rangeObj.Start))
		items.sort(key=lambda i: i.start)
		return items, total

	@staticmethod
	def _formatDescription(revision) -> str:
		try:
			return wordCom.cleanText(revision.FormatDescription)
		except COM_ERRORS:
			return ""

	def _text(self, revision, maxLength: int = 0) -> str:
		text = flatText(revision.Range.Text, maxLength)
		return text or self._formatDescription(revision)

	def summary(self, ctx, item):
		revision = item.obj
		text = f"{_revisionTypeName(revision.Type)}: {self._text(revision, 200)}"
		if conf("reportAuthor"):
			text += f", {revision.Author}"
		if conf("reportDate"):
			text += f", {formatDate(revision.Date)}"
		return text

	def row(self, ctx, item):
		revision = item.obj
		return [
			_revisionTypeName(revision.Type),
			self._text(revision, 150),
			str(revision.Author),
			formatDate(revision.Date),
		]

	def sectionTexts(self, ctx, item):
		revision = item.obj
		return [
			_revisionTypeName(revision.Type),
			str(revision.Author),
			formatDate(revision.Date),
			wordCom.cleanText(revision.Range.Text),
			self._formatDescription(revision),
		]

	def copyText(self, ctx, item):
		return wordCom.cleanText(item.obj.Range.Text) or self._formatDescription(item.obj)

	def actions(self, ctx, item, anchor, host):
		actions = []
		if item is not None:
			revision = item.obj
			author = str(revision.Author)
			actions += [
				# Translators: Action that accepts one tracked change.
				Action(_("Accept"), lambda: self._single(ctx, revision, True, host)),
				# Translators: Action that rejects one tracked change.
				Action(_("Reject"), lambda: self._single(ctx, revision, False, host)),
				Action(
					# Translators: Action that accepts every tracked change by one author.
					_("Accept all by {author}").format(author=author),
					lambda: self._byAuthor(ctx, author, True, host),
				),
				Action(
					# Translators: Action that rejects every tracked change by one author.
					_("Reject all by {author}").format(author=author),
					lambda: self._byAuthor(ctx, author, False, host),
				),
			]
		actions += [
			# Translators: Action that accepts every tracked change in the document.
			Action(_("Accept all in document..."), lambda: self._all(ctx, True, host)),
			# Translators: Action that rejects every tracked change in the document.
			Action(_("Reject all in document..."), lambda: self._all(ctx, False, host)),
		]
		return actions + super().actions(ctx, item, anchor, host)

	def _single(self, ctx, revision, accept: bool, host: Host) -> None:
		if accept:
			revision.Accept()
		else:
			revision.Reject()
		host.modified()
		# Translators: Reported after a tracked change was accepted or rejected.
		host.speak(_("Accepted") if accept else _("Rejected"))

	def _byAuthor(self, ctx, author: str, accept: bool, host: Host) -> None:
		revisions = ctx.document.Revisions
		count = 0
		# Translators: Name of the undo step for accepting or rejecting changes by one author.
		with undoRecord(ctx, _("Review changes by {author}").format(author=author)):
			for index in range(revisions.Count, 0, -1):
				revision = revisions.Item(index)
				if str(revision.Author) == author:
					if accept:
						revision.Accept()
					else:
						revision.Reject()
					count += 1
		host.modified()
		if accept:
			# Translators: Reported after accepting changes; {count} is a number.
			host.speak(_("{count} changes accepted").format(count=count))
		else:
			# Translators: Reported after rejecting changes; {count} is a number.
			host.speak(_("{count} changes rejected").format(count=count))

	def _all(self, ctx, accept: bool, host: Host) -> None:
		revisions = ctx.document.Revisions
		count = revisions.Count
		if accept:
			# Translators: Confirmation before accepting every change; {count} is a number.
			question = _("Accept all {count} changes in the document?").format(count=count)
			title = _("Accept all in document...").rstrip(".")
		else:
			# Translators: Confirmation before rejecting every change; {count} is a number.
			question = _("Reject all {count} changes in the document?").format(count=count)
			title = _("Reject all in document...").rstrip(".")
		if not host.confirm(title, question):
			return
		if accept:
			revisions.AcceptAll()
			host.modified()
			host.speak(_("{count} changes accepted").format(count=count))
		else:
			revisions.RejectAll()
			host.modified()
			host.speak(_("{count} changes rejected").format(count=count))


# The list of element types is in registry.py.

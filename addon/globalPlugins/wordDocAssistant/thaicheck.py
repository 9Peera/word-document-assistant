# Word Document Assistant for NVDA
# Thai word check in navigation mode (Z), using the offline word lists.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import os
from typing import Any

import addonHandler
import config
import globalVars

from . import thaidict, wordCom
from .elements import CONFIG_SECTION, Action, ElementType, Item, Section, conf
from .wordCom import COM_ERRORS, WordContext, flatText, undoRecord

addonHandler.initTranslation()

ORST_DICTIONARY = "https://dictionary.orst.go.th/"

_lists: thaidict.WordLists | None = None


def listsFolder() -> str:
	return os.path.join(globalVars.appArgs.configPath, "wordDocAssistant", "dictionaries")


def wordLists() -> thaidict.WordLists:
	"""The shared word lists, loaded on first use."""
	global _lists
	if _lists is None:
		_lists = thaidict.WordLists(listsFolder(), config.conf[CONFIG_SECTION])
		_lists.load()
	return _lists


def reload() -> None:
	if _lists is not None:
		_lists.load()


#: How far around the expected position a finding is looked for when field codes shift positions.
LOCATE_WINDOW = 400


def _locate(ctx: WordContext, partRange: Any, offset: int, word: str, text: str) -> Any:
	"""The Word range of a finding. Field codes can shift positions, so check and search if needed."""
	doc = ctx.document
	expected = int(partRange.Start) + offset
	rangeObj = doc.Range(expected, expected + len(word))
	try:
		if str(rangeObj.Text or "") == word:
			return rangeObj
		# Search near the expected position and take the closest match.
		window = doc.Range(
			max(int(partRange.Start), expected - LOCATE_WINDOW),
			min(int(partRange.End), expected + len(word) + LOCATE_WINDOW),
		)
	except COM_ERRORS:
		return None
	best = None
	for number, match in enumerate(wordCom.findAll(window, word)):
		if best is None or abs(match.Start - expected) < abs(best.Start - expected):
			best = match
		if number >= 20:
			break
	return best


class ThaiWords(ElementType):
	id = "thaiWords"
	vk = ord("Z")
	slow = True

	@property
	def label(self) -> str:
		# Translators: Element type name.
		return _("Thai word check")

	@property
	def columns(self) -> list[str]:
		# Translators: Column headers in the Thai word check list.
		return [_("Word"), _("Problem"), _("Suggestions"), _("Context")]

	@property
	def sections(self) -> list[Section]:
		return [
			# Translators: Section label in the Thai word check details.
			Section(_("Word")),
			# Translators: Section label in the Thai word check details.
			Section(_("Problem")),
			# Translators: Section label in the Thai word check details.
			Section(_("Suggestions"), multiline=True),
			# Translators: Section label in the Thai word check details.
			Section(_("Note")),
			# Translators: Section label in the Thai word check details.
			Section(_("Context"), multiline=True),
		]

	def collect(self, ctx, limit):
		dictionary = wordLists().dictionary
		items: list[Item] = []
		total = 0
		checkUnknown = bool(conf("thaiCheckUnknown"))
		budget = wordCom.budgetOf(ctx)
		# The text is read in large parts: reading it paragraph by paragraph takes minutes in a long document.
		for part in wordCom.chunks(ctx, budget):
			try:
				text = wordCom.sameLengthText(part.Text)
			except COM_ERRORS:
				continue
			for finding in dictionary.check(text, unknown=checkUnknown):
				total += 1
				if len(items) >= limit:
					continue
				if budget.over():
					break
				rangeObj = _locate(ctx, part, finding.start, finding.word, text)
				if rangeObj is None:
					continue
				context = flatText(text[max(0, finding.start - 25) : finding.end + 25])
				items.append(
					Item(obj=finding, navRange=rangeObj, start=rangeObj.Start, cache={"context": context}),
				)
		return items, total

	@staticmethod
	def _problem(finding: thaidict.Finding) -> str:
		if finding.kind == "misspelled":
			# Translators: Problem found by the Thai word check.
			return _("Commonly misspelled")
		# Translators: Problem found by the Thai word check.
		return _("Not in the word list")

	@staticmethod
	def _suggestions(finding: thaidict.Finding) -> str:
		# Translators: Shown when there are no suggestions.
		return ", ".join(finding.suggestions) or _("No suggestions")

	def summary(self, ctx, item):
		finding = item.obj
		if finding.kind == "misspelled":
			# Translators: Spoken for a commonly misspelled word, for example "อนุญาติ, should be อนุญาต".
			text = _("{word}, should be {right}").format(word=finding.word, right=finding.suggestions[0])
		else:
			# Translators: Spoken for a word not in the word list, followed by suggestions.
			text = _("{word}, not in the word list, {suggestions}").format(
				word=finding.word,
				suggestions=self._suggestions(finding),
			)
		if finding.note:
			text += f" ({finding.note})"
		return text

	def row(self, ctx, item):
		finding = item.obj
		return [finding.word, self._problem(finding), self._suggestions(finding), item.cache["context"]]

	def sectionTexts(self, ctx, item):
		finding = item.obj
		return [
			finding.word,
			self._problem(finding),
			self._suggestions(finding),
			finding.note,
			item.cache["context"],
		]

	def copyText(self, ctx, item):
		return item.obj.word

	def moveTo(self, ctx, item):
		item.navRange.Select()

	def edit(self, ctx, item):
		item.navRange.Select()
		return _("{word} selected, type to replace it").format(word=item.obj.word)

	def actions(self, ctx, item, anchor, host):
		lists = wordLists()
		actions: list[Action] = []
		if item is not None:
			finding = item.obj

			def replace(text: str) -> None:
				# Translators: Name of the undo step for replacing a word.
				with undoRecord(ctx, _("Replace word")):
					item.navRange.Text = text
				host.modified()
				host.speak(_("Replaced with {text}").format(text=text))

			def replaceAll(text: str) -> None:
				items, _total = self.collect(ctx, 100000)
				same = sorted(
					(i for i in items if i.obj.word == finding.word),
					key=lambda i: i.start,
					reverse=True,
				)
				with undoRecord(ctx, _("Replace word")):
					for other in same:
						other.navRange.Text = text
				host.modified()
				# Translators: Reported after every occurrence of a word was replaced; {count} is a number.
				host.speak(_("{count} replaced with {text}").format(count=len(same), text=text))

			def typed() -> None:
				text = host.prompt(
					_("Replace word"),
					_("Replace {word} with:").format(word=finding.word),
					finding.word,
				)
				if text:
					replace(text)

			def addWord() -> None:
				lists.addMyWord(finding.word)
				host.modified()
				# Translators: Reported after a word was added to the user's word list.
				host.speak(_("{word} added to my words").format(word=finding.word))

			def ignore() -> None:
				lists.dictionary.ignored.add(finding.word)
				host.modified()
				# Translators: Reported after a word is ignored until NVDA restarts.
				host.speak(_("{word} ignored until NVDA restarts").format(word=finding.word))

			def lookUp() -> None:
				host.copy(finding.word)
				try:
					os.startfile(ORST_DICTIONARY)
				except (OSError, AttributeError):
					return
				# Translators: Reported after the Royal Institute dictionary website was opened.
				host.speak(
					_(
						"Royal Institute dictionary opened. The word is on the clipboard, paste it in the search box",
					),
				)

			for suggestion in finding.suggestions:
				actions.append(
					Action(
						_("Replace with {suggestion}").format(suggestion=suggestion),
						lambda s=suggestion: replace(s),
					),
				)
			if finding.suggestions:
				right = finding.suggestions[0]
				actions.append(
					Action(
						# Translators: Action that replaces every occurrence of a word.
						_("Replace every {word} with {right}").format(word=finding.word, right=right),
						lambda: replaceAll(right),
					),
				)
			# Translators: Action that lets the user type a replacement for a misspelled word.
			actions.append(Action(_("Type a replacement..."), typed))
			if finding.kind == "unknown":
				# Translators: Action that adds a word to the user's own word list.
				actions.append(Action(_("Add {word} to my words").format(word=finding.word), addWord))
			# Translators: Action that ignores a word until NVDA restarts.
			actions.append(Action(_("Ignore {word}").format(word=finding.word), ignore))
			# Translators: Action that opens the Royal Institute dictionary website.
			actions.append(Action(_("Look up in the Royal Institute dictionary website"), lookUp))
		# Translators: Action that checks for and downloads newer word lists.
		actions.append(Action(_("Update the Thai word lists now..."), host.updateWordLists))
		return actions + super().actions(ctx, item, anchor, host)

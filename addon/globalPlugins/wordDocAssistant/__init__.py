# Word Document Assistant for NVDA
# Document navigation mode for Microsoft Word: spelling, grammar, comments,
# footnotes, endnotes and tracked changes.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

from typing import Any

import addonHandler
import api
import config
import core
import globalPluginHandler
import globalVars
import gui
import speech
import ui
import wx
from keyboardHandler import KeyboardInputGesture
from logHandler import log
from scriptHandler import script

from . import (
	ai,
	dialogs,
	elements,
	formatDialog,
	keySetup,
	presets,
	registry,
	settings,
	sounds,
	tools,
	wordCom,
	wordListUpdates,
)
from .elements import Action, ElementType, Item, conf
from .wordCom import COM_ERRORS, WordContext, WordUnavailable

addonHandler.initTranslation()

# Virtual key codes. Navigation mode reads keys by virtual key, so it works the same
# with the Thai (Kedmanee) and English keyboard layouts.
VK_BACK = 0x08
VK_RETURN = 0x0D
VK_ESCAPE = 0x1B
VK_SPACE = 0x20
VK_END = 0x23
VK_HOME = 0x24
VK_LEFT = 0x25
VK_UP = 0x26
VK_RIGHT = 0x27
VK_DOWN = 0x28
VK_APPS = 0x5D
VK_F8 = 0x77
VK_F10 = 0x79
VK_C = ord("C")

# Keys that only count as navigation keys when they are the real arrow block keys,
# so that the numeric keypad review commands keep working.
EXTENDED_ONLY = {
	VK_LEFT: ("kind", -1),
	VK_RIGHT: ("kind", 1),
	VK_UP: ("move", -1),
	VK_DOWN: ("move", 1),
	VK_HOME: ("edge", -1),
	VK_END: ("edge", 1),
}
PLAIN_KEYS = {
	VK_RETURN: ("edit",),
	VK_ESCAPE: ("escape",),
	VK_SPACE: ("list",),
	VK_BACK: ("back",),
	VK_APPS: ("actions",),
}

# Letters that run a command instead of moving to an element type.
CMD_TRACK = ord("T")
CMD_NOTE = ord("N")
CMD_POSITION = ord("P")
CMD_SUMMARY = ord("Q")
CMD_HELP = ord("H")
CMD_AI = ord("W")
CMD_FORMAT = ord("J")
CMD_AI_TOOLS = ord("V")
CMD_THAI_TOOLS = ord("Y")


def _helpText() -> str:
	# Translators: Help shown with H in navigation mode. Keep the key names in the translation.
	return _("""Word Document Assistant: navigation mode keys

Enter and leave the mode: NVDA+Shift+F8

Move to an element type (Shift with the letter moves to the previous one):
S  spelling errors
G  grammar issues
C  comments
F  footnotes
E  endnotes
R  revisions (tracked changes)
O  headings
U  tables
M  pictures and their alternative text
K  links
X  cross-references
I  citations
B  bookmarks
D  fields
L  text marked with the wrong language
A  accessibility problems
Z  Thai word check with the offline word lists (commonly misspelled words and words not in the list)

Inside the mode:
Down arrow / Up arrow  next or previous item of the current type
Left arrow / Right arrow  change the element type
Home / End  first or last item
Space  list of all items of the current type, with separate details and copy buttons
Applications or Shift+F10  actions for the current item
Control+C  copy the text of the current item
Control+Shift+C  copy all details of the current item

Commands:
T  report Track Changes status, Shift+T turns Track Changes on or off
N  read the footnote or endnote at the caret
P  report page, section, line and heading
Q  document summary
W  AI check of the current paragraph, Shift+W AI check of the text selected before the mode started
J  format the document with AI (official letter, memo, letter, report, article and more)
V  AI tools: summarize, rewrite, translate, summarize comments, check consistency
Y  Thai tools: Thai and Arabic digits, Buddhist and Christian era years
H  this help

Anywhere: NVDA+Control+Shift+F8 sets up or updates the AI service and API key

Leaving the mode:
Enter  leave the mode at the current item, ready to edit
Escape  leave the mode, the caret stays at the current item
Backspace  go back to where the caret was when the mode started, and leave the mode
Any other key leaves the mode and does its normal job in Word.""")


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	# Translators: Category of the add-on commands in the Input Gestures dialog.
	scriptCategory = _("Word Document Assistant")

	def __init__(self):
		super().__init__()
		settings.initConfig()
		ai.protectStoredKeys()
		self._wordListsMenuItem = None
		if not globalVars.appArgs.secure:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(settings.WordDocAssistantPanel)
			self._addToolsMenuItem()
			# Look for newer Thai word lists a minute after NVDA starts, at most once a week, if the user allowed it.
			core.callLater(60000, wordListUpdates.scheduledCheck)
		self.active = False
		self.ctx: WordContext | None = None
		#: Where the caret was when the mode started. Used for Backspace and to insert new items.
		self.anchor: Any = None
		self._kindIndex = 0
		self._items: dict[str, list[Item]] = {}
		self._totals: dict[str, int] = {}
		self._index: int | None = None
		self._indexKind = ""
		self._inclusive = True

	def _addToolsMenuItem(self) -> None:
		try:
			toolsMenu = gui.mainFrame.sysTrayIcon.toolsMenu
			# Translators: Item in the NVDA Tools menu.
			self._wordListsMenuItem = toolsMenu.Append(
				wx.ID_ANY, _("Update Thai word lists (Word Document Assistant)...")
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, lambda evt: wordListUpdates.updateNow(), self._wordListsMenuItem
			)
		except AttributeError:
			self._wordListsMenuItem = None

	def terminate(self):
		try:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(settings.WordDocAssistantPanel)
		except ValueError:
			pass
		if self._wordListsMenuItem is not None:
			try:
				gui.mainFrame.sysTrayIcon.toolsMenu.Remove(self._wordListsMenuItem)
			except (AttributeError, RuntimeError):
				pass
		super().terminate()

	def updateWordLists(self) -> None:
		wx.CallAfter(wordListUpdates.updateNow)

	@property
	def kind(self) -> ElementType:
		return registry.NAVIGATION_TYPES[self._kindIndex]

	# Gesture handling -------------------------------------------------------

	def getScript(self, gesture):
		if not self.active:
			found = super().getScript(gesture)
			# Outside Word, leave NVDA+Shift+F8 free for anything else that uses it.
			if (
				getattr(found, "__func__", None) is GlobalPlugin.script_toggleNavigationMode
				and not wordCom.isWordFocused()
			):
				return None
			return found
		# Keys pressed in NVDA dialogs, menus or other windows are never captured.
		if not wordCom.isDocumentFocused() or not isinstance(gesture, KeyboardInputGesture):
			return super().getScript(gesture)
		if gesture.isModifier:
			return None
		if self._classify(gesture) is not None:
			return self.script_modeKey
		# Any other key leaves the mode and keeps its normal behaviour in Word.
		self._exit(None)
		return super().getScript(gesture)

	@staticmethod
	def _classify(gesture: KeyboardInputGesture) -> tuple | None:
		vk = gesture.vkCode
		modifiers = frozenset(name.lower() for name in gesture.modifierNames)
		if vk == VK_F8 and modifiers == {"nvda", "shift"}:
			return ("toggle",)
		if not modifiers:
			if vk in EXTENDED_ONLY and gesture.isExtended:
				return EXTENDED_ONLY[vk]
			if vk in PLAIN_KEYS:
				return PLAIN_KEYS[vk]
		if vk == VK_F10 and modifiers == {"shift"}:
			return ("actions",)
		if vk == VK_C and modifiers == {"control"}:
			return ("copy", False)
		if vk == VK_C and modifiers == {"control", "shift"}:
			return ("copy", True)
		if ord("A") <= vk <= ord("Z") and modifiers <= {"shift"}:
			return ("letter", vk, "shift" in modifiers)
		if ord("0") <= vk <= ord("9") and modifiers <= {"shift"}:
			# Captured so a stray digit is not typed into the document.
			return ("unknown",)
		return None

	@script(
		# Translators: Description of the command in the Input Gestures dialog.
		description=_("Turns document navigation mode on or off in Microsoft Word"),
		gesture="kb:NVDA+shift+f8",
	)
	def script_toggleNavigationMode(self, gesture):
		if self.active:
			# Translators: Reported when navigation mode ends.
			self._exit(_("Navigation mode off"), collapse=True)
		else:
			self._enter()

	@script(
		# Translators: Description of the command in the Input Gestures dialog.
		description=_("Sets up or updates the AI service and API key of Word Document Assistant"),
		gesture="kb:NVDA+control+shift+f8",
	)
	def script_aiKeySetup(self, gesture):
		wx.CallAfter(keySetup.show)

	def script_modeKey(self, gesture):
		command = self._classify(gesture)
		if command is None:
			return
		try:
			self._run(command)
		except COM_ERRORS:
			log.debugWarning("Word command failed", exc_info=True)
			self.modified()
			sounds.play(sounds.ERROR)
			# Translators: Reported when Word does not answer a request.
			ui.message(_("Word did not respond, please try again"))

	# Entering and leaving the mode ------------------------------------------------

	def _enter(self) -> None:
		try:
			self.ctx = WordContext.fromFocus()
		except WordUnavailable as error:
			ui.message(str(error))
			return
		try:
			self.anchor = self.ctx.selection.Range.Duplicate
		except COM_ERRORS:
			self.anchor = None
		self._ensureFocusMode()
		self._items.clear()
		self._totals.clear()
		self._index = None
		self._inclusive = True
		self.active = True
		sounds.play(sounds.ENTER)
		# Translators: Reported when navigation mode starts; {kind} is the current element type.
		ui.message(_("Navigation mode, {kind}").format(kind=self.kind.label))

	@staticmethod
	def _ensureFocusMode() -> None:
		"""The mode moves the real Word caret, so browse mode is switched off."""
		treeInterceptor = getattr(api.getFocusObject(), "treeInterceptor", None)
		if treeInterceptor is not None and not getattr(treeInterceptor, "passThrough", True):
			try:
				treeInterceptor.passThrough = True
			except Exception:
				log.debugWarning("Cannot switch to focus mode", exc_info=True)

	def _exit(self, message: str | None, collapse: bool = False) -> None:
		if not self.active:
			return
		self.active = False
		if collapse and self._index is not None and self.ctx is not None:
			try:
				# Keep the caret at the current item without leaving text selected.
				self.ctx.selection.Collapse(wordCom.WD_COLLAPSE_START)
			except COM_ERRORS:
				pass
		self._items.clear()
		self._totals.clear()
		self._index = None
		sounds.play(sounds.EXIT)
		if message:
			speech.cancelSpeech()
			ui.message(message)

	# Commands in the mode -------------------------------------------------------

	def _run(self, command: tuple) -> None:
		name = command[0]
		if name == "toggle":
			self._exit(_("Navigation mode off"), collapse=True)
		elif name == "move":
			self._move(command[1])
		elif name == "kind":
			self._changeKind(command[1])
		elif name == "edge":
			self._edge(command[1])
		elif name == "letter":
			self._letter(command[1], command[2])
		elif name == "edit":
			self._edit()
		elif name == "escape":
			self._exit(_("Navigation mode off"), collapse=True)
		elif name == "back":
			self._back()
		elif name == "list":
			self._showList()
		elif name == "actions":
			self._showActions()
		elif name == "copy":
			self._copyCurrent(command[1])
		else:
			sounds.play(sounds.ERROR)

	def getItems(self, kind: ElementType) -> list[Item]:
		items = self._items.get(kind.id)
		if items is None:
			if kind.slow:
				# Translators: Reported while a slow list, such as spelling errors, is collected.
				ui.message(_("Collecting {kind}...").format(kind=kind.label))
			items, total = kind.collect(self.ctx, conf("maxItems"))
			self._items[kind.id] = items
			self._totals[kind.id] = total
			if kind.id == "thaiWords":
				wordListUpdates.announcePending()
				wx.CallAfter(wordListUpdates.askOnce)
		return items

	def isInvalidated(self, kind: ElementType) -> bool:
		return kind.id not in self._items

	def _currentItem(self) -> Item | None:
		items = self._items.get(self.kind.id)
		if items is None or self._index is None or self._indexKind != self.kind.id:
			return None
		if 0 <= self._index < len(items):
			return items[self._index]
		return None

	def _noItems(self, kind: ElementType) -> None:
		sounds.play(sounds.ERROR)
		# Translators: Reported when the document has no items of a type, for example "No Comments".
		ui.message(_("No {kind}").format(kind=kind.label))

	def _move(self, direction: int) -> None:
		kind = self.kind
		items = self.getItems(kind)
		if not items:
			self._noItems(kind)
			return
		if self._index is not None and self._indexKind == kind.id:
			index = self._index + direction
		else:
			index = self._indexFromCaret(items, direction)
		wrapped = False
		if index >= len(items):
			index, wrapped = 0, True
		elif index < 0:
			index, wrapped = len(items) - 1, True
		self._goto(kind, index, wrapped)

	def _indexFromCaret(self, items: list[Item], direction: int) -> int:
		position = wordCom.mainStoryPosition(self.ctx)
		inclusive, self._inclusive = self._inclusive, False
		if direction > 0:
			for index, item in enumerate(items):
				if item.start > position or (inclusive and item.start == position):
					return index
			return len(items)
		for index in range(len(items) - 1, -1, -1):
			if items[index].start < position:
				return index
		return -1

	def _goto(self, kind: ElementType, index: int, wrapped: bool = False) -> None:
		items = self.getItems(kind)
		item = items[index]
		kind.moveTo(self.ctx, item)
		self._index = index
		self._indexKind = kind.id
		text = kind.summary(self.ctx, item)
		if conf("reportPosition"):
			# Translators: Position of the item, for example "3 of 12".
			text += ", " + _("{number} of {total}").format(number=index + 1, total=self._totals[kind.id])
		if wrapped:
			sounds.play(sounds.WRAP)
			if not conf("playSounds"):
				# Translators: Reported when navigation wraps around the document.
				text = _("Wrapped") + ", " + text
		speech.cancelSpeech()
		ui.message(text)

	def _edge(self, direction: int) -> None:
		items = self.getItems(self.kind)
		if not items:
			self._noItems(self.kind)
			return
		self._goto(self.kind, 0 if direction < 0 else len(items) - 1)

	def _changeKind(self, direction: int) -> None:
		count = len(registry.NAVIGATION_TYPES)
		self._kindIndex = (self._kindIndex + direction) % count
		self._index = None
		kind = self.kind
		total = len(self.getItems(kind)) and self._totals[kind.id]
		speech.cancelSpeech()
		# Translators: Reported when the element type changes, for example "Comments, 5".
		ui.message(_("{kind}, {count}").format(kind=kind.label, count=total))

	def _letter(self, vk: int, shift: bool) -> None:
		kind = registry.TYPES_BY_VK.get(vk)
		if kind is not None:
			newIndex = registry.NAVIGATION_TYPES.index(kind)
			if newIndex != self._kindIndex:
				self._kindIndex = newIndex
				self._index = None
			self._move(-1 if shift else 1)
			return
		if vk == CMD_TRACK:
			message = wordCom.toggleTrackChanges(self.ctx) if shift else wordCom.trackChangesStatus(self.ctx)
			if shift:
				self.modified()
		elif vk == CMD_NOTE:
			# Translators: Reported when there is no footnote or endnote at the caret.
			message = wordCom.noteAtCaret(self.ctx) or _("No footnote or endnote at the caret")
		elif vk == CMD_POSITION:
			message = wordCom.positionInfo(self.ctx)
		elif vk == CMD_SUMMARY:
			# Translators: Reported while the document summary is being collected.
			ui.message(_("Counting..."))
			message = wordCom.documentSummary(self.ctx)
		elif vk == CMD_AI:
			wx.CallAfter(self._aiProofread, shift)
			return
		elif vk == CMD_FORMAT:
			wx.CallAfter(self._aiFormat)
			return
		elif vk == CMD_AI_TOOLS:
			wx.CallAfter(self._aiToolsMenu)
			return
		elif vk == CMD_THAI_TOOLS:
			wx.CallAfter(self._thaiToolsMenu)
			return
		elif vk == CMD_HELP:
			# Translators: Title of the help window.
			ui.browseableMessage(_helpText(), _("Word Document Assistant help"))
			return
		else:
			sounds.play(sounds.ERROR)
			return
		speech.cancelSpeech()
		ui.message(message)

	def _edit(self) -> None:
		item = self._currentItem()
		message = None
		if item is not None:
			message = self.kind.edit(self.ctx, item)
		# Translators: Reported when Enter leaves navigation mode.
		self._exit(message or _("Ready to edit"))

	def _back(self) -> None:
		if self.anchor is not None:
			try:
				self.anchor.Select()
			except COM_ERRORS:
				pass
		# Translators: Reported when Backspace returns the caret to where the mode started.
		self._exit(_("Back at the starting position"))

	def _copyCurrent(self, everything: bool) -> None:
		item = self._currentItem()
		if item is None:
			sounds.play(sounds.ERROR)
			# Translators: Reported when a command needs a current item but none is selected yet.
			ui.message(_("No current item, use the arrow keys first"))
			return
		text = self.kind.detailsText(self.ctx, item) if everything else self.kind.copyText(self.ctx, item)
		self.copy(text)

	def _showActions(self) -> None:
		kind = self.kind
		actions = kind.actions(self.ctx, self._currentItem(), self.anchor, self)
		if not actions:
			sounds.play(sounds.ERROR)
			return

		def show():
			gui.mainFrame.prePopup()
			try:
				action = dialogs.chooseAction(gui.mainFrame, actions)
			finally:
				gui.mainFrame.postPopup()
			if action is not None:
				wx.CallAfter(self.runAction, action)

		wx.CallAfter(show)

	def _showList(self) -> None:
		kind = self.kind
		items = self.getItems(kind)
		startIndex = self._index if self._indexKind == kind.id else None
		if startIndex is None and items:
			startIndex = min(self._indexFromCaret(items, 1), len(items) - 1)

		def show():
			try:
				_result, chosen = dialogs.runModal(
					lambda parent: dialogs.ElementListDialog(parent, self, kind, startIndex),
					lambda dialog: dialog.chosenIndex,
				)
			except COM_ERRORS:
				log.debugWarning("Element list failed", exc_info=True)
				self.speak(_("Word did not respond, please try again"))
				return
			if chosen is not None and self.active:
				core.callLater(150, self._gotoSafely, kind, chosen)

		wx.CallAfter(show)

	def _gotoSafely(self, kind: ElementType, index: int) -> None:
		try:
			if index < len(self.getItems(kind)):
				self._goto(kind, index)
		except COM_ERRORS:
			log.debugWarning("Cannot move to item", exc_info=True)

	# Host interface used by element actions --------------------------------------

	def runAction(self, action: Action) -> None:
		try:
			action.run()
		except COM_ERRORS:
			log.debugWarning("Action failed", exc_info=True)
			self.modified()
			# Translators: Reported when Word refuses a change, for example in a protected document.
			self.speak(_("Word could not complete this action"))

	def speak(self, message: str) -> None:
		# A short delay lets NVDA finish announcing the focus returning to Word.
		core.callLater(200, ui.message, message)

	def prompt(self, title: str, label: str, value: str = "") -> str | None:
		result, text = dialogs.runModal(
			lambda parent: dialogs.TextInputDialog(parent, title, label, value),
			lambda dialog: dialog.value,
		)
		return text if result == wx.ID_OK else None

	def confirm(self, title: str, message: str) -> bool:
		result, _value = dialogs.runModal(
			lambda parent: wx.MessageDialog(
				parent, message, title, wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION
			),
		)
		return result == wx.ID_YES

	def copy(self, text: str) -> None:
		if text and api.copyToClip(text):
			# Translators: Reported after text was copied to the clipboard.
			self.speak(_("Copied"))
		else:
			sounds.play(sounds.ERROR)

	def modified(self) -> None:
		"""The document changed: collect items again on the next command."""
		self._items.clear()
		self._totals.clear()
		self._index = None
		self._inclusive = True

	# AI suggestions ------------------------------------------------------------------

	def aiAvailable(self) -> bool:
		return ai.isEnabled()

	def _aiReady(self) -> bool:
		if not ai.isEnabled() or (conf("aiProvider") in ai.KEY_REQUIRED and not conf("aiApiKey")):
			# Nothing is set up yet: open the setup dialog straight away.
			# Translators: Reported when AI is used before it was set up.
			ui.message(_("AI is not set up yet. Opening the AI setup"))
			keySetup.show()
			return False
		if ai.isCloud() and not conf("aiConsent"):
			agreed = self.confirm(
				# Translators: Title of the question before text is first sent to an online AI service.
				_("Send text to the AI service"),
				# Translators: Question before text is first sent to an online AI service.
				_(
					"The text you check will be sent to the online AI service chosen in the settings. "
					"Free services may use it to improve their products. "
					"Do not use this for confidential documents. Continue?",
				),
			)
			if not agreed:
				return False
			config.conf[elements.CONFIG_SECTION]["aiConsent"] = True
		return True

	def _aiFailed(self, message: str, needsKey: bool = False) -> None:
		sounds.play(sounds.ERROR)
		ui.message(message)
		if needsKey and self.confirm(
			# Translators: Title of the question offered when the API key does not work.
			_("API key"),
			# Translators: Question offered when the API key is missing or rejected.
			_("{message}. Open the AI setup to get or update the API key now?").format(message=message),
		):
			keySetup.show()

	def _aiAsking(self) -> None:
		# Translators: Reported while waiting for the AI service.
		ui.message(_("Asking AI, please wait..."))

	def _aiProofread(self, selectionOnly: bool) -> None:
		if self.ctx is None or not self._aiReady():
			return
		try:
			if selectionOnly:
				rangeObj = self.anchor
				if rangeObj is None or rangeObj.End <= rangeObj.Start:
					sounds.play(sounds.ERROR)
					# Translators: Reported when Shift+W is used without selected text.
					ui.message(_("Select text before entering navigation mode, then press Shift+W"))
					return
				rangeObj = rangeObj.Duplicate
			else:
				item = self._currentItem()
				base = item.navRange if item is not None else self.ctx.selection.Range
				rangeObj = base.Paragraphs.Item(1).Range
			start = rangeObj.Start
			text = wordCom.sameLengthText(rangeObj.Text)[: ai.MAX_TEXT]
		except COM_ERRORS:
			log.debugWarning("Cannot read text for AI", exc_info=True)
			self._aiFailed(_("Word did not respond, please try again"))
			return
		if not text.strip():
			sounds.play(sounds.ERROR)
			# Translators: Reported when there is no text to check.
			ui.message(_("There is no text to check"))
			return
		self._aiAsking()
		ai.runInBackground(
			lambda: ai.proofread(text),
			lambda corrections: self._aiShowCorrections(start, text, corrections),
			self._aiFailed,
		)

	def _aiShowCorrections(self, start: int, text: str, corrections: list[dict[str, str]]) -> None:
		if not corrections:
			# Translators: Reported when AI finds nothing to correct.
			ui.message(_("AI found nothing to correct"))
			return
		# Translators: Labels of the detail fields in the AI suggestion dialog.
		originalLabel, suggestionLabel, reasonLabel = _("Original"), _("Suggestion"), _("Reason")
		options = [
			# Translators: One AI correction in the list, for example "teh to the".
			_("{original} to {suggestion}").format(original=c["original"], suggestion=c["suggestion"])
			for c in corrections
		]
		details = [
			[(originalLabel, c["original"]), (suggestionLabel, c["suggestion"]), (reasonLabel, c["reason"])]
			for c in corrections
		]
		result, chosen = dialogs.runModal(
			lambda parent: dialogs.ChoiceDialog(
				parent,
				# Translators: Title of the AI correction dialog; {count} is a number.
				_("AI suggestions: {count}").format(count=len(corrections)),
				# Translators: Label of the AI correction list.
				_("Check the corrections to use with the Space bar:"),
				options,
				details,
				multiple=True,
				# Translators: Button that applies the checked AI corrections.
				okLabel=_("&Apply checked"),
				suggestionLabel=suggestionLabel,
			),
			lambda dialog: dialog.chosen,
		)
		if result != wx.ID_OK or not chosen:
			return
		self._aiApply(start, text, [corrections[i] for i in chosen])

	def _aiApply(self, start: int, text: str, corrections: list[dict[str, str]]) -> None:
		# Find where each chosen correction is, using each occurrence only once.
		placed: list[tuple[int, dict[str, str]]] = []
		used: list[tuple[int, int]] = []
		for correction in corrections:
			original = correction["original"]
			offset = text.find(original)
			while offset >= 0 and any(a < offset + len(original) and offset < b for a, b in used):
				offset = text.find(original, offset + 1)
			if offset >= 0:
				used.append((offset, offset + len(original)))
				placed.append((offset, correction))
		applied = 0
		try:
			# Translators: Name of the undo step for AI corrections.
			with wordCom.undoRecord(self.ctx, _("AI corrections")):
				# From the end so earlier positions stay valid.
				for offset, correction in sorted(placed, key=lambda p: p[0], reverse=True):
					rangeObj = self.ctx.document.Range(
						start + offset, start + offset + len(correction["original"])
					)
					if rangeObj.Text != correction["original"]:
						continue
					rangeObj.Text = correction["suggestion"]
					applied += 1
		except COM_ERRORS:
			log.debugWarning("Cannot apply AI corrections", exc_info=True)
		self.modified()
		# Translators: Reported after AI corrections were applied.
		message = _("{applied} of {total} corrections applied").format(
			applied=applied, total=len(corrections)
		)
		if applied < len(corrections):
			# Translators: Added when some corrections could not be applied.
			message += ". " + _("Some text had already changed and was skipped")
		self.speak(message)

	def _aiPick(
		self,
		title: str,
		label: str,
		options: list[str],
		reasons: list[str],
		okLabel: str,
	) -> str | None:
		# Translators: Labels of the detail fields in the AI suggestion dialog.
		suggestionLabel, reasonLabel = _("Suggestion"), _("Reason")
		details = [[(suggestionLabel, o), (reasonLabel, r)] for o, r in zip(options, reasons, strict=True)]
		result, chosen = dialogs.runModal(
			lambda parent: dialogs.ChoiceDialog(
				parent,
				title,
				label,
				options,
				details,
				multiple=False,
				okLabel=okLabel,
				suggestionLabel=suggestionLabel,
			),
			lambda dialog: dialog.chosen,
		)
		if result != wx.ID_OK or not chosen:
			return None
		return options[chosen[0]]

	def aiWordSuggestions(self, word: str, context: str, apply) -> None:
		if not self._aiReady():
			return

		def done(suggestions: list[dict[str, str]]) -> None:
			if not suggestions:
				# Translators: Reported when AI has no suggestions.
				ui.message(_("AI has no suggestions"))
				return
			choice = self._aiPick(
				# Translators: Title of the dialog with AI suggestions for one word.
				_("AI suggestions for {word}").format(word=word),
				# Translators: Label of the list of AI suggestions for one word.
				_("Suggestions:"),
				[s["text"] for s in suggestions],
				[s["reason"] for s in suggestions],
				# Translators: Button that replaces the word with the chosen AI suggestion.
				_("&Replace"),
			)
			if choice:
				self._runSafely(lambda: apply(choice))

		self._aiAsking()
		ai.runInBackground(lambda: ai.suggestWord(word, context), done, self._aiFailed)

	def aiReplyDrafts(self, author: str, comment: str, commentedText: str, apply) -> None:
		if not self._aiReady():
			return

		def done(replies: list[str]) -> None:
			if not replies:
				ui.message(_("AI has no suggestions"))
				return
			choice = self._aiPick(
				# Translators: Title of the dialog with AI reply drafts.
				_("AI reply drafts"),
				# Translators: Label of the list of AI reply drafts.
				_("Choose a draft, you can edit it before it is posted:"),
				replies,
				[""] * len(replies),
				# Translators: Button that takes the chosen AI reply draft for editing.
				_("&Use this draft"),
			)
			if not choice:
				return
			# The user always reviews and edits the draft before it is posted.
			text = self.prompt(
				_("Reply to comment"),
				_("Reply to {author}:").format(author=author),
				choice,
			)
			if text:
				self._runSafely(lambda: apply(text))

		self._aiAsking()
		ai.runInBackground(lambda: ai.draftReplies(author, comment, commentedText), done, self._aiFailed)

	def _runSafely(self, func) -> None:
		self.runAction(Action("", func))

	# Formatting with AI --------------------------------------------------------------

	#: Most paragraphs formatted in one go.
	MAX_FORMAT_PARAGRAPHS = 3000

	def _aiFormat(self) -> None:
		if self.ctx is None or not self._aiReady():
			return
		try:
			content = self.ctx.document.Content
			sample = wordCom.cleanText(self.ctx.document.Range(0, min(content.End, 3000)).Text)
			anchor = self.anchor
			hasSelection = anchor is not None and anchor.End > anchor.Start
		except COM_ERRORS:
			self._aiFailed(_("Word did not respond, please try again"))
			return
		if not sample.strip():
			sounds.play(sounds.ERROR)
			ui.message(_("There is no text to check"))
			return
		result, choice = dialogs.runModal(
			lambda parent: formatDialog.FormatChooserDialog(parent, sample, hasSelection, recommend=True),
			lambda dialog: (dialog.chosen, dialog.selectionOnly, dialog.options),
		)
		if result != wx.ID_OK or choice[0] is None:
			return
		preset, selectionOnly, options = choice
		try:
			rangeObj = anchor.Duplicate if selectionOnly else self.ctx.document.Content
			paragraphs, texts = presets.collectParagraphs(rangeObj, self.MAX_FORMAT_PARAGRAPHS)
		except COM_ERRORS:
			log.debugWarning("Cannot read paragraphs", exc_info=True)
			self._aiFailed(_("Word did not respond, please try again"))
			return
		if not paragraphs:
			ui.message(_("There is no text to check"))
			return
		if len(paragraphs) >= self.MAX_FORMAT_PARAGRAPHS:
			# Translators: Reported when the document is too long to format at once.
			ui.message(
				_("Only the first {count} paragraphs will be formatted").format(
					count=self.MAX_FORMAT_PARAGRAPHS
				)
			)
		hints = {r: presets.ROLE_HINTS[r] for r in preset.roleIds if r in presets.ROLE_HINTS}
		# Translators: Reported while AI looks at the paragraphs; {count} is a number.
		ui.message(_("AI is reading {count} paragraphs, please wait...").format(count=len(paragraphs)))

		def progress(done: int, total: int) -> None:
			if done < total:
				# Translators: Progress while AI reads the paragraphs.
				wx.CallAfter(ui.message, _("{done} of {total} paragraphs").format(done=done, total=total))

		ai.runInBackground(
			lambda: ai.classifyParagraphs(texts, hints, preset.name, progress),
			lambda roles: self._aiFormatPreview(preset, options, paragraphs, texts, roles, not selectionOnly),
			self._aiFailed,
		)

	def _aiFormatPreview(self, preset, options, paragraphs, texts, roles, wholeDocument: bool) -> None:
		result, roles = dialogs.runModal(
			lambda parent: formatDialog.FormatPreviewDialog(parent, preset, texts, roles),
			lambda dialog: dialog.roles,
		)
		if result != wx.ID_OK:
			return
		# Translators: Reported while the document is being formatted.
		ui.message(_("Formatting..."))
		core.callLater(300, self._aiFormatApply, preset, options, paragraphs, roles, wholeDocument)

	def _aiFormatApply(self, preset, options, paragraphs, roles, wholeDocument: bool) -> None:
		try:
			counts = presets.apply(self.ctx, preset, paragraphs, roles, options, wholeDocument)
		except COM_ERRORS:
			log.debugWarning("Formatting failed", exc_info=True)
			self.modified()
			self._aiFailed(_("Word could not complete this action"))
			return
		self.modified()
		# Translators: Reported after formatting; {formatted} and {removed} are numbers.
		message = _("Formatted {formatted} paragraphs as {name}").format(
			formatted=counts["formatted"], name=preset.name
		)
		if counts["removed"]:
			# Translators: Added after formatting when empty paragraphs were removed.
			message += ", " + _("{removed} empty paragraphs removed").format(removed=counts["removed"])
		# Translators: Added after formatting, telling how to undo.
		message += ". " + _("Control+Z undoes it")
		self.speak(message)

	# Host methods used by element types ---------------------------------------------------

	def choose(self, title: str, label: str, options: list[str]) -> int | None:
		if not options:
			return None

		def make(parent):
			dialog = wx.SingleChoiceDialog(parent, label, title, options)
			dialog.SetSelection(0)
			return dialog

		result, index = dialogs.runModal(make, lambda dialog: dialog.GetSelection())
		return index if result == wx.ID_OK and index is not None and index >= 0 else None

	def aiDraft(self, title: str, label: str, work, apply) -> None:
		"""Run an AI request, let the user edit the answer, then apply it."""
		if not self._aiReady():
			return

		def done(text: str) -> None:
			edited = self.prompt(title, label, text)
			if edited:
				self._runSafely(lambda: apply(edited))

		self._aiAsking()
		ai.runInBackground(work, done, self._aiFailed)

	def aiFormat(self) -> None:
		wx.CallAfter(self._aiFormat)

	# Target text for tools ------------------------------------------------------------------

	def _toolRange(self, wholeDocumentDefault: bool):
		"""(range, description): the text selected before the mode, else the document or the current paragraph."""
		anchor = self.anchor
		if anchor is not None and anchor.End > anchor.Start:
			# Translators: What a tool works on.
			return anchor.Duplicate, _("selected text")
		if wholeDocumentDefault:
			# Translators: What a tool works on.
			return self.ctx.document.Content, _("whole document")
		item = self._currentItem()
		base = item.navRange if item is not None else self.ctx.selection.Range
		# Translators: What a tool works on.
		return base.Paragraphs.Item(1).Range, _("current paragraph")

	# AI tools (V) --------------------------------------------------------------------------------

	def _aiToolsMenu(self) -> None:
		if self.ctx is None or not self._aiReady():
			return
		tools_ = [
			# Translators: AI tool.
			(_("Summarize (short)"), lambda: self._aiResult("short")),
			# Translators: AI tool.
			(_("Summarize section by section"), lambda: self._aiResult("detailed")),
			# Translators: AI tool.
			(_("Key points, dates and actions"), lambda: self._aiResult("keyPoints")),
			# Translators: AI tool.
			(_("Rewrite in official Thai style"), lambda: self._aiRewrite("official")),
			# Translators: AI tool.
			(_("Rewrite more politely"), lambda: self._aiRewrite("polite")),
			# Translators: AI tool.
			(_("Rewrite shorter and clearer"), lambda: self._aiRewrite("concise")),
			# Translators: AI tool.
			(_("Rewrite in easy-to-read language"), lambda: self._aiRewrite("easy")),
			# Translators: AI tool.
			(_("Correct spelling and grammar only"), lambda: self._aiRewrite("fix")),
			# Translators: AI tool.
			(_("Translate to English"), lambda: self._aiTranslate("en")),
			# Translators: AI tool.
			(_("Translate to Thai"), lambda: self._aiTranslate("th")),
			# Translators: AI tool.
			(_("Summarize all comments"), self._aiComments),
			# Translators: AI tool.
			(_("Check consistency (spellings, digits, years, dates)"), self._aiConsistency),
		]
		# Translators: Title and label of the AI tools menu.
		index = self.choose(
			_("AI tools"),
			_("Tool (summaries use the selected text, or the whole document):"),
			[t[0] for t in tools_],
		)
		if index is not None:
			tools_[index][1]()

	def _readTool(self, wholeDocumentDefault: bool):
		try:
			rangeObj, where = self._toolRange(wholeDocumentDefault)
			text = wordCom.cleanText(rangeObj.Text)
		except COM_ERRORS:
			self._aiFailed(_("Word did not respond, please try again"))
			return None
		if not text.strip():
			sounds.play(sounds.ERROR)
			ui.message(_("There is no text to check"))
			return None
		return rangeObj, where, text

	def _showResult(self, title: str, text: str) -> None:
		dialogs.runModal(lambda parent: tools.ResultDialog(parent, title, text))

	def _aiResult(self, length: str) -> None:
		found = self._readTool(wholeDocumentDefault=True)
		if found is None:
			return
		_rangeObj, where, text = found
		self._aiAsking()
		# Translators: Title of the summary dialog; {where} is what was summarized.
		title = _("Summary of the {where}").format(where=where)
		ai.runInBackground(
			lambda: ai.summarize(text, length), lambda result: self._showResult(title, result), self._aiFailed
		)

	def _compare(self, title: str, rangeObj, original: str, proposed: str, canInsert: bool) -> None:
		result, value = dialogs.runModal(
			lambda parent: tools.CompareDialog(parent, title, original, proposed, canInsert),
			lambda dialog: dialog.value,
		)
		if result not in (tools.REPLACE, tools.INSERT_AFTER) or not value:
			return

		def apply() -> None:
			tools.replaceText(self.ctx, rangeObj, value, after=result == tools.INSERT_AFTER)
			self.modified()
			# Translators: Reported after AI text was put in the document.
			self.speak(_("Done. Control+Z undoes it"))

		self._runSafely(apply)

	def _aiRewrite(self, style: str) -> None:
		found = self._readTool(wholeDocumentDefault=False)
		if found is None:
			return
		rangeObj, where, text = found
		self._aiAsking()
		# Translators: Title of the rewrite dialog; {where} is what was rewritten.
		title = _("Rewrite the {where}").format(where=where)
		ai.runInBackground(
			lambda: ai.rewrite(text, style),
			lambda result: self._compare(title, rangeObj, text, result, canInsert=False),
			self._aiFailed,
		)

	def _aiTranslate(self, target: str) -> None:
		found = self._readTool(wholeDocumentDefault=False)
		if found is None:
			return
		rangeObj, where, text = found
		self._aiAsking()
		# Translators: Title of the translation dialog; {where} is what was translated.
		title = _("Translate the {where}").format(where=where)
		ai.runInBackground(
			lambda: ai.translate(text, target),
			lambda result: self._compare(title, rangeObj, text, result, canInsert=True),
			self._aiFailed,
		)

	def _aiComments(self) -> None:
		try:
			text = tools.commentsText(self.ctx)
		except COM_ERRORS:
			text = ""
		if not text:
			self._noItems(registry.TYPES_BY_VK[ord("C")])
			return
		self._aiAsking()
		ai.runInBackground(
			lambda: ai.summarizeComments(text),
			# Translators: Title of the comment summary dialog.
			lambda result: self._showResult(_("Summary of the comments"), result),
			self._aiFailed,
		)

	def _aiConsistency(self) -> None:
		found = self._readTool(wholeDocumentDefault=True)
		if found is None:
			return
		_rangeObj, where, text = found
		self._aiAsking()
		ai.runInBackground(
			lambda: ai.checkConsistency(text),
			# Translators: Title of the consistency check dialog; {where} is what was checked.
			lambda result: self._showResult(_("Consistency of the {where}").format(where=where), result),
			self._aiFailed,
		)

	# Thai tools (Y) -------------------------------------------------------------------------------

	def _thaiToolsMenu(self) -> None:
		if self.ctx is None:
			return
		try:
			rangeObj, where = self._toolRange(wholeDocumentDefault=True)
			text = str(rangeObj.Text or "")
		except COM_ERRORS:
			self._aiFailed(_("Word did not respond, please try again"))
			return
		thai, arabic = tools.countDigits(text)
		options = [
			# Translators: Thai tool; {count} is how many digits will change.
			(
				_("Arabic digits to Thai digits ({count})").format(count=arabic),
				lambda: self._digits(rangeObj, True),
			),
			# Translators: Thai tool; {count} is how many digits will change.
			(
				_("Thai digits to Arabic digits ({count})").format(count=thai),
				lambda: self._digits(rangeObj, False),
			),
			# Translators: Thai tool.
			(_("Christian era years (ค.ศ.) to Buddhist era (พ.ศ.)"), lambda: self._era(rangeObj, True)),
			# Translators: Thai tool.
			(_("Buddhist era years (พ.ศ.) to Christian era (ค.ศ.)"), lambda: self._era(rangeObj, False)),
		]
		# Translators: Title and label of the Thai tools menu; {where} is what the tools work on.
		index = self.choose(
			_("Thai tools"), _("Tool, works on the {where}:").format(where=where), [o[0] for o in options]
		)
		if index is not None:
			self._runSafely(options[index][1])

	def _digits(self, rangeObj, toThai: bool) -> None:
		changed = tools.convertDigits(self.ctx, rangeObj, toThai)
		self.modified()
		# Translators: Reported after digits were changed; {count} is a number.
		self.speak(_("{count} digits changed").format(count=changed) if changed else _("Nothing to change"))

	def _era(self, rangeObj, toBuddhist: bool) -> None:
		changed = tools.convertEra(self.ctx, rangeObj, toBuddhist)
		self.modified()
		# Translators: Reported after years were changed; {count} is a number.
		self.speak(_("{count} years changed").format(count=changed) if changed else _("Nothing to change"))

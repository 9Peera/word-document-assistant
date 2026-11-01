# Word Document Assistant for NVDA
# Dialogs: element list with separated details, text input.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any

import addonHandler
import wx
from gui import guiHelper
from logHandler import log

from .elements import Action, ElementType, Item
from .wordCom import COM_ERRORS

if TYPE_CHECKING:
	from . import GlobalPlugin

addonHandler.initTranslation()


class TextInputDialog(wx.Dialog):
	"""Multi-line text input. Enter adds a new line, Control+Enter confirms."""

	def __init__(self, parent: wx.Window, title: str, label: str, value: str = ""):
		super().__init__(parent, title=title)
		sizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		self.edit = helper.addLabeledControl(
			label,
			wx.TextCtrl,
			value=value,
			style=wx.TE_MULTILINE,
			size=(500, 120),
		)
		# Translators: Hint under the text field of the input dialog.
		helper.addItem(wx.StaticText(self, label=_("Press Control+Enter to confirm.")))
		helper.addDialogDismissButtons(wx.OK | wx.CANCEL)
		sizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizerAndFit(sizer)
		self.edit.Bind(wx.EVT_KEY_DOWN, self._onKey)
		self.edit.SetFocus()
		self.edit.SetInsertionPointEnd()
		self.CentreOnScreen()

	def _onKey(self, event: wx.KeyEvent) -> None:
		if event.GetKeyCode() in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER) and event.ControlDown():
			self.EndModal(wx.ID_OK)
			return
		event.Skip()

	@property
	def value(self) -> str:
		return self.edit.GetValue().strip()


class ElementListDialog(wx.Dialog):
	"""Lists all items of one element type, with each detail in its own read-only field."""

	def __init__(self, parent: wx.Window, plugin: GlobalPlugin, kind: ElementType, startIndex: int | None):
		# Translators: Title of the element list dialog, for example "Comments list".
		super().__init__(parent, title=_("{kind} list").format(kind=kind.label))
		self.plugin = plugin
		self.kind = kind
		self.chosenIndex: int | None = None
		self.items: list[Item] = []
		self.rows: list[list[str]] = []
		self.visible: list[int] = []

		mainSizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		# Translators: Label of the filter field in the element list dialog.
		self.filterEdit = helper.addLabeledControl(_("&Filter:"), wx.TextCtrl)
		self.list = helper.addLabeledControl(
			# Translators: Label of the list in the element list dialog.
			_("&Items:"),
			wx.ListCtrl,
			style=wx.LC_REPORT | wx.LC_SINGLE_SEL,
			size=(760, 220),
		)
		for column, header in enumerate(kind.columns):
			self.list.InsertColumn(column, header, width=200 if column else 160)

		# Each detail gets its own labelled read-only field so it can be read and copied on its own.
		self.sectionFields: list[wx.TextCtrl] = []
		for section in kind.sections:
			style = wx.TE_READONLY | (wx.TE_MULTILINE | wx.TE_RICH2 if section.multiline else 0)
			field = helper.addLabeledControl(
				f"{section.label}:",
				wx.TextCtrl,
				style=style,
				size=(760, 70 if section.multiline else -1),
			)
			self.sectionFields.append(field)

		buttons = guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: Button that moves the caret to the selected item.
		self.gotoButton = buttons.addButton(self, wx.ID_OK, label=_("&Go to"))
		# Translators: Button that opens the actions for the selected item.
		self.actionsButton = buttons.addButton(self, label=_("&Actions..."))
		# Translators: Button that copies the main text of the selected item.
		self.copyButton = buttons.addButton(self, label=_("&Copy text"))
		# Translators: Button that copies all details of the selected item.
		self.copyAllButton = buttons.addButton(self, label=_("Copy all &details"))
		# Translators: Button that closes the element list dialog.
		buttons.addButton(self, wx.ID_CLOSE, label=_("Clo&se"))
		helper.addItem(buttons)
		mainSizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizerAndFit(mainSizer)
		self.SetEscapeId(wx.ID_CLOSE)
		self.gotoButton.SetDefault()

		self.filterEdit.Bind(wx.EVT_TEXT, lambda evt: self._applyFilter())
		self.list.Bind(wx.EVT_LIST_ITEM_FOCUSED, lambda evt: self._showDetails())
		self.list.Bind(wx.EVT_LIST_ITEM_ACTIVATED, lambda evt: self._onGoto())
		self.list.Bind(wx.EVT_CONTEXT_MENU, lambda evt: self._onActions())
		self.Bind(wx.EVT_BUTTON, lambda evt: self._onGoto(), id=wx.ID_OK)
		self.actionsButton.Bind(wx.EVT_BUTTON, lambda evt: self._onActions())
		self.copyButton.Bind(wx.EVT_BUTTON, lambda evt: self._copy(False))
		self.copyAllButton.Bind(wx.EVT_BUTTON, lambda evt: self._copy(True))
		self.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(wx.ID_CLOSE), id=wx.ID_CLOSE)

		self._load(startIndex)
		self.list.SetFocus()
		self.CentreOnScreen()

	# Data -------------------------------------------------------------------

	def _load(self, selectIndex: int | None) -> None:
		ctx = self.plugin.ctx
		self.items = self.plugin.getItems(self.kind)
		self.rows = []
		for item in self.items:
			try:
				self.rows.append(self.kind.row(ctx, item))
			except COM_ERRORS:
				log.debugWarning("Cannot read item", exc_info=True)
				self.rows.append(["?"] * len(self.kind.columns))
		self._applyFilter(selectIndex)

	def _applyFilter(self, selectIndex: int | None = None) -> None:
		if selectIndex is None:
			selectIndex = self._selectedItemIndex()
		needle = self.filterEdit.GetValue().strip().lower()
		self.visible = [
			index for index, row in enumerate(self.rows) if not needle or needle in " ".join(row).lower()
		]
		self.list.DeleteAllItems()
		for index in self.visible:
			row = self.rows[index]
			position = self.list.InsertItem(self.list.GetItemCount(), row[0])
			for column, value in enumerate(row[1:], start=1):
				self.list.SetItem(position, column, value)
		if self.visible:
			listPosition = 0
			if selectIndex is not None:
				# Select the item itself, or the nearest one after it.
				listPosition = next(
					(pos for pos, index in enumerate(self.visible) if index >= selectIndex),
					len(self.visible) - 1,
				)
			self.list.Select(listPosition)
			self.list.Focus(listPosition)
		self._showDetails()

	def _selectedItemIndex(self) -> int | None:
		position = self.list.GetFirstSelected()
		if position < 0 or position >= len(self.visible):
			return None
		return self.visible[position]

	def _selectedItem(self) -> Item | None:
		index = self._selectedItemIndex()
		return None if index is None else self.items[index]

	def _showDetails(self) -> None:
		item = self._selectedItem()
		texts: list[str] = []
		if item is not None:
			try:
				texts = self.kind.sectionTexts(self.plugin.ctx, item)
			except COM_ERRORS:
				log.debugWarning("Cannot read item details", exc_info=True)
		for position, field in enumerate(self.sectionFields):
			field.SetValue(texts[position] if position < len(texts) else "")
		hasItem = item is not None
		self.gotoButton.Enable(hasItem)
		self.copyButton.Enable(hasItem)
		self.copyAllButton.Enable(hasItem)

	# Commands --------------------------------------------------------------

	def _onGoto(self) -> None:
		index = self._selectedItemIndex()
		if index is None:
			return
		self.chosenIndex = index
		self.EndModal(wx.ID_OK)

	def _copy(self, everything: bool) -> None:
		item = self._selectedItem()
		if item is None:
			return
		ctx = self.plugin.ctx
		try:
			text = self.kind.detailsText(ctx, item) if everything else self.kind.copyText(ctx, item)
		except COM_ERRORS:
			return
		self.plugin.copy(text)

	def _onActions(self) -> None:
		index = self._selectedItemIndex()
		item = None if index is None else self.items[index]
		try:
			actions = self.kind.actions(self.plugin.ctx, item, self.plugin.anchor, self.plugin)
		except COM_ERRORS:
			return
		action = chooseAction(self, actions)
		if action is None:
			return
		self.plugin.runAction(action)
		if self.plugin.isInvalidated(self.kind):
			self._load(index)


def chooseAction(window: wx.Window, actions: list[Action]) -> Action | None:
	"""Show the actions as a popup menu and return the one chosen."""
	if not actions:
		return None
	menu = wx.Menu()
	byId: dict[int, Action] = {}
	for action in actions:
		menuItem = menu.Append(wx.ID_ANY, action.label)
		menuItem.Enable(action.enabled)
		byId[menuItem.GetId()] = action
	try:
		chosenId = window.GetPopupMenuSelectionFromUser(menu)
	finally:
		menu.Destroy()
	return byId.get(chosenId)


def runModal(
	dialogFactory: Callable[[wx.Window], wx.Dialog],
	read: Callable[[Any], Any] = lambda dialog: None,
) -> tuple[int, Any]:
	"""Show a dialog modally between NVDA's prePopup and postPopup.

	read is called with the dialog after it closes and before it is destroyed.
	"""
	import gui

	gui.mainFrame.prePopup()
	try:
		dialog = dialogFactory(gui.mainFrame)
		try:
			result = dialog.ShowModal()
			return result, read(dialog)
		finally:
			dialog.Destroy()
	finally:
		gui.mainFrame.postPopup()


class ChoiceDialog(wx.Dialog):
	"""Lets the user pick AI suggestions. Nothing is chosen in advance.

	options are the lines of the list; details holds, for each option, a list of
	(label, text) pairs shown in separate read-only fields.
	"""

	def __init__(
		self,
		parent: wx.Window,
		title: str,
		label: str,
		options: list[str],
		details: list[list[tuple[str, str]]],
		multiple: bool,
		okLabel: str,
		suggestionLabel: str,
	):
		super().__init__(parent, title=title)
		#: Label of the detail field that holds the suggested text, used by the copy button.
		self.suggestionLabel = suggestionLabel
		self.options = options
		self.details = details
		self.multiple = multiple
		self.chosen: list[int] = []
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		listClass = wx.CheckListBox if multiple else wx.ListBox
		self.list = helper.addLabeledControl(label, listClass, choices=options, size=(640, 200))
		fieldLabels = [name for name, _text in details[0]] if details else []
		self.fields: list[wx.TextCtrl] = []
		for name in fieldLabels:
			self.fields.append(
				helper.addLabeledControl(
					f"{name}:",
					wx.TextCtrl,
					style=wx.TE_READONLY | wx.TE_MULTILINE,
					size=(640, 50),
				)
			)
		buttons = guiHelper.ButtonHelper(wx.HORIZONTAL)
		self.okButton = buttons.addButton(self, wx.ID_OK, label=okLabel)
		# Translators: Button that copies the selected AI suggestion.
		copyButton = buttons.addButton(self, label=_("&Copy suggestion"))
		# Translators: Button that closes the AI suggestion dialog without changing anything.
		buttons.addButton(self, wx.ID_CANCEL, label=_("Close without changes"))
		helper.addItem(buttons)
		mainSizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizerAndFit(mainSizer)
		self.list.Bind(wx.EVT_LISTBOX, lambda evt: self._showDetails())
		if not multiple:
			self.list.Bind(wx.EVT_LISTBOX_DCLICK, lambda evt: self._onOk())
		self.Bind(wx.EVT_BUTTON, lambda evt: self._onOk(), id=wx.ID_OK)
		copyButton.Bind(wx.EVT_BUTTON, lambda evt: self._copy())
		if options:
			self.list.SetSelection(0)
		self._showDetails()
		self.list.SetFocus()
		self.CentreOnScreen()

	def _showDetails(self) -> None:
		index = self.list.GetSelection()
		values = self.details[index] if 0 <= index < len(self.details) else []
		for position, field in enumerate(self.fields):
			field.SetValue(values[position][1] if position < len(values) else "")

	def _onOk(self) -> None:
		if self.multiple:
			self.chosen = list(self.list.GetCheckedItems())
			if not self.chosen:
				import ui

				# Translators: Reported when no AI suggestion was checked.
				ui.message(_("Check the suggestions to use with the Space bar first"))
				return
		else:
			index = self.list.GetSelection()
			if index < 0:
				return
			self.chosen = [index]
		self.EndModal(wx.ID_OK)

	def _copy(self) -> None:
		import api
		import ui

		index = self.list.GetSelection()
		if 0 <= index < len(self.details) and self.details[index]:
			text = next((t for name, t in self.details[index] if name == self.suggestionLabel), "")
			if text and api.copyToClip(text):
				# Translators: Reported after text was copied to the clipboard.
				ui.message(_("Copied"))

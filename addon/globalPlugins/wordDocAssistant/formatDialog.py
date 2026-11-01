# Word Document Assistant for NVDA
# Dialogs for formatting a document with AI.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import addonHandler
import ui
import wx
from gui import guiHelper

from . import ai, presets
from .presets import ApplyOptions, Preset

addonHandler.initTranslation()


class FormatChooserDialog(wx.Dialog):
	"""Choose a format; AI recommends one as soon as the dialog opens."""

	def __init__(self, parent: wx.Window, sample: str, hasSelection: bool, recommend: bool):
		# Translators: Title of the dialog for formatting a document with AI.
		super().__init__(parent, title=_("Format document with AI"))
		self.presets = presets.presets()
		self.chosen: Preset | None = None
		self.selectionOnly = False
		self.options = ApplyOptions()
		self._userMoved = False

		mainSizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		self.formats = helper.addLabeledControl(
			# Translators: Label of the list of document formats.
			_("&Format:"),
			wx.ListBox,
			choices=[p.name for p in self.presets],
			size=(560, 190),
		)
		self.description = helper.addLabeledControl(
			# Translators: Label of the field that describes the selected format.
			_("What this format does:"),
			wx.TextCtrl,
			style=wx.TE_READONLY | wx.TE_MULTILINE,
			size=(560, 80),
		)
		self.recommendation = helper.addLabeledControl(
			# Translators: Label of the field with the format AI recommends.
			_("AI recommendation:"),
			wx.TextCtrl,
			style=wx.TE_READONLY | wx.TE_MULTILINE,
			size=(560, 50),
		)
		self.scope = helper.addItem(
			wx.RadioBox(
				self,
				# Translators: Label of the choice of what to format.
				label=_("Format"),
				choices=[
					# Translators: Choice to format the whole document.
					_("&Whole document"),
					# Translators: Choice to format only the text selected before navigation mode.
					_("Selected &text only"),
				],
				majorDimension=1,
				style=wx.RA_SPECIFY_COLS,
			)
		)
		self.scope.EnableItem(1, hasSelection)
		self.pageSetup = helper.addItem(
			# Translators: Check box in the format dialog.
			wx.CheckBox(self, label=_("Set &paper size and margins (whole document only)")),
		)
		self.pageSetup.SetValue(True)
		self.fonts = helper.addItem(
			# Translators: Check box in the format dialog.
			wx.CheckBox(self, label=_("Set fo&nts and sizes")),
		)
		self.fonts.SetValue(True)
		self.removeEmpty = helper.addItem(
			# Translators: Check box in the format dialog.
			wx.CheckBox(self, label=_("Remove &empty paragraphs used for spacing")),
		)
		self.removeEmpty.SetValue(True)
		note = helper.addItem(
			wx.StaticText(
				self,
				label=_(
					# Translators: Note in the format dialog.
					"AI only decides what each paragraph is, such as subject, body or signature. "
					"The words are not changed. You can check and correct the result before anything is formatted.",
				),
			)
		)
		note.Wrap(560)
		buttons = guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: Button that lets AI analyse the document and shows a preview.
		buttons.addButton(self, wx.ID_OK, label=_("&Analyse and preview"))
		buttons.addButton(self, wx.ID_CANCEL, label=_("Cancel"))
		helper.addItem(buttons)
		mainSizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizerAndFit(mainSizer)

		self.formats.Bind(wx.EVT_LISTBOX, self._onFormat)
		self.scope.Bind(wx.EVT_RADIOBOX, lambda evt: self.pageSetup.Enable(self.scope.GetSelection() == 0))
		self.Bind(wx.EVT_BUTTON, lambda evt: self._onOk(), id=wx.ID_OK)
		self.formats.SetSelection(0)
		self._showDescription()
		self.formats.SetFocus()
		self.CentreOnScreen()
		if recommend and sample.strip():
			# Translators: Shown while AI looks at the document.
			self.recommendation.SetValue(_("Asking AI which format fits this document..."))
			ai.runInBackground(
				lambda: ai.recommendFormat(
					sample, [(p.id, p.name + ". " + p.description) for p in self.presets]
				),
				self._onRecommendation,
				lambda message, _key: self.recommendation.SetValue(message) if self else None,
			)
		else:
			# Translators: Shown when no recommendation is requested.
			self.recommendation.SetValue(_("Choose a format from the list"))

	def _onFormat(self, event) -> None:
		self._userMoved = True
		self._showDescription()

	def _showDescription(self) -> None:
		index = self.formats.GetSelection()
		if 0 <= index < len(self.presets):
			self.description.SetValue(self.presets[index].description)

	def _onRecommendation(self, result: dict[str, str]) -> None:
		if not self:
			return
		index = next((i for i, p in enumerate(self.presets) if p.id == result["format"]), -1)
		if index < 0:
			return
		name = self.presets[index].name
		# Translators: AI recommendation, for example "Thai official letter: starts with ที่ and เรื่อง".
		text = _("{name}: {reason}").format(name=name, reason=result["reason"]) if result["reason"] else name
		self.recommendation.SetValue(text)
		if not self._userMoved:
			self.formats.SetSelection(index)
			self._showDescription()
		# Translators: Spoken when AI recommends a format.
		ui.message(_("AI recommends {text}").format(text=text))

	def _onOk(self) -> None:
		index = self.formats.GetSelection()
		if index < 0:
			return
		self.chosen = self.presets[index]
		self.selectionOnly = self.scope.GetSelection() == 1
		self.options = ApplyOptions(
			pageSetup=self.pageSetup.GetValue() and not self.selectionOnly,
			fonts=self.fonts.GetValue(),
			removeEmpty=self.removeEmpty.GetValue(),
		)
		self.EndModal(wx.ID_OK)


class FormatPreviewDialog(wx.Dialog):
	"""Shows the role AI gave each paragraph; the user can change any of them before formatting."""

	def __init__(self, parent: wx.Window, preset: Preset, texts: list[str], roles: list[str]):
		# Translators: Title of the format preview dialog; {name} is the format.
		super().__init__(parent, title=_("Preview: {name}").format(name=preset.name))
		self.texts = texts
		self.roles = list(roles)
		names = presets.roleNames()
		self.roleIds = preset.roleIds
		self.roleLabels = [names.get(r, r) for r in self.roleIds]

		mainSizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		self.summary = helper.addItem(wx.StaticText(self, label=self._summaryText()))
		self.list = helper.addLabeledControl(
			# Translators: Label of the list of paragraphs in the format preview.
			_("&Paragraphs and their roles:"),
			wx.ListCtrl,
			style=wx.LC_REPORT | wx.LC_SINGLE_SEL,
			size=(700, 260),
		)
		# Translators: Column headers in the format preview.
		for column, (header, width) in enumerate(((_("No."), 50), (_("Role"), 200), (_("Text"), 440))):
			self.list.InsertColumn(column, header, width=width)
		for index, text in enumerate(texts):
			position = self.list.InsertItem(index, str(index + 1))
			self.list.SetItem(position, 1, names.get(self.roles[index], self.roles[index]))
			# Translators: Shown for a paragraph without text.
			self.list.SetItem(position, 2, " ".join(text.split())[:120] or _("(empty)"))
		self.role = helper.addLabeledControl(
			# Translators: Label of the choice that changes the role of the selected paragraph.
			_("&Role of the selected paragraph:"),
			wx.Choice,
			choices=self.roleLabels,
		)
		self.fullText = helper.addLabeledControl(
			# Translators: Label of the field with the full text of the selected paragraph.
			_("Text of the selected paragraph:"),
			wx.TextCtrl,
			style=wx.TE_READONLY | wx.TE_MULTILINE,
			size=(700, 70),
		)
		buttons = guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: Button that formats the document as shown in the preview.
		buttons.addButton(self, wx.ID_OK, label=_("&Format now"))
		buttons.addButton(self, wx.ID_CANCEL, label=_("Cancel"))
		helper.addItem(buttons)
		mainSizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizerAndFit(mainSizer)

		self.list.Bind(wx.EVT_LIST_ITEM_FOCUSED, lambda evt: self._showSelected())
		self.role.Bind(wx.EVT_CHOICE, self._onRole)
		if texts:
			self.list.Select(0)
			self.list.Focus(0)
		self._showSelected()
		self.list.SetFocus()
		self.CentreOnScreen()

	def _summaryText(self) -> str:
		counts: dict[str, int] = {}
		for role in self.roles:
			counts[role] = counts.get(role, 0) + 1
		names = presets.roleNames()
		parts = [f"{names.get(r, r)} {c}" for r, c in sorted(counts.items(), key=lambda x: -x[1])]
		# Translators: Summary at the top of the format preview; {count} is a number.
		return _("{count} paragraphs: {parts}").format(count=len(self.roles), parts=", ".join(parts))

	def _selected(self) -> int:
		return self.list.GetFirstSelected()

	def _showSelected(self) -> None:
		index = self._selected()
		if index < 0:
			return
		self.fullText.SetValue(self.texts[index])
		role = self.roles[index]
		self.role.SetSelection(self.roleIds.index(role) if role in self.roleIds else wx.NOT_FOUND)

	def _onRole(self, event) -> None:
		index = self._selected()
		choice = self.role.GetSelection()
		if index < 0 or choice < 0:
			return
		self.roles[index] = self.roleIds[choice]
		self.list.SetItem(index, 1, self.roleLabels[choice])
		self.summary.SetLabel(self._summaryText())

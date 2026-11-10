# Word Document Assistant for NVDA
# Configuration and the settings panel.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import addonHandler
import config
import ui
import wx
from gui import guiHelper, nvdaControls
from gui.settingsDialogs import SettingsPanel

from . import ai
from .elements import CONFIG_SECTION

addonHandler.initTranslation()

CONFIG_SPEC = {
	"playSounds": "boolean(default=True)",
	"reportPosition": "boolean(default=True)",
	"reportAuthor": "boolean(default=True)",
	"reportDate": "boolean(default=False)",
	"speakSpellingContext": "boolean(default=False)",
	"contextLength": "integer(default=20, min=5, max=100)",
	"maxItems": "integer(default=1000, min=50, max=10000)",
	# Longest time a search for items may take, so NVDA never waits a long time for Word.
	"collectSeconds": "integer(default=10, min=3, max=120)",
	"aiEnabled": "boolean(default=False)",
	"aiProvider": 'option("gemini", "anthropic", "openaiApi", "openai", "ollama", default="gemini")',
	"aiApiKey": 'string(default="")',
	"aiUrl": 'string(default="")',
	"aiModel": 'string(default="")',
	"aiLanguage": 'option("auto", "th", "en", default="auto")',
	"aiConsent": "boolean(default=False)",
	# The key of each service, so switching services does not lose a key.
	"aiKeyGemini": 'string(default="")',
	"aiKeyAnthropic": 'string(default="")',
	"aiKeyOpenaiApi": 'string(default="")',
	"aiKeyOpenai": 'string(default="")',
	# Thai word check
	"thaiCheckUnknown": "boolean(default=True)",
	"wordListUpdates": "boolean(default=False)",
	# The user has been asked once whether the weekly check may run.
	"wordListAsked": "boolean(default=False)",
	"wordListAutoDownload": "boolean(default=False)",
	"wordListLibreOffice": "boolean(default=False)",
	"misspellingListUrl": 'string(default="")',
}


def initConfig() -> None:
	config.conf.spec[CONFIG_SECTION] = CONFIG_SPEC


class WordDocAssistantPanel(SettingsPanel):
	# Translators: Title of the add-on settings panel.
	title = _("Word Document Assistant")

	def makeSettings(self, settingsSizer):
		helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		conf = config.conf[CONFIG_SECTION]
		self.playSounds = helper.addItem(
			# Translators: Check box in the settings panel.
			wx.CheckBox(self, label=_("Play &sounds for navigation mode and wrapping")),
		)
		self.playSounds.SetValue(conf["playSounds"])
		self.reportPosition = helper.addItem(
			# Translators: Check box in the settings panel.
			wx.CheckBox(self, label=_("Report item &position, for example 3 of 12")),
		)
		self.reportPosition.SetValue(conf["reportPosition"])
		self.reportAuthor = helper.addItem(
			# Translators: Check box in the settings panel.
			wx.CheckBox(self, label=_("Report the &author of comments and revisions")),
		)
		self.reportAuthor.SetValue(conf["reportAuthor"])
		self.reportDate = helper.addItem(
			# Translators: Check box in the settings panel.
			wx.CheckBox(self, label=_("Report the &date of comments and revisions")),
		)
		self.reportDate.SetValue(conf["reportDate"])
		self.speakContext = helper.addItem(
			# Translators: Check box in the settings panel.
			wx.CheckBox(self, label=_("Speak the &text around misspelled words")),
		)
		self.speakContext.SetValue(conf["speakSpellingContext"])
		self.contextLength = helper.addLabeledControl(
			# Translators: Label of a number field in the settings panel.
			_("&Characters of context on each side:"),
			nvdaControls.SelectOnFocusSpinCtrl,
			min=5,
			max=100,
			initial=conf["contextLength"],
		)
		self.maxItems = helper.addLabeledControl(
			# Translators: Label of a number field in the settings panel.
			_("&Maximum items collected for each type:"),
			nvdaControls.SelectOnFocusSpinCtrl,
			min=50,
			max=10000,
			initial=conf["maxItems"],
		)
		self.collectSeconds = helper.addLabeledControl(
			# Translators: Label of a number field in the settings panel.
			_("&Longest search time in seconds:"),
			nvdaControls.SelectOnFocusSpinCtrl,
			min=3,
			max=120,
			initial=conf["collectSeconds"],
		)
		self._makeAiSettings(helper, conf)
		self._makeWordListSettings(helper, conf)

	def _makeWordListSettings(self, helper: guiHelper.BoxSizerHelper, conf) -> None:
		# Translators: Group of Thai word check settings in the settings panel.
		groupSizer = wx.StaticBoxSizer(wx.VERTICAL, self, label=_("Thai word check (Z)"))
		box = groupSizer.GetStaticBox()
		group = guiHelper.BoxSizerHelper(box, sizer=groupSizer)
		helper.addItem(group)
		self.thaiUnknown = group.addItem(
			# Translators: Check box in the Thai word check settings.
			wx.CheckBox(
				box,
				label=_("Also report words that are not in the word list (names may be reported)"),
			),
		)
		self.thaiUnknown.SetValue(conf["thaiCheckUnknown"])
		self.wordListUpdates = group.addItem(
			# Translators: Check box in the Thai word check settings.
			wx.CheckBox(box, label=_("Check for new word lists once a week and tell me")),
		)
		self.wordListUpdates.SetValue(conf["wordListUpdates"])
		self.wordListAuto = group.addItem(
			# Translators: Check box in the Thai word check settings.
			wx.CheckBox(box, label=_("Download new word lists automatically")),
		)
		self.wordListAuto.SetValue(conf["wordListAutoDownload"])
		self.wordListLibreOffice = group.addItem(
			# Translators: Check box in the Thai word check settings.
			wx.CheckBox(
				box,
				label=_("Also use the LibreOffice Thai dictionary (NECTEC LEXiTRON, LGPL, downloaded)"),
			),
		)
		self.wordListLibreOffice.SetValue(conf["wordListLibreOffice"])
		self.misspellingUrl = group.addLabeledControl(
			# Translators: Label of the address of the organization's list of commonly misspelled words.
			_("Address of the organization's list of commonly misspelled words (optional):"),
			wx.TextCtrl,
			value=conf["misspellingListUrl"],
		)
		buttons = guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: Button that checks for new word lists now.
		updateButton = buttons.addButton(box, label=_("Update word lists &now"))
		# Translators: Button that opens the user's own word list.
		myWordsButton = buttons.addButton(box, label=_("Open my &words"))
		# Translators: Button that opens the user's own list of commonly misspelled words.
		myMisspellingsButton = buttons.addButton(box, label=_("Open my &misspellings"))
		group.addItem(buttons)
		updateButton.Bind(wx.EVT_BUTTON, self._onUpdateWordLists)
		myWordsButton.Bind(wx.EVT_BUTTON, lambda evt: self._openListFile("myWordsFile"))
		myMisspellingsButton.Bind(wx.EVT_BUTTON, lambda evt: self._openListFile("myMisspellingsFile"))

	def _saveWordLists(self) -> None:
		conf = config.conf[CONFIG_SECTION]
		conf["thaiCheckUnknown"] = self.thaiUnknown.GetValue()
		conf["wordListUpdates"] = self.wordListUpdates.GetValue()
		conf["wordListAsked"] = True
		conf["wordListAutoDownload"] = self.wordListAuto.GetValue()
		conf["wordListLibreOffice"] = self.wordListLibreOffice.GetValue()
		conf["misspellingListUrl"] = self.misspellingUrl.GetValue().strip()

	def _onUpdateWordLists(self, event) -> None:
		self._saveWordLists()
		from . import wordListUpdates

		wordListUpdates.updateNow(self)

	def _openListFile(self, name: str) -> None:
		import os

		from . import thaicheck

		lists = thaicheck.wordLists()
		lists.ensureMyFiles()
		try:
			os.startfile(getattr(lists, name))
		except OSError:
			ui.message(getattr(lists, name))

	def _makeAiSettings(self, helper: guiHelper.BoxSizerHelper, conf) -> None:
		# Translators: Group of AI settings in the settings panel.
		groupSizer = wx.StaticBoxSizer(wx.VERTICAL, self, label=_("AI suggestions"))
		box = groupSizer.GetStaticBox()
		group = guiHelper.BoxSizerHelper(box, sizer=groupSizer)
		helper.addItem(group)
		self.aiEnabled = group.addItem(
			# Translators: Check box that turns on AI suggestions.
			wx.CheckBox(box, label=_("&Use AI to suggest corrections and comment replies")),
		)
		self.aiEnabled.SetValue(conf["aiEnabled"])
		note = group.addItem(
			wx.StaticText(
				box,
				label=_(
					# Translators: Privacy note in the AI settings.
					"AI only suggests. Nothing changes in the document until you choose a suggestion. "
					"With an online service the text you check is sent to that service; "
					"free tiers may use it to improve their products. "
					"For confidential documents use Ollama on this computer.",
				),
			),
		)
		note.Wrap(520)
		self._providerIds = list(ai.PROVIDERS)
		names = ai.providerNames()
		self.aiProvider = group.addLabeledControl(
			# Translators: Label of the AI provider choice.
			_("AI &service:"),
			wx.Choice,
			choices=[names[p] for p in self._providerIds],
		)
		self.aiProvider.SetSelection(self._providerIds.index(conf["aiProvider"]))
		self.aiApiKey = group.addLabeledControl(
			# Translators: Label of the API key field.
			_("API &key:"),
			wx.TextCtrl,
			value=ai.currentKey(),
			style=wx.TE_PASSWORD,
		)
		self.aiUrl = group.addLabeledControl(
			# Translators: Label of the AI address field.
			_("Service &address (empty for the default):"),
			wx.TextCtrl,
			value=conf["aiUrl"],
		)
		self.aiModel = group.addLabeledControl(
			# Translators: Label of the AI model field.
			_("M&odel (empty for the default):"),
			wx.TextCtrl,
			value=conf["aiModel"],
		)
		self._languageIds = ["auto", "th", "en"]
		self.aiLanguage = group.addLabeledControl(
			# Translators: Label of the document language choice for AI.
			_("Document &language:"),
			wx.Choice,
			# Translators: Choices for the document language used by AI.
			choices=[_("Automatic"), _("Thai"), _("English")],
		)
		self.aiLanguage.SetSelection(self._languageIds.index(conf["aiLanguage"]))
		# Translators: Button that tests the AI connection.
		testButton = group.addItem(wx.Button(box, label=_("&Test connection")))
		testButton.Bind(wx.EVT_BUTTON, self._onTest)
		# Translators: Button that opens the guided AI key setup.
		setupButton = group.addItem(wx.Button(box, label=_("Get or &update the API key...")))
		setupButton.Bind(wx.EVT_BUTTON, self._onKeySetup)

	def _onKeySetup(self, event) -> None:
		from . import keySetup

		dialog = keySetup.KeySetupDialog(self)
		try:
			if dialog.ShowModal() == wx.ID_OK:
				conf = config.conf[CONFIG_SECTION]
				self.aiEnabled.SetValue(conf["aiEnabled"])
				self.aiProvider.SetSelection(self._providerIds.index(conf["aiProvider"]))
				self.aiApiKey.SetValue(ai.currentKey())
				self.aiUrl.SetValue(conf["aiUrl"])
				self.aiModel.SetValue(conf["aiModel"])
		finally:
			dialog.Destroy()

	def _saveAi(self) -> None:
		conf = config.conf[CONFIG_SECTION]
		provider = self._providerIds[self.aiProvider.GetSelection()]
		if provider != conf["aiProvider"]:
			# A different service needs a new agreement to send text.
			conf["aiConsent"] = False
		conf["aiEnabled"] = self.aiEnabled.GetValue()
		conf["aiProvider"] = provider
		ai.saveKey(provider, self.aiApiKey.GetValue().strip())
		conf["aiUrl"] = self.aiUrl.GetValue().strip()
		conf["aiModel"] = self.aiModel.GetValue().strip()
		conf["aiLanguage"] = self._languageIds[self.aiLanguage.GetSelection()]

	def _onTest(self, event) -> None:
		self._saveAi()
		# Translators: Reported while the AI connection is tested.
		ui.message(_("Testing..."))
		ai.runInBackground(ai.testConnection, ui.message, lambda message, needsKey: ui.message(message))

	def onSave(self):
		conf = config.conf[CONFIG_SECTION]
		conf["playSounds"] = self.playSounds.GetValue()
		conf["reportPosition"] = self.reportPosition.GetValue()
		conf["reportAuthor"] = self.reportAuthor.GetValue()
		conf["reportDate"] = self.reportDate.GetValue()
		conf["speakSpellingContext"] = self.speakContext.GetValue()
		conf["contextLength"] = self.contextLength.GetValue()
		conf["maxItems"] = self.maxItems.GetValue()
		conf["collectSeconds"] = self.collectSeconds.GetValue()
		self._saveAi()
		self._saveWordLists()
		from . import thaicheck

		thaicheck.reload()

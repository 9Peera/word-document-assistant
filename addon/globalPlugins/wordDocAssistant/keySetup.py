# Word Document Assistant for NVDA
# Setting up or updating the AI API key.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import os

import addonHandler
import api
import config
import ui
import wx
from gui import guiHelper
from logHandler import log

from . import ai
from .elements import CONFIG_SECTION

addonHandler.initTranslation()


def _clipboardText() -> str:
	try:
		return (api.getClipData() or "").strip()
	except Exception:  # noqa: BLE001
		return ""


class KeySetupDialog(wx.Dialog):
	"""Guides the user to get a free key with their own account, then saves and tests it."""

	def __init__(self, parent: wx.Window):
		# Translators: Title of the AI key setup dialog.
		super().__init__(parent, title=_("AI service and API key"))
		conf = config.conf[CONFIG_SECTION]
		self._providerIds = list(ai.PROVIDERS)
		self._lastClipboard = ""
		self._busy = False

		mainSizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		intro = helper.addItem(
			wx.StaticText(
				self,
				label=_(
					# Translators: Instructions at the top of the AI key setup dialog.
					"Choose a service, then press Get an API key. The service's page opens in your browser; "
					"sign in with your own account, create a key and copy it. "
					"When you come back to this dialog the key is found on the clipboard, tested and filled in. "
					"Google Gemini has a free tier; Anthropic Claude and OpenAI are paid and need billing set up "
					"on their website. Leave the model empty to choose one automatically, or press Find models. "
					"Ollama runs on this computer for free; no key is needed and text stays here.",
				),
			),
		)
		intro.Wrap(560)
		names = ai.providerNames()
		self.provider = helper.addLabeledControl(
			# Translators: Label of the AI provider choice.
			_("AI &service:"),
			wx.Choice,
			choices=[names[p] for p in self._providerIds],
		)
		self.provider.SetSelection(self._providerIds.index(conf["aiProvider"]))
		# Translators: Button that opens the page where the user gets a free API key or Ollama.
		self.getButton = helper.addItem(wx.Button(self, label=_("&Get a free API key (opens the browser)")))
		self.key = helper.addLabeledControl(
			# Translators: Label of the API key field.
			_("API &key:"),
			wx.TextCtrl,
			value=ai.currentKey(),
			style=wx.TE_PASSWORD,
		)
		# Translators: Button that pastes the API key from the clipboard.
		self.pasteButton = helper.addItem(wx.Button(self, label=_("&Paste key from clipboard and test")))
		self.url = helper.addLabeledControl(
			# Translators: Label of the AI address field.
			_("Service &address (empty for the default):"),
			wx.TextCtrl,
			value=conf["aiUrl"],
		)
		self.model = helper.addLabeledControl(
			# Translators: Label of the AI model field.
			_("M&odel (empty for the default):"),
			wx.ComboBox,
			value=conf["aiModel"],
			choices=[],
		)
		# Translators: Button that looks for Ollama on this computer and lists its models.
		self.ollamaButton = helper.addItem(wx.Button(self, label=_("Find &models")))
		# Translators: Button that tests the AI connection.
		self.testButton = helper.addItem(wx.Button(self, label=_("&Test connection")))
		self.status = helper.addLabeledControl(
			# Translators: Label of the read-only field that shows the last result.
			_("Status:"),
			wx.TextCtrl,
			style=wx.TE_READONLY,
			size=(560, -1),
		)
		buttons = guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: Button that saves the AI settings and turns AI suggestions on.
		buttons.addButton(self, wx.ID_OK, label=_("&Save and turn on AI"))
		buttons.addButton(self, wx.ID_CANCEL, label=_("Cancel"))
		helper.addItem(buttons)
		mainSizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizerAndFit(mainSizer)

		self.provider.Bind(wx.EVT_CHOICE, lambda evt: self._onProviderChange())
		self.getButton.Bind(wx.EVT_BUTTON, lambda evt: self._openKeyPage())
		self.pasteButton.Bind(wx.EVT_BUTTON, lambda evt: self._onPasteButton())
		self.ollamaButton.Bind(wx.EVT_BUTTON, lambda evt: self._findModels())
		self.testButton.Bind(wx.EVT_BUTTON, lambda evt: self._test())
		self.Bind(wx.EVT_BUTTON, lambda evt: self._save(), id=wx.ID_OK)
		# Coming back from the browser: look for a newly copied key.
		self.Bind(wx.EVT_ACTIVATE, self._onActivate)
		self._lastClipboard = _clipboardText()
		self._updateButtons()
		self.provider.SetFocus()
		self.CentreOnScreen()

	@property
	def providerId(self) -> str:
		return self._providerIds[self.provider.GetSelection()]

	def _values(self) -> dict:
		return {
			"aiEnabled": True,
			"aiProvider": self.providerId,
			"aiApiKey": ai.cleanKey(self.key.GetValue()),
			"aiUrl": self.url.GetValue().strip(),
			"aiModel": self.model.GetValue().strip(),
			"aiLanguage": config.conf[CONFIG_SECTION]["aiLanguage"],
		}

	def _onProviderChange(self) -> None:
		# Show the key and settings saved for this service; the model is chosen again.
		self.key.SetValue(ai.storedKey(self.providerId))
		self.model.SetItems([])
		self.model.SetValue("")
		self.url.SetValue(
			""
			if self.providerId != config.conf[CONFIG_SECTION]["aiProvider"]
			else config.conf[CONFIG_SECTION]["aiUrl"],
		)
		self._updateButtons()

	def _updateButtons(self) -> None:
		provider = self.providerId
		self.getButton.Enable(provider in ai.KEY_PAGES)
		if provider == ai.PROVIDER_OLLAMA:
			# Translators: Label of the button that opens the Ollama download page.
			self.getButton.SetLabel(_("&Get Ollama for free (opens the browser)"))
		elif provider == ai.PROVIDER_GEMINI:
			self.getButton.SetLabel(_("&Get a free API key (opens the browser)"))
		else:
			# Translators: Label of the button that opens the page to create a paid API key.
			self.getButton.SetLabel(_("&Get an API key (opens the browser)"))
		self.key.Enable(provider != ai.PROVIDER_OLLAMA)
		self.pasteButton.Enable(provider != ai.PROVIDER_OLLAMA)

	def _setStatus(self, message: str, speak: bool = True) -> None:
		if not self:
			return
		self.status.SetValue(message)
		if speak:
			ui.message(message)

	def _openKeyPage(self) -> None:
		url = ai.KEY_PAGES.get(self.providerId)
		if not url:
			return
		try:
			os.startfile(url)
		except OSError:
			log.debugWarning("Cannot open the browser", exc_info=True)
			api.copyToClip(url)
			# Translators: Reported when the browser cannot be opened; the address was copied instead.
			self._setStatus(_("Cannot open the browser. The address was copied: {url}").format(url=url))
			return
		if self.providerId != ai.PROVIDER_OLLAMA:
			# Translators: Reported after the page for creating a key was opened.
			self._setStatus(_("Browser opened. Create a key, copy it, then come back to this dialog"))

	def _onActivate(self, event: wx.ActivateEvent) -> None:
		event.Skip()
		if not event.GetActive() or self.providerId not in ai.KEY_REQUIRED:
			return
		text = ai.cleanKey(_clipboardText())
		if (
			text
			and text != self._lastClipboard
			and ai.looksLikeProviderKey(self.providerId, text)
			and text != self.key.GetValue()
		):
			self._lastClipboard = text
			wx.CallLater(400, self._paste, text, True)

	def _onPasteButton(self) -> None:
		clipboard = ai.cleanKey(_clipboardText())
		if ai.looksLikeKey(clipboard):
			self._paste(clipboard, announce=True)
			return
		typed = ai.cleanKey(self.key.GetValue())
		if typed:
			# Nothing usable on the clipboard: test the key typed in the field instead.
			self.key.SetValue(typed)
			self._test()
			return
		# Translators: Reported when neither the clipboard nor the field has an API key.
		self._setStatus(_("No API key on the clipboard or in the API key field. Copy or type the key first"))

	def _paste(self, text: str, announce: bool) -> None:
		if not self:
			return
		self.key.SetValue(ai.cleanKey(text))
		if announce:
			# Translators: Reported when an API key was taken from the clipboard.
			ui.message(_("API key found on the clipboard"))
		self._test()

	def _test(self) -> None:
		if self._busy:
			return
		self._busy = True
		self.testButton.Disable()
		# Translators: Reported while the AI connection is tested.
		self._setStatus(_("Testing..."))
		values = self._values()

		def done(message: str) -> None:
			self._finish()
			if self and not values["aiModel"] and ai.lastModels:
				# Offer the models this key can use; an empty field keeps the automatic choice.
				typed = self.model.GetValue()
				self.model.SetItems(ai.rankModels(values["aiProvider"], ai.lastModels))
				self.model.SetValue(typed)
			self._setStatus(message)

		def failed(message: str, needsKey: bool) -> None:
			self._finish()
			self._setStatus(message)

		ai.runInBackground(lambda: ai.testWith(values), done, failed)

	def _finish(self) -> None:
		self._busy = False
		if self:
			self.testButton.Enable()

	def _findModels(self) -> None:
		values = self._values()
		provider = values["aiProvider"]
		if provider in ai.KEY_REQUIRED and not values["aiApiKey"]:
			self._setStatus(_("Enter or paste an API key first"))
			return
		# Translators: Reported while the list of models is fetched.
		self._setStatus(_("Looking for models..."))

		def done(models: list[str]) -> None:
			if not self:
				return
			ranked = ai.rankModels(provider, models)
			if not ranked:
				if provider == ai.PROVIDER_OLLAMA:
					# Translators: Reported when Ollama runs but has no models installed.
					self._setStatus(
						_(
							"Ollama is running but has no models. Install one, for example: ollama pull qwen2.5:7b",
						),
					)
				else:
					# Translators: Reported when the service lists no text models for the key.
					self._setStatus(_("The service did not list any text model for this key"))
				return
			self.model.SetItems(ranked)
			current = self.model.GetValue().strip()
			self.model.SetValue(current if current in ranked else ranked[0])
			# Translators: Reported when models were found; {count} is a number, {model} a model name.
			self._setStatus(
				_("{count} models found, {model} selected. Use the Model box to choose another").format(
					count=len(ranked),
					model=self.model.GetValue(),
				),
			)

		ai.runInBackground(
			lambda: ai.listModels(provider, values["aiUrl"], values["aiApiKey"]),
			done,
			lambda message, _key: self._setStatus(message),
		)

	def _save(self) -> None:
		values = self._values()
		if values["aiProvider"] in ai.KEY_REQUIRED and not values["aiApiKey"]:
			# Translators: Reported when saving without an API key.
			self._setStatus(_("Enter or paste an API key first"))
			self.key.SetFocus()
			return
		conf = config.conf[CONFIG_SECTION]
		if values["aiProvider"] != conf["aiProvider"]:
			conf["aiConsent"] = False
		for name, value in values.items():
			if name != "aiApiKey":
				conf[name] = value
		ai.saveKey(values["aiProvider"], values["aiApiKey"])
		# Translators: Reported after the AI settings were saved.
		ui.message(_("AI settings saved"))
		self.EndModal(wx.ID_OK)


def show() -> None:
	"""Open the key setup dialog from anywhere (call through wx.CallAfter)."""
	from . import dialogs

	dialogs.runModal(KeySetupDialog)

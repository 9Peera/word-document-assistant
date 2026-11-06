# Word Document Assistant for NVDA
# Checking for, and downloading, newer Thai word lists.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import threading

import addonHandler
import config
import globalVars
import ui
import wx
from logHandler import log

from . import thaicheck, thaidict
from .elements import CONFIG_SECTION

addonHandler.initTranslation()

_busy = threading.Lock()
_toldThisSession = False


def _background(work, done, failed) -> None:
	def target():
		try:
			result = work()
		except Exception as error:  # noqa: BLE001
			log.debugWarning("Word list update failed", exc_info=True)
			wx.CallAfter(failed, error)
		else:
			wx.CallAfter(done, result)

	threading.Thread(target=target, name="wordDocAssistant.wordLists", daemon=True).start()


def _describe(updates: list[thaidict.WordSource]) -> str:
	lines = []
	for source in updates:
		# Translators: One word list in the update question; {license} is its licence, {homepage} its web page.
		lines.append(
			_("{name} (licence {license}, {homepage})").format(
				name=source.name, license=source.license or "-", homepage=source.homepage
			)
		)
	return "\n".join(lines)


def _failedMessage(error: Exception) -> str:
	if isinstance(error, ValueError):
		# Translators: Reported when a downloaded file is not a Thai word list.
		return _("The downloaded file is not a Thai word list, so it was not used")
	# Translators: Reported when the word lists cannot be checked or downloaded.
	return _("Cannot reach the word list website. Check the internet connection and try again")


def download(updates: list[thaidict.WordSource], announce: bool = True) -> None:
	lists = thaicheck.wordLists()

	def work() -> list[tuple[str, int]]:
		results = []
		for source in updates:
			results.append((source.name, lists.download(source)))
		lists.load()
		return results

	def done(results: list[tuple[str, int]]) -> None:
		_busy.release()
		if announce:
			parts = ", ".join(f"{name} {count}" for name, count in results)
			# Translators: Reported after word lists were updated; {lists} names them with their number of entries.
			ui.message(_("Thai word lists updated: {lists}").format(lists=parts))

	def failed(error: Exception) -> None:
		_busy.release()
		ui.message(_failedMessage(error))

	if not _busy.acquire(blocking=False):
		return
	# Translators: Reported while word lists are downloaded.
	ui.message(_("Downloading Thai word lists..."))
	_background(work, done, failed)


def updateNow(parent: wx.Window | None = None) -> None:
	"""Check now, then ask before downloading."""
	if globalVars.appArgs.secure:
		return
	if not _busy.acquire(blocking=False):
		# Translators: Reported when an update is already running.
		ui.message(_("The word lists are already being updated"))
		return
	lists = thaicheck.wordLists()
	# Translators: Reported while checking for newer word lists.
	ui.message(_("Checking for new Thai word lists..."))

	def done(updates: list[thaidict.WordSource]) -> None:
		_busy.release()
		if not updates:
			# Translators: Reported when the word lists are up to date.
			ui.message(_("The Thai word lists are up to date"))
			return
		from . import dialogs

		def make(window):
			return wx.MessageDialog(
				window,
				# Translators: Question before word lists are downloaded; {lists} names them.
				_("New Thai word lists are available:\n{lists}\n\nDownload them now?").format(
					lists=_describe(updates)
				),
				# Translators: Title of the word list update question.
				_("Thai word lists"),
				wx.YES_NO | wx.YES_DEFAULT | wx.ICON_QUESTION,
			)

		result, _value = dialogs.runModal(make)
		if result == wx.ID_YES:
			download(updates)

	def failed(error: Exception) -> None:
		_busy.release()
		ui.message(_failedMessage(error))

	_background(lists.checkForUpdates, done, failed)


def scheduledCheck() -> None:
	"""Weekly background check, started a while after NVDA starts."""
	settings = config.conf[CONFIG_SECTION]
	if globalVars.appArgs.secure or not settings["wordListUpdates"]:
		return
	lists = thaicheck.wordLists()
	if not lists.dueForCheck() or not _busy.acquire(blocking=False):
		return

	def done(updates: list[thaidict.WordSource]) -> None:
		_busy.release()
		if not updates:
			return
		if settings["wordListAutoDownload"]:
			download(updates)
		else:
			announcePending(updates)

	def failed(error: Exception) -> None:
		_busy.release()

	_background(lists.checkForUpdates, done, failed)


def askOnce() -> None:
	"""The first time the Thai word check is used, ask whether the weekly check may use the internet."""
	settings = config.conf[CONFIG_SECTION]
	if globalVars.appArgs.secure or settings["wordListAsked"] or settings["wordListUpdates"]:
		return
	settings["wordListAsked"] = True
	from . import dialogs

	def make(window):
		return wx.MessageDialog(
			window,
			# Translators: Question asked once, the first time the Thai word check is used.
			_(
				"The Thai word check works offline with the word list included in the add-on.\n\n"
				"Should the add-on look for newer Thai word lists on GitHub once a week and tell you when there "
				"is one? It only asks whether a newer file exists and never sends your documents. "
				"You can change this later in the settings."
			),
			# Translators: Title of the question about the weekly word list check.
			_("New Thai word lists"),
			wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION,
		)

	result, _value = dialogs.runModal(make)
	if result == wx.ID_YES:
		settings["wordListUpdates"] = True
		# Translators: Reported after the user allowed the weekly word list check.
		ui.message(_("The add-on will look for new Thai word lists once a week"))
		scheduledCheck()


def announcePending(updates: list[thaidict.WordSource] | None = None) -> None:
	"""Tell the user once per session that newer word lists can be downloaded."""
	global _toldThisSession
	if _toldThisSession:
		return
	updates = updates if updates is not None else thaicheck.wordLists().pendingUpdates()
	if not updates:
		return
	_toldThisSession = True
	ui.message(
		# Translators: Told when newer word lists are available; {lists} names them.
		_(
			"New Thai word lists are available: {lists}. To update, use the NVDA menu, Tools, "
			"Update Thai word lists, or the actions of the Thai word check (Z)."
		).format(lists=", ".join(s.name for s in updates))
	)

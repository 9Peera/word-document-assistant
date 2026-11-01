# Word Document Assistant for NVDA
# Earcons.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

import os

import config
import nvwave
from logHandler import log

from .elements import CONFIG_SECTION

_SOUNDS_DIR = os.path.join(os.path.dirname(__file__), "sounds")

ENTER = "enter"
EXIT = "exit"
WRAP = "wrap"
ERROR = "error"


def play(name: str) -> None:
	if not config.conf[CONFIG_SECTION]["playSounds"]:
		return
	path = os.path.join(_SOUNDS_DIR, f"{name}.wav")
	try:
		nvwave.playWaveFile(path, asynchronous=True)
	except Exception:
		log.debugWarning(f"Cannot play {path}", exc_info=True)

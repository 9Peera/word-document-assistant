# Word Document Assistant for NVDA
# Protects API keys with Windows DPAPI, so the NVDA configuration file does not hold them in plain text.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

"""API keys are encrypted for the current Windows user with CryptProtectData.

A protected key only works for the same Windows user on the same computer.
If DPAPI is not available the key is stored as before, so the add-on keeps working.
"""

from __future__ import annotations

import base64
import ctypes
import sys

from logHandler import log

PREFIX = "dpapi:"
_ENTROPY = b"wordDocAssistant.apiKey"
CRYPTPROTECT_UI_FORBIDDEN = 0x1


class _Blob(ctypes.Structure):
	_fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _blob(data: bytes) -> tuple[_Blob, ctypes.Array]:
	buffer = ctypes.create_string_buffer(data, len(data))
	return _Blob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char))), buffer


def _crypt(data: bytes, protect: bool) -> bytes | None:
	if sys.platform != "win32":
		return None
	try:
		from ctypes import wintypes

		crypt32 = ctypes.windll.crypt32
		kernel32 = ctypes.windll.kernel32
		function = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
		function.argtypes = [
			ctypes.POINTER(_Blob),
			ctypes.c_void_p,
			ctypes.POINTER(_Blob),
			ctypes.c_void_p,
			ctypes.c_void_p,
			wintypes.DWORD,
			ctypes.POINTER(_Blob),
		]
		function.restype = wintypes.BOOL
		kernel32.LocalFree.argtypes = [ctypes.c_void_p]
		kernel32.LocalFree.restype = ctypes.c_void_p
		inBlob, _inBuffer = _blob(data)
		entropy, _entropyBuffer = _blob(_ENTROPY)
		outBlob = _Blob()
		if not function(
			ctypes.byref(inBlob),
			None,
			ctypes.byref(entropy),
			None,
			None,
			CRYPTPROTECT_UI_FORBIDDEN,
			ctypes.byref(outBlob),
		):
			return None
		try:
			return ctypes.string_at(outBlob.pbData, outBlob.cbData)
		finally:
			kernel32.LocalFree(ctypes.cast(outBlob.pbData, ctypes.c_void_p))
	except Exception:  # noqa: BLE001
		log.debugWarning("DPAPI is not available", exc_info=True)
		return None


def protect(key: str) -> str:
	"""Value to store in the configuration for this key."""
	if not key or key.startswith(PREFIX):
		return key
	encrypted = _crypt(key.encode("utf-8"), protect=True)
	if encrypted is None:
		return key
	return PREFIX + base64.b64encode(encrypted).decode("ascii")


def unprotect(value: str) -> str:
	"""The key from a stored value. Empty when it was protected by another user or computer."""
	if not value or not value.startswith(PREFIX):
		return value or ""
	try:
		encrypted = base64.b64decode(value[len(PREFIX) :])
	except ValueError:
		return ""
	decrypted = _crypt(encrypted, protect=False)
	if decrypted is None:
		log.debugWarning("Cannot decrypt the stored API key; it was saved by another user or computer")
		return ""
	return decrypted.decode("utf-8", "replace")


def isProtected(value: str) -> bool:
	return bool(value) and value.startswith(PREFIX)

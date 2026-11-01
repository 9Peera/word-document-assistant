"""Minimal stand-ins for the NVDA modules the add-on imports, so its logic can be tested without NVDA."""

import builtins
import sys
import types

builtins._ = lambda text: text  # type: ignore[attr-defined]
builtins.pgettext = lambda context, text: text  # type: ignore[attr-defined]

spoken: list[str] = []
clipboard: list[str] = []


def _module(name: str, **attributes):
	module = types.ModuleType(name)
	module.__dict__.update(attributes)
	sys.modules[name] = module
	return module


class Obj:
	pass


focus = Obj()
focus.windowClassName = "_WwG"
focus.windowHandle = 1
focus.appModule = Obj()
focus.appModule.appName = "winword"
focus.treeInterceptor = None


class Conf(dict):
	spec: dict = {}


conf = Conf()

_module("addonHandler", initTranslation=lambda: None)
_module(
	"api",
	getFocusObject=lambda: focus,
	getForegroundObject=lambda: focus,
	copyToClip=lambda text, notify=False: clipboard.append(text) or True,
	getClipData=lambda: clipboard[-1] if clipboard else "",
)
_module("config", conf=conf)
_module("core", callLater=lambda ms, func, *a, **k: func(*a, **k))


class _GlobalPlugin:
	def __init__(self):
		pass

	def getScript(self, gesture):
		if gesture.vkCode == 0x77 and sorted(gesture.modifierNames) == ["NVDA", "shift"]:
			return self.script_toggleNavigationMode
		return None

	def terminate(self):
		pass


_module("globalPluginHandler", GlobalPlugin=_GlobalPlugin)
_module("globalVars", appArgs=types.SimpleNamespace(secure=True))
_module("speech", cancelSpeech=lambda: None)
_module(
	"ui", message=lambda text: spoken.append(text), browseableMessage=lambda *a, **k: spoken.append("HELP")
)
wx = _module("wx", Dialog=object, Window=object, ID_HIGHEST=5999, CallAfter=lambda func, *a: func(*a))
gui = _module("gui")
gui.guiHelper = _module("gui.guiHelper")
gui.nvdaControls = _module("gui.nvdaControls")
gui.settingsDialogs = _module("gui.settingsDialogs", SettingsPanel=object)
_module("windowUtils")


class KeyboardInputGesture:
	def __init__(self, vk, modifiers=(), extended=False):
		self.vkCode = vk
		self.modifierNames = list(modifiers)
		self.isExtended = extended
		self.isModifier = False


_module("keyboardHandler", KeyboardInputGesture=KeyboardInputGesture)
_noop = lambda *a, **k: None  # noqa: E731
_module("logHandler", log=types.SimpleNamespace(debugWarning=_noop, info=_noop, error=_noop))
_module("scriptHandler", script=lambda **kw: lambda func: func)


class COMError(Exception):
	pass


_module("comtypes", COMError=COMError)
_module("comtypes.client")
_module("comtypes.automation", IDispatch=None)
_module("oleacc")
_module(
	"winUser",
	getClassName=lambda hwnd: {1: "_WwG", 2: "child"}.get(hwnd, ""),
	getAncestor=lambda hwnd, flag: 1 if hwnd == 2 else 0,
)
_module("nvwave", playWaveFile=_noop)

DEFAULT_CONFIG = {
	"playSounds": True,
	"reportPosition": True,
	"reportAuthor": True,
	"reportDate": False,
	"speakSpellingContext": False,
	"contextLength": 20,
	"maxItems": 1000,
	"aiEnabled": True,
	"aiProvider": "gemini",
	"aiApiKey": "",
	"aiUrl": "",
	"aiModel": "",
	"aiLanguage": "auto",
	"aiConsent": True,
	"aiKeyGemini": "",
	"aiKeyAnthropic": "",
	"aiKeyOpenaiApi": "",
	"aiKeyOpenai": "",
}


def resetConfig(**values):
	conf["wordDocAssistant"] = {**DEFAULT_CONFIG, **values}
	spoken.clear()
	clipboard.clear()

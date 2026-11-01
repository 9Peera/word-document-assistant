import nvdaStubs
from fakeword import Context, Doc
from nvdaStubs import KeyboardInputGesture as Key

import wordDocAssistant as plugin
from wordDocAssistant import registry

F8 = 0x77


def makePlugin(monkeypatch, doc):
	monkeypatch.setattr(plugin.WordContext, "fromFocus", classmethod(lambda cls: Context(doc)))
	instance = plugin.GlobalPlugin()
	return instance


def press(instance, vk, modifiers=(), extended=False):
	gesture = Key(vk, modifiers, extended)
	script = instance.getScript(gesture)
	if script is not None:
		script(gesture)
	return script


def spellingDoc():
	doc = Doc(["helo world. thsi is bad. end"])
	doc.spelling = [doc.find("helo"), doc.find("thsi")]
	return doc


def test_every_letter_belongs_to_one_type_and_commands_do_not_clash():
	letters = {chr(k.vk) for k in registry.NAVIGATION_TYPES}
	assert letters == set("SGCFERUOMKXIBDLAZ")
	commands = {
		chr(c)
		for c in (
			plugin.CMD_TRACK,
			plugin.CMD_NOTE,
			plugin.CMD_POSITION,
			plugin.CMD_SUMMARY,
			plugin.CMD_HELP,
			plugin.CMD_AI,
			plugin.CMD_FORMAT,
			plugin.CMD_AI_TOOLS,
			plugin.CMD_THAI_TOOLS,
		)
	}
	assert not letters & commands


def test_mode_moves_with_letters_arrows_and_leaves_at_the_item(monkeypatch):
	doc = spellingDoc()
	instance = makePlugin(monkeypatch, doc)
	press(instance, F8, ["NVDA", "shift"])
	assert instance.active and nvdaStubs.spoken[-1] == "Navigation mode, Spelling errors"
	press(instance, ord("S"))
	assert nvdaStubs.spoken[-1] == "helo, 1 of 2"
	press(instance, 0x28, extended=True)
	assert nvdaStubs.spoken[-1] == "thsi, 2 of 2"
	press(instance, ord("S"), ["shift"])
	assert nvdaStubs.spoken[-1] == "helo, 1 of 2"
	press(instance, 0x0D)
	assert not instance.active
	assert (doc.selection.Start, doc.selection.End) == doc.find("helo")


def test_backspace_returns_to_the_start(monkeypatch):
	doc = spellingDoc()
	doc.selection.Start = doc.selection.End = 25
	instance = makePlugin(monkeypatch, doc)
	press(instance, F8, ["NVDA", "shift"])
	press(instance, ord("S"))
	press(instance, 0x08)
	assert not instance.active and doc.selection.Start == 25


def test_numpad_arrow_leaves_the_mode(monkeypatch):
	instance = makePlugin(monkeypatch, spellingDoc())
	press(instance, F8, ["NVDA", "shift"])
	press(instance, 0x28, extended=False)
	assert not instance.active


def test_nvda_f8_alone_does_nothing_and_outside_word_the_key_is_free(monkeypatch):
	instance = makePlugin(monkeypatch, spellingDoc())
	press(instance, F8, ["NVDA"])
	assert not instance.active
	nvdaStubs.focus.appModule.appName = "explorer"
	try:
		assert instance.getScript(Key(F8, ["NVDA", "shift"])) is None
	finally:
		nvdaStubs.focus.appModule.appName = "winword"


def test_change_type_with_right_arrow_and_tool_letters(monkeypatch):
	instance = makePlugin(monkeypatch, spellingDoc())
	calls = []
	monkeypatch.setattr(instance, "_aiToolsMenu", lambda: calls.append("V"))
	monkeypatch.setattr(instance, "_thaiToolsMenu", lambda: calls.append("Y"))
	monkeypatch.setattr(instance, "_aiFormat", lambda: calls.append("J"))
	press(instance, F8, ["NVDA", "shift"])
	press(instance, 0x27, extended=True)
	assert nvdaStubs.spoken[-1].startswith("Grammar issues")
	for letter in "VYJ":
		press(instance, ord(letter))
	assert calls == ["V", "Y", "J"] and instance.active


def test_copy_current_item(monkeypatch):
	instance = makePlugin(monkeypatch, spellingDoc())
	press(instance, F8, ["NVDA", "shift"])
	press(instance, ord("S"))
	press(instance, ord("C"), ["control"])
	assert nvdaStubs.clipboard[-1] == "helo"

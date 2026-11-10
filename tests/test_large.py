"""Long documents: work is done in parts, with Word's own search, and stops at the time limit."""

import nvdaStubs
from fakeword import ENGLISH, THAI, Coll, Context, Doc
from nvdaStubs import KeyboardInputGesture as Key

import wordDocAssistant as plugin
from wordDocAssistant import a11y, elements, wordCom

F8 = 0x77
RIGHT = 0x27
DOWN = 0x28


def press(instance, vk, modifiers=(), extended=False):
	gesture = Key(vk, modifiers, extended)
	script = instance.getScript(gesture)
	if script is not None:
		script(gesture)


def longDoc(paragraphs=300):
	texts = [f"line {n} helo" if n % 50 == 0 else f"line {n} text" for n in range(paragraphs)]
	doc = Doc(texts)
	start = 0
	while True:
		start = doc.text.find("helo", start)
		if start < 0:
			break
		doc.spelling.append((start, start + 4))
		start += 4
	return doc


class CountingColl(Coll):
	"""A collection that records how it was read."""

	def __init__(self, items, iterable=True):
		super().__init__(items)
		self.itemCalls = 0
		self.iterable = iterable

	def Item(self, index):
		self.itemCalls += 1
		return super().Item(index)

	def __iter__(self):
		if not self.iterable:
			raise TypeError("no enumerator")
		return super().__iter__()


def test_iterate_uses_the_enumerator_and_falls_back_to_item():
	enumerable = CountingColl(["a", "b", "c"])
	assert list(wordCom.iterate(enumerable)) == ["a", "b", "c"]
	assert enumerable.itemCalls == 0
	plain = CountingColl(["a", "b"], iterable=False)
	assert list(wordCom.iterate(plain)) == ["a", "b"]
	assert plain.itemCalls == 2
	assert list(wordCom.iterate(CountingColl([]))) == []


def test_budget_stops_iteration():
	budget = wordCom.Budget(0)
	assert list(wordCom.iterate(CountingColl(["a", "b", "c"]), budget)) == ["a"]
	assert budget.expired


def test_spelling_is_read_in_parts_that_end_at_paragraph_ends(monkeypatch):
	monkeypatch.setattr(wordCom, "CHUNK_SIZE", 500)
	doc = longDoc()
	items, total = elements.SpellingErrors().collect(Context(doc), 1000)
	assert total == len(doc.spelling) == 6
	assert [item.start for item in items] == [start for start, _end in doc.spelling]
	assert len(doc.proofingCalls) > 5
	ends = {p.end for p in doc.paragraphs()}
	assert all(end in ends for _start, end in doc.proofingCalls)


def test_time_limit_gives_what_was_found_and_says_so(monkeypatch):
	monkeypatch.setattr(wordCom, "CHUNK_SIZE", 500)
	doc = longDoc()
	monkeypatch.setattr(plugin.WordContext, "fromFocus", classmethod(lambda cls: Context(doc)))
	nvdaStubs.resetConfig(collectSeconds=0)
	instance = plugin.GlobalPlugin()
	press(instance, F8, ["NVDA", "shift"])
	press(instance, ord("S"))
	assert instance._totals["spelling"] < 6
	assert any("stopped after 0 seconds" in text for text in nvdaStubs.spoken)


def test_arrows_do_not_search_slow_types(monkeypatch):
	doc = longDoc()
	monkeypatch.setattr(plugin.WordContext, "fromFocus", classmethod(lambda cls: Context(doc)))
	instance = plugin.GlobalPlugin()
	press(instance, F8, ["NVDA", "shift"])
	press(instance, RIGHT, extended=True)
	assert nvdaStubs.spoken[-1] == "Grammar issues"
	assert doc.proofingCalls == []
	press(instance, DOWN, extended=True)
	assert doc.proofingCalls


def test_language_problems_are_found_with_search_not_paragraph_by_paragraph():
	doc = Doc(["English only", "ไทยถูกต้อง", "ไทยผิด", "mixed ไทย", "Latin marked Thai"])
	start, end = doc.find("ไทยผิด")
	doc.lang[start:end] = [ENGLISH] * (end - start)
	start, end = doc.find("Latin marked Thai")
	doc.lang[start:end] = [THAI] * (end - start)
	start, end = doc.find("mixed")
	doc.lang[start:end] = [THAI] * (end - start)
	problems = a11y.findLanguageProblems(Context(doc), 100)
	assert [(p.kind, p.text.strip()) for p in problems] == [
		("thai", "ไทยผิด"),
		("mixed", "mixed ไทย"),
		("latin", "Latin marked Thai"),
	]


def test_empty_paragraph_runs_are_found_with_search():
	doc = Doc(["one", "", "", "", "two", "", "three", "", ""])
	runs = list(a11y._emptyParagraphRuns(Context(doc), wordCom.Budget(60)))
	assert [len(run) for run in runs] == [3, 2]

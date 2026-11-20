"""Support for many languages: language check, digits, word lists, AI language and formats."""

import nvdaStubs
from fakeword import ENGLISH, THAI, Context, Doc
from wordDocAssistant import a11y, ai, languages, presets, thaidict, tools

RUSSIAN = 1049


def test_scripts_and_languages():
	assert languages.scriptsIn("Hello สวัสดี Привет") == ["latin", "thai", "cyrillic"]
	assert languages.scriptOf("Привет,") == "cyrillic"
	assert languages.scriptOf("123 ...") is None
	assert languages.isWrongLanguage("thai", ENGLISH)
	assert languages.isWrongLanguage("latin", THAI)
	assert languages.isWrongLanguage("cyrillic", ENGLISH)
	assert not languages.isWrongLanguage("cyrillic", RUSSIAN)
	assert not languages.isWrongLanguage("latin", 1036)  # French is fine for Latin-script text
	assert not languages.isWrongLanguage("latin", 12345)  # unknown languages are never called wrong
	assert not languages.isWrongLanguage("thai", languages.LANG_NO_PROOFING)
	assert languages.parseLanguageList("uk, th,xx,uk") == [1058, THAI]
	assert languages.preferredLanguage("cyrillic", [1058], ENGLISH) == 1058
	assert languages.preferredLanguage("cyrillic", [], ENGLISH) == RUSSIAN
	assert languages.preferredLanguage("latin", [THAI], 1036) == 1036
	assert languages.languageName(RUSSIAN) == "Russian"
	assert languages.languageName(4242) == "language 4242"


def russianDoc():
	doc = Doc(["English text", "Русский текст", "Ещё текст по-русски", "ไทย"])
	start, end = doc.find("Ещё текст по-русски")
	doc.lang[start:end] = [RUSSIAN] * (end - start)
	return doc


def test_text_of_any_script_with_a_wrong_language_is_found():
	doc = russianDoc()
	problems = a11y.findLanguageProblems(Context(doc), 100)
	assert [(p.kind, p.text.strip()) for p in problems] == [("cyrillic", "Русский текст")]
	assert a11y.LanguageProblems._problemText(problems[0]) == "Cyrillic text marked with a wrong language"


def test_fix_gives_each_word_the_language_of_its_script():
	doc = Doc(["Русский text ไทย"])
	start, end = doc.find("ไทย")
	doc.lang[start:end] = [ENGLISH] * (end - start)
	nvdaStubs.resetConfig(documentLanguages="uk")
	ctx = Context(doc)
	problems = a11y.findLanguageProblems(ctx, 100)
	assert [p.kind for p in problems] == ["mixed"]
	changed = a11y.fixParagraphLanguage(ctx, problems[0].paragraph)
	assert changed == 2
	assert doc.lang[doc.find("Русский")[0]] == 1058  # the user's Cyrillic language
	assert doc.lang[doc.find("ไทย")[0]] == THAI
	assert doc.lang[doc.find("text")[0]] == ENGLISH


def test_paragraph_actions_offer_the_language_of_each_script():
	doc = Doc(["Русский text"])
	ctx = Context(doc)
	kind = a11y.LanguageProblems()
	items, _total = kind.collect(ctx, 100)
	names = [action.label for action in kind.actions(ctx, items[0], None, None)]
	assert "Mark whole paragraph as English (United States)" in names
	assert "Mark whole paragraph as Russian" in names


def test_digits_of_many_scripts():
	assert languages.countDigitSystems("12 ๓ ٤٥ ۶") == {"western": 2, "thai": 1, "arabicIndic": 2, "persian": 1}
	doc = Doc(["Price 2026, ٣٤"])
	ctx = Context(doc)
	assert tools.convertDigits(ctx, doc.Content, "western", "arabicIndic") == 4
	assert "٢٠٢٦" in doc.text
	assert tools.convertDigits(ctx, doc.Content, "arabicIndic", "western") == 6
	assert "2026, 34" in doc.text


def test_misspelling_lists_work_for_words_with_spaces():
	dictionary = thaidict.Dictionary()
	dictionary.load([], thaidict.BUNDLED_MISSPELLING_FILES)
	text = "I recieve a sandwich. Recieve it, wich is weird; definately."
	found = [(f.word, f.suggestions[0]) for f in dictionary.check(text, unknown=False)]
	assert found == [
		("recieve", "receive"),
		("Recieve", "Receive"),
		("wich", "which"),
		("definately", "definitely"),
	]
	# Thai lists still work alongside.
	assert [f.word for f in dictionary.check("ขออนุญาติ", unknown=False)] == ["อนุญาติ"]


def test_ai_answers_in_the_chosen_language(config):
	config.update(aiLanguage="fr")
	assert ai._languageHint() == "Answer in French."
	config.update(aiLanguage="auto")
	assert ai._languageHint() == "Answer in the same language as the text."
	asked = []
	ai_answer = "Bonjour"
	original = ai._textAnswer
	try:
		ai._textAnswer = lambda system, text: asked.append(system) or ai_answer
		assert ai.translate("Hello", "ja") == "Bonjour"
	finally:
		ai._textAnswer = original
	assert "into natural, correct Japanese" in asked[0]


def test_style_guide_formats_and_documents_without_thai():
	doc = Doc(["My paper", "Method", "Some text", "Details"])
	ctx = Context(doc)
	preset = presets.presetById("apa")
	assert preset is not None and preset.paper is None and preset.group == "general"
	paragraphs, texts = presets.collectParagraphs(doc.Content, 100)
	normal = doc.Styles.Item(presets.STYLE_NORMAL)
	normal.Font.NameBi = "Arial"  # a complex-script font chosen for Arabic, for example
	presets.apply(
		ctx,
		preset,
		paragraphs,
		["title", "heading1", "body", "heading3"],
		presets.ApplyOptions(),
		True,
		hasThai=False,
	)
	assert normal.Font.NameBi == "Arial"
	assert normal.Font.Name in ("Times New Roman", "Calibri") and normal.Font.Size == 12
	assert doc.Styles.Item(presets.STYLE_HEADING3).Font.Italic is True
	assert round(doc.paraProps[2]["FirstLineIndent"] / presets.CM, 2) == 1.27
	thaiGroups = {p.id for p in presets.presets() if p.group == "thai"}
	assert thaiGroups == {"officialLetter", "memo", "announcement", "minutes"}

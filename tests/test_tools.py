from fakeword import Context, Doc
from wordDocAssistant import tools


def test_digit_conversion_both_ways():
	doc = Doc(["ปี 2569 วันที่ 13"])
	ctx = Context(doc)
	assert tools.convertDigits(ctx, doc.Content, toThai=True) == 6
	assert "๒๕๖๙" in doc.text and "๑๓" in doc.text
	assert tools.convertDigits(ctx, doc.Content, toThai=False) == 6
	assert "2569" in doc.text
	assert ctx.application.undo == ["Change digits", "Change digits"]


def test_era_replacements():
	assert tools.eraReplacements("ค.ศ. 2026 และ คศ 1999", toBuddhist=True) == [
		("ค.ศ. 2026", "พ.ศ. 2569"),
		("คศ 1999", "พศ 2542"),
	]
	assert tools.eraReplacements("พ.ศ. ๒๕๖๙", toBuddhist=False) == [("พ.ศ. ๒๕๖๙", "ค.ศ. ๒๐๒๖")]
	assert tools.eraReplacements("พ.ศ. 2569", toBuddhist=True) == []


def test_convert_era_in_document():
	doc = Doc(["เมื่อ ค.ศ. 2026 และ ค.ศ. 2026 อีกครั้ง"])
	ctx = Context(doc)
	assert tools.convertEra(ctx, doc.Content, toBuddhist=True) == 1
	assert doc.text.count("พ.ศ. 2569") == 2


def test_replace_keeps_the_paragraph_mark_and_insert_after_adds_a_paragraph():
	doc = Doc(["old text", "next"])
	ctx = Context(doc)
	paragraph = doc.paragraphs()[0].Range
	tools.replaceText(ctx, paragraph, "new text", after=False)
	assert doc.text == "new text\rnext\r"
	tools.replaceText(ctx, doc.paragraphs()[0].Range, "translation\nline 2", after=True)
	assert doc.text == "new text\rtranslation\rline 2\rnext\r"


def test_count_digits():
	assert tools.countDigits("12 ๓๔๕") == (3, 2)

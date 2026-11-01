import base64

from fakeword import ENGLISH, THAI, Context, Doc, Host, InlineShape, Link, Table
from wordDocAssistant import a11y


def makeDoc():
	doc = Doc(
		[
			"Report",  # 0 heading 1
			"Intro text",  # 1
			"Details",  # 2 heading 3 (skips level 2)
			"",  # 3 empty
			"",  # 4 empty
			"ข้อความภาษาไทย",  # 5 Thai, will be marked English
			"see here for more",  # 6 link "here"
			"X",  # 7 picture placeholder
		]
	)
	doc.setLevel(0, 1)
	doc.setLevel(2, 3)
	start, end = doc.find("ข้อความภาษาไทย")
	doc.lang[start:end] = [ENGLISH] * (end - start)
	linkStart, linkEnd = doc.find("here")
	doc.links.append(Link(doc, linkStart, linkEnd, "https://example.com"))
	doc.inlineShapes.append(InlineShape(doc, doc.find("X")[0]))
	tableStart = doc.find("Intro text")[0]
	doc.tables.append(Table(doc, tableStart, tableStart + 5, rows=3, columns=2))
	return doc


def kinds(issues):
	return sorted(issue.kind for issue in issues)


def test_checker_finds_every_kind_of_problem():
	ctx = Context(makeDoc())
	issues = a11y.checkDocument(ctx, 100)
	assert kinds(issues) == sorted(
		["altText", "tableHeader", "headingSkip", "emptyParagraphs", "linkText", "title", "language"]
	)
	skip = next(i for i in issues if i.kind == "headingSkip")
	assert skip.extra == 2


def test_fixes_change_the_document():
	doc = makeDoc()
	ctx = Context(doc)
	checker = a11y.AccessibilityIssues()
	host = Host(prompts=["Annual report", "A chart of sales"])
	items, _total = checker.collect(ctx, 100)
	byKind = {item.obj.kind: item for item in items}
	for kind in ("tableHeader", "headingSkip", "emptyParagraphs", "title"):
		checker.actions(ctx, byKind[kind], None, host)[0].run()
	checker.actions(ctx, byKind["altText"], None, host)[0].run()
	assert doc.tables[0].Rows.Item(1).HeadingFormat is True
	assert doc.paraProps[2]["level"] == 2
	assert "\r\r\r" not in doc.text
	assert doc.properties["Title"] == "Annual report"
	assert doc.inlineShapes[0].AlternativeText == "A chart of sales"
	assert host.changes == 5


def test_language_problems_and_fix():
	doc = makeDoc()
	ctx = Context(doc)
	kind = a11y.LanguageProblems()
	items, _total = kind.collect(ctx, 100)
	assert len(items) == 1 and items[0].obj.kind == "thai"
	host = Host()
	kind.actions(ctx, items[0], None, host)[0].run()
	start, end = doc.find("ข้อความภาษาไทย")
	assert set(doc.lang[start:end]) == {THAI}
	assert kind.collect(ctx, 100)[0] == []


def test_mixed_paragraph_with_one_language_is_reported():
	doc = Doc(["ทดสอบ Word"])
	doc.lang = [THAI] * len(doc.text)
	ctx = Context(doc)
	items, _total = a11y.LanguageProblems().collect(ctx, 10)
	assert items[0].obj.kind == "mixed"
	a11y.fixParagraphLanguage(ctx, items[0].obj.paragraph)
	start, end = doc.find("Word")
	assert set(doc.lang[start:end]) == {ENGLISH}


def test_image_data_comes_from_open_xml():
	doc = makeDoc()
	ctx = Context(doc)
	picture = a11y.collectPictures(ctx, 10)[0]
	data, mime = a11y.imageData(ctx, picture)
	assert (data, mime) == (b"PNGDATA", "image/png")


def test_drawing_formats_are_refused_with_a_message():
	doc = makeDoc()
	doc.inlineShapes[0].mime = "image/x-emf"
	ctx = Context(doc)
	picture = a11y.collectPictures(ctx, 10)[0]
	try:
		a11y.imageData(ctx, picture)
	except a11y.ImageUnavailable as error:
		assert "image/x-emf" in str(error)
	else:
		raise AssertionError("EMF should not be sent to AI")


def test_ai_alt_text_is_saved_after_editing(monkeypatch):
	doc = makeDoc()
	ctx = Context(doc)
	monkeypatch.setattr(
		a11y.ai, "describeImage", lambda image, mime, context, detailed: f"AI {len(image)} {detailed}"
	)
	kind = a11y.Pictures()
	items, _total = kind.collect(ctx, 10)
	host = Host()
	labels = [a.label for a in kind.actions(ctx, items[0], None, host)]
	assert "Describe with AI (short alternative text)..." in labels
	kind.actions(ctx, items[0], None, host)[1].run()
	assert doc.inlineShapes[0].AlternativeText == "AI 7 False"


def test_unclear_links():
	assert a11y._isUnclearLink("คลิกที่นี่", "https://a.th")
	assert a11y._isUnclearLink("https://a.th", "https://a.th")
	assert not a11y._isUnclearLink("ระเบียบงานสารบรรณ", "https://a.th")


def test_base64_roundtrip_in_fake():
	assert base64.b64decode(base64.b64encode(b"x")) == b"x"

import types

from fakeword import Context, Doc
from wordDocAssistant import presets


def letterDoc():
	return Doc(
		[
			"ที่ ศธ 0123/456",
			"13 ตุลาคม 2569",
			"",
			"เรื่อง ขอเชิญประชุม",
			"เรียน ผู้อำนวยการ",
			"ด้วยสำนักงาน",
			"ขอแสดงความนับถือ",
			"(นายสมชาย ใจดี)",
		]
	)


ROLES = ["refNo", "date", "empty", "subject", "recipient", "body", "closing", "signature"]


def test_every_preset_has_body_and_known_roles():
	names = set(presets.roleNames())
	for preset in presets.presets():
		assert "body" in preset.roles, preset.id
		assert set(preset.roles) <= names, preset.id
		assert set(preset.roles) <= set(presets.ROLE_HINTS), preset.id


def test_official_letter_is_applied_in_one_undo_step():
	doc = letterDoc()
	ctx = Context(doc)
	preset = presets.presetById("officialLetter")
	paragraphs, texts = presets.collectParagraphs(doc.Content, 100)
	counts = presets.apply(ctx, preset, paragraphs, ROLES, presets.ApplyOptions(), wholeDocument=True)
	assert counts == {"formatted": 7, "removed": 1}
	assert ctx.application.undo == ["Format document: Thai official letter (หนังสือภายนอก)"]
	assert ctx.application.ScreenUpdating is True
	body = doc.paraProps[5]
	assert body["Alignment"] == presets.ALIGN_THAI and round(body["FirstLineIndent"] / presets.CM, 1) == 2.5


def test_lists_and_tables_are_not_broken(monkeypatch):
	doc = letterDoc()
	ctx = Context(doc)
	paragraphs, _texts = presets.collectParagraphs(doc.Content, 100)
	monkeypatch.setattr(presets, "_inTable", lambda p: p.index == 5)
	monkeypatch.setattr(presets, "_hasNumbering", lambda p: p.index == 7)
	presets.apply(
		ctx,
		presets.presetById("report"),
		paragraphs,
		["body"] * 8,
		presets.ApplyOptions(removeEmpty=False),
		True,
	)
	assert "Alignment" not in doc.paraProps.get(5, {})
	assert "Style" not in doc.paraProps.get(7, {})


def test_every_preset_applies_without_errors():
	for preset in presets.presets():
		doc = letterDoc()
		paragraphs, _texts = presets.collectParagraphs(doc.Content, 100)
		roles = [r if r in preset.roles else "body" for r in ROLES]
		presets.apply(Context(doc), preset, paragraphs, roles, presets.ApplyOptions(), True)


def test_missing_font_falls_back():
	ctx = types.SimpleNamespace(
		application=types.SimpleNamespace(FontNames=types.SimpleNamespace(Count=1, Item=lambda i: "Tahoma"))
	)
	assert presets._availableFont(ctx, "TH SarabunIT๙", ("TH Sarabun New", "Tahoma")) == "Tahoma"

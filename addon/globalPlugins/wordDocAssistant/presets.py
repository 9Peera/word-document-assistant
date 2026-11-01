# Word Document Assistant for NVDA
# Document format presets and how they are applied through the Word object model.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

"""Format presets.

AI only decides the role of each paragraph (title, subject, body, signature...).
The look of every role comes from the preset below, so the result is predictable,
the text itself never changes, and an organization can adjust the numbers here.
The measurements of the Thai official formats follow common practice; check them
against your organization's own rules before relying on them.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import addonHandler

from .wordCom import COM_ERRORS, WordContext, undoRecord

addonHandler.initTranslation()

CM = 28.3465  # points per centimetre

# Word constants
ALIGN_LEFT = 0
ALIGN_CENTER = 1
ALIGN_RIGHT = 2
ALIGN_JUSTIFY = 3
ALIGN_THAI = 9  # wdAlignParagraphThaiJustify: justified without stretching Thai words
STYLE_NORMAL = -1
STYLE_HEADING1 = -2
STYLE_HEADING2 = -3
STYLE_HEADING3 = -4
STYLE_TITLE = -63
STYLE_SUBTITLE = -75
STYLE_LIST_PARAGRAPH = -180
STYLE_QUOTE = -181
COLOR_AUTOMATIC = -16777216
LINE_SPACE_MULTIPLE = 5
PAPER_A4 = 7
WITHIN_TABLE = 12


def roleNames() -> dict[str, str]:
	"""Paragraph roles that AI can assign, with the name shown to the user."""
	return {
		# Translators: Role of a paragraph when formatting a document.
		"title": _("Title"),
		# Translators: Role of a paragraph when formatting a document.
		"subtitle": _("Subtitle or author"),
		# Translators: Role of a paragraph when formatting a document.
		"heading1": _("Heading level 1"),
		# Translators: Role of a paragraph when formatting a document.
		"heading2": _("Heading level 2"),
		# Translators: Role of a paragraph when formatting a document.
		"heading3": _("Heading level 3"),
		# Translators: Role of a paragraph when formatting a document.
		"body": _("Body text"),
		# Translators: Role of a paragraph when formatting a document.
		"list": _("List item"),
		# Translators: Role of a paragraph when formatting a document.
		"quote": _("Quotation"),
		# Translators: Role of a paragraph when formatting a document (Thai letters: ที่ ...).
		"refNo": _("Reference number"),
		# Translators: Role of a paragraph when formatting a document.
		"sender": _("Sender name and address"),
		# Translators: Role of a paragraph when formatting a document.
		"date": _("Date"),
		# Translators: Role of a paragraph when formatting a document (เรื่อง ...).
		"subject": _("Subject"),
		# Translators: Role of a paragraph when formatting a document (เรียน ..., To ...).
		"recipient": _("Recipient"),
		# Translators: Role of a paragraph when formatting a document (อ้างถึง ...).
		"reference": _("Reference to earlier letter"),
		# Translators: Role of a paragraph when formatting a document (สิ่งที่ส่งมาด้วย ...).
		"enclosure": _("Enclosures"),
		# Translators: Role of a paragraph when formatting a document (Dear ...).
		"salutation": _("Salutation"),
		# Translators: Role of a paragraph when formatting a document (ขอแสดงความนับถือ, Yours sincerely).
		"closing": _("Closing"),
		# Translators: Role of a paragraph when formatting a document.
		"signature": _("Signature, name or position"),
		# Translators: Role of a paragraph when formatting a document.
		"contact": _("Contact details of the office"),
		# Translators: Role of a paragraph when formatting a document.
		"abstract": _("Abstract"),
		# Translators: Role of a paragraph when formatting a document.
		"keywords": _("Keywords"),
		# Translators: Role of a paragraph when formatting a document.
		"bibliography": _("Bibliography entry"),
		# Translators: Role of a paragraph when formatting a document.
		"caption": _("Caption of a table or figure"),
		# Translators: Role of a paragraph when formatting a document.
		"meta": _("Meeting details (date, place, attendees)"),
		# Translators: Role of a paragraph when formatting a document.
		"empty": _("Empty paragraph used for spacing"),
	}


#: Short English definitions sent to the AI.
ROLE_HINTS = {
	"title": "main title of the document (e.g. บันทึกข้อความ, ประกาศ..., report title)",
	"subtitle": "subtitle, author or organization line under the title",
	"heading1": "main section heading",
	"heading2": "sub-section heading, or a meeting agenda item (ระเบียบวาระที่ ...)",
	"heading3": "minor heading",
	"body": "ordinary paragraph",
	"list": "item of a bulleted or numbered list",
	"quote": "quoted passage",
	"refNo": "letter number line starting with ที่ or No.",
	"sender": "sender organization name or address block",
	"date": "date line of a letter or announcement (including ประกาศ ณ วันที่ / สั่ง ณ วันที่)",
	"subject": "subject line starting with เรื่อง or Subject/Re:",
	"recipient": "addressee line starting with เรียน, To, or the inside address",
	"reference": "line starting with อ้างถึง",
	"enclosure": "line starting with สิ่งที่ส่งมาด้วย or Enclosure",
	"salutation": "greeting such as Dear ...",
	"closing": "closing such as ขอแสดงความนับถือ, Yours sincerely",
	"signature": "signature block lines: (name), position",
	"contact": "office contact at the end: ส่วนราชการเจ้าของเรื่อง, โทร., e-mail",
	"abstract": "abstract paragraph or its heading บทคัดย่อ/Abstract",
	"keywords": "keywords line คำสำคัญ/Keywords",
	"bibliography": "entry of the reference list or bibliography",
	"caption": "caption of a table or figure (ตารางที่, ภาพที่, Table, Figure)",
	"meta": "meeting date, time, place, list of attendees or absentees",
	"empty": "empty paragraph used only to make space",
}


@dataclass
class RoleFormat:
	style: int | None = None
	align: int | None = None
	leftCm: float | None = None
	firstLineCm: float | None = None
	before: float | None = None
	after: float | None = None
	bold: bool | None = None
	size: float | None = None  # Thai size in points; Latin text is scaled to match
	keepWithNext: bool | None = None


@dataclass
class Preset:
	id: str
	name: str
	description: str
	fontThai: str
	fontLatin: str
	sizeThai: float
	sizeLatin: float
	lineSpacing: float = 1.0
	spaceAfter: float = 0
	bodyAlign: int = ALIGN_THAI
	margins: tuple[float, float, float, float] = (2.5, 2.0, 3.0, 2.0)  # top, bottom, left, right in cm
	roles: dict[str, RoleFormat] = field(default_factory=dict)

	@property
	def roleIds(self) -> list[str]:
		return list(self.roles)


def _official(**extra: RoleFormat) -> dict[str, RoleFormat]:
	"""Roles shared by Thai official letters."""
	roles = {
		"refNo": RoleFormat(align=ALIGN_LEFT, firstLineCm=0, leftCm=0),
		"sender": RoleFormat(align=ALIGN_LEFT, leftCm=10.0, firstLineCm=0),
		"date": RoleFormat(align=ALIGN_LEFT, leftCm=7.5, firstLineCm=0, before=6),
		"subject": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=6),
		"recipient": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=6),
		"reference": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=6),
		"enclosure": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0),
		"body": RoleFormat(align=ALIGN_THAI, leftCm=0, firstLineCm=2.5, before=6),
		"list": RoleFormat(align=ALIGN_THAI, leftCm=2.5, firstLineCm=0),
		"closing": RoleFormat(align=ALIGN_LEFT, leftCm=7.5, firstLineCm=0, before=12),
		"signature": RoleFormat(align=ALIGN_CENTER, leftCm=5.0, firstLineCm=0),
		"contact": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=12),
		"empty": RoleFormat(),
	}
	roles.update(extra)
	return roles


def presets() -> list[Preset]:
	"""All presets, in the order shown to the user."""
	thai, thaiOfficial, latin = "TH Sarabun New", "TH SarabunIT๙", "Calibri"
	return [
		Preset(
			id="officialLetter",
			# Translators: Name of a document format.
			name=_("Thai official letter (หนังสือภายนอก)"),
			# Translators: Description of a document format.
			description=_(
				"A4, margins 2.5 top, 2 bottom, 3 left, 2 right cm; TH SarabunIT๙ 16 pt. "
				"ที่, ส่วนราชการ, วันที่, เรื่อง, เรียน, อ้างถึง, สิ่งที่ส่งมาด้วย, body indented 2.5 cm, "
				"ขอแสดงความนับถือ and signature block in their usual places.",
			),
			fontThai=thaiOfficial,
			fontLatin=thaiOfficial,
			sizeThai=16,
			sizeLatin=16,
			roles=_official(),
		),
		Preset(
			id="memo",
			# Translators: Name of a document format.
			name=_("Internal memo (บันทึกข้อความ)"),
			# Translators: Description of a document format.
			description=_(
				"บันทึกข้อความ title in bold 29 pt; ส่วนราชการ, ที่, วันที่, เรื่อง in bold; "
				"เรียน, body indented 2.5 cm and signature block. TH SarabunIT๙ 16 pt.",
			),
			fontThai=thaiOfficial,
			fontLatin=thaiOfficial,
			sizeThai=16,
			sizeLatin=16,
			margins=(1.5, 2.0, 3.0, 2.0),
			roles=_official(
				title=RoleFormat(align=ALIGN_CENTER, bold=True, size=29, after=6, keepWithNext=True),
				sender=RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, bold=True),
				refNo=RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, bold=True),
				date=RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, bold=True),
				subject=RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, bold=True, after=6),
			),
		),
		Preset(
			id="announcement",
			# Translators: Name of a document format.
			name=_("Announcement or order (ประกาศ / คำสั่ง)"),
			# Translators: Description of a document format.
			description=_(
				"Title and subject centred in bold, body indented 2.5 cm, "
				"ประกาศ ณ วันที่ and signature centred. TH SarabunIT๙ 16 pt.",
			),
			fontThai=thaiOfficial,
			fontLatin=thaiOfficial,
			sizeThai=16,
			sizeLatin=16,
			roles=_official(
				title=RoleFormat(align=ALIGN_CENTER, bold=True, leftCm=0, firstLineCm=0, keepWithNext=True),
				subject=RoleFormat(align=ALIGN_CENTER, bold=True, leftCm=0, firstLineCm=0, after=12),
				date=RoleFormat(align=ALIGN_CENTER, leftCm=0, firstLineCm=0, before=12),
				signature=RoleFormat(align=ALIGN_CENTER, leftCm=3.0, firstLineCm=0),
			),
		),
		Preset(
			id="businessLetter",
			# Translators: Name of a document format.
			name=_("Formal or business letter"),
			# Translators: Description of a document format.
			description=_(
				"Sender block, date, recipient, subject in bold, salutation, body with space between "
				"paragraphs, closing and signature, all aligned left. TH Sarabun New 16 pt, Calibri 11 pt.",
			),
			fontThai=thai,
			fontLatin=latin,
			sizeThai=16,
			sizeLatin=11,
			spaceAfter=6,
			bodyAlign=ALIGN_THAI,
			margins=(2.54, 2.54, 2.54, 2.54),
			roles={
				"sender": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0),
				"date": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=12, after=12),
				"refNo": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0),
				"recipient": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0),
				"subject": RoleFormat(
					align=ALIGN_LEFT, leftCm=0, firstLineCm=0, bold=True, before=12, after=6
				),
				"salutation": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=6, after=6),
				"body": RoleFormat(align=ALIGN_THAI, leftCm=0, firstLineCm=0, after=8),
				"list": RoleFormat(style=STYLE_LIST_PARAGRAPH, align=ALIGN_LEFT, leftCm=1.27),
				"enclosure": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=12),
				"closing": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=12, after=24),
				"signature": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, after=0),
				"contact": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, before=12),
				"empty": RoleFormat(),
			},
		),
		Preset(
			id="personalLetter",
			# Translators: Name of a document format.
			name=_("General letter"),
			# Translators: Description of a document format.
			description=_(
				"Sender address and date at the right, salutation, body indented, closing and name at the right. "
				"TH Sarabun New 16 pt, Calibri 12 pt.",
			),
			fontThai=thai,
			fontLatin=latin,
			sizeThai=16,
			sizeLatin=12,
			spaceAfter=6,
			margins=(2.54, 2.54, 2.54, 2.54),
			roles={
				"sender": RoleFormat(align=ALIGN_RIGHT, leftCm=0, firstLineCm=0),
				"date": RoleFormat(align=ALIGN_RIGHT, leftCm=0, firstLineCm=0, before=6, after=12),
				"salutation": RoleFormat(align=ALIGN_LEFT, leftCm=0, firstLineCm=0, after=6),
				"body": RoleFormat(align=ALIGN_THAI, leftCm=0, firstLineCm=1.27, after=6),
				"list": RoleFormat(style=STYLE_LIST_PARAGRAPH, align=ALIGN_LEFT, leftCm=1.27),
				"closing": RoleFormat(align=ALIGN_RIGHT, leftCm=0, firstLineCm=0, before=12),
				"signature": RoleFormat(align=ALIGN_RIGHT, leftCm=0, firstLineCm=0),
				"empty": RoleFormat(),
			},
		),
		Preset(
			id="report",
			# Translators: Name of a document format.
			name=_("Formal report"),
			# Translators: Description of a document format.
			description=_(
				"Word Title and Heading styles (so the headings can be listed and navigated), "
				"body justified for Thai, captions centred. Margins 3.81 top and left, 2.54 bottom and right cm. "
				"TH Sarabun New 16 pt, Calibri 11 pt.",
			),
			fontThai=thai,
			fontLatin=latin,
			sizeThai=16,
			sizeLatin=11,
			spaceAfter=6,
			margins=(3.81, 2.54, 3.81, 2.54),
			roles={
				"title": RoleFormat(style=STYLE_TITLE, align=ALIGN_CENTER, bold=True, size=24, after=12),
				"subtitle": RoleFormat(style=STYLE_SUBTITLE, align=ALIGN_CENTER, size=18, after=18),
				"heading1": RoleFormat(style=STYLE_HEADING1, bold=True, size=20, before=18, after=6),
				"heading2": RoleFormat(style=STYLE_HEADING2, bold=True, size=18, before=12, after=6),
				"heading3": RoleFormat(style=STYLE_HEADING3, bold=True, size=16, before=12, after=6),
				"body": RoleFormat(align=ALIGN_THAI, firstLineCm=1.27),
				"list": RoleFormat(style=STYLE_LIST_PARAGRAPH, leftCm=1.27),
				"quote": RoleFormat(style=STYLE_QUOTE, leftCm=1.27),
				"caption": RoleFormat(align=ALIGN_CENTER, bold=True, before=6, after=6),
				"bibliography": RoleFormat(align=ALIGN_LEFT, leftCm=1.27, firstLineCm=-1.27),
				"empty": RoleFormat(),
			},
		),
		Preset(
			id="article",
			# Translators: Name of a document format.
			name=_("Article"),
			# Translators: Description of a document format.
			description=_(
				"Title and author line, Heading styles for sections, body with 1.15 line spacing and "
				"space between paragraphs, quotations indented. TH Sarabun New 16 pt, Calibri 11 pt.",
			),
			fontThai=thai,
			fontLatin=latin,
			sizeThai=16,
			sizeLatin=11,
			lineSpacing=1.15,
			spaceAfter=8,
			bodyAlign=ALIGN_LEFT,
			margins=(2.54, 2.54, 2.54, 2.54),
			roles={
				"title": RoleFormat(style=STYLE_TITLE, align=ALIGN_LEFT, bold=True, size=26, after=6),
				"subtitle": RoleFormat(style=STYLE_SUBTITLE, align=ALIGN_LEFT, size=16, after=18),
				"heading1": RoleFormat(style=STYLE_HEADING1, bold=True, size=20, before=18, after=6),
				"heading2": RoleFormat(style=STYLE_HEADING2, bold=True, size=18, before=12, after=4),
				"heading3": RoleFormat(style=STYLE_HEADING3, bold=True, size=16, before=12, after=4),
				"body": RoleFormat(align=ALIGN_LEFT, firstLineCm=0),
				"list": RoleFormat(style=STYLE_LIST_PARAGRAPH, leftCm=1.27),
				"quote": RoleFormat(style=STYLE_QUOTE, leftCm=1.27),
				"caption": RoleFormat(align=ALIGN_CENTER, before=4, after=8),
				"empty": RoleFormat(),
			},
		),
		Preset(
			id="academic",
			# Translators: Name of a document format.
			name=_("Academic paper"),
			# Translators: Description of a document format.
			description=_(
				"Centred title and authors, abstract and keywords, Heading styles, body justified and "
				"indented, captions, bibliography with hanging indent. Margins 3.81 top and left, "
				"2.54 bottom and right cm.",
			),
			fontThai=thai,
			fontLatin="Times New Roman",
			sizeThai=16,
			sizeLatin=12,
			margins=(3.81, 2.54, 3.81, 2.54),
			roles={
				"title": RoleFormat(style=STYLE_TITLE, align=ALIGN_CENTER, bold=True, size=20, after=6),
				"subtitle": RoleFormat(style=STYLE_SUBTITLE, align=ALIGN_CENTER, size=16, after=12),
				"abstract": RoleFormat(align=ALIGN_THAI, leftCm=1.0, firstLineCm=1.27, after=6),
				"keywords": RoleFormat(align=ALIGN_LEFT, leftCm=1.0, after=12),
				"heading1": RoleFormat(style=STYLE_HEADING1, bold=True, size=18, before=12, after=6),
				"heading2": RoleFormat(style=STYLE_HEADING2, bold=True, size=16, before=12, after=6),
				"heading3": RoleFormat(style=STYLE_HEADING3, bold=True, size=16, before=6, after=6),
				"body": RoleFormat(align=ALIGN_THAI, firstLineCm=1.27),
				"list": RoleFormat(style=STYLE_LIST_PARAGRAPH, leftCm=1.27),
				"quote": RoleFormat(style=STYLE_QUOTE, leftCm=1.27),
				"caption": RoleFormat(align=ALIGN_CENTER, bold=True, before=6, after=6),
				"bibliography": RoleFormat(align=ALIGN_LEFT, leftCm=1.27, firstLineCm=-1.27, after=4),
				"empty": RoleFormat(),
			},
		),
		Preset(
			id="minutes",
			# Translators: Name of a document format.
			name=_("Meeting minutes (รายงานการประชุม)"),
			# Translators: Description of a document format.
			description=_(
				"Centred title and meeting details, attendee lists, each agenda item as a heading "
				"(ระเบียบวาระที่ ...), body indented, recorder and checker signatures at the right.",
			),
			fontThai=thaiOfficial,
			fontLatin=thaiOfficial,
			sizeThai=16,
			sizeLatin=16,
			roles={
				"title": RoleFormat(style=STYLE_TITLE, align=ALIGN_CENTER, bold=True, size=18, after=0),
				"meta": RoleFormat(align=ALIGN_CENTER, leftCm=0, firstLineCm=0),
				"heading1": RoleFormat(style=STYLE_HEADING1, bold=True, size=16, before=12, after=0),
				"heading2": RoleFormat(style=STYLE_HEADING2, bold=True, size=16, before=12, after=0),
				"body": RoleFormat(align=ALIGN_THAI, leftCm=0, firstLineCm=2.5),
				"list": RoleFormat(align=ALIGN_LEFT, leftCm=2.5, firstLineCm=0),
				"signature": RoleFormat(align=ALIGN_CENTER, leftCm=7.0, firstLineCm=0),
				"empty": RoleFormat(),
			},
		),
		Preset(
			id="accessible",
			# Translators: Name of a document format.
			name=_("Easy to read and accessible document"),
			# Translators: Description of a document format.
			description=_(
				"Real Word headings for screen reader navigation, text aligned left (not justified), "
				"1.5 line spacing, space between paragraphs instead of empty lines, larger text. "
				"Good for people who use screen readers or have low vision.",
			),
			fontThai=thai,
			fontLatin="Arial",
			sizeThai=18,
			sizeLatin=13,
			lineSpacing=1.5,
			spaceAfter=10,
			bodyAlign=ALIGN_LEFT,
			margins=(2.54, 2.54, 2.54, 2.54),
			roles={
				"title": RoleFormat(style=STYLE_TITLE, align=ALIGN_LEFT, bold=True, size=28, after=12),
				"subtitle": RoleFormat(style=STYLE_SUBTITLE, align=ALIGN_LEFT, size=20),
				"heading1": RoleFormat(style=STYLE_HEADING1, bold=True, size=24, before=18, after=8),
				"heading2": RoleFormat(style=STYLE_HEADING2, bold=True, size=21, before=14, after=6),
				"heading3": RoleFormat(style=STYLE_HEADING3, bold=True, size=19, before=12, after=6),
				"body": RoleFormat(align=ALIGN_LEFT, firstLineCm=0, leftCm=0),
				"list": RoleFormat(style=STYLE_LIST_PARAGRAPH, align=ALIGN_LEFT, leftCm=1.27),
				"quote": RoleFormat(style=STYLE_QUOTE, align=ALIGN_LEFT, leftCm=1.27),
				"caption": RoleFormat(align=ALIGN_LEFT, bold=True),
				"empty": RoleFormat(),
			},
		),
	]


def presetById(presetId: str) -> Preset | None:
	return next((p for p in presets() if p.id == presetId), None)


# Applying --------------------------------------------------------------------------


@dataclass
class ApplyOptions:
	pageSetup: bool = True
	fonts: bool = True
	removeEmpty: bool = True


def _availableFont(ctx: WordContext, wanted: str, fallbacks: tuple[str, ...]) -> str:
	"""The wanted font if installed, otherwise the first installed fallback."""
	try:
		names = ctx.application.FontNames
		installed = {str(names.Item(i)) for i in range(1, names.Count + 1)}
	except COM_ERRORS:
		return wanted
	for name in (wanted, *fallbacks):
		if name in installed:
			return name
	return wanted


def _setFont(font: Any, latin: str, thai: str, latinSize: float, thaiSize: float) -> None:
	font.Name = latin
	font.NameBi = thai
	font.Size = latinSize
	font.SizeBi = thaiSize


def _latinSize(preset: Preset, thaiSize: float) -> float:
	return round(thaiSize * preset.sizeLatin / preset.sizeThai * 2) / 2


def collectParagraphs(rangeObj: Any, limit: int) -> tuple[list[Any], list[str]]:
	"""Paragraph objects of the range and their text, walking with Next for speed."""
	paragraphs, texts = [], []
	paragraph = rangeObj.Paragraphs.First
	end = rangeObj.End
	while paragraph is not None and len(paragraphs) < limit:
		paraRange = paragraph.Range
		if paraRange.Start >= end:
			break
		paragraphs.append(paragraph)
		texts.append(str(paraRange.Text or "").rstrip("\r\x07"))
		try:
			paragraph = paragraph.Next()
		except COM_ERRORS:
			break
	return paragraphs, texts


def _isRemovableEmpty(paragraph: Any) -> bool:
	paraRange = paragraph.Range
	if str(paraRange.Text or "").strip("\r\n\x07\x0b\x0c \t"):
		return False
	try:
		if paraRange.InlineShapes.Count or paraRange.Information(WITHIN_TABLE):
			return False
	except COM_ERRORS:
		return False
	return True


def apply(
	ctx: WordContext,
	preset: Preset,
	paragraphs: list[Any],
	roles: list[str],
	options: ApplyOptions,
	wholeDocument: bool,
) -> dict[str, int]:
	"""Format the paragraphs. Returns counts of what was done. One undo step."""
	app = ctx.application
	fontThai = _availableFont(
		ctx, preset.fontThai, ("TH SarabunIT๙", "TH Sarabun New", "TH SarabunPSK", "Leelawadee UI", "Tahoma")
	)
	fontLatin = _availableFont(ctx, preset.fontLatin, (fontThai, "Calibri", "Arial"))
	counts = {"formatted": 0, "removed": 0}
	try:
		app.ScreenUpdating = False
	except COM_ERRORS:
		pass
	try:
		_applyAll(ctx, preset, paragraphs, roles, options, wholeDocument, fontThai, fontLatin, counts)
	finally:
		try:
			app.ScreenUpdating = True
		except COM_ERRORS:
			pass
	return counts


def _applyAll(ctx, preset, paragraphs, roles, options, wholeDocument, fontThai, fontLatin, counts) -> None:
	doc = ctx.document
	app = ctx.application
	# Translators: Name of the undo step for formatting a document.
	with undoRecord(ctx, _("Format document: {name}").format(name=preset.name)):
		if options.pageSetup and wholeDocument:
			setup = doc.PageSetup
			try:
				setup.PaperSize = PAPER_A4
			except COM_ERRORS:
				pass
			top, bottom, left, right = preset.margins
			setup.TopMargin = top * CM
			setup.BottomMargin = bottom * CM
			setup.LeftMargin = left * CM
			setup.RightMargin = right * CM
		if options.fonts:
			normal = doc.Styles.Item(STYLE_NORMAL)
			_setFont(normal.Font, fontLatin, fontThai, preset.sizeLatin, preset.sizeThai)
			paragraphFormat = normal.ParagraphFormat
			paragraphFormat.SpaceAfter = preset.spaceAfter
			paragraphFormat.SpaceBefore = 0
			if preset.lineSpacing != 1.0:
				paragraphFormat.LineSpacingRule = LINE_SPACE_MULTIPLE
				paragraphFormat.LineSpacing = app.LinesToPoints(preset.lineSpacing)
			else:
				paragraphFormat.LineSpacingRule = 0
			# Headings and titles in the same fonts, in black, at the preset sizes.
			for role in preset.roles.values():
				if role.style is None or role.style == STYLE_NORMAL:
					continue
				try:
					style = doc.Styles.Item(role.style)
					size = role.size or preset.sizeThai
					_setFont(style.Font, fontLatin, fontThai, _latinSize(preset, size), size)
					style.Font.Color = COLOR_AUTOMATIC
					if role.bold is not None:
						style.Font.Bold = role.bold
						style.Font.BoldBi = role.bold
				except COM_ERRORS:
					continue
		toRemove = []
		for paragraph, roleId in zip(paragraphs, roles, strict=True):
			role = preset.roles.get(roleId) or preset.roles.get("body") or RoleFormat()
			if roleId == "empty":
				if options.removeEmpty and _isRemovableEmpty(paragraph):
					toRemove.append(paragraph)
				continue
			try:
				_formatParagraph(paragraph, preset, role, options)
				counts["formatted"] += 1
			except COM_ERRORS:
				continue
		# From the end, so the paragraphs before stay where they are.
		for paragraph in reversed(toRemove):
			try:
				paragraph.Range.Delete()
				counts["removed"] += 1
			except COM_ERRORS:
				continue


def _inTable(paragraph: Any) -> bool:
	try:
		return bool(paragraph.Range.Information(WITHIN_TABLE))
	except COM_ERRORS:
		return False


def _hasNumbering(paragraph: Any) -> bool:
	try:
		return paragraph.Range.ListFormat.ListType != 0
	except COM_ERRORS:
		return False


def _formatParagraph(paragraph: Any, preset: Preset, role: RoleFormat, options: ApplyOptions) -> None:
	if _inTable(paragraph):
		# Indents and alignment would break table layout; tables get the fonts through the Normal style.
		return
	if role.style is not None:
		paragraph.Style = role.style
	elif not _hasNumbering(paragraph):
		# Keep the style of bulleted and numbered paragraphs so their numbers are not lost.
		paragraph.Style = STYLE_NORMAL
	paragraph.Alignment = role.align if role.align is not None else preset.bodyAlign
	if role.leftCm is not None:
		paragraph.LeftIndent = role.leftCm * CM
	if role.firstLineCm is not None:
		paragraph.FirstLineIndent = role.firstLineCm * CM
	if role.before is not None:
		paragraph.SpaceBefore = role.before
	if role.after is not None:
		paragraph.SpaceAfter = role.after
	if role.keepWithNext is not None:
		paragraph.KeepWithNext = role.keepWithNext
	font = paragraph.Range.Font
	if role.bold is not None:
		font.Bold = role.bold
		font.BoldBi = role.bold
	if role.size is not None and options.fonts:
		font.SizeBi = role.size
		font.Size = _latinSize(preset, role.size)

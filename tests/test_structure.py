import xml.dom.minidom

from fakeword import Context, Doc, Field, Host, Link, Table
from wordDocAssistant import structure


def headingDoc():
	doc = Doc(["Title", "body", "Part one", "body", "Sub part", "body"])
	doc.setLevel(0, 1)
	doc.setLevel(2, 2)
	doc.setLevel(4, 3)
	return doc


def test_headings_are_found_with_goto_and_can_change_level():
	doc = headingDoc()
	ctx = Context(doc)
	kind = structure.Headings()
	items, total = kind.collect(ctx, 100)
	assert total == 3
	assert [kind.summary(ctx, i) for i in items] == [
		"Level 1: Title",
		"Level 2: Part one",
		"Level 3: Sub part",
	]
	host = Host()
	labels = [a.label for a in kind.actions(ctx, items[2], None, host)]
	assert labels[:2] == ["Promote to level 2", "Demote to level 4"]
	kind.actions(ctx, items[2], None, host)[0].run()
	assert doc.paraProps[4]["level"] == 2
	body = next(a for a in kind.actions(ctx, items[1], None, host) if a.label == "Change to body text")
	body.run()
	assert doc.paraProps[2]["level"] == 10


def test_tables_summary_header_and_ai_description(monkeypatch):
	doc = Doc(["Name\tAge", "after"])
	ctx = Context(doc)
	table = Table(doc, 0, 9, rows=4, columns=2)
	doc.tables.append(table)
	kind = structure.Tables()
	items, _total = kind.collect(ctx, 10)
	assert kind.summary(ctx, items[0]).startswith("Table 1: 4 rows, 2 columns, header row no")
	host = Host()
	kind.actions(ctx, items[0], None, host)[0].run()
	assert table.Rows.Item(1).HeadingFormat is True
	monkeypatch.setattr(structure.ai, "describeTable", lambda text, caption: "Ages of staff")
	describe = next(a for a in kind.actions(ctx, items[0], None, host) if "AI" in a.label)
	describe.run()
	assert table.Descr == "Ages of staff"


def test_broken_internal_link():
	doc = Doc(["go to results"])
	start, end = doc.find("results")
	doc.links.append(Link(doc, start, end, subAddress="_Results"))
	ctx = Context(doc)
	kind = structure.Links()
	items, _total = kind.collect(ctx, 10)
	assert kind.summary(ctx, items[0]).endswith("broken")
	doc.Bookmarks.Add("_Results", doc.Range(0, 2))
	assert not kind.summary(ctx, items[0]).endswith("broken")
	host = Host()
	remove = next(a for a in kind.actions(ctx, items[0], None, host) if a.label.startswith("Remove link"))
	remove.run()
	assert doc.links[0].deleted


def test_bookmark_names_are_checked():
	doc = Doc(["text"])
	ctx = Context(doc)
	anchor = doc.Range(0, 4)
	kind = structure.Bookmarks()
	host = Host(prompts=["1bad"])
	next(a for a in kind.actions(ctx, None, anchor, host) if a.label.startswith("Add")).run()
	assert doc.bookmarks == [] and "Not a valid" in host.spoken[-1]
	host = Host(prompts=["good name"])
	next(a for a in kind.actions(ctx, None, anchor, host) if a.label.startswith("Add")).run()
	assert doc.bookmarks[0].Name == "good_name"


def test_fields_errors_and_cross_reference_targets():
	doc = Doc(["See Error! Reference source not found. and 3", "Results"])
	ctx = Context(doc)
	errStart, errEnd = doc.find("Error! Reference source not found.")
	okStart, okEnd = doc.find("3")
	doc.fields += [
		Field(doc, "REF _Ref1 \\h", errStart, errEnd),
		Field(doc, "REF _Ref2 \\h", okStart, okEnd),
		Field(doc, "PAGE", 0, 0),
	]
	resultsStart = doc.find("Results")[0]
	doc.Bookmarks.Add("_Ref2", doc.Range(resultsStart, resultsStart + 7))
	refs = structure.CrossReferences()
	items, total = refs.collect(ctx, 10)
	assert total == 2
	summaries = [refs.summary(ctx, i) for i in items]
	assert summaries[0].endswith("Error") and "(target not found)" in summaries[0]
	assert summaries[1] == "3, points to Results"
	fields = structure.Fields()
	assert fields.collect(ctx, 10)[1] == 3
	host = Host()
	updateAll = next(a for a in fields.actions(ctx, None, None, host) if a.label.startswith("Update all"))
	updateAll.run()
	assert all(f.updated for f in doc.fields)


def test_insert_cross_reference_asks_type_target_and_kind():
	doc = Doc(["text"])
	ctx = Context(doc)
	doc.crossReferenceItems[structure.REF_TYPE_HEADING] = ["Introduction", "Results"]
	anchor = doc.Range(0, 4)
	host = Host(choices=[0, 1, 1])
	structure.insertCrossReference(ctx, anchor, host)
	calls = doc.inserted
	assert calls == [(structure.REF_TYPE_HEADING, structure.KIND_PAGE, 2, True)]


def test_sources_and_citations():
	xmlText = structure.sourceXml(
		"Pee26", "Book", "Peera Pilarit; Somchai Jaidee", "Thai & Word", "2026", "SAINT", "",
	)
	xml.dom.minidom.parseString(xmlText)
	assert "<b:Last>Pilarit</b:Last>" in xmlText and "&amp;" in xmlText
	doc = Doc(["text"])
	ctx = Context(doc)
	host = Host(choices=[0, 0], prompts=["Peera Pilarit", "Word accessibility", "2026", "SE-ED"])
	anchor = doc.Range(0, 4)
	structure.insertCitation(ctx, anchor, host)
	assert doc.sources[0].Tag == "Pilarit26"
	assert doc.fields[0].Code.Text.strip() == "CITATION Pilarit26"
	citations = structure.Citations()
	items, _total = citations.collect(ctx, 10)
	assert "Pilarit" in citations.summary(ctx, items[0])

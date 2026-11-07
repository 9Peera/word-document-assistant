import json
import os
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import nvdaStubs
import pytest
from fakeword import Context, Doc, Host
from wordDocAssistant import thaicheck, thaidict


@pytest.fixture(scope="module")
def dictionary():
	d = thaidict.Dictionary()
	d.load([(thaidict.BUNDLED_WORDS, "words")], [thaidict.BUNDLED_MISSPELLINGS])
	return d


def test_starter_misspellings_are_safe(dictionary):
	pairs = thaidict.readMisspellings(thaidict.BUNDLED_MISSPELLINGS)
	assert len(pairs) >= 60
	rights = {right for right, _note in pairs.values()}
	for wrong, (right, _note) in pairs.items():
		assert wrong not in right, wrong
		assert not any(wrong in r for r in rights), wrong
		assert right in dictionary.words or right in rights, right


def test_collected_misspellings_do_not_count_as_correct(dictionary):
	assert "อีเมล์" not in dictionary.words and "อีเมล" in dictionary.words


def test_finds_common_misspellings_with_the_right_form(dictionary):
	findings = dictionary.check("กรุณาส่งอีเมล์ภายในสับดาห์นี้ ขอบคุณนะค่ะ")
	assert [(f.word, f.suggestions[0]) for f in findings] == [
		("อีเมล์", "อีเมล"),
		("สับดาห์", "สัปดาห์"),
		("นะค่ะ", "นะคะ"),
	]


def test_misspelling_inside_a_correct_longer_word_is_not_reported(dictionary):
	dictionary.words.add("ชลอม")
	try:
		assert [f.word for f in dictionary.check("ระบบจะชลอ และชลอมใบนี้")] == ["ชลอ"]
	finally:
		dictionary.words.discard("ชลอม")


def test_incomplete_word_is_reported_whole_with_a_completion(dictionary):
	findings = dictionary.check("ผมพิมพ์ผิดว่า คอมพิวเตอ")
	assert [(f.word, f.suggestions[:1]) for f in findings] == [("คอมพิวเตอ", ["คอมพิวเตอร์"])]


def test_names_after_a_title_are_not_reported(dictionary):
	assert dictionary.check("ประธานคือนายสมชชาย ใจดี", unknown=True) == [] or all(
		f.word != "สมชชาย" for f in dictionary.check("ประธานคือนายสมชชาย ใจดี")
	)


def test_segmentation_never_splits_a_vowel_from_its_consonant(dictionary):
	for start, end, _known in dictionary.segment("ผมพิมพ์ผิดว่า"):
		assert "ผมพิมพ์ผิดว่า"[start] not in thaidict._NO_START


def test_one_edit_suggestions(dictionary):
	assert "ประชุม" in dictionary.suggest("ประชุ่ม")


def test_large_text_is_fast(dictionary):
	import time

	text = "ขอเชิญประชุมคณะกรรมการเรื่องการพัฒนาเว็บไซท์สำหรับคนตาบอด " * 500
	start = time.time()
	dictionary.check(text)
	assert time.time() - start < 3


# Word lists and updates -------------------------------------------------------------


class ListServer:
	def __init__(self, body: str, etag: str):
		self.body, self.etag, self.requests = body, etag, []
		server = self

		class Handler(BaseHTTPRequestHandler):
			def _reply(self, withBody):
				server.requests.append((self.command, self.headers.get("If-None-Match")))
				if self.headers.get("If-None-Match") == server.etag:
					self.send_response(304)
					self.end_headers()
					return
				data = server.body.encode()
				self.send_response(200)
				self.send_header("ETag", server.etag)
				self.send_header("Content-Length", str(len(data)))
				self.end_headers()
				if withBody:
					self.wfile.write(data)

			def do_HEAD(self):
				self._reply(False)

			def do_GET(self):
				self._reply(True)

			def log_message(self, *args):
				pass

		self.http = HTTPServer(("127.0.0.1", 0), Handler)
		threading.Thread(target=self.http.serve_forever, daemon=True).start()
		self.url = f"http://127.0.0.1:{self.http.server_port}/list"


@pytest.fixture
def lists(tmp_path, config, monkeypatch):
	config.update(
		thaiCheckUnknown=True,
		wordListUpdates=False,
		wordListAutoDownload=False,
		wordListLibreOffice=False,
		misspellingListUrl="",
	)
	monkeypatch.setattr(
		nvdaStubs.sys.modules["globalVars"].appArgs, "configPath", str(tmp_path), raising=False,
	)
	monkeypatch.setattr(thaicheck, "_lists", None)
	return thaicheck.wordLists()


def test_bundled_list_is_used_until_a_download_exists(lists):
	assert len(lists.dictionary.words) > 50000
	assert lists.folder.endswith(os.path.join("wordDocAssistant", "dictionaries"))


def test_organization_list_update_cycle(lists, config, monkeypatch):
	# Only the local test server: the real word list websites change over time.
	monkeypatch.setattr(thaidict, "sources", lambda: [])
	server = ListServer("# list\nอนุญาติ\tอนุญาต\nกะเพา\tกะเพรา\nเกมส\tเกม\nอีเมว\tอีเมล\nสังเกด\tสังเกต\n", '"v1"')
	config["misspellingListUrl"] = server.url
	updates = lists.checkForUpdates()
	assert [s.id for s in updates] == ["organization"]
	assert lists.download(updates[0]) == 5
	lists.load()
	assert "กะเพา" in lists.dictionary.misspellings
	assert lists.checkForUpdates() == []
	assert server.requests[-1] == ("HEAD", '"v1"')
	server.etag = '"v2"'
	assert [s.id for s in lists.checkForUpdates()] == ["organization"]
	assert lists.pendingUpdates()[0].id == "organization"
	assert lists.dueForCheck() is False
	server.http.shutdown()


def test_a_web_page_is_not_accepted_as_a_word_list(lists, config):
	server = ListServer("<html><body>Not found</body></html>\n" * 20, '"x"')
	config["misspellingListUrl"] = server.url
	source = lists.enabledSources()[-1]
	with pytest.raises(ValueError):
		lists.download(source)
	assert not os.path.exists(os.path.join(lists.folder, "organization.tsv"))
	server.http.shutdown()


def test_my_words_are_never_reported(lists):
	lists.addMyWord("สมชชาย")
	lists.load()
	assert all(f.word != "สมชชาย" for f in lists.dictionary.check("คุณสมชชายมาแล้ว ทดสอบสมชชาย"))
	state = json.dumps(lists.state())
	assert state == "{}"


def test_element_type_replaces_and_ignores(lists, config):
	doc = Doc(["กรุณาส่งอีเมล์ภายในวันนี้", "อีเมล์ฉบับที่สอง"])
	ctx = Context(doc)
	kind = thaicheck.ThaiWords()
	items, total = kind.collect(ctx, 100)
	assert total == 2 and kind.summary(ctx, items[0]) == "อีเมล์, should be อีเมล"
	host = Host()
	replaceAll = next(
		a for a in kind.actions(ctx, items[0], None, host) if a.label.startswith("Replace every")
	)
	replaceAll.run()
	assert doc.text == "กรุณาส่งอีเมลภายในวันนี้\rอีเมลฉบับที่สอง\r"
	assert kind.collect(ctx, 100)[1] == 0


def test_weekly_check_is_off_until_the_user_agrees(lists, config, monkeypatch):
	from wordDocAssistant import dialogs, wordListUpdates

	monkeypatch.setattr(nvdaStubs.sys.modules["globalVars"].appArgs, "secure", False)
	checks = []
	monkeypatch.setattr(wordListUpdates, "scheduledCheck", lambda: checks.append("check"))
	answers = [nvdaStubs.sys.modules["wx"].ID_NO]
	monkeypatch.setattr(dialogs, "runModal", lambda make, read=None: (answers.pop(0), None))
	assert config["wordListUpdates"] is False
	wordListUpdates.askOnce()
	assert config["wordListAsked"] is True and config["wordListUpdates"] is False and checks == []
	wordListUpdates.askOnce()  # asked only once
	config["wordListAsked"] = False
	answers.append(nvdaStubs.sys.modules["wx"].ID_YES)
	wordListUpdates.askOnce()
	assert config["wordListUpdates"] is True and checks == ["check"]


def test_nothing_is_asked_on_secure_screens(lists, config):
	from wordDocAssistant import wordListUpdates

	wordListUpdates.askOnce()
	assert config["wordListAsked"] is False and config["wordListUpdates"] is False

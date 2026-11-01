import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest
from wordDocAssistant import ai, keyStore


class Server:
	"""A fake AI service. Set .reply(path, body, headers) -> (status, json)."""

	def __init__(self):
		self.requests = []
		self.reply = None
		server = self

		class Handler(BaseHTTPRequestHandler):
			def _answer(self, body):
				server.requests.append(
					(self.command, self.path, body, {k.lower(): v for k, v in self.headers.items()})
				)
				status, data = server.reply(self.path, body, self.headers)
				payload = json.dumps(data).encode()
				self.send_response(status)
				self.end_headers()
				self.wfile.write(payload)

			def do_GET(self):
				self._answer(None)

			def do_POST(self):
				self._answer(json.loads(self.rfile.read(int(self.headers["Content-Length"]))))

			def log_message(self, *args):
				pass

		self.http = HTTPServer(("127.0.0.1", 0), Handler)
		threading.Thread(target=self.http.serve_forever, daemon=True).start()
		self.url = f"http://127.0.0.1:{self.http.server_port}"


@pytest.fixture
def server(config):
	s = Server()
	ai.RETRY_DELAYS = (0, 0)
	ai._autoModel.clear()
	ai.lastModels.clear()
	yield s
	s.http.shutdown()


ANSWER = {
	"corrections": [
		{"original": "thsi", "suggestion": "this", "reason": "typo"},
		{"original": "absent", "suggestion": "x"},
	]
}


def wrap(provider, data):
	text = json.dumps(data, ensure_ascii=False)
	if provider == "gemini":
		return {"candidates": [{"content": {"parts": [{"text": text}]}}]}
	if provider == "anthropic":
		return {"content": [{"type": "text", "text": text}]}
	if provider == "ollama":
		return {"message": {"content": "```json\n" + text + "\n```"}}
	return {"choices": [{"message": {"content": "Here: " + text}}]}


@pytest.mark.parametrize("provider", ["gemini", "anthropic", "openaiApi", "openai", "ollama"])
def test_proofread_with_every_service(server, config, provider):
	config.update(
		aiProvider=provider,
		aiUrl=server.url,
		aiModel="m",
		aiApiKey="sk-ant-key" if provider == "anthropic" else "key",
	)
	server.reply = lambda path, body, headers: (200, wrap(provider, ANSWER))
	assert ai.proofread("helo thsi") == [{"original": "thsi", "suggestion": "this", "reason": "typo"}]


@pytest.mark.parametrize("provider", ["gemini", "anthropic", "openaiApi", "ollama"])
def test_pictures_are_sent_in_each_service_format(server, config, provider):
	config.update(aiProvider=provider, aiUrl=server.url, aiModel="m", aiApiKey="key")
	server.reply = lambda path, body, headers: (200, wrap(provider, {"text": "A bar chart"}))
	assert ai.describeImage(b"\x89PNG", "image/png", "Sales 2569", detailed=False) == "A bar chart"
	body = json.dumps(server.requests[-1][2])
	assert "iVBORw==" in body
	expected = {
		"gemini": "inlineData",
		"anthropic": '"type": "image"',
		"openaiApi": "data:image/png;base64",
		"ollama": '"images"',
	}
	assert expected[provider] in body


def test_busy_model_is_retried_then_replaced(server, config):
	config.update(aiProvider="gemini", aiUrl=server.url, aiModel="", aiApiKey="AQ.key")
	models = ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-3.8-flash-tts"]

	def reply(path, body, headers):
		if body is None:
			return 200, {
				"models": [
					{"name": f"models/{m}", "supportedGenerationMethods": ["generateContent"]} for m in models
				]
			}
		if "gemini-3.8-flash:" in path:
			return 503, {"error": {"status": "UNAVAILABLE"}}
		return 200, wrap("gemini", {"ok": True})

	server.reply = reply
	assert ai.testConnection() == "The AI service is working, model gemini-3.5-flash-lite"
	posts = [r[1] for r in server.requests if r[0] == "POST"]
	assert posts.count("/models/gemini-3.8-flash:generateContent") == 3


def test_unsupported_key_type_gives_a_clear_message(server, config):
	config.update(aiProvider="gemini", aiUrl=server.url, aiModel="gemini-x", aiApiKey="AQ.key")
	server.reply = lambda path, body, headers: (
		401,
		{"error": {"details": [{"reason": "ACCESS_TOKEN_TYPE_UNSUPPORTED"}]}},
	)
	with pytest.raises(ai.AIKeyError, match="did not accept this kind of key"):
		ai.testConnection()


def test_missing_key_asks_for_setup(config):
	config.update(aiProvider="anthropic", aiApiKey="")
	with pytest.raises(ai.AIKeyError):
		ai.testConnection()


def test_claude_models_prefer_sonnet_then_haiku():
	ranked = ai.rankModels(
		"anthropic", ["claude-opus-5-5", "claude-haiku-5-5", "claude-sonnet-4-5", "claude-sonnet-5-5"]
	)
	assert ranked[:3] == ["claude-sonnet-5-5", "claude-sonnet-4-5", "claude-haiku-5-5"]


def test_key_detection():
	assert ai.looksLikeProviderKey("gemini", " AQ.Ab8RN6abcdefghijklmnopqrstuvwxyz\n")
	assert ai.looksLikeProviderKey("anthropic", "sk-ant-api03-abcdefghijklmnopqrstu")
	assert not ai.looksLikeProviderKey("openaiApi", "sk-ant-api03-abcdefghijklmnopqrstu")
	assert ai.cleanKey('"abc"') == "abc"


def test_json_is_found_in_wrapped_answers():
	assert ai._parseJson('Sure!\n```json\n{"a": 1}\n```') == {"a": 1}
	assert ai._parseJson('answer: {"a": 2} done') == {"a": 2}


def test_paragraph_roles_are_checked(server, config):
	config.update(aiProvider="openai", aiUrl=server.url, aiModel="m")
	server.reply = lambda path, body, headers: (
		200,
		wrap(
			"openai",
			{"roles": [{"i": 0, "role": "subject"}, {"i": 1, "role": "bogus"}, {"i": 9, "role": "title"}]},
		),
	)
	roles = ai.classifyParagraphs(
		["เรื่อง ขอเชิญ", "text", ""], {"subject": "", "body": "", "empty": ""}, "letter"
	)
	assert roles == ["subject", "body", "empty"]


def test_keys_are_stored_through_the_key_store(config, monkeypatch):
	monkeypatch.setattr(keyStore, "_crypt", lambda data, protect: data[::-1])
	ai.saveKey("gemini", " AQ.secret ")
	assert config["aiApiKey"].startswith("dpapi:") and "secret" not in config["aiApiKey"]
	assert ai.currentKey() == "AQ.secret" and ai.storedKey("gemini") == "AQ.secret"
	config["aiKeyAnthropic"] = "sk-ant-plain"
	ai.protectStoredKeys()
	assert config["aiKeyAnthropic"].startswith("dpapi:")
	assert ai._conf()["aiKeyAnthropic"] == "sk-ant-plain"


def test_key_store_without_windows_keeps_working():
	assert keyStore.protect("abc") == "abc"
	assert keyStore.unprotect("abc") == "abc"
	assert keyStore.unprotect("dpapi:@@") == ""

# Word Document Assistant for NVDA
# Optional AI proofreading and suggestions. Nothing is changed in the document
# without the user choosing it.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from __future__ import annotations

import base64
import json
import re
import threading
import time
import urllib.error
import urllib.request
from collections.abc import Callable
from typing import Any

import addonHandler
import config
import wx
from logHandler import log

from . import keyStore
from .elements import CONFIG_SECTION

addonHandler.initTranslation()

PROVIDER_GEMINI = "gemini"
PROVIDER_ANTHROPIC = "anthropic"
PROVIDER_OPENAI_API = "openaiApi"
#: Any service with the OpenAI chat format at an address the user enters (Typhoon, Groq, OpenRouter, LM Studio...).
PROVIDER_OPENAI = "openai"
PROVIDER_OLLAMA = "ollama"
PROVIDERS = (PROVIDER_GEMINI, PROVIDER_ANTHROPIC, PROVIDER_OPENAI_API, PROVIDER_OPENAI, PROVIDER_OLLAMA)

# An empty model means: ask the service which models the key can use and pick one.
DEFAULT_MODELS = {provider: "" for provider in PROVIDERS}
DEFAULT_URLS = {
	PROVIDER_GEMINI: "https://generativelanguage.googleapis.com/v1beta",
	PROVIDER_ANTHROPIC: "https://api.anthropic.com/v1",
	PROVIDER_OPENAI_API: "https://api.openai.com/v1",
	PROVIDER_OPENAI: "",
	PROVIDER_OLLAMA: "http://localhost:11434",
}
#: Used when the model list cannot be read.
FALLBACK_MODELS = {
	PROVIDER_ANTHROPIC: "claude-sonnet-5-5",
}
#: Providers that need an API key.
KEY_REQUIRED = (PROVIDER_GEMINI, PROVIDER_ANTHROPIC, PROVIDER_OPENAI_API)
ANTHROPIC_VERSION = "2023-06-01"

#: Longest text sent in one request, in characters.
MAX_TEXT = 4000
TIMEOUT = 45


def providerNames() -> dict[str, str]:
	return {
		# Translators: AI provider name in the settings panel.
		PROVIDER_GEMINI: _("Google Gemini (free tier, or paid with billing turned on)"),
		# Translators: AI provider name in the settings panel.
		PROVIDER_ANTHROPIC: _("Anthropic Claude (paid, needs an API key)"),
		# Translators: AI provider name in the settings panel.
		PROVIDER_OPENAI_API: _("OpenAI (paid, needs an API key)"),
		# Translators: AI provider name in the settings panel.
		PROVIDER_OPENAI: _("OpenAI compatible service (enter the address)"),
		# Translators: AI provider name in the settings panel.
		PROVIDER_OLLAMA: _("Ollama on this computer (offline, text stays here)"),
	}


class AIError(Exception):
	"""A request to the AI service failed. The message is shown to the user."""


class AIKeyError(AIError):
	"""The API key is missing or was rejected; the user should set up a key."""


class AIBusy(AIError):
	"""The service is overloaded or briefly unavailable (HTTP 5xx); trying again may work."""


class AIQuota(AIBusy):
	"""The free quota of this model is used up (HTTP 429); another model may still work."""


class AIModelNotFound(AIError):
	"""The service does not know the model (HTTP 404)."""


#: Pages where the user signs in with their own account and gets a free key or the free app.
KEY_PAGES = {
	PROVIDER_GEMINI: "https://aistudio.google.com/apikey",
	PROVIDER_ANTHROPIC: "https://console.anthropic.com/settings/keys",
	PROVIDER_OPENAI_API: "https://platform.openai.com/api-keys",
	PROVIDER_OLLAMA: "https://ollama.com/download",
}

_GEMINI_KEY = re.compile(r"^AIza[0-9A-Za-z_\-]{30,}$")
#: Any single token that could be an API key. The service itself decides whether it is valid.
_ANY_KEY = re.compile(r"^[0-9A-Za-z_\-.:]{20,300}$")


def cleanKey(text: str) -> str:
	"""Remove spaces, line breaks and quotes that come along when a key is copied."""
	return (text or "").strip().strip("\"'`").strip()


def looksLikeGeminiKey(text: str) -> bool:
	"""Typical Google key; used only to notice a key copied from the browser."""
	return bool(_GEMINI_KEY.match(cleanKey(text)))


def looksLikeProviderKey(provider: str, text: str) -> bool:
	"""Typical key of this provider; used only to notice a key copied from the browser."""
	key = cleanKey(text)
	if provider == PROVIDER_GEMINI:
		return looksLikeGeminiKey(key) or (key.startswith("AQ.") and looksLikeKey(key))
	if provider == PROVIDER_ANTHROPIC:
		return key.startswith("sk-ant-") and looksLikeKey(key)
	if provider == PROVIDER_OPENAI_API:
		return key.startswith("sk-") and not key.startswith("sk-ant-") and looksLikeKey(key)
	return False


def looksLikeKey(text: str) -> bool:
	return bool(_ANY_KEY.match(cleanKey(text)))


#: Values tried by the key setup dialog before they are saved.
_override: dict[str, Any] = {}


class _Settings:
	"""Configuration values, with the key setup dialog's unsaved values taking priority."""

	def __getitem__(self, name: str) -> Any:
		if name in _override:
			return _override[name]
		value = config.conf[CONFIG_SECTION][name]
		if name == "aiApiKey" or name.startswith("aiKey"):
			return keyStore.unprotect(value)
		return value


def _conf() -> Any:
	return _Settings()


def _keySetting(provider: str) -> str:
	return "aiKey" + provider[0].upper() + provider[1:]


def storedKey(provider: str) -> str:
	"""The key saved for this service, if any."""
	try:
		conf = config.conf[CONFIG_SECTION]
		key = keyStore.unprotect(conf[_keySetting(provider)] or "")
		if not key and conf["aiProvider"] == provider:
			# Keys saved before each service had its own key.
			key = keyStore.unprotect(conf["aiApiKey"] or "")
		return key
	except KeyError:
		return ""


def currentKey() -> str:
	"""The key of the selected service, decrypted."""
	try:
		return keyStore.unprotect(config.conf[CONFIG_SECTION]["aiApiKey"] or "")
	except KeyError:
		return ""


def rememberKey(provider: str, key: str) -> None:
	"""Store the key of a service, encrypted when possible."""
	try:
		config.conf[CONFIG_SECTION][_keySetting(provider)] = keyStore.protect(cleanKey(key))
	except KeyError:
		pass


def saveKey(provider: str, key: str) -> None:
	"""Store the key as the current key and as the key of its service."""
	config.conf[CONFIG_SECTION]["aiApiKey"] = keyStore.protect(cleanKey(key))
	rememberKey(provider, key)


def protectStoredKeys() -> None:
	"""Encrypt keys saved by earlier versions in plain text."""
	try:
		conf = config.conf[CONFIG_SECTION]
		for name in ("aiApiKey", *(_keySetting(p) for p in PROVIDERS)):
			value = conf[name] or ""
			if value and not keyStore.isProtected(value):
				protected = keyStore.protect(value)
				if protected != value:
					conf[name] = protected
	except KeyError:
		log.debugWarning("Cannot protect stored keys", exc_info=True)


def isEnabled() -> bool:
	return bool(_conf()["aiEnabled"])


def isCloud() -> bool:
	"""True when text leaves this computer."""
	if _conf()["aiProvider"] != PROVIDER_OLLAMA:
		return True
	url = _setting("aiUrl", DEFAULT_URLS).lower()
	return not any(host in url for host in ("localhost", "127.0.0.1", "[::1]"))


def _setting(name: str, defaults: dict[str, str]) -> str:
	value = (_conf()[name] or "").strip()
	return value or defaults.get(_conf()["aiProvider"], "")


def _post(url: str, body: dict | None, headers: dict[str, str]) -> dict:
	"""POST body as JSON, or GET when body is None. Returns the JSON answer."""
	request = urllib.request.Request(
		url,
		data=None if body is None else json.dumps(body).encode("utf-8"),
		headers={"Content-Type": "application/json", **headers},
		method="GET" if body is None else "POST",
	)
	try:
		with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
			return json.loads(response.read().decode("utf-8"))
	except urllib.error.HTTPError as error:
		detail = ""
		try:
			detail = error.read().decode("utf-8", "replace")[:1500]
		except Exception:  # noqa: BLE001
			pass
		log.debugWarning(f"AI service returned {error.code}: {detail}")
		if "ACCESS_TOKEN_TYPE_UNSUPPORTED" in detail.upper():
			raise AIKeyError(
				# Translators: Reported when Google does not accept the new kind of key for this account.
				_(
					"Google did not accept this kind of key for your account. Create a new key in "
					"Google AI Studio, or use Ollama on this computer instead",
				),
			) from error
		if error.code == 404:
			raise AIModelNotFound(
				# Translators: Reported when the AI service does not know the model.
				_(
					"The AI service does not have this model. Leave the model empty to choose one automatically",
				),
			) from error
		if error.code in (401, 403) or (error.code == 400 and "API_KEY" in detail.upper()):
			# Translators: Reported when the AI service rejects the API key.
			raise AIKeyError(_("The AI service rejected the API key. Check it in the settings")) from error
		if error.code in (500, 502, 503, 504):
			raise AIBusy(
				# Translators: Reported when the AI service is overloaded and retries did not help.
				_("The AI service is busy right now. Wait a minute and try again"),
			) from error
		if error.code == 429:
			# Translators: Reported when the free quota of the AI service is used up.
			raise AIQuota(_("The AI service limit was reached. Wait a moment and try again")) from error
		# Translators: Reported when the AI service returns an error; {code} is the HTTP status.
		raise AIError(_("The AI service returned error {code}").format(code=error.code)) from error
	except (urllib.error.URLError, TimeoutError, OSError) as error:
		log.debugWarning("Cannot reach the AI service", exc_info=True)
		# Translators: Reported when the AI service cannot be reached.
		raise AIError(
			_("Cannot reach the AI service. Check the internet connection and the address"),
		) from error
	except ValueError as error:
		# Translators: Reported when the AI service sends an answer the add-on cannot read.
		raise AIError(_("The AI service sent an answer that could not be read")) from error


def _complete(system: str, user: str, image: tuple[bytes, str] | None = None) -> str:
	"""Send one request and return the text of the answer."""
	provider = _conf()["aiProvider"]
	model = _setting("aiModel", DEFAULT_MODELS)
	baseUrl = _setting("aiUrl", DEFAULT_URLS).rstrip("/")
	apiKey = cleanKey(_conf()["aiApiKey"])
	if provider in KEY_REQUIRED and not apiKey:
		# Translators: Reported when the AI API key has not been entered.
		raise AIKeyError(_("Enter an API key in the Word Document Assistant settings first"))
	if not baseUrl:
		# Translators: Reported when the AI address or model has not been entered.
		raise AIError(_("Enter the AI service address and model in the settings first"))
	try:
		if model:
			return _withRetry(lambda: _generate(provider, baseUrl, apiKey, model, system, user, image))
		return _autoComplete(provider, baseUrl, apiKey, system, user, image)
	except (KeyError, IndexError, TypeError) as error:
		log.debugWarning("Unexpected AI answer", exc_info=True)
		raise AIError(_("The AI service sent an answer that could not be read")) from error


# Requests for each kind of service ---------------------------------------------


def _bearer(apiKey: str) -> dict[str, str]:
	return {"Authorization": f"Bearer {apiKey}"} if apiKey else {}


def _anthropicHeaders(apiKey: str) -> dict[str, str]:
	return {"x-api-key": apiKey, "anthropic-version": ANTHROPIC_VERSION}


def _generate(
	provider: str,
	baseUrl: str,
	apiKey: str,
	model: str,
	system: str,
	user: str,
	image: tuple[bytes, str] | None = None,
) -> str:
	imageData, mime = (base64.b64encode(image[0]).decode("ascii"), image[1]) if image else ("", "")
	if provider == PROVIDER_GEMINI:
		data = _post(
			f"{baseUrl}/models/{model}:generateContent",
			{
				"systemInstruction": {"parts": [{"text": system}]},
				"contents": [
					{
						"role": "user",
						"parts": ([{"inlineData": {"mimeType": mime, "data": imageData}}] if image else [])
						+ [{"text": user}],
					},
				],
				"generationConfig": {"temperature": 0.2, "responseMimeType": "application/json"},
			},
			{"x-goog-api-key": apiKey},
		)
		return "".join(part.get("text", "") for part in data["candidates"][0]["content"]["parts"])
	if provider == PROVIDER_ANTHROPIC:
		data = _post(
			f"{baseUrl}/messages",
			{
				"model": model,
				"max_tokens": 4096,
				"system": system,
				"messages": [
					{
						"role": "user",
						"content": (
							[
								{
									"type": "image",
									"source": {"type": "base64", "media_type": mime, "data": imageData},
								},
								{"type": "text", "text": user},
							]
							if image
							else user
						),
					},
				],
			},
			_anthropicHeaders(apiKey),
		)
		return "".join(block.get("text", "") for block in data["content"] if block.get("type") == "text")
	messages: list[dict[str, Any]] = [
		{"role": "system", "content": system},
		{"role": "user", "content": user},
	]
	if provider == PROVIDER_OLLAMA:
		if image:
			messages[1]["images"] = [imageData]
		data = _post(
			f"{baseUrl}/api/chat",
			{
				"model": model,
				"messages": messages,
				"stream": False,
				"format": "json",
				"options": {"temperature": 0.2},
			},
			{},
		)
		return data["message"]["content"]
	if image:
		messages[1]["content"] = [
			{"type": "text", "text": user},
			{"type": "image_url", "image_url": {"url": f"data:{mime};base64,{imageData}"}},
		]
	body: dict[str, Any] = {"model": model, "messages": messages}
	if provider == PROVIDER_OPENAI_API:
		# Newer OpenAI models only accept the default temperature.
		body["response_format"] = {"type": "json_object"}
	else:
		body["temperature"] = 0.2
	data = _post(f"{baseUrl}/chat/completions", body, _bearer(apiKey))
	return data["choices"][0]["message"]["content"]


def listModels(provider: str, baseUrl: str, apiKey: str) -> list[str]:
	"""Models this key can use for text, as the service reports them."""
	baseUrl = (baseUrl or DEFAULT_URLS.get(provider, "")).rstrip("/")
	if provider == PROVIDER_GEMINI:
		return listGeminiModels(baseUrl, apiKey)
	if provider == PROVIDER_OLLAMA:
		return listOllamaModels(baseUrl)
	if provider == PROVIDER_ANTHROPIC:
		data = _post(f"{baseUrl}/models?limit=1000", None, _anthropicHeaders(apiKey))
	else:
		data = _post(f"{baseUrl}/models", None, _bearer(apiKey))
	return [str(m["id"]) for m in data.get("data", []) if isinstance(m, dict) and m.get("id")]


#: Model chosen automatically for each service and key, so the list is not fetched every time.
_autoModel: dict[tuple[str, str], str] = {}
#: Models the last key could use, shown in the setup dialog.
lastModels: list[str] = []


def listGeminiModels(baseUrl: str, apiKey: str) -> list[str]:
	"""Models this key can use to generate text, without the "models/" prefix."""
	names: list[str] = []
	pageToken = ""
	for _page in range(5):
		url = f"{baseUrl}/models?pageSize=200" + (f"&pageToken={pageToken}" if pageToken else "")
		data = _post(url, None, {"x-goog-api-key": apiKey})
		for model in data.get("models", []):
			methods = model.get("supportedGenerationMethods", [])
			name = str(model.get("name", "")).removeprefix("models/")
			if name and "generateContent" in methods:
				names.append(name)
		pageToken = data.get("nextPageToken", "")
		if not pageToken:
			break
	return names


_SKIP_WORDS = (
	"tts",
	"image",
	"live",
	"audio",
	"embedding",
	"vision",
	"robotics",
	"thinking",
	"exp",
	"learnlm",
	"realtime",
	"transcribe",
	"search",
	"moderation",
	"instruct",
	"codex",
	"whisper",
	"dall",
	"babbage",
	"davinci",
)


def _versionKey(name: str) -> tuple:
	"""Numbers in the model name, for example (3, 8) for gemini-3.8-flash or (5, 5) for claude-sonnet-5-5."""
	return tuple(int(n) for n in re.findall(r"\d+", name) if len(n) < 6) or (0,)


def _ordered(groups: list[list[str]], names: list[str]) -> list[str]:
	seen: list[str] = []
	for name in [n for group in groups for n in sorted(group, key=_versionKey, reverse=True)] + names:
		if name not in seen:
			seen.append(name)
	return seen


def rankGeminiModels(names: list[str]) -> list[str]:
	"""Newest stable Flash first, then Flash-Lite, previews, then the rest."""
	usable = [n for n in names if n.startswith("gemini") and not any(w in n for w in _SKIP_WORDS)]
	return (
		_ordered(
			[
				[n for n in usable if "flash" in n and "lite" not in n and "preview" not in n],
				[n for n in usable if "flash" in n and "lite" in n and "preview" not in n],
				[n for n in usable if "flash" in n and "preview" in n],
				[n for n in usable if "flash" not in n],
			],
			[],
		)
		or names
	)


def rankModels(provider: str, names: list[str]) -> list[str]:
	"""Models in order of preference for proofreading documents."""
	if provider == PROVIDER_GEMINI:
		return rankGeminiModels(names)
	if provider == PROVIDER_ANTHROPIC:
		# Sonnet balances quality for Thai and English with cost; Haiku is cheaper, Opus the strongest.
		claude = [n for n in names if n.startswith("claude")]
		return _ordered(
			[
				[n for n in claude if "sonnet" in n],
				[n for n in claude if "haiku" in n],
				[n for n in claude if "opus" in n],
			],
			claude,
		)
	if provider == PROVIDER_OPENAI_API:
		chat = [n for n in names if n.startswith("gpt-") and not any(w in n for w in _SKIP_WORDS)]
		return _ordered(
			[
				[n for n in chat if not any(w in n for w in ("mini", "nano", "preview"))],
				[n for n in chat if "mini" in n and "preview" not in n],
				[n for n in chat if "nano" in n or "preview" in n],
			],
			[],
		)
	if provider == PROVIDER_OLLAMA:
		return _ordered(
			[
				[n for n in names if "typhoon" in n],
				[n for n in names if "qwen" in n],
				[n for n in names if "gemma" in n],
				[n for n in names if "llama" in n],
			],
			names,
		)
	return names


def pickGeminiModel(names: list[str]) -> str:
	ranked = rankGeminiModels(names)
	return ranked[0] if ranked else ""


#: Seconds to wait before each new try when the service is busy.
RETRY_DELAYS = (2, 5)


def _withRetry(func: Callable[[], str]) -> str:
	"""Call func, trying again after a short wait when the service is busy (not when the quota is used up)."""
	for delay in (*RETRY_DELAYS, None):
		try:
			return func()
		except AIQuota:
			raise
		except AIBusy:
			if delay is None:
				raise
			log.debugWarning(f"AI service busy, trying again in {delay} s")
			time.sleep(delay)
	raise AIError(_("The AI request failed"))


def _autoComplete(
	provider: str,
	baseUrl: str,
	apiKey: str,
	system: str,
	user: str,
	image: tuple[bytes, str] | None = None,
) -> str:
	"""Pick a model automatically; move to the next one if it is busy, out of quota or retired."""
	key = (provider, apiKey)
	if key not in _autoModel or not lastModels:
		try:
			names = listModels(provider, baseUrl, apiKey)
		except AIKeyError:
			raise
		except AIError:
			if provider not in FALLBACK_MODELS:
				raise
			names = []
		lastModels[:] = names
	ranked = rankModels(provider, lastModels) or (
		[FALLBACK_MODELS[provider]] if provider in FALLBACK_MODELS else []
	)
	if not ranked:
		# Translators: Reported when the key cannot use any text model.
		raise AIError(_("No text model is available for this key. Enter a model in the settings"))
	saved = _autoModel.get(key)
	current = saved if saved is not None and saved in ranked else ranked[0]
	candidates = [current] + [m for m in ranked if m != current][:3]
	lastError: AIError | None = None
	for candidate in candidates:
		try:
			result = _withRetry(
				lambda m=candidate: _generate(provider, baseUrl, apiKey, m, system, user, image),
			)
		except (AIBusy, AIModelNotFound) as error:
			log.debugWarning(f"AI model {candidate} failed: {error}")
			lastError = error
			continue
		if _autoModel.get(key) != candidate:
			log.info(f"Word Document Assistant: using {provider} model {candidate}")
			_autoModel[key] = candidate
		return result
	raise lastError or AIError(_("The AI request failed"))


def currentModel() -> str:
	"""The model in use, for messages."""
	model = _setting("aiModel", DEFAULT_MODELS)
	if not model:
		model = _autoModel.get((_conf()["aiProvider"], cleanKey(_conf()["aiApiKey"])), "")
	return model


def _parseJson(text: str) -> Any:
	text = text.strip()
	# Some models wrap the JSON in a Markdown code block or add a sentence around it.
	fenced = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL)
	if fenced:
		text = fenced.group(1).strip()
	for candidate in (text, _between(text, "{", "}"), _between(text, "[", "]")):
		if not candidate:
			continue
		try:
			return json.loads(candidate)
		except ValueError:
			continue
	raise AIError(_("The AI service sent an answer that could not be read"))


def _between(text: str, start: str, end: str) -> str:
	first, last = text.find(start), text.rfind(end)
	return text[first : last + 1] if 0 <= first < last else ""


def _languageHint() -> str:
	language = _conf()["aiLanguage"]
	if language == "th":
		return "The text is in Thai. Answer in Thai."
	if language == "en":
		return "The text is in English. Answer in English."
	return "Answer in the same language as the text (Thai or English)."


_RULES = (
	"Never invent facts. Keep the author's meaning, names, numbers and terminology. "
	"Return only JSON, with no explanation outside the JSON."
)


def proofread(text: str) -> list[dict[str, str]]:
	"""Return proposed corrections: dicts with original, suggestion and reason."""
	system = (
		"You are a careful proofreader for Thai and English documents. "
		"Find spelling mistakes, wrong Thai spellings (for example confusing tone marks, "
		"การันต์, ร/ล, ศ/ษ/ส), typing errors, wrong word choice and clear grammar errors. "
		"Do not rewrite style. Each 'original' must be copied exactly from the text, "
		"short (a word or phrase), and appear in the text. "
		f"{_languageHint()} {_RULES} "
		'Format: {"corrections": [{"original": "...", "suggestion": "...", "reason": "..."}]}. '
		'If there is nothing to correct return {"corrections": []}.'
	)
	data = _parseJson(_complete(system, text[:MAX_TEXT]))
	corrections = data.get("corrections", []) if isinstance(data, dict) else data
	result = []
	for entry in corrections if isinstance(corrections, list) else []:
		if not isinstance(entry, dict):
			continue
		original = str(entry.get("original", "")).strip()
		suggestion = str(entry.get("suggestion", "")).strip()
		if original and suggestion and original != suggestion and original in text:
			result.append(
				{
					"original": original,
					"suggestion": suggestion,
					"reason": str(entry.get("reason", "")).strip(),
				},
			)
	return result


def suggestWord(word: str, context: str) -> list[dict[str, str]]:
	"""Return correction candidates for one misspelled word."""
	system = (
		"You help correct one word that a spell checker marked as misspelled. "
		"Give up to 5 likely correct spellings for that word in its context, best first. "
		"If the word is actually correct (for example a name or a technical term), "
		"include it unchanged as one option and say so in the reason. "
		f"{_languageHint()} {_RULES} "
		'Format: {"suggestions": [{"text": "...", "reason": "..."}]}.'
	)
	user = f"Word: {word}\nContext: {context}"
	data = _parseJson(_complete(system, user[:MAX_TEXT]))
	entries = data.get("suggestions", []) if isinstance(data, dict) else data
	result = []
	for entry in entries if isinstance(entries, list) else []:
		if isinstance(entry, str):
			entry = {"text": entry}
		if isinstance(entry, dict) and str(entry.get("text", "")).strip():
			result.append(
				{"text": str(entry["text"]).strip(), "reason": str(entry.get("reason", "")).strip()},
			)
	return result[:5]


def draftReplies(author: str, comment: str, commentedText: str) -> list[str]:
	"""Return short reply drafts for a comment."""
	system = (
		"You help a document author reply to a review comment in Microsoft Word. "
		"Write 3 different short, polite, professional replies the author could send, "
		"for example agreeing and saying what was changed, asking for clarification, "
		"or explaining why the text stays as it is. "
		f"{_languageHint()} {_RULES} "
		'Format: {"replies": ["...", "...", "..."]}.'
	)
	user = f"Comment by {author}: {comment}\nText the comment refers to: {commentedText}"
	data = _parseJson(_complete(system, user[:MAX_TEXT]))
	replies = data.get("replies", []) if isinstance(data, dict) else data
	return [str(r).strip() for r in replies if str(r).strip()][:5] if isinstance(replies, list) else []


def recommendFormat(sample: str, options: list[tuple[str, str]]) -> dict[str, str]:
	"""Suggest a format for the document. options are (id, description) pairs."""
	listing = "\n".join(f"- {presetId}: {description}" for presetId, description in options)
	system = (
		"You help format Microsoft Word documents. Read the beginning of a document and choose "
		"the most suitable format from the list. Thai government letters start with ที่ and "
		"ส่วนราชการ and contain เรื่อง and เรียน; internal memos start with บันทึกข้อความ; "
		"announcements and orders start with ประกาศ or คำสั่ง; minutes contain รายงานการประชุม "
		f"or ระเบียบวาระ. {_languageHint()} {_RULES} "
		'Format: {"format": "<id from the list>", "reason": "<one short sentence>"}.'
	)
	user = f"Formats:\n{listing}\n\nBeginning of the document:\n{sample[:2500]}"
	data = _parseJson(_complete(system, user))
	ids = {presetId for presetId, _description in options}
	choice = str(data.get("format", "")).strip() if isinstance(data, dict) else ""
	if choice not in ids:
		# Translators: Reported when AI could not recommend a format.
		raise AIError(_("AI could not recommend a format. Choose one from the list"))
	return {"format": choice, "reason": str(data.get("reason", "")).strip()}


#: Paragraphs sent in one request when classifying, and characters of each paragraph.
CLASSIFY_BATCH = 120
CLASSIFY_CHARS = 160


def classifyParagraphs(
	texts: list[str],
	roleHints: dict[str, str],
	formatName: str,
	progress: Callable[[int, int], None] | None = None,
) -> list[str]:
	"""Give each paragraph a role id. The text of the document is not changed."""
	roleList = "\n".join(f"- {roleId}: {hint}" for roleId, hint in roleHints.items())
	system = (
		f"You are preparing a Microsoft Word document in the format: {formatName}. "
		"For every numbered paragraph decide its role, using only these role ids:\n"
		f"{roleList}\n"
		"Use body when unsure. Use empty only for paragraphs with no text. "
		"Do not change, rewrite or return the text. "
		f"{_RULES} "
		'Format: {"roles": [{"i": <number>, "role": "<role id>"}, ...]} with one entry per paragraph.'
	)
	allowed = set(roleHints)
	roles = ["body"] * len(texts)
	for start in range(0, len(texts), CLASSIFY_BATCH):
		batch = texts[start : start + CLASSIFY_BATCH]
		lines = []
		for offset, text in enumerate(batch):
			flat = " ".join(text.split())
			lines.append(f"[{start + offset}] {flat[:CLASSIFY_CHARS] if flat else '(empty)'}")
		data = _parseJson(_complete(system, "\n".join(lines)))
		entries = data.get("roles", []) if isinstance(data, dict) else data
		for entry in entries if isinstance(entries, list) else []:
			if not isinstance(entry, dict):
				continue
			try:
				index = int(entry.get("i", ""))
			except (TypeError, ValueError):
				continue
			role = str(entry.get("role", "")).strip()
			if start <= index < start + len(batch) and role in allowed:
				roles[index] = role
		if progress:
			progress(min(start + len(batch), len(texts)), len(texts))
	# Paragraphs without text are always empty, whatever the AI said.
	for index, text in enumerate(texts):
		if not text.strip():
			roles[index] = "empty" if "empty" in allowed else roles[index]
		elif roles[index] == "empty":
			roles[index] = "body"
	return roles


#: Longest text sent for summaries and other whole-document tasks, in characters.
MAX_LONG_TEXT = 60000

_TEXT_FORMAT = 'Format: {"text": "..."}. Use \\n for line breaks inside the text.'


def _textAnswer(system: str, user: str, image: tuple[bytes, str] | None = None) -> str:
	data = _parseJson(_complete(system, user, image))
	text = str(data.get("text", "")).strip() if isinstance(data, dict) else str(data).strip()
	if not text:
		raise AIError(_("The AI service sent an answer that could not be read"))
	return text


def describeImage(image: bytes, mime: str, context: str, detailed: bool) -> str:
	"""Alternative text for a picture in a document."""
	if detailed:
		task = (
			"Write a detailed description of the picture for a blind reader: what it shows, any text in it "
			"(copied exactly), and for charts or diagrams the type, axes, values and the main message. "
			"Use short paragraphs."
		)
	else:
		task = (
			"Write alternative text for the picture for a screen reader user: one or two sentences, at most "
			"about 250 characters, saying what matters in this document. Do not start with 'image of' or "
			"'picture of'. Copy any important text in the picture exactly. For a chart, give its main message."
		)
	system = (
		f"You describe pictures in Microsoft Word documents for blind and low vision readers. {task} "
		"Use the nearby document text only to understand the purpose of the picture; do not guess names "
		f"of people you cannot identify from the picture or the text. {_languageHint()} {_RULES} {_TEXT_FORMAT}"
	)
	return _textAnswer(system, f"Text near the picture in the document: {context[:1500]}", (image, mime))


def describeTable(tableText: str, caption: str) -> str:
	"""A summary of a table to use as its alternative text."""
	system = (
		"You write the alternative text of a table in a Microsoft Word document for screen reader users: "
		"in two to four sentences say what the table is about, what its columns are, and its most important "
		f"figures or trend. {_languageHint()} {_RULES} {_TEXT_FORMAT}"
	)
	user = f"Caption or text before the table: {caption[:500]}\nTable (cells separated by tabs):\n{tableText[:MAX_TEXT]}"
	return _textAnswer(system, user)


def summarize(text: str, length: str = "short") -> str:
	"""A summary of the text. length is short, detailed or keyPoints."""
	tasks = {
		"short": "Summarize it in one short paragraph.",
		"detailed": "Summarize it section by section, with a short heading line for each section.",
		"keyPoints": "List its key points, decisions, dates, amounts and actions requested, one per line, each starting with '- '.",
	}
	system = (
		"You help a blind person understand a long Microsoft Word document quickly. "
		f"{tasks.get(length, tasks['short'])} Keep names, numbers and dates exactly as written. "
		f"{_languageHint()} {_RULES} {_TEXT_FORMAT}"
	)
	return _textAnswer(system, text[:MAX_LONG_TEXT])


REWRITE_STYLES = ("official", "polite", "concise", "easy", "fix")


def rewrite(text: str, style: str) -> str:
	"""The text rewritten in a style. The meaning, names and numbers stay the same."""
	tasks = {
		"official": "Rewrite it in formal Thai government style (ภาษาราชการ) as used in official letters, or formal English if the text is English.",
		"polite": "Rewrite it to be more polite and respectful, keeping it professional.",
		"concise": "Rewrite it shorter and clearer, removing repetition.",
		"easy": "Rewrite it in plain, easy-to-read language with short sentences (Easy Read), for people with reading or learning difficulties.",
		"fix": "Correct spelling, typing errors and grammar only. Do not change the style.",
	}
	system = (
		"You edit text from a Microsoft Word document. "
		f"{tasks.get(style, tasks['fix'])} Keep the meaning, names, numbers, dates and paragraph breaks. "
		f"Return only the rewritten text. {_RULES} {_TEXT_FORMAT}"
	)
	return _textAnswer(system, text[: MAX_TEXT * 3])


def translate(text: str, target: str) -> str:
	"""The text translated to Thai (th) or English (en)."""
	language = "Thai" if target == "th" else "English"
	system = (
		f"Translate the text from a Microsoft Word document into natural, correct {language}. Keep names, "
		"numbers, dates and paragraph breaks. For Thai official text use the usual official terms. "
		f"Return only the translation. {_RULES} {_TEXT_FORMAT}"
	)
	return _textAnswer(system, text[: MAX_TEXT * 3])


def summarizeComments(comments: str) -> str:
	"""What reviewers asked for, grouped, with what is still open."""
	system = (
		"You read the review comments of a Microsoft Word document. Summarize what the reviewers ask for, "
		"grouped by topic, then list the open (not resolved) requests the author still has to handle, "
		f"one per line starting with '- '. Mention who asked. {_languageHint()} {_RULES} {_TEXT_FORMAT}"
	)
	return _textAnswer(system, comments[:MAX_LONG_TEXT])


def checkConsistency(text: str) -> str:
	"""Inconsistencies in the text, as a list."""
	system = (
		"Check a Microsoft Word document for inconsistencies only: the same word or name spelled in different "
		"ways, Thai digits mixed with Arabic digits, Buddhist and Christian era years mixed, dates or amounts that "
		"contradict each other, inconsistent terms or abbreviations, and numbering that skips. "
		"List each problem on its own line starting with '- ', quoting the text so it can be found. "
		f"If there is nothing, say so in one line. {_languageHint()} {_RULES} {_TEXT_FORMAT}"
	)
	return _textAnswer(system, text[:MAX_LONG_TEXT])


def testWith(values: dict[str, Any]) -> str:
	"""Test unsaved settings. Runs in the background thread."""
	_override.clear()
	_override.update(values)
	try:
		return testConnection()
	finally:
		_override.clear()


def listOllamaModels(baseUrl: str = "") -> list[str]:
	"""Names of the models installed in Ollama on this computer."""
	url = (baseUrl or DEFAULT_URLS[PROVIDER_OLLAMA]).rstrip("/") + "/api/tags"
	try:
		with urllib.request.urlopen(url, timeout=5) as response:
			data = json.loads(response.read().decode("utf-8"))
	except (urllib.error.URLError, TimeoutError, OSError, ValueError) as error:
		# Translators: Reported when Ollama is not running on this computer.
		raise AIError(
			_("Ollama was not found on this computer. Install it, then start it and try again"),
		) from error
	return [m["name"] for m in data.get("models", []) if isinstance(m, dict) and m.get("name")]


def testConnection() -> str:
	system = f'Reply with {{"ok": true}}. {_RULES}'
	_parseJson(_complete(system, "ping"))
	model = currentModel()
	if model:
		# Translators: Reported when the AI connection test works; {model} is the model name.
		return _("The AI service is working, model {model}").format(model=model)
	# Translators: Reported when the AI connection test works.
	return _("The AI service is working")


def runInBackground(
	work: Callable[[], Any],
	onDone: Callable[[Any], None],
	onError: Callable[[str, bool], None],
) -> None:
	"""Run a network request without freezing NVDA; callbacks run on the main thread."""

	def target():
		try:
			result = work()
		except AIError as error:
			wx.CallAfter(onError, str(error), isinstance(error, AIKeyError))
		except Exception:
			log.error("AI request failed", exc_info=True)
			# Translators: Reported when the AI request fails for an unknown reason.
			wx.CallAfter(onError, _("The AI request failed"), False)
		else:
			wx.CallAfter(onDone, result)

	threading.Thread(target=target, name="wordDocAssistant.ai", daemon=True).start()

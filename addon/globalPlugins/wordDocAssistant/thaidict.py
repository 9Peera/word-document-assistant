# Word Document Assistant for NVDA
# Offline Thai word check: open word lists, commonly misspelled words, updates.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

"""Thai word check that works without AI or internet.

Word lists come from sources that allow reuse (see data/README.md). Updates are only
looked for and downloaded from the addresses in SOURCES, or from an address the
organization sets for its own list of commonly misspelled words.
"""

from __future__ import annotations

import bisect
import json
import os
import re
import threading
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Any

import addonHandler
from logHandler import log

addonHandler.initTranslation()

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
BUNDLED_WORDS = os.path.join(DATA_DIR, "words_th.txt")
BUNDLED_MISSPELLINGS = os.path.join(DATA_DIR, "misspellings_th.tsv")
MAX_WORD = 30
TIMEOUT = 20
CHECK_INTERVAL = 7 * 24 * 3600


@dataclass(frozen=True)
class WordSource:
	id: str
	name: str
	url: str
	license: str
	homepage: str
	kind: str = "words"  # words, hunspell or misspellings
	bundledEtag: str = ""


def sources() -> list[WordSource]:
	return [
		WordSource(
			"pythainlp",
			# Translators: Name of a Thai word list.
			_("PyThaiNLP Thai word list"),
			"https://raw.githubusercontent.com/PyThaiNLP/pythainlp/dev/pythainlp/corpus/words_th.txt",
			"CC0 1.0",
			"https://github.com/PyThaiNLP/pythainlp",
			bundledEtag='"59162035b39b3b618f2b785993c20f147d0964a05edfa36f500cc143291d9401"',
		),
		WordSource(
			"libreoffice",
			# Translators: Name of a Thai word list.
			_("LibreOffice Thai dictionary (NECTEC LEXiTRON)"),
			"https://raw.githubusercontent.com/LibreOffice/dictionaries/master/th_TH/th_TH.dic",
			"LGPL",
			"https://github.com/LibreOffice/dictionaries/tree/master/th_TH",
			kind="hunspell",
		),
	]


# Thai characters that cannot start a word: vowels and marks written above, below or after a consonant.
_NO_START = set("ะัาำิีึืฺุู็่้๊๋์ํ๎ๅ")
# Vowels written before the consonant: a word cannot end right after them.
_NO_END = set("เแโใไ")
_THAI_RUN = re.compile(r"[ก-๎]+")
_CONSONANT = re.compile(r"[ก-ฮ]")
#: Words before a person's name; the unknown word after them is not reported.
NAME_PREFIXES = ("นาย", "นาง", "นางสาว", "เด็กชาย", "เด็กหญิง", "ด.ช.", "ด.ญ.", "ดร.", "คุณ", "ศ.", "รศ.", "ผศ.")


@dataclass
class Finding:
	start: int
	end: int
	word: str
	kind: str  # "misspelled" or "unknown"
	suggestions: list[str] = field(default_factory=list)
	note: str = ""


class Dictionary:
	"""Known words and commonly misspelled words."""

	def __init__(self) -> None:
		self.words: set[str] = set()
		self.misspellings: dict[str, tuple[str, str]] = {}
		self.ignored: set[str] = set()
		self.alphabet: set[str] = set()
		self._sorted: list[str] = []

	def load(
		self,
		wordFiles: list[tuple[str, str]],
		misspellingFiles: list[str],
		myWordsFile: str = "",
	) -> None:
		words: set[str] = set()
		for path, kind in wordFiles:
			words |= readWords(path, kind)
		misspellings: dict[str, tuple[str, str]] = {}
		for path in misspellingFiles:
			misspellings.update(readMisspellings(path))
		if myWordsFile:
			words |= readWords(myWordsFile, "words")
		# The collected word list contains common misspellings; they must not count as correct.
		words -= set(misspellings)
		self.words = {w for w in words if len(w) <= MAX_WORD}
		self.misspellings = misspellings
		self.alphabet = {ch for w in self.words for ch in w if "ก" <= ch <= "๎"}
		self._sorted = sorted(self.words)

	@property
	def loaded(self) -> bool:
		return bool(self.words)

	# Checking --------------------------------------------------------------------

	def check(self, text: str, unknown: bool = True) -> list[Finding]:
		findings: list[Finding] = []
		for finding in sorted(self._misspelled(text), key=lambda f: (f.start, -(f.end - f.start))):
			if not findings or finding.start >= findings[-1].end:
				findings.append(finding)
		taken = [(f.start, f.end) for f in findings]
		if unknown:
			for finding in self._unknown(text):
				if not any(a < finding.end and finding.start < b for a, b in taken):
					findings.append(finding)
		findings.sort(key=lambda f: f.start)
		return findings

	def _misspelled(self, text: str) -> list[Finding]:
		found = []
		for wrong, (right, note) in self.misspellings.items():
			if wrong in self.ignored:
				continue
			position = text.find(wrong)
			while position >= 0:
				end = position + len(wrong)
				if not self._insideLongerWord(text, position, end):
					found.append(Finding(position, end, wrong, "misspelled", [right], note))
				position = text.find(wrong, end)
		return found

	def _insideLongerWord(self, text: str, start: int, end: int) -> bool:
		"""True when a correct, longer word covers this place, such as ชลอม around ชลอ."""
		for left in range(max(0, end - MAX_WORD), start + 1):
			for right in range(end, min(len(text), left + MAX_WORD) + 1):
				if (right - left) > (end - start) and text[left:right] in self.words:
					return True
		return False

	def segment(self, run: str) -> list[tuple[int, int, bool]]:
		"""Split a run of Thai letters into (start, end, known) with as few unknown letters as possible."""
		length = len(run)
		boundaries = [
			i == 0 or i == length or (run[i] not in _NO_START and run[i - 1] not in _NO_END)
			for i in range(length + 1)
		]
		inf = (length + 1, length + 1)
		best: list[tuple[int, int]] = [inf] * (length + 1)
		back: list[tuple[int, bool]] = [(0, False)] * (length + 1)
		best[0] = (0, 0)
		for i in range(length):
			if best[i] == inf or not boundaries[i]:
				continue
			unknownCost, tokens = best[i]
			for j in range(i + 1, min(length, i + MAX_WORD) + 1):
				if boundaries[j] and run[i:j] in self.words:
					candidate = (unknownCost, tokens + 1)
					if candidate < best[j]:
						best[j], back[j] = candidate, (i, True)
			# One letter cluster as unknown.
			j = i + 1
			while j < length and not boundaries[j]:
				j += 1
			candidate = (unknownCost + (j - i), tokens + 1)
			if candidate < best[j]:
				best[j], back[j] = candidate, (i, False)
		pieces = []
		position = length
		while position > 0:
			start, known = back[position]
			pieces.append((start, position, known))
			position = start
		pieces.reverse()
		merged: list[tuple[int, int, bool]] = []
		for piece in pieces:
			if merged and not piece[2] and not merged[-1][2]:
				merged[-1] = (merged[-1][0], piece[1], False)
			else:
				merged.append(piece)
		return merged

	def _unknown(self, text: str) -> list[Finding]:
		found = []
		for match in _THAI_RUN.finditer(text):
			run = match.group(0)
			pieces = self.segment(run)
			used = -1
			for index, (start, end, known) in enumerate(pieces):
				word = run[start:end]
				if (
					known
					or index <= used
					or len(word) < 2
					or not _CONSONANT.search(word)
					or word in self.ignored
				):
					continue
				before = (
					run[pieces[index - 1][0] : pieces[index - 1][1]]
					if index
					else text[max(0, match.start() - 8) : match.start()].strip()
				)
				if before.endswith(NAME_PREFIXES):
					continue
				first, last, suggestions = self._bestSpan(run, pieces, index)
				used = last
				spanStart, spanEnd = pieces[first][0], pieces[last][1]
				absolute = match.start() + spanStart
				found.append(
					Finding(
						absolute,
						absolute + spanEnd - spanStart,
						run[spanStart:spanEnd],
						"unknown",
						suggestions,
					),
				)
		return found

	def _bestSpan(
		self,
		run: str,
		pieces: list[tuple[int, int, bool]],
		index: int,
	) -> tuple[int, int, list[str]]:
		"""The unknown piece, joined with its neighbours when that gives a known word nearby.

		A word typed incompletely, such as คอมพิวเตอ, is split into pieces; joining them finds คอมพิวเตอร์.
		"""
		last = len(pieces) - 1
		# Widest first, so a replacement covers the whole mistyped word.
		spans = [
			(max(0, index - 1), min(last, index + 2)),
			(max(0, index - 1), min(last, index + 1)),
			(index, min(last, index + 2)),
			(max(0, index - 1), index),
			(index, min(last, index + 1)),
			(index, index),
		]
		seen = set()
		for first, final in spans:
			if (first, final) in seen:
				continue
			seen.add((first, final))
			text = run[pieces[first][0] : pieces[final][1]]
			suggestions = self.suggest(text)
			if len(text) >= 4:
				suggestions += [w for w in self.completions(text) if w not in suggestions]
			if suggestions:
				return first, final, suggestions[:5]
		return index, index, []

	def completions(self, prefix: str, limit: int = 3) -> list[str]:
		"""Known words that start with the text and are at most three letters longer."""
		position = bisect.bisect_left(self._sorted, prefix)
		result = []
		while position < len(self._sorted) and self._sorted[position].startswith(prefix):
			word = self._sorted[position]
			if len(word) - len(prefix) <= 3:
				result.append(word)
			position += 1
		return sorted(result, key=len)[:limit]

	def suggest(self, word: str, limit: int = 5) -> list[str]:
		"""Known words one edit away: a letter or mark added, removed, replaced or swapped."""
		candidates: set[str] = set()
		letters = self.alphabet
		for i in range(len(word) + 1):
			left, right = word[:i], word[i:]
			if right:
				candidates.add(left + right[1:])
				for letter in letters:
					candidates.add(left + letter + right[1:])
			if len(right) > 1:
				candidates.add(left + right[1] + right[0] + right[2:])
			for letter in letters:
				candidates.add(left + letter + right)
		candidates.discard(word)
		known = [c for c in candidates if c in self.words]
		return sorted(known, key=lambda c: (abs(len(c) - len(word)), c))[:limit]


def readWords(path: str, kind: str) -> set[str]:
	words: set[str] = set()
	try:
		with open(path, encoding="utf-8", errors="replace") as file:
			lines = file.read().splitlines()
	except OSError:
		return words
	if kind == "hunspell" and lines and lines[0].strip().isdigit():
		lines = lines[1:]
	for line in lines:
		word = line.split("/", 1)[0].strip() if kind == "hunspell" else line.strip()
		if word and not word.startswith("#"):
			words.add(word)
	return words


def readMisspellings(path: str) -> dict[str, tuple[str, str]]:
	pairs: dict[str, tuple[str, str]] = {}
	try:
		with open(path, encoding="utf-8", errors="replace") as file:
			for line in file:
				if not line.strip() or line.startswith("#"):
					continue
				parts = [p.strip() for p in line.rstrip("\n").split("\t")]
				if len(parts) >= 2 and parts[0] and parts[1] and parts[0] != parts[1]:
					pairs[parts[0]] = (parts[1], parts[2] if len(parts) > 2 else "")
	except OSError:
		pass
	return pairs


# Files and updates -------------------------------------------------------------------


class WordLists:
	"""Where the lists are kept, which are in use, and their updates."""

	def __init__(self, folder: str, settings: Any):
		self.folder = folder
		self.settings = settings  # mapping with the add-on configuration
		self.dictionary = Dictionary()
		self._lock = threading.Lock()

	# Paths and state
	def path(self, name: str) -> str:
		return os.path.join(self.folder, name)

	@property
	def myWordsFile(self) -> str:
		return self.path("mywords.txt")

	@property
	def myMisspellingsFile(self) -> str:
		return self.path("mymisspellings.tsv")

	def _statePath(self) -> str:
		return self.path("state.json")

	def state(self) -> dict:
		try:
			with open(self._statePath(), encoding="utf-8") as file:
				return json.load(file)
		except (OSError, ValueError):
			return {}

	def saveState(self, state: dict) -> None:
		os.makedirs(self.folder, exist_ok=True)
		temporary = self._statePath() + ".tmp"
		with open(temporary, "w", encoding="utf-8") as file:
			json.dump(state, file, ensure_ascii=False, indent=1)
		os.replace(temporary, self._statePath())

	def enabledSources(self) -> list[WordSource]:
		enabled = [s for s in sources() if s.id != "libreoffice" or self.settings["wordListLibreOffice"]]
		url = (self.settings["misspellingListUrl"] or "").strip()
		if url:
			enabled.append(
				WordSource(
					"organization",
					# Translators: Name of the organization's own list of commonly misspelled words.
					_("Organization list of commonly misspelled words"),
					url,
					"",
					url,
					kind="misspellings",
				),
			)
		return enabled

	def _sourceFile(self, source: WordSource) -> str:
		extension = {"hunspell": "dic", "misspellings": "tsv"}.get(source.kind, "txt")
		return self.path(f"{source.id}.{extension}")

	def load(self) -> None:
		wordFiles: list[tuple[str, str]] = []
		misspellingFiles = [BUNDLED_MISSPELLINGS]
		for source in self.enabledSources():
			downloaded = self._sourceFile(source)
			if source.kind == "misspellings":
				if os.path.isfile(downloaded):
					misspellingFiles.append(downloaded)
			elif os.path.isfile(downloaded):
				wordFiles.append((downloaded, source.kind))
			elif source.id == "pythainlp":
				wordFiles.append((BUNDLED_WORDS, "words"))
		misspellingFiles.append(self.myMisspellingsFile)
		dictionary = Dictionary()
		dictionary.load(wordFiles, misspellingFiles, self.myWordsFile)
		dictionary.ignored = self.dictionary.ignored
		with self._lock:
			self.dictionary = dictionary

	def addMyWord(self, word: str) -> None:
		os.makedirs(self.folder, exist_ok=True)
		with open(self.myWordsFile, "a", encoding="utf-8") as file:
			file.write(word.strip() + "\n")
		self.dictionary.words.add(word.strip())

	def ensureMyFiles(self) -> None:
		os.makedirs(self.folder, exist_ok=True)
		if not os.path.isfile(self.myWordsFile):
			with open(self.myWordsFile, "w", encoding="utf-8") as file:
				# Translators: First line of the user's own word list file.
				file.write("# " + _("My words: one word per line. These words are never reported.") + "\n")
		if not os.path.isfile(self.myMisspellingsFile):
			with open(self.myMisspellingsFile, "w", encoding="utf-8") as file:
				# Translators: First line of the user's own list of commonly misspelled words.
				file.write(
					"# "
					+ _("My commonly misspelled words: wrong form, Tab, right form, Tab, optional note.")
					+ "\n",
				)

	# Updates
	def _request(self, source: WordSource, etag: str, method: str) -> tuple[int, str, bytes]:
		headers = {"User-Agent": "wordDocAssistant-NVDA-addon"}
		if etag:
			headers["If-None-Match"] = etag
		request = urllib.request.Request(source.url, headers=headers, method=method)
		try:
			with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
				body = response.read() if method == "GET" else b""
				return (
					response.status,
					response.headers.get("ETag", "") or response.headers.get("Last-Modified", ""),
					body,
				)
		except urllib.error.HTTPError as error:
			if error.code == 304:
				return 304, etag, b""
			raise

	def checkForUpdates(self) -> list[WordSource]:
		"""Sources with a newer version than the one in use. Network; run in the background."""
		state = self.state()
		updates = []
		for source in self.enabledSources():
			known = state.get(source.id, {}).get("etag") or source.bundledEtag
			downloaded = os.path.isfile(self._sourceFile(source))
			try:
				status, etag, _body = self._request(
					source,
					known if (downloaded or source.bundledEtag) else "",
					"HEAD",
				)
			except (urllib.error.URLError, TimeoutError, OSError):
				log.debugWarning(f"Cannot check {source.url}", exc_info=True)
				continue
			if status == 304:
				continue
			if not (downloaded or source.bundledEtag) or (etag and etag != known):
				updates.append(source)
		state["lastCheck"] = time.time()
		state["updates"] = [s.id for s in updates]
		self.saveState(state)
		return updates

	def download(self, source: WordSource) -> int:
		"""Download one source. Returns the number of entries. Network; run in the background."""
		_status, etag, body = self._request(source, "", "GET")
		text = body.decode("utf-8")
		entries = [line for line in text.splitlines() if line.strip() and not line.startswith("#")]
		thai = sum(1 for line in entries[:500] if _THAI_RUN.search(line))
		minimum = 5 if source.kind == "misspellings" else 1000
		if len(entries) < minimum or thai < min(len(entries[:500]), minimum) // 2:
			raise ValueError(f"{source.url} does not look like a Thai word list")
		os.makedirs(self.folder, exist_ok=True)
		target = self._sourceFile(source)
		with open(target + ".tmp", "w", encoding="utf-8") as file:
			file.write(text)
		os.replace(target + ".tmp", target)
		state = self.state()
		state[source.id] = {"etag": etag, "downloaded": time.time(), "entries": len(entries)}
		state["updates"] = [i for i in state.get("updates", []) if i != source.id]
		self.saveState(state)
		return len(entries)

	def dueForCheck(self) -> bool:
		return time.time() - float(self.state().get("lastCheck", 0)) >= CHECK_INTERVAL

	def pendingUpdates(self) -> list[WordSource]:
		ids = set(self.state().get("updates", []))
		return [s for s in self.enabledSources() if s.id in ids]

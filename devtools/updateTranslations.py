"""Refresh the Thai translation after code changes.

Extracts the translatable strings, merges them into addon/locale/th/LC_MESSAGES/nvda.po
(existing translations are kept), and lists the entries that still need a Thai translation.

Usage: python devtools/updateTranslations.py
Requires gettext (xgettext, msgmerge, msgfmt).
"""

import glob
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PO = os.path.join(ROOT, "addon", "locale", "th", "LC_MESSAGES", "nvda.po")


def main() -> int:
	sources = sorted(glob.glob(os.path.join(ROOT, "addon", "globalPlugins", "wordDocAssistant", "*.py")))
	sources.append(os.path.join(ROOT, "buildVars.py"))
	with tempfile.TemporaryDirectory() as folder:
		pot = os.path.join(folder, "wordDocAssistant.pot")
		subprocess.run(
			["xgettext", "--from-code=UTF-8", "--keyword=pgettext:1c,2", "-c", "--no-location", "-o", pot, *sources],
			check=True,
		)
		with open(pot, encoding="utf-8") as file:
			text = file.read().replace("charset=CHARSET", "charset=UTF-8")
		with open(pot, "w", encoding="utf-8") as file:
			file.write(text)
		subprocess.run(["msgmerge", "--quiet", "--no-fuzzy-matching", "--no-location", "--update", "--backup=none", PO, pot], check=True)
	result = subprocess.run(["msgfmt", "--check", "--statistics", "-o", os.devnull, PO], capture_output=True, text=True)
	print(result.stderr.strip())
	untranslated = subprocess.run(["msgattrib", "--untranslated", "--no-obsolete", PO], capture_output=True, text=True).stdout
	missing = [line[7:-1] for line in untranslated.splitlines() if line.startswith('msgid "') and line != 'msgid ""']
	if missing:
		print("\nStill needs a Thai translation (edit nvda.po):")
		for entry in missing:
			print("  " + entry)
	return 0


if __name__ == "__main__":
	sys.exit(main())

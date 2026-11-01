# Word Document Assistant for NVDA
# All element types in the order of the left and right arrow keys.
# Copyright (C) 2026 Peera Pilarit
# This file is covered by the GNU General Public License version 2.

from . import a11y, elements, structure, thaicheck

NAVIGATION_TYPES: list[elements.ElementType] = [
	elements.SpellingErrors(),
	elements.GrammarErrors(),
	elements.Comments(),
	elements.Footnotes(),
	elements.Endnotes(),
	elements.Revisions(),
	structure.Headings(),
	structure.Tables(),
	a11y.Pictures(),
	structure.Links(),
	structure.CrossReferences(),
	structure.Citations(),
	structure.Bookmarks(),
	structure.Fields(),
	a11y.LanguageProblems(),
	thaicheck.ThaiWords(),
	a11y.AccessibilityIssues(),
]

TYPES_BY_VK = {kind.vk: kind for kind in NAVIGATION_TYPES}
assert len(TYPES_BY_VK) == len(NAVIGATION_TYPES), "two element types share a letter"

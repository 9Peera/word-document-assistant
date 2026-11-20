# Word lists

* `words_th.txt` - Thai word list from PyThaiNLP (https://github.com/PyThaiNLP/pythainlp),
  released under the Creative Commons Zero 1.0 Universal Public Domain Dedication (CC0 1.0).
  It is collected from real text, so it also contains common misspellings; the add-on removes
  every wrong form listed in `misspellings_th.tsv` before it uses the list.
* `misspellings_th.tsv` - commonly misspelled Thai words, written for this add-on (GPL v2 like the add-on).
* `misspellings_en.tsv` - commonly misspelled English words, written for this add-on (GPL v2 like the add-on).
  Every `misspellings_<language>.tsv` file in this folder is loaded, so lists for more languages can be added.

Updated copies are downloaded to the NVDA configuration folder (`wordDocAssistant/dictionaries`).
The optional LibreOffice Thai dictionary (LGPL, from NECTEC LEXiTRON and libthai) is only downloaded when it is turned on in the settings.

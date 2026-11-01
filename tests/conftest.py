import os
import sys

HERE = os.path.dirname(__file__)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "..", "addon", "globalPlugins"))

import nvdaStubs  # noqa: E402,F401  (registers the NVDA stand-ins before the add-on is imported)
import pytest  # noqa: E402


@pytest.fixture(autouse=True)
def config():
	nvdaStubs.resetConfig()
	return nvdaStubs.conf["wordDocAssistant"]

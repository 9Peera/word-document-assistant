import ast
import os

ROOT = os.path.join(os.path.dirname(__file__), "..")


def test_build_variables_are_valid_python():
	with open(os.path.join(ROOT, "buildVars.py"), encoding="utf-8") as file:
		ast.parse(file.read())

import types

from fakeword import Context, Doc, Host
from nvdaStubs import COMError
from wordDocAssistant import elements


class ModernComment:
	"""A Microsoft 365 comment that refuses changes through the object model."""

	def __init__(self, doc):
		self.Author = "Somchai"
		self.Done = False
		self.edited = False
		self.Scope = doc.Range(0, 4)
		self.Ancestor = None

		def refuse(*args):
			raise COMError("not allowed")

		self.Replies = types.SimpleNamespace(Count=0, Item=lambda i: None, Add=refuse)
		self.Range = types.SimpleNamespace(Text="Check this")

	def Edit(self):
		self.edited = True


class FrozenRange:
	@property
	def Text(self):
		return "Check this"

	@Text.setter
	def Text(self, value):
		raise COMError("not allowed")


def test_reply_and_edit_fall_back_to_the_clipboard():
	doc = Doc(["text here"])
	ctx = Context(doc)
	comment = ModernComment(doc)
	comments = elements.Comments()
	host = Host(prompts=["My reply"])
	comments._reply(ctx, comment, host)
	assert host.copied == ["My reply"] and comment.edited
	assert "copied" in host.spoken[-1]
	comment.Range = FrozenRange()
	comment.edited = False
	host = Host(prompts=["New text"])
	comments._editText(ctx, comment, host)
	assert host.copied == ["New text"] and comment.edited and host.changes == 0

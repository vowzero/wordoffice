"""
        公共节点
"""
from typing import override

from core.node import OoxmlNode
from .paragraph import Paragraph


class BlockLevelElements(OoxmlNode):
	"""
	块级正文节点区域
	"""
	
	def __init__(self, *args, **kwargs):
		super().__init__(args, kwargs)
		self._empty_paragraph:Paragraph = None
	
	def add_p(self, _p) -> BlockLevelElements: ...
	
	def add_table(self, _t) -> BlockLevelElements: ...
	


class HeaderFooter(BlockLevelElements):
	namespace: dict[str, str]
	
	def __init__(self, *args, **kwargs): ...
	
	@override
	def to_xml(self): ...



"""
        公共节点
"""
from typing import Iterable

from wordoffice.core.node import OoxmlNode, children_node_to_xml
from wordoffice.ooxml.run import Run
from .utils import xml_header


class BlockLevelElements(OoxmlNode):
	"""
	块级正文节点区域
	"""
	
	def add_p(self, _p):
		from wordoffice.ooxml.paragraph import Paragraph
		if isinstance(_p, Paragraph):
			self.append(_p)
		elif isinstance(_p, str | Run):
			self.append(Paragraph(_p))
		elif isinstance(_p, Iterable):
			for p in _p:
				if isinstance(p, Paragraph):
					self.append(p)
				else:
					self.append(Paragraph(p))
		else:
			raise TypeError(f"Unsupported type: {type(_p)}")
		return self
	
	def add_table(self, _t):
		from wordoffice.ooxml.table import Table
		if isinstance(_t, Table):
			self.append(_t)
		else:
			raise TypeError(f"Unsupported type: {type(_t)}")
		return self
	
	def before_to_xml(self):
		from .paragraph import Paragraph
		data = super().before_to_xml()
		origin = data[3]
		children_length = len(origin)
		if children_length == 0 or (children_length == 1 and origin[0].tag_name != 'w:p'):
			origin.append(Paragraph())
			return data
		if origin[-1].tag_name != 'w:p':
			origin.append(Paragraph())
		new = []
		for i in range(1, len(origin)):
			new.append(origin[i - 1])
			if origin[i].tag_name != 'w:p':
				if origin[i - 1].tag_name != 'w:p':
					new.append('<w:p/>')
		new.append(origin[-1])
		
		data[3] = new
		return data
	
	def to_xml(self):
		if self.tag_name == None:
			children = self.before_to_xml()[3]
			return ''.join(children_node_to_xml(children))
		else:
			return super().to_xml()


class HeaderFooter(BlockLevelElements):
	# TODO: 应为EG_BlockLevelElts的容器，现在是EG_BlockLevelElts的子集paragraphs
	namespace = {
		'xmlns:mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
		'xmlns:r' : "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		'xmlns:m' : "http://schemas.openxmlformats.org/officeDocument/2006/math",
		'xmlns:wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		'xmlns:w' : "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	}
	
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._attrs.update(self.namespace)
	
	def to_xml(self):
		res = xml_header + super().to_xml()
		return res

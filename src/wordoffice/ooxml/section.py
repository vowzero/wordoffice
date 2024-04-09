from wordoffice.core.field import AttrField, SubField
from wordoffice.core.node import OoxmlNode, OoxmlNodeBase
from wordoffice.core.opc import NodePart, OPC_BUILTIN_MAP, PartBase
from wordoffice.core.type import *
from .enums import *
from .shared import HeaderFooter, BlockLevelElements
from .utils import create_switch_node, SimpleValNode


class HeaderReference(OoxmlNode):
	tag = "w:headerReference"
	enum = ST_HdrFtr
	type = AttrField('w:type', ('enum', {'enums': enum}), default=enum.DEFAULT, required=True)
	rId = AttrField('r:id', 'string', required=True)
	
	@classmethod
	def DEFAULT(cls, part: PartBase):
		return cls(type=ST_HdrFtr.DEFAULT, rId=part.part_rid)


class FooterReference(OoxmlNode):
	tag = "w:footerReference"
	enum = ST_HdrFtr
	type = AttrField('w:type', ('enum', {'enums': enum}), default=enum.DEFAULT, required=True)
	rId = AttrField('r:id', 'string', required=True)


class PageSize(OoxmlNode):
	tag = "w:pgSz"
	enum = ST_PageOrientation
	width = AttrField('w:w', OoxmlTwipsMeasure, required=True)
	height = AttrField('w:h', OoxmlTwipsMeasure, required=True)
	code = AttrField('w:code', 'integer')
	orient = AttrField('w:orient', ('enum', {'enums': enum}))


class PageMargin(OoxmlNode):
	tag = "w:pgMar"
	left = AttrField('w:left', OoxmlTwipsMeasure, required=True)
	top = AttrField('w:top', OoxmlSignedTwipsMeasure, required=True)
	right = AttrField('w:right', OoxmlTwipsMeasure, required=True)
	bottom = AttrField('w:bottom', OoxmlSignedTwipsMeasure, required=True)
	header = AttrField('w:header', OoxmlTwipsMeasure, required=True)
	footer = AttrField('w:footer', OoxmlTwipsMeasure, required=True)
	gutter = AttrField('w:gutter', OoxmlTwipsMeasure, required=True)


class Col(OoxmlNode):
	tag = "w:col"
	space = AttrField('w:space', OoxmlTwipsMeasure, default='0')
	"""文本列之间的间距"""
	width = AttrField('w:w', OoxmlTwipsMeasure)
	"""文本列的宽度"""


class Cols(OoxmlNode):
	tag = "w:cols"
	_children_type = (Col,)
	
	equalWidth = AttrField('w:equalWidth', 'ooxml_boolean', default=False)
	"""指定当前列是否具有相同的宽度，会对col重新计算，目前不支持"""
	num = AttrField('w:num', 'integer', default='1')
	"""等宽列时，列的数量，equalWidth=True生效"""
	sep = AttrField('w:sep', 'ooxml_boolean', default=False)
	"""指定是否在此节中的每个文本列之间绘制垂直线"""
	space = AttrField('w:space', OoxmlTwipsMeasure, default='720')
	"""指定当前节中文本列之间的间距，equalWidth=True生效"""
	
	def add_col(self, col: Col):
		if len(self._children) == 45:
			assert not "最多只能有45个列"
		self.append(col)


class DocGrid(OoxmlNode):
	tag = "w:docGrid"
	enum = ST_DocGrid
	type = AttrField('w:type', ('enum', {'enums': enum}))
	linePitch = AttrField('w:linePitch', 'integer')
	charSpace = AttrField('w:charSpace', 'integer')


Bidi = create_switch_node('Bidi', 'w:bidi', False)
FormProt = create_switch_node('FormProtection', 'w:formProt', False)


class LnNumType(OoxmlNode):
	tag = 'w:lnNumType'
	enum = ST_LineNumberRestart
	countBy = AttrField('w:countBy', 'integer')
	start = AttrField('w:start', 'integer', default="1")
	distance = AttrField('w:distance', OoxmlTwipsMeasure)
	restart = AttrField('w:restart', ('enum', {'enums': enum}), default=enum.NEW_PAGE)


class PgNumType(OoxmlNode):
	tag = 'w:pgNumType'
	enum_fmt = ST_NumberFormat
	enum_chapSep = ST_ChapterSep
	fmt = AttrField('w:fmt', ('enum', {'enums': enum_fmt}), default=enum_fmt.DECIMAL)
	start = AttrField('w:start', 'integer')
	chapStyle = AttrField('w:chapStyle', 'integer')
	chapSep = AttrField('w:chapSep', ('enum', {'enums': enum_chapSep}), default=enum_chapSep.HYPHEN)


RtlGutter = create_switch_node('RtlGutter', 'w:rtlGutter', False)


class TextDirection(SimpleValNode):
	tag = 'w:textDirection'
	enum = ST_TextDirection
	val = AttrField('w:val', ('enum', {'enums': enum}), required=True)


class SectionType(SimpleValNode):
	tag = 'w:type'
	enum = ST_SectionMark
	val = AttrField('w:val', ('enum', {'enums': enum}), required=True)


class VAlign(SimpleValNode):
	tag = 'w:vAlign'
	enum = ST_VerticalJc
	val = AttrField('w:val', ('enum', {'enums': enum}), required=True)


class SectionPr(OoxmlNode):
	tag = "w:sectPr"
	# headerReference = OoxmlSubField(HeaderReference,property=False)
	# footerReference = OoxmlSubField(FooterReference,property=False)
	type = SubField(SectionType,index=3)
	pageSize = SubField(PageSize,index=4)
	pageMargin = SubField(PageMargin,index=5)
	bidi = SubField(Bidi)
	cols = SubField(Cols,index=6)
	docGrid = SubField(DocGrid,index=100)
	formProt = SubField(FormProt)
	lnNumType = SubField(LnNumType)
	pgNumType = SubField(PgNumType)
	rtlGutter = SubField(RtlGutter)
	textDirection = SubField(TextDirection)
	vAlign = SubField(VAlign)
	
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self.hdrs = {}
		self.ftrs = {}
	
	def setHeaderReference(self, headerReference: HeaderReference):
		self.hdrs[str(headerReference.type)] = headerReference
	
	def delHeaderReference(self, headerReferenceType: ST_HdrFtr):
		self.hdrs[str(headerReferenceType)] = None
	
	def setFooterReference(self, footerReference: FooterReference):
		self.ftrs[str(footerReference.type)] = footerReference
	
	def delFooterReference(self, footerReferenceType: ST_HdrFtr):
		self.ftrs[str(footerReferenceType)] = None
	
	def iter_children(self):
		for headerReference in self.hdrs.values():
			if headerReference is not None:
				yield headerReference
		for footerReference in self.ftrs.values():
			if footerReference is not None:
				yield footerReference
		yield from super().iter_children()


class Section(BlockLevelElements):
	sectPr: SectionPr = SubField(type=SectionPr)
	
	def __init__(self, sectPr=True, *args, **kwargs):
		super().__init__(sectPr=sectPr, *args, **kwargs)
		self.isLastSection = False
	
	@property
	def tag_name(self):
		return None
	
	def before_to_xml(self):
		data = super().before_to_xml()
		children = data[3]
		children_xml = children[:-1]
		
		last_child = children[-1]
		
		if not self.isLastSection:
			last_child.set_temp_section(self.sectPr)
			children_xml.append(last_child.to_xml())
		else:
			
			children_xml.append(last_child.to_xml())
			if self.sectPr is not None:
				children_xml.append(self.sectPr.to_xml())
		
		self.isLastSection = False
		data[3] = children_xml
		return data


class _HeaderFooterPartBase(HeaderFooter, NodePart):
	opc_info = None
	
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		super(OoxmlNodeBase, self).__init__(self, self.opc_info)


class Header(_HeaderFooterPartBase):
	tag = 'w:hdr'
	
	opc_info = OPC_BUILTIN_MAP['header']
	
	def __init__(self, word: "Docx" = None, *args, **kwargs):
		super().__init__(*args, **kwargs)
		if word is not None:
			word.add_word_part(self)


class Footer(_HeaderFooterPartBase):
	tag = 'w:ftr'
	
	opc_info = OPC_BUILTIN_MAP['footer']

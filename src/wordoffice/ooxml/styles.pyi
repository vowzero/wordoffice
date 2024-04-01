from typing import Any

from core import opc
from core.node import OoxmlNode, OoxmlNodeBase
from core.opc import NodePart
from .enums import *
from .paragraph import ParagraphPr
from .run import RunPr
from .table import TablePr, TableRowPr, TableCellPr
from .utils import create_switch_node, SimpleValNode


class RunPrDefault(OoxmlNode):
	rPr:RunPr


class ParagraphPrDefault(OoxmlNode):
	pPr:ParagraphPr


class DocDefaults(OoxmlNode):
	rPrDefault :RunPrDefault
	pPrDefault :ParagraphPrDefault


class LsdException(OoxmlNode):
	name :Any
	locked :Any
	uiPriority :Any
	semiHidden :Any
	unhideWhenUsed :Any
	qFormat :Any


class LatentStyles(OoxmlNode):
	
	defLockedState :Any
	defUIPriority :Any
	defSemiHidden :Any
	defUnhideWhenUsed :Any
	defQFormat :Any
	count :Any
	_children_type = (LsdException,)


class StyleName(SimpleValNode):
	
	val :Any


class StyleAliases(SimpleValNode):
	
	val :Any


class StyleBasedOn(SimpleValNode):
	
	val :Any


class StyleNext(SimpleValNode):
	
	val :Any


class StyleLink(SimpleValNode):
	
	val :Any


StyleAutoRedefine = create_switch_node('StyleAutoRedefine', 'w:autoRedefine', False)
StyleHidden = create_switch_node('StyleHidden', 'w:hidden', False)


class StyleUiPriority(SimpleValNode):
	
	val :Any


StyleSemiHidden = create_switch_node('StyleSemiHidden', 'w:semiHidden', False)
StyleUnhideWhenUsed = create_switch_node('StyleUnhideWhenUsed', 'w:unhideWhenUsed', False)
StyleQFormat = create_switch_node('StyleQFormat', 'w:qFormat', False)
StyleLocked = create_switch_node('StyleLocked', 'w:locked', False)
StylePersonal = create_switch_node('StylePersonal', 'w:personal', False)
StylePersonalCompose = create_switch_node('StylePersonalCompose', 'w:personalCompose', False)
StylePersonalReply = create_switch_node('StylePersonalReply', 'w:personalReply', False)


class TableStylePr(OoxmlNode):
	
	enum = ST_TblStyleOverrideType
	type :Any
	pPr :ParagraphPr
	rPr :RunPr
	tblPr :TablePr
	trPr :TableRowPr
	tcPr :TableCellPr


class Style(OoxmlNode):
	
	enum_type = ST_StyleType
	type :Any
	styleId :Any
	default :Any
	customStyle :Any
	
	name :StyleName
	aliases :StyleAliases
	basedOn :StyleBasedOn
	next :StyleNext
	link :StyleLink
	autoRedefine :StyleAutoRedefine
	hidden :StyleHidden
	uiPriority :StyleUiPriority
	semiHidden :StyleSemiHidden
	unhideWhenUsed :StyleUnhideWhenUsed
	qFormat :StyleQFormat
	locked :StyleLocked
	personal :StylePersonal
	personalCompose :StylePersonalCompose
	personalReply :StylePersonalReply
	# rsid = OoxmlSubField(StyleRsid)
	pPr :ParagraphPr
	rPr :RunPr
	tblPr :TablePr
	trPr :TableRowPr
	tcPr :TableCellPr
	tblStylePr :TableStylePr
	
	@classmethod
	def Character(cls, styleId, **kwargs):...
	
	@classmethod
	def Paragraph(cls, styleId, **kwargs)...

class Styles(OoxmlNode):
	
	docDefaults :DocDefaults
	latentStyles :LatentStyles
	_children_type = (Style,)
	
	def get_styleId_by_name(self, name, type=None):
		for style in self._children:
			if style.name.val == name:
				if type is None:
					return style.styleId
				elif style.type.val == type:
					return style.styleId
		return None


class StylesPart(Styles, NodePart):
	opc_info = opc.OPC_BUILTIN_MAP['styles']
	namespace = {
		'xmlns:mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
		'xmlns:r' : "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		'xmlns:w' : "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	}
	
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		super(OoxmlNodeBase, self).__init__(self, self.opc_info)
		self._attrs.update(self.namespace)
	
	def to_xml(self):...


_builtin_lsds = [{"name": "Normal", "uiPriority": "0", "qFormat": "1"},
		 {"name": "heading 1", "uiPriority": "9", "qFormat": "1"},
		 {"name": "heading 2", "semiHidden": "1", "uiPriority": "9", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "heading 3", "semiHidden": "1", "uiPriority": "9", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "heading 4", "semiHidden": "1", "uiPriority": "9", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "heading 5", "semiHidden": "1", "uiPriority": "9", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "heading 6", "semiHidden": "1", "uiPriority": "9", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "heading 7", "semiHidden": "1", "uiPriority": "9", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "heading 8", "semiHidden": "1", "uiPriority": "9", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "heading 9", "semiHidden": "1", "uiPriority": "9", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "index 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index 5", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index 6", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index 7", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index 8", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index 9", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "toc 1", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "toc 2", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "toc 3", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "toc 4", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "toc 5", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "toc 6", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "toc 7", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "toc 8", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "toc 9", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1"},
		 {"name": "Normal Indent", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "footnote text", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "annotation text", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "header", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "footer", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "index heading", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "caption", "semiHidden": "1", "uiPriority": "35", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "table of figures", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "envelope address", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "envelope return", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "footnote reference", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "annotation reference", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "line number", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "page number", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "endnote reference", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "endnote text", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "table of authorities", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "macro", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "toa heading", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Bullet", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Number", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List 5", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Bullet 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Bullet 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Bullet 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Bullet 5", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Number 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Number 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Number 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Number 5", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Title", "uiPriority": "10", "qFormat": "1"},
		 {"name": "Closing", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Signature", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Default Paragraph Font", "semiHidden": "1", "uiPriority": "1", "unhideWhenUsed": "1"},
		 {"name": "Body Text", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Body Text Indent", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Continue", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Continue 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Continue 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Continue 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "List Continue 5", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Message Header", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Subtitle", "uiPriority": "11", "qFormat": "1"},
		 {"name": "Salutation", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Date", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Body Text First Indent", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Body Text First Indent 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Note Heading", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Body Text 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Body Text 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Body Text Indent 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Body Text Indent 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Block Text", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Hyperlink", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "FollowedHyperlink", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Strong", "uiPriority": "22", "qFormat": "1"},
		 {"name": "Emphasis", "uiPriority": "20", "qFormat": "1"},
		 {"name": "Document Map", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Plain Text", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "E-mail Signature", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Top of Form", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Bottom of Form", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Normal (Web)", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Acronym", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Address", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Cite", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Code", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Definition", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Keyboard", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Preformatted", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Sample", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Typewriter", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "HTML Variable", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Normal Table", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "annotation subject", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "No List", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Outline List 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Outline List 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Outline List 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Simple 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Simple 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Simple 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Classic 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Classic 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Classic 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Classic 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Colorful 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Colorful 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Colorful 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Columns 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Columns 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Columns 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Columns 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Columns 5", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid 5", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid 6", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid 7", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid 8", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table List 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table List 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table List 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table List 4", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table List 5", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table List 6", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table List 7", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table List 8", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table 3D effects 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table 3D effects 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table 3D effects 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Contemporary", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Elegant", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Professional", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Subtle 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Subtle 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Web 1", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Web 2", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Web 3", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Balloon Text", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Table Grid", "uiPriority": "39"},
		 {"name": "Table Theme", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Placeholder Text", "semiHidden": "1"},
		 {"name": "No Spacing", "uiPriority": "1", "qFormat": "1"},
		 {"name": "Light Shading", "uiPriority": "60"},
		 {"name": "Light List", "uiPriority": "61"},
		 {"name": "Light Grid", "uiPriority": "62"},
		 {"name": "Medium Shading 1", "uiPriority": "63"},
		 {"name": "Medium Shading 2", "uiPriority": "64"},
		 {"name": "Medium List 1", "uiPriority": "65"},
		 {"name": "Medium List 2", "uiPriority": "66"},
		 {"name": "Medium Grid 1", "uiPriority": "67"},
		 {"name": "Medium Grid 2", "uiPriority": "68"},
		 {"name": "Medium Grid 3", "uiPriority": "69"},
		 {"name": "Dark List", "uiPriority": "70"},
		 {"name": "Colorful Shading", "uiPriority": "71"},
		 {"name": "Colorful List", "uiPriority": "72"},
		 {"name": "Colorful Grid", "uiPriority": "73"},
		 {"name": "Light Shading Accent 1", "uiPriority": "60"},
		 {"name": "Light List Accent 1", "uiPriority": "61"},
		 {"name": "Light Grid Accent 1", "uiPriority": "62"},
		 {"name": "Medium Shading 1 Accent 1", "uiPriority": "63"},
		 {"name": "Medium Shading 2 Accent 1", "uiPriority": "64"},
		 {"name": "Medium List 1 Accent 1", "uiPriority": "65"},
		 {"name": "Revision", "semiHidden": "1"},
		 {"name": "List Paragraph", "uiPriority": "34", "qFormat": "1"},
		 {"name": "Quote", "uiPriority": "29", "qFormat": "1"},
		 {"name": "Intense Quote", "uiPriority": "30", "qFormat": "1"},
		 {"name": "Medium List 2 Accent 1", "uiPriority": "66"},
		 {"name": "Medium Grid 1 Accent 1", "uiPriority": "67"},
		 {"name": "Medium Grid 2 Accent 1", "uiPriority": "68"},
		 {"name": "Medium Grid 3 Accent 1", "uiPriority": "69"},
		 {"name": "Dark List Accent 1", "uiPriority": "70"},
		 {"name": "Colorful Shading Accent 1", "uiPriority": "71"},
		 {"name": "Colorful List Accent 1", "uiPriority": "72"},
		 {"name": "Colorful Grid Accent 1", "uiPriority": "73"},
		 {"name": "Light Shading Accent 2", "uiPriority": "60"},
		 {"name": "Light List Accent 2", "uiPriority": "61"},
		 {"name": "Light Grid Accent 2", "uiPriority": "62"},
		 {"name": "Medium Shading 1 Accent 2", "uiPriority": "63"},
		 {"name": "Medium Shading 2 Accent 2", "uiPriority": "64"},
		 {"name": "Medium List 1 Accent 2", "uiPriority": "65"},
		 {"name": "Medium List 2 Accent 2", "uiPriority": "66"},
		 {"name": "Medium Grid 1 Accent 2", "uiPriority": "67"},
		 {"name": "Medium Grid 2 Accent 2", "uiPriority": "68"},
		 {"name": "Medium Grid 3 Accent 2", "uiPriority": "69"},
		 {"name": "Dark List Accent 2", "uiPriority": "70"},
		 {"name": "Colorful Shading Accent 2", "uiPriority": "71"},
		 {"name": "Colorful List Accent 2", "uiPriority": "72"},
		 {"name": "Colorful Grid Accent 2", "uiPriority": "73"},
		 {"name": "Light Shading Accent 3", "uiPriority": "60"},
		 {"name": "Light List Accent 3", "uiPriority": "61"},
		 {"name": "Light Grid Accent 3", "uiPriority": "62"},
		 {"name": "Medium Shading 1 Accent 3", "uiPriority": "63"},
		 {"name": "Medium Shading 2 Accent 3", "uiPriority": "64"},
		 {"name": "Medium List 1 Accent 3", "uiPriority": "65"},
		 {"name": "Medium List 2 Accent 3", "uiPriority": "66"},
		 {"name": "Medium Grid 1 Accent 3", "uiPriority": "67"},
		 {"name": "Medium Grid 2 Accent 3", "uiPriority": "68"},
		 {"name": "Medium Grid 3 Accent 3", "uiPriority": "69"},
		 {"name": "Dark List Accent 3", "uiPriority": "70"},
		 {"name": "Colorful Shading Accent 3", "uiPriority": "71"},
		 {"name": "Colorful List Accent 3", "uiPriority": "72"},
		 {"name": "Colorful Grid Accent 3", "uiPriority": "73"},
		 {"name": "Light Shading Accent 4", "uiPriority": "60"},
		 {"name": "Light List Accent 4", "uiPriority": "61"},
		 {"name": "Light Grid Accent 4", "uiPriority": "62"},
		 {"name": "Medium Shading 1 Accent 4", "uiPriority": "63"},
		 {"name": "Medium Shading 2 Accent 4", "uiPriority": "64"},
		 {"name": "Medium List 1 Accent 4", "uiPriority": "65"},
		 {"name": "Medium List 2 Accent 4", "uiPriority": "66"},
		 {"name": "Medium Grid 1 Accent 4", "uiPriority": "67"},
		 {"name": "Medium Grid 2 Accent 4", "uiPriority": "68"},
		 {"name": "Medium Grid 3 Accent 4", "uiPriority": "69"},
		 {"name": "Dark List Accent 4", "uiPriority": "70"},
		 {"name": "Colorful Shading Accent 4", "uiPriority": "71"},
		 {"name": "Colorful List Accent 4", "uiPriority": "72"},
		 {"name": "Colorful Grid Accent 4", "uiPriority": "73"},
		 {"name": "Light Shading Accent 5", "uiPriority": "60"},
		 {"name": "Light List Accent 5", "uiPriority": "61"},
		 {"name": "Light Grid Accent 5", "uiPriority": "62"},
		 {"name": "Medium Shading 1 Accent 5", "uiPriority": "63"},
		 {"name": "Medium Shading 2 Accent 5", "uiPriority": "64"},
		 {"name": "Medium List 1 Accent 5", "uiPriority": "65"},
		 {"name": "Medium List 2 Accent 5", "uiPriority": "66"},
		 {"name": "Medium Grid 1 Accent 5", "uiPriority": "67"},
		 {"name": "Medium Grid 2 Accent 5", "uiPriority": "68"},
		 {"name": "Medium Grid 3 Accent 5", "uiPriority": "69"},
		 {"name": "Dark List Accent 5", "uiPriority": "70"},
		 {"name": "Colorful Shading Accent 5", "uiPriority": "71"},
		 {"name": "Colorful List Accent 5", "uiPriority": "72"},
		 {"name": "Colorful Grid Accent 5", "uiPriority": "73"},
		 {"name": "Light Shading Accent 6", "uiPriority": "60"},
		 {"name": "Light List Accent 6", "uiPriority": "61"},
		 {"name": "Light Grid Accent 6", "uiPriority": "62"},
		 {"name": "Medium Shading 1 Accent 6", "uiPriority": "63"},
		 {"name": "Medium Shading 2 Accent 6", "uiPriority": "64"},
		 {"name": "Medium List 1 Accent 6", "uiPriority": "65"},
		 {"name": "Medium List 2 Accent 6", "uiPriority": "66"},
		 {"name": "Medium Grid 1 Accent 6", "uiPriority": "67"},
		 {"name": "Medium Grid 2 Accent 6", "uiPriority": "68"},
		 {"name": "Medium Grid 3 Accent 6", "uiPriority": "69"},
		 {"name": "Dark List Accent 6", "uiPriority": "70"},
		 {"name": "Colorful Shading Accent 6", "uiPriority": "71"},
		 {"name": "Colorful List Accent 6", "uiPriority": "72"},
		 {"name": "Colorful Grid Accent 6", "uiPriority": "73"},
		 {"name": "Subtle Emphasis", "uiPriority": "19", "qFormat": "1"},
		 {"name": "Intense Emphasis", "uiPriority": "21", "qFormat": "1"},
		 {"name": "Subtle Reference", "uiPriority": "31", "qFormat": "1"},
		 {"name": "Intense Reference", "uiPriority": "32", "qFormat": "1"},
		 {"name": "Book Title", "uiPriority": "33", "qFormat": "1"},
		 {"name": "Bibliography", "semiHidden": "1", "uiPriority": "37", "unhideWhenUsed": "1"},
		 {"name": "TOC Heading", "semiHidden": "1", "uiPriority": "39", "unhideWhenUsed": "1", "qFormat": "1"},
		 {"name": "Plain Table 1", "uiPriority": "41"},
		 {"name": "Plain Table 2", "uiPriority": "42"},
		 {"name": "Plain Table 3", "uiPriority": "43"},
		 {"name": "Plain Table 4", "uiPriority": "44"},
		 {"name": "Plain Table 5", "uiPriority": "45"},
		 {"name": "Grid Table Light", "uiPriority": "40"},
		 {"name": "Grid Table 1 Light", "uiPriority": "46"},
		 {"name": "Grid Table 2", "uiPriority": "47"},
		 {"name": "Grid Table 3", "uiPriority": "48"},
		 {"name": "Grid Table 4", "uiPriority": "49"},
		 {"name": "Grid Table 5 Dark", "uiPriority": "50"},
		 {"name": "Grid Table 6 Colorful", "uiPriority": "51"},
		 {"name": "Grid Table 7 Colorful", "uiPriority": "52"},
		 {"name": "Grid Table 1 Light Accent 1", "uiPriority": "46"},
		 {"name": "Grid Table 2 Accent 1", "uiPriority": "47"},
		 {"name": "Grid Table 3 Accent 1", "uiPriority": "48"},
		 {"name": "Grid Table 4 Accent 1", "uiPriority": "49"},
		 {"name": "Grid Table 5 Dark Accent 1", "uiPriority": "50"},
		 {"name": "Grid Table 6 Colorful Accent 1", "uiPriority": "51"},
		 {"name": "Grid Table 7 Colorful Accent 1", "uiPriority": "52"},
		 {"name": "Grid Table 1 Light Accent 2", "uiPriority": "46"},
		 {"name": "Grid Table 2 Accent 2", "uiPriority": "47"},
		 {"name": "Grid Table 3 Accent 2", "uiPriority": "48"},
		 {"name": "Grid Table 4 Accent 2", "uiPriority": "49"},
		 {"name": "Grid Table 5 Dark Accent 2", "uiPriority": "50"},
		 {"name": "Grid Table 6 Colorful Accent 2", "uiPriority": "51"},
		 {"name": "Grid Table 7 Colorful Accent 2", "uiPriority": "52"},
		 {"name": "Grid Table 1 Light Accent 3", "uiPriority": "46"},
		 {"name": "Grid Table 2 Accent 3", "uiPriority": "47"},
		 {"name": "Grid Table 3 Accent 3", "uiPriority": "48"},
		 {"name": "Grid Table 4 Accent 3", "uiPriority": "49"},
		 {"name": "Grid Table 5 Dark Accent 3", "uiPriority": "50"},
		 {"name": "Grid Table 6 Colorful Accent 3", "uiPriority": "51"},
		 {"name": "Grid Table 7 Colorful Accent 3", "uiPriority": "52"},
		 {"name": "Grid Table 1 Light Accent 4", "uiPriority": "46"},
		 {"name": "Grid Table 2 Accent 4", "uiPriority": "47"},
		 {"name": "Grid Table 3 Accent 4", "uiPriority": "48"},
		 {"name": "Grid Table 4 Accent 4", "uiPriority": "49"},
		 {"name": "Grid Table 5 Dark Accent 4", "uiPriority": "50"},
		 {"name": "Grid Table 6 Colorful Accent 4", "uiPriority": "51"},
		 {"name": "Grid Table 7 Colorful Accent 4", "uiPriority": "52"},
		 {"name": "Grid Table 1 Light Accent 5", "uiPriority": "46"},
		 {"name": "Grid Table 2 Accent 5", "uiPriority": "47"},
		 {"name": "Grid Table 3 Accent 5", "uiPriority": "48"},
		 {"name": "Grid Table 4 Accent 5", "uiPriority": "49"},
		 {"name": "Grid Table 5 Dark Accent 5", "uiPriority": "50"},
		 {"name": "Grid Table 6 Colorful Accent 5", "uiPriority": "51"},
		 {"name": "Grid Table 7 Colorful Accent 5", "uiPriority": "52"},
		 {"name": "Grid Table 1 Light Accent 6", "uiPriority": "46"},
		 {"name": "Grid Table 2 Accent 6", "uiPriority": "47"},
		 {"name": "Grid Table 3 Accent 6", "uiPriority": "48"},
		 {"name": "Grid Table 4 Accent 6", "uiPriority": "49"},
		 {"name": "Grid Table 5 Dark Accent 6", "uiPriority": "50"},
		 {"name": "Grid Table 6 Colorful Accent 6", "uiPriority": "51"},
		 {"name": "Grid Table 7 Colorful Accent 6", "uiPriority": "52"},
		 {"name": "List Table 1 Light", "uiPriority": "46"},
		 {"name": "List Table 2", "uiPriority": "47"},
		 {"name": "List Table 3", "uiPriority": "48"},
		 {"name": "List Table 4", "uiPriority": "49"},
		 {"name": "List Table 5 Dark", "uiPriority": "50"},
		 {"name": "List Table 6 Colorful", "uiPriority": "51"},
		 {"name": "List Table 7 Colorful", "uiPriority": "52"},
		 {"name": "List Table 1 Light Accent 1", "uiPriority": "46"},
		 {"name": "List Table 2 Accent 1", "uiPriority": "47"},
		 {"name": "List Table 3 Accent 1", "uiPriority": "48"},
		 {"name": "List Table 4 Accent 1", "uiPriority": "49"},
		 {"name": "List Table 5 Dark Accent 1", "uiPriority": "50"},
		 {"name": "List Table 6 Colorful Accent 1", "uiPriority": "51"},
		 {"name": "List Table 7 Colorful Accent 1", "uiPriority": "52"},
		 {"name": "List Table 1 Light Accent 2", "uiPriority": "46"},
		 {"name": "List Table 2 Accent 2", "uiPriority": "47"},
		 {"name": "List Table 3 Accent 2", "uiPriority": "48"},
		 {"name": "List Table 4 Accent 2", "uiPriority": "49"},
		 {"name": "List Table 5 Dark Accent 2", "uiPriority": "50"},
		 {"name": "List Table 6 Colorful Accent 2", "uiPriority": "51"},
		 {"name": "List Table 7 Colorful Accent 2", "uiPriority": "52"},
		 {"name": "List Table 1 Light Accent 3", "uiPriority": "46"},
		 {"name": "List Table 2 Accent 3", "uiPriority": "47"},
		 {"name": "List Table 3 Accent 3", "uiPriority": "48"},
		 {"name": "List Table 4 Accent 3", "uiPriority": "49"},
		 {"name": "List Table 5 Dark Accent 3", "uiPriority": "50"},
		 {"name": "List Table 6 Colorful Accent 3", "uiPriority": "51"},
		 {"name": "List Table 7 Colorful Accent 3", "uiPriority": "52"},
		 {"name": "List Table 1 Light Accent 4", "uiPriority": "46"},
		 {"name": "List Table 2 Accent 4", "uiPriority": "47"},
		 {"name": "List Table 3 Accent 4", "uiPriority": "48"},
		 {"name": "List Table 4 Accent 4", "uiPriority": "49"},
		 {"name": "List Table 5 Dark Accent 4", "uiPriority": "50"},
		 {"name": "List Table 6 Colorful Accent 4", "uiPriority": "51"},
		 {"name": "List Table 7 Colorful Accent 4", "uiPriority": "52"},
		 {"name": "List Table 1 Light Accent 5", "uiPriority": "46"},
		 {"name": "List Table 2 Accent 5", "uiPriority": "47"},
		 {"name": "List Table 3 Accent 5", "uiPriority": "48"},
		 {"name": "List Table 4 Accent 5", "uiPriority": "49"},
		 {"name": "List Table 5 Dark Accent 5", "uiPriority": "50"},
		 {"name": "List Table 6 Colorful Accent 5", "uiPriority": "51"},
		 {"name": "List Table 7 Colorful Accent 5", "uiPriority": "52"},
		 {"name": "List Table 1 Light Accent 6", "uiPriority": "46"},
		 {"name": "List Table 2 Accent 6", "uiPriority": "47"},
		 {"name": "List Table 3 Accent 6", "uiPriority": "48"},
		 {"name": "List Table 4 Accent 6", "uiPriority": "49"},
		 {"name": "List Table 5 Dark Accent 6", "uiPriority": "50"},
		 {"name": "List Table 6 Colorful Accent 6", "uiPriority": "51"},
		 {"name": "List Table 7 Colorful Accent 6", "uiPriority": "52"},
		 {"name": "Mention", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Smart Hyperlink", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Hashtag", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Unresolved Mention", "semiHidden": "1", "unhideWhenUsed": "1"},
		 {"name": "Smart Link", "semiHidden": "1", "unhideWhenUsed": "1"}, ]


class StylesBuiltinPart(StylesPart):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		# 没有latent styles，潜在样式由office app内置生成
		self._init_docDefaults()
		self._init_styles()
	
	def _init_docDefaults(self):...
	
	def _init_styles(self):...

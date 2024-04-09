from typing import Iterable

from wordoffice.core.field import SubField, AttrField, SubChoiceField
from wordoffice.core.node import OoxmlNode
from wordoffice.core.type import OoxmlSignedTwipsMeasure, OoxmlHpsMeasure, OoxmlSignedHpsMeasure, OoxmlTwipsMeasure
from wordoffice.common import ColorAuto
from .dml import Drawing
from .enums import *
from .shared_node import Shading, BorderAttrFields as SharedBorder
from .utils import SimpleValNode, create_switch_node, create_empty_node

Bold = create_switch_node("Bold", "w:b", False)
BoldCs = create_switch_node("BoldCs", "w:bCs", False)
Italic = create_switch_node("Italic", "w:i", False)
ItalicCs = create_switch_node("ItalicCs", "w:iCs", False)
Caps = create_switch_node("Caps", "w:caps", False)
SmallCaps = create_switch_node("SmallCaps", "w:smallCaps", False)
Strike = create_switch_node("strike", "w:strike", False)
DoubleStrike = create_switch_node("DStrike", "w:dstrike", False)
Outline = create_switch_node("Outline", "w:outline", False)
Shadow = create_switch_node("Shadow", "w:shadow", False)
Emboss = create_switch_node("Emboss", "w:emboss", False)
Imprint = create_switch_node("Imprint", "w:imprint", False)
NoProof = create_switch_node("NoProof", "w:noProof", False)
SnapToGrid = create_switch_node("SnapToGridRun", "w:snapToGrid", True)
Vanish = create_switch_node("Vanish", "w:vanish", False)
WebHidden = create_switch_node("WebHidden", "w:webHidden", False)


# rtl
# cs
# SepcVanish = create_switch_node("SpecVanish", "w:specVanish", False)
# oMath

class Fonts(OoxmlNode):
	tag = "w:rFonts"
	enum = ST_Hint
	ascii = AttrField('w:ascii', 'string')
	hAnsi = AttrField('w:hAnsi', 'string')
	eastAsia = AttrField('w:eastAsia', 'string')
	cs = AttrField('w:cs', 'string')
	hint = AttrField('w:hint', ('enum', {'enums': enum}), default=enum.DEFAULT)
	asciiTheme = AttrField('w:asciiTheme', 'string')
	hAnsiTheme = AttrField('w:hAnsiTheme', 'string')
	eastAsiaTheme = AttrField('w:eastAsiaTheme', 'string')
	cstheme = AttrField('w:csTheme', 'string')


class Color(SimpleValNode):
	tag = "w:color"
	val = AttrField('w:val', 'ooxml_color', required=True)


class SpacingRun(SimpleValNode):
	tag = "w:spacing"
	val = AttrField('w:val', OoxmlSignedTwipsMeasure, required=True)


Spacing = SpacingRun


class W(SimpleValNode):
	tag = "w:w"
	val = AttrField('w:val', 'ooxml_attr_scale', required=True)


class Kern(SimpleValNode):
	tag = "w:kern"
	val = AttrField('w:val', OoxmlHpsMeasure, required=True)


class Position(SimpleValNode):
	tag = "w:position"
	val = AttrField('w:val', OoxmlSignedHpsMeasure, required=True)


class FontSize(SimpleValNode):
	tag = "w:sz"
	val = AttrField('w:val', OoxmlHpsMeasure, required=True)


class FontSizeCs(SimpleValNode):
	tag = "w:szCs"
	val = AttrField('w:val', OoxmlHpsMeasure, required=True)


class Highlight(SimpleValNode):
	tag = "w:highlight"
	enum = ST_HighlightColor
	val = AttrField('w:val', ('enum', {'enums': enum}), default=enum.NONE, required=True)


class Underline(OoxmlNode):
	tag = "w:u"
	enum = ST_Underline
	val = AttrField('w:val', ('enum', {'enums': ST_Underline}), default=ST_Underline.NONE, required=True)
	color = AttrField('w:color', 'ooxml_color', default=ColorAuto)


class BorderAttrFields(SharedBorder):
	tag = "w:bdr"


class FitText(SimpleValNode):
	tag = "w:fitText"
	val = AttrField('w:val', OoxmlTwipsMeasure, required=True)
	id = AttrField('w:id', 'string')


class VertAlign(SimpleValNode):
	tag = "w:vertAlign"
	enum = ST_VerticalAlignRun
	val = AttrField('w:val', ('enum', {'enums': enum}), default=ST_VerticalAlignRun.BASELINE, required=True)


class Em(SimpleValNode):
	tag = "w:em"
	enum = ST_Em
	val = AttrField('w:val', ('enum', {'enums': ST_Em}), default=enum.NONE, required=True)


class Lang(OoxmlNode):
	tag = "w:lang"
	val = AttrField('w:val', 'string')
	eastAsia = AttrField('w:eastAsia', 'string')
	bidi = AttrField('w:bidi', 'string')


class RunStyle(SimpleValNode):
	tag = "w:rStyle"
	val = AttrField('w:val', 'string', required=True)


class RunPr(OoxmlNode):
	tag = "w:rPr"
	style = SubField(RunStyle)
	fonts: Fonts = SubField(Fonts)
	bold: Bold = SubField(Bold)
	boldCs = SubField(BoldCs)
	italic = SubField(Italic)
	italicCs = SubField(ItalicCs)
	caps = SubField(Caps)
	smallCaps = SubField(SmallCaps)
	strike = SubField(Strike)
	doubleStrike = SubField(DoubleStrike)
	# outline shadow emboss imprint
	_effect = SubChoiceField({
		'outline': SubField(Outline),
		'shadow' : SubField(Shadow),
		'emboss' : SubField(Emboss),
		'imprint': SubField(Imprint)
	})
	
	noProof = SubField(NoProof)
	snapToGrid = SubField(SnapToGrid)
	vanish: Vanish = SubField(Vanish)
	webHidden = SubField(WebHidden)
	
	color = SubField(Color)
	spacing = SubField(Spacing)
	w = SubField(W)
	kern = SubField(Kern)
	position = SubField(Position)
	fontSize = SubField(FontSize)
	fontSizeCs = SubField(FontSizeCs)
	highlight = SubField(Highlight)
	underline: Underline = SubField(Underline)
	# effect
	border = SubField(BorderAttrFields)
	shade = SubField(Shading)
	fitText = SubField(FitText)
	vertAlign = SubField(VertAlign)
	# rtl cs
	em = SubField(Em)
	lang = SubField(Lang)
	
	# eastAsianLayout specVanish oMath
	
	def to_xml(self):
		res = super().to_xml()
		if res == '<w:rPr/>':
			return ''
		else:
			return res


class _TextSpace(EnumStrBase):
	PRESERVE = "preserve"
	DEFAULT = "default"


class Text(OoxmlNode):
	tag = "w:t"
	enum = _TextSpace
	space = AttrField("xml:space", ("enum", {"enums": enum}))
	
	def __init__(self, text: str = None , *args, **kwargs):
		super().__init__(*args, **kwargs)
		self.innerText = text


class Break(OoxmlNode):
	tag = "w:br"
	enum_type = ST_BrType
	enum_clear = ST_BrClear
	type = AttrField('w:type', ('enum', {'enums': enum_type}))
	clear = AttrField('w:type', ('enum', {'enums': enum_clear}))


NoBreakHyphen = create_empty_node("NoBreakHyphen", "w:noBreakHyphen")
SoftHyphen = create_empty_node("SoftHyphen", "w:softHyphen")
DayShort = create_empty_node("DayShort", "w:dayShort")
MonthShort = create_empty_node("MonthShort", "w:monthShort")
YearShort = create_empty_node("YearShort", "w:yearShort")
DayLong = create_empty_node("DayLong", "w:dayLong")
MonthLong = create_empty_node("MonthLong", "w:monthLong")
YearLong = create_empty_node("YearLong", "w:yearLong")
AnnotationRef = create_empty_node("AnnotationRef", "w:annotationRef")
EndnoteRef = create_empty_node("EndnoteRef", "w:endnoteRef")
Separator = create_empty_node("Separator", "w:separator")
ContinuationSeparator = create_empty_node("ContinuationSeparator", "w:continuationSeparator")
PageNum = create_empty_node("PageNum", "w:pgNum")
CarriageReturn = create_empty_node("CarriageReturn", "w:cr")
Tab = create_empty_node("Tab", "w:tab")
LastRenderedPageBreak = create_empty_node("LastRenderedPageBreak", "w:lastRenderedPageBreak")


class ContentPart(OoxmlNode):
	tag = "w:contentPart"
	rid = AttrField("r:id", "string")


class DelText(Text):
	tag = "w:delText"


class InstrText(Text):
	tag = "w:instrText"


class DelInstrText(Text):
	tag = "w:delInstrText"


class RunObject(OoxmlNode):
	tag = "w:object"
	# TODO: implement
	...


class RunPicture(OoxmlNode):
	tag = "w:pict"
	# TODO: implement
	...


class FieldChar(OoxmlNode):
	tag = "w:fldChar"
	# TODO: implement
	...


class Ruby(OoxmlNode):
	tag = "w:ruby"
	# TODO: implement
	...


class FootnoteReference(OoxmlNode):
	tag = "w:footnoteReference"
	# TODO: implement
	...


class EndnoteReference(OoxmlNode):
	tag = "w:endnoteReference"
	# TODO: implement
	...


class CommentReference(OoxmlNode):
	tag = "w:commentReference"
	# TODO: implement
	...


class Ptab(OoxmlNode):
	tag = "w:ptab"
	# TODO: implement
	...


class Run(OoxmlNode):
	tag = "w:r"
	
	rPr: RunPr = SubField(RunPr)
	_children_type = (Text,
			  DelText,
			  InstrText,
			  DelInstrText,
			  Break,
			  NoBreakHyphen,
			  SoftHyphen,
			  DayShort,
			  MonthShort,
			  YearShort,
			  DayLong,
			  MonthLong,
			  YearLong,
			  AnnotationRef,
			  EndnoteRef,
			  Separator,
			  ContinuationSeparator,
			  PageNum,
			  CarriageReturn,
			  Tab,
			  LastRenderedPageBreak,
			  ContentPart,
			  RunObject,
			  RunPicture,
			  FieldChar,
			  Ruby,
			  FootnoteReference,
			  EndnoteReference,
			  CommentReference,
			  Drawing,
			  Ptab
			  
			  )
	
	def __init__(self, content: str | OoxmlNode = None, rPr=None):
		if isinstance(rPr, str):
			rPr = {'style': rPr}
		super().__init__(rPr=rPr)
		if content is not None:
			if isinstance(content, Drawing):
				self.append(content)
			else:
				self.add_text(content)
	
	def add_text(self, _text: Text | str | Iterable):
		if isinstance(_text, Text):
			self.append(_text)
		elif isinstance(_text, str):
			if self._children and isinstance(self._children[-1], Text):
				self._children[-1]._text += _text
			else:
				self.append(Text(_text))
		elif isinstance(_text, Iterable):
			for t in _text:
				self.add_text(t)
		else:
			raise TypeError(f"Unsupported type: {type(_text)}")
		return self
	
	def add_image(self, image):  # TODO 类型检查
		self.append(image)
		return self

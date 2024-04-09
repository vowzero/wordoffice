from typing import Iterable, Optional

from wordoffice.core.field import AttrField, SubField
from wordoffice.core.node import OoxmlNode
from wordoffice.core.type import *
from .enums import *
from .run import Run, RunPr
from .section import SectionPr
from .utils import SimpleValNode, create_switch_node, EmptyHiddenNode


class Jc(SimpleValNode):
	tag = "w:jc"
	enum = ST_Jc
	val = AttrField('w:val', ('enum', {'enums': enum}), default=enum.BOTH, required=True)


class OutlineLevel(SimpleValNode, EmptyHiddenNode):
	tag = "w:outlineLvl"
	val = AttrField('w:val', ('integer_range', {'range': (0, 9)}), default='9', required=True)


class Indent(OoxmlNode):
	tag = "w:ind"
	start = AttrField('w:start', OoxmlSignedTwipsMeasure)
	startChars = AttrField('w:startChars', 'integer')
	end = AttrField('w:end', OoxmlSignedTwipsMeasure)
	endChars = AttrField('w:endChars', 'integer')
	left = AttrField('w:left', OoxmlSignedTwipsMeasure)
	leftChars = AttrField('w:leftChars', 'integer')
	right = AttrField('w:right', OoxmlSignedTwipsMeasure)
	rightChars = AttrField('w:rightChars', 'integer')
	hanging = AttrField('w:hanging', OoxmlTwipsMeasure)
	hangingChars = AttrField('w:hangingChars', 'integer')
	firstLine = AttrField('w:firstLine', OoxmlTwipsMeasure)
	firstLineChars = AttrField('w:firstLineChars', 'integer')


MirrorIndents = create_switch_node('MirrorIndents', 'w:mirrorIndents', False)
AdjustRightIndent = create_switch_node('AdjustRightIndent', 'w:adjustRightInd', True)
SnapToGrid = create_switch_node("SnapToGridPara", "w:snapToGrid", True)


class SpacingPara(OoxmlNode):
	tag = "w:spacing"
	enum = ST_LineSpacingRule
	before = AttrField('w:before', OoxmlTwipsMeasure, default='0')
	beforeLines = AttrField('w:beforeLines', 'integer', default='0')
	beforeAutospacing = AttrField('w:beforeAutospacing', 'ooxml_boolean', default=False)
	after = AttrField('w:after', OoxmlTwipsMeasure, default='0')
	afterLines = AttrField('w:afterLines', 'integer')
	afterAutospacing = AttrField('w:afterAutospacing', 'ooxml_boolean', default=False)
	line = AttrField('w:line', OoxmlSignedTwipsMeasure, default='0')
	lineRule = AttrField('w:lineRule', ('enum', {'enums': enum}), default=enum.AUTO)


Spacing = SpacingPara
ContextualSpacing = create_switch_node('ContextualSpacing', 'w:contextualSpacing', False)
WidowControl = create_switch_node('WidowControl', 'w:widowControl', False)
KeepNext = create_switch_node('KeepNext', 'w:keepNext', False)
KeepLines = create_switch_node('KeepLines', 'w:keepLines', False)
PageBreakBefore = create_switch_node('PageBreakBefore', 'w:pageBreakBefore', False)
SuppressAutoHyphens = create_switch_node('SuppressAutoHyphens', 'w:suppressAutoHyphens', False)
SuppressLineNumbers = create_switch_node('SuppressLineNumbers', 'w:suppressLineNumbers', False)
Kinsoku = create_switch_node('Kinsoku', 'w:kinsoku', True)
OverflowPunct = create_switch_node('OverflowPunct', 'w:overflowPunct', True)
WordWrap = create_switch_node('WordWrap', 'w:wordWrap', False)
AutoSpaceDE = create_switch_node('AutoSpaceDE', 'w:autoSpaceDE', True)
AutoSpaceDN = create_switch_node('AutoSpaceDN', 'w:autoSpaceDN', True)
TopLinePunct = create_switch_node('ToplinePunct', 'w:topLinePunct', False)


class TextAlignment(SimpleValNode):
	tag = "w:textAlignment"
	enum = ST_TextAlignment
	val = AttrField('w:val', ('enum', {'enums': enum}), default=enum.AUTO)


class ParagraphStyle(SimpleValNode):
	tag = "w:pStyle"
	val = AttrField('w:val', 'string', required=True)


class ParagraphPr(EmptyHiddenNode):
	tag = "w:pPr"
	style = SubField(ParagraphStyle)
	rPr = SubField(RunPr)
	jc = SubField(Jc)
	outlineLevel = SubField(OutlineLevel)
	indent = SubField(Indent)
	mirrorIndents = SubField(MirrorIndents)
	adjustRightIndent = SubField(AdjustRightIndent)
	snapToGrid = SubField(SnapToGrid)
	spacing = SubField(SpacingPara)
	contextualSpacing = SubField(ContextualSpacing)
	widowControl = SubField(WidowControl)
	keepNext = SubField(KeepNext)
	keepLines = SubField(KeepLines)
	pageBreakBefore = SubField(PageBreakBefore)
	suppressAutoHyphens = SubField(SuppressAutoHyphens)
	suppressLineNumbers = SubField(SuppressLineNumbers)
	kinsoku = SubField(Kinsoku)
	overflowPunct = SubField(OverflowPunct)
	wordWrap = SubField(WordWrap)
	autoSpaceDE = SubField(AutoSpaceDE)
	autoSpaceDN = SubField(AutoSpaceDN)
	topLinePunct = SubField(TopLinePunct)
	textAlignment = SubField(TextAlignment)
	
	sectPr = SubField(SectionPr)


class Paragraph(OoxmlNode):
	tag = 'w:p'
	
	pPr = SubField(ParagraphPr)
	
	def __init__(self, runs: Optional[str | Run | Iterable[Run | str]] = None, pPr: ParagraphPr | dict | str = None):
		if isinstance(pPr, str):
			pPr = {'style': pPr}
		super().__init__(pPr=pPr)
		if runs is not None:
			self.add_run(runs)
		self.isPlaceholder = False
	
	def add_run(self, _run: Run | str | Iterable):
		if isinstance(_run, str):
			self.append(Run(_run))
		elif isinstance(_run, Run):
			self.append(_run)
		elif isinstance(_run, Iterable):
			for t in _run:
				self.add_run(t)
		else:
			raise TypeError(f"Unsupported type: {type(_run)}")
		return self
	
	def to_xml(self):
		xml = super().to_xml()
		if self.isPlaceholder:
			self.pPr.sectPr = None
		return xml
	
	def set_temp_section(self, sectPr):
		self.isPlaceholder = True
		if self.pPr is None:
			self.pPr = ParagraphPr(sectPr=sectPr)
		else:
			self.pPr.sectPr = sectPr

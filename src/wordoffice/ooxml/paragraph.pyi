from typing import Iterable, Optional

from wordoffice.core.node import OoxmlNode
from wordoffice.core.type import *
from wordoffice.ooxml.enums import *
from wordoffice.ooxml.run import Run
from wordoffice.ooxml.section import SectionPr
from wordoffice.ooxml.utils import SimpleValNode, EmptyHiddenNode, SwitchNode
from wordoffice.common import Measure


class Jc(SimpleValNode):
	enum = ST_Jc
	val: ST_Jc | str


class OutlineLevel(SimpleValNode, EmptyHiddenNode):
	val: int


class Indent(OoxmlNode):
	start: Measure | str | int | float
	startChars: int
	end: Measure | str | int | float
	endChars: int
	hanging: Measure | str | int | float
	hangingChars: int
	firstLine: Measure | str | int | float
	firstLineChars: int
	left : Measure | str | int | float
	leftChars: int
	right : Measure | str | int | float
	rightChars : int


class MirrorIndents(SwitchNode): ...


class AdjustRightIndent(SwitchNode): ...


class SnapToGrid(SwitchNode): ...


class SpacingPara(OoxmlNode):
	enum = ST_LineSpacingRule
	before: Measure | str | int | float
	beforeLines: int | str
	beforeAutospacing: bool
	after: Measure | str | int | float
	afterLines: int | str
	afterAutospacing: bool
	line: Measure | str | int | float
	lineRule: ST_LineSpacingRule | str


Spacing = SpacingPara


class ContextualSpacing(SwitchNode): ...


class WidowControl(SwitchNode): ...


class KeepNext(SwitchNode): ...


class KeepLines(SwitchNode): ...


class PageBreakBefore(SwitchNode): ...


class SuppressAutoHyphens(SwitchNode): ...


class SuppressLineNumbers(SwitchNode): ...


class Kinsoku(SwitchNode): ...


class OverflowPunct(SwitchNode): ...


class WordWrap(SwitchNode): ...


class AutoSpaceDE(SwitchNode): ...


class AutoSpaceDN(SwitchNode): ...


class TopLinePunct(SwitchNode): ...


class TextAlignment(SimpleValNode):
	enum = ST_TextAlignment
	val: ST_TextAlignment | str


class ParagraphStyle(SimpleValNode):
	val: str


class ParagraphPr(EmptyHiddenNode):
	def __init__(self, *args, **kwargs):
		super().__init__(args, kwargs)
	
	jc: Jc
	style: ParagraphStyle
	outlineLevel: OutlineLevel
	indent: Indent
	mirrorIndents: MirrorIndents
	adjustRightIndent: AdjustRightIndent
	snapToGrid: SnapToGrid
	spacing: SpacingPara
	contextualSpacing: ContextualSpacing
	widowControl: WidowControl
	keepNext: KeepNext
	keepLines: KeepLines
	pageBreakBefore: PageBreakBefore
	suppressAutoHyphens: SuppressAutoHyphens
	suppressLineNumbers: SuppressLineNumbers
	kinsoku: Kinsoku
	overflowPunct: OverflowPunct
	wordWrap: WordWrap
	autoSpaceDE: AutoSpaceDE
	autoSpaceDN: AutoSpaceDN
	topLinePunct: TopLinePunct
	textAlignment: TextAlignment
	
	sectPr: SectionPr


class Paragraph(OoxmlNode):
	pPr: ParagraphPr
	
	def __init__(self, runs: Optional[str | Run | Iterable[Run | str]] = None, pPr: ParagraphPr | dict | str = None):
		"""
		12323213
		"""
		...
	
	def add_run(self, _run: Run | str | Iterable) -> Paragraph: ...
	
	def set_temp_section(self, sectPr): ...

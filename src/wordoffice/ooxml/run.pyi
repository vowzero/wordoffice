from typing import Iterable

from core.node import OoxmlNode
from .enums import *
from .shared import Shading, BorderAttrFields as SharedBorder
from .utils import SimpleValNode, SwitchNode
from ..common import Measure


class Bold(SwitchNode): ...


class BoldCs(SwitchNode): ...


class Italic(SwitchNode): ...


class ItalicCs(SwitchNode): ...


class Caps(SwitchNode): ...


class SmallCaps(SwitchNode): ...


class Strike(SwitchNode): ...


class DoubleStrike(SwitchNode): ...


class Outline(SwitchNode): ...


class Shadow(SwitchNode): ...


class Emboss(SwitchNode): ...


class Imprint(SwitchNode): ...


class NoProof(SwitchNode): ...


class SnapToGrid(SwitchNode): ...


class Vanish(SwitchNode): ...


class WebHidden(SwitchNode): ...


class Fonts(OoxmlNode):
	enum = ST_Hint
	ascii: str
	hAnsi: str
	eastAsia: str
	cs: str
	hint: ST_Hint | str
	asciiTheme: str
	hAnsiTheme: str
	eastAsiaTheme: str
	cstheme: str


class Color(SimpleValNode):
	val: str


class SpacingRun(SimpleValNode):
	val: str


Spacing = SpacingRun


class W(SimpleValNode):
	val: str


class Kern(SimpleValNode):
	val: str


class Position(SimpleValNode):
	val: str


class FontSize(SimpleValNode):
	val: str


class FontSizeCs(SimpleValNode):
	val: str


class Highlight(SimpleValNode):
	enum = ST_HighlightColor
	val: ST_HighlightColor | str


class Underline(OoxmlNode):
	enum = ST_Underline
	val: str | ST_Underline
	color: str


class BorderAttrFields(SharedBorder): ...


class FitText(SimpleValNode):
	val: Measure | str | int | float
	id: str


class VertAlign(SimpleValNode):
	enum = ST_VerticalAlignRun
	val: str | ST_VerticalAlignRun


class Em(SimpleValNode):
	enum = ST_Em
	val: ST_Em | str


class Lang(OoxmlNode):
	val: str
	eastAsia: str
	bidi: str


class RunPr(OoxmlNode):
	tag = "w:rPr"
	# rStyle
	
	fonts: Fonts
	bold: Bold
	boldCs: BoldCs
	italic: Italic
	italicCs: ItalicCs
	caps: Caps
	smallCaps: SmallCaps
	strike: Strike
	doubleStrike: DoubleStrike
	
	outline: Outline
	shadow: Shadow
	emboss: Emboss
	imprint: Imprint
	
	noProof: NoProof
	snapToGrid: SnapToGrid
	vanish: Vanish:Vanish
	webHidden: WebHidden
	
	color: Color
	spacing: Spacing
	w: W
	kern: Kern
	position: Position
	fontSize: FontSize
	fontSizeCs: FontSizeCs
	highlight: Highlight
	underline: Underline:Underline
	# effect
	border: BorderAttrFields
	shade: Shading
	fitText: FitText
	vertAlign: VertAlign
	# rtl cs
	em: Em
	lang: Lang


class Text(OoxmlNode):
	
	def __init__(self, text: str = None, *args, **kwargs): ...
	
	@property
	def innerText(self) -> str | None: ...
	
	@innerText.setter
	def innerText(self, value): ...


class NoBreakHyphen(OoxmlNode): ...


class SoftHyphen(OoxmlNode): ...


class DayShort(OoxmlNode): ...


class MonthShort(OoxmlNode): ...


class YearShort(OoxmlNode): ...


class DayLong(OoxmlNode): ...


class MonthLong(OoxmlNode): ...


class YearLong(OoxmlNode): ...


class AnnotationRef(OoxmlNode): ...


class EndnoteRef(OoxmlNode): ...


class Separator(OoxmlNode): ...


class ContinuationSeparator(OoxmlNode): ...


class PageNum(OoxmlNode): ...


class CarriageReturn(OoxmlNode): ...


class Tab(OoxmlNode): ...


class LastRenderedPageBreak(OoxmlNode): ...


class Run(OoxmlNode):
	rPr: RunPr:RunPr
	
	def __init__(self, text: str | OoxmlNode  = None, rPr=None):...
	
	def add_text(self, _text: Text | str | Iterable) -> Run: ...
	
	def add_image(self, image) -> Run: ...


class Break(OoxmlNode):...
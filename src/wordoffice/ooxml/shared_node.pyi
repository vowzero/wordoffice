from core.node import OoxmlNode
from wordoffice.common import Measure
from wordoffice.ooxml.enums import ST_Shd, ST_Border


class Shading(OoxmlNode):
	enum = ST_Shd
	val: ST_Shd | str
	color: str
	fill: str


class BorderAttrFields(OoxmlNode):
	enum = ST_Border
	val: ST_Border | str
	color: str
	sz: Measure | str | int | float
	space: Measure | str | int | float
	frame: bool
	shadow: bool

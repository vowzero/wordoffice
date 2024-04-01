from wordoffice.core.field import AttrField
from wordoffice.core.node import OoxmlNode
from wordoffice.core.type import OoxmlEightPointMeasure, OoxmlPointMeasure
from wordoffice.common import ColorAuto, Measure
from wordoffice.ooxml.enums import ST_Shd, ST_Border
from wordoffice.ooxml.utils import SimpleValNode


class Shading(OoxmlNode):
	tag = "w:shd"
	enum = ST_Shd
	val = AttrField('w:val', ('enum', {'enums': enum}), required=True)
	color = AttrField('w:color', 'ooxml_color', default=ColorAuto)
	fill = AttrField('w:fill', 'ooxml_color', default=ColorAuto)


class BorderAttrFields(SimpleValNode):
	enum = ST_Border
	val = AttrField('w:val', ('enum', {'enums': ST_Border}), default=ST_Border.NONE, required=True)
	color = AttrField('w:color', 'ooxml_color', default=ColorAuto)
	sz = AttrField('w:sz', OoxmlEightPointMeasure, default=Measure.PT_8(4))
	space = AttrField('w:space', OoxmlPointMeasure, default=Measure.PT(0))
	frame = AttrField('w:frame', 'ooxml_boolean', default=False)
	shadow = AttrField('w:shadow', 'ooxml_boolean', default=False)
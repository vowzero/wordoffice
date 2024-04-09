from wordoffice.core.field import SubField, AttrField, SubChoiceField
from wordoffice.core.node import OoxmlNode
from wordoffice.ooxml.enums import *
from wordoffice.common import MeasureZero, Measure
from wordoffice.core.type import OoxmlPointMeasure, OoxmlEmusMeasure, OoxmlSignedPointMeasure


class _CT_NonVisualDrawingProps(OoxmlNode):
	
	id = AttrField('id', 'integer_unsigned', required=True)
	name = AttrField('name', 'string', required=True)
	descr = AttrField('descr', 'string', default='')
	hidden = AttrField('hidden', 'xsd_boolean', default=False)
	title = AttrField('title', 'string', default='')


# _children_type =(hlinkClick, hlinkHover, extLst,) # TODO: 支持替代文本超链接


class DocPr(_CT_NonVisualDrawingProps):
	tag = 'wp:docPr'


class CNvPr(_CT_NonVisualDrawingProps):
	tag = 'pic:cNvPr'


	

class CNvPicPr(OoxmlNode):
	tag = 'pic:cNvPicPr'
	'''
	  <xsd:complexType name="CT_NonVisualPictureProperties">
    <xsd:sequence>
      <xsd:element name="picLocks" "CT_PictureLocking" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" "CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst1" "CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="preferRelativeResize" "xsd:boolean" use="optional" default="true"/>
  </xsd:complexType>'''
	preferRelativeResize = AttrField('preferRelativeResize', 'ooxml_boolean', default=True)


# picLocks = OoxmlSubField(PicLocks)
# extLst = OoxmlSubField(ExtLst)


class NvPicPr(OoxmlNode):
	tag = 'pic:nvPicPr'
	"""
	    <xsd:sequence>
      <xsd:element name="cNvPr" "a:CT_NonVisualDrawingProps" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="cNvPicPr" "a:CT_NonVisualPictureProperties" minOccurs="1"
	maxOccurs="1"/>
    </xsd:sequence>
    """
	cNvPr = SubField(CNvPr)
	cNvPicPr = SubField(CNvPicPr)


# extLst = OoxmlSubField(ExtLst)


class Blip(OoxmlNode):
	tag = 'a:blip'
	
	enum = ST_BlipCompression
	
	embed = AttrField('r:embed', 'string')  # rid
	link = AttrField('r:link', 'string')  # rid
	cstate = AttrField('cstate', ('enum', {'enums': enum}), default=enum.NONE)
	'''
    <xsd:sequence>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
	<xsd:element name="alphaBiLevel" "CT_AlphaBiLevelEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="alphaCeiling" "CT_AlphaCeilingEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="alphaFloor" "CT_AlphaFloorEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="alphaInv" "CT_AlphaInverseEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="alphaMod" "CT_AlphaModulateEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="alphaModFix" "CT_AlphaModulateFixedEffect" minOccurs="1"
	  maxOccurs="1"/>
	<xsd:element name="alphaRepl" "CT_AlphaReplaceEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="biLevel" "CT_BiLevelEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="blur" "CT_BlurEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="clrChange" "CT_ColorChangeEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="clrRepl" "CT_ColorReplaceEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="duotone" "CT_DuotoneEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="fillOverlay" "CT_FillOverlayEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="grayscl" "CT_GrayscaleEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="hsl" "CT_HSLEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="lum" "CT_LuminanceEffect" minOccurs="1" maxOccurs="1"/>
	<xsd:element name="tint" "CT_TintEffect" minOccurs="1" maxOccurs="1"/>
      </xsd:choice>
      <xsd:element name="extLst" "CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>

    <xsd:attribute name="cstate" "ST_BlipCompression" use="optional" default="none"/>
	'''


class _CT_RelativeRect(OoxmlNode):
	l = AttrField('l', 'integer', default=0)  # TODO: ST_Percentage
	t = AttrField('t', 'integer', default=0)
	r = AttrField('r', 'integer', default=0)
	b = AttrField('b', 'integer', default=0)


class SrcRect(_CT_RelativeRect):
	tag = 'a:srcRect'


class Tile(OoxmlNode):
	tag = 'a:tile'
	
	enum_flip = ST_TileFlipMode
	enum_algn = ST_RectAlignment
	tx = AttrField('tx', 'integer')  # TODO: ST_Coordinate
	ty = AttrField('ty', 'integer')
	sx = AttrField('sx', 'integer')  # TODO: ST_Percentage
	sy = AttrField('sy', 'integer')
	flip = AttrField('flip', ('enum', {'enums': enum_flip}), default=enum_flip.NONE)
	algn = AttrField('algn', ('enum', {'enums': enum_algn}))


class FillRect(_CT_RelativeRect):
	tag = 'a:fillRect'


class Stretch(OoxmlNode):
	tag = 'a:stretch'
	
	fillRect = SubField(FillRect)


class BlipFill(OoxmlNode):
	tag = 'pic:blipFill'
	
	dpi = AttrField('dpi', 'integer_unsigned')
	rotWithShape = AttrField('rotWithShape', 'ooxml_boolean')
	
	blip = SubField(Blip)
	srcRect = SubField(SrcRect)
	
	# 二选一
	tile = SubField(Tile)
	stretch = SubField(Stretch)
	
	def to1_xml(self):
		return f'''
                                
                            <pic:blipFill>
                                <a:blip r:embed="rId1" cstate="print">
                                    <a:extLst>
                                        <a:ext uri="{{28A0092B-C50C-407E-A947-70E740481C1C}}">
                                            <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                        </a:ext>
                                    </a:extLst>
                                </a:blip>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                                    
                '''


class ShapePrTransformOffset(OoxmlNode):
	tag = 'a:off'
	
	x = AttrField('x', OoxmlPointMeasure, required=True)
	y = AttrField('y', OoxmlPointMeasure, required=True)


class ShapePrTransformExtent(OoxmlNode):
	tag = 'a:ext'
	
	cx = AttrField('cx', OoxmlEmusMeasure, required=True)
	cy = AttrField('cy', OoxmlEmusMeasure, required=True)


class ShapePrTransform(OoxmlNode):
	tag = 'a:xfrm'
	
	rot = AttrField('rot', 'integer', default=0)
	flipH = AttrField('flipH', 'xsd_boolean', default=False)
	flipV = AttrField('flipV', 'xsd_boolean', default=False)
	
	off = SubField(ShapePrTransformOffset)
	ext = SubField(ShapePrTransformExtent)


class AvLst(OoxmlNode):
	tag = 'a:avLst'


# TODO


class PrstGeom(OoxmlNode):
	tag = 'a:prstGeom'
	enum = ST_ShapeType
	prst = AttrField('prst', ('enum', {'enums': enum}), required=True)
	avLst = SubField(AvLst)


# TODO


class SpPr(OoxmlNode):
	tag = 'pic:spPr'
	'''
	  <xsd:complexType name="CT_ShapeProperties">
    <xsd:sequence>
      <xsd:group ref="EG_Geometry" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_FillProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="ln" "CT_LineProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_EffectProperties" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="scene3d" "CT_Scene3D" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sp3d" "CT_Shape3D" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" "CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" "ST_BlackWhiteMode" use="optional"/>
  </xsd:complexType>'''
	bwMode = AttrField('bwMode', ('enum', {'enums': ST_BlackWhiteMode}))
	
	xfrm = SubField(ShapePrTransform)
	prstGeom = SubField(PrstGeom)
	
	def to1_xml(self):
		return f'''

                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="3023178" cy="2015331"/>
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst/>
                                        </a:prstGeom>
                                    </pic:spPr>

                '''


class Picture(OoxmlNode):
	tag = 'pic:pic'
	
	nvPicPr = SubField(NvPicPr)
	blipFill = SubField(BlipFill)
	spPr = SubField(SpPr)
	
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._attrs.update({
			'xmlns:pic': "http://schemas.openxmlformats.org/drawingml/2006/picture",
		})


class GraphicData(OoxmlNode):
	tag = 'a:graphicData'
	
	uri = AttrField('uri', 'string', required=True)  # TODO  Ooxml Token


# _children_type =(chart, externalData, graphic, legacy, rels,)


class Graphic(OoxmlNode):
	tag = 'a:graphic'
	
	graphicData = SubField(GraphicData)
	
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._attrs.update({
			'xmlns:a': "http://schemas.openxmlformats.org/drawingml/2006/main",
		})


class Extent(OoxmlNode):
	"""
		指定 Drawing Object 在Word中占位框大小
	"""
	tag = 'wp:extent'
	cx = AttrField('cx', OoxmlEmusMeasure, required=True)
	cy = AttrField('cy', OoxmlEmusMeasure, required=True)


class EffectExtent(OoxmlNode):
	tag = 'wp:extentExtent'


class DrawingInline(OoxmlNode):
	tag = 'wp:inline'
	
	distT = AttrField('distT', OoxmlEmusMeasure, default=MeasureZero)
	distB = AttrField('distB', OoxmlEmusMeasure, default=MeasureZero)
	distL = AttrField('distL', OoxmlEmusMeasure, default=MeasureZero)
	distR = AttrField('distR', OoxmlEmusMeasure, default=MeasureZero)
	
	extent = SubField(Extent, index=1)
	effectExtent = SubField(EffectExtent, index=2)
	docPr = SubField(DocPr, index=3)
	graphic: Graphic = SubField(Graphic, index=4)


class SimplePos(OoxmlNode):
	tag = 'wp:simplePos'
	x = AttrField('x', OoxmlSignedPointMeasure, required=True)
	y = AttrField('y', OoxmlSignedPointMeasure, required=True)


class PositionOffset(OoxmlNode):
	tag = 'wp:posOffset'
	
	def __init__(self, val:Measure=MeasureZero, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._val = val
	
	def to_xml(self):
		return f'<{self.tag_name}>{self._val.emus}</{self.tag_name}>'


class PositionAlignH(OoxmlNode):
	tag = 'wp:align'
	enum = ST_AlignH
	
	def __init__(self, val=None, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._val = val
	
	def to_xml(self):
		if self._val is not None:
			self._text = str(self._val)
		return super().to_xml()


class PositionH(OoxmlNode):
	tag = 'wp:positionH'
	enum = ST_RelFromH
	enum_align = ST_AlignH
	relativeFrom = AttrField('relativeFrom', ('enum', {'enums': enum}), required=True)
	align = SubField(PositionAlignH)
	posOffset = SubField(PositionOffset)


class PositionAlignV(OoxmlNode):
	tag = 'wp:align'
	enum = ST_AlignV
	
	def __init__(self, val=None, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._val = val
	
	def to_xml(self):
		if self._val is not None:
			self._text = str(self._val)
		return super().to_xml()


class PositionV(OoxmlNode):
	tag = 'wp:positionV'
	enum = ST_RelFromV
	
	relativeFrom = AttrField('relativeFrom', ('enum', {'enums': enum}), required=True)
	align = SubField(PositionAlignV)
	posOffset = SubField(PositionOffset)


class WrapNone(OoxmlNode):
	tag = 'wp:wrapNone'


class WrapSquare(OoxmlNode):
	tag = 'wp:wrapSquare'


class WrapTight(OoxmlNode):
	tag = 'wp:wrapTight'


class WrapThrough(OoxmlNode):
	tag = 'wp:wrapThrough'


class WrapTopBottom(OoxmlNode):
	tag = 'wp:wrapTopAndBottom'


class DrawingAnchor(OoxmlNode):
	tag = 'wp:anchor'
	
	distT: int = AttrField('distT', OoxmlEmusMeasure, default=MeasureZero)
	distB: int = AttrField('distB', OoxmlEmusMeasure, default=MeasureZero)
	distL: int = AttrField('distL', OoxmlEmusMeasure, default=MeasureZero)
	distR: int = AttrField('distR', OoxmlEmusMeasure, default=MeasureZero)
	simplePosSwitch: bool = AttrField('simplePos', 'xsd_boolean')
	relativeHeight: int = AttrField('relativeHeight', 'integer_unsigned', required=True)
	behindDoc: bool = AttrField('behindDoc', 'xsd_boolean', required=True)
	locked: bool = AttrField('locked', 'xsd_boolean', required=True)
	layoutInCell: bool = AttrField('layoutInCell', 'xsd_boolean', required=True)
	hidden: bool = AttrField('hidden', 'xsd_boolean')
	allowOverlap: bool = AttrField('allowOverlap', 'xsd_boolean', required=True)
	
	simplePos = SubField(SimplePos, index=1)
	positionH = SubField(PositionH, index=2)
	positionV = SubField(PositionV, index=3)
	extent = SubField(Extent, index=4)
	# effectExtent = OoxmlSubField(EffectExtent)
	# EG_WrapType
	wrap = SubChoiceField({
		'wrapNone'        : SubField(WrapNone),
		'wrapSquare'      : SubField(WrapSquare),
		'wrapTight'       : SubField(WrapTight),
		'wrapThrough'     : SubField(WrapThrough),
		'wrapTopAndBottom': SubField(WrapTopBottom),
	}, index=5, required=True)
	
	docPr = SubField(DocPr, index=6)
	graphic: Graphic = SubField(Graphic, index=7)


class Drawing(OoxmlNode):
	tag = 'w:drawing'
	
	inline = SubField(DrawingInline)
	anchor = SubField(DrawingAnchor)
	
	def to_xml(self):
		if self.inline and self.anchor:
			assert False, "inline和anchor只能有一个"
		return super().to_xml()

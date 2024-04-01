from core import opc
from core.field import AttrField, SubField
from core.node import OoxmlNode, OoxmlNodeBase
from core.opc import NodePart
from .enums import *
from .utils import create_switch_node, xml_header, SimpleValNode


class FontAltName(SimpleValNode):
	tag = "w:altName"
	val = AttrField('w:val', 'string', required=True)


class FontPanose1(SimpleValNode):
	tag = "w:panose1"
	val = AttrField('w:val', 'string', required=True)  # TODO:CT_Panose


class FontCharset(SimpleValNode):
	tag = "w:charset"
	val = AttrField('w:val', 'string')  # TODO: OoxmlUcharHexNumber
	characterSet = AttrField('w:characterSet', 'string', default="ISO-8859-1")


class FontFamily(SimpleValNode):
	tag = "w:family"
	enum = ST_FontFamily
	val = AttrField('w:val', ('enum', {'enums': enum}), required=True)


FontNotTrueType = create_switch_node("FontNotTrueType", "w:notTrueType", False)


class FontPitch(SimpleValNode):
	tag = "w:pitch"
	enum = ST_FontPitch
	val = AttrField('w:val', ('enum', {'enums': enum}), required=True)


class FontSig(OoxmlNode):
	tag = "w:sig"
	usb0 = AttrField('w:usb0', 'ooxml_long_hex_number', required=True)
	usb1 = AttrField('w:usb1', 'ooxml_long_hex_number', required=True)
	usb2 = AttrField('w:usb2', 'ooxml_long_hex_number', required=True)
	usb3 = AttrField('w:usb3', 'ooxml_long_hex_number', required=True)
	csb0 = AttrField('w:csb0', 'ooxml_long_hex_number', required=True)
	csb1 = AttrField('w:csb1', 'ooxml_long_hex_number', required=True)


class _FontRel(OoxmlNode):
	rId = AttrField("r:id", 'string', required=True)
	fontKey = AttrField("w:fontKey", 'string')  # TODO:ST_Guid
	subsetted = AttrField("w:subsetted", 'ooxml_boolean')


class FontEmbedRegular(_FontRel):
	tag = "w:embedRegular"


class FontEmbedBold(_FontRel):
	tag = "w:embedBold"


class FontEmbedItalic(_FontRel):
	tag = "w:embedItalic"


class FontEmbedBoldItalic(_FontRel):
	tag = "w:embedBoldItalic"


class Font(OoxmlNode):
	tag = "w:font"
	name = AttrField("w:name", 'string', required=True)
	
	altName = SubField(FontAltName)
	panose1 = SubField(FontPanose1)
	charset = SubField(FontCharset)
	family = SubField(FontFamily)
	notTrueType = SubField(FontNotTrueType)
	pitch = SubField(FontPitch)
	sig = SubField(FontSig)
	embedRegular = SubField(FontEmbedRegular)
	embedBold = SubField(FontEmbedBold)
	embedItalic = SubField(FontEmbedItalic)
	embedBoldItalic = SubField(FontEmbedBoldItalic)


class Fonts(OoxmlNode):
	tag = "w:fonts"
	_children_type = (Font,)


class FontTable(Fonts, NodePart):
	opc_info = opc.OPC_BUILTIN_MAP['fontTable']
	namespace = {
		'xmlns:mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
		'xmlns:r' : "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		'xmlns:m' : "http://schemas.openxmlformats.org/officeDocument/2006/math",
		'xmlns:wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		'xmlns:w' : "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	}
	
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		super(OoxmlNodeBase, self).__init__(self, self.opc_info)
		self._attrs.update(self.namespace)
	
	def to_xml(self):
		res = xml_header + super().to_xml()
		return res

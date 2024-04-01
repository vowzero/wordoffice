"""
        ooxml中OPC部分实现
"""

from wordoffice.ooxml.enums import EnumStrBase
from wordoffice.driver import FileDriver


class ContentTypeEnum(EnumStrBase):
	DOC_APP = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	DOC_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
	WORD_DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	WORD_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
	WORD_SETTINGS = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
	WORD_WEBSETTINGS = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
	WORD_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
	WORD_ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
	WORD_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	WORD_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	WORD_FONTTABLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
	WORD_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"


class RelationshipTypeEnum(EnumStrBase):
	WORD_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	WORD_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	WORD_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	WORD_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	WORD_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	WORD_CUSTOMXML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
	WORD_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
	WORD_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	WORD_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
	WORD_FONTTABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
	WORD_WEBSETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
	
	MEDIA_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	
	APP_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	APP_APP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"


OPC_BUILTIN_MAP = {
	'app'        : ('docProps/app.xml', ContentTypeEnum.DOC_APP, RelationshipTypeEnum.APP_APP),
	'core'       : ('docProps/core.xml', ContentTypeEnum.DOC_CORE, RelationshipTypeEnum.APP_CORE),
	'document'   : ('word/document.xml', ContentTypeEnum.WORD_DOCUMENT, RelationshipTypeEnum.WORD_DOCUMENT),
	'styles'     : ('styles.xml', ContentTypeEnum.WORD_STYLES, RelationshipTypeEnum.WORD_STYLES),
	'settings'   : ('settings.xml', ContentTypeEnum.WORD_SETTINGS, RelationshipTypeEnum.WORD_SETTINGS),
	'webSettings': ('webSettings.xml', ContentTypeEnum.WORD_WEBSETTINGS, RelationshipTypeEnum.WORD_WEBSETTINGS),
	'fontTable'  : ('fontTable.xml', ContentTypeEnum.WORD_FONTTABLE, RelationshipTypeEnum.WORD_FONTTABLE),
	'theme'      : ('theme/theme{}.xml', ContentTypeEnum.WORD_THEME, RelationshipTypeEnum.WORD_THEME),
	'header'     : ('header{}.xml', ContentTypeEnum.WORD_HEADER, RelationshipTypeEnum.WORD_HEADER),
	'footer'     : ('footer{}.xml', ContentTypeEnum.WORD_FOOTER, RelationshipTypeEnum.WORD_FOOTER),
}


class PartBase:
	def __init__(self, pack_uri, content_type, rel_type):
		
		self._pack_uri: str = pack_uri
		self._content_type = content_type
		self._rel_type = rel_type
		self._part_rid = None
	
	@property
	def filename(self):
		"""获取文件名"""
		return self._pack_uri.split('/')[-1]
	
	@property
	def ext(self):
		"""获取文件后缀名"""
		return self.filename.split('.')[-1]
	
	@property
	def content_type(self):
		"""获取内容类型(Content_Type)"""
		return self._content_type
	
	@property
	def rel_type(self):
		"""获取关系类型(Relationship_Type)"""
		return self._rel_type
	
	@property
	def pack_uri(self):
		"""获取包内URI"""
		return self._pack_uri
	
	@property
	def physical_uri(self):
		"""获取物理URI"""
		return self._pack_uri[1:]
	
	@property
	def part_rid(self):
		"""获取关系ID，未被添加到关系时不存在"""
		if self._part_rid is None:
			raise Exception("Part_rid is None")
		return self._part_rid
	
	@part_rid.setter
	def part_rid(self, value):
		
		self._part_rid = value
	
	def write_to(self, writer: FileDriver):
		"""写出到Writer"""
		raise NotImplementedError()
	
	@pack_uri.setter
	def pack_uri(self, value):
		
		self._pack_uri = value


class FilePart(PartBase):
	"""文件类型的Part"""
	
	def __init__(self, filename: str, opc_info):
		super().__init__(*opc_info)
		self._filename = filename
	
	def write_to(self, writer: FileDriver):
		writer.write_file(self.physical_uri, self._filename)


class StringPart(PartBase):
	"""字符串类型的Part"""
	
	def __init__(self, content: str, opc_info):
		super().__init__(*opc_info)
		self._content = content
	
	def write_to(self, writer: FileDriver):
		writer.write(self.physical_uri, self._content)


class NodePart(PartBase):
	"""OoxmlNode类型的Part"""
	
	def __init__(self, node, opc_info):
		super().__init__(*opc_info)
		self._node = node
	
	def write_to(self, writer: FileDriver):
		writer.write(self.physical_uri, self._node.to_xml())


class ContentTypes(PartBase):
	"""[Content_Types].xml管理类"""
	
	def __init__(self):
		super().__init__('/[Content_Types].xml', None, None)
		self._pack_parts: list[PartBase] = []
		self._default_parts: list[str] = [
			'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
			'<Default Extension="xml" ContentType="application/xml"/>'
		]
	
	def add_part(self, part: PartBase):
		"""添加Part到Content_Types"""
		self._pack_parts.append(part)
	
	def to_xml(self):
		"""生成xml"""
		default_nodes: list[str] = self._default_parts.copy()
		override_nodes: list[str] = []
		xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
		xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
		for item in self._pack_parts:
			content_type = item.content_type
			if isinstance(content_type, (str, ContentTypeEnum)):
				override_nodes.append(f'<Override PartName="{item.pack_uri}" ContentType="{content_type}"/>')
			elif isinstance(content_type, tuple):
				# noinspection PyTypeChecker
				default_nodes.append(f'<Default Extension="{content_type[0]}" ContentType="{content_type[1]}"/>')
		xml += ''.join(default_nodes)
		xml += ''.join(override_nodes)
		xml += '</Types>'
		
		return xml
	
	def write_to(self, writer: FileDriver):
		writer.write(self.physical_uri, self.to_xml())


class Relationships(PartBase):
	opc_content_type = "application/vnd.openxmlformats-package.relationships+xml"
	
	def __init__(self, pack_uri, basename='/'):
		super().__init__(pack_uri, self.opc_content_type, None)
		self._pack_parts: list[PartBase] = []
		self.basename = basename
		self.part_count = 0
		self.same_part_counter = {}
	
	def add_part(self, part: PartBase):
		"""添加Part到_rels并赋予rid"""
		origin_pack_uri = part.pack_uri
		same_name_counter = self.same_part_counter.setdefault(origin_pack_uri, 0)
		new_counter = self.same_part_counter[origin_pack_uri] = same_name_counter + 1
		self.part_count += 1
		
		new_uri = self.basename + origin_pack_uri
		
		if "{}" in new_uri:
			new_uri = new_uri.format(new_counter)
		
		part.pack_uri = new_uri
		part.part_rid = f"rId{self.part_count}"
		
		self._pack_parts.append(part)
	
	def to_xml(self):
		"""生成xml"""
		basename_length = len(self.basename)
		
		xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
		xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		for part in self._pack_parts:
			xml += f'<Relationship Id="{part.part_rid}" Type="{part.rel_type}" Target="{part.pack_uri[basename_length:]}"/>'
		xml += '</Relationships>'
		return xml
	
	def write_to(self, writer: FileDriver):
		for part in self._pack_parts:
			part.write_to(writer)
		writer.write(self.physical_uri, self.to_xml())

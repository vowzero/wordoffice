import os
import zipfile
from typing import Iterable, Optional

from wordoffice.common import Measure
from wordoffice.core.node import OoxmlNode
from wordoffice.ooxml.paragraph import Paragraph
from wordoffice.ooxml.table import Table
from wordoffice.core import opc
from wordoffice.driver import FileDriver
from wordoffice.ooxml.section import Section, DocGrid, PageSize, SectionPr, PageMargin, Cols

template_base = os.path.join(os.path.dirname(os.path.realpath(__file__)), '../templates/normal')
template_path = os.path.join(template_base, './word')


class Docx(OoxmlNode):
	tag = 'w:body'
	
	def __init__(self, default_template_dir_path: Optional[str] = None):
		super().__init__()
		self._children.append(Section())
		
		self._app_rels = opc.Relationships('/_rels/.rels')
		self._word_rels = opc.Relationships('/word/_rels/document.xml.rels', '/word/')
		self._content_types = opc.ContentTypes()
		self._init_body()
		
		self._custom_templates_config: list[str | dict[str, None | str]] = [template_path, {'styles': None, 'settings': None, 'webSettings': None, 'fontTable': None}]
		if default_template_dir_path is not None:
			self.set_template_path(default_template_dir_path)
		
		self._driver = FileDriver
	
	@property
	def lastSection(self)->Section:
		"""获取最后一个节，文档默认拥有一个节"""
		return self._children[-1]
	
	def add_p(self, paragraph: Paragraph | Iterable[Paragraph|Iterable|str]):
		"""添加段落"""
		self._children[-1].add_p(paragraph)
	
	def add_table(self, table: Table):
		"""添加表格"""
		self._children[-1].add_table(table)
	
	def add_section(self, section: Section):
		"""添加节，在最一个节（默认的文档节）之前添加节"""
		self._children.insert(-1, section)
	
	def set_template_path(self, path: str):
		"""设置模板路径，默认为内部模板路径"""
		self._custom_templates_config[0] = path
	
	def set_template(self, part_name: str, part_path: str | opc.PartBase):
		"""为单独部分独立设置模板\n
		:param part_name: part名称 支持['styles','settings', 'webSettings', 'fontTable']
		:param part_path: 模板路径|模板Part
		"""
		self._custom_templates_config[1][part_name] = part_path
	
	def _init_body(self):
		self.lastSection.sectPr = SectionPr(
			pageSize=PageSize(width=Measure.CM(21), height=Measure.CM(29.7)),
			pageMargin=PageMargin(
				left=1800,
				top=1440,
				right=1800,
				bottom=1440,
				header=851,
				footer=992,
				gutter=0
			),
			cols=Cols(space=Measure.TWIPS(425)),
			docGrid=DocGrid(type=DocGrid.enum.LINES, linePitch=312, charSpace=312),
		)
	
	def to_xml(self):
		with open(os.path.join(template_base, 'word/document.xml'), 'r') as f:
			doc_str = f.read()
		
		self._children[-1].isLastSection = True
		doc_str = doc_str.replace('<!--main_content-->', super().to_xml())
		self._children[-1].isLastSection = False
		
		return doc_str
	
	def save(self, docx_path, debug=False):
		"""保存文档\n
		:param docx_path: 保存路径
		:param debug: 是否开启调试模式，开启后会自动解压文件到debug文件夹下
		"""
		docx_writer = self._driver(docx_path)
		
		part_document = opc.NodePart(self, opc.OPC_BUILTIN_MAP['document'])
		part_app = opc.FilePart(os.path.join(template_base, 'docProps/app.xml'), opc.OPC_BUILTIN_MAP['app'])
		part_core = opc.FilePart(os.path.join(template_base, 'docProps/core.xml'), opc.OPC_BUILTIN_MAP['core'])
		
		# app part
		self.add_app_part(part_app)
		self.add_app_part(part_core)
		self.add_app_part(part_document)
		
		# word part
		template_path = self._custom_templates_config[0]
		for part_name, part_value in self._custom_templates_config[1].items():
			if part_value is None:
				file_path = os.path.join(template_path, f'./{part_name}.xml')
				part = opc.FilePart(file_path, opc.OPC_BUILTIN_MAP[part_name])
			elif isinstance(part_value, opc.NodePart):
				part = part_value
			elif isinstance(part_value, str):
				part = opc.FilePart(part_value, opc.OPC_BUILTIN_MAP[part_name])
			else:
				raise ValueError(f"Invalid part value: {part_value}")
			self.add_word_part(part)
		
		# part_theme = opc.FilePart(os.path.join(template_base, 'word/theme/theme1.xml'), OPC_BUILTIN_MAP['theme'])
		
		# self.add_word_part(part_theme)
		
		# 从rels写入docx
		self._app_rels.write_to(docx_writer)
		self._word_rels.write_to(docx_writer)
		# 从contentTypes写入docx
		self._content_types.write_to(docx_writer)
		
		docx_writer.close()
		if debug:
			with zipfile.ZipFile(docx_path, mode='r') as f:
				f.extractall('./debug/' + os.path.basename(docx_path).split('.')[0])
				content = f.read('word/document.xml')
				with open('./debug/' + os.path.basename(docx_path).split('.')[0] + '.xml', 'wb') as f:
					f.write(content)
	
	def add_word_part(self, part):
		"""注册主文档关系"""
		self._content_types.add_part(part)
		self._word_rels.add_part(part)
		return part
	
	def add_app_part(self, part):
		"""注册应用程序关系"""
		self._content_types.add_part(part)
		self._app_rels.add_part(part)
		return part

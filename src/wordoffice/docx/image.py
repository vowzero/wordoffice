from PIL import Image as PILImage

from wordoffice.core.opc import RelationshipTypeEnum, FilePart
from wordoffice.common import Measure
from wordoffice.docx.document import Docx
from wordoffice.ooxml.dml import *

image_count_id = 1


class ImageJpegPart(FilePart):
	opc_info = ("media/image{}.jpeg", ('jpeg', 'image/jpeg'), RelationshipTypeEnum.MEDIA_IMAGE)
	
	def __init__(self, filename, word: Docx = None):
		super().__init__(filename, self.opc_info)
		
		image = PILImage.open(filename)
		self.size = image.size
		if 'dpi' in image.info:
			self.dpi = image.info['dpi']
		else:
			self.dpi = None
		image.close()
		
		if word is not None:
			word.add_word_part(self)
	
	def get_origin_inches_size(self):
		if self.dpi is None:
			return self.size[0] / 72, self.size[1] / 72
		return self.size[0] / self.dpi[0], self.size[1] / self.dpi[1]


class ImageInline(Drawing):
	def __init__(self, image_part, size=None):
		super().__init__()
		self._image_part: ImageJpegPart = image_part
		self._picture = Picture()
		self._init_inline()
		
		self._size = None
		
		self.size = size
		self.non_visual = {}
	
	@property
	def non_visual(self):
		return self._non_visual
	
	@non_visual.setter
	def non_visual(self, non_visual: dict):
		global image_count_id
		if not isinstance(non_visual, dict):
			raise ValueError(f"non_visual must be a dict")
		
		new_non_visual = {'id': image_count_id, 'name': '图片 ' + str(image_count_id), 'descr': '图片描述 ' + str(image_count_id)}
		new_non_visual.update(non_visual)
		
		image_count_id += 1
		
		self._non_visual = new_non_visual
		self.inline.docPr = new_non_visual
		self._picture.nvPicPr.cNvPr = new_non_visual
	
	@property
	def size(self):
		return self._size
	
	@size.setter
	def size(self, size):
		if size is None:
			size = self._image_part.get_origin_inches_size()
			size = (Measure.INCH(size[0]), Measure.INCH(size[1]))
		new_size_dict = {'cx': size[0], 'cy': size[1]}
		self.inline.extent = new_size_dict
		self._picture.spPr.xfrm.ext = new_size_dict
		self._size = size
	
	@property
	def xfrm(self) -> ShapePrTransform:
		return self._picture.spPr.xfrm
	
	@xfrm.setter
	def xfrm(self, xfrm):
		self._picture.spPr.xfrm = xfrm
	
	def _init_inline(self):
		picture = self._picture
		blipFill = BlipFill()
		spPr = SpPr()
		self.inline = DrawingInline()
		
		# self.inline.extent = None  # 固定extent顺序位置
		# self.inline.effectExtent = EffectExtent()
		# 替换可选文字描述
		self.inline.docPr = None
		# self.inline.cNvGraphicFramePr # 盲人相关
		self.inline.graphic = Graphic()
		self.inline.graphic.graphicData = GraphicData(uri="http://schemas.openxmlformats.org/drawingml/2006/picture")
		self.inline.graphic.graphicData.append(picture)
		
		picture.nvPicPr = NvPicPr()
		picture.nvPicPr.cNvPr = None
		picture.nvPicPr.cNvPicPr = CNvPicPr()
		picture.blipFill = blipFill
		picture.spPr = spPr
		
		blipFill.blip = Blip(embed=self._image_part.part_rid)
		# office word在Blip中有个魔法子元素extLst，功能未知
		'''
		    <a:extLst>
			<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
			    <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
			</a:ext>
		    </a:extLst>
		'''
		blipFill.stretch = Stretch(fillRect=FillRect())
		
		spPr.xfrm = ShapePrTransform()
		spPr.xfrm.off = ShapePrTransformOffset(x=0, y=0)
		spPr.xfrm.ext = None
		
		spPr.prstGeom = PrstGeom(prst=PrstGeom.enum.RECT)
		spPr.prstGeom.avLst = AvLst()


class ImageAnchor(Drawing):
	tag = 'w:drawing'
	POS_H = ST_RelFromH
	POS_V = ST_RelFromV
	
	def __init__(self, image_part, size=None, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._image_part: ImageJpegPart = image_part
		self._picture = Picture()
		self._init_anchor()
		
		self._size = None
		
		self.size = size
		self.non_visual = {}
	
	@property
	def non_visual(self):
		return self._non_visual
	
	@non_visual.setter
	def non_visual(self, non_visual: dict):
		global image_count_id
		if not isinstance(non_visual, dict):
			raise ValueError(f"non_visual must be a dict")
		
		new_non_visual = {'id': image_count_id, 'name': '图片 ' + str(image_count_id), 'descr': '图片描述 ' + str(image_count_id)}
		new_non_visual.update(non_visual)
		
		image_count_id += 1
		
		self._non_visual = new_non_visual
		self.anchor.docPr = new_non_visual
		self._picture.nvPicPr.cNvPr = new_non_visual
	
	@property
	def size(self):
		return self._size
	
	@size.setter
	def size(self, size):
		if size is None:
			size = self._image_part.get_origin_inches_size()
			size = (Measure.INCH(size[0]), Measure.INCH(size[1]))
		new_size_dict = {'cx': size[0], 'cy': size[1]}
		self.anchor.extent = new_size_dict
		self._picture.spPr.xfrm.ext = new_size_dict
		self._size = size
	
	@property
	def xfrm(self) -> ShapePrTransform:
		return self._picture.spPr.xfrm
	
	@xfrm.setter
	def xfrm(self, xfrm):
		self._picture.spPr.xfrm = xfrm
	
	@property
	def overlap(self):
		return self.anchor.allowOverlap
	
	@overlap.setter
	def overlap(self, overlap):
		self.anchor.allowOverlap = overlap
	
	@property
	def behindDoc(self):
		return self.anchor.behindDoc
	
	@behindDoc.setter
	def behindDoc(self, behindDoc):
		self.anchor.behindDoc = behindDoc
	
	@property
	def position(self):
		return self.anchor.positionH, self.anchor.positionV
	
	@position.setter
	def position(self, position):
		self.anchor.positionH, self.anchor.positionV = position
	
	def _init_anchor(self):
		picture = self._picture
		blipFill = BlipFill()
		spPr = SpPr()
		self.anchor = DrawingAnchor(relativeHeight=2, behindDoc=False, layoutInCell=True, locked=False, allowOverlap=True, simplePosSwitch=False)
		
		self.anchor.simplePos = {'x': 0, 'y': 0}
		self.anchor.positionH = {'relativeFrom': PositionH.enum.PAGE, 'posOffset': Measure.PT(0)}
		self.anchor.positionV = {'relativeFrom': PositionV.enum.PAGE, 'posOffset': Measure.PT(0)}
		
		self.anchor.extent = None  # 固定extent顺序位置
		self.anchor.wrapNone = True
		
		# # self.anchor.effectExtent = EffectExtent()
		# # 替换可选文字描述
		self.anchor.docPr = None
		# # self.inline.cNvGraphicFramePr # 盲人相关
		self.anchor.graphic = Graphic()
		self.anchor.graphic.graphicData = GraphicData(uri="http://schemas.openxmlformats.org/drawingml/2006/picture")
		self.anchor.graphic.graphicData.append(picture)
		#
		picture.nvPicPr = NvPicPr()
		picture.nvPicPr.cNvPr = None
		picture.nvPicPr.cNvPicPr = CNvPicPr()
		picture.blipFill = blipFill
		picture.spPr = spPr
		#
		blipFill.blip = Blip(embed=self._image_part.part_rid)
		# # office word在Blip中有个魔法子元素extLst，功能未知
		# '''
		#     <a:extLst>
		#         <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
		#             <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
		#         </a:ext>
		#     </a:extLst>
		# '''
		blipFill.stretch = Stretch(fillRect=FillRect())
		#
		spPr.xfrm = ShapePrTransform()
		spPr.xfrm.off = ShapePrTransformOffset(x=0, y=0)
		spPr.xfrm.ext = None
		
		spPr.prstGeom = PrstGeom(prst=PrstGeom.enum.RECT)
		spPr.prstGeom.avLst = AvLst()


def Image(image_part, inline=True):
	if inline:
		return ImageInline(image_part)
	else:
		return ImageAnchor(image_part)

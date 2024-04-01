from typing import Iterable

from wordoffice.core.field import SubField, AttrField
from wordoffice.core.node import OoxmlNode
from wordoffice.core.type import *
from .enums import *
from .shared import BlockLevelElements
from .shared_node import Shading, BorderAttrFields
from .utils import SimpleValNode, create_switch_node


class TableStyle(SimpleValNode):
	tag = "w:tblStyle"
	val = AttrField('w:val', 'string')


class TablePositionPr(OoxmlNode):
	tag = "w:tblpPr"
	enum_vAnchor = ST_VAnchor
	enum_hAnchor = ST_HAnchor
	enum_xAlign = ST_XAlign
	enum_yAlign = ST_YAlign
	leftFromText = AttrField('w:leftFromText', OoxmlTwipsMeasure)
	rightFromText = AttrField('w:rightFromText', OoxmlTwipsMeasure)
	topFromText = AttrField('w:topFromText', OoxmlTwipsMeasure)
	bottomFromText = AttrField('w:bottomFromText', OoxmlTwipsMeasure)
	vertAnchor = AttrField('w:vertAnchor', ('enum', {'enums': enum_vAnchor}))
	horzAnchor = AttrField('w:horzAnchor', ('enum', {'enums': enum_hAnchor}))
	tblpXSpec = AttrField('w:tblpXSpec', ('enum', {'enums': enum_xAlign}))
	tblpX = AttrField('w:tblpX', OoxmlSignedTwipsMeasure)
	tblpYSpec = AttrField('w:tblpYSpec', ('enum', {'enums': enum_yAlign}))
	tblpY = AttrField('w:tblpY', OoxmlSignedTwipsMeasure)


class TableOverlap(SimpleValNode):
	tag = "w:tblOverlap"
	enum = ST_TblOverlap
	val = AttrField('w:val', ('enum', {'enums': enum}))


BidiVisual = create_switch_node('BidiVisual', 'w:bidiVisual', False)


class TableStyleRowBandSize(SimpleValNode):
	tag = "w:tblStyleRowBandSize"
	val = AttrField('w:val', 'integer')


class TableStyleColBandSize(SimpleValNode):
	tag = "w:tblStyleColBandSize"
	val = AttrField('w:val', 'integer')


class _CT_TblWidth(OoxmlNode):
	enum = ST_TblWidth
	type = AttrField('w:type', ('enum', {'enums': enum}))
	w = AttrField('w:w', OoxmlMeasurementOrPercent)


class TableWidth(_CT_TblWidth):
	tag = "w:tblW"


class TableJc(SimpleValNode):
	tag = "w:jc"
	enum = ST_JcTable
	val = AttrField('w:val', ('enum', {'enums': enum}))


class TableCellSpacing(_CT_TblWidth):
	tag = "w:tblCellSpacing"


class TableInd(_CT_TblWidth):
	tag = "w:tblInd"


class TableBorderTop(BorderAttrFields):
	tag = "w:top"


class TableBorderBottom(BorderAttrFields):
	tag = "w:bottom"


class TableBorderLeft(BorderAttrFields):
	tag = "w:left"


class TableBorderRight(BorderAttrFields):
	tag = "w:right"


class TableBorderStart(BorderAttrFields):
	tag = "w:start"


class TableBorderEnd(BorderAttrFields):
	tag = "w:end"


class TableBorderInsideH(BorderAttrFields):
	tag = "w:insideH"


class TableBorderInsideV(BorderAttrFields):
	tag = "w:insideV"


class TableBorders(OoxmlNode):
	tag = "w:tblBorders"
	top = SubField(TableBorderTop)
	start = SubField(TableBorderStart)
	left = SubField(TableBorderLeft)
	bottom = SubField(TableBorderBottom)
	end = SubField(TableBorderEnd)
	right = SubField(TableBorderRight)
	insideH = SubField(TableBorderInsideH)
	insideV = SubField(TableBorderInsideV)


class TableLayoutType(SimpleValNode):
	tag = "w:tblLayout"
	enum = ST_TblLayoutType
	type = AttrField('w:type', ('enum', {'enums': enum}))


class TableCellMarginTop(_CT_TblWidth):
	tag = "w:top"


class TableCellMarginStart(_CT_TblWidth):
	tag = "w:start"


class TableCellMarginLeft(_CT_TblWidth):
	tag = "w:left"


class TableCellMarginBottom(_CT_TblWidth):
	tag = "w:bottom"


class TableCellMarginEnd(_CT_TblWidth):
	tag = "w:end"


class TableCellMarginRight(_CT_TblWidth):
	tag = "w:right"


class TableCellMar(OoxmlNode):
	tag = "w:tblCellMar"
	top = SubField(TableCellMarginTop)
	start = SubField(TableCellMarginStart)
	left = SubField(TableCellMarginLeft)
	bottom = SubField(TableCellMarginBottom)
	end = SubField(TableCellMarginEnd)
	right = SubField(TableCellMarginRight)


class TableLook(OoxmlNode):
	tag = "w:tblLook"
	
	firstRow = AttrField('w:firstRow', 'ooxml_boolean')
	lastRow = AttrField('w:lastRow', 'ooxml_boolean')
	firstColumn = AttrField('w:firstColumn', 'ooxml_boolean')
	lastColumn = AttrField('w:lastColumn', 'ooxml_boolean')
	noHBand = AttrField('w:noHBand', 'ooxml_boolean')
	noVBand = AttrField('w:noVBand', 'ooxml_boolean')
	val = AttrField('w:val', 'ooxml_short_hex_number')


class TableCaption(SimpleValNode):
	tag = "w:tblCaption"
	val = AttrField('w:val', 'string')


class TableDescription(SimpleValNode):
	tag = "w:tblDescription"
	val = AttrField('w:val', 'string')


class TablePr(OoxmlNode):
	tag = "w:tblPr"
	tblStyle = SubField(TableStyle)
	tblpPr = SubField(TablePositionPr)
	tblOverlap = SubField(TableOverlap)
	bidiVisual = SubField(BidiVisual)
	tblStyleRowBandSize = SubField(TableStyleRowBandSize)
	tblStyleColBandSize = SubField(TableStyleColBandSize)
	tblW = SubField(TableWidth)
	jc = SubField(TableJc)
	tblCellSpacing = SubField(TableCellSpacing)
	tblInd = SubField(TableInd)
	tblBorders = SubField(TableBorders)
	shd = SubField(Shading)
	tblLayout = SubField(TableLayoutType)
	tblCellMar = SubField(TableCellMar)
	tblLook = SubField(TableLook)
	tblCaption = SubField(TableCaption)
	tblDescription = SubField(TableDescription)


class GridCol(OoxmlNode):
	tag = "w:gridCol"
	w = AttrField('w:w', OoxmlTwipsMeasure)


class TableGrid(OoxmlNode):
	tag = "w:tblGrid"
	_children_type = (GridCol,)
	
	def add_grid_col(self, gridCol):
		self.append(gridCol)


class TablePrEx(OoxmlNode):
	tag = "w:tblPrEx"
	tblW = SubField(TableWidth)
	jc = SubField(TableJc)
	tblCellSpacing = SubField(TableCellSpacing)
	tblInd = SubField(TableInd)
	tblBorders = SubField(TableBorders)
	shd = SubField(Shading)
	tblLayout = SubField(TableLayoutType)
	tblCellMar = SubField(TableCellMar)
	tblLook = SubField(TableLook)


class TableRowCnfStyle(OoxmlNode):
	tag = "w:cnfStyle"
	val = AttrField('w:val', 'string')  # TODO: ST_Cnf
	firstRow = AttrField('w:firstRow', 'ooxml_boolean')
	lastRow = AttrField('w:lastRow', 'ooxml_boolean')
	firstColumn = AttrField('w:firstColumn', 'ooxml_boolean')
	lastColumn = AttrField('w:lastColumn', 'ooxml_boolean')
	oddVBand = AttrField('w:oddVBand', 'ooxml_boolean')
	evenVBand = AttrField('w:evenVBand', 'ooxml_boolean')
	oddHBand = AttrField('w:oddHBand', 'ooxml_boolean')
	evenHBand = AttrField('w:evenHBand', 'ooxml_boolean')
	firstRowFirstColumn = AttrField('w:firstRowFirstColumn', 'ooxml_boolean')
	firstRowLastColumn = AttrField('w:firstRowLastColumn', 'ooxml_boolean')
	lastRowFirstColumn = AttrField('w:lastRowFirstColumn', 'ooxml_boolean')
	lastRowLastColumn = AttrField('w:lastRowLastColumn', 'ooxml_boolean')


class TableRowDivId(SimpleValNode):
	tag = "w:divId"
	val = AttrField('w:val', 'integer')


class TableRowGridBefore(SimpleValNode):
	tag = "w:gridBefore"
	val = AttrField('w:val', 'integer')


class TableRowGridAfter(SimpleValNode):
	tag = "w:gridAfter"
	val = AttrField('w:val', 'integer')


class TableRowWidthBefore(_CT_TblWidth):
	tag = "w:wBefore"


class TableRowWidthAfter(_CT_TblWidth):
	tag = "w:wAfter"


TableRowCantSplit = create_switch_node('CantSplit', 'w:cantSplit', False)


class TableRowHeight(OoxmlNode):
	tag = "w:trHeight"
	enum = ST_HeightRule
	val = AttrField('w:val', OoxmlTwipsMeasure)
	hRule = AttrField('w:hRule', ('enum', {'enums': enum}))


TableRowTblHeader = create_switch_node('TblHeader', 'w:tblHeader', False)


class TableRowTblCellSpacing(_CT_TblWidth):
	tag = "w:tblCellSpacing"


TableRowJc = TableJc

TableRowHidden = create_switch_node('Hidden', 'w:hidden', False)


class TableRowPr(OoxmlNode):
	tag = "w:trPr"
	cnfStyle = SubField(TableRowCnfStyle)
	divId = SubField(TableRowDivId)
	gridBefore = SubField(TableRowGridBefore)
	gridAfter = SubField(TableRowGridAfter)
	wBefore = SubField(TableRowWidthBefore)
	wAfter = SubField(TableRowWidthAfter)
	cantSplit = SubField(TableRowCantSplit)
	trHeight = SubField(TableRowHeight)
	tblHeader = SubField(TableRowTblHeader)
	tblCellSpacing = SubField(TableRowTblCellSpacing)
	jc = SubField(TableRowJc)
	hidden = SubField(TableRowHidden)


TableCellCnfStyle = TableRowCnfStyle


class TableCellWidth(_CT_TblWidth):
	tag = "w:tcW"


class TableCellGridSpan(SimpleValNode):
	tag = "w:gridSpan"
	val = AttrField('w:val', 'integer')


class _TableCellMerge(SimpleValNode):
	enum = ST_Merge
	val = AttrField('w:val', ('enum', {'enums': enum}))


class TableCellHMerge(_TableCellMerge):
	tag = "w:hMerge"


class TableCellVMerge(_TableCellMerge):
	tag = "w:vMerge"


class TableBorderTl2Br(BorderAttrFields):
	tag = "w:tl2br"


class TableBorderTr2Bl(BorderAttrFields):
	tag = "w:tr2bl"


class TableCellBorders(OoxmlNode):
	tag = "w:tcBorders"
	top = SubField(TableBorderTop)
	start = SubField(TableBorderStart)
	left = SubField(TableBorderLeft)
	bottom = SubField(TableBorderBottom)
	end = SubField(TableBorderEnd)
	right = SubField(TableBorderRight)
	insideH = SubField(TableBorderInsideH)
	insideV = SubField(TableBorderInsideV)
	tl2br = SubField(TableBorderTl2Br)
	tr2bl = SubField(TableBorderTr2Bl)


TableCellNoWrap = create_switch_node('NoWrap', 'w:noWrap', False)


class TableCellTextDirection(SimpleValNode):
	tag = "w:textDirection"
	enum = ST_TextDirection
	val = AttrField('w:val', ('enum', {'enums': enum}))


TableCellFitText = create_switch_node('FitText', 'w:tcFitText', False)


class TableCellVerticalJc(SimpleValNode):
	tag = "w:vAlign"
	enum = ST_VerticalJc
	val = AttrField('w:val', ('enum', {'enums': enum}))


TableCellHideMark = create_switch_node('HideMark', 'w:hideMark', False)


class TableCellHeader(SimpleValNode):
	tag = "w:header"
	val = AttrField('w:val', 'string')


class TableCellHeaders(OoxmlNode):
	tag = "w:headers"
	_children_type = (TableCellHeader,)


class TableCellPr(OoxmlNode):
	tag = "w:tcPr"
	cnfStyle = SubField(TableRowCnfStyle)
	tcW = SubField(TableCellWidth)
	gridSpan = SubField(TableCellGridSpan)
	hMerge = SubField(TableCellHMerge)
	vMerge = SubField(TableCellVMerge)
	tcBorders = SubField(TableCellBorders)
	shd = SubField(Shading)
	noWrap = SubField(TableCellNoWrap)
	tcMar = SubField(TableCellMar)
	textDirection = SubField(TableCellTextDirection)
	tcFitText = SubField(TableCellFitText)
	vAlign = SubField(TableCellVerticalJc)
	hideMark = SubField(TableCellHideMark)
	headers = SubField(TableCellHeaders)


class TableCell(BlockLevelElements):
	tag = "w:tc"
	id = AttrField('w:id', 'string')
	tcPr: TableCellPr = SubField(TableCellPr)
	
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self.hMerge = kwargs.get('hMerge', False)
		self.vMerge = kwargs.get('vMerge', False)
	
	def to_xml(self):
		if self.hMerge is True:
			return ''
		else:
			return super().to_xml()


class TableRow(OoxmlNode):
	tag = "w:tr"
	tblPrEx = SubField(TablePrEx)
	trPr = SubField(TableRowPr)
	_children_type = (TableCell,)
	
	# TODO: Table Cell Levels
	@property
	def main_cols(self) -> list[TableCell]:
		# noinspection PyTypeChecker
		for index, cell in enumerate(self._children):
			if not cell.hMerge and not cell.vMerge:
				yield index, cell
	
	@property
	def cols(self) -> list[TableCell]:
		# noinspection PyTypeChecker
		return self._children
	
	@cols.setter
	def cols(self, _cols):
		if not isinstance(_cols, Iterable):
			assert False, "cols必须是可迭代对象(列表元组)"
		else:
			for col in _cols:
				if not isinstance(col, TableCell):
					assert False, "col元素必须是TableCell对象"
			self._children = list(_cols)
	
	def __getitem__(self, indices) -> TableCell | list[TableCell]:
		if isinstance(indices, (int, slice)):
			return self.cols[indices]
		else:
			assert False, "indices必须是切片或col索引"


class Table(OoxmlNode):
	tag = "w:tbl"
	tblPr = SubField(TablePr)
	tblGrid = SubField(TableGrid)
	_children_type = (TableRow,)
	
	def __init__(self, size=(1, 1), *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._size = size
	
	@property
	def size(self):
		return self._size
	
	@property
	def rows(self) -> list[TableRow]:
		# noinspection PyTypeChecker
		return self._children
	
	@rows.setter
	def rows(self, _rows):
		if not isinstance(_rows, Iterable):
			assert False, "rows必须是可迭代对象(列表元组)"
		else:
			for row in _rows:
				if not isinstance(row, TableRow):
					assert False, "row元素必须是TableRow对象"
			self._children = list(_rows)
	
	def __getitem__(self, indices) -> TableRow | list[TableRow]:
		if isinstance(indices, (int, slice)):
			return self.rows[indices]
		else:
			assert False, "indices必须是切片或row索引"
	
	def before_to_xml(self):
		data = super().before_to_xml()
		rows: list[TableRow] = list(data[3])
		rows_length = len(rows)
		
		for col in rows[0]:
			if col.tcPr.vMerge is True:
				assert False, "表格第一行的单元格不能向上合并"
		
		for i in range(rows_length):
			cols: list[TableCell] = rows[i].cols
			cols_length = len(cols)
			if cols_length >= 1 and cols[0].hMerge is True:
				assert False, "表格行的第一列单元格不能向左合并"
			
			for j in range(cols_length):
				col_span = 1
				cur_col = cols[j]
				
				# 列内向上合并单元格
				if cur_col.vMerge is False:
					with_restart = False
					for k in range(i + 1, rows_length):
						cur_next_col = rows[k][j]
						if cur_next_col.vMerge is True:
							with_restart = True
							cur_next_col.tcPr.vMerge = TableCellVMerge(val=TableCellVMerge.enum.CONTINUE)
						else:
							break
					if with_restart:
						cur_col.tcPr.vMerge = TableCellVMerge(val=TableCellVMerge.enum.RESTART)
				
				# 行内向左合并单元格
				for k in range(j + 1, cols_length):
					if cols[k].hMerge is True:
						col_span += 1
					else:
						break
				
				if col_span > 1:
					cur_col.tcPr.gridSpan = TableCellGridSpan(val=col_span)
				else:
					cur_col.tcPr.gridSpan = None
		
		data[3] = rows
		return data
	
	def merge_cell(self, start: tuple, end: tuple) -> TableCell:
		start_row, start_col = start
		end_row, end_col = end
		if start_row > end_row or start_col > end_col:
			assert False, "起始坐标不能大于结束坐标"
		start_cell = self[start_row][start_col]
		if start_cell.hMerge or start_cell.vMerge:
			assert False, "起始单元格已经合并"
		
		for row in self[start_row:end_row + 1]:
			for col in row[start_col + 1:end_col + 1]:
				col.hMerge = True
		for row in self[start_row + 1:end_row + 1]:
			row[start_col].vMerge = True
		
		return start_cell

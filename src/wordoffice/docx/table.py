from wordoffice.ooxml import table as ooxml_table


class Table(ooxml_table.Table):
	tag = "w:tbl"
	
	def __init__(self, size, *args, **kwargs):
		"""
		创建表格
		:param size: 表格大小，元组(行数,列数)
		"""
		super().__init__(size, *args, **kwargs)
		self._init_table_pr(kwargs.get('tblPr'))
		self._init_table_size()
	
	def _init_table_size(self):
		row_size, col_size = self._size
		self.tblGrid = ooxml_table.TableGrid()
		for i in range(col_size):
			self.tblGrid.append(ooxml_table.GridCol(w=2074))
		for i in range(row_size):
			tr = ooxml_table.TableRow()
			for j in range(col_size):
				cur_cell = ooxml_table.TableCell()
				cur_cell.tcPr = ooxml_table.TableCellPr()
				cur_cell.tcPr.tcW = ooxml_table.TableCellWidth(w=2074, type='dxa')
				
				tr.append(cur_cell)
			self.append(tr)
	
	def _init_table_pr(self, tblPr):
		"""初始化表格属性"""
		if tblPr is None:
			self.tblPr = ooxml_table.TablePr()
		# 初始化默认表格风格
		# if self.tblPr.tblStyle is None:
		#         self.tblPr.tblStyle = ooxml_table.TableStyle(val="a3")
		# 初始化默认表格宽度
		if self.tblPr.tblW is None:
			self.tblPr.tblW = ooxml_table.TableWidth(type=ooxml_table.TableWidth.enum.AUTO, w=0)
		# 初始化默认表格条件样式
		if self.tblPr.tblLook is None:
			self.tblPr.tblLook = ooxml_table.TableLook(val="#04A0", firstRow=True, lastRow=False, firstColumn=True, lastColumn=False, noHBand=False, noVBand=True)

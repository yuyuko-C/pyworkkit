import openpyxl.worksheet.merge as o_merge
from openpyxl.worksheet.cell_range import CellRange


from ..cell import Cell


class MergedCellRange(o_merge.MergedCellRange):
    def __init__(self, worksheet, coord):
        from .worksheet import Worksheet
        super().__init__(worksheet, coord)
        self.ws:Worksheet = worksheet
        self.start_cell:Cell=self.start_cell

    def unmerge(self):
        # 这里的行和列的起始值（索引），和Excel的一样，从1开始，并不是从0开始（注意）
        r1, r2, c1, c2 = self.min_row, self.max_row, self.min_col, self.max_col
        self.ws.unmerge_cells(start_row=r1, end_row=r2, start_column=c1, end_column=c2)
        
    def fill(self,value):
        r1, r2, c1, c2 = self.min_row, self.max_row, self.min_col, self.max_col
        # 数据填充
        for r in range(r1, r2+1):  # 遍历行
            if c2 - c1 > 0:  # 多列
                # 遍历列
                for c in range(c1, c2+1):
                    self.ws.cell(r, c).value = value
            else:  # 单列
                self.ws.cell(r, c1).value = value
                self.start_cell
                
    @property
    def shape(self):
        return self.max_col-self.min_col+1,self.max_row-self.min_row+1
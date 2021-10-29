import openpyxl.cell as o_cell
from copy import copy


class Cell(o_cell.Cell):
    """
    继承了原 Cell 的类，为了实现扩展方法
    """

    def copy_style(self, cell):
        self.copy_font(cell)
        self.copy_border(cell)
        self.copy_fill(cell)
        self.copy_alignment(cell)
        self.copy_number_format(cell)

    def copy_font(self, cell):
        self.font = copy(cell.font)

    def copy_border(self, cell):
        self.border = copy(cell.border)

    def copy_fill(self, cell):
        self.fill = copy(cell.fill)

    def copy_alignment(self, cell):
        self.alignment = copy(cell.alignment)

    def copy_number_format(self, cell):
        self.number_format = copy(cell.number_format)

    def set_number_format(self, format: str):
        self.number_format = format

    def set_font_bold(self, enable):
        font = copy(self.font)
        font.bold = enable
        self.font = font


class MergedCell(o_cell.MergedCell):
    def __init__(self, worksheet, row, column) -> None:
        super().__init__(worksheet, row=row, column=column)

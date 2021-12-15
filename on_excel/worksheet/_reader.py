
import openpyxl.worksheet._reader as o_reader
from openpyxl.xml.constants import (
    SHEET_MAIN_NS,
    EXT_TYPES,
)


from ..cell import Cell
from .worksheet import Worksheet

VALUE_TAG = '{%s}v' % SHEET_MAIN_NS

class WorkSheetParser(o_reader.WorkSheetParser):
    """
    继承了原 WorkSheetParser 的类，为了使用自定义的表格读取规则
    """

    # 重写此方法，用于使用自定义的表格读取规则
    def parse_row(self, row):
        attrs = dict(row.attrib)

        if "r" in attrs:
            try:
                self.row_counter = int(attrs['r'])
            except ValueError:
                val = float(attrs['r'])
                if val.is_integer():
                    self.row_counter = int(val)
                else:
                    raise ValueError(f"{attrs['r']} is not a valid row number")
        else:
            self.row_counter += 1
        self.col_counter = 0

        keys = {k for k in attrs if not k.startswith('{')}
        if keys - {'r', 'spans'}:
            # don't create dimension objects unless they have relevant information
            self.row_dimensions[str(self.row_counter)] = attrs

        # 自定义：不读取值为空的元素
        # 可以改变maxrow与maxcolumn以及to_Dataframe方法
        # 不可改变保存后的结果
        cells = [self.parse_cell(el)
                 for el in row if el.findtext(VALUE_TAG, None)]

        return self.row_counter, cells


class WorksheetReader(o_reader.WorksheetReader):
    """
    继承了原 WorksheetReader 的类，为了使用自定义的 Cell
    """

    def __init__(self, ws, xml_source, shared_strings, data_only):
        self.ws:Worksheet = ws
        self.parser = WorkSheetParser(xml_source, shared_strings,
                data_only, ws.parent.epoch, ws.parent._date_formats,
                ws.parent._timedelta_formats)
        self.tables = []


    # 重写此方法，用于使用自定义的 Cell
    def bind_cells(self):
        for idx, row in self.parser.parse():
            for cell in row:
                style = self.ws.parent._cell_styles[cell['style_id']]
                c = Cell(
                    self.ws, row=cell['row'], column=cell['column'], style_array=style)
                c._value = cell['value']
                c.data_type = cell['data_type']
                self.ws._cells[(cell['row'], cell['column'])] = c
        self.ws.formula_attributes = self.parser.array_formulae
        if self.ws._cells:
            self.ws._current_row = self.ws.max_row  # use cells not row dimensions


    def bind_merged_cells(self):
        from openpyxl.worksheet.cell_range import MultiCellRange
        from .merge import MergedCellRange
        
        if not self.parser.merged_cells:
            return

        ranges = []
        for cr in self.parser.merged_cells.mergeCell:
            mcr = MergedCellRange(self.ws, cr.ref)
            self.ws._clean_merge_range(mcr)
            ranges.append(mcr)
        self.ws.merged_cells = MultiCellRange(ranges)


import openpyxl.worksheet.worksheet as o_sheet
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter, column_index_from_string
from inspect import isgenerator
import pandas as pd


from on_excel.cell import Cell, MergedCell


class Worksheet(o_sheet.Worksheet):
    """
    继承了原 Worksheet 的类，为了实现扩展方法
    """

    def __init__(self, parent, title=None) -> None:
        super(Worksheet, self).__init__(parent, title)
        if self.is_huge():
            for col in self.columns:
                print(col)
            pass

    @property
    def visiable(self):
        return self.sheet_state == 'visible'

    @visiable.setter
    def visiable(self, enable: bool):
        if enable:
            self.sheet_state = 'visible'
        else:
            self.sheet_state = 'hidden'

    def is_huge(self):
        if self.max_column > 50 or self.max_row > 100000:
            return True
        else:
            return False

    def set_freeze_panes(self, coordinate: str):
        self.freeze_panes = coordinate

    def delete_blank_rows_and_columns(self):
        self.max_column
        col_index = 1
        for col in self.iter_cols(values_only=True, max_col=100):
            cells = [str(cell).strip() for cell in col if cell != None]
            if all([not cell for cell in cells]):
                self.delete_cols(col_index, self.max_column-col_index+1)
                break
            col_index += 1
        row_index = 1
        for row in self.iter_rows(values_only=True):
            cells = [str(cell).strip() for cell in row if cell != None]
            if all([not cell for cell in cells]):
                self.delete_rows(row_index, self.max_row-row_index+1)
                break
            row_index += 1
        print(self.title, '已清空多余的行列')
        return self

    def clear(self, style=True):
        """Clear all content in sheet.

        Args:
            style (bool, optional): is or not clear style . Defaults to True.
        """
        row = 1
        maxrow = self.max_row
        if style:
            self.delete_rows(1, maxrow)
        else:
            self.unmerge_cells(True)
            maxcol = self.max_column
            while row <= maxrow:
                col = 1
                while col <= maxcol:
                    self.cell(row, col).value = ''
                    col += 1
                row += 1

    def match_value_rows(self, value, on_col: str, agg='equal'):
        """Get rows fitted with condition.

        Args:
            value ([type]): [description]
            on_col (str): [description]
            agg (str, optional): 'equal'  'start'  'end'  'in'  'not_in'. Defaults to 'equal'.

        Returns:
            [type]: [description]
        """
        if isinstance(value, str):
            if agg == 'equal':
                return [cell.row for cell in self[on_col] if cell.value != None and cell.value == value]
            elif agg == 'start':
                return [cell.row for cell in self[on_col] if cell.value != None and cell.value.startswith(value)]
            elif agg == 'in':
                return [cell.row for cell in self[on_col] if cell.value != None and value in cell.value]
            elif agg == 'end':
                return [cell.row for cell in self[on_col] if cell.value != None and cell.value.endswith(value)]
            elif agg == 'not_in':
                return [cell.row for cell in self[on_col] if cell.value != None and value not in cell.value]
        elif isinstance(value, list):
            ret = []
            for x in value:
                ret.extend(self.match_value_rows(x, on_col, agg))
            return ret

    def append_df(self, df: pd.DataFrame):
        df.reset_index(drop=False, inplace=True)
        df = df.drop(df.columns.tolist()[0], axis=1)
        self.append(df.columns.values.tolist())
        [self.append(row) for row in df.values.tolist()]

    def unmerge_cell(self, fill):

        m_list = list(self.merged_cells)

        # 拆分合并的单元格 并填充内容
        for m_area in m_list:
            # 这里的行和列的起始值（索引），和Excel的一样，从1开始，并不是从0开始（注意）
            r1, r2, c1, c2 = m_area.min_row, m_area.max_row, m_area.min_col, m_area.max_col

            self.unmerge_cells(
                start_row=r1, end_row=r2, start_column=c1, end_column=c2)
            # print('区域:', m_area, '  坐标:', r1, r2, c1, c2)

            if fill:
                # 获取一个单元格的内容
                first_value = self.cell(r1, c1).value

                # 数据填充
                for r in range(r1, r2+1):  # 遍历行
                    if c2 - c1 > 0:  # 多个列，遍历列
                        for c in range(c1, c2+1):
                            self.cell(r, c).value = first_value
                    else:  # 一个列
                        self.cell(r, c1).value = first_value

        return self

    def toDataframe(self, columns_row=0):
        df = pd.DataFrame(self.values)
        # 重设df的columns，并丢掉包含column在内的之前的行
        df.columns = df.iloc[columns_row].tolist()
        df.drop(range(0, columns_row+1), inplace=True, axis=0)  # 丢弃行
        return df

    def merge_colist(self, by, colist,  alig=False):
        def get_samerows_on_col(self, by: str):
            ret_rowlists = []
            row = 0
            _col = self[by]
            _max_row = self.max_row-1
            while row < _max_row:
                cur_cell = _col[row]
                # 如果当前cell存在值且值不等于空
                if cur_cell.value and cur_cell.value != '':
                    check_row = row+1
                    # 寻找下面的cell截止到哪里不一致
                    while check_row <= _max_row:
                        check_cell = _col[check_row]
                        if check_cell.value and check_cell.value != '':
                            if cur_cell.value == check_cell.value:
                                check_row += 1
                            else:
                                check_row -= 1
                                break
                        else:
                            check_row -= 1
                            break

                    # 如果有连续两格以上的单元格内容相同，进行融合
                    if check_row-row >= 1:
                        ret_rowlists.append([row+1, check_row+1])

                    # 忽略已经合并的单元格
                    row = check_row
                row += 1
            return ret_rowlists

        def merge_cells_on_col(self, pair_rows, cols, alig=False):
            for rows in pair_rows:
                row, check_row = rows[0], rows[1]
                for col in cols:
                    msg = '%s%d:%s%d' % (col, row, col, check_row)
                    self.merge_cells(msg)
                    if alig:
                        self['%s%d' % (col, row)].alignment = Alignment(horizontal='center',
                                                                        vertical='center',
                                                                        wrap_text=True)
            return self

        pair_rows = get_samerows_on_col(self, by)
        return merge_cells_on_col(self, pair_rows, list(colist), alig)

    # 样式属性
    def sys_style(self, tar_sheet):
        rows = self.iter_rows()
        for row in rows:
            for cell in row:
                if isinstance(cell, Cell):
                    cell.copy_style(tar_sheet[cell.coordinate])

    # 行列属性
    def set_columns_hide(self, columns, enable: bool):
        if isinstance(columns, str):
            self.column_dimensions[columns].hidden = enable
        elif isinstance(columns, tuple):
            self.column_dimensions.group(
                start=columns[0], end=columns[1], hidden=enable)
        else:
            raise TypeError('columns is a tuple or str')

    def set_column_width(self, column, width):
        letter = None
        if isinstance(column, str):
            letter = column
        elif isinstance(column, int):
            letter = get_column_letter(column)

        if width == 0:
            # print(letter+":"+str(max_len))
            self.column_dimensions[letter].width = 2
        elif width <= 12:
            # print(letter+":"+str(max_len))
            self.column_dimensions[letter].width = width + 2
        elif width <= 50:
            # print(letter+":"+str(max_len))
            self.column_dimensions[letter].width = width + 2
        else:
            # print(letter+":"+str(max_len))
            self.column_dimensions[letter].width = 50
            for cell in self[letter]:
                cell.alignment = Alignment(wrap_text=True)

    def auto_width(self, col_list=None, begin_cal_row=1):
        def get_maxlength(self, col, begin_cal_row):
            """
            获取一个类型为object的Series中的最大占位长度，用于确定导出的xlsx文件的列宽
            col : 表头，也参与比较，解决有时候表头过长的问题
            """
            str_list = [str(cell.value) for cell in self[col] if cell != None]
            str_list = str_list[begin_cal_row-1:]
            len_list = []
            for elem in str_list:
                if elem == "None":
                    continue
                elem_split = list(elem)
                length = 0
                for c in elem_split:
                    if ord(c) <= 256:
                        length += 1
                    else:
                        length += 2
                len_list.append(length)
            if len(len_list) == 0:
                return 0
            return max(len_list)

        loop_range = None
        if col_list == None:
            # 全部列都自动宽度
            max_col = self.max_column
            # 由列数获得对应的列字母
            loop_range = [get_column_letter(i) for i in range(1, max_col+1)]
        elif isinstance(col_list, (tuple, list)):
            loop_range = col_list
        elif isinstance(col_list, str):
            loop_range = (col_list,)
        for letter in loop_range:
            max_len = get_maxlength(self, letter, begin_cal_row)
            self.set_column_width(letter, max_len)

    # 重写此方法，用于快速获取表格中的单元格
    def __getitem__(self, index):
        if type(index) == tuple:
            return self.cell(index[0], index[1])
        return super(Worksheet, self).__getitem__(index)

    # 重写此方法，用于获取自定义类型
    def cell(self, row, column, value=None):
        """
        Returns a cell object based on the given coordinates.

        Usage: cell(row=15, column=1, value=5)

        Calling `cell` creates cells in memory when they
        are first accessed.

        :param row: row index of the cell (e.g. 4)
        :type row: int

        :param column: column index of the cell (e.g. 3)
        :type column: int

        :param value: value of the cell (e.g. 5)
        :type value: numeric or time or string or bool or none

        :rtype: openpyxl.cell.cell.Cell
        """

        if row < 1 or column < 1:
            raise ValueError("Row or column values must be at least 1")

        cell = self._get_cell(row, column)
        if value is not None:
            cell.value = value

        return cell

    # 重写此方法，用于获取自定义类型
    def _get_cell(self, row, column):
        """
        Internal method for getting a cell from a worksheet.
        Will create a new cell if one doesn't already exist.
        """
        if not 0 < row < 1048577:
            raise ValueError("Row numbers must be between 1 and 1048576")
        coordinate = (row, column)
        if not coordinate in self._cells:
            cell = Cell(self, row=row, column=column)
            self._add_cell(cell)
        cell = self._cells[coordinate]
        if isinstance(cell, (Cell, MergedCell)):
            return cell

    # 重写此方法，用于获取自定义类型
    def append(self, iterable):
        """Appends a group of values at the bottom of the current sheet.

        * If it's a list: all values are added in order, starting from the first column
        * If it's a dict: values are assigned to the columns indicated by the keys (numbers or letters)

        :param iterable: list, range or generator, or dict containing values to append
        :type iterable: list|tuple|range|generator or dict

        Usage:

        * append(['This is A1', 'This is B1', 'This is C1'])
        * **or** append({'A' : 'This is A1', 'C' : 'This is C1'})
        * **or** append({1 : 'This is A1', 3 : 'This is C1'})

        :raise: TypeError when iterable is neither a list/tuple nor a dict

        """
        row_idx = self._current_row + 1

        if (isinstance(iterable, (list, tuple, range))
                or isgenerator(iterable)):
            for col_idx, content in enumerate(iterable, 1):
                if isinstance(content, Cell):
                    # compatible with write-only mode
                    cell = content
                    if cell.parent and cell.parent != self:
                        raise ValueError(
                            "Cells cannot be copied from other worksheets")
                    cell.parent = self
                    cell.column = col_idx
                    cell.row = row_idx
                else:
                    cell = Cell(self, row=row_idx,
                                column=col_idx, value=content)
                self._cells[(row_idx, col_idx)] = cell

        elif isinstance(iterable, dict):
            for col_idx, content in iterable.items():
                if isinstance(col_idx, str):
                    col_idx = column_index_from_string(col_idx)
                cell = Cell(self, row=row_idx, column=col_idx, value=content)
                self._cells[(row_idx, col_idx)] = cell

        else:
            self._invalid_row(iterable)

        self._current_row = row_idx

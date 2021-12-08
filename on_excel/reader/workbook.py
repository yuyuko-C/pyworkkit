import openpyxl.reader.workbook as o_book

from ..workbook.workbook import Workbook


class WorkbookParser(o_book.WorkbookParser):
    """
    继承了原 WorkbookParser 的类，为了使用自定义的 Workbook
    """

    # 重写此方法，用于使用自定义的 Workbook
    def __init__(self, archive, workbook_part_name, keep_links=True) -> None:
        self.archive = archive
        self.workbook_part_name = workbook_part_name
        self.wb = Workbook()
        self.keep_links = keep_links
        self.sheets = []
    pass

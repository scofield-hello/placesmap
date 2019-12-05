from logging import getLogger

from scrapy.exporters import BaseItemExporter
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.styles.colors import BLACK
from collections import OrderedDict


class ExcelItemExporter(BaseItemExporter):
    """
    导出为 Excel
    """
    __logger = getLogger('ExcelItemExporter')

    def __init__(self, filename: str, sheet_columns: dict, **kwargs):
        self._configure(kwargs, dont_fail=True)
        self.__filename = filename
        self.__workbook = Workbook()
        header_font = Font(name='宋体', bold=True, color=BLACK, size=11)
        header_alignment = Alignment(horizontal="center", vertical="center")
        active_sheet = self.__workbook.get_active_sheet()
        self.__workbook.remove(active_sheet)
        self.__workbook_sheets = []
        for (sheet_name, columns) in sheet_columns.items():
            worksheet = self.__workbook.create_sheet(title=sheet_name)
            worksheet.sheet_properties.tabColor = '6FB7B7'
            worksheet.page_setup.fitToHeight = 1
            worksheet.page_setup.fitToWidth = 1
            worksheet.row_dimensions[1].height = 20
            worksheet.row_dimensions[2].height = 20
            worksheet.merge_cells(start_row=1,
                                  start_column=1,
                                  end_row=2,
                                  end_column=len(columns))
            worksheet_header_cell = worksheet['A1']
            worksheet_header_cell.value = sheet_name
            worksheet_header_cell.font = header_font
            worksheet_header_cell.alignment = header_alignment
            worksheet.row_dimensions[3].height = 25
            for row in worksheet.iter_rows(min_row=3,
                                           max_row=3,
                                           min_col=1,
                                           max_col=len(columns)):
                for cell in row:
                    cell.value = columns[cell.col_idx - 1]
                    cell.alignment = header_alignment
                    cell.font = header_font
            self.__workbook_sheets.append(worksheet)
        self.__logger.info(
            '本次任务共创建%s个表格:[%s]' %
            (len(self.__workbook_sheets), self.__workbook_sheets))

    def export_item(self, item):
        item_dict = dict(self._get_serialized_fields(item))
        serialized_item_dict = {}
        sheet_list = []
        for (key, value) in item_dict.items():
            if key.startswith("F"):
                if isinstance(value, str) or value is None:
                    serialized_item_dict[key] = value
                elif isinstance(value, list):
                    sheet_list.append(value)
                else:
                    self.__logger.warning("字段被忽略，字段: %s = %s, 类型: %s" %
                                          (key, value, type(value)))
        ordered_dict = OrderedDict(
            sorted(serialized_item_dict.items(), key=lambda t: int(t[0][1:])))
        sheet_list.insert(0, ordered_dict)
        for index, sheet_item in enumerate(sheet_list):
            if isinstance(sheet_item, dict):
                self.__workbook_sheets[index].append(
                    list(ordered_dict.values()))
            elif isinstance(sheet_item, list):
                for data in sheet_item:
                    ordered_dict = OrderedDict(
                        sorted(data.items(), key=lambda t: int(t[0][1:])))
                    self.__workbook_sheets[index].append(
                        list(ordered_dict.values()))

    def finish_exporting(self):
        self.__logger.info('导出结束, 文档保存路径:%s' % self.__filename)
        self.__workbook.save(self.__filename)

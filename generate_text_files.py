import os
import sys
from abc import ABC
from pathlib import Path
from typing import Optional, Dict, Callable, Tuple

import xlrd  # Note: Use version 1.2.0 for .xlsx files
from xlrd.sheet import Sheet

from core.model import CoreBaseModel


class Column(CoreBaseModel):
    index: int = 0
    text_alteration: Optional[Callable]


class TextFileColumns(CoreBaseModel):
    file_name: Column
    content: Column


class Excel(CoreBaseModel):
    columns: Dict[str, TextFileColumns]


class ExcelTextFileGenerator(ABC):
    excel: Excel
    excel_sheet: Sheet
    has_header: bool = False

    def __init__(self,
                 excel_file_name: str,
                 excel: Excel,
                 sheet_index: Optional[int] = 0,
                 has_header: bool = False):
        self.excel = excel
        self.has_header = has_header,

        template_location = os.path.join(sys.path[0], f"{excel_file_name}")
        excel_workbook = xlrd.open_workbook(template_location)
        self.excel_sheet = excel_workbook.sheet_by_index(sheet_index)

    @staticmethod
    def mapping_to_excel_column_conversion(mapping: Dict[str, Dict[str, Tuple[int, Optional[Callable]]]]) -> Excel:
        '''
        Intended Mapping Usage:
                {"description_of_column": {"file_name": ("index", "text_alteration_callable"),
                                   "content": ("index", "text_alteration_callable")}}
        '''
        excel = Excel(columns=[])
        for description, column in mapping.items():
            converted_column = TextFileColumns(file_name=Column(index=column["file_name"][0],
                                                                text_alteration=column["file_name"][1]),
                                               content=Column(index=column["content"][0],
                                                              text_alteration=column["content"][1]))
            excel.columns[description] = converted_column
        return excel

    def create_templates(self,
                         name_column: Column,
                         content_column: Column,
                         description: str = "Template",
                         directory_name: str = "templates"):
        first_data_row = 1 if self.has_header else 0
        directory = os.path.join(*f"{directory_name}".split("."))
        Path(directory).mkdir(parents=True, exist_ok=True)
        for row in range(first_data_row, self.excel_sheet.nrows):
            print(f"{description} Generation: {row} of {self.excel_sheet.nrows}")
            file_name = self.excel_sheet.cell_value(row, name_column.index)
            content = self.excel_sheet.cell_value(row, content_column.index)
            if file_name != "" and content != "":
                if name_column.text_alteration is not None:
                    file_name = name_column.text_alteration(file_name)
                if content_column.text_alteration is not None:
                    content = content_column.text_alteration(content)
                text_file = open(os.path.join(directory, f"{file_name}.txt"), "w+")
                text_file.write(content)
                text_file.close()

    def generate_all_text_files(self):
        for type_of_file, column_indexes in self.excel.columns.items():
            self.create_templates(name_column=column_indexes.file_name,
                                  content_column=column_indexes.content,
                                  description=f"{type_of_file.upper()} File",
                                  directory_name=f"templates.{type_of_file.lower()}s")

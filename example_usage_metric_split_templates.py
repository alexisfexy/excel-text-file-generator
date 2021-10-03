import time
from typing import Optional, Dict, Callable, Tuple

from generate_text_files import ExcelTextFileGenerator

METRIC_SPLIT_TEXT_FILE_MAPPINGS: Dict[str, Dict[str, Tuple[int, Optional[Callable]]]] = {
    "metric":
        {"file_name": (0, None),
         "content": (1, None)},
    "split":
        {"file_name": (2, lambda string: string.removeprefix("    SplitKey.").removesuffix(": ")),
         "content": (4, None)}
}


def generate_metric_split_templates():
    excel = ExcelTextFileGenerator.mapping_to_excel_column_conversion(METRIC_SPLIT_TEXT_FILE_MAPPINGS)
    template_generator = ExcelTextFileGenerator(excel_file_name="Graphics_Templates.xlsx",
                                                excel=excel,
                                                sheet_index=0,
                                                has_header=False)
    template_generator.generate_all_text_files()


if __name__ == "__main__":
    start_time = time.time()
    generate_metric_split_templates()
    print(f"Generated all Templates in: {time.time() - start_time}")

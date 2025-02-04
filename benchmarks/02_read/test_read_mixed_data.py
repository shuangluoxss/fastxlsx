import pytest
from fastxlsx import RangeInfo, DShape, DType
from utils_read import LIBRARIES_READ
from base_read import BaseWorkbook


@pytest.fixture(scope="module")
def mixed_data():
    n = 16
    return {
        "title": RangeInfo((0, 0), DShape.Scalar(), dtype=DType.Str),
        "datetime": RangeInfo((0, 1), DShape.Scalar(), dtype=DType.DateTime),
        "flag": RangeInfo((1, 1), DShape.Scalar(), dtype=DType.Bool),
        "table_name": RangeInfo((2, 2), DShape.Scalar(), dtype=DType.Str),
        "table_header": RangeInfo((2, 3), DShape.Row(n), dtype=DType.Str),
        "table_index": RangeInfo((3, 2), DShape.Column(n), dtype=DType.Int),
        "table_value": RangeInfo((3, 3), DShape.Matrix(n, n), dtype=DType.Float),
        "date_list": RangeInfo((3, 1), DShape.Column(n), dtype=DType.Date),
    }


@pytest.mark.parametrize("Workbook", LIBRARIES_READ)
def test_read_mixed_data(benchmark, Workbook, mixed_data, xlsx_dir):
    benchmark.group = "read_mixed_data"
    library_name = Workbook.__name__.rstrip("Workbook")
    print(library_name)

    def read_with_library():
        input_filename = f"{xlsx_dir}/test_mixed_data_FastXLSX.xlsx"
        wb: BaseWorkbook = Workbook(input_filename)
        ws = wb.get_sheet("MixedData")
        wb.read_from_sheet(ws, mixed_data)

    benchmark(read_with_library)

import pytest
from fastxlsx import RangeInfo, DShape, DType
from utils_read import LIBRARIES_READ
from base_read import BaseWorkbook


@pytest.fixture(scope="module")
def big_matrix():
    n_rows, n_cols = 5000, 10
    return {
        "matrix": RangeInfo((0, 0), DShape.Matrix(n_rows, n_cols), dtype=DType.Float)
    }


@pytest.mark.parametrize("Workbook", LIBRARIES_READ)
def test_read_big_matrix(benchmark, Workbook, big_matrix, xlsx_dir):
    benchmark.group = "read_big_matrix"
    library_name = Workbook.__name__.rstrip("Workbook")
    print(library_name)

    def read_with_library():
        input_filename = f"{xlsx_dir}/test_big_matrix_FastXLSX.xlsx"
        wb: BaseWorkbook = Workbook(input_filename)
        ws = wb.get_sheet(0)
        wb.read_from_sheet(ws, big_matrix)

    benchmark(read_with_library)

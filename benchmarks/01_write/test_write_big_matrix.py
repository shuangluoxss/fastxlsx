import pytest
import numpy as np
from fastxlsx import RangeInfo, DShape, DType
from utils_write import LIBRARIES_WRITE
from base_write import BaseWorkbook


@pytest.fixture(scope="module")
def big_matrix():
    n_rows, n_cols = 5000, 10
    return np.random.random((n_rows, n_cols))


@pytest.mark.parametrize("Workbook", LIBRARIES_WRITE)
def test_write_big_matrix(benchmark, Workbook, big_matrix, xlsx_dir):
    benchmark.group = "write_big_matrix"
    library_name = Workbook.__name__.rstrip("Workbook")
    print(library_name)

    def write_with_library():
        output_filename = f"{xlsx_dir}/test_big_matrix_{library_name}.xlsx"
        sheetname = "BigMatrix"
        wb: BaseWorkbook = Workbook()
        # Special optimization for PyExcelerate and OpenPyXLWriteonly
        if library_name == "PyExcelerate":
            wb.wb.new_sheet(sheetname, data=big_matrix)
        elif library_name == "OpenPyXLWriteonly":
            ws = wb.create_sheet(sheetname)
            for row in big_matrix:
                ws.append(row.tolist())
        else:
            ws = wb.create_sheet(sheetname)
            range_info = RangeInfo(
                (0, 0), DShape.Matrix(*big_matrix.shape), dtype=DType.Float
            )
            wb.write_to_sheet(ws, [(range_info, big_matrix)])
        wb.save(output_filename)

    benchmark(write_with_library)

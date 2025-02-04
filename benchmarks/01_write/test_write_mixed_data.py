import pytest
import numpy as np
import datetime
from fastxlsx import RangeInfo, DShape, DType
from utils_write import LIBRARIES_WRITE
from base_write import BaseWorkbook


@pytest.fixture(scope="module")
def mixed_data():
    n = 16
    return [
        (RangeInfo((0, 0), DShape.Scalar(), dtype=DType.Str), "Datetime:"),
        (
            RangeInfo((0, 1), DShape.Scalar(), dtype=DType.DateTime),
            datetime.datetime.now(),
        ),
        (RangeInfo((1, 0), DShape.Scalar(), dtype=DType.Str), "Flag"),
        (RangeInfo((1, 1), DShape.Scalar(), dtype=DType.Bool), True),
        (RangeInfo((2, 2), DShape.Scalar(), dtype=DType.Str), f"A {n}x{n} table"),
        (
            RangeInfo((2, 3), DShape.Row(n), dtype=DType.Str),
            [f"Var_{i+1}" for i in range(n)],
        ),
        (
            RangeInfo((3, 2), DShape.Column(n), dtype=DType.Int),
            [i + 1 for i in range(n)],
        ),
        (
            RangeInfo((3, 3), DShape.Matrix(n, n), dtype=DType.Float),
            np.random.rand(n, n).tolist(),
        ),
        (
            RangeInfo((3, 1), DShape.Column(n), dtype=DType.Date),
            [datetime.date.today() + datetime.timedelta(days=i) for i in range(n)],
        ),
    ]


@pytest.mark.parametrize("Workbook", LIBRARIES_WRITE)
def test_write_mixed_data(benchmark, Workbook, mixed_data, xlsx_dir):
    benchmark.group = "write_mixed_data"
    library_name = Workbook.__name__.rstrip("Workbook")
    print(library_name)

    def write_with_library():
        output_filename = f"{xlsx_dir}/test_mixed_data_{library_name}.xlsx"
        wb: BaseWorkbook = Workbook()
        ws = wb.create_sheet("MixedData")
        wb.write_to_sheet(ws, mixed_data)
        wb.save(output_filename)

    benchmark(write_with_library)

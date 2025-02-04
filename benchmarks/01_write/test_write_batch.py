import pytest
import numpy as np
import datetime
from fastxlsx import RangeInfo, DShape, DType, write_many, WriteOnlyWorksheet
from utils_write import LIBRARIES_WRITE, range_info_to_tuple, range_info_from_tuple
from base_write import BaseWorkbook
from multiprocessing import Pool
import os


@pytest.fixture(scope="module")
def multi_sheet_data():
    n_sheets = 6
    n_workbooks = 10
    n = 16

    def gen_data():
        return [
            (range_info_to_tuple(range_info), value)
            for (range_info, value) in [
                (RangeInfo((0, 0), DShape.Scalar(), dtype=DType.Str), "Datetime:"),
                (
                    RangeInfo((0, 1), DShape.Scalar(), dtype=DType.DateTime),
                    datetime.datetime.now(),
                ),
                (RangeInfo((1, 0), DShape.Scalar(), dtype=DType.Str), "Flag"),
                (RangeInfo((1, 1), DShape.Scalar(), dtype=DType.Bool), True),
                (
                    RangeInfo((2, 2), DShape.Scalar(), dtype=DType.Str),
                    f"A {n}x{n} table",
                ),
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
                    [
                        datetime.date.today() + datetime.timedelta(days=i)
                        for i in range(n)
                    ],
                ),
            ]
        ]

    return [
        {f"sheet{i+1}": gen_data() for i in range(n_sheets)} for _ in range(n_workbooks)
    ]


def write_single_file(arg):
    Workbook, sheet_dict, output_filename = arg
    wb: BaseWorkbook = Workbook()
    for sheetname, data in sheet_dict.items():
        ws = wb.create_sheet(sheetname)
        wb.write_to_sheet(
            ws,
            [
                (range_info_from_tuple(*range_info), value)
                for (range_info, value) in data
            ],
        )
    wb.save(output_filename)


@pytest.mark.parametrize("Workbook", LIBRARIES_WRITE)
def test_write_batch(benchmark, Workbook, multi_sheet_data, xlsx_dir):
    benchmark.group = "write_batch"
    library_name = Workbook.__name__.rstrip("Workbook")

    def write_with_library():
        if library_name.startswith("FastXLSX"):
            workbooks_to_write = {}
            wb: BaseWorkbook = Workbook()
            for i, sheet_dict in enumerate(multi_sheet_data, start=1):
                output_filename = f"{xlsx_dir}/test_batch_{library_name}_{i:02d}.xlsx"
                ws_list = []
                for sheetname, data in sheet_dict.items():
                    ws = WriteOnlyWorksheet(sheetname)
                    wb.write_to_sheet(
                        ws,
                        [
                            (range_info_from_tuple(*range_info), value)
                            for (range_info, value) in data
                        ],
                    )
                    ws_list.append(ws)
                workbooks_to_write[output_filename] = ws_list
            write_many(workbooks_to_write)
        else:
            with Pool(processes=os.cpu_count()) as pool:
                pool.map(
                    write_single_file,
                    [
                        (
                            Workbook,
                            sheet_dict,
                            f"{xlsx_dir}/test_batch_{library_name}_{i:02d}.xlsx",
                        )
                        for i, sheet_dict in enumerate(multi_sheet_data, start=1)
                    ],
                )
            pool.close()
            pool.join()

    benchmark(write_with_library)

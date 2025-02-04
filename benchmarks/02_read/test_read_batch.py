import pytest
import numpy as np
from fastxlsx import RangeInfo, DShape, DType, read_many
from utils_read import LIBRARIES_READ, range_info_to_tuple, range_info_from_tuple
from base_read import BaseWorkbook
from multiprocessing import Pool
import os


@pytest.fixture(scope="module")
def multi_sheet_data():
    n_sheets = 6
    n_workbooks = 10
    n = 16

    def gen_data():
        range_info_template = {
            "title": RangeInfo((0, 0), DShape.Scalar(), dtype=DType.Str),
            "datetime": RangeInfo((0, 1), DShape.Scalar(), dtype=DType.DateTime),
            "flag": RangeInfo((1, 1), DShape.Scalar(), dtype=DType.Bool),
            "table_name": RangeInfo((2, 2), DShape.Scalar(), dtype=DType.Str),
            "table_header": RangeInfo((2, 3), DShape.Row(n), dtype=DType.Str),
            "table_index": RangeInfo((3, 2), DShape.Column(n), dtype=DType.Int),
            "table_value": RangeInfo((3, 3), DShape.Matrix(n, n), dtype=DType.Float),
            "date_list": RangeInfo((3, 1), DShape.Column(n), dtype=DType.Date),
        }
        return {
            key: range_info_to_tuple(range_info)
            for key, range_info in range_info_template.items()
        }

    return [
        {f"sheet{i_sheet+1}": gen_data() for i_sheet in range(n_sheets)}
        for _ in range(n_workbooks)
    ]


def read_single_file(arg):
    Workbook, sheet_dict, input_filename = arg
    wb: BaseWorkbook = Workbook(input_filename)
    res = {}
    for sheetname, range_to_read in sheet_dict.items():
        ws = wb.get_sheet(sheetname)
        res[sheetname] = wb.read_from_sheet(
            ws,
            {
                key: range_info_from_tuple(*range_info)
                for (key, range_info) in range_to_read.items()
            },
        )
    return res


@pytest.mark.parametrize("Workbook", LIBRARIES_READ)
def test_read_batch(benchmark, Workbook, multi_sheet_data, xlsx_dir):
    benchmark.group = "read_batch"
    library_name = Workbook.__name__.rstrip("Workbook")
    input_filename_template = "{xlsx_dir}/test_batch_FastXLSX_{i_workbook:02d}.xlsx"

    def read_with_library():
        if library_name == "FastXLSX":
            workbooks_to_read = {}
            for i_workbook, sheet_dict in enumerate(multi_sheet_data, start=1):
                input_filename = input_filename_template.format(
                    xlsx_dir=xlsx_dir, i_workbook=i_workbook
                )
                workbooks_to_read[input_filename] = {
                    sheetname: {
                        key: range_info_from_tuple(*range_info)
                        for (key, range_info) in range_to_read.items()
                    }
                    for sheetname, range_to_read in sheet_dict.items()
                }
            dat = list(read_many(workbooks_to_read).values())
            assert dat[0]["sheet1"]["table_value"].shape == (16, 16)
        else:
            with Pool(processes=os.cpu_count()) as pool:
                dat = pool.map(
                    read_single_file,
                    [
                        (
                            Workbook,
                            sheet_dict,
                            input_filename_template.format(
                                xlsx_dir=xlsx_dir, i_workbook=i_workbook
                            ),
                        )
                        for i_workbook, sheet_dict in enumerate(
                            multi_sheet_data, start=1
                        )
                    ],
                )
            pool.close()
            pool.join()
            assert np.asarray(dat[0]["sheet1"]["table_value"]).shape == (16, 16)

    benchmark(read_with_library)

Getting Started
===============
FastXLSX use seprated class for write and read xlsx file, here are some simple examples:

Writing
-------
.. code-block:: python

    import datetime
    import numpy as np
    from fastxlsx import DType, WriteOnlyWorkbook, WriteOnlyWorksheet, write_many

    # Initialize workbook
    wb = WriteOnlyWorkbook()
    ws = wb.create_sheet("sheet1")

    ws.write_cell((0, 0), "Hello World!")
    ws.write_cell((1, 0), True, dtype=DType.Bool)
    ws.write_cell("B1", datetime.datetime.now(), dtype=DType.DateTime)
    ws.write_row((4, 2), ["var_a", "var_b", "var_c"], dtype=DType.Str)
    ws.write_column((4, 0), [2.5, "xyz", datetime.date.today()], dtype=DType.Any)
    # If `dtype` is one of [DType.Bool, DType.Int, DType.Float], must pass a numpy array
    ws.write_matrix((5, 2), np.random.random((3, 3)), dtype=DType.Float)

    # Save to file
    wb.save("./example.xlsx")

    # Write multiple files in parallel
    workbooks_to_write = {}
    for i_workbook in range(10):
        ws_list = []
        for i_sheet in range(6):
            ws = WriteOnlyWorksheet(f"Sheet{i_sheet}")
            ws.write_cell("A1", 10 * i_workbook + i_sheet, dtype=DType.Int)
            ws.write_matrix((1, 1), np.random.random((3, 3)), dtype=DType.Float)
            ws_list.append(ws)
        workbooks_to_write[f"example_{i_workbook:02d}.xlsx"] = ws_list
    write_many(workbooks_to_write)

Reading
-------
.. code-block:: python

    from fastxlsx import DShape, DType, RangeInfo, ReadOnlyWorkbook, read_many

    # Load xlsx file
    wb = ReadOnlyWorkbook("./example.xlsx")
    # List all sheet names
    wb.sheetnames
    # Get a worksheet by index or name
    ws = wb.get_by_idx(0)
    # Read a single cell, notice the index is 0-based
    print(ws.cell_value((0, 0)))
    print(ws.cell_value("B1", dtype=DType.DateTime))
    # Read a column with `read_value` and `RangeInfo`
    print(ws.read_value(RangeInfo((4, 0), DShape.Column(3), dtype=DType.Any)))
    print(
        ws.read_values(
            {
                "var_a": RangeInfo((5, 2), DShape.Column(3), dtype=DType.Float),
                "matrix": RangeInfo((5, 2), DShape.Matrix(3, 3), dtype=DType.Float),
            }
        )
    )

    # Read multiple sheets
    print(wb.read_worksheets({"sheet1": [RangeInfo((2, 2), DShape.Scalar())]}))
    # Read multiple files in parallel
    print(
        read_many(
            {
                f"./example_{i_workbook:02d}.xlsx": {
                    f"Sheet{i_sheet}": [
                        RangeInfo((0, 0), DShape.Scalar()),
                        RangeInfo((1, 1), DShape.Matrix(3, 3)),
                    ]
                    for i_sheet in range(6)
                }
                for i_workbook in range(10)
            }
        )
    )
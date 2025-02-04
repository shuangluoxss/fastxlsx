# FastXLSX

[![PyPI](https://img.shields.io/pypi/v/fastxlsx)](https://pypi.org/project/fastxlsx/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/shuangluoxss/fastxlsx/blob/main/LICENSE)

**A lightweight, high-performance Python library for blazing-fast XLSX I/O operations.**  
Powered by Rust's [Calamine](https://github.com/tafia/calamine) (reading) and [Rust-XlsxWriter](https://github.com/jmcnamara/rust_xlsxwriter) (writing), with seamless Python integration via `PyO3`.

## ✨ Key Features

### ✅ Supported Capabilities

- **Data Types**: Native support for `bool`, `int`, `float`, `date`, `datetime`, and `str`.
- **Data Operations**: Scalars, rows, columns, matrices, and batch processing.
- **Coordinate Systems**: Dual support for **A1** (e.g., `B2`) and **R1C1** (e.g., `(2, 3)`) notation.
- **Parallel Processing**: Multi-threaded read/write operations for massive datasets.
- **Type Safety**: Full type hints and IDE-friendly documentation.
- **Blasting Performance**: 5-10x faster compared to `openpyxl`.

### 🚫 Current Limitations

- **File Formats**: Only XLSX (no XLS/XLSB support).
- **Formulas & Styling**: Cell formulas, merged cells, and formatting not supported.
- **Modifications**: Append/update operations on existing files unavailable.
- **Advanced Features**: Charts, images, and other advanced features not supported.

## 🏆 Performance Benchmarks

Tested on AMD Ryzen 7 5600X @ 3.7GHz (Ubuntu 24.04 VM) using `pytest-benchmark`.  
Full details could be obtained from [benchmarks](./benchmarks).

### 📝 Writing Performance (Lower is Better)

| library              | Mixed Data (ms) | 5000x10 Matrix(ms) | Batch Write (ms) |
| :------------------- | :-------------- | :----------------- | :--------------- |
| **fastxlsx**         | 0.97(1.00x)     | 62.06(1.00x)       | 7.77(1.00x)      |
| pyexcelerate         | 2.65(2.73x)     | 256.89(4.14x)      | 50.33(6.48x)     |
| xlsxwriter           | 5.03(5.19x)     | 297.14(4.79x)      | 61.25(7.89x)     |
| openpyxl(write_only) | 5.91(6.09x)     | 422.22(6.80x)      | 83.89(10.80x)    |
| openpyxl             | 6.25(6.44x)     | 737.30(11.88x)     | 83.65(10.77x)    |

### 📖 Reading Performance (Lower is Better)

| library         | Mixed Data (ms) | 5000x10 Matrix(ms) | Batch Write (ms) |
| :-------------- | :-------------- | :----------------- | :--------------- |
| **fastxlsx**    | 0.24(1.00x)     | 24.22(1.00x)       | 3.14(1.00x)      |
| pycalamine      | 0.32(1.30x)     | 33.51(1.38x)       | 28.25(8.99x)     |
| openpyxl        | 3.93(16.07x)    | 330.63(13.65x)     | 62.71(19.96x)    |

⚠️ **Windows Users Note**: Batch operations use `multiprocessing.Pool`, which may underperform due to `spawn` method limitations.

## 🛠️ Installation

### PyPI Install

```bash
pip install fastxlsx
```

### Source Build (Requires Rust Toolchain)

```bash
git clone https://github.com/shuangluoxss/fastxlsx.git
cd fastxlsx
pip install .
```

## 🚀 Quick Start Guide

### Writing

```python
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
```

### Reading

```python
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
```

_For full details, see [docs](./docs)._

## 📖 Motivation

As is well known, Excel is not a good format for performance, but due to its widely used nature, sometimes we have to handle massive XLSX datasets. When I do some postprocessing work in Python, a lot of time is wasted on reading and writing; and when I tried to speed it up by parallelization, the spawn feature in Windows disturb me again. Therefore, I decided to develop a xlsx read-write library with Rust+PyO3 to solve that.

Thanks to the high performance of `calamine` and `rust_xlsxwriter`, as well as the great work of `PyO3` and `maturin`, it is possible to do that by just binding them together with Python. Also thanks to the help of Deepseek enable me, a Rust beginner, could finish that.

## 📌 Future Plans

- Add support for formula and cell formatting  
  `rust_xlsxwriter` supports formula and cell formatting well so that is not too hard to implent them into `fastxlsx`. But personally, when I export a large amount of data, format is usually not important, so the priority of this item is not high.
- Improve error handling

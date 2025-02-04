import random
import string
from fastxlsx import RangeInfo, DShape, DType
from adapters_write import (
    FastXLSXWorkbook,
    OpenPyXLWorkbook,
    OpenPyXLWriteonlyWorkbook,
    XlsxWriterWorkbook,
    PyExcelerateWorkbook,
    PyLightXLWorkbook,
)

LIBRARIES_WRITE = [
    FastXLSXWorkbook,
    OpenPyXLWorkbook,
    OpenPyXLWriteonlyWorkbook,
    XlsxWriterWorkbook,
    PyExcelerateWorkbook,
    PyLightXLWorkbook,
]


def generate_random_string(length=10):
    characters = string.ascii_letters + string.digits
    random_string = "".join(random.choice(characters) for _ in range(length))
    return random_string


def range_info_to_tuple(range_info: RangeInfo):
    """Since RangeInfo could not be pickle, we need to convert it to a tuple"""
    pos_tuple = range_info.pos
    data_shape = range_info.data_shape
    if isinstance(data_shape, DShape.Scalar):
        data_shape_tuple = (0,)
    elif isinstance(data_shape, DShape.Row):
        data_shape_tuple = (1, data_shape.n_cols)
    elif isinstance(data_shape, DShape.Column):
        data_shape_tuple = (2, data_shape.n_rows)
    elif isinstance(data_shape, DShape.Matrix):
        data_shape_tuple = (3, data_shape.n_rows, data_shape.n_cols)
    dtype = range_info.dtype
    if dtype == DType.Int:
        dtype_int = 0
    elif dtype == DType.Float:
        dtype_int = 1
    elif dtype == DType.Str:
        dtype_int = 2
    elif dtype == DType.Bool:
        dtype_int = 3
    elif dtype == DType.Date:
        dtype_int = 4
    elif dtype == DType.DateTime:
        dtype_int = 5
    elif dtype == DType.Any:
        dtype_int = 6

    return (pos_tuple, data_shape_tuple, dtype_int)


def range_info_from_tuple(pos_tuple, data_shape_tuple, dtype_int):
    if data_shape_tuple[0] == 0:
        data_shape = DShape.Scalar()
    elif data_shape_tuple[0] == 1:
        data_shape = DShape.Row(data_shape_tuple[1])
    elif data_shape_tuple[0] == 2:
        data_shape = DShape.Column(data_shape_tuple[1])
    elif data_shape_tuple[0] == 3:
        data_shape = DShape.Matrix(data_shape_tuple[1], data_shape_tuple[2])
    dtype = [
        DType.Int,
        DType.Float,
        DType.Str,
        DType.Bool,
        DType.Date,
        DType.DateTime,
        DType.Any,
    ][dtype_int]
    return RangeInfo(pos_tuple, data_shape, dtype=dtype)

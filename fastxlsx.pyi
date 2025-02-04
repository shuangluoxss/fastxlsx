from typing import Any, Dict, List, Tuple, Union, NamedTuple, overload, Optional
import numpy as np
from enum import IntEnum

class DType(IntEnum):
    """Enumeration for data types."""

    Int: int
    Float: int
    Str: int
    Bool: int
    Date: int
    DateTime: int
    Any: int

class DShape:
    """Class to describe the shape of data."""
    class Scalar(NamedTuple): ...

    class Row(NamedTuple):
        n_cols: int

    class Column(NamedTuple):
        n_rows: int

    class Matrix(NamedTuple):
        n_rows: int
        n_cols: int

class RangeInfo:
    """Class to describe the range of data."""

    pos: Tuple[int, int]
    data_shape: DShape
    dtype: DType
    strict: bool
    def __init__(
        self,
        pos: Tuple[int, int],
        data_shape: DShape,
        *,
        dtype: DType = DType.Any,
        strict: bool = True,
    ) -> "RangeInfo":
        """Generate a RangeInfo object.

        Parameters
        ----------
        pos : Tuple[int, int]
            The 0-based starting position of the range as (row, col). \n
            Negative refer counted from the end of the sheet.
        data_shape : DShape
            The shape of the data in the range.
        dtype : DType, default DType.Any
            The data type of the range.
        strict : bool, default True
            Whether to enforce strict type checking. If True, raise error when
            the `dtype` does not match, else use default value.
        """
        ...
    @property
    def start(self) -> Tuple[int, int]:
        """The 0-based starting position of the range as (row, col)"""
        ...
    @property
    def end(self) -> Tuple[int, int]:
        """The 0-based ending position of the range as (row, col)"""
        ...
    @property
    def shape(self) -> Tuple[int, int]:
        """The shape of the range as (n_rows, n_cols)"""
        ...

class ReadOnlyWorksheet:
    """Read-only worksheet class"""
    def read_value(self, range_info: RangeInfo) -> Any:
        """Read a single value from the worksheet based on the specified range.

        Parameters
        ----------
        range_info : RangeInfo
            The range information describing the position, shape, and data type of the value to read.

        Returns
        -------
        Any
            The value read from the specified range. Could be scalar or 1d-array or 2d-array.
        """
        ...
    @overload
    def read_values(self, range_info_list: List[RangeInfo]) -> List[Any]:
        """Read multiple values from the worksheet based on a list of ranges.

        Parameters
        ----------
        range_info_list : List[RangeInfo]
            A list of `RangeInfo` objects, each describing a range of cells to read.

        Returns
        -------
        List[Any]
            A list of values read from the specified ranges.
        """
        ...
    @overload
    def read_values(self, range_info_list: Dict[str, RangeInfo]) -> Dict[str, Any]:
        """Read multiple values from the worksheet based on named ranges.

        Parameters
        ----------
        range_info_list : Dict[str, RangeInfo]
            A dictionary mapping string keys (representing named ranges) to `RangeInfo` objects,
            each describing a range of cells to read.

        Returns
        -------
        Dict[str, Any]
            A dictionary mapping the same string keys to the corresponding values read from the
            specified ranges.
        """
        ...
    def cell_value(
        self,
        cell_addr: Union[Tuple[int, int], str],
        *,
        dtype: DType = DType.Any,
        strict: bool = True,
    ) -> Any:
        """Read a value from a specific cell in the worksheet.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The cell address, either as a tuple of (row, col) or a string (e.g., "A1").
        dtype : DType, default DType.Any
            The expected data type of the cell value. If use DType.Any, will determine
            automatically.
        strict : bool, default True
            Whether to enforce strict type checking. If True, raise error when
            the `dtype` does not match, else use default value.

        Returns
        -------
        Any
            The value read from the specified cell as the specified `dtype`.
        """
        ...

class ReadOnlyWorkbook:
    """Read-only workbook class"""
    def __init__(self, path: str) -> "ReadOnlyWorkbook":
        """Generate a `ReadOnlyWorkbook` object.

        Parameters
        ----------
        path : str
            The path to the xlsx file.
        """
        ...
    def get_by_name(self, name: str) -> ReadOnlyWorksheet:
        """Get the sheet by sheet name.

        Parameters
        ----------
        name : str
            The name of the sheet.

        Returns
        -------
        ReadOnlyWorksheet
        """
        ...
    def get_by_idx(self, idx: int) -> ReadOnlyWorksheet:
        """Get the sheet by index.

        Parameters
        ----------
        idx : int
            The 0-based index of the sheet.

        Returns
        -------
        ReadOnlyWorksheet
        """
        ...
    def get(self, idx_or_name: Union[int, str]) -> ReadOnlyWorksheet:
        """Get the sheet by index or name.

        Parameters
        ----------
        idx_or_name : Union[int, str]
            The 0-based index or name of the sheet.

        Returns
        -------
        ReadOnlyWorksheet
        """
        ...
    @property
    def sheetnames(self) -> List[str]:
        """The names of the sheets."""
        ...
    @property
    def worksheets(self) -> List[ReadOnlyWorksheet]:
        """The sheets."""
        ...
    @overload
    def read_worksheets(
        self,
        worksheets_to_read: Dict[Union[str, int], List[RangeInfo]],
    ) -> Dict[Union[str, int], List[Any]]:
        """Read values from multiple worksheets based on specified ranges.

        Parameters
        ----------
        worksheets_to_read : Dict[Union[str, int], List[RangeInfo]]
            A dictionary mapping worksheet identifiers (either by name or index) to a list of
            `RangeInfo` objects. Each `RangeInfo` object describes a range of cells to read
            from the corresponding worksheet.

        Returns
        -------
        Dict[Union[str, int], List[Any]]
            A dictionary mapping worksheet identifiers (either by name or index) to a list of
            values read from the specified ranges, where each range is replaced by the corresponding 
            data.

        Examples
        --------
        >>> from fastxlsx import WorkbookReanonly, RangeInfo, DShape, DType
        >>> wb = WorkbookReanonly("example.xlsx")
        >>> worksheets_to_read = {
                "Sheet1": [
                    RangeInfo((0, 0), DShape.Matrix(2, 2), dtype=DType.Int),
                    RangeInfo((2, 0), DShape.Column(3), dtype=DType.Str),
                ],
                1: [RangeInfo((0, 0), DShape.Row(5), dtype=DType.Float)],
            }
        >>> res = wb.read_worksheets(worksheets_to_read)
        >>> res
        {
            "Sheet1": [np.array([[1, 2], [3, 4]]), ["A", "B", "C"]],
            1: [np.array([1.1, 2.2, 3.3, 4.4, 5.5])],
        }
        """
        ...

    @overload
    def read_worksheets(
        self,
        worksheets_to_read: Dict[Union[str, int], Dict[str, RangeInfo]],
    ) -> Dict[Union[str, int], Dict[str, Any]]:
        """Read values from multiple worksheets based on named ranges.

        This method reads data from one or more worksheets in the workbook, using the provided
        named ranges for each worksheet. Each named range is defined by a string key and a
        `RangeInfo` object, which specifies the position, shape, and data type of the data to
        be read.

        Parameters
        ----------
        worksheets_to_read : Dict[Union[str, int], Dict[str, RangeInfo]]
            A dictionary mapping worksheet identifiers (either by name or index) to a nested
            dictionary. The nested dictionary maps string keys (representing named ranges) to
            `RangeInfo` objects, which describe the ranges of cells to read from the
            corresponding worksheet.

        Returns
        -------
        Dict[Union[str, int], Dict[str, Any]]
            A dictionary mapping worksheet identifiers (either by name or index) to dictionary 
            of named ranges, where each named range is replaced by the corresponding data.

        Examples
        --------
        >>> from fastxlsx import WorkbookReanonly, RangeInfo, DShape, DType
        >>> wb = WorkbookReanonly("example.xlsx")
        >>> worksheets_to_read = {
                "Sheet1": {
                    "mat_1": RangeInfo((0, 0), DShape.Matrix(2, 2), dtype=DType.Int),
                    "col_1": RangeInfo((2, 0), DShape.Column(3), dtype=DType.Str),
                },
                1: {"row_1": RangeInfo((0, 0), DShape.Row(5), dtype=DType.Float)},
            }
        >>> res = wb.read_worksheets(worksheets_to_read)
        >>> res
        {
            "Sheet1": {"mat_1": np.array([[1, 2], [3, 4]]), "col_1": ["A", "B", "C"]},
            1: {"row_1": np.array([1.1, 2.2, 3.3, 4.4, 5.5])},
        }
        """
        ...

class WriteOnlyWorksheet:
    """Write-only worksheet class"""
    def __init__(self, title: str) -> "WriteOnlyWorksheet":
        """Initialize a new worksheet with the specified title.

        Parameters
        ----------
        title : str
            The title of the worksheet.
        """
        ...
    def write_cell(
        self,
        cell_addr: Union[Tuple[int, int], str],
        value: Any,
        *,
        dtype: Optional[DType] = None,
    ):
        """Write a value to a specific cell in the worksheet.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The cell address, either as a 0-based tuple of (row, col) or a string (e.g., "A1").
        value : Any
            The value to write to the cell.
        dtype : DType, default DType.Any
            The data type to enforce for the value. If None, try for each type automatically.
        """
        ...
    def write_row(
        self,
        cell_addr: Union[Tuple[int, int], str],
        value: Union[np.ndarray, List[Any]],
        *,
        dtype: Optional[DType] = None,
    ):
        """Write a row of values starting from a specific cell.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The starting cell address, either as a tuple of (row, col) or a string (e.g., "A1").
        value : Union[np.ndarray, List[Any]]
            The row of values to write. 1d array-like.\n
            MUST be `np.ndarray` with correct dtype if `dtype` is one of [`DType.Bool`, `DType.Int`, `DType.Float`]
        dtype : Optional[DType], default None
            The data type to enforce for the values. If None, will try for every possible types.\n
            DType.Any allow each item have different type, but will also increase the time cost.
        """
        ...
    def write_column(
        self,
        cell_addr: Union[Tuple[int, int], str],
        value: Union[np.ndarray, List[Any]],
        *,
        dtype: Optional[DType] = None,
    ):
        """Write a column of values starting from a specific cell.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The starting cell address, either as a tuple of (row, col) or a string (e.g., "A1").
        value : Union[np.ndarray, List[Any]]
            The column of values to write.\n
            MUST be `np.ndarray` with correct dtype if `dtype` is one of [`DType.Bool`, `DType.Int`, `DType.Float`]
        dtype : Optional[DType], default None
            The data type to enforce for the values. If None, will try for every possible types.\n
            DType.Any allow each item have different type, but will also increase the time cost.
        """
        ...
    def write_matrix(
        self,
        cell_addr: Union[Tuple[int, int], str],
        value: Union[np.ndarray, List[List[Any]]],
        *,
        dtype: Optional[DType] = None,
    ):
        """Write a matrix of values starting from a specific cell.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The starting cell address, either as a tuple of (row, col) or a string (e.g., "A1").
        value : Union[np.ndarray, List[List[Any]]]
            The matrix of values to write.\n
            MUST be `np.ndarray` with correct dtype if `dtype` is one of [`DType.Bool`, `DType.Int`, `DType.Float`]
        dtype : Optional[DType], default None
            The data type to enforce for the values. If None, will try for every possible types.\n
            DType.Any allow each item have different type, but will also increase the time cost.
        """
        ...

class WriteOnlyWorkbook:
    """Write-only workbook class"""
    def __init__(self) -> "WriteOnlyWorkbook": ...
    def create_sheet(self, name: str) -> WriteOnlyWorksheet:
        """Create a new worksheet with the specified name.

        Parameters
        ----------
        name : str
            The name of the new worksheet.

        Returns
        -------
        WriteOnlyWorksheet
            The newly created worksheet.
        """
        ...
    def get_by_idx(self, idx: int) -> WriteOnlyWorksheet:
        """Get a worksheet by its index.

        Parameters
        ----------
        idx : int
            The 0-based index of the worksheet.

        Returns
        -------
        WriteOnlyWorksheet
            The worksheet at the specified index.
        """
        ...
    def get_by_name(self, name: str) -> WriteOnlyWorksheet:
        """Get a worksheet by its name.

        Parameters
        ----------
        name : str
            The name of the worksheet.

        Returns
        -------
        WriteOnlyWorksheet
            The worksheet with the specified name.
        """
        ...
    def get(self, idx_or_name: Union[int, str]) -> WriteOnlyWorksheet:
        """Get a worksheet by its index or name.

        Parameters
        ----------
        idx_or_name : Union[int, str]
            The 0-based index or name of the worksheet.

        Returns
        -------
        WriteOnlyWorksheet
            The worksheet at the specified index or with the specified name.
        """
        ...
    def save(self, path: str):
        """Save the workbook to the specified file path.

        Parameters
        ----------
        path : str
            The file path where the workbook will be saved.
        """
        ...
    def sheetnames(self) -> List[str]:
        """Get the names of all worksheets in the workbook.

        Returns
        -------
        List[str]
            A list of worksheet names.
        """
        ...

@overload
def read_many(
    workbooks_to_read: Dict[str, Dict[Union[int, str], List[RangeInfo]]],
) -> Dict[str, Dict[Union[int, str], List[Any]]]:
    """Read values from multiple workbooks based on specified ranges.

    This function reads data from multiple workbooks, where each workbook is identified by its
    file path. For each workbook, a dictionary maps worksheet identifiers (either by name or
    index) to a list of `RangeInfo` objects, which describe the ranges of cells to read.

    Parameters
    ----------
    workbooks_to_read : Dict[str, Dict[Union[int, str], List[RangeInfo]]]
        A dictionary mapping workbook file paths to nested dictionaries. Each nested dictionary
        maps worksheet identifiers (either by name or index) to a list of `RangeInfo` objects.

    Returns
    -------
    Dict[str, Dict[Union[int, str], List[Any]]]
        A dictionary mapping workbook file paths to nested dictionaries. Each nested dictionary
        maps worksheet identifiers (either by name or index) to a list of values read from the
        specified ranges.

        The structure of the returned dictionary mirrors the input `workbooks_to_read`, with
        each `RangeInfo` replaced by the corresponding data.

    Examples
    --------
    >>> from fastxlsx import RangeInfo, DShape, DType, read_many
    >>> wb = WorkbookReanonly("example.xlsx")
    >>> workbooks_to_read = {
            "workbook1.xlsx": {
                "Sheet1": [
                    RangeInfo((0, 0), DShape.Matrix(2, 2), dtype=DType.Int),
                    RangeInfo((2, 0), DShape.Column(3), dtype=DType.Str),
                ],
                1: [RangeInfo((0, 0), DShape.Row(5), dtype=DType.Float)],
            },
            "workbook2.xlsx": {
                0: [RangeInfo((0, 0), DShape.Scalar(), dtype=DType.Str)],
            },
        }
    >>> res = read_many(workbooks_to_read)
    >>> res
        {
            "workbook1.xlsx": {
                "Sheet1": [np.array([[1, 2], [3, 4]]), ["A", "B", "C"]],
                1: [np.array([1.1, 2.2, 3.3, 4.4, 5.5])],
            },
            "workbook2.xlsx": {0: ["Some string"]},
        }
    """
    ...

@overload
def read_many(
    workbooks_to_read: Dict[str, Dict[Union[int, str], Dict[str, RangeInfo]]],
) -> Dict[str, Dict[Union[int, str], Dict[str, Any]]]:
    """Read values from multiple workbooks based on named ranges.

    This function reads data from multiple workbooks, where each workbook is identified by its
    file path. For each workbook, a dictionary maps worksheet identifiers (either by name or
    index) to a dictionary of named ranges.

    Parameters
    ----------
    workbooks_to_read : Dict[str, Dict[Union[int, str], Union[List[RangeInfo], Dict[str, RangeInfo]]]]
        A dictionary mapping workbook file paths to nested dictionaries. Each nested dictionary
        maps worksheet identifiers (either by name or index) to a dictionary of named ranges (string keys
        mapped to `RangeInfo` objects).

    Returns
    -------
    Dict[str, Dict[Union[int, str], Dict[str, Any]]]
        A dictionary mapping workbook file paths to nested dictionaries. Each nested dictionary
        maps worksheet identifiers (either by name or index) to a dictionary of named ranges,
        where each named range is replaced by the corresponding data.

        The structure of the returned dictionary mirrors the input `workbooks_to_read`, with
        each `RangeInfo` replaced by the corresponding data.

    Examples
    --------
    >>> from fastxlsx import RangeInfo, DShape, DType, read_many
    >>> wb = WorkbookReanonly("example.xlsx")
    >>> workbooks_to_read = {
            "workbook1.xlsx": {
                "Sheet1": {
                    "mat_1": RangeInfo((0, 0), DShape.Matrix(2, 2), dtype=DType.Int),
                    "col_1": RangeInfo((2, 0), DShape.Column(3), dtype=DType.Str),
                },
                1: {"row_1": RangeInfo((0, 0), DShape.Row(5), dtype=DType.Float)},
            },
            "workbook2.xlsx": {
                0: {"scalar_1": RangeInfo((0, 0), DShape.Scalar(), dtype=DType.Str)},
            },
        }
    >>> res = read_many(workbooks_to_read)
    >>> res
        {
            "workbook1.xlsx": {
                "Sheet1": {"mat_1": np.array([[1, 2], [3, 4]]), "col_1": ["A", "B", "C"]},
                1: {"row_1": np.array([1.1, 2.2, 3.3, 4.4, 5.5])},
            },
            "workbook2.xlsx": {0: {"scalar_1": "Some string"}},
        }
    """
    ...

def write_many(
    workbooks_to_write: Dict[str, List[WriteOnlyWorksheet]],
) -> None:
    """Write multiple workbooks to disk.

    This function writes multiple workbooks to their respective file paths. Each workbook is
    represented by a list of `WriteOnlyWorksheet` objects, which contain the data to be written.

    Parameters
    ----------
    workbooks_to_write : Dict[str, List[WriteOnlyWorksheet]]
        A dictionary mapping workbook file paths to lists of `WriteOnlyWorksheet` objects.
        Each `WriteOnlyWorksheet` object represents a worksheet containing data to be written.

    Examples
    --------
    >>> from fastxlsx import DType, WriteOnlyWorksheet, write_many
    >>> workbooks_to_write = {}
    >>> for i_workbook in range(10):
            ws_list = []
            for i_sheet in range(6):
                ws = WriteOnlyWorksheet(f"Sheet{i_sheet}")
                ws.write_cell("A1", 10 * i_workbook + i_sheet, dtype=DType.Int)
                ws.write_matrix((1, 1), np.random.random((3, 3)), dtype=DType.Float)
                ws_list.append(ws)
            workbooks_to_write[f"workboopk_{i_workbook}.xlsx"] = ws_list
    >>> write_many(workbooks_to_write)
    
    """
    ...

def idx_to_addr(row: int, col: int) -> str:
    """Convert a 0-based (row, col) index to a cell address string (e.g., "A1").

    Parameters
    ----------
    row : int
        The 0-based row index.
    col : int
        The 0-based column index.

    Returns
    -------
    str
        The cell address string (e.g., "A1").
    """
    ...

def addr_to_idx(addr: str) -> Tuple[int, int]:
    """Convert a cell address string (e.g., "A1") to a 0-based (row, col) index.

    Parameters
    ----------
    addr : str
        The cell address string (e.g., "A1").

    Returns
    -------
    Tuple[int, int]
        A tuple of (row, col) indices.
    """
    ...

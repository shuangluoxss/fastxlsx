from abc import ABC, abstractmethod
from typing import Any, List, Tuple
from fastxlsx import RangeInfo, DType, DShape


class BaseWorkbook(ABC):
    @abstractmethod
    def create_sheet(self, title: str) -> Any:
        pass

    @abstractmethod
    def write_func(self, ws: Any, row: int, col: int, value: Any, dtype: DType) -> None:
        pass

    @abstractmethod
    def save(self, filename: str) -> None:
        pass

    def write_to_sheet(
        self,
        ws: Any,
        data_to_write: List[Tuple[RangeInfo, Any]],
    ):
        for range_info, value in data_to_write:
            (row_start, col_start) = range_info.start
            data_shape = range_info.data_shape
            if isinstance(data_shape, DShape.Scalar):
                self.write_func(ws, row_start, col_start, value, range_info.dtype)
            elif isinstance(data_shape, DShape.Row):
                for i_col, v in enumerate(value):
                    self.write_func(
                        ws, row_start, col_start + i_col, v, range_info.dtype
                    )
            elif isinstance(data_shape, DShape.Column):
                for i_row, v in enumerate(value):
                    self.write_func(
                        ws, row_start + i_row, col_start, v, range_info.dtype
                    )
            elif isinstance(data_shape, DShape.Matrix):
                for i_row, row in enumerate(value):
                    for i_col, v in enumerate(row):
                        self.write_func(
                            ws,
                            row_start + i_row,
                            col_start + i_col,
                            v,
                            range_info.dtype,
                        )

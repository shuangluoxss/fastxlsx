from abc import ABC, abstractmethod
from typing import Any, Dict, Union
from fastxlsx import RangeInfo, DType, DShape


class BaseWorkbook(ABC):
    @abstractmethod
    def get_sheet(self, idx_or_name: Union[int, str]) -> Any:
        pass

    @abstractmethod
    def read_func(self, ws: Any, row: int, col: int, dtype: DType) -> None:
        pass

    def read_from_sheet(
        self,
        ws: Any,
        range_to_read: Dict[str, RangeInfo],
    ) -> Dict[str, Any]:
        res_dict = {}
        for key, range_info in range_to_read.items():
            (row_start, col_start) = range_info.start
            data_shape = range_info.data_shape
            dtype = range_info.dtype
            if isinstance(data_shape, DShape.Scalar):
                res_dict[key] = self.read_func(ws, row_start, col_start, dtype)
            elif isinstance(range_info.data_shape, DShape.Row):
                res_dict[key] = [
                    self.read_func(ws, row_start, col_start + i_col, dtype)
                    for i_col in range(data_shape.n_cols)
                ]
            elif isinstance(range_info.data_shape, DShape.Column):
                res_dict[key] = [
                    self.read_func(ws, row_start + i_row, col_start, dtype)
                    for i_row in range(data_shape.n_rows)
                ]
            elif isinstance(range_info.data_shape, DShape.Matrix):
                res_dict[key] = [
                    [
                        self.read_func(ws, row_start + i_row, col_start + i_col, dtype)
                        for i_col in range(data_shape.n_cols)
                    ]
                    for i_row in range(data_shape.n_rows)
                ]
        return res_dict

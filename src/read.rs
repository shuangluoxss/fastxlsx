use crate::fromcell::FromCell;
use crate::types::{Array1Container, Array2Container, ValueContainer, WrappedValue};
use crate::types::{CalamineData, CellAddr, DShape, DType, IdxOrName, ListOrDict, RangeInfo};
use crate::utils::adjust_idx;
use calamine::{open_workbook, Data, Range, Reader, Xlsx};
use chrono::{NaiveDate, NaiveDateTime};
use indexmap::IndexMap;
use ndarray::{Array1, Array2};
use pyo3::exceptions::PyFileExistsError;
use pyo3::prelude::*;
use rayon::prelude::*;
use std::fs::File;
use std::io::BufReader;
use std::path::PathBuf;

/// Read-only worksheet class
#[pyclass]
pub struct ReadOnlyWorksheet {
    pub sheet: Range<Data>,
    #[pyo3(get)]
    pub n_rows: usize,
    #[pyo3(get)]
    pub n_cols: usize,
    #[pyo3(get)]
    pub title: String,
}
impl ReadOnlyWorksheet {
    pub fn new(sheet: Range<Data>, title: String) -> Self {
        let (n_rows, n_cols) = sheet.get_size();
        Self {
            sheet,
            n_rows,
            n_cols,
            title,
        }
    }
    pub fn get_value_rs<T: FromCell>(&self, range_info: &RangeInfo) -> PyResult<ValueContainer<T>> {
        let data_shape = range_info.data_shape;
        let pos = (
            adjust_idx(range_info.pos.0, self.n_rows),
            adjust_idx(range_info.pos.1, self.n_cols),
        );
        let strict = range_info.strict;
        match data_shape {
            DShape::Scalar {} => Ok(ValueContainer::Scalar(T::from_cell(
                self.sheet.get(pos),
                strict,
            )?)),
            DShape::Row { n_cols } => {
                let arr_vec = (0..n_cols)
                    .into_iter()
                    .map(|j| T::from_cell(self.sheet.get((pos.0, pos.1 + j)), strict))
                    .collect::<PyResult<Vec<_>>>()?;
                let arr = { unsafe { Array1::from_shape_vec_unchecked(n_cols, arr_vec) } };
                Ok(ValueContainer::Array1(Array1Container { value: arr }))
            }
            DShape::Column { n_rows } => {
                let arr_vec = (0..n_rows)
                    .into_iter()
                    .map(|i| T::from_cell(self.sheet.get((pos.0 + i, pos.1)), strict))
                    .collect::<PyResult<Vec<_>>>()?;
                let arr = { unsafe { Array1::from_shape_vec_unchecked(n_rows, arr_vec) } };
                Ok(ValueContainer::Array1(Array1Container { value: arr }))
            }
            DShape::Matrix { n_rows, n_cols } => {
                let mut arr_vec = Vec::with_capacity(n_rows * n_cols);

                for i in 0..n_rows {
                    for j in 0..n_cols {
                        let cell = self.sheet.get((pos.0 + i, pos.1 + j));
                        let value = T::from_cell(cell, strict)?;
                        arr_vec.push(value);
                    }
                }
                let arr =
                    { unsafe { Array2::from_shape_vec_unchecked((n_rows, n_cols), arr_vec) } };
                Ok(ValueContainer::Array2(Array2Container { value: arr }))
            }
        }
    }
}

#[pymethods]
impl ReadOnlyWorksheet {
    /**
        Read a single value from the worksheet based on the specified range.

        Parameters
        ----------
        range_info : RangeInfo
            The range information describing the position, shape, and data type of the value to read.

        Returns
        -------
        Any
            The value read from the specified range. Could be scalar or 1d-array or 2d-array.
    */
    fn read_value(&self, range_info: &RangeInfo) -> PyResult<WrappedValue> {
        match range_info.dtype {
            DType::Int => self.get_value_rs::<i64>(range_info).map(WrappedValue::Int),
            DType::Float => self
                .get_value_rs::<f64>(range_info)
                .map(WrappedValue::Float),
            DType::Str => self
                .get_value_rs::<String>(range_info)
                .map(WrappedValue::Str),
            DType::Bool => self
                .get_value_rs::<bool>(range_info)
                .map(WrappedValue::Bool),
            DType::Date => self
                .get_value_rs::<NaiveDate>(range_info)
                .map(WrappedValue::Date),
            DType::DateTime => self
                .get_value_rs::<NaiveDateTime>(range_info)
                .map(WrappedValue::DateTime),
            DType::Any => self
                .get_value_rs::<CalamineData>(range_info)
                .map(WrappedValue::Any),
        }
    }
    /**
        Read multiple values from the worksheet based on a list of ranges.

        Parameters
        ----------
        range_info_list : List[RangeInfo] | Dict[str, RangeInfo]
            A list or dict of `RangeInfo` objects, each describing a range of cells to read.

        Returns
        -------
        List[Any] | Dict[str, Any]
            A list or dict of values read from the specified ranges.
    */
    fn read_values(
        &self,
        range_infos: ListOrDict<String, RangeInfo>,
    ) -> PyResult<ListOrDict<String, WrappedValue>> {
        match range_infos {
            ListOrDict::List(range_info_list) => range_info_list
                .iter()
                .map(|range_info| self.read_value(range_info))
                .collect::<PyResult<Vec<_>>>()
                .map(ListOrDict::List),
            ListOrDict::Dict(range_info_dict) => range_info_dict
                .iter()
                .map(|(k, range_info)| self.read_value(range_info).map(|v| (k.clone(), v)))
                .collect::<PyResult<IndexMap<_, _>>>()
                .map(ListOrDict::Dict),
        }
    }
    /**
        Read a value from a specific cell in the worksheet.

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
    */
    #[pyo3(signature = (cell_addr, *, dtype = DType::Any, strict = true))]
    fn cell_value(
        &self,
        cell_addr: CellAddr,
        dtype: DType,
        strict: bool,
    ) -> PyResult<WrappedValue> {
        let pos = cell_addr.as_idx()?;
        let pos = (pos.0 as i32, pos.1 as i32);
        let range_info = RangeInfo {
            pos,
            data_shape: DShape::Scalar {},
            dtype,
            strict,
        };
        self.read_value(&range_info)
    }
    fn __repr__(&self) -> String {
        format!("<ReadOnlyWorksheet \"{}\">", self.title)
    }
}

/// Read-only workbook class
#[pyclass]
pub struct ReadOnlyWorkbook {
    #[pyo3(get)]
    pub path: PathBuf,
    pub xlsx: Xlsx<BufReader<File>>,
    #[pyo3(get)]
    pub n_sheets: usize,
    #[pyo3(get)]
    pub sheetnames: Vec<String>,
}
#[pymethods]
impl ReadOnlyWorkbook {
    /**
        Generate a `ReadOnlyWorkbook` object.

        Parameters
        ----------
        path : str
            The path to the xlsx file.
    */
    #[new]
    pub fn new(path: PathBuf) -> PyResult<Self> {
        let xlsx: Xlsx<BufReader<File>> = open_workbook(&path).map_err(|e| {
            PyErr::new::<PyFileExistsError, _>(format!("{path:?} could not be read as xlsx: {e}"))
        })?;
        let sheetnames = xlsx
            .sheet_names()
            .iter()
            .map(|s| s.to_string())
            .collect::<Vec<_>>();
        let n_sheets = sheetnames.len();
        Ok(Self {
            path,
            xlsx,
            n_sheets,
            sheetnames,
        })
    }
    /**
        Get the sheet by sheet name.

        Parameters
        ----------
        name : str
            The name of the sheet.

        Returns
        -------
        ReadOnlyWorksheet
    */
    fn get_by_name(&mut self, sheet_name: String) -> PyResult<ReadOnlyWorksheet> {
        match self.xlsx.worksheet_range(&sheet_name) {
            Ok(sheet) => return Ok(ReadOnlyWorksheet::new(sheet, sheet_name)),
            Err(e) => return Err(PyErr::new::<PyFileExistsError, _>(format!("{e}"))),
        }
    }
    /**
        Get the sheet by index.

        Parameters
        ----------
        idx : int
            The 0-based index of the sheet.

        Returns
        -------
        ReadOnlyWorksheet
    */
    fn get_by_idx(&mut self, idx: usize) -> PyResult<ReadOnlyWorksheet> {
        if let Some(result) = self.xlsx.worksheet_range_at(idx) {
            match result {
                Ok(sheet) => {
                    return Ok(ReadOnlyWorksheet::new(
                        sheet,
                        self.sheetnames.get(idx).unwrap().to_owned(),
                    ))
                }
                Err(e) => return Err(PyErr::new::<PyFileExistsError, _>(format!("{e}"))),
            }
        } else {
            return Err(PyErr::new::<PyFileExistsError, _>(format!(
                "No sheet at index {idx}"
            )));
        }
    }
    /**
        Get the sheet by index or name.

        Parameters
        ----------
        idx_or_name : Union[int, str]
            The 0-based index or name of the sheet.

        Returns
        -------
        ReadOnlyWorksheet
    */
    fn get(&mut self, idx_or_name: IdxOrName) -> PyResult<ReadOnlyWorksheet> {
        match idx_or_name {
            IdxOrName::Idx(idx) => self.get_by_idx(adjust_idx(idx, self.n_sheets)),
            IdxOrName::Name(name) => self.get_by_name(name),
        }
    }
    /**
        Read values from multiple worksheets based on specified ranges.

        Parameters
        ----------
        worksheets_to_read : Dict[Union[str, int], Union[List[RangeInfo], Dict[str, RangeInfo]]]
            A dictionary mapping worksheet identifiers (either by name or index) to a list or dict of 
            `RangeInfo` objects. Each `RangeInfo` object describes a range of cells to read
            from the corresponding worksheet.

        Returns
        -------
        Dict[Union[str, int], Union[List[Any], Dict[str, Any]]]
            A dictionary mapping worksheet identifiers (either by name or index) to a list or dict of
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
    */
    fn read_worksheets(
        &mut self,
        worksheets_to_read: IndexMap<IdxOrName, ListOrDict<String, RangeInfo>>,
    ) -> PyResult<IndexMap<IdxOrName, ListOrDict<String, WrappedValue>>> {
        worksheets_to_read
            .into_iter()
            .map(|(idx_or_name, range_infos)| {
                let ws = self.get(idx_or_name.clone())?;
                Ok((idx_or_name, ws.read_values(range_infos)?))
            })
            .collect::<PyResult<IndexMap<IdxOrName, ListOrDict<String, WrappedValue>>>>()
    }
    /// The sheets.
    #[getter]
    fn worksheets(&mut self) -> PyResult<Vec<ReadOnlyWorksheet>> {
        (0..self.n_sheets)
            .into_iter()
            .map(|idx| self.get_by_idx(idx))
            .collect::<PyResult<Vec<_>>>()
    }
}
/**
    Read values from multiple workbooks based on specified ranges.

    This function reads data from multiple workbooks, where each workbook is identified by its
    file path. For each workbook, a dictionary maps worksheet identifiers (either by name or
    index) to a list of `RangeInfo` objects, which describe the ranges of cells to read.

    Parameters
    ----------
    workbooks_to_read : Dict[str, Dict[Union[int, str], Union[List[RangeInfo], Dict[str, RangeInfo]]]]
        A dictionary mapping workbook file paths to nested dictionaries. Each nested dictionary
        maps worksheet identifiers (either by name or index) to a list or dict of `RangeInfo` objects.

    Returns
    -------
    Dict[str, Dict[Union[int, str], Union[List[Any], Dict[str, Any]]]]
        A dictionary mapping workbook file paths to nested dictionaries. Each nested dictionary
        maps worksheet identifiers (either by name or index) to a list or dict of values read from the
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
*/
#[pyfunction]
pub fn read_many(
    workbooks_to_read: IndexMap<String, IndexMap<IdxOrName, ListOrDict<String, RangeInfo>>>,
) -> PyResult<IndexMap<String, IndexMap<IdxOrName, ListOrDict<String, WrappedValue>>>> {
    workbooks_to_read
        .into_par_iter()
        .map(|(path, worksheets)| {
            let mut workbook = ReadOnlyWorkbook::new(PathBuf::from(path.clone()))?;
            Ok((path, workbook.read_worksheets(worksheets)?))
        })
        .collect()
}

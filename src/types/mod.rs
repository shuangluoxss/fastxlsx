use indexmap::IndexMap;
use pyo3::prelude::*;
use std::hash::Hash;
mod calamine_data;
mod cell_addr;
mod containers;
pub use calamine_data::CalamineData;
pub use cell_addr::CellAddr;
pub use containers::{
    Array1Container, Array2Container, ValueContainer, WrappedValue, WriteToSheet,
};

/// Enumeration for data types.
#[pyclass(eq, eq_int)]
#[derive(PartialEq, Clone, Copy)]
pub enum DType {
    Int,
    Float,
    Str,
    Bool,
    Date,
    DateTime,
    Any,
}

/// Class to describe the shape of data.
#[pyclass(eq)]
#[derive(PartialEq, Clone, Copy)]
pub enum DShape {
    Scalar {},
    Row { n_cols: usize },
    Column { n_rows: usize },
    Matrix { n_rows: usize, n_cols: usize },
}

/// Class to describe the range of data.
#[pyclass]
#[derive(Clone)]
pub struct RangeInfo {
    #[pyo3(get, set)]
    pub pos: (i32, i32),
    #[pyo3(get, set)]
    pub data_shape: DShape,
    #[pyo3(get, set)]
    pub dtype: DType,
    #[pyo3(get, set)]
    pub strict: bool,
}
#[pymethods]
impl RangeInfo {
    /**
        Generate a RangeInfo object.

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
    */
    #[new]
    #[pyo3(signature = (pos, data_shape = DShape::Scalar{}, *, dtype = DType::Any, strict = true))]
    pub fn new(pos: (i32, i32), data_shape: DShape, dtype: DType, strict: bool) -> Self {
        Self {
            pos,
            data_shape,
            dtype,
            strict,
        }
    }
    /// The shape of the range as (n_rows, n_cols)
    #[getter]
    pub fn shape(&self) -> (usize, usize) {
        match self.data_shape {
            DShape::Scalar {} => (1, 1),
            DShape::Row { n_cols } => (1, n_cols),
            DShape::Column { n_rows } => (n_rows, 1),
            DShape::Matrix { n_rows, n_cols } => (n_rows, n_cols),
        }
    }
    /// The 0-based starting position of the range as (row, col)
    #[getter]
    pub fn start(&self) -> (i32, i32) {
        self.pos
    }
    /// The 0-based ending position of the range as (row, col)
    #[getter]
    pub fn end(&self) -> (i32, i32) {
        let (n_rows, n_cols) = self.shape();
        (
            self.pos.0 + n_rows as i32 - 1,
            self.pos.1 + n_cols as i32 - 1,
        )
    }
}

#[derive(Clone, PartialEq, Eq, Hash, FromPyObject, IntoPyObject)]
pub enum IdxOrName {
    Idx(i32),
    Name(String),
}

#[derive(Clone, FromPyObject, IntoPyObject)]
pub enum ListOrDict<K: PartialEq + Eq + Hash, T> {
    List(Vec<T>),
    Dict(IndexMap<K, T>),
}

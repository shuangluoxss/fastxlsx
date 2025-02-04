use crate::types::CalamineData;
use chrono::{NaiveDate, NaiveDateTime};
use ndarray::Array2;
use numpy::{PyArray2, PyArrayMethods, ToPyArray};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::PyAny;

use super::WriteToSheet;

#[derive(Clone)]
pub struct Array2Container<T> {
    pub value: Array2<T>,
}
impl<T: Clone> Array2Container<T> {
    pub fn new(value: Array2<T>) -> Self {
        Self { value }
    }
    pub fn len(&self) -> usize {
        self.value.len()
    }
    pub fn nrows(&self) -> usize {
        self.value.nrows()
    }
    pub fn ncols(&self) -> usize {
        self.value.ncols()
    }
    pub fn mapv<U>(&self, f: impl Fn(T) -> U) -> Array2Container<U> {
        Array2Container {
            value: self.value.mapv(f),
        }
    }
}
macro_rules! impl_into_pyobject_for_array2_container_numeric {
    ($($type:ty)*) => ($(
        impl<'py> IntoPyObject<'py> for Array2Container<$type> {
            type Target = PyAny;
            type Output = Bound<'py, Self::Target>;
            type Error = PyErr;
            fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
                Ok(self.value.to_pyarray(py).into_any())
            }
        }
    )*)
}
macro_rules! impl_into_pyobject_for_array2_container_other {
    ($($type:ty)*) => ($(
        impl<'py> IntoPyObject<'py> for Array2Container<$type> {
            type Target = PyAny;
            type Output = Bound<'py, Self::Target>;
            type Error = PyErr;
            fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
                Ok(self.value
                    .outer_iter()
                    .map(|row| row.to_vec())
                    .collect::<Vec<_>>()
                    .into_pyobject(py)?
                    .into_any())
            }
        }
    )*)
}

impl_into_pyobject_for_array2_container_numeric!(bool i64 f64);
impl_into_pyobject_for_array2_container_other!(String NaiveDate NaiveDateTime CalamineData);

macro_rules! impl_from_py_for_array2_container_numeric {
    ($err_msg:literal, $target_type:ty, $(($($source_type:ty)*))?) => {
        impl<'py> FromPyObject<'py> for Array2Container<$target_type> {
            fn extract_bound(ob: &Bound<'py, PyAny>) -> PyResult<Self> {
                if let Ok(array) = ob.downcast::<PyArray2<$target_type>>() {
                    return Ok(Self { value: array.to_owned_array() });
                }
                $(
                    $(
                        if let Ok(array) = ob.downcast::<PyArray2<$source_type>>() {
                            return Ok(Self {
                                value: array.to_owned_array().mapv(<$target_type>::from)
                            });
                        }
                    )*
                )?

                Err(PyTypeError::new_err(
                    concat!("Not a NumPy array of ", $err_msg, " type")
                ))
            }
        }
    };
}
macro_rules! impl_from_py_for_array2_container_other {
    ($($type:ty)*) => ($(
        impl<'py> FromPyObject<'py> for Array2Container<$type> {
            fn extract_bound(ob: &Bound<'py, PyAny>) -> PyResult<Self> {
                if let Ok(mat) = ob.extract::<Vec<Vec<$type>>>() {
                    let mut res_vec: Vec<$type> = Vec::new();
                    let n_rows = mat.len();
                    let n_cols = mat[0].len();
                    for row in &mat {
                        if row.len() != n_cols {
                            return Err(PyTypeError::new_err(
                                "Not a 2d-list: Item have different size",
                            ));
                        };
                        res_vec.extend_from_slice(row);
                    }
                    return Ok(Self {
                        value: {
                            unsafe { Array2::from_shape_vec_unchecked((n_rows, n_cols), res_vec) }
                        },
                    });
                }
                Err(PyTypeError::new_err(format!("Not a 2d-list of {}", stringify!($type))))
            }
        }
    )*)
}

impl_from_py_for_array2_container_numeric!("Bool", bool, ());
impl_from_py_for_array2_container_numeric!("Int", i64, (i32 i16 i8));
impl_from_py_for_array2_container_numeric!("Float", f64, (f32));
impl_from_py_for_array2_container_other!(String NaiveDate NaiveDateTime CalamineData);

macro_rules! impl_into_exceldata_for_array2_container_other {
    ($($type:ty)*) => ($(
        impl WriteToSheet for Array2Container<$type> {
            fn write_to_sheet(
                &self,
                sheet: &mut rust_xlsxwriter::Worksheet,
                start: (u32, u16),
                _is_column: bool,
            ) -> Result<(), rust_xlsxwriter::XlsxError> {
                sheet.write_row_matrix(
                    start.0,
                    start.1,
                    self.value.outer_iter().map(|row| row.to_vec()),
                )?;
                Ok(())
            }
        }
    )*)
}

macro_rules! impl_into_exceldata_for_array2_container_datetime {
    ($($type:ty)* => $format:expr) => ($(
        impl WriteToSheet for Array2Container<$type> {
            fn write_to_sheet(
                &self,
                sheet: &mut rust_xlsxwriter::Worksheet,
                start: (u32, u16),
                _is_column: bool,
            ) -> Result<(), rust_xlsxwriter::XlsxError> {
                let date_format = rust_xlsxwriter::Format::new().set_num_format($format);
                self.value.indexed_iter().try_for_each(|((i, j), v)| {
                    sheet
                        .write_datetime_with_format(start.0 + i as u32, start.1 + j as u16, v, &date_format)
                        .map(|_| ())
                })?;
                Ok(())
            }
        }
    )*)
}
impl_into_exceldata_for_array2_container_other!(bool i64 f64 String CalamineData);
impl_into_exceldata_for_array2_container_datetime! {NaiveDate => "yyyy/mm/dd"}
impl_into_exceldata_for_array2_container_datetime! {NaiveDateTime => "yyyy/mm/dd hh:mm:ss"}

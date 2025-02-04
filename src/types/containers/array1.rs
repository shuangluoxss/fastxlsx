use crate::types::CalamineData;
use chrono::{NaiveDate, NaiveDateTime};
use ndarray::Array1;
use numpy::{PyArray1, PyArrayMethods, ToPyArray};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::PyAny;

use super::WriteToSheet;

#[derive(Clone)]
pub struct Array1Container<T> {
    pub value: Array1<T>,
}

impl<T: Clone> Array1Container<T> {
    pub fn new(value: Array1<T>) -> Self {
        Self { value }
    }
    pub fn len(&self) -> usize {
        self.value.len()
    }
    pub fn mapv<U>(&self, f: impl Fn(T) -> U) -> Array1Container<U> {
        Array1Container {
            value: self.value.mapv(f),
        }
    }
}

macro_rules! impl_into_pyobject_for_array1_container_numeric {
    ($($type:ty)*) => ($(
        impl<'py> IntoPyObject<'py> for Array1Container<$type> {
            type Target = PyAny;
            type Output = Bound<'py, Self::Target>;
            type Error = PyErr;
            fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
                Ok(self.value.to_pyarray(py).into_any())
            }
        }
    )*)
}
macro_rules! impl_into_pyobject_for_array1_container_other {
    ($($type:ty)*) => ($(
        impl<'py> IntoPyObject<'py> for Array1Container<$type> {
            type Target = PyAny;
            type Output = Bound<'py, Self::Target>;
            type Error = PyErr;
            fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
                Ok(self.value.to_vec().into_pyobject(py)?.into_any())
            }
        }
    )*)
}

impl_into_pyobject_for_array1_container_numeric!(bool i64 f64);
impl_into_pyobject_for_array1_container_other!(String NaiveDate NaiveDateTime CalamineData);

macro_rules! impl_from_py_for_array1_container_numeric {
    ($err_msg:literal, $target_type:ty, $(($($source_type:ty)*))?) => {
        impl<'py> FromPyObject<'py> for Array1Container<$target_type> {
            fn extract_bound(ob: &Bound<'py, PyAny>) -> PyResult<Self> {
                if let Ok(array) = ob.downcast::<PyArray1<$target_type>>() {
                    return Ok(Self { value: array.to_owned_array() });
                }
                $(
                    $(
                        if let Ok(array) = ob.downcast::<PyArray1<$source_type>>() {
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
macro_rules! impl_from_py_for_array1_container_other {
    ($($type:ty)*) => ($(
        impl<'py> FromPyObject<'py> for Array1Container<$type> {
            fn extract_bound(ob: &Bound<'py, PyAny>) -> PyResult<Self> {
                if let Ok(array) = ob.extract::<Vec<$type>>() {
                    return Ok(Self{value: Array1::from_vec(array)});
                }
                Err(PyTypeError::new_err(format!("Not a List of {}", stringify!($type))))
            }
        }
    )*)
}
impl_from_py_for_array1_container_numeric!("Bool", bool, ());
impl_from_py_for_array1_container_numeric!("Int", i64, (i32 i16 i8));
impl_from_py_for_array1_container_numeric!("Float", f64, (f32));
impl_from_py_for_array1_container_other!(String NaiveDate NaiveDateTime CalamineData);

macro_rules! impl_into_exceldata_for_array1_container_other {
    ($($type:ty)*) => ($(
        impl WriteToSheet for Array1Container<$type> {
            fn write_to_sheet(
                &self,
                sheet: &mut rust_xlsxwriter::Worksheet,
                start: (u32, u16),
                is_column: bool,
            ) -> Result<(), rust_xlsxwriter::XlsxError> {
                if is_column {
                    sheet.write_column(start.0, start.1, self.value.clone())?
                } else {
                    sheet.write_row(start.0, start.1, self.value.clone())?
                };
                Ok(())
            }
        }
    )*)
}

macro_rules! impl_into_exceldata_for_array1_container_datetime {
    ($($type:ty)* => $format:expr) => ($(
        impl WriteToSheet for Array1Container<$type> {
            fn write_to_sheet(
                &self,
                sheet: &mut rust_xlsxwriter::Worksheet,
                start: (u32, u16),
                is_column: bool,
            ) -> Result<(), rust_xlsxwriter::XlsxError> {
                let date_format = rust_xlsxwriter::Format::new().set_num_format($format);
                if is_column {
                    self.value.indexed_iter().try_for_each(|(i, v)| {
                        sheet
                            .write_datetime_with_format(start.0 + i as u32, start.1, v, &date_format)
                            .map(|_| ())
                    })?
                } else {
                    self.value.indexed_iter().try_for_each(|(j, v)| {
                        sheet
                            .write_datetime_with_format(start.0, start.1 + j as u16, v, &date_format)
                            .map(|_| ())
                    })?
                };
                Ok(())
            }
        }
    )*)
}
impl_into_exceldata_for_array1_container_other!(bool i64 f64 String CalamineData);
impl_into_exceldata_for_array1_container_datetime! {NaiveDate => "yyyy/mm/dd"}
impl_into_exceldata_for_array1_container_datetime! {NaiveDateTime => "yyyy/mm/dd hh:mm:ss"}

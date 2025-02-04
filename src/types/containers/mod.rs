use crate::types::{CalamineData, DShape, DType};
use chrono::{NaiveDate, NaiveDateTime};
use pyo3::prelude::*;
use pyo3::types::PyAny;

mod array1;
mod array2;
pub use array1::Array1Container;
pub use array2::Array2Container;

#[derive(Clone)]
pub enum ValueContainer<T> {
    Scalar(T),
    Array1(Array1Container<T>),
    Array2(Array2Container<T>),
}
impl<T: Clone> ValueContainer<T> {
    pub fn get_shape(&self, is_column: bool) -> DShape {
        match self {
            ValueContainer::Scalar(_) => DShape::Scalar {},
            ValueContainer::Array1(arr1) => {
                if is_column {
                    DShape::Column { n_rows: arr1.len() }
                } else {
                    DShape::Row { n_cols: arr1.len() }
                }
            }
            ValueContainer::Array2(arr2) => DShape::Matrix {
                n_rows: arr2.nrows(),
                n_cols: arr2.ncols(),
            },
        }
    }
}
macro_rules! impl_into_pyobject_for_value_container {
    ($($type:ty)*) => ($(
        impl<'py> IntoPyObject<'py> for ValueContainer<$type> {
            type Target = PyAny;
            type Output = Bound<'py, Self::Target>;
            type Error = PyErr;

            fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
                match self {
                    ValueContainer::Scalar(v) => Ok(v.into_pyobject(py)?.as_any().clone()),
                    ValueContainer::Array1(v) => v.into_pyobject(py),
                    ValueContainer::Array2(v) => v.into_pyobject(py),
                }
            }
        }
    )*)
}

impl_into_pyobject_for_value_container!(bool i64 f64 String NaiveDate NaiveDateTime CalamineData);

// macro_rules! impl_from_pyobject_for_value_container {
//     ($($type:ty)*) => ($(
//         impl<'py> FromPyObject<'py> for ValueContainer<$type> {
//             fn extract_bound(ob: &Bound<'py, PyAny>) -> PyResult<Self> {
//                 if let Ok(val) = ob.extract::<$type>() {
//                     Ok(Self::Scalar(val))
//                 } else if let Ok(arr1) = ob.extract::<Array1Container<$type>>() {
//                     Ok(Self::Array1(arr1))
//                 } else if let Ok(arr2) = ob.extract::<Array2Container<$type>>() {
//                     Ok(Self::Array2(arr2))
//                 }
//                 else{
//                     Err(PyTypeError::new_err("Unsupported Python type"))
//                 }
//             }
//         }
//     )*)
// }
// impl_from_pyobject_for_value_container!(bool i64 f64 String NaiveDate NaiveDateTime CalamineData);

impl<T: Clone> ValueContainer<T> {
    pub fn mapv<U>(self, f: impl Fn(T) -> U) -> ValueContainer<U> {
        match self {
            ValueContainer::Scalar(value) => ValueContainer::Scalar(f(value)),
            ValueContainer::Array1(array) => ValueContainer::Array1(array.mapv(f)),
            ValueContainer::Array2(array) => ValueContainer::Array2(array.mapv(f)),
        }
    }
}

pub trait WriteToSheet {
    fn write_to_sheet(
        &self,
        sheet: &mut rust_xlsxwriter::Worksheet,
        start: (u32, u16),
        is_column: bool,
    ) -> Result<(), rust_xlsxwriter::XlsxError>;
}

macro_rules! impl_into_exceldata_for_value_container_other {
    ($($type:ty)*) => ($(
        impl WriteToSheet for ValueContainer<$type> {
            fn write_to_sheet(
                &self,
                sheet: &mut rust_xlsxwriter::Worksheet,
                start: (u32, u16),
                is_column: bool,
            ) -> Result<(), rust_xlsxwriter::XlsxError> {
                match self {
                    ValueContainer::Scalar(v) => sheet.write(start.0, start.1, v.clone()).map(|_| ()),
                    ValueContainer::Array1(arr1) => arr1.write_to_sheet(sheet, start, is_column),
                    ValueContainer::Array2(arr2) => arr2.write_to_sheet(sheet, start, is_column),
                }
            }
        }
    )*)
}

macro_rules! impl_into_exceldata_for_value_container_datetime {
    ($($type:ty)* => $format:expr) => ($(
        impl WriteToSheet for ValueContainer<$type> {
            fn write_to_sheet(
                &self,
                sheet: &mut rust_xlsxwriter::Worksheet,
                start: (u32, u16),
                is_column: bool,
            ) -> Result<(), rust_xlsxwriter::XlsxError> {
                let date_format = rust_xlsxwriter::Format::new().set_num_format($format);
                match self {
                    ValueContainer::Scalar(v) => sheet.write_datetime_with_format(start.0, start.1, v, &date_format).map(|_| ()),
                    ValueContainer::Array1(arr1) => arr1.write_to_sheet(sheet, start, is_column),
                    ValueContainer::Array2(arr2) => arr2.write_to_sheet(sheet, start, is_column)
                }
            }
        }
    )*)
}

impl_into_exceldata_for_value_container_other!(bool i64 f64 String CalamineData);
impl_into_exceldata_for_value_container_datetime! {NaiveDate => "yyyy/mm/dd"}
impl_into_exceldata_for_value_container_datetime! {NaiveDateTime => "yyyy/mm/dd hh:mm:ss"}

#[derive(Clone, IntoPyObject)]
pub enum WrappedValue {
    Bool(ValueContainer<bool>),
    Int(ValueContainer<i64>),
    Float(ValueContainer<f64>),
    Str(ValueContainer<String>),
    Date(ValueContainer<NaiveDate>),
    DateTime(ValueContainer<NaiveDateTime>),
    Any(ValueContainer<CalamineData>),
}
impl WrappedValue {
    pub fn get_dtype(&self) -> DType {
        match self {
            WrappedValue::Int(_) => DType::Int,
            WrappedValue::Float(_) => DType::Float,
            WrappedValue::Str(_) => DType::Str,
            WrappedValue::Bool(_) => DType::Bool,
            WrappedValue::Date(_) => DType::Date,
            WrappedValue::DateTime(_) => DType::DateTime,
            WrappedValue::Any(_) => DType::Any,
        }
    }

    pub fn get_shape(&self, is_column: bool) -> DShape {
        match self {
            WrappedValue::Int(ref v) => v.get_shape(is_column),
            WrappedValue::Float(ref v) => v.get_shape(is_column),
            WrappedValue::Str(ref v) => v.get_shape(is_column),
            WrappedValue::Bool(ref v) => v.get_shape(is_column),
            WrappedValue::Date(ref v) => v.get_shape(is_column),
            WrappedValue::DateTime(ref v) => v.get_shape(is_column),
            WrappedValue::Any(ref v) => v.get_shape(is_column),
        }
    }
}
impl WriteToSheet for WrappedValue {
    fn write_to_sheet(
        &self,
        sheet: &mut rust_xlsxwriter::Worksheet,
        start: (u32, u16),
        is_column: bool,
    ) -> Result<(), rust_xlsxwriter::XlsxError> {
        match self {
            WrappedValue::Int(ref v) => v.write_to_sheet(sheet, start, is_column),
            WrappedValue::Float(ref v) => v.write_to_sheet(sheet, start, is_column),
            WrappedValue::Str(ref v) => v.write_to_sheet(sheet, start, is_column),
            WrappedValue::Bool(ref v) => v.write_to_sheet(sheet, start, is_column),
            WrappedValue::Date(ref v) => v.write_to_sheet(sheet, start, is_column),
            WrappedValue::DateTime(ref v) => v.write_to_sheet(sheet, start, is_column),
            WrappedValue::Any(ref v) => v.write_to_sheet(sheet, start, is_column),
        }
    }
}

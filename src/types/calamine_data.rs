use chrono::{NaiveDate, NaiveDateTime};
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::IntoPyObject;
use rust_xlsxwriter::IntoExcelData;
#[derive(Clone)]
pub enum CalamineData {
    Int(i64),
    Float(f64),
    Str(String),
    Bool(bool),
    Date(NaiveDate),
    DateTime(NaiveDateTime),
    Empty,
}
impl<'py> IntoPyObject<'py> for CalamineData {
    type Target = PyAny;
    type Output = Bound<'py, Self::Target>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        match self {
            CalamineData::Int(v) => Ok(v.into_pyobject(py)?.into_any()),
            CalamineData::Float(v) => Ok(v.into_pyobject(py)?.into_any()),
            CalamineData::Str(v) => Ok(v.into_pyobject(py)?.into_any()),
            CalamineData::Bool(v) => Ok(v.into_pyobject(py)?.as_any().clone()),
            CalamineData::Date(v) => Ok(v.into_pyobject(py)?.as_any().clone()),
            CalamineData::DateTime(v) => Ok(v.into_pyobject(py)?.as_any().clone()),
            CalamineData::Empty => Ok(py.None().into_pyobject(py)?.into_any()),
        }
    }
}

macro_rules! impl_into_excel_data {
    ($($variant:ident),+) => {
        impl IntoExcelData for CalamineData {
            fn write(
                self,
                worksheet: &mut rust_xlsxwriter::Worksheet,
                row: rust_xlsxwriter::RowNum,
                col: rust_xlsxwriter::ColNum,
            ) -> Result<&mut rust_xlsxwriter::Worksheet, rust_xlsxwriter::XlsxError> {
                match self {
                    $(CalamineData::$variant(v) => v.write(worksheet, row, col),)+
                    CalamineData::Date(t) => {
                        let date_format = rust_xlsxwriter::Format::new().set_num_format("yyyy/mm/dd");
                        worksheet.write_datetime_with_format(row, col, &t, &date_format)
                    },
                    CalamineData::DateTime(t) => {
                        let date_format = rust_xlsxwriter::Format::new().set_num_format("yyyy/mm/dd hh:mm:ss");
                        worksheet.write_datetime_with_format(row, col, &t, &date_format)
                    },
                    CalamineData::Empty => Ok(worksheet),
                }
            }

            fn write_with_format<'a>(
                self,
                worksheet: &'a mut rust_xlsxwriter::Worksheet,
                row: rust_xlsxwriter::RowNum,
                col: rust_xlsxwriter::ColNum,
                format: &rust_xlsxwriter::Format,
            ) -> Result<&'a mut rust_xlsxwriter::Worksheet, rust_xlsxwriter::XlsxError> {
                match self {
                    $(CalamineData::$variant(v) => v.write_with_format(worksheet, row, col, format),)+
                    CalamineData::Date(t) => t.write_with_format(worksheet, row, col, format),
                    CalamineData::DateTime(t) => t.write_with_format(worksheet, row, col, format),
                    CalamineData::Empty => Ok(worksheet),
                }
            }
        }
    };
}
impl_into_excel_data!(Int, Float, Str, Bool);

impl<'py> FromPyObject<'py> for CalamineData {
    fn extract_bound(ob: &Bound<'py, PyAny>) -> PyResult<Self> {
        if ob.is_none() {
            return Ok(Self::Empty);
        } else if let Ok(bool_val) = ob.extract::<bool>() {
            return Ok(Self::Bool(bool_val));
        } else if let Ok(int_val) = ob.extract::<i64>() {
            return Ok(Self::Int(int_val));
        } else if let Ok(float_val) = ob.extract::<f64>() {
            return Ok(Self::Float(float_val));
        } else if let Ok(str_val) = ob.extract::<String>() {
            return Ok(Self::Str(str_val));
        } else if let Ok(datetime_val) = ob.extract::<NaiveDateTime>() {
            // ! Notice that datetime.datetime could also be extract as NaiveDate, so must check datetime first
            return Ok(Self::DateTime(datetime_val));
        } else if let Ok(date_val) = ob.extract::<NaiveDate>() {
            return Ok(Self::Date(date_val));
        }
        return Err(PyValueError::new_err("Invalid type"));
    }
}

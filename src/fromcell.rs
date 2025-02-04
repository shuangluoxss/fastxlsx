use crate::types::CalamineData;
use calamine::{Data, DataType};
use chrono::{NaiveDate, NaiveDateTime};
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
pub trait FromCell {
    fn from_cell(cell: Option<&Data>, strict: bool) -> PyResult<Self>
    where
        Self: Sized;
}

macro_rules! impl_from_cell {
    ($type:ty, $convert_method:ident, $default:expr, $type_name:expr) => {
        impl FromCell for $type {
            fn from_cell(cell: Option<&Data>, strict: bool) -> PyResult<Self> {
                let result = match cell {
                    Some(cell_data) => cell_data.$convert_method().ok_or_else(|| {
                        PyValueError::new_err(format!(
                            "Cell could not be parsed as {}: {:?}",
                            $type_name, cell_data
                        ))
                    }),
                    None => Err(PyValueError::new_err("Empty cell")),
                };
                if strict {
                    result
                } else {
                    Ok(result.unwrap_or($default))
                }
            }
        }
    };
}
impl_from_cell!(f64, as_f64, f64::NAN, "Float");
impl_from_cell!(i64, as_i64, i64::default(), "Int");
impl_from_cell!(String, as_string, String::default(), "String");
impl_from_cell!(bool, get_bool, bool::default(), "Bool");
impl_from_cell!(NaiveDate, as_date, NaiveDate::default(), "Date");
impl_from_cell!(
    NaiveDateTime,
    as_datetime,
    NaiveDateTime::default(),
    "DateTime"
);

impl FromCell for CalamineData {
    fn from_cell(cell: Option<&Data>, _strict: bool) -> PyResult<Self> {
        let value = match cell {
            Some(cell_data) => match cell_data {
                Data::Int(v) => CalamineData::Int(*v),
                Data::Float(v) => CalamineData::Float(*v),
                Data::String(v) => CalamineData::Str(v.clone()),
                Data::Bool(v) => CalamineData::Bool(*v),
                Data::DateTime(v) => CalamineData::DateTime(v.as_datetime().unwrap()),
                Data::DateTimeIso(v) => CalamineData::Str(v.clone()),
                Data::DurationIso(v) => CalamineData::Str(v.clone()),
                Data::Error(_) => CalamineData::Str("error".into()),
                Data::Empty => CalamineData::Empty,
            },
            None => CalamineData::Empty,
        };
        Ok(value)
    }
}

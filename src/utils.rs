pub fn adjust_idx(i: i32, n: usize) -> usize {
    if i < 0 {
        return (n as i32 + i) as usize;
    } else {
        return i as usize;
    }
}

#[macro_export]
macro_rules! define_extract_macro {
    ($(($type:ty, $type_extract:ty, $variant:ident)),* , $array_variant:ident, $macro_name:ident) => {
        #[macro_export]
        macro_rules! $macro_name{
            ($v:expr, $dtype:expr) =>{
                if let Some(dtype) = $dtype {
                    match dtype {
                        $(
                            DType::$variant => $v.extract::<$type_extract>()
                            .map(|arr| WrappedValue::$variant(ValueContainer::$array_variant(arr))),
                        )*
                    }
                }
                else{
                    loop {
                        $(
                            if let Ok(value) = $v.extract::<$type_extract>(){
                                break Ok(WrappedValue::$variant(ValueContainer::$array_variant(value)))
                            }
                        )*
                        break Err(PyTypeError::new_err("Unsupported array type"))
                    }
                }
            }
        }

    };
}
define_extract_macro!(
    (bool, bool, Bool),
    (i64, i64, Int),
    (f64, f64, Float),
    (String, String, Str),
    // ! DateTime must be in front of Date
    (NaiveDateTime, NaiveDateTime, DateTime),
    (NaiveDate, NaiveDate, Date),
    (Any, CalamineData, Any),
    Scalar,
    extract_scalar
);

define_extract_macro!(
    (bool, Array1Container<bool>, Bool),
    (i64, Array1Container<i64>, Int),
    (f64, Array1Container<f64>, Float),
    (String, Array1Container<String>, Str),
    // ! DateTime must be in front of Date
    (NaiveDateTime, Array1Container<NaiveDateTime>, DateTime),
    (NaiveDate, Array1Container<NaiveDate>, Date),
    (Any, Array1Container<CalamineData>, Any),
    Array1,
    extract_array1
);

define_extract_macro!(
    (bool, Array2Container<bool>, Bool),
    (i64, Array2Container<i64>, Int),
    (f64, Array2Container<f64>, Float),
    (String, Array2Container<String>, Str),
    // ! DateTime must be in front of Date
    (NaiveDateTime, Array2Container<NaiveDateTime>, DateTime),
    (NaiveDate, Array2Container<NaiveDate>, Date),
    (Any, Array2Container<CalamineData>, Any),
    Array2,
    extract_array2
);

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;

#[derive(Clone, FromPyObject, IntoPyObject)]
pub enum CellAddr {
    Idx((usize, usize)),
    Name(String),
}
impl CellAddr {
    pub fn as_idx(&self) -> PyResult<(usize, usize)> {
        match self {
            CellAddr::Idx(idx) => Ok(*idx),
            CellAddr::Name(name) => {
                let (letters, number) =
                    name.split_at(name.chars().position(|c| c.is_digit(10)).ok_or(
                        PyValueError::new_err(format!("Invalid cell address: {name}")),
                    )?);
                let mut col: usize = 0;
                let _ = letters.chars().try_for_each(|c| {
                    let x = c as usize - 'A' as usize + 1;
                    if x > 26_usize {
                        return Err(PyValueError::new_err(format!(
                            "Invalid cell address: {name}"
                        )));
                    } else {
                        col = col * 26 + x
                    }
                    Ok(())
                })?;
                col -= 1;

                let row = number
                    .parse::<usize>()
                    .map_err(|_| PyValueError::new_err(format!("Invalid cell address: {name}")))?
                    - 1;

                Ok((row, col))
            }
        }
    }
    pub fn as_addr(&self) -> PyResult<String> {
        match self {
            CellAddr::Idx((row, col)) => {
                let mut letters = String::new();
                let mut col = *col;
                while col >= 26 {
                    letters.push((col % 26 + 'A' as usize) as u8 as char);
                    col /= 26;
                }
                letters.push((col + 'A' as usize) as u8 as char);
                Ok(format!("{}{}", letters, row + 1))
            }
            CellAddr::Name(name) => Ok(name.clone()),
        }
    }
}

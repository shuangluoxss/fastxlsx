pub mod conversations;
pub mod fromcell;
pub mod read;
pub mod types;
#[macro_use]
pub mod utils;
pub mod write;
use pyo3::prelude::*;

/**
    Convert a cell address string (e.g., "A1") to a 0-based (row, col) index.

    Parameters
    ----------
    addr : str
        The cell address string (e.g., "A1").

    Returns
    -------
    Tuple[int, int]
        A tuple of (row, col) indices.
*/
#[pyfunction]
fn addr_to_idx(addr: String) -> PyResult<(usize, usize)> {
    types::CellAddr::Name(addr).as_idx()
}

/**
    Convert a 0-based (row, col) index to a cell address string (e.g., "A1").

    Parameters
    ----------
    row : int
        The 0-based row index.
    col : int
        The 0-based column index.

    Returns
    -------
    str
        The cell address string (e.g., "A1").
*/
#[pyfunction]
fn idx_to_addr(row: usize, col: usize) -> PyResult<String> {
    types::CellAddr::Idx((row, col)).as_addr()
}
/// Returns the current version of the library.
#[pyfunction]
fn version() -> PyResult<String> {
    Ok(env!("CARGO_PKG_VERSION").to_string())
}

/// A fast library for reading and writing XLSX files.
#[pymodule]
fn fastxlsx(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<types::DType>()?;
    m.add_class::<types::DShape>()?;
    m.add_class::<types::RangeInfo>()?;
    m.add_class::<read::ReadOnlyWorkbook>()?;
    m.add_class::<read::ReadOnlyWorksheet>()?;
    m.add_class::<write::WriteOnlyWorkbook>()?;
    m.add_class::<write::WriteOnlyWorksheet>()?;
    m.add_function(wrap_pyfunction!(read::read_many, m)?)?;
    m.add_function(wrap_pyfunction!(write::write_many, m)?)?;
    m.add_function(wrap_pyfunction!(version, m)?)?;
    m.add_function(wrap_pyfunction!(addr_to_idx, m)?)?;
    m.add_function(wrap_pyfunction!(idx_to_addr, m)?)?;
    Ok(())
}

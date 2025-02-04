FastXLSX
======================

.. toctree::
   :maxdepth: 2
   :hidden:
   :caption: CONTENT

   getting_started
   advanced_usage
   api

A lightweight, high-performance Python library for blazing-fast XLSX I/O operations.

Powered by Rust's `Calamine <https://github.com/tafia/calamine>`_ (reading) and `Rust-XlsxWriter <https://github.com/jmcnamara/rust_xlsxwriter>`_ (writing), with seamless Python integration via `PyO3 <https://github.com/PyO3/pyo3>`_.

Key Features
------------
Supported Capabilities
^^^^^^^^^^^^^^^^^^^^^^

- **Data Types**: Native support for `bool`, `int`, `float`, `date`, `datetime`, and `str`.
- **Data Operations**: Scalars, rows, columns, matrices, and batch processing.
- **Coordinate Systems**: Dual support for **A1** (e.g., `B2`) and **R1C1** (e.g., `(2, 3)`) notation.
- **Parallel Processing**: Multi-threaded read/write operations for massive datasets.
- **Type Safety**: Full type hints and IDE-friendly documentation.
- **Blasting Performance**: 5-10x faster compared to `openpyxl`.

Current Limitations
^^^^^^^^^^^^^^^^^^^

- **File Formats**: Only XLSX (no XLS/XLSB support).
- **Formulas & Styling**: Cell formulas, merged cells, and formatting not supported.
- **Modifications**: Append/update operations on existing files unavailable.
- **Advanced Features**: Charts, images, and other advanced features not supported.

Installation
------------

PyPI Install
^^^^^^^^^^^^

.. code-block:: bash

    pip install fastxlsx

Source Build (Requires Rust Toolchain)
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


.. code-block:: bash

    git clone https://github.com/shuangluoxss/fastxlsx.git
    cd fastxlsx
    pip install .
Advanced Usage
==============

While FastXLSX offers a minimal API for simplicity, these advanced features require special attention for optimal performance and data integrity.


Type Handling
-------------

Core Concepts
~~~~~~~~~~~~~
FastXLSX uses the ``DType`` enum to enforce type consistency. Explicit type declaration provides:

- **Speed boost** in I/O operations  
- **Memory efficiency** through optimized data structures  
- **Type safety** for critical data pipelines  

Implementation Patterns
~~~~~~~~~~~~~~~~~~~~~~~
- **Reading**: Specify types via ``RangeInfo`` objects or directly in ``cell_value`` method
- **Writing**: Declare types directly in ``write_xxx()`` methods

.. hint::
   Always specify ``dtype`` when possible - auto-detection add ~0.3μs/cell.

Numerical Optimization
----------------------

Return Type Behavior
~~~~~~~~~~~~~~~~~~~~
FastXLSX employs specialized memory layouts for numerical data:

+-----------------------+-------------------+
|       Data Type       | Return Structure  |
+=======================+===================+
| Bool/Int/Float        | ``numpy.ndarray`` |
+-----------------------+-------------------+
| Str/Date/DateTime/Any | Python ``list``   |
+-----------------------+-------------------+

Code Examples
~~~~~~~~~~~~~

.. code-block:: python

   # Reading with/without explicit typing
   from fastxlsx import DType, RangeInfo
   
   # Auto-detected (slower, returns list)
   data_list = ws.read_value(RangeInfo((5,2), DShape.Row(3)))
   
   # Type-enforced (faster, returns ndarray)
   data_array = ws.read_value(
       RangeInfo((5,2), DShape.Row(3), dtype=DType.Float)
   )

.. code-block:: python
   :emphasize-lines: 7

   # Writing numerical data requirements
   import numpy as np
   
   # Valid: NumPy array with matching dtype
   ws.write_row((1,0), np.array([4,5,6]), dtype=DType.Int)
   
   # Invalid: Python list with numerical dtype
   ws.write_row((2,0), [7,8,9], dtype=DType.Int)  # Raises TypeError

Type Enforcement Modes
----------------------

Strict Mode (Default)
~~~~~~~~~~~~~~~~~~~~~
- **Behavior**: Raises ``ValueError`` on type mismatch
- **Use Case**: Data validation scenarios

.. code-block:: python

   # Strict type checking example
   try:
       ws.read_value(
           RangeInfo((0,0), DShape.Row(3), dtype=DType.Str)
       )
   except ValueError as e:
       print(f"Data integrity violation: {e}")

Lenient Mode (strict=False)
~~~~~~~~~~~~~~~~~~~~~~~~~~~
- **Behavior**: Replaces invalid values with type-specific defaults
- **Warning**: May cause silent data corruption

+---------------+-------------------+
|     DType     |   Default Value   |
+===============+===================+
| Bool          | False             |
+---------------+-------------------+
| Int           | 0                 |
+---------------+-------------------+
| Float         | nan               |
+---------------+-------------------+
| Date/DateTime | 1970-01-01        |
+---------------+-------------------+
| Str           | Empty string ("") |
+---------------+-------------------+
| Any           | None              |
+---------------+-------------------+

.. code-block:: python

   # Lenient mode example
   ws.read_value(
       RangeInfo((0,0), DShape.Row(3), dtype=DType.Date, strict=False)
   )
   # Returns: [date(1970,1,1), date(2025,2,3), date(1970,1,1)]

.. danger::
   Use lenient mode **ONLY IF YOU KNOW WHAT YOU ARE DOING**.

Parallel Processing
-------------------

FastXLSX implements Rust-native parallelism through the ``rayon`` library, offering true multi-threaded I/O operations without Python's GIL limitations.

Core Advantages
~~~~~~~~~~~~~~~
- **Cross-Platform Efficiency**: Bypasses Windows ``spawn`` method limitations
- **Scalability**: Linear throughput scaling with CPU cores
- **Memory Safety**: Zero-copy data pipelines with Rust's ownership model

Implementation Patterns
~~~~~~~~~~~~~~~~~~~~~~~
Batch Writing
""""""""""""""
Optimal for generating multiple files with identical schemas (same worksheet structure/columns):

.. code-block:: python
   :caption: Writing 10 files with 6 sheets each

   import numpy as np
   from fastxlsx import DType, write_many, WriteOnlyWorksheet
   
   workbooks = {}
   for fid in range(10):  # 10 output files
       sheets = []
       for sid in range(6):  # 6 sheets per file
           ws = WriteOnlyWorksheet(f"Sheet{sid}")
           # Header with workbook/sheet ID
           ws.write_cell("A1", 10*fid + sid, dtype=DType.Int)
           # 3x3 float matrix
           ws.write_matrix((1,1), np.random.rand(100,100), dtype=DType.Float)
           sheets.append(ws)
       workbooks[f"batch_{fid:02d}.xlsx"] = sheets
   
   write_many(workbooks)  # Parallelized write

Batch Reading
""""""""""""""
Ideal for aggregating data from multiple files with consistent layouts:

.. code-block:: python
   :caption: Reading 10 files with matrix extraction

   from fastxlsx import read_many, RangeInfo, DShape, DType
   
   results = read_many({
       f"batch_{fid:02d}.xlsx": {
           f"Sheet{sid}": [
               RangeInfo((0,0), DShape.Scalar(), dtype=DType.Int),  # Read header
               RangeInfo((1,1), DShape.Matrix(100,100), dtype=DType.Float) # Extract 3x3 matrix
           ]
           for sid in range(6)
       }
       for fid in range(10)
   })

Performance Characteristics
~~~~~~~~~~~~~~~~~~~~~~~~~~~
+------------------+---------------+----------------+
|    Operation     | 10 Files (ms) | 100 Files (ms) |
+==================+===============+================+
| Sequential Write | 584           | 5860           |
+------------------+---------------+----------------+
| Parallel Write   | 96.1 (6.07x)  | 924 (6.34x)    |
+------------------+---------------+----------------+
| Sequential Read  | 367           | 3730           |
+------------------+---------------+----------------+
| Parallel Read    | 65.9 (5.57x)  | 620 (6.01x)    |
+------------------+---------------+----------------+

.. note::
   Benchmark environment: AMD Ryzen 7 5600X (6-core), 64GB DDR4, SATA SSD, Win10

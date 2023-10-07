namespace Exceleration
{
    /// <summary>
    /// Represents a cell within a worksheet.
    /// </summary>
    public class Cell
    {
        private readonly int _rowIndex;
        private readonly int _colIndex;

        /// <summary>
        /// Gets the value stored in the cell.
        /// </summary>
        public object Value { get; private set; }

        /// <summary>
        /// Gets the address of the cell in A1-style notation.
        /// </summary>
        public string Address { get; private set; }

        /// <summary>
        /// Gets the parent worksheet to which this cell belongs.
        /// </summary>
        public Worksheet Parent { get; private set; }

        /// <summary>
        /// Gets the data type of the cell's value.
        /// </summary>
        public Type DataType { get; private set; }

        /// <summary>
        /// Gets the row number (1-based) of the cell.
        /// </summary>
        public int Row
        {
            get
            {
                return _rowIndex + 1;
            }
        }

        /// <summary>
        /// Gets the column number (1-based) of the cell.
        /// </summary>
        public int Column
        {
            get
            {
                return _colIndex + 1;
            }
        }

        /// <summary>
        /// Gets the column letter (e.g., "A") corresponding to the cell's column index.
        /// </summary>
        public string ColumnLetter
        {
            get
            {
                return ConvertNumberToColLetter(_colIndex + 1);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// </summary>
        /// <param name="value">The value stored in the cell.</param>
        /// <param name="address">The address of the cell in A1-style notation.</param>
        /// <param name="rowIndex">The row index of the cell (0-based).</param>
        /// <param name="colIndex">The column index of the cell (0-based).</param>
        /// <param name="parentWorksheet">The parent worksheet to which this cell belongs.</param>
        /// <param name="dataType">The data type of the cell's value.</param>
        public Cell(object value, string address, int rowIndex, int colIndex, Worksheet parentWorksheet, Type dataType)
        {            
            _rowIndex = rowIndex;
            _colIndex = colIndex;
            Value = value;
            Address = address;
            Parent = parentWorksheet;
            DataType = dataType;
        }

        /// <summary>
        /// Gets a cell that is offset from the current cell by the specified number of rows and columns.
        /// </summary>
        /// <param name="rowOffset">The number of rows to offset (positive or negative).</param>
        /// <param name="colOffset">The number of columns to offset (positive or negative).</param>
        /// <returns>The cell that is offset from the current cell.</returns>
        public Cell Offset(int rowOffset, int colOffset)
        {
            int newRow = _rowIndex + rowOffset;
            int newCol = _colIndex + colOffset;

            return Parent.GetCell(newRow, newCol);
        }

        /// <summary>
        /// Returns a string representation of the cell's value.
        /// </summary>
        /// <returns>A string representation of the cell's value.</returns>
        public override string? ToString()
        {
            return Value.ToString();
        }

        /// <summary>
        /// Converts the cell's value to the specified type and returns the result. Constrained to type struct.
        /// </summary>
        /// <typeparam name="T">The target type to which the value should be converted.</typeparam>
        /// <param name="returnDefaultOnConversionError">If true, returns the default value of the target type on conversion error. If false, throws an exception on error.</param>
        /// <returns>The converted value of the cell, or the default value of the target type on conversion error (if returnDefaultOnConversionError is true).</returns>
        /// <exception cref="InvalidCastException">Thrown if the value cannot be converted to the specified type and returnDefaultOnConversionError is false.</exception>
        public T To<T>(bool returnDefaultOnConversionError = true) where T : struct
        {
            try
            {
                return (T)Convert.ChangeType(Value, typeof(T));
            }
            catch
            {
                if (returnDefaultOnConversionError) return default;
                else throw new InvalidCastException($"Cannot convert cell value to type {typeof(T)}.");
            }
        }

        /// <summary>
        /// Converts the cell's value to a nullable value of the specified type and returns the result. Constrained to type struct.
        /// </summary>
        /// <typeparam name="T">The target nullable value type to which the value should be converted.</typeparam>
        /// <param name="returnNullOnConversionError">If true, returns null on conversion error. If false, returns the default nullable value of the target type on error.</param>
        /// <returns>The converted nullable value of the cell, or null on conversion error (if returnNullOnConversionError is true).</returns>
        public T? ToNullable<T>(bool returnNullOnConversionError = true) where T : struct
        {
            if (string.IsNullOrEmpty(Value?.ToString()?.Trim()))
            {
                if (returnNullOnConversionError) return null;
                else return default;
            }

            try
            {
                return (T)Convert.ChangeType(Value, typeof(T));
            }

            catch
            {
                if (returnNullOnConversionError) return null;
                else return default;
            }
        }

        /// <summary>
        /// Determines whether the cell's value can be successfully parsed into a value of the specified type. Constrained to type struct.
        /// </summary>
        /// <typeparam name="T">The target type to check for parseability.</typeparam>
        /// <returns>True if the value can be parsed into the specified type; otherwise, false.</returns>
        public bool IsParseable<T>() where T : struct
        {
            if (string.IsNullOrEmpty(Value?.ToString()?.Trim())) return false;

            try
            {
                Convert.ChangeType(Value, typeof(T));
                return true;
            }

            catch
            {
                return false;
            }
        }

        private static string ConvertNumberToColLetter(int colNumber)
        {
            int dividend = colNumber;
            string colLetter = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colLetter = Convert.ToChar(65 + modulo).ToString() + colLetter;
                dividend = ((dividend - modulo) / 26);
            }

            return colLetter;
        }
    }
}

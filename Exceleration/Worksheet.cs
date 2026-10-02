using System.Data;

namespace Exceleration
{
    /// <summary>
    /// Represents a worksheet within a workbook.
    /// </summary>
    public class Worksheet
    {
        /// <summary>
        /// Gets the internal DataTable associated with the worksheet.
        /// </summary>
        internal DataTable DataTable { get; set; }

        /// <summary>
        /// Gets the parent workbook to which this worksheet belongs.
        /// </summary>
        public Workbook Parent { get; private set; }

        /// <summary>
        /// Gets the name of the worksheet.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Worksheet"/> class.
        /// </summary>
        /// <param name="table">The DataTable representing the worksheet data.</param>
        /// <param name="parent">The parent workbook to which this worksheet belongs.</param>
        internal Worksheet(DataTable table, Workbook parent)
        {
            DataTable = table;
            Parent = parent;
            Name = table.TableName;
        }

        /// <summary>
        /// Gets a list of all cells in the worksheet.
        /// </summary>
        public List<Cell> Cells
        {
            get
            {
                var allCells = new List<Cell>();

                for (int rowIndex = 0; rowIndex < DataTable.Rows.Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < DataTable.Columns.Count; colIndex++)
                    {
                        allCells.Add(GetCell(rowIndex + 1, colIndex + 1));
                    }
                }

                return allCells;
            }
        }

        /// <summary>
        /// Gets a list of rows in the worksheet, where each row is represented as a list of cells.
        /// </summary>
        public List<List<Cell>> Rows
        {
            get
            {
                var rows = new List<List<Cell>>();

                for (int rowIndex = 0; rowIndex < DataTable.Rows.Count; rowIndex++)
                {
                    var rowCells = new List<Cell>();
                    for (int colIndex = 0; colIndex < DataTable.Columns.Count; colIndex++)
                    {
                        rowCells.Add(GetCell(rowIndex + 1, colIndex + 1));
                    }
                    rows.Add(rowCells);
                }

                return rows;
            }
        }

        /// <summary>
        /// Gets a list of columns in the worksheet, where each column is represented as a list of cells.
        /// </summary>
        public List<List<Cell>> Columns
        {
            get
            {
                var columns = new List<List<Cell>>();

                for (int colIndex = 0; colIndex < DataTable.Columns.Count; colIndex++)
                {
                    var colCells = new List<Cell>();
                    for (int rowIndex = 0; rowIndex < DataTable.Rows.Count; rowIndex++)
                    {
                        colCells.Add(GetCell(rowIndex + 1, colIndex + 1));
                    }
                    columns.Add(colCells);
                }

                return columns;
            }
        }

        /// <summary>
        /// Gets the cell at the specified address in A1-style notation.
        /// </summary>
        /// <param name="cellAddress">The address of the cell in A1-style notation: column letters, in either case, then the row number, such as "B12".</param>
        /// <returns>The cell at the specified address.</returns>
        /// <exception cref="ArgumentNullException">Thrown if the address is null.</exception>
        /// <exception cref="ArgumentException">Thrown if the address is not an A1-style reference.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the cell is outside the sheet's used range.</exception>
        public Cell this[string cellAddress]
        {
            get
            {
                (int rowNumber, int colNumber) = ParseA1Reference(cellAddress, nameof(cellAddress));

                return GetCell(rowNumber, colNumber);
            }
        }

        /// <summary>
        /// Gets the cell at the specified row and column indices.
        /// </summary>
        /// <param name="rowNumber">The row index (1-based).</param>
        /// <param name="colNumber">The column index (1-based).</param>
        /// <returns>The cell at the specified row and column.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the row or column index is out of range.</exception>
        public Cell GetCell(int rowNumber, int colNumber)
        {
            CheckRow(rowNumber);
            CheckColumn(colNumber);

            int rowIndex = rowNumber - 1;
            int colIndex = colNumber - 1;
            object value = DataTable.Rows[rowIndex][colIndex];
            string address = ConvertToA1Style(rowIndex, colIndex);
            Type dataType = value.GetType();

            return new Cell(value, address, rowIndex, colIndex, this, dataType);
        }

        /// <summary>
        /// Gets the cell at the specified A1-style reference.
        /// </summary>
        /// <param name="a1Reference">The A1-style reference of the cell: column letters, in either case, then the row number, such as "B12".</param>
        /// <returns>The cell at the specified A1-style reference.</returns>
        /// <exception cref="ArgumentNullException">Thrown if the reference is null.</exception>
        /// <exception cref="ArgumentException">Thrown if the reference is not an A1-style reference.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the cell is outside the sheet's used range.</exception>
        public Cell GetCell(string a1Reference)
        {
            var (row, col) = ParseA1Reference(a1Reference, nameof(a1Reference));
            return GetCell(row, col);
        }

        /// <summary>
        /// Gets the value of the cell at the specified row and column indices.
        /// </summary>
        /// <param name="rowNumber">The row index (1-based).</param>
        /// <param name="colNumber">The column index (1-based).</param>
        /// <returns>The value of the cell at the specified row and column.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the row or column index is out of range.</exception>
        public object GetCellValue(int rowNumber, int colNumber)
        {
            CheckRow(rowNumber);
            CheckColumn(colNumber);

            return DataTable.Rows[rowNumber - 1][colNumber - 1];
        }

        /// <summary>
        /// Gets the value of the cell at the specified A1-style cell address.
        /// </summary>
        /// <param name="cellAddress">The A1-style cell address: column letters, in either case, then the row number, such as "B12".</param>
        /// <returns>The value of the cell at the specified cell address.</returns>
        /// <exception cref="ArgumentNullException">Thrown if the cell address is null.</exception>
        /// <exception cref="ArgumentException">Thrown if the cell address is not an A1-style reference.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the cell is outside the sheet's used range.</exception>
        public object GetCellValue(string cellAddress)
        {
            var (rowNumber, colNumber) = ParseA1Reference(cellAddress, nameof(cellAddress));

            return GetCellValue(rowNumber, colNumber);
        }

        /// <summary>
        /// Gets a list of cells in the specified row.
        /// </summary>
        /// <param name="rowNumber">The row index (1-based).</param>
        /// <returns>A list of cells in the specified row.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the row index is out of range.</exception>
        public List<Cell> GetRow(int rowNumber)
        {
            CheckRow(rowNumber);

            var rowCells = new List<Cell>();

            for (int colIndex = 0; colIndex < DataTable.Columns.Count; colIndex++)
            {
                rowCells.Add(GetCell(rowNumber, colIndex + 1));
            }

            return rowCells;
        }

        /// <summary>
        /// Gets a list of cells in the specified column by its column letter (e.g., "A").
        /// </summary>
        /// <param name="colLetter">The column letters, in either case (e.g., "A" or "ab").</param>
        /// <returns>A list of cells in the specified column.</returns>
        /// <exception cref="ArgumentNullException">Thrown if the column letters are null.</exception>
        /// <exception cref="ArgumentException">Thrown if the column letters are empty or hold anything but the letters A to Z.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the column is outside the sheet's used range.</exception>
        public List<Cell> GetColumn(string colLetter)
        {
            ArgumentNullException.ThrowIfNull(colLetter);
            if (colLetter.Length == 0 || !colLetter.All(char.IsAsciiLetter))
            {
                throw new ArgumentException($"'{colLetter}' is not a column such as \"B\" or \"AB\".", nameof(colLetter));
            }

            return GetColumn(ColumnLettersToNumber(colLetter));
        }

        /// <summary>
        /// Gets a list of cells in the specified column by its column index (1-based).
        /// </summary>
        /// <param name="colNumber">The column index (1-based).</param>
        /// <returns>A list of cells in the specified column.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the column index is out of range.</exception>
        public List<Cell> GetColumn(int colNumber)
        {
            CheckColumn(colNumber);

            var columnCells = new List<Cell>();

            for (int i = 1; i <= DataTable.Rows.Count; i++)
            {
                columnCells.Add(GetCell(i, colNumber));
            }

            return columnCells;
        }

        /// <summary>
        /// Returns a copy of the worksheet's data as a DataTable. Its columns are named Column0, Column1 and so on,
        /// and row 1 of the sheet is the table's first row.
        /// </summary>
        /// <returns>A copy of the DataTable representing the worksheet data.</returns>
        public DataTable ToDataTable()
        {
            return DataTable.Copy();
        }

        private void CheckRow(int rowNumber)
        {
            if (rowNumber < 1 || rowNumber > DataTable.Rows.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(rowNumber), rowNumber,
                    $"Row {rowNumber} is outside the sheet '{Name}', which has {DataTable.Rows.Count} rows.");
            }
        }

        private void CheckColumn(int colNumber)
        {
            if (colNumber < 1 || colNumber > DataTable.Columns.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(colNumber), colNumber,
                    $"Column {colNumber} is outside the sheet '{Name}', which has {DataTable.Columns.Count} columns.");
            }
        }

        // Column letters then row digits, nothing else: "B12" or "b12". A row or column too large
        // for an int becomes int.MaxValue, which no sheet has, so the lookup reports it as out of range.
        private static (int row, int col) ParseA1Reference(string reference, string paramName)
        {
            ArgumentNullException.ThrowIfNull(reference, paramName);

            int letters = 0;
            while (letters < reference.Length && char.IsAsciiLetter(reference[letters]))
            {
                letters++;
            }

            int end = letters;
            while (end < reference.Length && char.IsAsciiDigit(reference[end]))
            {
                end++;
            }

            if (letters == 0 || end == letters || end != reference.Length)
            {
                throw new ArgumentException($"'{reference}' is not an A1-style reference such as \"B12\".", paramName);
            }

            long row = 0;
            foreach (char digit in reference.AsSpan(letters))
            {
                row = Math.Min(row * 10 + (digit - '0'), int.MaxValue);
            }

            return ((int)row, ColumnLettersToNumber(reference.AsSpan(0, letters)));
        }

        private static int ColumnLettersToNumber(ReadOnlySpan<char> letters)
        {
            long number = 0;
            foreach (char letter in letters)
            {
                number = Math.Min(number * 26 + (char.ToUpperInvariant(letter) - 'A' + 1), int.MaxValue);
            }

            return (int)number;
        }

        private static string ConvertToA1Style(int row, int col)
        {
            string columnPart = IndexToColumn(col + 1);
            int rowPart = row + 1;
            return $"{columnPart}{rowPart}";
        }

        private static string IndexToColumn(int index)
        {
            string columnName = "";
            while (index > 0)
            {
                int modulo = (index - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                index = (index - modulo) / 26;
            }
            return columnName;
        }

    }
}

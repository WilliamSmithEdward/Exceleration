using System.Data;
using System.Text.RegularExpressions;

namespace Exceleration
{
    /// <summary>
    /// Represents a worksheet within a workbook.
    /// </summary>
    public partial class Worksheet
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
                        allCells.Add(GetCell(rowIndex, colIndex));
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
        /// <param name="cellAddress">The address of the cell in A1-style notation.</param>
        /// <returns>The cell at the specified address.</returns>
        public Cell this[string cellAddress]
        {
            get
            {
                (int rowIndex, int colIndex) = ConvertFromA1Style(cellAddress);

                return GetCell(rowIndex, colIndex);
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
            int rowIndex = rowNumber - 1;
            int colIndex = colNumber - 1;

            if (rowIndex < 0 || rowIndex >= DataTable.Rows.Count ||
                colIndex < 0 || colIndex >= DataTable.Columns.Count)
            {
                throw new ArgumentOutOfRangeException($"Invalid row { rowNumber } or column { colNumber } index.");
            }

            object value = DataTable.Rows[rowIndex][colIndex];
            string address = ConvertToA1Style(rowIndex, colIndex);
            Type dataType = value.GetType();

            return new Cell(value, address, rowIndex, colIndex, this, dataType);
        }

        /// <summary>
        /// Gets the cell at the specified A1-style reference.
        /// </summary>
        /// <param name="a1Reference">The A1-style reference of the cell.</param>
        /// <returns>The cell at the specified A1-style reference.</returns>
        public Cell GetCell(string a1Reference)
        {
            var (row, col) = ConvertFromA1Style(a1Reference);
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
            int rowIndex = rowNumber - 1;
            int colIndex = colNumber - 1;

            if (rowIndex < 0 || rowIndex >= DataTable.Rows.Count ||
                colIndex < 0 || colIndex >= DataTable.Columns.Count)
            {
                throw new ArgumentOutOfRangeException($"Invalid row { rowNumber } or column { colNumber } index.");
            }

            return DataTable.Rows[rowIndex][colIndex];
        }

        /// <summary>
        /// Gets the value of the cell at the specified A1-style cell address.
        /// </summary>
        /// <param name="cellAddress">The A1-style cell address (e.g., "A1").</param>
        /// <returns>The value of the cell at the specified cell address.</returns>
        /// <exception cref="ArgumentException">Thrown if the cell address is invalid.</exception>
        public object GetCellValue(string cellAddress)
        {
            int colNumber = 0;
            int rowNumber = 0;
            int multiplier = 1;

            for (int i = cellAddress.Length - 1; i >= 0; i--)
            {
                char ch = cellAddress[i];
                if (char.IsLetter(ch))
                {
                    colNumber += (ch - 'A' + 1) * multiplier;
                    multiplier *= 26;
                }
                else if (char.IsDigit(ch))
                {
                    rowNumber = rowNumber * 10 + (ch - '0');
                }
                else
                {
                    throw new ArgumentException("Invalid cell address.");
                }
            }

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
        /// <param name="colLetter">The column letter (e.g., "A").</param>
        /// <returns>A list of cells in the specified column.</returns>
        /// <exception cref="ArgumentException">Thrown if the column letter is invalid.</exception>
        public List<Cell> GetColumn(string colLetter)
        {
            int colNumber = ConvertColLetterToColNumber(colLetter);

            return GetColumn(colNumber);
        }

        /// <summary>
        /// Gets a list of cells in the specified column by its column index (1-based).
        /// </summary>
        /// <param name="colNumber">The column index (1-based).</param>
        /// <returns>A list of cells in the specified column.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the column index is out of range.</exception>
        public List<Cell> GetColumn(int colNumber)
        {
            var columnCells = new List<Cell>();

            for (int i = 1; i <= DataTable.Rows.Count; i++)
            {
                columnCells.Add(GetCell(i, colNumber));
            }

            return columnCells;
        }

        /// <summary>
        /// Converts the worksheet data to a DataTable and returns it.
        /// </summary>
        /// <returns>A copy of the DataTable representing the worksheet data.</returns>
        public DataTable ToDataTable()
        {
            return DataTable.Copy();
        }

        private (int row, int col) ConvertFromA1Style(string a1Reference)
        {
            var match = MyRegex().Match(a1Reference);
            if (!match.Success)
            {
                throw new ArgumentException("Invalid A1 style reference.");
            }

            string columnPart = match.Groups[1].Value;
            string rowPart = match.Groups[2].Value;

            int row = int.Parse(rowPart);
            int col = ColumnToIndex(columnPart);

            return (row, col);
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

        private int ColumnToIndex(string column)
        {
            int index = 0;
            foreach (char ch in column)
            {
                index *= 26;
                index += ch - 'A' + 1;
            }
            return index;
        }

        private int ConvertColLetterToColNumber(string colLetter)
        {
            int colNumber = 0;
            for (int i = 0; i < colLetter.Length; i++)
            {
                colNumber = colNumber * 26 + colLetter[i] - 'A' + 1;
            }

            return colNumber;
        }

        [GeneratedRegex("([A-Za-z]+)(\\d+)")]
        private static partial Regex MyRegex();
    }
}

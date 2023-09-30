using System.Data;
using System.Text.RegularExpressions;

namespace Exceleration
{
    public class Worksheet
    {
        internal DataTable DataTable { get; set; }
        public Workbook Parent { get; private set; }
        public string Name { get; private set; }
        
        internal Worksheet(DataTable table, Workbook parent)
        {
            DataTable = table;
            Parent = parent;
            Name = table.TableName;
        }

        public List<Cell> Cells
        {
            get
            {
                List<Cell> allCells = new List<Cell>();
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

        public Cell this[string cellAddress]
        {
            get
            {
                (int rowIndex, int colIndex) = ConvertFromA1Style(cellAddress);

                return GetCell(rowIndex, colIndex);
            }
        }

        public Cell GetCell(int rowIndex, int colIndex)
        {
            if (rowIndex < 0 || rowIndex >= DataTable.Rows.Count ||
                colIndex < 0 || colIndex >= DataTable.Columns.Count)
            {
                throw new ArgumentOutOfRangeException("Invalid row or column index.");
            }

            object value = DataTable.Rows[rowIndex][colIndex];
            string address = ConvertToA1Style(rowIndex, colIndex);
            Type dataType = value.GetType();

            return new Cell(value, address, rowIndex, colIndex, this, dataType);
        }

        public Cell GetCell(string a1Reference)
        {
            var (row, col) = ConvertFromA1Style(a1Reference);
            return GetCell(row, col);
        }

        public object GetCellValue(int rowIndex, int colIndex)
        {
            if (rowIndex < 0 || rowIndex >= DataTable.Rows.Count ||
                colIndex < 0 || colIndex >= DataTable.Columns.Count)
            {
                throw new ArgumentOutOfRangeException("Invalid row or column index.");
            }

            return DataTable.Rows[rowIndex][colIndex];
        }

        public object GetCellValue(string cellAddress)
        {
            int colIndex = 0;
            int rowIndex = 0;
            int multiplier = 1;
            for (int i = cellAddress.Length - 1; i >= 0; i--)
            {
                char ch = cellAddress[i];
                if (Char.IsLetter(ch))
                {
                    colIndex += (ch - 'A' + 1) * multiplier;
                    multiplier *= 26;
                }
                else if (Char.IsDigit(ch))
                {
                    rowIndex = rowIndex * 10 + (ch - '0');
                }
                else
                {
                    throw new ArgumentException("Invalid cell address.");
                }
            }

            rowIndex--;
            colIndex--;

            return GetCellValue(rowIndex, colIndex);
        }

        public List<Cell> Rows(int rowNumber)
        {
            int rowIndex = rowNumber--;

            List<Cell> rowCells = new List<Cell>();

            if (rowIndex >= 0 && rowIndex < DataTable.Rows.Count)
            {
                for (int colIndex = 0; colIndex < DataTable.Columns.Count; colIndex++)
                {
                    rowCells.Add(GetCell(rowIndex, colIndex));
                }
            }

            return rowCells;
        }

        public List<Cell> Columns(string colLetter)
        {
            int colIndex = ConvertColLetterToIndex(colLetter);

            return Columns(colIndex + 1);
        }

        public List<Cell> Columns(int colNumber)
        {
            List<Cell> columnCells = new List<Cell>();

            int colIndex = colNumber - 1;

            for (int i = 0; i < DataTable.Rows.Count; i++)
            {
                columnCells.Add(GetCell(i, colIndex));
            }

            return columnCells;
        }

        private (int row, int col) ConvertFromA1Style(string a1Reference)
        {
            var match = Regex.Match(a1Reference, @"([A-Za-z]+)(\d+)");
            if (!match.Success)
            {
                throw new ArgumentException("Invalid A1 style reference.");
            }

            string columnPart = match.Groups[1].Value;
            string rowPart = match.Groups[2].Value;

            int row = int.Parse(rowPart) - 1;
            int col = ColumnToIndex(columnPart) - 1;

            return (row, col);
        }

        private string ConvertToA1Style(int row, int col)
        {
            string columnPart = IndexToColumn(col + 1);
            int rowPart = row + 1;
            return $"{columnPart}{rowPart}";
        }

        private string IndexToColumn(int index)
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

        private int ConvertColLetterToIndex(string colLetter)
        {
            int colIndex = 0;
            for (int i = 0; i < colLetter.Length; i++)
            {
                colIndex = colIndex * 26 + (colLetter[i] - 'A' + 1);
            }
            colIndex--;

            return colIndex;
        }
    }
}

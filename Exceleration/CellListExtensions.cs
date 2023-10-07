namespace Exceleration
{
    /// <summary>
    /// Provides extension methods for working with lists of cells.
    /// </summary>
    public static class CellListExtensions
    {
        /// <summary>
        /// Gets the first cell in a list of cells with a specific row number.
        /// </summary>
        /// <param name="cells">The list of cells to search.</param>
        /// <param name="rowNumber">The row number to match.</param>
        /// <returns>The first cell with the specified row number.</returns>
        public static Cell GetFirstCellByRowNumber(this List<Cell> cells, int rowNumber)
        {
            return cells.First(x => x.Row == rowNumber);
        }

        /// <summary>
        /// Gets the first cell in a list of cells with a specific column letter (e.g., "A").
        /// </summary>
        /// <param name="cells">The list of cells to search.</param>
        /// <param name="columnLetter">The column letter to match.</param>
        /// <returns>The first cell with the specified column letter.</returns>
        public static Cell GetFirstCellByColumnLetter(this List<Cell> cells, string columnLetter)
        {
            return cells.First(x => x.ColumnLetter.Equals(columnLetter));
        }

        /// <summary>
        /// Gets the first cell in a list of cells with a specific column number.
        /// </summary>
        /// <param name="cells">The list of cells to search.</param>
        /// <param name="columnNumber">The column number to match.</param>
        /// <returns>The first cell with the specified column number.</returns>
        public static Cell GetFirstCellByColumnNumber(this List<Cell> cells, int columnNumber)
        {
            return cells.First(x => x.Column == columnNumber);
        }

        /// <summary>
        /// Gets a list of cells in a specific row from a list of cells.
        /// </summary>
        /// <param name="cells">The list of cells to search.</param>
        /// <param name="rowNumber">The row number to match.</param>
        /// <returns>A list of cells in the specified row.</returns>
        public static List<Cell> GetRow(this List<Cell> cells, int rowNumber)
        {
            return cells.Where(x => x.Row == rowNumber).ToList();
        }

        /// <summary>
        /// Gets a list of cells in a specific column by its column number (1-based) from a list of cells.
        /// </summary>
        /// <param name="cells">The list of cells to search.</param>
        /// <param name="columnNumber">The column number to match.</param>
        /// <returns>A list of cells in the specified column.</returns>
        public static List<Cell> GetColumn(this List<Cell> cells, int columnNumber)
        {
            return cells.Where(x => x.Column == columnNumber).ToList();
        }

        /// <summary>
        /// Gets a list of cells in a specific column by its column letter (e.g., "A") from a list of cells.
        /// </summary>
        /// <param name="cells">The list of cells to search.</param>
        /// <param name="columnLetter">The column letter to match.</param>
        /// <returns>A list of cells in the specified column.</returns>
        public static List<Cell> GetColumn(this List<Cell> cells, string columnLetter)
        {
            return cells.Where(x => x.ColumnLetter.Equals(columnLetter)).ToList();
        }
    }
}

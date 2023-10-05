namespace Exceleration
{
    public static class CellListExtensions
    {
        public static Cell GetFirstCellByRowNumber(this List<Cell> cells, int rowNumber)
        {
            return cells.First(x => x.Row == rowNumber);
        }

        public static Cell GetFirstCellByColumnNumber(this List<Cell> cells, int columnNumber)
        {
            return cells.First(x => x.Column == columnNumber);
        }

        public static Cell GetFirstCellByColumnLetter(this List<Cell> cells, string columnLetter)
        {
            return cells.First(x => x.ColumnLetter.Equals(columnLetter));
        }

        public static List<Cell> GetRow(this List<Cell> cells, int rowNumber)
        {
            return cells.Where(x => x.Row == rowNumber).ToList();
        }

        public static List<Cell> GetColumn(this List<Cell> cells, int columnNumber)
        {
            return cells.Where(x => x.Column == columnNumber).ToList();
        }

        public static List<Cell> GetColumn(this List<Cell> cells, string columnLetter)
        {
            return cells.Where(x => x.ColumnLetter.Equals(columnLetter)).ToList();
        }
    }
}

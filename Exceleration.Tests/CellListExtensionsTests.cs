namespace Exceleration.Tests;

public class CellListExtensionsTests
{
    private static List<Cell> AllCells(TempFolder folder) =>
        new Workbook(Xlsx.Prices(folder.File("prices.xlsx")))["Sheet1"].Rows.SelectMany(r => r).ToList();

    [Fact]
    public void GetFirstCell_finds_by_row_column_letter_and_column_number()
    {
        using var folder = new TempFolder();
        var cells = AllCells(folder);

        Assert.Equal("A2", cells.GetFirstCellByRowNumber(2).Address);
        Assert.Equal("C1", cells.GetFirstCellByColumnLetter("C").Address);
        Assert.Equal("B1", cells.GetFirstCellByColumnNumber(2).Address);
        Assert.Throws<InvalidOperationException>(() => cells.GetFirstCellByRowNumber(9));
    }

    [Fact]
    public void GetRow_and_GetColumn_filter_the_list()
    {
        using var folder = new TempFolder();
        var cells = AllCells(folder);

        Assert.Equal(new[] { "A3", "B3", "C3" }, cells.GetRow(3).Select(c => c.Address));
        Assert.Equal(new[] { "B1", "B2", "B3" }, cells.GetColumn("B").Select(c => c.Address));
        Assert.Equal(new[] { "C1", "C2", "C3" }, cells.GetColumn(3).Select(c => c.Address));
        Assert.Empty(cells.GetRow(9));
    }

    [Fact]
    public void The_README_filter_finds_the_expensive_fruit()
    {
        using var folder = new TempFolder();
        var ws = new Workbook(Xlsx.Prices(folder.File("prices.xlsx")))["Sheet1"];

        var expensive = ws.Rows
            .Skip(1)
            .Where(r => r.GetFirstCellByColumnLetter("C").To<double>() > 2)
            .Select(r => r.GetFirstCellByColumnNumber(1).Value);

        Assert.Equal(new object[] { "Pear" }, expensive);
    }
}

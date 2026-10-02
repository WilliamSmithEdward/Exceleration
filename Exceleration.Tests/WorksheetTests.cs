namespace Exceleration.Tests;

public class WorksheetTests
{
    private static Worksheet Prices(TempFolder folder) =>
        new Workbook(Xlsx.Prices(folder.File("prices.xlsx")))["Sheet1"];

    [Fact]
    public void Reads_a_cell_by_A1_reference_and_by_row_and_column()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Equal("Apple", ws["A2"].Value);
        Assert.Equal("Apple", ws.GetCell("A2").Value);
        var b3 = ws.GetCell(3, 2);
        Assert.Equal("B3", b3.Address);
        Assert.Equal(3, b3.Row);
        Assert.Equal(2, b3.Column);
        Assert.Equal("B", b3.ColumnLetter);
        Assert.Equal(7d, b3.Value);
        Assert.Same(ws, b3.Parent);
    }

    [Fact]
    public void GetCellValue_reads_by_reference_and_by_row_and_column()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Equal(2.25, ws.GetCellValue("C3"));
        Assert.Equal("Qty", ws.GetCellValue(1, 2));
    }

    [Fact]
    public void Row_one_is_row_one_and_leading_empty_rows_and_columns_are_kept()
    {
        using var folder = new TempFolder();
        var ws = new Workbook(Xlsx.Write(folder.File("offset.xlsx"), new SheetSpec("S", ("B3", "here"))))["S"];

        Assert.Equal("here", ws["B3"].Value);
        Assert.Equal(DBNull.Value, ws["A1"].Value);
        var table = ws.ToDataTable();
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(2, table.Columns.Count);
    }

    [Fact]
    public void Rows_and_Columns_list_every_cell_of_the_used_range()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Equal(new[] { "A1,B1,C1", "A2,B2,C2", "A3,B3,C3" },
            ws.Rows.Select(r => string.Join(",", r.Select(c => c.Address))));
        Assert.Equal(new[] { "A1,A2,A3", "B1,B2,B3", "C1,C2,C3" },
            ws.Columns.Select(c => string.Join(",", c.Select(x => x.Address))));
    }

    [Fact]
    public void GetRow_and_GetColumn_return_one_row_or_column()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Equal(new object[] { "Apple", 3d, 1.5 }, ws.GetRow(2).Select(c => c.Value));
        Assert.Equal(new object[] { "Qty", 3d, 7d }, ws.GetColumn("B").Select(c => c.Value));
        Assert.Equal(new object[] { "Price", 1.5, 2.25 }, ws.GetColumn(3).Select(c => c.Value));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(4, 1)]
    [InlineData(1, 4)]
    public void A_cell_outside_the_used_range_throws_ArgumentOutOfRangeException(int row, int column)
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Throws<ArgumentOutOfRangeException>(() => ws.GetCell(row, column));
        Assert.Throws<ArgumentOutOfRangeException>(() => ws.GetCellValue(row, column));
    }

    [Fact]
    public void A_reference_that_is_not_A1_style_throws_ArgumentException()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Throws<ArgumentException>(() => ws["$A$1"]);
        Assert.Throws<ArgumentException>(() => ws.GetCellValue("$A$1"));
    }

    [Fact]
    public void ToDataTable_returns_a_copy_with_numbered_columns()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        var table = ws.ToDataTable();
        table.Rows[1][0] = "Changed";

        Assert.Equal(new[] { "Column0", "Column1", "Column2" }, table.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName));
        Assert.Equal("Apple", ws["A2"].Value);
        Assert.Equal("Sheet1", table.TableName);
    }
}

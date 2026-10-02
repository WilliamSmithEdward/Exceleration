namespace Exceleration.Tests;

/// <summary>
/// Cell addressing: A1 references, row and column numbers, offsets and the cell lists.
/// </summary>
public class AddressingTests
{
    private static Worksheet Prices(TempFolder folder) =>
        new Workbook(Xlsx.Prices(folder.File("prices.xlsx")))["Sheet1"];

    /// <summary>A sheet with values in A1 to B12, so two-digit rows can be told apart.</summary>
    private static Worksheet Tall(TempFolder folder) =>
        new Workbook(Xlsx.Write(folder.File("tall.xlsx"), new SheetSpec("S",
            Enumerable.Range(1, 12).SelectMany(r => new (string, object)[] { ($"A{r}", $"A{r}"), ($"B{r}", $"B{r}") }).ToArray())))["S"];

    [Fact]
    public void Cells_lists_every_cell_row_by_row()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Equal(new[] { "A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3" }, ws.Cells.Select(c => c.Address));
        Assert.Equal("Apple", ws.Cells[3].Value);
    }

    [Theory]
    [InlineData("A1", 0, 0, "A1")]
    [InlineData("A1", 1, 0, "A2")]
    [InlineData("A1", 0, 1, "B1")]
    [InlineData("B2", 1, 1, "C3")]
    [InlineData("C3", -2, -1, "B1")]
    public void Offset_moves_by_the_given_rows_and_columns(string from, int rows, int columns, string to)
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        var cell = ws[from].Offset(rows, columns);

        Assert.Equal(to, cell.Address);
        Assert.Equal(ws[to].Value, cell.Value);
    }

    [Fact]
    public void Offset_outside_the_sheet_throws_ArgumentOutOfRangeException()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Throws<ArgumentOutOfRangeException>(() => ws["A1"].Offset(-1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => ws["C3"].Offset(0, 1));
    }

    [Theory]
    [InlineData("A12")]
    [InlineData("B10")]
    [InlineData("A1")]
    public void GetCellValue_reads_multi_digit_rows_in_order(string reference)
    {
        using var folder = new TempFolder();
        var ws = Tall(folder);

        Assert.Equal(reference, ws.GetCellValue(reference));
        Assert.Equal(reference, ws[reference].Value);
    }

    [Theory]
    [InlineData("a1", "A1")]
    [InlineData("b12", "B12")]
    [InlineData("b3", "B3")]
    public void Lower_case_references_read_the_same_cell(string reference, string expected)
    {
        using var folder = new TempFolder();
        var ws = Tall(folder);

        Assert.Equal(expected, ws[reference].Address);
        Assert.Equal(expected, ws.GetCell(reference).Address);
        Assert.Equal(expected, ws.GetCellValue(reference));
    }

    [Fact]
    public void Lower_case_column_letters_find_the_column()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);
        var cells = ws.Rows.SelectMany(r => r).ToList();

        Assert.Equal(new[] { "B1", "B2", "B3" }, ws.GetColumn("b").Select(c => c.Address));
        Assert.Equal(new[] { "B1", "B2", "B3" }, cells.GetColumn("b").Select(c => c.Address));
        Assert.Equal("C1", cells.GetFirstCellByColumnLetter("c").Address);
    }

    [Theory]
    [InlineData("A1B")]
    [InlineData("A1:B2")]
    [InlineData("?A1")]
    [InlineData(" A1")]
    [InlineData("1A")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("")]
    [InlineData("A-1")]
    public void A_reference_with_anything_but_letters_then_digits_throws_ArgumentException(string reference)
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        var indexer = Assert.ThrowsAny<ArgumentException>(() => ws[reference]);
        var getCell = Assert.ThrowsAny<ArgumentException>(() => ws.GetCell(reference));
        var getValue = Assert.ThrowsAny<ArgumentException>(() => ws.GetCellValue(reference));
        Assert.All(new[] { indexer, getCell, getValue }, e => Assert.IsType<ArgumentException>(e));
    }

    [Theory]
    [InlineData("A99999999999")]
    [InlineData("ZZZZZZZZZZ1")]
    [InlineData("A0")]
    [InlineData("D1")]
    [InlineData("A4")]
    public void A_reference_outside_the_sheet_throws_ArgumentOutOfRangeException(string reference)
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Throws<ArgumentOutOfRangeException>(() => ws[reference]);
        Assert.Throws<ArgumentOutOfRangeException>(() => ws.GetCellValue(reference));
    }

    [Theory]
    [InlineData("")]
    [InlineData("B1")]
    [InlineData("$B")]
    public void A_column_letter_that_is_not_letters_throws_ArgumentException(string letters)
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.IsType<ArgumentException>(Assert.ThrowsAny<ArgumentException>(() => ws.GetColumn(letters)));
    }

    [Fact]
    public void A_null_reference_throws_ArgumentNullException_naming_the_parameter()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        Assert.Equal("cellAddress", Assert.Throws<ArgumentNullException>(() => ws[null!]).ParamName);
        Assert.Equal("a1Reference", Assert.Throws<ArgumentNullException>(() => ws.GetCell(null!)).ParamName);
        Assert.Equal("cellAddress", Assert.Throws<ArgumentNullException>(() => ws.GetCellValue(null!)).ParamName);
        Assert.Equal("colLetter", Assert.Throws<ArgumentNullException>(() => ws.GetColumn((string)null!)).ParamName);
    }

    [Fact]
    public void An_index_out_of_range_names_the_parameter_and_says_why()
    {
        using var folder = new TempFolder();
        var ws = Prices(folder);

        var row = Assert.Throws<ArgumentOutOfRangeException>(() => ws.GetCell(4, 1));
        var column = Assert.Throws<ArgumentOutOfRangeException>(() => ws.GetCellValue(1, 0));

        Assert.Equal("rowNumber", row.ParamName);
        Assert.Equal(4, row.ActualValue);
        Assert.Contains("3 rows", row.Message);
        Assert.Equal("colNumber", column.ParamName);
        Assert.Equal(0, column.ActualValue);
        Assert.Contains("3 columns", column.Message);
    }

    [Fact]
    public void GetRow_and_GetColumn_check_the_index_on_a_sheet_with_no_cells()
    {
        using var folder = new TempFolder();
        var path = Xlsx.WriteRaw(folder.File("empty.xlsx"),
            ("Empty", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>"));
        var ws = new Workbook(path)["Empty"];

        Assert.Empty(ws.Cells);
        Assert.Throws<ArgumentOutOfRangeException>(() => ws.GetColumn(5));
        Assert.Throws<ArgumentOutOfRangeException>(() => ws.GetRow(5));
    }
}

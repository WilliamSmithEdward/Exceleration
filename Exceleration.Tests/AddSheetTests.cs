using System.Data;

namespace Exceleration.Tests;

/// <summary>
/// Adding sheets to a workbook: names, ownership and the caller's DataTable.
/// </summary>
public class AddSheetTests
{
    private static Workbook Prices(TempFolder folder, string name = "prices.xlsx") =>
        new(Xlsx.Prices(folder.File(name)));

    private static DataTable OneCell(string value)
    {
        var table = new DataTable("Original");
        table.Columns.Add("Column0", typeof(object));
        table.Rows.Add(value);
        return table;
    }

    [Theory]
    [InlineData("Sheet1")]
    [InlineData("sheet1")]
    [InlineData("SHEET1")]
    public void AddSheet_refuses_a_name_that_differs_only_in_case(string name)
    {
        using var folder = new TempFolder();
        var wb = Prices(folder);
        var other = new Workbook(Xlsx.Write(folder.File("other.xlsx"), new SheetSpec(name, ("A1", 1))));

        Assert.Throws<ArgumentException>(() => wb.AddSheet(OneCell("x"), name));
        Assert.Throws<ArgumentException>(() => wb.AddSheet(other[name]));
        Assert.Single(wb.Sheets);
    }

    [Fact]
    public void AddSheet_from_a_DataTable_leaves_the_callers_table_alone()
    {
        using var folder = new TempFolder();
        var wb = Prices(folder);
        var table = OneCell("before");

        wb.AddSheet(table, "Added");
        table.Rows[0][0] = "after";
        table.Rows.Add("extra");

        Assert.Equal("Original", table.TableName);
        Assert.Equal("before", wb["Added"]["A1"].Value);
        Assert.Throws<ArgumentOutOfRangeException>(() => wb["Added"]["A2"]);
    }

    [Fact]
    public void AddSheet_from_another_workbook_adds_a_sheet_this_workbook_owns()
    {
        using var folder = new TempFolder();
        var wb = Prices(folder);
        var other = new Workbook(Xlsx.Write(folder.File("other.xlsx"), new SheetSpec("Other", ("A1", "from other"))));
        var sheet = other["Other"];

        wb.AddSheet(sheet);

        Assert.Same(wb, wb["Other"].Parent);
        Assert.Equal("from other", wb["Other"]["A1"].Value);
        Assert.Same(wb, wb["Other"]["A1"].Parent.Parent);
        Assert.Same(other, sheet.Parent);
        Assert.Same(sheet, other["Other"]);
    }

    [Fact]
    public void AddSheet_with_null_throws_ArgumentNullException()
    {
        using var folder = new TempFolder();
        var wb = Prices(folder);

        Assert.Equal("sheet", Assert.Throws<ArgumentNullException>(() => wb.AddSheet(null!)).ParamName);
        Assert.Equal("table", Assert.Throws<ArgumentNullException>(() => wb.AddSheet(null!, "Name")).ParamName);
        Assert.Equal("workSheetName", Assert.Throws<ArgumentNullException>(() => wb.AddSheet(OneCell("x"), null!)).ParamName);
    }
}

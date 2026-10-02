namespace Exceleration.Tests;

public class CellTests
{
    private static Worksheet Values(TempFolder folder) =>
        new Workbook(Xlsx.Write(folder.File("values.xlsx"), new SheetSpec("S",
            ("A1", "text"), ("B1", 42), ("C1", true), ("D1", new DateTime(2024, 6, 27)), ("F1", "   "),
            ("A2", 1.5), ("B2", 2.5), ("C2", "12"))))["S"];

    [Fact]
    public void Values_are_strings_doubles_bools_DateTimes_or_DBNull()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        Assert.Equal("text", ws["A1"].Value);
        Assert.Equal(42d, ws["B1"].Value);
        Assert.Equal(true, ws["C1"].Value);
        Assert.Equal(new DateTime(2024, 6, 27), ws["D1"].Value);
        Assert.Equal(DBNull.Value, ws["E1"].Value);
        Assert.Equal(new[] { typeof(string), typeof(double), typeof(bool), typeof(DateTime), typeof(DBNull) },
            new[] { "A1", "B1", "C1", "D1", "E1" }.Select(r => ws[r].DataType));
    }

    [Fact]
    public void ToString_is_the_value_as_text_and_empty_for_an_empty_cell()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        Assert.Equal("text", ws["A1"].ToString());
        Assert.Equal("", ws["E1"].ToString());
    }

    [Fact]
    public void To_converts_and_rounds_to_even()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        Assert.Equal(42, ws["B1"].To<int>());
        Assert.Equal(1.5m, ws["A2"].To<decimal>());
        Assert.Equal(2, ws["A2"].To<int>());
        Assert.Equal(2, ws["B2"].To<int>());
        Assert.True(ws["C1"].To<bool>());
        Assert.Equal(new DateTime(2024, 6, 27), ws["D1"].To<DateTime>());
    }

    [Fact]
    public void To_returns_the_default_on_failure_or_throws_when_asked()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        Assert.Equal(0, ws["A1"].To<int>());
        Assert.Equal(0, ws["E1"].To<int>());
        Assert.Throws<InvalidCastException>(() => ws["A1"].To<int>(false));
    }

    [Fact]
    public void ToNullable_returns_null_for_empty_blank_or_unconvertible_cells()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        Assert.Equal(42, ws["B1"].ToNullable<int>());
        Assert.Null(ws["A1"].ToNullable<int>());
        Assert.Null(ws["E1"].ToNullable<int>());
        Assert.Null(ws["F1"].ToNullable<int>());
    }

    [Fact]
    public void IsParseable_says_whether_the_value_converts()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        Assert.True(ws["B1"].IsParseable<int>());
        Assert.True(ws["C2"].IsParseable<int>());
        Assert.False(ws["A1"].IsParseable<int>());
        Assert.False(ws["E1"].IsParseable<int>());
        Assert.False(ws["F1"].IsParseable<int>());
    }
}

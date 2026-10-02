namespace Exceleration.Tests;

/// <summary>
/// The returnNullOnConversionError flag of ToNullable.
/// </summary>
public class ConversionTests
{
    private static Worksheet Values(TempFolder folder) =>
        new Workbook(Xlsx.Write(folder.File("values.xlsx"), new SheetSpec("S",
            ("A1", "text"), ("B1", 42), ("C1", "   "), ("D1", 4.5), ("F1", 1))))["S"];

    [Fact]
    public void ToNullable_with_false_returns_the_type_default_when_the_value_does_not_convert()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        int? number = ws["A1"].ToNullable<int>(false);
        DateTime? date = ws["A1"].ToNullable<DateTime>(false);

        Assert.True(number.HasValue);
        Assert.Equal(0, number);
        Assert.Equal(default(DateTime), date);
    }

    [Fact]
    public void ToNullable_with_false_still_returns_null_for_an_empty_or_blank_cell()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        Assert.Null(ws["E1"].ToNullable<int>(false));
        Assert.Null(ws["C1"].ToNullable<int>(false));
    }

    [Fact]
    public void ToNullable_converts_a_value_that_converts_whatever_the_flag()
    {
        using var folder = new TempFolder();
        var ws = Values(folder);

        Assert.Equal(42, ws["B1"].ToNullable<int>(false));
        Assert.Equal(4.5, ws["D1"].ToNullable<double>(true));
    }
}

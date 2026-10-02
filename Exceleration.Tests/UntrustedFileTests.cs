namespace Exceleration.Tests;

/// <summary>
/// Workbooks from elsewhere: one exception type for a file that does not parse, and a limit on
/// the cells a sheet may make the library hold.
/// </summary>
public class UntrustedFileTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static TheoryData<string, string> BrokenWorksheets => new()
    {
        { "unclosed", $"<worksheet xmlns=\"{Main}\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v>" },
        { "not xml", "this is not XML" },
    };

    [Theory]
    [MemberData(nameof(BrokenWorksheets))]
    public void A_worksheet_that_does_not_parse_throws_InvalidDataException_naming_the_file(string label, string xml)
    {
        using var folder = new TempFolder();
        var path = Xlsx.WriteRaw(folder.File($"{label.Replace(' ', '-')}.xlsx"), ("Sheet1", xml));

        var e = Assert.Throws<InvalidDataException>(() => new Workbook(path));

        Assert.NotNull(e.InnerException);
        Assert.Contains($"'{Path.GetFileName(path)}'", e.Message);
        Assert.DoesNotContain(folder.Root, e.Message);
    }

    [Fact]
    public void A_zip_file_that_is_not_a_workbook_throws_InvalidDataException()
    {
        using var folder = new TempFolder();
        var path = folder.File("other.xlsx");
        using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Create))
        {
            using var writer = new StreamWriter(zip.CreateEntry("readme.txt").Open());
            writer.Write("not a workbook");
        }

        var e = Assert.Throws<InvalidDataException>(() => new Workbook(path));
        Assert.IsType<ExcelDataReader.Exceptions.HeaderException>(e.InnerException);
    }

    [Fact]
    public void A_sheet_with_more_cells_than_the_limit_is_refused_before_it_is_read()
    {
        using var folder = new TempFolder();
        // Two values a million rows apart: a 1.6 KB file that reads as 1,048,576 rows.
        var path = Xlsx.Write(folder.File("tall.xlsx"),
            new SheetSpec("Small", ("A1", 1)),
            new SheetSpec("Tall", ("A1", 1), ("A1048576", 2)));

        long before = GC.GetAllocatedBytesForCurrentThread();
        var e = Assert.Throws<InvalidDataException>(() => new Workbook(path, false, maxCellsPerSheet: 1_000_000));
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Contains("'Tall'", e.Message);
        Assert.Contains("1048576 rows and 1 columns", e.Message);
        Assert.Contains("1000000", e.Message);
        Assert.True(allocated < 100_000_000, $"allocated {allocated:N0} bytes before refusing");
    }

    [Fact]
    public void A_workbook_within_the_limit_reads_as_without_one()
    {
        using var folder = new TempFolder();
        var path = Xlsx.Prices(folder.File("prices.xlsx"));

        var limited = new Workbook(path, false, maxCellsPerSheet: 9);

        Assert.Equal(9, limited["Sheet1"].Cells.Count);
        Assert.Equal("Pear", limited["Sheet1"]["A3"].Value);
        Assert.Throws<InvalidDataException>(() => new Workbook(path, false, maxCellsPerSheet: 8));
    }

    [Fact]
    public void The_limit_applies_to_the_copy_option_too()
    {
        using var folder = new TempFolder();
        var name = $"limited-{Guid.NewGuid():N}.xlsx";
        var source = Xlsx.Prices(folder.File(name));
        var copy = Path.Combine(AppContext.BaseDirectory, name);
        try
        {
            Assert.Throws<InvalidDataException>(() => new Workbook(source, true, maxCellsPerSheet: 8));
            Assert.Equal("Apple", new Workbook(source, true, maxCellsPerSheet: 9)["Sheet1"]["A2"].Value);
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void A_limit_below_one_throws_ArgumentOutOfRangeException(long limit)
    {
        using var folder = new TempFolder();
        var path = Xlsx.Prices(folder.File("prices.xlsx"));

        var e = Assert.Throws<ArgumentOutOfRangeException>(() => new Workbook(path, false, limit));
        Assert.Equal("maxCellsPerSheet", e.ParamName);
    }
}

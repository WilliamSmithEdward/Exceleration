using System.Data;
using System.Xml;
using ExcelDataReader.Exceptions;

namespace Exceleration.Tests;

public class WorkbookTests
{
    [Fact]
    public void Reads_every_sheet_in_workbook_order()
    {
        using var folder = new TempFolder();
        var path = Xlsx.Write(folder.File("book.xlsx"),
            new SheetSpec("First", ("A1", 1)),
            new SheetSpec("Second", ("A1", 2)),
            new SheetSpec("Third", ("A1", 3)));

        var wb = new Workbook(path);

        Assert.Equal(new[] { "First", "Second", "Third" }, wb.Sheets.Select(s => s.Name));
        Assert.All(wb.Sheets, s => Assert.Same(wb, s.Parent));
        Assert.Equal("book.xlsx", wb.Name);
        Assert.Equal(path, wb.FilePath);
    }

    [Fact]
    public void Finds_a_sheet_by_name_ignoring_case()
    {
        using var folder = new TempFolder();
        var wb = new Workbook(Xlsx.Prices(folder.File("prices.xlsx")));

        Assert.Same(wb["Sheet1"], wb["sHEET1"]);
    }

    [Fact]
    public void A_missing_sheet_throws_ArgumentException()
    {
        using var folder = new TempFolder();
        var wb = new Workbook(Xlsx.Prices(folder.File("prices.xlsx")));

        var e = Assert.Throws<ArgumentException>(() => wb["Nope"]);
        Assert.Contains("'Nope'", e.Message);
    }

    [Fact]
    public void Reads_the_file_into_memory_and_closes_it()
    {
        using var folder = new TempFolder();
        var path = Xlsx.Prices(folder.File("prices.xlsx"));

        var wb = new Workbook(path);
        File.Delete(path);

        Assert.False(File.Exists(path));
        Assert.Equal("Apple", wb["Sheet1"]["A2"].Value);
    }

    [Fact]
    public void A_missing_file_throws_FileNotFoundException()
    {
        using var folder = new TempFolder();

        Assert.Throws<FileNotFoundException>(() => new Workbook(folder.File("none.xlsx")));
    }

    [Theory]
    [InlineData("data.csv", "a,b\n1,2\n")]
    [InlineData("text.xlsx", "not a workbook")]
    [InlineData("empty.xlsx", "")]
    public void A_file_that_is_not_a_workbook_throws_HeaderException(string name, string content)
    {
        using var folder = new TempFolder();
        File.WriteAllText(folder.File(name), content);

        Assert.Throws<HeaderException>(() => new Workbook(folder.File(name)));
    }

    [Fact]
    public void A_truncated_xlsx_throws_InvalidDataException()
    {
        using var folder = new TempFolder();
        var bytes = File.ReadAllBytes(Xlsx.Prices(folder.File("prices.xlsx")));
        File.WriteAllBytes(folder.File("cut.xlsx"), bytes[..(bytes.Length / 2)]);

        Assert.Throws<InvalidDataException>(() => new Workbook(folder.File("cut.xlsx")));
    }

    [Fact]
    public void A_worksheet_with_a_DTD_is_refused_and_its_entities_are_not_resolved()
    {
        using var folder = new TempFolder();
        var secret = folder.File("secret.txt");
        File.WriteAllText(secret, "SECRET");
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + $"<!DOCTYPE worksheet [<!ENTITY e SYSTEM \"{new Uri(secret).AbsoluteUri}\">]>"
            + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
            + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>&e;</t></is></c></row></sheetData></worksheet>";
        var path = Xlsx.WriteRaw(folder.File("dtd.xlsx"), ("Sheet1", xml));

        Assert.Throws<XmlException>(() => new Workbook(path));
    }

    [Fact]
    public void AddSheet_adds_a_sheet_made_from_a_DataTable()
    {
        using var folder = new TempFolder();
        var wb = new Workbook(Xlsx.Prices(folder.File("prices.xlsx")));

        wb.AddSheet(wb["Sheet1"].ToDataTable(), "Copy");

        Assert.Equal(new[] { "Sheet1", "Copy" }, wb.Sheets.Select(s => s.Name));
        Assert.Equal("Pear", wb["Copy"]["A3"].Value);
        Assert.Same(wb, wb["Copy"].Parent);
    }

    [Fact]
    public void AddSheet_refuses_a_name_that_is_already_there()
    {
        using var folder = new TempFolder();
        var wb = new Workbook(Xlsx.Prices(folder.File("prices.xlsx")));

        Assert.Throws<ArgumentException>(() => wb.AddSheet(new DataTable(), "Sheet1"));
        Assert.Throws<ArgumentException>(() => wb.AddSheet(wb["Sheet1"]));
        Assert.Single(wb.Sheets);
    }
}

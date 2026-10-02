using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;

namespace Exceleration.Tests;

/// <summary>
/// A folder of its own under the system temp folder for one test, deleted afterwards.
/// </summary>
internal sealed class TempFolder : IDisposable
{
    public TempFolder()
    {
        Root = Path.Combine(Path.GetTempPath(), "exceleration-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Root);
    }

    public string Root { get; }

    public string File(string name) => Path.Combine(Root, name);

    public void Dispose()
    {
        try { Directory.Delete(Root, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}

/// <summary>
/// A worksheet for <see cref="Xlsx"/>: its name and its cells by A1 reference. A string is written
/// as an inline string, a bool as a boolean, a DateTime as a serial date with a date format, and
/// any other value as a number.
/// </summary>
internal sealed record SheetSpec(string Name, params (string Reference, object Value)[] Cells);

/// <summary>
/// Writes small .xlsx files for the tests, so no workbook is committed and none comes from
/// elsewhere. Each holds only the parts ExcelDataReader needs.
/// </summary>
internal static class Xlsx
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>The workbook the README samples read: Name, Qty, Price, then two rows.</summary>
    public static string Prices(string path) => Write(path, new SheetSpec("Sheet1",
        ("A1", "Name"), ("B1", "Qty"), ("C1", "Price"),
        ("A2", "Apple"), ("B2", 3), ("C2", 1.5),
        ("A3", "Pear"), ("B3", 7), ("C3", 2.25)));

    public static string Write(string path, params SheetSpec[] sheets) =>
        WriteRaw(path, sheets.Select(s => (s.Name, SheetXml(s.Cells))).ToArray());

    /// <summary>Writes the given worksheet XML as is, for workbooks a test wants malformed.</summary>
    public static string WriteRaw(string path, params (string Name, string Xml)[] sheets)
    {
        using var file = System.IO.File.Create(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Create);

        var types = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
            .Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">")
            .Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>")
            .Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>")
            .Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>")
            .Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        var workbook = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets>");
        var rels = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{PackageRel}\">")
            .Append($"<Relationship Id=\"rIdStyles\" Type=\"{Rel}/styles\" Target=\"styles.xml\"/>");
        for (int i = 1; i <= sheets.Length; i++)
        {
            types.Append($"<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            workbook.Append($"<sheet name=\"{SecurityElement.Escape(sheets[i - 1].Name)}\" sheetId=\"{i}\" r:id=\"rId{i}\"/>");
            rels.Append($"<Relationship Id=\"rId{i}\" Type=\"{Rel}/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>");
        }
        types.Append("</Types>");
        workbook.Append("</sheets></workbook>");
        rels.Append("</Relationships>");

        Add(zip, "[Content_Types].xml", types.ToString());
        Add(zip, "_rels/.rels", $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{PackageRel}\"><Relationship Id=\"rId1\" Type=\"{Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        Add(zip, "xl/workbook.xml", workbook.ToString());
        Add(zip, "xl/_rels/workbook.xml.rels", rels.ToString());
        // Style 0 is General; style 1 is the built-in date format 14 (m/d/yyyy).
        Add(zip, "xl/styles.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"{Main}\"><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>");
        for (int i = 1; i <= sheets.Length; i++)
        {
            Add(zip, $"xl/worksheets/sheet{i}.xml", sheets[i - 1].Xml);
        }
        return path;
    }

    /// <summary>Worksheet XML for the given cells, rows in order.</summary>
    public static string SheetXml(params (string Reference, object Value)[] cells)
    {
        var xml = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"{Main}\"><sheetData>");
        var rows = cells
            .GroupBy(c => int.Parse(new string(c.Reference.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture))
            .OrderBy(g => g.Key);
        foreach (var row in rows)
        {
            xml.Append($"<row r=\"{row.Key}\">");
            foreach (var (reference, value) in row)
            {
                xml.Append(value switch
                {
                    string s => $"<c r=\"{reference}\" t=\"inlineStr\"><is><t>{SecurityElement.Escape(s)}</t></is></c>",
                    bool b => $"<c r=\"{reference}\" t=\"b\"><v>{(b ? 1 : 0)}</v></c>",
                    DateTime d => $"<c r=\"{reference}\" s=\"1\"><v>{d.ToOADate().ToString(CultureInfo.InvariantCulture)}</v></c>",
                    _ => $"<c r=\"{reference}\"><v>{Convert.ToString(value, CultureInfo.InvariantCulture)}</v></c>",
                });
            }
            xml.Append("</row>");
        }
        return xml.Append("</sheetData></worksheet>").ToString();
    }

    private static void Add(ZipArchive zip, string name, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}

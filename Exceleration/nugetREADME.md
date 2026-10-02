# Exceleration

Exceleration reads Excel workbooks into memory through [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) and lets you address the cells the way Excel does: by sheet name, by A1 reference, or by row and column number counted from 1. It reads; it does not write workbooks.

```
dotnet add package Exceleration
```

Everything is in the `Exceleration` namespace. The package targets net8.0, net9.0 and net10.0 and depends on ExcelDataReader and ExcelDataReader.DataSet 3.7.0.

---

## Read cells

The samples below read `prices.xlsx`, whose first sheet, `Sheet1`, holds:

|   | A | B | C |
|---|---|---|---|
| 1 | Name | Qty | Price |
| 2 | Apple | 3 | 1.5 |
| 3 | Pear | 7 | 2.25 |

```csharp
using Exceleration;

var wb = new Workbook("prices.xlsx");
var ws = wb["Sheet1"];

Console.WriteLine(ws["A2"].Value);            // Apple
Console.WriteLine(ws.GetCell(3, 2).Address);  // B3
Console.WriteLine(ws.GetCellValue("C3"));     // 2.25
Console.WriteLine(ws.GetCellValue(1, 2));     // Qty
```

## Convert values

```csharp
using Exceleration;

var ws = new Workbook("prices.xlsx")["Sheet1"];

int qty = ws["B2"].To<int>();              // 3
decimal price = ws["C2"].To<decimal>();    // 1.5
int bad = ws["A2"].To<int>();              // 0: "Apple" is not a number
int? none = ws["A2"].ToNullable<int>();    // null
bool ok = ws["B3"].IsParseable<int>();     // True

Console.WriteLine($"{qty} {price} {bad} {none is null} {ok}");
```

## Rows and columns

```csharp
using Exceleration;

var ws = new Workbook("prices.xlsx")["Sheet1"];

foreach (var row in ws.Rows)
{
    Console.WriteLine(string.Join(", ", row.Select(c => c.Value)));
}
// Name, Qty, Price
// Apple, 3, 1.5
// Pear, 7, 2.25

var names = ws.GetColumn("A").Skip(1).Select(c => c.Value);
Console.WriteLine(string.Join(", ", names));  // Apple, Pear

var header = ws.GetRow(1);
Console.WriteLine(header.GetFirstCellByColumnLetter("C").Value);  // Price

var expensive = ws.Rows
    .Skip(1)
    .Where(r => r.GetFirstCellByColumnLetter("C").To<double>() > 2)
    .Select(r => r.GetFirstCellByColumnNumber(1).Value);
Console.WriteLine(string.Join(", ", expensive));  // Pear
```

`ws.Rows` and `ws.Columns` are properties that return a list of lists. `GetRow`, `GetColumn` and the `CellListExtensions` methods (`GetFirstCellByRowNumber`, `GetFirstCellByColumnLetter`, `GetFirstCellByColumnNumber`, and `GetRow` and `GetColumn` on a `List<Cell>`) pick cells out of them.

## DataTable and added sheets

```csharp
using System.Data;
using Exceleration;

var wb = new Workbook("prices.xlsx");

DataTable table = wb["Sheet1"].ToDataTable();
Console.WriteLine(table.Columns[0].ColumnName);  // Column0
Console.WriteLine(table.Rows[1][0]);             // Apple

wb.AddSheet(table, "Copy");
Console.WriteLine(wb["Copy"]["A3"].Value);       // Pear
```

`ToDataTable` returns a copy of the sheet's data. Its columns are named `Column0`, `Column1` and so on, and row 1 of the sheet is the table's first row.

---

## How a workbook is read

- `new Workbook(filePath)` reads every sheet of the file into memory and closes it before returning. Later changes to the file are not seen; construct a new `Workbook` to read them.
- It reads the formats ExcelDataReader's `ExcelReaderFactory.CreateReader` detects from the file's content: .xlsx and .xlsm, .xlsb, and .xls. It does not read CSV: a .csv file throws `ExcelDataReader.Exceptions.HeaderException` ("Invalid file signature").
- The file is opened for reading without sharing. If another process has it open, Excel included, the constructor throws `IOException`. See "Known problems" below.
- `new Workbook(filePath, true)` copies the file first, into the folder `Exceleration.dll` was loaded from, under the same file name, and reads the copy. It replaces a file of that name in that folder, leaves the copy there afterwards, and `FilePath` then names the copy.
- Each constructor registers `CodePagesEncodingProvider.Instance` with `Encoding.RegisterProvider`, which ExcelDataReader needs for the code pages of older .xls files. The registration applies to the whole process.
- `Name` is the file name of the path you passed, and `Sheets` lists the sheets in workbook order.

## Sheets, cells and values

- `wb["Sheet1"]` finds a sheet by name, ignoring case. A name that is not there throws `ArgumentException`.
- Row 1 of a sheet is row 1: no row is taken as a header. Empty rows and columns before the first value are kept, so a sheet whose only value is in B3 still has B3 at row 3, column 2.
- Rows and columns are numbered from 1. `GetCell`, `GetCellValue`, `GetRow` and `GetColumn` throw `ArgumentOutOfRangeException` for a row or column outside the sheet's used range.
- `Value` is what ExcelDataReader read: a `string`, a `double` for every number, a `bool`, a `DateTime` for a cell with a date format, or `DBNull.Value` for an empty cell. A formula cell holds the value Excel last calculated and saved. `DataType` is the type of `Value`, `typeof(DBNull)` for an empty cell.
- `To<T>()` converts with `Convert.ChangeType`, which uses the current culture and rounds a fractional number to the nearest even integer (1.5 and 2.5 both give 2). It returns `default(T)` when the conversion fails, or throws `InvalidCastException` with `To<T>(false)`.
- `ToNullable<T>()` returns `null` for an empty cell, a cell that holds only spaces, or a value that does not convert. `IsParseable<T>()` returns `false` in the same cases.
- `ToString()` returns `Value.ToString()`, an empty string for an empty cell.
- `wb.AddSheet(table, name)` adds a sheet made from a `DataTable`, and `wb.AddSheet(sheet)` adds an existing `Worksheet`. Both throw `ArgumentException` when a sheet with the same name, compared with case, is already there. Nothing is written back to the file.

## Reading files you did not create

- The whole workbook is held in memory, and a sheet takes room for every cell from A1 to its last used cell, empty or not. A 1.6 KB .xlsx whose only values are in A1 and CV100000 becomes a 100,000 by 100 table and took about 0.9 GB to read. Exceleration has no size limit, so check where a file comes from and how large it is before you read it.
- A file that is not a workbook, or is damaged, throws whatever ExcelDataReader or .NET raised while reading it, for example `ExcelDataReader.Exceptions.HeaderException`, `System.IO.InvalidDataException` for a damaged .zip, or `System.Xml.XmlException` for broken XML. An .xlsx part with a DTD is refused with `XmlException`, so external entities are not resolved.
- Pass `true` for the copy only when you control the file name, because the copy replaces any file of that name in the application's folder.

## Known problems in 1.1.1.4

- `Worksheet.Cells` throws `ArgumentOutOfRangeException` for any sheet that has a cell.
- `Cell.Offset` is off by one row and one column: `ws["B2"].Offset(0, 0)` returns A1, and an offset from row 1 or column A throws.
- `GetCellValue(string)` reads the row digits backwards, so `GetCellValue("A12")` reads A21, and it accepts letters after the digits ("1A").
- Lower-case references read the wrong column: `ws["a1"]` and `GetCellValue("a1")` look for column 33, and `GetColumn("b")` for column 34. The `CellListExtensions` methods compare column letters with case.
- A reference with extra characters is read as the part that looks like one: `ws["A1B"]` and `ws["A1:B2"]` return A1. A row number too large for an `int` throws `OverflowException`.
- `ArgumentOutOfRangeException` from the cell methods carries its message as the parameter name, so it reads "Specified argument was out of the range of valid values. (Parameter 'Invalid row 0 or column 0 index.')".
- The constructor opens the file without sharing, so it fails while Excel, or another `Workbook` on another thread, has the file open.
- `new Workbook(path, true)` throws `IOException` when the file is already in the folder it copies to, because it copies the file onto itself.
- `ToNullable<T>(false)` returns `null` on a failed conversion, exactly like `ToNullable<T>()`.
- `AddSheet` compares names with case while `wb[name]` ignores it, so a sheet named "sheet1" can be added next to "Sheet1" and cannot be found by name.
- `AddSheet(table, name)` renames the caller's `DataTable` and keeps it, so later changes to the table change the sheet. `AddSheet(sheet)` with a sheet from another workbook leaves its `Parent` pointing at that workbook.

## Attributions

Icon "Exceleration.png" designed by Stockes Design on Freepik.com.

## License

[MIT](https://github.com/WilliamSmithEdward/Exceleration/blob/main/LICENSE)

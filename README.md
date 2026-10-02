# Exceleration

[![NuGet version](https://img.shields.io/nuget/v/Exceleration)](https://www.nuget.org/packages/Exceleration)
[![Downloads](https://img.shields.io/nuget/dt/Exceleration)](https://www.nuget.org/packages/Exceleration)
[![CI](https://github.com/WilliamSmithEdward/Exceleration/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/WilliamSmithEdward/Exceleration/actions/workflows/ci.yml)
[![Security](https://github.com/WilliamSmithEdward/Exceleration/actions/workflows/security.yml/badge.svg?branch=main)](https://github.com/WilliamSmithEdward/Exceleration/actions/workflows/security.yml)
[![Malware scan](https://github.com/WilliamSmithEdward/Exceleration/actions/workflows/malware-scan.yml/badge.svg?branch=main)](https://github.com/WilliamSmithEdward/Exceleration/actions/workflows/malware-scan.yml)
[![OpenSSF Scorecard](https://api.scorecard.dev/projects/github.com/WilliamSmithEdward/Exceleration/badge)](https://scorecard.dev/viewer/?uri=github.com/WilliamSmithEdward/Exceleration)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](https://github.com/WilliamSmithEdward/Exceleration/blob/main/LICENSE)

Exceleration reads Excel workbooks into memory through [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) and lets you address the cells the way Excel does: by sheet name, by A1 reference, or by row and column number counted from 1. It reads; it does not write workbooks.

```
dotnet add package Exceleration
```

Everything is in the `Exceleration` namespace. The package targets net8.0, net9.0 and net10.0 and depends on ExcelDataReader and ExcelDataReader.DataSet 3.9.0.

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
Console.WriteLine(ws["b2"].Offset(1, 1).Value); // 2.25
Console.WriteLine(ws.Cells.Count);            // 9
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
- The file is opened for reading and shared with other readers and writers, so a workbook that Excel, another process or another thread has open can be read. It throws `IOException` only when another process holds the file without sharing it.
- `new Workbook(filePath, true)` copies the file first, into the folder `Exceleration.dll` was loaded from, under the same file name, and reads the copy. It replaces a file of that name in that folder, leaves the copy there afterwards, and `FilePath` then names the copy. A file that is already in that folder is read where it is.
- Each constructor registers `CodePagesEncodingProvider.Instance` with `Encoding.RegisterProvider`, which ExcelDataReader needs for the code pages of older .xls files. The registration applies to the whole process.
- `Name` is the file name of the path you passed, and `Sheets` lists the sheets in workbook order.

## Sheets, cells and values

- `wb["Sheet1"]` finds a sheet by name, ignoring case. A name that is not there throws `ArgumentException`.
- Row 1 of a sheet is row 1: no row is taken as a header. Empty rows and columns before the first value are kept, so a sheet whose only value is in B3 still has B3 at row 3, column 2.
- An A1 reference is column letters, in either case, then the row number, and nothing else: "B12" and "b12" are the same cell. Anything else, such as "$B$12", "B12:C13" or " B12", throws `ArgumentException`, and a null reference throws `ArgumentNullException`.
- Rows and columns are numbered from 1. `GetCell`, `GetCellValue`, `GetRow`, `GetColumn` and `Offset` throw `ArgumentOutOfRangeException` for a row or column outside the sheet's used range; its `ParamName` is `rowNumber` or `colNumber`, and its message gives the sheet's size.
- `Offset(rows, columns)` returns the cell that many rows down and columns to the right; negative numbers go up and left. `Cells` lists every cell of the used range, row by row.
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

- `ToNullable<T>(false)` returns `null` on a failed conversion, exactly like `ToNullable<T>()`.
- `AddSheet` compares names with case while `wb[name]` ignores it, so a sheet named "sheet1" can be added next to "Sheet1" and cannot be found by name.
- `AddSheet(table, name)` renames the caller's `DataTable` and keeps it, so later changes to the table change the sheet. `AddSheet(sheet)` with a sheet from another workbook leaves its `Parent` pointing at that workbook.

## Attributions

Icon "Exceleration.png" designed by Stockes Design on Freepik.com.

## License

[MIT](https://github.com/WilliamSmithEdward/Exceleration/blob/main/LICENSE)

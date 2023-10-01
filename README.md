# Exceleration
Exceleration is a C# library that wraps ExcelDataReader. It helps you better interact with Excel files by simplifying the development experience.

## Example

### Basic Usage
```csharp
using Exceleration;

var wb = new Workbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xlsx"));

var ws = wb["Sheet1"];

var cell = ws["A1"];

Console.WriteLine(cell.Address + ": " + cell.Value);

cell = cell.Offset(1, 0);

Console.WriteLine(cell.Address + ": " + cell.Value);

cell = cell.Offset(0, 1);

Console.WriteLine(cell.Address + ": " + cell.Value);
```

### Get All Cells
```csharp
using Exceleration;

var wb = new Workbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xlsx"));

var ws = wb["Sheet1"];

var cells = ws.Cells;

Console.WriteLine(cells.First().Address + ": " + cells.First().Value);
```

### Using LINQ to Filter Cells
```csharp
using Exceleration;

var wb = new Workbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xlsx"));

var ws = wb["Sheet1"];

var cells = ws.Cells.Where(x => x.ColLetter.Equals("A"));

foreach (var item in cells)
{
    Console.WriteLine(item.Address + ": " + item.Value + " " + item.ColLetter);
}
```

### Accessing Rows and Columns Collections
```csharp
using Exceleration;

var wb = new Workbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xlsx"));

var ws = wb["Sheet1"];

var columns = ws.Columns("A");

var rows = ws.Rows(1);

Console.WriteLine(string.Join(",", columns.Select(x => x.Value)));

Console.WriteLine(string.Join(",", rows.Select(x => x.Value)));
```

## Attributions
NuGet icon "SQLServerInteraction.png" designed by Stockes Design on Freepik.com

# Exceleration
Exceleration is a C# library that wraps ExcelDataReader. It helps you better interact with Excel files by simplifying the development experience.

## Example

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

## Attributions
NuGet icon "SQLServerInteraction.png" designed by Stockes Design on Freepik.com

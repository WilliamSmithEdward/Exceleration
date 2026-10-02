# Changelog

Each release's notes. The Publish workflow takes the section for the
version it releases as the GitHub release's body, so a section is written
here before the version is tagged.

The sections up to 1.1.1.4 were gathered from the nuget.org version history,
with the UTC date the nuget.org catalog records for each upload. Neither
nuget.org nor the READMEs carried release notes for them. Versions 1.0.0 to
1.1.1.3 are unlisted on nuget.org; 1.1.1.4 is the listed version. All of
them target net7.0.

## [2.0.0] - 2026-10-02

Cell addressing works as documented: `Worksheet.Cells` and `Cell.Offset` no longer throw or land on the wrong cell, references read in either case and multi-digit rows read correctly, and anything that is not a reference is refused. The constructor reads a workbook that Excel has open, a file that does not parse throws one exception type, and a new overload limits how many cells a sheet from an untrusted file may make the library hold. Several of the fixes change what callers see, hence the major version. ExcelDataReader and ExcelDataReader.DataSet are 3.9.0, up from 3.7.0.

### Breaking changes

* The package targets net8.0, net9.0 and net10.0. net7.0, which is out of support, is dropped.
* A file that is not a workbook, or is damaged, throws `InvalidDataException` naming the file, with the exception ExcelDataReader or .NET raised as its `InnerException`. It used to throw that exception itself, such as `ExcelDataReader.Exceptions.HeaderException` for a CSV file or `XmlException` for broken XML.
* An A1 reference must be column letters then the row number and nothing else. "A1B", "A1:B2", " A1" and "1A" throw `ArgumentException`; they used to read some cell. A row number too large for an `int` throws `ArgumentOutOfRangeException` instead of `OverflowException`. `GetColumn(string)` throws `ArgumentException` for anything but letters.
* `ArgumentOutOfRangeException` from the cell methods names `rowNumber` or `colNumber` as its `ParamName`, carries the value, and says how many rows or columns the sheet has. Its message used to be passed as the parameter name.
* `Cell.Offset(rows, columns)` returns the cell that many rows and columns away. It used to return the cell one row up and one column left of that, or throw.
* The constructor opens the file shared with other readers and writers, so another process may write it during the read. It used to share nothing, and failed whenever anything else had the file open.
* `ToNullable<T>(false)` returns `default(T)` for a value that does not convert. It used to return `null`, as `ToNullable<T>()` does.
* `AddSheet` compares sheet names ignoring case, as `wb[name]` finds them. `AddSheet(table, name)` adds a copy of the table and no longer renames or keeps the caller's. `AddSheet(sheet)` adds a copy owned by the workbook, so the object in `Sheets` is not the one passed in.

### Fixes

* `Worksheet.Cells` threw `ArgumentOutOfRangeException` for any sheet with a cell. It lists every cell, row by row.
* `GetCellValue(string)` read the row digits backwards, so "A12" read A21, and accepted "1A". It reads references as the indexer does.
* Lower-case references read the wrong column ("a1" looked for column 33), and the `CellListExtensions` methods compared column letters with case. Letters are read in either case.
* The constructor failed with `IOException` while Excel or another thread had the workbook open. The file is opened with `FileShare.ReadWrite`, as the source had it until the copy option was added.
* `new Workbook(path, true)` failed with `IOException` when the workbook was already in the folder it copies to. The copy is skipped and the file read in place.
* `AddSheet` let "sheet1" be added next to "Sheet1", where it could not be found by name, and `AddSheet(sheet)` with a sheet from another workbook left its `Parent` pointing at that workbook.
* `GetRow` and `GetColumn` on a sheet with no cells returned an empty list for any index instead of throwing. A null reference or argument throws `ArgumentNullException` naming the method's own parameter.
* The build is deterministic, so the same commit gives the same dll.

### Additions

* `Workbook(string filePath, bool copyFileToExeDirectoryBeforeRead, long maxCellsPerSheet)` refuses a workbook with `InvalidDataException`, before any sheet is read, when one sheet's used range has more cells than the limit. A sheet holds every cell from A1 to its last used cell, so a 1.6 KB file can otherwise take gigabytes.
* The READMEs are rewritten against the code, with samples that compile and run, and SECURITY.md says what the library reads and writes and how to read untrusted workbooks.
* The package is built in CI from the tagged commit, tested on all three frameworks, scanned for vulnerabilities and malware, and published through nuget.org trusted publishing. The GitHub release carries the package's signed build provenance.

## [1.1.1.4] - 2024-06-27

No notes were recorded.

## [1.1.1.3] - 2023-10-09

No notes were recorded.

## [1.1.1.2] - 2023-10-09

No notes were recorded.

## [1.1.1.1] - 2023-10-09

No notes were recorded.

## [1.1.1] - 2023-10-09

No notes were recorded.

## [1.1.0] - 2023-10-07

No notes were recorded.

## [1.0.5.9] - 2023-10-05

No notes were recorded.

## [1.0.5.8] - 2023-10-05

No notes were recorded.

## [1.0.5.7] - 2023-10-05

No notes were recorded.

## [1.0.5.6] - 2023-10-05

No notes were recorded.

## [1.0.5.5] - 2023-10-05

No notes were recorded.

## [1.0.5.4] - 2023-10-05

No notes were recorded.

## [1.0.5.3] - 2023-10-05

No notes were recorded.

## [1.0.5.2] - 2023-10-05

No notes were recorded.

## [1.0.5.1] - 2023-10-05

No notes were recorded.

## [1.0.5] - 2023-10-05

No notes were recorded.

## [1.0.4.1] - 2023-10-05

No notes were recorded.

## [1.0.4] - 2023-10-05

No notes were recorded.

## [1.0.3] - 2023-10-05

No notes were recorded.

## [1.0.2] - 2023-10-01

No notes were recorded.

## [1.0.1] - 2023-09-30

No notes were recorded.

## [1.0.0.1] - 2023-09-30

No notes were recorded.

## [1.0.0] - 2023-09-30

No notes were recorded.

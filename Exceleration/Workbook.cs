using ExcelDataReader;
using System.Data;
using System.Reflection;
using System.Text;

namespace Exceleration
{
    /// <summary>
    /// A workbook read into memory from a file through ExcelDataReader, as a list of worksheets.
    /// </summary>
    public class Workbook
    {
        /// <summary>
        /// Gets the path the workbook was read from: the path passed to the constructor, or the path of the copy
        /// when the constructor copied the file first.
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// Gets the file name of the path passed to the constructor.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the worksheets in workbook order. Adding to this list directly skips the name check of <see cref="AddSheet(Worksheet)"/>.
        /// </summary>
        public List<Worksheet> Sheets { get; private set; }

        /// <summary>
        /// Reads every sheet of the workbook at <paramref name="filePath"/> into memory and closes the file.
        /// </summary>
        /// <remarks>
        /// The formats are those ExcelDataReader's <c>ExcelReaderFactory.CreateReader</c> detects: .xlsx, .xlsm, .xlsb and .xls, not CSV.
        /// The file is opened for reading and shared with other readers and writers, so a workbook Excel or another
        /// thread has open can be read. The constructor also registers
        /// <see cref="CodePagesEncodingProvider.Instance"/> with <see cref="Encoding.RegisterProvider(EncodingProvider)"/> for the whole process.
        /// </remarks>
        /// <param name="filePath">The path of the workbook file.</param>
        /// <param name="copyFileToExeDirectoryBeforeRead">When true, copies the file into the folder Exceleration.dll was loaded from,
        /// under the same file name and replacing a file of that name, and reads the copy, which is left there. A file already
        /// in that folder is read where it is. Optional; false by default.</param>
        /// <exception cref="IOException">The file cannot be opened, for example because another process holds it without sharing.</exception>
        /// <exception cref="FileNotFoundException">There is no file at <paramref name="filePath"/>.</exception>
        /// <exception cref="InvalidDataException">The file is not a workbook ExcelDataReader can read, or is damaged. The exception
        /// ExcelDataReader or .NET raised is the <see cref="Exception.InnerException"/>.</exception>
        public Workbook(string filePath, bool copyFileToExeDirectoryBeforeRead = false)
            : this(filePath, copyFileToExeDirectoryBeforeRead, long.MaxValue)
        {
        }

        /// <summary>
        /// Reads every sheet of the workbook at <paramref name="filePath"/> into memory and closes the file, refusing a
        /// workbook with a sheet larger than <paramref name="maxCellsPerSheet"/> before any sheet is read.
        /// </summary>
        /// <remarks>
        /// A sheet holds a cell for every row and column from A1 to its last used cell, empty or not, so a small file can
        /// need a great deal of memory. Pass a limit when the workbook comes from somewhere you do not control. Otherwise
        /// this reads as <see cref="Workbook(string, bool)"/> does.
        /// </remarks>
        /// <param name="filePath">The path of the workbook file.</param>
        /// <param name="copyFileToExeDirectoryBeforeRead">As for <see cref="Workbook(string, bool)"/>.</param>
        /// <param name="maxCellsPerSheet">The most cells, rows times columns of the used range, any one sheet may have.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="maxCellsPerSheet"/> is less than 1.</exception>
        /// <exception cref="IOException">The file cannot be opened, for example because another process holds it without sharing.</exception>
        /// <exception cref="FileNotFoundException">There is no file at <paramref name="filePath"/>.</exception>
        /// <exception cref="InvalidDataException">A sheet has more cells than the limit, or the file is not a workbook ExcelDataReader
        /// can read, or is damaged. For an unreadable file, the exception ExcelDataReader or .NET raised is the
        /// <see cref="Exception.InnerException"/>.</exception>
        public Workbook(string filePath, bool copyFileToExeDirectoryBeforeRead, long maxCellsPerSheet)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(maxCellsPerSheet, 1);

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            FilePath = filePath;
            Name = Path.GetFileName(filePath);

            if (copyFileToExeDirectoryBeforeRead)
            {
                string exeDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
                string destinationPath = Path.Combine(exeDirectory, Name);
                if (!IsSameFile(filePath, destinationPath))
                {
                    File.Copy(filePath, destinationPath, true);
                }
                FilePath = destinationPath;
            }

            Sheets = new List<Worksheet>();

            using var stream = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            DataSet? result = null;
            string? overLimit;
            try
            {
                using var reader = ExcelReaderFactory.CreateReader(stream);
                overLimit = FindSheetOverLimit(reader, maxCellsPerSheet);
                if (overLimit is null)
                {
                    result = reader.AsDataSet(new ExcelDataSetConfiguration());
                }
            }
            // Whatever the parse hit, from a wrong signature to broken XML or a damaged zip, comes out as one type the
            // caller can catch. Running out of memory is not a property of the file, so it passes through.
            catch (Exception e) when (e is not OutOfMemoryException)
            {
                throw new InvalidDataException($"'{Name}' is not a workbook Exceleration can read: {e.Message}", e);
            }

            if (overLimit is not null)
            {
                throw new InvalidDataException(overLimit);
            }

            foreach (DataTable table in result!.Tables)
            {
                Sheets.Add(new Worksheet(table, this));
            }
        }

        // ExcelDataReader knows each sheet's used range before it reads the cells, and AsDataSet allocates all of it.
        // Returns why the workbook is refused, or null, with the reader back at its first sheet.
        private string? FindSheetOverLimit(IExcelDataReader reader, long maxCellsPerSheet)
        {
            if (maxCellsPerSheet == long.MaxValue)
            {
                return null;
            }

            do
            {
                long cells = (long)reader.RowCount * reader.FieldCount;
                if (cells > maxCellsPerSheet)
                {
                    return $"Sheet '{reader.Name}' of '{Name}' has {reader.RowCount} rows and {reader.FieldCount} columns, "
                        + $"{cells} cells, more than the limit of {maxCellsPerSheet}.";
                }
            } while (reader.NextResult());

            reader.Reset();
            return null;
        }

        /// <summary>
        /// Gets the worksheet with the specified name from the workbook, comparing names without regard to case.
        /// </summary>
        /// <param name="sheetName">The name of the worksheet to retrieve.</param>
        /// <returns>The worksheet with the specified name.</returns>
        /// <exception cref="ArgumentException">Thrown if no worksheet is found with the given name.</exception>
        public Worksheet this[string sheetName]
        {
            get
            {
                var sheet = Sheets.FirstOrDefault(s => s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)) ?? throw new ArgumentException($"No worksheet found with the name '{sheetName}'.");
                return sheet;
            }
        }

        /// <summary>
        /// Adds a copy of a worksheet, from this workbook or another, to the workbook. The copy belongs to this workbook;
        /// the worksheet passed in is not changed.
        /// </summary>
        /// <param name="sheet">The worksheet to copy.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="sheet"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown if a worksheet with the same name, compared without regard to case, already exists.</exception>
        public void AddSheet(Worksheet sheet)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            CheckNameIsFree(sheet.Name);

            Sheets.Add(new Worksheet(sheet.DataTable.Copy(), this));
        }

        /// <summary>
        /// Adds a worksheet with the specified name to the workbook, made from a copy of the given DataTable.
        /// The table passed in is not changed, and later changes to it do not reach the worksheet.
        /// </summary>
        /// <param name="table">The DataTable representing the worksheet data.</param>
        /// <param name="workSheetName">The name of the worksheet to add.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="table"/> or <paramref name="workSheetName"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown if a worksheet with the same name, compared without regard to case, already exists.</exception>
        public void AddSheet(DataTable table, string workSheetName)
        {
            ArgumentNullException.ThrowIfNull(table);
            ArgumentNullException.ThrowIfNull(workSheetName);
            CheckNameIsFree(workSheetName);

            var copy = table.Copy();
            copy.TableName = workSheetName;

            Sheets.Add(new Worksheet(copy, this));
        }

        // Sheet names compare as the indexer finds them, and as Excel compares them: without regard to case.
        private void CheckNameIsFree(string name)
        {
            if (Sheets.Any(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException($"Worksheet named '{name}' already exists.");
            }
        }

        // Copying a file onto itself fails, because File.Copy opens the destination for writing while it
        // reads the source. Windows and macOS file systems ignore case by default; Linux ones do not.
        private static bool IsSameFile(string first, string second)
        {
            var comparison = OperatingSystem.IsLinux() ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            return string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), comparison);
        }
    }
}
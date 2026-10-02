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
        /// The file is opened for reading without sharing. The constructor also registers
        /// <see cref="CodePagesEncodingProvider.Instance"/> with <see cref="Encoding.RegisterProvider(EncodingProvider)"/> for the whole process.
        /// </remarks>
        /// <param name="filePath">The path of the workbook file.</param>
        /// <param name="copyFileToExeDirectoryBeforeRead">When true, copies the file into the folder Exceleration.dll was loaded from,
        /// under the same file name and replacing a file of that name, and reads the copy, which is left there. Optional; false by default.</param>
        /// <exception cref="IOException">The file cannot be opened, for example because another process has it open.</exception>
        /// <exception cref="FileNotFoundException">There is no file at <paramref name="filePath"/>.</exception>
        /// <exception cref="ExcelDataReader.Exceptions.HeaderException">The file is not in a format ExcelDataReader reads.</exception>
        public Workbook(string filePath, bool copyFileToExeDirectoryBeforeRead = false)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            FilePath = filePath;
            Name = Path.GetFileName(filePath);

            if (copyFileToExeDirectoryBeforeRead)
            {
                string exeDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
                string destinationPath = Path.Combine(exeDirectory, Name);
                File.Copy(filePath, destinationPath, true);
                FilePath = destinationPath;
            }

            Sheets = new List<Worksheet>();

            using var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var result = reader.AsDataSet(new ExcelDataSetConfiguration());

            foreach (DataTable table in result.Tables)
            {
                Sheets.Add(new Worksheet(table, this));
            }
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
        /// Adds a worksheet to the workbook.
        /// </summary>
        /// <param name="sheet">The worksheet to add.</param>
        /// <exception cref="ArgumentException">Thrown if a worksheet with the same name, compared with case, already exists.</exception>
        public void AddSheet(Worksheet sheet)
        {
            if (Sheets.Any(x => x.Name.Equals(sheet.Name))) throw new ArgumentException($"Worksheet named '{ sheet.Name }' already exists.");

            Sheets.Add(sheet);
        }

        /// <summary>
        /// Adds a worksheet with the specified name to the workbook, using the given DataTable.
        /// </summary>
        /// <param name="table">The DataTable representing the worksheet data.</param>
        /// <param name="workSheetName">The name of the worksheet to add. The table's <see cref="DataTable.TableName"/> is set to it.</param>
        /// <exception cref="ArgumentException">Thrown if a worksheet with the same name, compared with case, already exists.</exception>
        public void AddSheet(DataTable table, string workSheetName)
        {
            if (Sheets.Any(x => x.Name.Equals(workSheetName))) throw new ArgumentException($"Worksheet named '{ workSheetName }' already exists.");

            table.TableName = workSheetName;

            Sheets.Add(new Worksheet(table, this));
        }
    }
}
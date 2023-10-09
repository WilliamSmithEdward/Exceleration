using ExcelDataReader;
using System.Data;
using System.Reflection;
using System.Text;

namespace Exceleration
{
    /// <summary>
    /// Represents a workbook, which is a collection of worksheets.
    /// </summary>
    public class Workbook
    {
        /// <summary>
        /// Gets the file path of the workbook.
        /// </summary>
        public string FilePath { get; private set; }
        
        /// <summary>
        /// Gets the name of the workbook.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the list of worksheets in the workbook.
        /// </summary>
        public List<Worksheet> Sheets { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Workbook"/> class.
        /// </summary>
        /// <param name="filePath">The file path of the workbook.</param>
        /// /// <param name="copyFileToExeDirectoryBeforeRead">Should the file be copied to the exe directory prior to reading? (Optional)</param>
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
        /// Gets the worksheet with the specified name from the workbook.
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
        /// <exception cref="ArgumentException">Thrown if a worksheet with the same name already exists.</exception>
        public void AddSheet(Worksheet sheet)
        {
            if (Sheets.Any(x => x.Name.Equals(sheet.Name))) throw new ArgumentException($"Worksheet named '{ sheet.Name }' already exists.");

            Sheets.Add(sheet);
        }

        /// <summary>
        /// Adds a worksheet with the specified name to the workbook, using the given DataTable.
        /// </summary>
        /// <param name="table">The DataTable representing the worksheet data.</param>
        /// <param name="workSheetName">The name of the worksheet to add.</param>
        /// <exception cref="ArgumentException">Thrown if a worksheet with the same name already exists.</exception>
        public void AddSheet(DataTable table, string workSheetName)
        {
            if (Sheets.Any(x => x.Name.Equals(workSheetName))) throw new ArgumentException($"Worksheet named '{ workSheetName }' already exists.");

            table.TableName = workSheetName;

            Sheets.Add(new Worksheet(table, this));
        }
    }
}
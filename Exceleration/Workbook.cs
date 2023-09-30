using ExcelDataReader;
using System.Data;
using System.Text;

namespace Exceleration
{
    public class Workbook
    {
        public string FilePath { get; private set; }
        public string Name { get; private set; }
        public List<Worksheet> Sheets { get; private set; }

        public Workbook(string filePath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            FilePath = filePath;
            Name = Path.GetFileName(filePath);

            Sheets = new List<Worksheet>();

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var result = reader.AsDataSet(new ExcelDataSetConfiguration());

            foreach (DataTable table in result.Tables)
            {
                Sheets.Add(new Worksheet(table));
            }
        }

        public Worksheet this[string sheetName]
        {
            get
            {
                var sheet = Sheets.FirstOrDefault(s => s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet == null)
                {
                    throw new ArgumentException($"No worksheet found with the name '{sheetName}'.");
                }
                return sheet;
            }
        }
    }
}
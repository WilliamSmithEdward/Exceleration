using ExcelDataReader;
using System.Data;
using System.Text;

namespace Exceleration
{
    public class Workbook
    {
        public List<Worksheet> Sheets { get; private set; }

        public Workbook(string filePath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var result = reader.AsDataSet(new ExcelDataSetConfiguration());

            foreach (DataTable table in result.Tables)
            {
                Sheets.Add(new Worksheet(table));
            }
        }
    }
}
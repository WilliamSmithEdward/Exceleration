using System.Data;

namespace Exceleration
{
    public class Worksheet
    {
        private DataTable DataTable { get; set; }
        public string Name { get; set; }
        
        internal Worksheet(DataTable table)
        {
            DataTable = table;
        }
    }
}

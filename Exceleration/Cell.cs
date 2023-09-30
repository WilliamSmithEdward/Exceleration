using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration
{
    public class Cell
    {
        private object _value;

        public object Value
        {
            get { return _value; }
            set
            {
                // Update the value in the underlying DataTable as well, if needed
                Parent.DataTable.Rows[RowIndex][ColIndex] = value;

                // Update the value in the Cell object
                _value = value;
            }
        }

        public string Address { get; private set; }
        public int RowIndex { get; private set; }
        public int ColIndex { get; private set; }
        public Worksheet Parent { get; private set; }
        public Type DataType { get; private set; }

        public Cell(object value, string address, int rowIndex, int colIndex, Worksheet parentWorksheet, Type dataType)
        {
            _value = value;
            Address = address;
            RowIndex = rowIndex;
            ColIndex = colIndex;
            Parent = parentWorksheet;
            DataType = dataType;
        }

        public Cell Offset(int rowOffset, int colOffset)
        {
            int newRow = RowIndex + rowOffset;
            int newCol = ColIndex + colOffset;

            return Parent.GetCell(newRow, newCol);
        }
    }
}

﻿namespace Exceleration
{
    public class Cell
    {
        private int _rowIndex;
        private int _colIndex;

        public object Value { get; private set; }
        public string Address { get; private set; }
        public Worksheet Parent { get; private set; }
        public Type DataType { get; private set; }
        public int Row
        {
            get
            {
                return _rowIndex + 1;
            }
        }
        public int Column
        {
            get
            {
                return _colIndex + 1;
            }
        }
        public string ColumnLetter
        {
            get
            {
                return ConvertNumberToColLetter(_colIndex + 1);
            }
        }

        public Cell(object value, string address, int rowIndex, int colIndex, Worksheet parentWorksheet, Type dataType)
        {            
            _rowIndex = rowIndex;
            _colIndex = colIndex;
            Value = value;
            Address = address;
            Parent = parentWorksheet;
            DataType = dataType;
        }

        public Cell Offset(int rowOffset, int colOffset)
        {
            int newRow = _rowIndex + rowOffset;
            int newCol = _colIndex + colOffset;

            return Parent.GetCell(newRow, newCol);
        }

        public override string? ToString()
        {
            return Value.ToString();
        }

        public T To<T>(bool returnDefaultOnConversionError = true) where T : struct
        {
            try
            {
                return (T)Convert.ChangeType(Value, typeof(T));
            }
            catch
            {
                if (returnDefaultOnConversionError) return default;
                else throw new InvalidCastException($"Cannot convert cell value to type {typeof(T)}.");
            }
        }

        public T? ToNullable<T>(bool returnNullOnConversionError = true) where T : struct
        {
            if (string.IsNullOrEmpty(Value?.ToString()?.Trim()))
            {
                if (returnNullOnConversionError) return null;
                else return default;
            }

            try
            {
                return (T)Convert.ChangeType(Value, typeof(T));
            }

            catch
            {
                if (returnNullOnConversionError) return null;
                else return default;
            }
        }

        public bool IsParseable<T>() where T : struct
        {
            if (string.IsNullOrEmpty(Value?.ToString()?.Trim())) return false;

            try
            {
                Convert.ChangeType(Value, typeof(T));
                return true;
            }

            catch
            {
                return false;
            }
        }

        private static string ConvertNumberToColLetter(int colNumber)
        {
            int dividend = colNumber;
            string colLetter = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colLetter = Convert.ToChar(65 + modulo).ToString() + colLetter;
                dividend = ((dividend - modulo) / 26);
            }

            return colLetter;
        }
    }
}

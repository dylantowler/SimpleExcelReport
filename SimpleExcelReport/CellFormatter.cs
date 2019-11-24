*&^@#H﻿
using System;
using Microsoft.Office.Interop.Excel;

namespace SimpleExcelReport
{
    public abstract class CellFormatter<TRow, TProperty, TFormatType> : ICellFormatter<TRow>
    {
        private readonly Func<TRow, TProperty, TFormatType> _formatter;

        protected CellFormatter(Func<TRow, TProperty, TFormatType> formatter)
        {
            _formatter = formatter;
        }

        public void SetCellFormat(Range range, TRow row, object value)
        {
            TFormatType format = _formatter.Invoke(row, (TProperty) value);

            FormatObject(range, format);
        }

        protected abstract void FormatObject(Range range, object formatValue);
    }
}

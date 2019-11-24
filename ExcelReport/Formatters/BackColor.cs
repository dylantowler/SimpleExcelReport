using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace ExcelReport.Formatters
{
    public class BackColor<TRow, TProperty> : CellFormatter<TRow, TProperty, Color>
    {
        public BackColor(Func<TRow, TProperty, Color> formatter) : base(formatter)
        {
        }

        protected override void FormatObject(Range range, object formatValue)
        {
            Interior interior = range.Interior;
            interior.Color = formatValue;
        }
    }
}

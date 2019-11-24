using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Font = Microsoft.Office.Interop.Excel.Font;

namespace SimpleExcelReport.Formatters
{
    public class TextColor<TRow, TProperty> : CellFormatter<TRow, TProperty, Color>
    {
        public TextColor(Func<TRow, TProperty, Color> formatter) : base(formatter)
        {
        }

        protected override void FormatObject(Range range, object formatValue)
        {
            Font font = range.Font;
            font.Color = formatValue;
        }
    }
}

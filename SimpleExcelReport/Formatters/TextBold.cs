using System;
using Microsoft.Office.Interop.Excel;

namespace SimpleExcelReport.Formatters
{
    public class TextBold<TRow, TProperty> : CellFormatter<TRow, TProperty, bool>
    {
        public TextBold(Func<TRow, TProperty, bool> textBold) : base(textBold)
        {
            
        }

        protected override void FormatObject(Range range, object formatValue)
        {
            Font font = range.Font;
            font.Bold = formatValue;
        }
    }
}

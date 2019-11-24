using Microsoft.Office.Interop.Excel;

namespace ExcelReport.Formatters
{
    public class NumberFormat<TRow> : ICellFormatter<TRow>
    {
        private readonly string _formatString;

        public NumberFormat(string formatString)
        {
            _formatString = formatString;
        }

        public void SetCellFormat(Range range, TRow row, object value)
        {
            range.NumberFormat = _formatString;
        }
    }
}

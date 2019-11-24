using Microsoft.Office.Interop.Excel;

namespace ExcelReport
{
    public interface ICellFormatter<TRow>
    {
        void SetCellFormat(Range range, TRow row, object value);
    }
}
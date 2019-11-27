using Microsoft.Office.Interop.Excel;

namespace SimpleExcelReport
{
    public interface ICellFormatter<in TRow>
    {
        void SetCellFormat(Range range, TRow row, object value);
    }
}
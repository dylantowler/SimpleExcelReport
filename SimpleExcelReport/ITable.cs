using Microsoft.Office.Interop.Excel;

namespace SimpleExcelReport
{
    public interface ITable
    {
        void Write(Worksheet worksheet, int originX, int originY);
    }
}
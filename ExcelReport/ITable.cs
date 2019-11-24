using Microsoft.Office.Interop.Excel;

namespace ExcelReport
{
    public interface ITable
    {
        void Write(Worksheet worksheet, int originX, int originY);
    }
}
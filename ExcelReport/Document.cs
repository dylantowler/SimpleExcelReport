using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelReport
{
    public class Document : IDisposable
    {
        private void ReleaseUnmanagedResources()
        {
            _workbook.Close();
        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        private Application _excel;
        private Workbook _workbook;
        public readonly Worksheet _worksheet;

        public Document()
        {
            _excel = new Application();
            _workbook = _excel.Workbooks.Add(Type.Missing);
            _worksheet = _workbook.ActiveSheet as Worksheet;
        }

        public Worksheet Sheet => _worksheet;

        public void SaveAs(string filename)
        {
            _workbook.SaveAs(filename);
        }
    }
}

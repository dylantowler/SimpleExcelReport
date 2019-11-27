using System;
using Microsoft.Office.Interop.Excel;

namespace SimpleExcelReport
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

        private readonly Workbook _workbook;
        public readonly Worksheet Worksheet;

        public Document()
        {
            Application excel = new Application();
            _workbook = excel.Workbooks.Add(Type.Missing);
            Worksheet = _workbook.ActiveSheet as Worksheet;
        }

        public Worksheet Sheet => Worksheet;

        public void SaveAs(string filename)
        {
            _workbook.SaveAs(filename);
        }
    }
}

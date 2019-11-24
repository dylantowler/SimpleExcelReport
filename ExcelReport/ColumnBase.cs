using System.Collections.Generic;

namespace ExcelReport
{
    public abstract class ColumnBase<TRow>
    {
        public abstract object GetDisplayValue(TRow row);
        public abstract object GetValue(TRow row);
        public string Heading { get; set; }
        internal List<ICellFormatter<TRow>> Formatters = new List<ICellFormatter<TRow>>();
    }
}

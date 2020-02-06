using System;
using System.Collections.Generic;

namespace SimpleExcelReport
{
    public abstract class ColumnBase<TRow>
    {
        public abstract Type PropertyType { get; }
        public abstract bool Empty(TRow row);
        public abstract object GetDisplayValue(TRow row);
        public abstract object GetValue(TRow row);
        public string Heading { get; set; }
        public byte? Width { get; set; } = null;
        internal List<ICellFormatter<TRow>> Formatters = new List<ICellFormatter<TRow>>();
    }
}

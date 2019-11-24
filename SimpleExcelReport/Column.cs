using System;
using System.Drawing;
using System.Linq.Expressions;
using SimpleExcelReport.Formatters;

namespace SimpleExcelReport
{
    public class Column<TRow, TProperty> : ColumnBase<TRow>
    {
        private readonly Func<TRow, TProperty> _getter;
        private Func<TProperty, string> _stringConverter;

        public Column(Expression<Func<TRow, TProperty>> expression)
        {
            _getter = expression.Compile();

            if (typeof(TProperty).IsEnum)
            {
                _stringConverter = property => property.ToString();
            }

            SetDefaultHeading(expression);
        }

        private void SetDefaultHeading(Expression<Func<TRow, TProperty>> expression)
        {
            Heading = (expression.Body as MemberExpression)?.Member.Name;
        }

        public Column<TRow, TProperty> SetHeading(string heading)
        {
            Heading = heading;

            return this;
        }

        public Column<TRow, TProperty> TextBold(Func<TRow, TProperty, bool> textBold)
        {
            TextBold<TRow, TProperty> formatter = new TextBold<TRow, TProperty>(textBold);

            Formatters.Add(formatter);

            return this;
        }

        public Column<TRow, TProperty> BackColor(Func<TRow, TProperty, Color> backColor)
        {
            BackColor<TRow, TProperty> formatter = new BackColor<TRow, TProperty>(backColor);

            Formatters.Add(formatter);

            return this;
        }

        public Column<TRow, TProperty> TextColor(Func<TRow, TProperty, Color> textColor)
        {
            TextColor<TRow, TProperty> formatter = new TextColor<TRow, TProperty>(textColor);

            Formatters.Add(formatter);

            return this;
        }

        public Column<TRow, TProperty> AsString(Func<TProperty, string> toString)
        {
            _stringConverter = toString;

            return this;
        }

        public Column<TRow, TProperty> NumberFormat(string formatString)
        {
            NumberFormat<TRow> formatter =  new NumberFormat<TRow>(formatString);
            
            Formatters.Add(formatter);

            return this;
        }

        public override object GetDisplayValue(TRow row)
        {
            if (_stringConverter != null)
            {
                return _stringConverter(_getter(row));
            }

            return _getter(row);
        }

        public override object GetValue(TRow row) => _getter(row);
    }
}
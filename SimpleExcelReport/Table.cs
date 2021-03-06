﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Microsoft.Office.Interop.Excel;

namespace SimpleExcelReport
{
    public class Table<TRow>
    {
        private readonly List<Group<TRow>> _groups = new List<Group<TRow>>();
        private readonly IEnumerable<TRow> _dataSource;

        public bool HeadingBorder { get; set; } = false;
        public string Title { get; set; }
        public int ColumnCount => Columns.Count;
        public List<ColumnBase<TRow>> Columns { get; } = new List<ColumnBase<TRow>>();

        public Table(IEnumerable<TRow> dataSource)
        {
            _dataSource = dataSource ?? throw new ArgumentNullException(nameof(dataSource));
        }

        public Column<TRow, TProperty> AddColumn<TProperty>(Expression<Func<TRow, TProperty>> memberExpression)
        {
            Column<TRow, TProperty> column = new Column<TRow, TProperty>(memberExpression);

            Columns.Add(column);

            return column;
        }

        public int Write(Worksheet worksheet, int originX, int originY)
        {
            // Looks like the excel interop stuff is 1 based!
            if (originX < 1) throw new ArgumentOutOfRangeException(nameof(originX), "Excel cell access is 1 based.");
            if (originY < 1) throw new ArgumentOutOfRangeException(nameof(originY), "Excel cell access is 1 based.");

            int y = originY;

            SetColumnWidths(worksheet, originX);

            CreateHeader(worksheet, originX, ref y);

            CreateRows(worksheet, originX, ref y);

            GroupBorders(worksheet, originX, originY, y - 1);

            return y;
        }

        private void SetColumnWidths(Worksheet worksheet, int originX)
        {
            int x = originX;

            foreach (ColumnBase<TRow> column in Columns)
            {
                if (column.Width != null)
                {
                    var range = worksheet.Cells[1, x];
                    range.ColumnWidth = column.Width;
                }

                x++;
            }
        }

        private (int left, int right) GroupSpan(Group<TRow> group)
        {
            int left = group.Columns.Min(c => Columns.IndexOf(c));
            int right = group.Columns.Max(c => Columns.IndexOf(c));

            return (left, right);
        }

        private Range GroupRange(Worksheet worksheet, Group<TRow> group, int xOffset, int top, int bottom)
        {
            (int left, int right) = GroupSpan(group);

            Range range = worksheet.Range[worksheet.Cells[top, left + xOffset], worksheet.Cells[bottom, right + xOffset]];

            return range;
        }

        private void GroupBorders(Worksheet worksheet, int originX, int originY, int y)
        {
            foreach (Group<TRow> group in _groups.Where(g => g.HasBorder))
            {
                Range range = GroupRange(worksheet, group, originX, originY, y);

                range.BorderAround();
            }
        }

        /// <summary>
        /// If you read code in this method (or methods it calls) and think I'm unnecessarily using
        /// temp variables read the following two links (has to do with dynamic Excel interop):
        /// https://stackoverflow.com/questions/11740856/changing-font-size-of-one-cell-in-excel-using-c-sharp
        /// https://stackoverflow.com/questions/158706/how-do-i-properly-clean-up-excel-interop-objects/159419#159419 
        /// </summary>
        private void CreateHeader(Worksheet worksheet, int originX, ref int y)
        {
            int x = originX;

            bool groupHeading = false;

            if (!string.IsNullOrWhiteSpace(Title))
            {
                Range range = worksheet.Range[worksheet.Cells[y, originX], worksheet.Cells[y, originX + Columns.Count - 1]];
                range.Merge();
                range.Value = Title;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                Font font = range.Font;
                font.Bold = true;

                if (HeadingBorder)
                {
                    range.BorderAround();
                }

                y++;
            }

            foreach (Group<TRow> group in _groups.Where(g => g.HasHeading))
            {
                Range range = GroupRange(worksheet, group, originX, y, y);
                range.Merge();
                range.Value = group.Heading;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                Font font = range.Font;
                font.Bold = true;

                if (group.HasBorder)
                {
                    range.BorderAround();
                }

                groupHeading = true;
            }

            if (groupHeading)
            {
                y++;
            }

            foreach (ColumnBase<TRow> tableColumn in Columns)
            {
                // Set header value
                dynamic cell = worksheet.Cells[y, x];
                cell.Value = tableColumn.Heading;

                Range range = worksheet.Range[worksheet.Cells[y, x], worksheet.Cells[y, x]];
                Font font = range.Font;
                font.Bold = true;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                x++;
            }

            if (HeadingBorder)
            {
                Range range = worksheet.Range[worksheet.Cells[y, originX], worksheet.Cells[y, x - 1]];
                range.BorderAround();
            }

            y++;
        }

        private void CreateRows(Worksheet worksheet, int originX, ref int y)
        {
            foreach (TRow row in _dataSource)
            {
                int x = originX;

                foreach (ColumnBase<TRow> tableColumn in Columns)
                {
                    dynamic cell = worksheet.Cells[y, x];
                    if (!tableColumn.Empty(row))
                    {
                        cell.Value = tableColumn.GetDisplayValue(row);
                    }

                    Range range = worksheet.Range[worksheet.Cells[y, x], worksheet.Cells[y, x]];

                    foreach (ICellFormatter<TRow> formatter in tableColumn.Formatters)
                    {
                        formatter.SetCellFormat(range, row, tableColumn.GetValue(row));
                    }

                    x++;
                }

                y++;
            }
        }

        public Group<TRow> Group(ColumnBase<TRow>[] columns)
        {
            if (!columns.Where(c => c != null).Select(c => Columns.IndexOf(c)).Contiguous())
            {
                throw new ArgumentException("Column grouping must contain contiguous columns.", nameof(columns));
            }

            Group<TRow> group = new Group<TRow>(columns.Where(c => c != null).ToArray());

            _groups.Add(group);

            return group;
        }
    }
}

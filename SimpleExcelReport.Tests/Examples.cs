using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using NUnit.Framework;

namespace SimpleExcelReport.Tests
{
    public class Examples
    {
        private enum Sex
        {
            M,
            F
        }

        private struct Grades
        {
            public double Math;
            public double English;
            public double History;
            public double Economics;
            public double Science;
        }

        private class Student
        {
            public string Name { get; set; }
            public Sex Sex { get; set; }
            public Grades Grades { get; set; }
        }

        [Test]
        [Category("RequiresExcel")]
        public void LaunchTestReport()
        {
            List<Student> students = TestData();

            Table<Student> studentTable = new Table<Student>(students);
            studentTable.AddColumn(s => s.Name).SetHeading("Name");
            studentTable.AddColumn(s => s.Sex.ToString()).SetHeading("Gender");
            studentTable.AddColumn(s => s.Grades.Math).SetHeading("Math");
            studentTable.AddColumn(s => s.Grades.English).SetHeading("English");
            studentTable.AddColumn(s => s.Grades.Science).SetHeading("Science");
            studentTable.AddColumn(s => s.Grades.Economics).SetHeading("Economics");
            studentTable.AddColumn(s => s.Grades.History).SetHeading("History");

            string tempFilename = Path.GetTempFileName();
            tempFilename = Path.ChangeExtension(tempFilename, ".xlsx");

            using (Document excelDocument = new Document())
            {
                studentTable.Write(excelDocument.Sheet, 1, 1);
                excelDocument.SaveAs(tempFilename);
            }

            System.Diagnostics.Process.Start(tempFilename);
        }

        private string SexAsString(Sex sex)
        {
            switch (sex)
            {
                case Sex.M:
                    return "Boy";
                case Sex.F:
                    return "Girl";
                default:
                    throw new ArgumentOutOfRangeException(nameof(sex), sex, null);
            }
        }

        [Test]
        [Category("RequiresExcel")]
        public void LaunchSimpleTestReport()
        {
            List<Student> students = TestData();

            Table<Student> studentTable = new Table<Student>(students);
            var name = studentTable.AddColumn(s => s.Name);
            var sex = studentTable.AddColumn(s => s.Sex).SetHeading("Gender").AsString(SexAsString).BackColor((s, g) => g == Sex.F ? Color.Pink : Color.Aqua);
            var math = studentTable.AddColumn(s => s.Grades.Math).TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80).NumberFormat("#0.00\"%\"");
            var english = studentTable.AddColumn(s => s.Grades.English).TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80).NumberFormat("#0.00\"%\"");
            var science = studentTable.AddColumn(s => s.Grades.Science).TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80).NumberFormat("#0.00\"%\"");
            var economics = studentTable.AddColumn(s => s.Grades.Economics).TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80).NumberFormat("#0.00\"%\"");
            var history = studentTable.AddColumn(s => s.Grades.History).TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80).NumberFormat("#0.00\"%\"");

            studentTable.Group(new ColumnBase<Student>[]{ name, sex }).SetHeading("Student").Border();
            studentTable.Group(new ColumnBase<Student>[] { math, english, science, economics, history }).SetHeading("Grades").Border();

            studentTable.HeadingBorder = true;

            string tempFilename = Path.GetTempFileName();
            tempFilename = Path.ChangeExtension(tempFilename, ".xlsx");

            using (Document excelDocument = new Document())
            {
                studentTable.Write(excelDocument.Sheet, 2, 2);
                excelDocument.SaveAs(tempFilename);
            }

            System.Diagnostics.Process.Start(tempFilename);
        }

        private static List<Student> TestData() => new List<Student>
        {
            new Student
            {
                Name = "Nicholas",
                Sex = Sex.M,
                Grades = new Grades
                {
                    Math = 85,
                    English = 88,
                    History = 75,
                    Economics = 75,
                    Science = 63
                }
            },

            new Student
            {
                Name = "Ian",
                Sex = Sex.F,
                Grades = new Grades
                {
                    Math = 85,
                    English = 88,
                    History = 75,
                    Economics = 75,
                    Science = 63
                }
            },

            new Student
            {
                Name = "Alec",
                Sex = Sex.M,
                Grades = new Grades
                {
                    Math = 40,
                    English = 35,
                    History = 42,
                    Economics = 24,
                    Science = 44
                }
            },

            new Student
            {
                Name = "Abagail",
                Sex = Sex.F,
                Grades = new Grades
                {
                    Math = 85.00000001,
                    English = 88.45765786,
                    History = 75.11,
                    Economics = 75.4567,
                    Science = 63.005415
                }
            },
        };
    }
}

[![Build status](https://ci.appveyor.com/api/projects/status/onyfc7crpqfis72h?svg=true)](https://ci.appveyor.com/project/DylanTowler/simpleexcelreport)
[![Nuget](https://img.shields.io/nuget/v/SimpleExcelReport)](https://www.nuget.org/packages/SimpleExcelReport)

[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=SimpleExcelReport&metric=alert_status)](https://sonarcloud.io/dashboard?id=SimpleExcelReport)
[![Bugs](https://sonarcloud.io/api/project_badges/measure?project=SimpleExcelReport&metric=bugs)](https://sonarcloud.io/dashboard?id=SimpleExcelReport)
[![Code Smells](https://sonarcloud.io/api/project_badges/measure?project=SimpleExcelReport&metric=code_smells)](https://sonarcloud.io/dashboard?id=SimpleExcelReport)
[![Duplicated Lines (%)](https://sonarcloud.io/api/project_badges/measure?project=SimpleExcelReport&metric=duplicated_lines_density)](https://sonarcloud.io/dashboard?id=SimpleExcelReport)

# Welcome to the SimpleExcelReport project

SimpleExcelReport is a .NET component that allows you to quickly create tabular Excel reports using a fluent API and conditional formatting.

It currently relies on Microsoft Excel Interop but eventually I would like to use something to write the xslx directly so that it can run on servers (and target .NET Standard).  If you would like to help with that then please fork!

## Adding SimpleExcelReport to your project

#### NuGet
You can use NuGet to quickly add SimpleExcelReport to your project. Just search for `SimpleExcelReport` and install the package.

## Super simple example

### Code
```c#
List<Student> students = TestData();

Table<Student> studentTable = new Table<Student>(students);
studentTable.AddColumn(s => s.Name);
studentTable.AddColumn(s => s.Sex);
studentTable.AddColumn(s => s.Grades.Math);
studentTable.AddColumn(s => s.Grades.English);
studentTable.AddColumn(s => s.Grades.Science);
studentTable.AddColumn(s => s.Grades.Economics);
studentTable.AddColumn(s => s.Grades.History);
```

### Result
![alt text](https://github.com/dylantowler/SimpleExcelReport/blob/master/ReadMeImages/SuperSimple.PNG)

## Custom display and conditional formatting

### Code
```c#
List<Student> students = TestData();

Table<Student> studentTable = new Table<Student>(students);
studentTable.AddColumn(s => s.Name);
studentTable.AddColumn(s => s.Sex).AsString(sex => sex == Sex.F ? "Girl" : "Boy").BackColor((s, g) => g == Sex.F ? Color.Pink : Color.Aqua);
studentTable.AddColumn(s => s.Grades.Math);
studentTable.AddColumn(s => s.Grades.English);
studentTable.AddColumn(s => s.Grades.Science);
studentTable.AddColumn(s => s.Grades.Economics);
studentTable.AddColumn(s => s.Grades.History);
```

### Result
![alt text](https://github.com/dylantowler/SimpleExcelReport/blob/master/ReadMeImages/CustomStringDisplay.PNG)

## Excel formatting, borders, group headings and more...
### Code
```c#
List<Student> students = TestData();

Table<Student> studentTable = new Table<Student>(students);
studentTable.HeadingBorder = true;

var name = studentTable.AddColumn(s => s.Name);
var sex = studentTable.AddColumn(s => s.Sex).SetHeading("Gender").AsString(SexAsString)
	.BackColor((s, g) => g == Sex.F ? Color.Pink : Color.Aqua);
var math = studentTable.AddColumn(s => s.Grades.Math).TextColor((s, g) => g < 50 ? Color.Red : Color.Green)
	.TextBold((s, g) => g > 80).NumberFormat("#0.00\"%\"");
var english = studentTable.AddColumn(s => s.Grades.English)
	.TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80)
	.NumberFormat("#0.00\"%\"");
var science = studentTable.AddColumn(s => s.Grades.Science)
	.TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80)
	.NumberFormat("#0.00\"%\"");
var economics = studentTable.AddColumn(s => s.Grades.Economics)
	.TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80)
	.NumberFormat("#0.00\"%\"");
var history = studentTable.AddColumn(s => s.Grades.History)
	.TextColor((s, g) => g < 50 ? Color.Red : Color.Green).TextBold((s, g) => g > 80)
	.NumberFormat("#0.00\"%\"");

studentTable.Group(new ColumnBase<Student>[] {name, sex}).SetHeading("Student").Border();
studentTable.Group(new ColumnBase<Student>[] {math, english, science, economics, history}).SetHeading("Grades").Border();
```

### Result
![alt text](https://github.com/dylantowler/SimpleExcelReport/blob/master/ReadMeImages/NumberFormatCustomHeadingBordersGroupHeadingsEtc.PNG)


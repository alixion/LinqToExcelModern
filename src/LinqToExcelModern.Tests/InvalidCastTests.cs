﻿namespace LinqToExcelModern.Tests;

[Author("Alberto Chvaicer", "achvaicer@gmail.com")]
[Category("Integration")]
[TestFixture]
public class InvalidCastTests
{
    private IExcelQueryFactory _factory;
    private string _excelFileName;

    [OneTimeSetUp]
    public void fs()
    {
        var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
        _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");

    }

    [SetUp]
    public void s()
    {
        _factory = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
    }

    [Test]
    public void invalid_number_cast_with_header()
    {
        Assert.That(() => (from x in ExcelQueryFactory.Worksheet<Company>("Invalid Cast", _excelFileName, new LogManagerFactory())
            select x).ToList(), Throws.TypeOf<LinqToExcelModern.Exceptions.ExcelException>(), "Error on row 8 and column name 'EmployeeCount'.");
    }
}
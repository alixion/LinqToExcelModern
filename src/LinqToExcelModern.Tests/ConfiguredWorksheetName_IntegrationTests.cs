﻿using System.Data;

namespace LinqToExcelModern.Tests;

[Author("Paul Yoder", "paulyoder@gmail.com")]
[Category("Integration")]
[TestFixture]
public class ConfiguredWorksheetName_IntegrationTests : SQLLogStatements_Helper
{
    private string _excelFileName;

    [OneTimeSetUp]
    public void fs()
    {
        var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
        _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        InstantiateLogger();
    }

    [SetUp]
    public void s()
    {
        ClearLogEvents();
    }

    [Test]
    public void data_is_read_from_correct_worksheet()
    {
        var companies = from c in ExcelQueryFactory.Worksheet<Company>("More Companies", _excelFileName, new LogManagerFactory())
            select c;

        Assert.AreEqual(3, companies.ToList().Count);
    }

    [Test]
    public void worksheetIndex_of_2_uses_third_table_name_orderedby_name()
    {
        var companies = (from c in ExcelQueryFactory.Worksheet<Company>(3, _excelFileName, new LogManagerFactory())
            select c).ToList();

        var expectedSql = "SELECT * FROM [More Companies$]";
        Assert.AreEqual(expectedSql, GetSQLStatement(), "SQL Statement");
    }

    [Test]
    public void worksheetIndex_too_high_throws_exception()
    {
        Assert.That(() => from c in ExcelQueryFactory.Worksheet<Company>(100, _excelFileName, new LogManagerFactory())
                select c,
            Throws.TypeOf<DataException>());
    }
}
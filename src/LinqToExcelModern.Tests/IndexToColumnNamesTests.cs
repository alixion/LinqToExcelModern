namespace LinqToExcelModern.Tests;

[TestFixture]
[Category("Unit")]
[Author("Alberto Chvaicer")]
public class IndexToColumnNamesTests
{
    [Test]
    public void valid_excel_column_indexes()
    {
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(1), "A");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(2), "B");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(3), "C");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(25), "Y");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(26), "Z");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(27), "AA");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(28), "AB");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(51), "AY");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(52), "AZ");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(53), "BA");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(54), "BB");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(701), "ZY");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(702), "ZZ");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(703), "AAA");
        Assert.AreEqual(LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(704), "AAB");
    }

    [Test]
    public void negative_integer_should_throw_exception()
    {
        Assert.That(() => LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(-1), Throws.ArgumentException, "Index should be a positive integer");
    }

    [Test]
    public void zero_should_throw_exception()
    {
        Assert.That(() => LinqToExcelModern.Query.ExcelUtilities.ColumnIndexToExcelColumnName(-1), Throws.ArgumentException, "Index should be a positive integer");
    }
}
using LinqToExcelModern.Domain;
using LinqToExcelModern.Exceptions;

namespace LinqToExcelModern.Tests;

class CompanyGoodWithAllowFieldTypeConversionExceptions : IAllowFieldTypeConversionExceptions
{
    public CompanyGoodWithAllowFieldTypeConversionExceptions()
    {
        FieldTypeConversionExceptions = new List<ExcelException>();
    }

    public string Name { get; set; }
    public string CEO { get; set; }
    public int EmployeeCount { get; set; }

    public IList<ExcelException> FieldTypeConversionExceptions { get; private set; }
}
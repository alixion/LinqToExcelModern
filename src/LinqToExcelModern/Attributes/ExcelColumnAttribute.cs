using System;

namespace LinqToExcelModern.Attributes
{
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        private readonly string _columnName;

        public ExcelColumnAttribute(string columnName)
        {
            _columnName = columnName;
        }

        public string ColumnName
        {
            get { return _columnName; }
        }

    }
}

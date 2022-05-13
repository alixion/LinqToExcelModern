﻿namespace LinqToExcelModern.Query;

internal class ExcelQueryConstructorArgs
{
    internal bool Lazy { get; set; }

    public ExcelQueryConstructorArgs()
    {
        OleDbServices = OleDbServices.AllServices;
    }

    internal string FileName { get; set; }
    internal Dictionary<string, string> ColumnMappings { get; set; }
    internal Dictionary<string, Func<string, object>> Transformations { get; set; }
    internal StrictMappingType? StrictMapping { get; set; }
    internal bool UsePersistentConnection { get; set; }
    internal TrimSpacesType TrimSpaces { get; set; }
    internal bool ReadOnly { get; set; }
    internal OleDbServices OleDbServices { get; set; }
    internal int CodePageIdentifier { get; set; }
    internal bool SkipEmptyRows { get; set; }
}
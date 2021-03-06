using System.Data;
using System.Data.OleDb;
using LinqToExcelModern.Extensions;

namespace LinqToExcelModern.Query;

public static class ExcelUtilities
{
    internal static string GetConnectionString(ExcelQueryArgs args)
    {
        var connString = "";
        var fileNameLower = args.FileName.ToLower();

        if (fileNameLower.EndsWith("xlsx") ||
            fileNameLower.EndsWith("xlsm"))
        {
            connString = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};OLE DB Services={1:d};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                args.FileName,
                args.OleDbServices);
        }
        else if (fileNameLower.EndsWith("xlsb"))
        {
            connString = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};OLE DB Services={1:d};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                args.FileName,
                args.OleDbServices);
        }
        else if (fileNameLower.EndsWith("csv"))
        {
            connString = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};OLE DB Services={1:d};Extended Properties=""text;Excel 12.0;HDR=YES;IMEX=1""",
                Path.GetDirectoryName(args.FileName),
                args.OleDbServices);                
        }
        else
        {
            connString = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};OLE DB Services={1:d};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                args.FileName,
                args.OleDbServices);
        }

        if (args.NoHeader)
            connString = connString.Replace("HDR=YES", "HDR=NO");

        if (args.ReadOnly)
            connString = connString.Replace("IMEX=1", "IMEX=1;READONLY=TRUE");

        if(args.CodePageIdentifier > 0)
            connString = connString.Replace("IMEX=1",string.Format ("IMEX=1;CharacterSet={0:000}",args.CodePageIdentifier));

        return connString;
    }

    internal static IEnumerable<string> GetWorksheetNames(string fileName)
    {
        return GetWorksheetNames(fileName, new ExcelQueryArgs());
    }

    internal static IEnumerable<string> GetWorksheetNames(string fileName, ExcelQueryArgs origArgs)
    {
        var args = new ExcelQueryArgs(origArgs);
        args.FileName = fileName;
        args.ReadOnly = true;
        return GetWorksheetNames(args);
    }

    internal static OleDbConnection GetConnection(ExcelQueryArgs args)
    {
        if (args.UsePersistentConnection)
        {
            if (args.PersistentConnection == null)
                args.PersistentConnection = new OleDbConnection(GetConnectionString(args));

            return args.PersistentConnection;
        }

        return new OleDbConnection(GetConnectionString(args));
    }

    internal static IEnumerable<string> GetWorksheetNames(ExcelQueryArgs args)
    {
        var worksheetNames = new List<string>();

        var conn = GetConnection(args);
        try
        {
            if (conn.State == ConnectionState.Closed)
                conn.Open();

            var excelTables = conn.GetOleDbSchemaTable(
                OleDbSchemaGuid.Tables,
                new Object[] { null, null, null, "TABLE" });

            worksheetNames.AddRange(
                from DataRow row in excelTables.Rows
                where IsTable(row)
                let tableName = row["TABLE_NAME"].ToString()
                    .RegexReplace("(^'|'$)", "")
                    .RegexReplace(@"\$$", "")
                    .Replace("''", "'")
                where IsNotBuiltinTable(tableName)
                select tableName);

            excelTables.Dispose();
        }
        finally
        {
            if (!args.UsePersistentConnection)
                conn.Dispose();
        }

        return worksheetNames;
    }

    internal static bool IsTable(DataRow row)
    {
        var tableName = row["TABLE_NAME"].ToString();

        return tableName.EndsWith("$") || (tableName.StartsWith("'") && tableName.EndsWith("$'"));
    }

    internal static bool IsNamedRange(DataRow row)
    {
        var tableName = row["TABLE_NAME"].ToString();

        return (tableName.Contains("$") && !tableName.EndsWith("$") && !tableName.EndsWith("$'")) || !tableName.Contains("$");
    }

    internal static bool IsWorkseetScopedNamedRange(DataRow row)
    {
        var tableName = row["TABLE_NAME"].ToString();

        return IsNamedRange(row) && tableName.Contains("$");
    }

    internal static bool IsNotBuiltinTable(string tableName)
    {
        return !tableName.Contains("FilterDatabase") && !tableName.Contains("Print_Area");
    }

    internal static IEnumerable<string> GetColumnNames(string worksheetName, string fileName)
    {
        return GetColumnNames(worksheetName, fileName, new ExcelQueryArgs());
    }

    internal static IEnumerable<string> GetColumnNames(string worksheetName, string fileName, ExcelQueryArgs origArgs)
    {
        var args = new ExcelQueryArgs(origArgs);
        args.WorksheetName = worksheetName;
        args.FileName = fileName;
        return GetColumnNames(args);
    }

    internal static IEnumerable<string> GetColumnNames(string worksheetName, string namedRange, string fileName)
    {
        return GetColumnNames(worksheetName, namedRange, fileName, new ExcelQueryArgs());
    }

    internal static IEnumerable<string> GetColumnNames(string worksheetName, string namedRange, string fileName, ExcelQueryArgs origArgs)
    {
        var args = new ExcelQueryArgs(origArgs);
        args.WorksheetName = worksheetName;
        args.NamedRangeName = namedRange;
        args.FileName = fileName;
        return GetColumnNames(args);
    }

    internal static IEnumerable<string> GetColumnNames(ExcelQueryArgs args)
    {
        var columns = new List<string>();
        var conn = GetConnection(args);
        try
        {
            if (conn.State == ConnectionState.Closed)
                conn.Open();

            using (var command = conn.CreateCommand())
            {
                command.CommandText = string.Format("SELECT TOP 1 * FROM [{0}{1}]", string.Format("{0}{1}", args.WorksheetName, "$"), args.NamedRangeName);
                var data = command.ExecuteReader();
                columns.AddRange(GetColumnNames(data));
            }
        }
        finally
        {
            if (!args.UsePersistentConnection)
                conn.Dispose();
        }

        return columns;
    }

    internal static IEnumerable<string> GetColumnNames(IDataReader data)
    {
        var columns = new List<string>();
        var sheetSchema = data.GetSchemaTable();
        foreach (DataRow row in sheetSchema.Rows)
            columns.Add(row["ColumnName"].ToString());

        return columns;
    }

    internal static IEnumerable<string> GetNamedRanges(string fileName, string worksheetName)
    {
        return GetNamedRanges(fileName, worksheetName, new ExcelQueryArgs());
    }

    internal static IEnumerable<string> GetNamedRanges(string fileName)
    {
        return GetNamedRanges(fileName, new ExcelQueryArgs());
    }

    internal static IEnumerable<string> GetNamedRanges(string fileName, ExcelQueryArgs origArgs)
    {
        var args = new ExcelQueryArgs(origArgs);
        args.FileName = fileName;
        args.ReadOnly = true;
        return GetNamedRanges(args);
    }

    internal static IEnumerable<string> GetNamedRanges(string fileName, string worksheetName, ExcelQueryArgs origArgs)
    {
        var args = new ExcelQueryArgs(origArgs);
        args.FileName = fileName;
        args.WorksheetName = worksheetName;
        args.ReadOnly = true;
        return GetNamedRanges(args);
    }

    internal static IEnumerable<string> GetNamedRanges(ExcelQueryArgs args)
    {
        var namedRanges = new List<string>();

        var conn = GetConnection(args);
        try
        {
            if (conn.State == ConnectionState.Closed)
                conn.Open();

            var excelTables = conn.GetOleDbSchemaTable(
                OleDbSchemaGuid.Tables,
                new Object[] { null, null, null, "TABLE" });

            namedRanges.AddRange(
                from DataRow row in excelTables.Rows
                where IsNamedRange(row)
                      && (!string.IsNullOrEmpty(args.WorksheetName) ? row["TABLE_NAME"].ToString().StartsWith(args.WorksheetName) : !IsWorkseetScopedNamedRange(row))
                let tableName = row["TABLE_NAME"].ToString()
                    .Replace("''", "'")
                where IsNotBuiltinTable(tableName)
                select tableName.Split('$').Last());

            excelTables.Dispose();
        }
        finally
        {
            if (!args.UsePersistentConnection)
                conn.Dispose();
        }

        return namedRanges;
    }

    public static string ColumnIndexToExcelColumnName(int index)
    {
        if (index < 1) throw new ArgumentException("Index should be a positive integer");
        var quotient = (--index) / 26;

        if (quotient > 0)
        {
            return ColumnIndexToExcelColumnName(quotient) + (char)((index % 26) + 65);
        }
        else

        {
            return ((char)((index % 26) + 65)).ToString();
        }
    }

}
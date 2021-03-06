using Remotion.Linq;
using System.Data;
using System.Data.OleDb;
using System.Reflection;
using Remotion.Linq.Clauses.ResultOperators;
using LinqToExcelModern.Extensions;
using System.Text.RegularExpressions;
using System.Text;
using LinqToExcelModern.Domain;
using LinqToExcelModern.Logging;

namespace LinqToExcelModern.Query;

internal class ExcelQueryExecutor : IQueryExecutor
{
    private readonly ILogManagerFactory _logManagerFactory;
    private readonly ILogProvider _log;
    private readonly ExcelQueryArgs _args;

    internal ExcelQueryExecutor(ExcelQueryArgs args, ILogManagerFactory logManagerFactory)
    {
        ValidateArgs(args);
        _args = args;

        if (logManagerFactory != null) {
            _logManagerFactory = logManagerFactory;
            _log = _logManagerFactory.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        }

        if (_log != null && _log.IsDebugEnabled == true)
            _log.DebugFormat("Connection String: {0}", ExcelUtilities.GetConnection(args).ConnectionString);

        GetWorksheetName();
    }

    private void ValidateArgs(ExcelQueryArgs args)
    {
        if (_log != null && _log.IsDebugEnabled == true)
            _log.DebugFormat("ExcelQueryArgs = {0}", args);

        if (args.FileName == null)
            throw new ArgumentNullException("FileName", "FileName property cannot be null.");

        if (!String.IsNullOrEmpty(args.StartRange) &&
            !Regex.Match(args.StartRange, "^[a-zA-Z]{1,3}[0-9]{1,7}$").Success)
            throw new ArgumentException(string.Format(
                "StartRange argument '{0}' is invalid format for cell name", args.StartRange));

        if (!String.IsNullOrEmpty(args.EndRange) &&
            !Regex.Match(args.EndRange, "^[a-zA-Z]{1,3}[0-9]{1,7}$").Success)
            throw new ArgumentException(string.Format(
                "EndRange argument '{0}' is invalid format for cell name", args.EndRange));

        if (args.NoHeader &&
            !String.IsNullOrEmpty(args.StartRange) &&
            args.FileName.ToLower().Contains(".csv"))
            throw new ArgumentException("Cannot use WorksheetRangeNoHeader on csv files");
    }

    /// <summary>
    /// Executes a query with a scalar result, i.e. a query that ends with a result operator such as Count, Sum, or Average.
    /// </summary>
    public T ExecuteScalar<T>(QueryModel queryModel)
    {
        return ExecuteSingle<T>(queryModel, false);
    }

    /// <summary>
    /// Executes a query with a single result object, i.e. a query that ends with a result operator such as First, Last, Single, Min, or Max.
    /// </summary>
    public T ExecuteSingle<T>(QueryModel queryModel, bool returnDefaultWhenEmpty)
    {
        var results = ExecuteCollection<T>(queryModel);

        foreach (var resultOperator in queryModel.ResultOperators)
        {
            if (resultOperator is LastResultOperator)
                return results.LastOrDefault();
        }

        return (returnDefaultWhenEmpty) ?
            results.FirstOrDefault() :
            results.First();
    }

    /// <summary>
    /// Executes a query with a collection result.
    /// </summary>
    public IEnumerable<T> ExecuteCollection<T>(QueryModel queryModel)
    {
        var sql = GetSqlStatement(queryModel);
        LogSqlStatement(sql);

        IEnumerable<T> returnResults = Project<T>(queryModel, sql);

        if (!_args.Lazy)
        {
            returnResults = returnResults.ToList();
        }

        foreach (var resultOperator in queryModel.ResultOperators)
        {
            if (resultOperator is ReverseResultOperator)
                returnResults = returnResults.Reverse();
            if (resultOperator is SkipResultOperator)
                returnResults = returnResults.Skip(resultOperator.Cast<SkipResultOperator>().GetConstantCount());
        }

        return returnResults;
    }

    private IEnumerable<T> Project<T>(QueryModel queryModel, SqlParts sql)
    {
        Func<object, T> projector = null;

        var objectResults = GetDataResults(sql, queryModel);
        foreach (var row in objectResults)
        {
            if (projector == null)
            {
                projector = GetSelectProjector<T>(row, queryModel);
            }

            yield return projector(row);
        }
    }

    protected Func<object, T> GetSelectProjector<T>(object firstResult, QueryModel queryModel)
    {
        Func<object, T> projector = (result) => result.Cast<T>();
        if (ShouldBuildResultObjectMapping<T>(firstResult, queryModel))
        {
            var proj = ProjectorBuildingExpressionTreeVisitor.BuildProjector<T>(queryModel.SelectClause.Selector);
            projector = (result) => proj(new ResultObjectMapping(queryModel.MainFromClause, result));
        }
        return projector;
    }

    protected bool ShouldBuildResultObjectMapping<T>(object firstResult, QueryModel queryModel)
    {
        var ignoredResultOperators = new List<Type>()
        {
            typeof (MaxResultOperator),
            typeof (CountResultOperator),
            typeof (LongCountResultOperator),
            typeof (MinResultOperator),
            typeof (SumResultOperator)
        };

        var resultType = firstResult?.GetType();

        return (firstResult != null &&
                resultType != typeof(T) &&
                resultType != Nullable.GetUnderlyingType(typeof(T)) &&
                !queryModel.ResultOperators.Any(x => ignoredResultOperators.Contains(x.GetType())));
    }

    protected SqlParts GetSqlStatement(QueryModel queryModel)
    {
        var sqlVisitor = new SqlGeneratorQueryModelVisitor(_args);
        sqlVisitor.VisitQueryModel(queryModel);
        return sqlVisitor.SqlStatement;
    }

    private void GetWorksheetName()
    {
        if (_args.FileName.ToLower().EndsWith("csv"))
            _args.WorksheetName = Path.GetFileName(_args.FileName);
        else if (_args.WorksheetIndex.HasValue)
        {
            var worksheetNames = ExcelUtilities.GetWorksheetNames(_args);
            if (_args.WorksheetIndex.Value < worksheetNames.Count())
                _args.WorksheetName = worksheetNames.ElementAt(_args.WorksheetIndex.Value);
            else
                throw new DataException("Worksheet Index Out of Range");
        }
        else if (String.IsNullOrEmpty(_args.WorksheetName) && String.IsNullOrEmpty(_args.NamedRangeName))
        {
            _args.WorksheetName = "Sheet1";
        }
    }

    /// <summary>
    /// Executes the sql query and returns the data results
    /// </summary>
    /// <typeparam name="T">Data type in the main from clause (queryModel.MainFromClause.ItemType)</typeparam>
    /// <param name="queryModel">Linq query model</param>
    protected IEnumerable<object> GetDataResults(SqlParts sql, QueryModel queryModel)
    {
        IEnumerable<object> results;
        OleDbDataReader data = null;

        var conn = ExcelUtilities.GetConnection(_args);
        var command = conn.CreateCommand();
        try
        {
            if (conn.State == ConnectionState.Closed)
                conn.Open();

            command.CommandText = sql.ToString();
            command.Parameters.AddRange(sql.Parameters.ToArray());
            try { data = command.ExecuteReader(); }
            catch (OleDbException e)
            {
                if (e.Message.Contains(_args.WorksheetName))
                    throw new DataException(
                        string.Format("'{0}' is not a valid worksheet name in file {3}. Valid worksheet names are: '{1}'. Error received: {2}",
                            _args.WorksheetName, string.Join("', '", ExcelUtilities.GetWorksheetNames(_args.FileName).ToArray()), e.Message, _args.FileName), e);
                if (!CheckIfInvalidColumnNameUsed(sql))
                    throw e;
            }

            var columns = ExcelUtilities.GetColumnNames(data);
            LogColumnMappingWarnings(columns);
            if (columns.Count() == 1 && columns.First() == "Expr1000")
                results = GetScalarResults(data);
            else if (queryModel.MainFromClause.ItemType == typeof(Row))
                results = GetRowResults(data, columns);
            else if (queryModel.MainFromClause.ItemType == typeof(RowNoHeader))
                results = GetRowNoHeaderResults(data);
            else
                results = GetTypeResults(data, columns, queryModel);

            foreach (var result in results)
            {
                yield return result;
            }
        }
        finally
        {
            command.Dispose();

            if (!_args.UsePersistentConnection)
            {
                conn.Dispose();
                _args.PersistentConnection = null;
            }
        }
    }

    /// <summary>
    /// Logs a warning for any property to column mappings that do not exist in the excel worksheet
    /// </summary>
    /// <param name="Columns">List of columns in the worksheet</param>
    private void LogColumnMappingWarnings(IEnumerable<string> columns)
    {
        foreach (var kvp in _args.ColumnMappings)
        {
            if (!columns.Contains(kvp.Value))
            {
                if (_log != null)
                    _log.WarnFormat("'{0}' column that is mapped to the '{1}' property does not exist in the '{2}' worksheet",
                        kvp.Value, kvp.Key, _args.WorksheetName);
            }
        }
    }

    private bool CheckIfInvalidColumnNameUsed(SqlParts sql)
    {
        var usedColumns = sql.ColumnNamesUsed;
        var tableColumns = ExcelUtilities.GetColumnNames(_args.WorksheetName, _args.NamedRangeName, _args.FileName);
        foreach (var column in usedColumns)
        {
            if (!tableColumns.Contains(column))
            {
                throw new DataException(string.Format(
                    "'{0}' is not a valid column name. " +
                    "Valid column names are: '{1}'",
                    column,
                    string.Join("', '", tableColumns.ToArray())));
            }
        }
        return false;
    }

    private IEnumerable<object> GetRowResults(IDataReader data, IEnumerable<string> columns)
    {
        var columnIndexMapping = new Dictionary<string, int>();
        for (var i = 0; i < columns.Count(); i++)
            columnIndexMapping[columns.ElementAt(i)] = i;

        var currentRowNumber = 0;
        while (data.Read())
        {
            if (this._args.SkipEmptyRows && IsLineEmpty(data))
            {
                continue;
            }

            currentRowNumber++;
            IList<Cell> cells = new List<Cell>();
            for (var i = 0; i < columns.Count(); i++)
            {
                try
                {
                    var value = data[i];
                    value = TrimStringValue(value);
                    cells.Add(new Cell(value));
                }
                catch (Exception exception)
                {
                    throw new Exceptions.ExcelException(currentRowNumber, i, columns.ElementAtOrDefault(i), exception);
                }
            }
            yield return new Row(cells, columnIndexMapping);
        }
    }

    private IEnumerable<object> GetRowNoHeaderResults(OleDbDataReader data)
    {
        var currentRowNumber = 0;
        while (data.Read())
        {
            if (this._args.SkipEmptyRows && IsLineEmpty(data))
            {
                continue;
            }

            currentRowNumber++;
            IList<Cell> cells = new List<Cell>();
            for (var i = 0; i < data.FieldCount; i++)
            {
                try
                {
                    var value = data[i];
                    value = TrimStringValue(value);
                    cells.Add(new Cell(value));
                }
                catch (Exception exception)
                {
                    throw new Exceptions.ExcelException(currentRowNumber, i, exception);
                }
            }
            yield return new RowNoHeader(cells);
        }
    }

    private IEnumerable<object> GetTypeResults(IDataReader data, IEnumerable<string> columns, QueryModel queryModel)
    {
        var fromType = queryModel.MainFromClause.ItemType;
        var props = fromType.GetProperties();
        if (_args.StrictMapping.Value != StrictMappingType.None)
            this.ConfirmStrictMapping(columns, props, _args.StrictMapping.Value);
        var gatherUnmappedCells = typeof(IContainsUnmappedCells).IsAssignableFrom(fromType);
        var gatherFieldTypeConversionExceptions = typeof(IAllowFieldTypeConversionExceptions).IsAssignableFrom(fromType);

        var currentRowNumber = 0;
        while (data.Read())
        {
            if (this._args.SkipEmptyRows && IsLineEmpty(data))
            {
                continue;
            }

            currentRowNumber++;
            var result = Activator.CreateInstance(fromType);

            var propMapping = props.Select(prop => new
            {
                prop,
                columnName = (_args.ColumnMappings.ContainsKey(prop.Name)) ?
                    _args.ColumnMappings[prop.Name] :
                    prop.Name
            });

            foreach (var mapping in propMapping)
            {
                if (columns.Contains(mapping.columnName))
                {
                    try
                    {
                        var value = GetColumnValue(data, mapping.columnName, mapping.prop.Name, fromType.Name).Cast(mapping.prop.PropertyType);
                        value = TrimStringValue(value);
                        result.SetProperty(mapping.prop.Name, value);
                    }
                    catch (Exception exception)
                    {
                        var excelException = new Exceptions.ExcelException(currentRowNumber, mapping.columnName, exception);

                        if (gatherFieldTypeConversionExceptions)
                        {
                            var gatherer = (IAllowFieldTypeConversionExceptions)result;
                            if (gatherer.FieldTypeConversionExceptions != null)
                            {
                                gatherer.FieldTypeConversionExceptions.Add(excelException);
                            }
                            else
                            {
                                throw excelException;
                            }
                        }
                        else
                        {
                            throw excelException;
                        }
                    }
                }
            }

            if (gatherUnmappedCells)
            {
                var gatherer = (IContainsUnmappedCells)result;
                if (gatherer.UnmappedCells != null)
                {
                    foreach (var col in columns.Except(propMapping.Select(x => x.columnName)))
                    {
                        var value = data[col];
                        value = TrimStringValue(value);
                        gatherer.UnmappedCells[col] = new Cell(value);
                    }
                }
            }
            yield return result;
        }
    }

    /// <summary>
    /// Trims leading and trailing spaces, based on the value of _args.TrimSpaces
    /// </summary>
    /// <param name="value">Input string value</param>
    /// <returns>Trimmed string value</returns>
    private object TrimStringValue(object value)
    {
        if (value == null || value.GetType() != typeof(string))
            return value;

        switch (_args.TrimSpaces)
        {
            case TrimSpacesType.Start:
                return ((string)value).TrimStart();
            case TrimSpacesType.End:
                return ((string)value).TrimEnd();
            case TrimSpacesType.Both:
                return ((string)value).Trim();
            case TrimSpacesType.None:
            default:
                return value;
        }
    }

    private void ConfirmStrictMapping(IEnumerable<string> columns, PropertyInfo[] properties, StrictMappingType strictMappingType)
    {
        var propertyNames = properties.Select(x => x.Name);
        if (strictMappingType == StrictMappingType.ClassStrict || strictMappingType == StrictMappingType.Both)
        {
            foreach (var propertyName in propertyNames)
            {
                if (!columns.Contains(propertyName) && PropertyIsNotMapped(propertyName))
                    throw new StrictMappingException("'{0}' property is not mapped to a column", propertyName);
            }
        }

        if (strictMappingType == StrictMappingType.WorksheetStrict || strictMappingType == StrictMappingType.Both)
        {
            foreach (var column in columns)
            {
                if (!propertyNames.Contains(column) && ColumnIsNotMapped(column))
                    throw new StrictMappingException("'{0}' column is not mapped to a property", column);
            }
        }
    }

    private bool PropertyIsNotMapped(string propertyName)
    {
        return !_args.ColumnMappings.Keys.Contains(propertyName);
    }

    private bool ColumnIsNotMapped(string columnName)
    {
        return !_args.ColumnMappings.Values.Contains(columnName);
    }

    private object GetColumnValue(IDataRecord data, string columnName, string propertyName, string typeName)
    {
        var transformationKey = string.Format("{0}.{1}", typeName, propertyName);
        //Perform the property transformation if there is one
        return (_args.Transformations.ContainsKey(transformationKey)) ?
            _args.Transformations[transformationKey](data[columnName].ToString()) :
            data[columnName];
    }

    private IEnumerable<object> GetScalarResults(IDataReader data)
    {
        data.Read();
        return new List<object> { data[0] };
    }

    private void LogSqlStatement(SqlParts sqlParts)
    {
        if (_log != null && _log.IsDebugEnabled == true)
        {
            var logMessage = new StringBuilder();
            logMessage.AppendFormat("{0};", sqlParts.ToString());
            for (var i = 0; i < sqlParts.Parameters.Count(); i++)
            {
                var paramValue = sqlParts.Parameters.ElementAt(i).Value.ToString();
                var paramMessage = string.Format(" p{0} = '{1}';",
                    i, sqlParts.Parameters.ElementAt(i).Value.ToString());

                if (paramValue.IsNumber())
                    paramMessage = paramMessage.Replace("'", "");
                logMessage.Append(paramMessage);
            }

            if (_logManagerFactory != null) {
                var sqlLog = _logManagerFactory.GetLogger("LinqToExcelModern.SQL");
                sqlLog.Debug(logMessage.ToString());
            }
        }
    }

    private bool IsLineEmpty(IDataReader data)
    {
        var objects = new object[data.FieldCount];
        data.GetValues(objects);
        return objects.All(obj => obj is DBNull);
    }
}

using LinqToExcelModern.Domain;

namespace LinqToExcelModern.Tests;

class CompanyNameWithUnmappedCells : IContainsUnmappedCells
{
	public CompanyNameWithUnmappedCells()
	{
		UnmappedCells = new Dictionary<string, Cell>();
	}

	public string Name { get; set; }
	public IDictionary<string, Cell> UnmappedCells { get; private set; }
}
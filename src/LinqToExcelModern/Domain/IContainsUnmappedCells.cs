namespace LinqToExcelModern.Domain
{
	/// <summary>
	/// Implement this interface to receive values for cells that
	/// were not mapped.
	/// </summary>
	public interface IContainsUnmappedCells
	{
		IDictionary<string, Cell> UnmappedCells { get; }
	}
}

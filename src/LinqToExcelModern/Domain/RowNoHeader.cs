namespace LinqToExcelModern.Domain
{
    public class RowNoHeader : List<Cell>
    {
        /// <param name="cells">Cells contained within the row</param>
        public RowNoHeader(IEnumerable<Cell> cells)
        {
            base.AddRange(cells);
        }
    }
}

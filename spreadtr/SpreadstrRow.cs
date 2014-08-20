namespace spreadstr
{
    using System.Collections.Generic;

    public class SpreadstrRow
    {
        private readonly List<string> _columns = new List<string>();

        public SpreadstrRow()
        {
        }

        public SpreadstrRow(IEnumerable<string> columns)
        {
            _columns.AddRange(columns);
        }

        public void AddColumn(string column)
        {
            _columns.Add(column);
        }

        public IEnumerable<string> Columns
        {
            get { return _columns; }
        }

        public int ColumnCount
        {
            get { return _columns.Count; }
        }
    }
}
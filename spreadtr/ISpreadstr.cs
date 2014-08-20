namespace spreadstr
{
    using System.Collections.Generic;

    public interface ISpreadstr
    {
        Spreadsheet Generate<T>(IEnumerable<T> items, ISpreadstrConverter<T> converter, string title);
    }
}
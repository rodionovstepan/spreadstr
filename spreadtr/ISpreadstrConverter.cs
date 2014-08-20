namespace spreadstr
{
    public interface ISpreadstrConverter<in T>
    {
        SpreadstrRow Convert(T item);
    }
}
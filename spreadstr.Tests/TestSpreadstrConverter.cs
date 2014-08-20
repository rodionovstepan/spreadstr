namespace spreadstr.Tests
{
    public class TestSpreadstrConverter : ISpreadstrConverter<TestItem>
    {
        public SpreadstrRow Convert(TestItem item)
        {
            return new SpreadstrRow(new[] {item.FirstName, item.LastName});
        }
    }
}
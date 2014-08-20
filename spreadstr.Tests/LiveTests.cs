namespace spreadstr.Tests
{
    using System.IO;
    using Impl;
    using NUnit.Framework;

    public class LiveTests
    {
        [Test]
        public void Test1()
        {
            var spreadstr = new Spreadstr();
            var converter = new TestSpreadstrConverter();
            var items = new[]
                        {
                            new TestItem("Ivan", "Ivanov"),
                            new TestItem("Semen", "Semenov"),
                            new TestItem("Sergey", "Sergeev"),
                            new TestItem("Viktor", "Viktorov"),
                            new TestItem("Vladimir", "Vladimirov")
                        };

            var spreadsheet = spreadstr.Generate(items, converter);

            File.WriteAllBytes("result.xlsx", spreadsheet.Content);
        }
    }
}
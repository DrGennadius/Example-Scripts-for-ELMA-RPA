using ELMA.RPA.Scripts;
using NUnit.Framework;
using System.IO;

namespace Tests
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
            if (File.Exists("test.xlsx"))
            {
                File.Delete("test.xlsx");
            }
        }

        [Test]
        public void ExcelHelperTest()
        {
            ExcelHelper excel = new ExcelHelper("test.xlsx");
            var worksheet = excel.GetOrCreateWorksheet("test");
            var cell = excel.GetOrCreateCell(worksheet, "E", 11);
            excel.SetCellValue(cell, "test E11");
            excel.Close();
            Assert.Pass();
        }
    }
}
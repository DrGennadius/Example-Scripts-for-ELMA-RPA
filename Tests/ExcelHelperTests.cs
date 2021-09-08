using DocumentFormat.OpenXml.Spreadsheet;
using ELMA.RPA.Scripts;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;

namespace Tests
{
    public class ExcelHelperTests
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
        public void GetCreateTest1()
        {
            ExcelHelper excel = new ExcelHelper("test.xlsx");

            var worksheet = excel.GetOrCreateWorksheet("test");

            var cell1 = excel.GetOrCreateCell(worksheet, "E", 11);
            excel.SetCellValue(cell1, "test E11");
            var cell2 = excel.GetOrCreateCell(worksheet, "E11");

            Assert.AreEqual(cell1, cell2);

            excel.Close();

            File.Delete("test.xlsx");

            Assert.Pass();
        }

        [Test]
        public void GetCreateTest2()
        {
            ExcelHelper excel = new ExcelHelper("test.xlsx");

            var worksheet = excel.GetOrCreateWorksheet("test");

            // 1. Text

            var cell1 = excel.GetOrCreateCell(worksheet, "A", 1);
            excel.SetCellValue(cell1, "1");
            cell1 = excel.GetOrCreateCell(worksheet, "A", 1);

            var cell2 = excel.GetOrCreateCell(worksheet, "A", 2);
            excel.SetCellValue(cell2, 1, CellValues.String);
            cell2 = excel.GetOrCreateCell(worksheet, "A", 1);

            Assert.AreEqual(cell1.CellValue.Text, cell2.CellValue.Text);

            // 2. Number

            excel.SetCellValue(cell1, 1);
            cell1 = excel.GetOrCreateCell(worksheet, "A", 1);

            excel.SetCellValue(cell2, "1", CellValues.Number);
            cell2 = excel.GetOrCreateCell(worksheet, "A", 1);

            Assert.AreEqual(cell1.CellValue.Text, cell2.CellValue.Text);

            //

            excel.Close();

            File.Delete("test.xlsx");

            Assert.Pass();
        }

        [Test]
        public void IncrementColumnAddressTest()
        {
            Dictionary<string, string> samples = new()
            {
                { "A", "B" },
                { "b", "C" },
                { "Z", "AA" },
                { "AA", "AB" },
                { "ZZZZZZ", "AAAAAAA" }
            };

            foreach (var item in samples)
            {
                string result = ExcelHelper.IncrementColumnAddress(item.Key);
                Assert.AreEqual(item.Value, result);
            }

            Assert.Pass();
        }

        [Test]
        public void GetOrCreateRightCellTest()
        {
            Dictionary<string, string> samples = new()
            {
                { "A1", "B1" },
                { "b1", "C1" },
                { "Z1", "AA1" },
                { "AA1", "AB1" },
                { "ZZZZZZ1", "AAAAAAA1" }
            };

            ExcelHelper excel = new ExcelHelper("test.xlsx");
            var worksheet = excel.GetOrCreateWorksheet("test");

            foreach (var item in samples)
            {
                var originCell = excel.GetOrCreateCell(worksheet, item.Key);
                var rigthCell = excel.GetOrCreateRightCell(worksheet, originCell);
                Assert.AreEqual(item.Value, (string)rigthCell.CellReference);
            }

            excel.Close();

            File.Delete("test.xlsx");

            Assert.Pass();
        }

        [Test]
        public void FindCellByTextTest()
        {
            string testString = "I'm Gennadius!";

            ExcelHelper excel = new ExcelHelper("test.xlsx");

            var worksheet = excel.GetOrCreateWorksheet("test");

            var originCell = excel.GetOrCreateCell(worksheet, "GENNADY45");
            excel.SetCellValue(originCell, testString);

            var foundCell = excel.FindCellByText(worksheet, testString);

            Assert.IsNotNull(foundCell);

            Assert.AreEqual(originCell, foundCell);

            excel.Close();

            File.Delete("test.xlsx");

            Assert.Pass();
        }

        [Test]
        public void GetWorksheetNameTest()
        {
            string worksheetNameTest = "test";

            ExcelHelper excel = new ExcelHelper("test.xlsx");

            var worksheet = excel.GetOrCreateWorksheet(worksheetNameTest);

            string worksheetName = excel.GetWorksheetName(worksheet);

            Assert.AreEqual(worksheetName, worksheetNameTest);

            excel.Close();

            File.Delete("test.xlsx");

            Assert.Pass();
        }
    }
}
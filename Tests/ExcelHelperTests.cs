using DocumentFormat.OpenXml.Spreadsheet;
using ELMA.RPA.Scripts;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;

namespace Tests
{
    public class ExcelHelperTests
    {
        private const string path = "test.xlsx";

        ExcelHelper excel = null;

        [SetUp]
        public void Setup()
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        [TearDown]
        public void Cleanup()
        {
            if (excel != null)
            {
                excel.Dispose();
                excel = null;
            }
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        [Test]
        public void GetCreateTest1()
        {
            excel = new ExcelHelper(path);

            var worksheet = excel.GetOrCreateWorksheet("test");

            var cell1 = excel.GetOrCreateCell(worksheet, "E", 11);
            excel.SetCellValue(cell1, "test E11");
            var cell2 = excel.GetOrCreateCell(worksheet, "E11");

            Assert.AreEqual(cell1, cell2);

            Assert.Pass();
        }

        [Test]
        public void GetCreateTest2()
        {
            excel = new ExcelHelper(path);

            var worksheet = excel.GetOrCreateWorksheet("test");

            // 1. Text

            var cell1 = excel.GetOrCreateCell(worksheet, "A", 1);
            excel.SetCellValue(cell1, "1");
            cell1 = excel.GetOrCreateCell(worksheet, "A", 1);

            var cell2 = excel.GetOrCreateCell(worksheet, "A", 2);
            excel.SetCellValue(cell2, 1, CellValues.String);
            cell2 = excel.GetOrCreateCell(worksheet, "A", 1);

            Assert.AreEqual(excel.GetCellStringValue(cell1), excel.GetCellStringValue(cell2));

            // 2. Number

            excel.SetCellValue(cell1, 1);
            cell1 = excel.GetOrCreateCell(worksheet, "A", 1);

            excel.SetCellValue(cell2, "1", CellValues.Number);
            cell2 = excel.GetOrCreateCell(worksheet, "A", 1);

            Assert.AreEqual(excel.GetCellStringValue(cell1), excel.GetCellStringValue(cell2));

            //

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

            excel = new ExcelHelper(path);
            var worksheet = excel.GetOrCreateWorksheet("test");

            foreach (var item in samples)
            {
                var originCell = excel.GetOrCreateCell(worksheet, item.Key);
                var rigthCell = excel.GetOrCreateRightCell(worksheet, originCell);
                Assert.AreEqual(item.Value, (string)rigthCell.CellReference);
            }

            Assert.Pass();
        }

        [Test]
        public void FindCellByTextTest()
        {
            string testString = "I'm Gennadius!";

            excel = new ExcelHelper(path);

            var worksheet = excel.GetOrCreateWorksheet("test");

            var originCell = excel.GetOrCreateCell(worksheet, "GENNADY45");
            excel.SetCellValue(originCell, testString);

            var foundCell = excel.FindCellByText(worksheet, testString);

            Assert.IsNotNull(foundCell);

            Assert.AreEqual(originCell, foundCell);

            Assert.Pass();
        }

        [Test]
        public void GetWorksheetNameTest()
        {
            string worksheetNameTest = "test";

            excel = new ExcelHelper(path);

            var worksheet = excel.GetOrCreateWorksheet(worksheetNameTest);

            string worksheetName = excel.GetWorksheetName(worksheet);

            Assert.AreEqual(worksheetName, worksheetNameTest);

            Assert.Pass();
        }
    }
}
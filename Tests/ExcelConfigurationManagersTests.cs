using DocumentFormat.OpenXml.Spreadsheet;
using ELMA.RPA.Scripts;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Tests
{
    public class ExcelConfigurationManagerTests
    {
        ExcelConfigurationManager configurationManager = null;
        const int testCount = 200;
        const int multTestCount = 100;

        [SetUp]
        public void Setup()
        {
            if (File.Exists("test.xlsx"))
            {
                File.Delete("test.xlsx");
            }
            configurationManager = new ExcelConfigurationManager("test.xlsx");
            configurationManager.Read();
        }

        [TearDown]
        public void Cleanup()
        {
            if (configurationManager != null)
            {
                configurationManager.Dispose();
                configurationManager = null;
            }
            if (File.Exists("test.xlsx"))
            {
                File.Delete("test.xlsx");
            }
        }

        [Test]
        public void Test1()
        {
            Assert.IsTrue(configurationManager.SingleParams.Count == 0, "При инициализации листов не должно быть, а следовательно тут кол-во 0 должно быть.");

            configurationManager.SingleParams.Add("sheet1", new Dictionary<string, string>()
            {
                { "param1", "value1" },
                { "param2", "value2" },
                { "param3", "value3" },
                { "param4", "value4" },
                { "param5", "value5" },
                { "param6", "value6" }
            });

            configurationManager.Save();
            configurationManager.Dispose();
            configurationManager = new ExcelConfigurationManager("test.xlsx");
            configurationManager.Read();

            Assert.IsTrue(configurationManager.SingleParams.Count == 1, "Ожидалось один лист.");
            Assert.IsTrue(configurationManager.SingleParams["sheet1"].Count == 6, "Ожидалось 6 параметров.");

            Assert.Pass();
        }

        [Test]
        public void Test2()
        {
            Assert.IsTrue(configurationManager.SingleParams.Count == 0, "При инициализации листов не должно быть, а следовательно тут кол-во 0 должно быть.");

            for (int i = 1; i <= testCount; i++)
            {
                Dictionary<string, string> keyValues = new();
                for (int j = 1; j <= testCount; j++)
                {
                    keyValues.Add("param" + j, "value" + j);
                }
                configurationManager.SingleParams.Add("sheet" + i, keyValues);
            }

            // Множественый параметр.
            for (int i = 1; i <= multTestCount; i++)
            {
                Dictionary<string, IEnumerable<string>> keyValues = new();
                for (int j = 1; j <= multTestCount; j++)
                {
                    List<string> values = new();
                    for (int k = 1; k <= multTestCount; k++)
                    {
                        values.Add($"value({j},{k})");
                    }
                    keyValues.Add("multParam" + j, values);
                }
                configurationManager.MultipleParams.Add("multSheet" + i, keyValues);
            }

            configurationManager.Save();
            configurationManager.Dispose();
            configurationManager = new ExcelConfigurationManager("test.xlsx");
            configurationManager.Read();

            Assert.IsTrue(configurationManager.SingleParams.Count == testCount, "Ожидалось" + testCount + "листов.");
            Assert.IsTrue(configurationManager.SingleParams["sheet1"].Count == testCount, "Ожидалось" + testCount + "параметров.");
            Assert.IsTrue(configurationManager.SingleParams["sheet" + testCount].Count == testCount, "Ожидалось" + testCount + "параметров.");

            Assert.AreEqual(configurationManager.SingleParams["sheet" + testCount]["param" + testCount], "value" + testCount);

            // Множественый параметр.
            Assert.IsTrue(configurationManager.MultipleParams.Count == multTestCount, "Ожидалось" + multTestCount + "листов.");
            Assert.IsTrue(configurationManager.MultipleParams["multSheet1"].Count == multTestCount, "Ожидалось" + multTestCount + "параметров.");
            Assert.IsTrue(configurationManager.MultipleParams["multSheet" + multTestCount].Count == multTestCount, "Ожидалось" + multTestCount + "параметров.");

            Assert.AreEqual(
                configurationManager.MultipleParams["multSheet" + multTestCount]["multParam" + multTestCount].ToArray()[multTestCount - 1], 
                $"value({multTestCount},{multTestCount})"
            );

            Assert.Pass();
        }
    }
}
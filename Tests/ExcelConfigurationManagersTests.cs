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
        readonly int testCount = 200;

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
            Assert.IsTrue(configurationManager.Params.Count == 0, "При инициализации листов не должно быть, а следовательно тут кол-во 0 должно быть.");

            configurationManager.Params.Add("sheet1", new Dictionary<string, string>()
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

            Assert.IsTrue(configurationManager.Params.Count == 1, "Ожидалось один лист.");
            Assert.IsTrue(configurationManager.Params["sheet1"].Count == 6, "Ожидалось 6 параметров.");

            Assert.Pass();
        }

        [Test]
        public void Test2()
        {
            Assert.IsTrue(configurationManager.Params.Count == 0, "При инициализации листов не должно быть, а следовательно тут кол-во 0 должно быть.");

            for (int i = 1; i <= testCount; i++)
            {
                Dictionary<string, string> keyValues = new Dictionary<string, string>();
                for (int j = 1; j <= testCount; j++)
                {
                    keyValues.Add("param" + j, "value" + j);
                }
                configurationManager.Params.Add("sheet" + i, keyValues);
            }

            configurationManager.Save();
            configurationManager.Dispose();
            configurationManager = new ExcelConfigurationManager("test.xlsx");
            configurationManager.Read();

            Assert.IsTrue(configurationManager.Params.Count == testCount, "Ожидалось" + testCount + "листов.");
            Assert.IsTrue(configurationManager.Params["sheet1"].Count == testCount, "Ожидалось" + testCount + "параметров.");
            Assert.IsTrue(configurationManager.Params["sheet" + testCount].Count == testCount, "Ожидалось" + testCount + "параметров.");

            Assert.AreEqual(configurationManager.Params["sheet" + testCount]["param" + testCount], "value" + testCount);

            Assert.Pass();
        }
    }
}
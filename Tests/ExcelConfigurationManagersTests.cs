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

        [SetUp]
        public void Setup()
        {
            if (File.Exists("test.xlsx"))
            {
                File.Delete("test.xlsx");
            }
        }

        [TearDown]
        public void Cleanup()
        {
            if (configurationManager != null)
            {
                configurationManager.Dispose();
            }
            if (File.Exists("test.xlsx"))
            {
                File.Delete("test.xlsx");
            }
        }

        [Test]
        public void Test1()
        {
            configurationManager = new ExcelConfigurationManager("test.xlsx");
            configurationManager.Read();

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
    }
}
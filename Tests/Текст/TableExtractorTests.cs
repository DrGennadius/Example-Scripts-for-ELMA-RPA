using ELMA.RPA.Scripts;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Tests
{
    public class TableExtractorTests
    {
        string sampleText1;

        [SetUp]
        public void Setup()
        {
            sampleText1 = File.ReadAllText("Пример текста.txt");
        }

        [Test]
        public void SimpleExtractTest()
        {
            TableDetectFeatures tableDetectFeatures = new()
            {
                FirstCellText = "№",
                HasStartSequentialNumberingCells = true,
                SplitPattern = @"\s{2,}"
            };
            TableExtractor tableExtractor = new(tableDetectFeatures);
            tableExtractor.Extract(sampleText1);

            Assert.Pass();
        }
    }
}

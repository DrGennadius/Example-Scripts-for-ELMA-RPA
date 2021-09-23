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
            // TODO: Еще можно сделать преобразование переходов строк на Linux (\n),
            // но в этом нет необходимости пока.
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

            Assert.AreEqual(tableExtractor.Data.GetLength(0), 5);
            Assert.AreEqual(tableExtractor.Data.GetLength(1), 4);

            Assert.AreEqual(tableExtractor.Data[0, 0], "№");
            Assert.AreEqual(tableExtractor.Data[1, 0], "1");
            Assert.AreEqual(tableExtractor.Data[3, 2], "Я какой-то документ 3.docx");
            Assert.AreEqual(
                tableExtractor.Data[2, 1],
                "Признак предоставления" 
                + Environment.NewLine 
                + "гарантии на качество"
                + Environment.NewLine
                + "предлагаемых работ и услуг"
            );

            Assert.Pass();
        }
    }
}

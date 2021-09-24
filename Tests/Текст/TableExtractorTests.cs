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
        string sampleText2;

        [SetUp]
        public void Setup()
        {
            sampleText1 = File.ReadAllText("Пример текста.txt");
            sampleText2 = File.ReadAllText("Пример текста 2.txt");
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

        [Test]
        public void MultiplyExtractTest()
        {
            TableDetectFeatures tableDetectFeatures = new()
            {
                FirstCellText = "№",
                HasStartSequentialNumberingCells = false,
                SplitPattern = @"\s{2,}",
                LineSkipPattern = "(Это пример текста колонтитула, бла-бла-бла ООО.+$)"
                                + "|(Документ сформирован порталом портала АО.+$)"
            };

            TableExtractor tableExtractor = new(tableDetectFeatures);
            tableExtractor.ExtractNext(sampleText2);
            var data1 = tableExtractor.Data;

            tableExtractor.ExtractNext(sampleText2);
            var data2 = tableExtractor.Data;

            Assert.IsFalse(tableExtractor.ExtractNext(sampleText2));

            Assert.AreEqual(data1.GetLength(0), data2.GetLength(0));
            Assert.AreEqual(data1.GetLength(1), data2.GetLength(1));

            Assert.AreEqual(data1[0, 0], data2[0, 0]);
            Assert.AreEqual(data1[1, 0], data2[1, 0]);
            Assert.AreEqual(data1[3, 2], data2[3, 2]);
            Assert.AreEqual(data1[2, 1], data2[2, 1]);

            Assert.Pass();
        }
    }
}

using ELMA.RPA.Scripts;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tests
{
    public class TableDetectTests
    {
        string sampleText1;

        [SetUp]
        public void Setup()
        {
            sampleText1 = File.ReadAllText("Пример текста.txt");
        }

        [Test]
        public void Test1()
        {
            TableDetectFeatures tableDetectFeatures1 = new()
            {
                FirstCellText = "№",
                HasStartSequentialNumberingCells = false,
                SplitPattern = @"\s{2,}"
            };
            TableDetectFeatures tableDetectFeatures2 = new()
            {
                FirstCellText = "№",
                HasStartSequentialNumberingCells = true,
                SplitPattern = @"\s{2,}"
            };
            var tableParameters1 = CommonTestBody(tableDetectFeatures1);
            var tableParameters2 = CommonTestBody(tableDetectFeatures2);

            Assert.AreEqual(tableParameters1.FirstCharIndex, tableParameters2.FirstCharIndex);
            Assert.AreEqual(tableParameters1.LastCharIndex, tableParameters2.LastCharIndex);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes.Length, tableParameters2.BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[0], tableParameters2.BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[1], tableParameters2.BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[2], tableParameters2.BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[3], tableParameters2.BeginColumnIndexes[3]);

            Assert.Pass();
        }

        private TableParameters CommonTestBody(TableDetectFeatures tableDetectFeatures)
        {
            TableDetector tableDetector = new(tableDetectFeatures);
            var tableParameters = tableDetector.Detect(sampleText1);

            Assert.IsTrue(tableParameters.HasValue);
            int firstCharIndex = tableParameters.Value.FirstCharIndex;
            int lastCharIndex = tableParameters.Value.LastCharIndex;
            Assert.AreEqual(sampleText1[firstCharIndex], '№');

            int[] beginColumnIndexes = tableParameters.Value.BeginColumnIndexes;
            Assert.AreEqual(beginColumnIndexes.Length, 4);

            int[] beginColumnIndexesOffset = beginColumnIndexes.Select(x => x + tableParameters.Value.FirstCharIndex).ToArray();
            Assert.AreEqual(sampleText1[beginColumnIndexesOffset[0]], '№');
            Assert.AreEqual(sampleText1[beginColumnIndexesOffset[1]], 'Н');
            Assert.AreEqual(sampleText1[beginColumnIndexesOffset[2]], 'Н');
            Assert.AreEqual(sampleText1[beginColumnIndexesOffset[3]], 'К');

            int lastRowIndex = tableDetector.LastTableRowIndex(ref sampleText1, lastCharIndex);
            Assert.AreEqual(lastRowIndex, 25);

            return tableParameters.Value;
        }
    }
}

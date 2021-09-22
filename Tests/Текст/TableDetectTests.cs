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
            TableDetector tableDetector1 = new(tableDetectFeatures1);
            var tableParameters = tableDetector1.Detect(sampleText1);

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

            int lastRowIndex = tableDetector1.LastTableRowIndex(ref sampleText1, lastCharIndex);
            Assert.AreEqual(lastRowIndex, 25);

            Assert.Pass();
        }
    }
}

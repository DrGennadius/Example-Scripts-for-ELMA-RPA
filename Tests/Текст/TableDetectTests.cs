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
    public class TableDetectTests
    {
        string sampleText1;
        string sampleText2;

        [SetUp]
        public void Setup()
        {
            sampleText1 = File.ReadAllText("Пример текста.txt");
            sampleText2 = File.ReadAllText("Пример текста 2.txt");
        }

        [Test]
        public void SimpleDetectTest()
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
            var tableParameters1 = SimpleDetectTestBody(tableDetectFeatures1);
            var tableParameters2 = SimpleDetectTestBody(tableDetectFeatures2);

            Assert.AreEqual(tableParameters1.FirstCharIndex, tableParameters2.FirstCharIndex);
            Assert.AreEqual(tableParameters1.LastCharIndex, tableParameters2.LastCharIndex);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes.Length, tableParameters2.BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[0], tableParameters2.BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[1], tableParameters2.BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[2], tableParameters2.BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[3], tableParameters2.BeginColumnIndexes[3]);

            Assert.Pass();
        }

        [Test]
        public void GenerateAndDetectTest()
        {
            Console.WriteLine("GenerateAndDetectTest");
            CancellationTokenSource tokenSource = new();
            var cancellationToken = tokenSource.Token;
            Task task = new(() => PrintPrivateMemory(), cancellationToken);
            task.Start();

            TableDetectFeatures tableDetectFeatures = new()
            {
                FirstCellText = "№",
                HasStartSequentialNumberingCells = false,
                SplitPattern = @"\s{2,}"
            };

            string text = GenerateTextWithTable();
            TableDetector tableDetector = new(tableDetectFeatures);

            long start = Stopwatch.GetTimestamp();
            var tableParameters = tableDetector.Detect(text).Value;
            long end = Stopwatch.GetTimestamp();
            Console.WriteLine($"На обнаруживание таблицы 1 потрачено {end - start} тиков");

            int firstCharIndex = tableParameters.FirstCharIndex;
            int lastCharIndex = tableParameters.LastCharIndex;

            Assert.AreEqual(firstCharIndex, 0);
            Assert.AreEqual(lastCharIndex, text.Length);

            string preText = "Bla-bla-bla!" + Environment.NewLine;
            text = preText + text + Environment.NewLine + "Bla-bla-bla!";
            tableDetector = new(tableDetectFeatures);

            start = Stopwatch.GetTimestamp();
            tableParameters = tableDetector.Detect(text).Value;
            end = Stopwatch.GetTimestamp();
            Console.WriteLine($"На обнаруживание таблицы 2 потрачено {end - start} тиков");

            int firstCharIndex2 = tableParameters.FirstCharIndex;
            int lastCharIndex2 = tableParameters.LastCharIndex;

            Assert.AreEqual(firstCharIndex2 - preText.Length, firstCharIndex);
            Assert.AreEqual(lastCharIndex2 - preText.Length - Environment.NewLine.Length, lastCharIndex);

            tokenSource.Cancel();

            Assert.Pass();
        }

        private TableParameters SimpleDetectTestBody(TableDetectFeatures tableDetectFeatures)
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

            int lastRowIndex = tableDetector.LastTableRowIndex(sampleText1, lastCharIndex);
            Assert.AreEqual(lastRowIndex, 25);

            return tableParameters.Value;
        }

        [Test]
        public void MultiplyDetectTest()
        {
            TableDetectFeatures tableDetectFeatures = new()
            {
                FirstCellText = "№",
                HasStartSequentialNumberingCells = false,
                SplitPattern = @"\s{2,}",
                LineSkipPattern = "(Это пример текста колонтитула, бла-бла-бла ООО.+$)"
                                + "|(Документ сформирован порталом портала АО.+$)"
            };

            TableDetector tableDetector = new(tableDetectFeatures);

            var tableParameters1 = tableDetector.Detect(sampleText2).Value;
            var tableParameters2 = tableDetector.Detect(sampleText2, tableParameters1).Value;
            var tableParameters3 = tableDetector.Detect(sampleText2, tableParameters2);

            Assert.IsFalse(tableParameters3.HasValue);

            Assert.AreEqual(tableParameters1.BeginColumnIndexes.Length, tableParameters2.BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[0], tableParameters2.BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[1], tableParameters2.BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[2], tableParameters2.BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[3], tableParameters2.BeginColumnIndexes[3]);

            // Сразу все.
            var tablesParameters = tableDetector.DetectAll(sampleText2);

            Assert.AreEqual(tablesParameters.Length, 2);

            Assert.AreEqual(tableParameters1.BeginColumnIndexes.Length, tablesParameters[0].BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[0], tablesParameters[0].BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[1], tablesParameters[0].BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[2], tablesParameters[0].BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[3], tablesParameters[0].BeginColumnIndexes[3]);

            Assert.AreEqual(tableParameters1.BeginColumnIndexes.Length, tablesParameters[0].BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[0], tablesParameters[0].BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[1], tablesParameters[0].BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[2], tablesParameters[0].BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.BeginColumnIndexes[3], tablesParameters[0].BeginColumnIndexes[3]);

            Assert.Pass();
        }

        private string GenerateTextWithTable()
        {
            StringBuilder builder = new();
            const int rowsCount = 10000;
            const int columnsCount = 1000;
            int cellLength = (int)Math.Log10(columnsCount) + 1;
            GenerateHeader(builder, rowsCount, columnsCount, cellLength);
            GenerateTableBody(builder, rowsCount, columnsCount, cellLength);
            return builder.ToString();
        }

        private void GenerateHeader(StringBuilder builder, int rowsCount, int columnsCount, int cellLength)
        {
            builder.Append($"{GenerateCellText($"№", cellLength)}  ");
            for (int i = 1; i < columnsCount; i++)
            {
                builder.Append($"{GenerateCellText($"{i}", cellLength)}  ");
            }
            builder.Append(Environment.NewLine);
        }

        private void GenerateTableBody(StringBuilder builder, int rowsCount, int columnsCount, int cellLength)
        {
            for (int i = 1; i < rowsCount; i++)
            {
                for (int c = 0; c < columnsCount; c++)
                {
                    builder.Append($"{GenerateCellText($"{c}", cellLength)}  ");
                }
                builder.Append(Environment.NewLine);
            }
        }

        private string GenerateCellText(string value, int cellLength)
        {
            int valueLength = value.Length;
            string appendString = "";
            for (int i = valueLength; i < cellLength; i++)
            {
                appendString += ' ';
            }
            return value + appendString;
        }

        private void PrintPrivateMemory()
        {
            while(true)
            {
                Process proc = Process.GetCurrentProcess();
                Console.WriteLine("Private Memory: {0:0.##} MB", (double)proc.PrivateMemorySize64 / (1024 * 1024));
                Thread.Sleep(10000);
            }
        }
    }
}

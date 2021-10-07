using ELMA.RPA.Scripts;
using NUnit.Framework;
using System;
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
            TableFeatures tableDetectFeatures1 = new()
            {
                FirstTableCellWordPattern = "№",
                HasStartSequentialNumberingCells = false,
                SplitPattern = @"\s{2,}"
            };
            TableFeatures tableDetectFeatures2 = new()
            {
                FirstTableCellWordPattern = "№",
                HasStartSequentialNumberingCells = true,
                SplitPattern = @"\s{2,}"
            };
            var tableParameters1 = SimpleDetectTestBody(tableDetectFeatures1);
            var tableParameters2 = SimpleDetectTestBody(tableDetectFeatures2);

            Assert.AreEqual(tableParameters1.FirstCharIndex, tableParameters2.FirstCharIndex);
            Assert.AreEqual(tableParameters1.LastCharIndex, tableParameters2.LastCharIndex);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes.Length, tableParameters2.RowInfoItems[0].BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[0], tableParameters2.RowInfoItems[0].BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[1], tableParameters2.RowInfoItems[0].BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[2], tableParameters2.RowInfoItems[0].BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[3], tableParameters2.RowInfoItems[0].BeginColumnIndexes[3]);

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

            TableFeatures tableDetectFeatures = new()
            {
                FirstTableCellWordPattern = "№",
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
            // Вычитаем длину Environment.NewLine, т.к. теперь не считается пустая строка вместе с переносом.
            Assert.AreEqual(lastCharIndex, text.Length - Environment.NewLine.Length);

            text += Environment.NewLine;

            tableDetector = new(tableDetectFeatures);

            start = Stopwatch.GetTimestamp();
            tableParameters = tableDetector.Detect(text).Value;
            end = Stopwatch.GetTimestamp();
            Console.WriteLine($"На повторное обнаруживание таблицы 1 потрачено {end - start} тиков");

            firstCharIndex = tableParameters.FirstCharIndex;
            int lastCharIndexX = tableParameters.LastCharIndex;

            Assert.AreEqual(firstCharIndex, 0);
            Assert.AreEqual(lastCharIndex, lastCharIndexX);

            text += Environment.NewLine;

            string preText = "Bla-bla-bla!" + Environment.NewLine;
            text = preText + text + "Bla-bla-bla!";
            tableDetector = new(tableDetectFeatures);

            start = Stopwatch.GetTimestamp();
            tableParameters = tableDetector.Detect(text).Value;
            end = Stopwatch.GetTimestamp();
            Console.WriteLine($"На обнаруживание таблицы 2 потрачено {end - start} тиков");

            int firstCharIndex2 = tableParameters.FirstCharIndex;
            int lastCharIndex2 = tableParameters.LastCharIndex;

            // Тут считаем отклонение в размере префиксного текста.
            int char1 = firstCharIndex2 - preText.Length;
            // Тут берем снова просто последний индекс,
            // т.к. теперь переделано определение таблиц
            // и добавлена компенсация на переходы на пустые строки в конце.
            int char2 = lastCharIndex2;

            // Вычитаем кол-во символов перехода на новую строку, т.к. мы переносили перед таблицей.
            string testSubtext1 = text[char1..(firstCharIndex2 - Environment.NewLine.Length)];
            string testSubtext2 = text[char2..];

            Assert.AreEqual(testSubtext1, testSubtext2);

            tokenSource.Cancel();

            Assert.Pass();
        }

        private TableParameters SimpleDetectTestBody(TableFeatures tableDetectFeatures)
        {
            TableDetector tableDetector = new(tableDetectFeatures);
            var tableParameters = tableDetector.Detect(sampleText1);

            Assert.IsTrue(tableParameters.HasValue);
            int firstCharIndex = tableParameters.Value.FirstCharIndex;
            int lastCharIndex = tableParameters.Value.LastCharIndex;
            Assert.AreEqual(sampleText1[firstCharIndex], '№');

            int[] beginColumnIndexes = tableParameters.Value.RowInfoItems[0].BeginColumnIndexes;
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
            TableFeatures tableDetectFeatures = new()
            {
                FirstTableCellWordPattern = "№",
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

            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes.Length, tableParameters2.RowInfoItems[0].BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[0], tableParameters2.RowInfoItems[0].BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[1], tableParameters2.RowInfoItems[0].BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[2], tableParameters2.RowInfoItems[0].BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[3], tableParameters2.RowInfoItems[0].BeginColumnIndexes[3]);

            // Сразу все.
            var tablesParameters = tableDetector.DetectAll(sampleText2);

            Assert.AreEqual(tablesParameters.Length, 2);

            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes.Length, tablesParameters[0].RowInfoItems[0].BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[0], tablesParameters[0].RowInfoItems[0].BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[1], tablesParameters[0].RowInfoItems[0].BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[2], tablesParameters[0].RowInfoItems[0].BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[3], tablesParameters[0].RowInfoItems[0].BeginColumnIndexes[3]);

            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes.Length, tablesParameters[0].RowInfoItems[0].BeginColumnIndexes.Length);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[0], tablesParameters[0].RowInfoItems[0].BeginColumnIndexes[0]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[1], tablesParameters[0].RowInfoItems[0].BeginColumnIndexes[1]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[2], tablesParameters[0].RowInfoItems[0].BeginColumnIndexes[2]);
            Assert.AreEqual(tableParameters1.RowInfoItems[0].BeginColumnIndexes[3], tablesParameters[0].RowInfoItems[0].BeginColumnIndexes[3]);

            Assert.Pass();
        }

        private string GenerateTextWithTable()
        {
            StringBuilder builder = new();
            const int rowsCount = 1000;
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

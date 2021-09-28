using ELMA.RPA.Scripts;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace Tests
{
    public class TableExtractorTests
    {
        string sampleText1;
        string sampleText2;
        string sampleText3;

        [SetUp]
        public void Setup()
        {
            sampleText1 = File.ReadAllText("Пример текста.txt");
            sampleText2 = File.ReadAllText("Пример текста 2.txt");
            sampleText3 = File.ReadAllText("Пример текста 3.txt");
            // TODO: Еще можно сделать преобразование переходов строк на Linux (\n),
            // но в этом нет необходимости пока.
        }

        [Test]
        public void SimpleExtractTest()
        {
            TableFeatures tableDetectFeatures = new()
            {
                FirstTableCellWordPattern = "№",
                HasStartSequentialNumberingCells = true,
                SplitPattern = @"\s{2,}"
            };
            TableExtractor tableExtractor = new(tableDetectFeatures);
            tableExtractor.Extract(sampleText1);

            Assert.AreEqual(tableExtractor.Data.Length, 5);
            Assert.AreEqual(tableExtractor.Data[0].Length, 4);

            Assert.AreEqual(tableExtractor.Data[0][0], "№");
            Assert.AreEqual(tableExtractor.Data[1][0], "1");
            Assert.AreEqual(tableExtractor.Data[3][2], "Я какой-то документ 3.docx");
            Assert.AreEqual(
                tableExtractor.Data[2][1],
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
            TableFeatures tableDetectFeatures = new()
            {
                FirstTableCellWordPattern = "№",
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

            Assert.AreEqual(data1.Length, data2.Length);
            Assert.AreEqual(data1[0].Length, data2[0].Length);

            Assert.AreEqual(data1[0][0], data2[0][0]);
            Assert.AreEqual(data1[1][0], data2[1][0]);
            Assert.AreEqual(data1[3][2], data2[3][2]);
            Assert.AreEqual(data1[2][1], data2[2][1]);

            // Сразу все.
            tableExtractor.Clear();
            var data = tableExtractor.GetAll(sampleText2).ToArray();
            Assert.AreEqual(data.Length, 2);

            Assert.AreEqual(data1.Length, data[0].Length);
            Assert.AreEqual(data1[0].Length, data[0][0].Length);

            Assert.AreEqual(data1[0][0], data[0][0][0]);
            Assert.AreEqual(data1[1][0], data[0][1][0]);
            Assert.AreEqual(data1[3][2], data[0][3][2]);
            Assert.AreEqual(data1[2][1], data[0][2][1]);

            // На всякий случай второй еще проверим.
            Assert.AreEqual(data1.Length, data[1].Length);
            Assert.AreEqual(data1[0].Length, data[1][0].Length);

            Assert.AreEqual(data1[0][0], data[1][0][0]);
            Assert.AreEqual(data1[1][0], data[1][1][0]);
            Assert.AreEqual(data1[3][2], data[1][3][2]);
            Assert.AreEqual(data1[2][1], data[1][2][1]);

            Assert.Pass();
        }

        [Test]
        public void ExtractJsonTest()
        {
            TableFeatures tableDetectFeatures = new()
            {
                FirstTableCellWordPattern = "№",
                HasStartSequentialNumberingCells = false,
                SplitPattern = @"\s{2,}",
                LineSkipPattern = "(Это пример текста колонтитула, бла-бла-бла ООО.+$)"
                                + "|(Документ сформирован порталом портала АО.+$)"
            };

            TableExtractor tableExtractor = new(tableDetectFeatures);
            tableExtractor.ExtractNext(sampleText2);
            var masterData = tableExtractor.Data;
            string jsonData1 = tableExtractor.JsonData;

            tableExtractor.ExtractNext(sampleText2);
            string jsonData2 = tableExtractor.JsonData;

            jsonData2 = jsonData2.Replace("333339", "333333");

            Assert.AreEqual(jsonData1, jsonData2);

            var options = new JsonSerializerOptions
            {
                // Кодировка для Unicode: Basic Latin и Cyrillic
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
            };
            var data1 = JsonSerializer.Deserialize<string[][]>(jsonData1, options);
            var data2 = JsonSerializer.Deserialize<string[][]>(jsonData2, options);

            Assert.IsFalse(tableExtractor.ExtractNext(sampleText2));

            Assert.AreEqual(masterData.Length, data1.Length);
            Assert.AreEqual(masterData[0].Length, data1[0].Length);

            Assert.AreEqual(data1.Length, data2.Length);
            Assert.AreEqual(data1[0].Length, data2[0].Length);

            Assert.AreEqual(masterData[0][0], data1[0][0]);
            Assert.AreEqual(masterData[1][0], data1[1][0]);
            Assert.AreEqual(masterData[3][2], data1[3][2]);
            Assert.AreEqual(masterData[2][1], data1[2][1]);

            Assert.AreEqual(data1[0][0], data2[0][0]);
            Assert.AreEqual(data1[1][0], data2[1][0]);
            Assert.AreEqual(data1[3][2], data2[3][2]);
            Assert.AreEqual(data1[2][1], data2[2][1]);

            Assert.Pass();
        }

        [Test]
        public void SimpleExtractTest2()
        {
            TableFeatures tableDetectFeatures = new()
            {
                HeaderCellPatterns = new()
                {
                    @"Номер\s+лота",
                    @"Наименование\s+закупаемых\s+товаров,\s+работ\s+и\s+услуг",
                    @"Краткая\s+характеристика\s+\(описание\)\s+товаров,\s+работ\s+и\s+услуг\s+с\s+указанием\s+СТ\s+РК,\s+ГОСТ,\s+ТУ\s+и\s+т.д.",
                    @"Количество",
                    @"Сумма,\s+выделенная\s+для\s+закупки\s+без\s+учета\s+НДС",
                    @"Приоритет\s+закупки"
                },
                FirstTableCellWordPattern = "Номер",
                FirstBodyRowCellWordPattern = "№",
                HasStartSequentialNumberingCells = false,
                SplitPattern = @"\s{2,}"
            };
            TableExtractor tableExtractor = new(tableDetectFeatures);
            tableExtractor.Extract(sampleText3);

            Assert.AreEqual(tableExtractor.Data.Length, 5);
            Assert.AreEqual(tableExtractor.Data[0].Length, 6);

            Assert.Pass();
        }
    }
}

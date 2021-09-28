using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Text.Unicode;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Извлекатель таблицы из текста.
    /// </summary>
    public class TableExtractor
    {
        readonly TableDetector _tableDetector = new();

        private string[][] _data = Array.Empty<string[]>();

        private int _lastIndex = -1;

        public TableExtractor() { }

        public TableExtractor(TableFeatures detectFeatures)
        {
            _tableDetector = new(detectFeatures);
        }

        public TableExtractor(TableDetector detector)
        {
            _tableDetector = detector;
        }

        /// <summary>
        /// Данные.
        /// </summary>
        public string[][] Data => _data;

        /// <summary>
        /// Данные как Json строка.
        /// </summary>
        public string JsonData
        {
            get
            {
                var options = new JsonSerializerOptions
                {
                    // Кодировка для Unicode: Basic Latin и Cyrillic
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
                };
                return JsonSerializer.Serialize(_data, options);
            }
        }

        /// <summary>
        /// Данные как матрица.
        /// </summary>
        public string[,] MatrixData
        {
            get
            {
                if (_data.Length == 0)
                {
                    return new string[0, 0];
                }
                int len1 = _data.Length;
                int len2 = _data[0].Length;
                string[,] matrix = new string[len1, len2];
                int r = 0;
                for (int i = 0; i < len1; i++)
                {
                    for (int k = 0; k < len2; k++)
                    {
                        matrix[r, k] = _data[i][k];
                    }
                }
                return matrix;
            }
        }

        /// <summary>
        /// Извлечение данных таблицы.
        /// </summary>
        /// <param name="text">Текст.</param>
        /// <param name="startIndex">Начальный индекс.</param>
        /// <returns></returns>
        public bool Extract(string text, int startIndex = 0)
        {
            var tableParameters = _tableDetector.Detect(text, startIndex);
            return tableParameters.HasValue
                && Extract(text, tableParameters.Value);
        }

        /// <summary>
        /// Извлечение данных таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="tableParameters"></param>
        /// <returns></returns>
        public bool Extract(string text, TableParameters tableParameters)
        {
            bool isSucces = true;
            int columnsLength = tableParameters.BeginColumnIndexes.Length;

            string[][] rawData = GetRawTableData(text, tableParameters);
            if (rawData.Length == 0 || rawData[0].Length == 0)
            {
                return false;
            }

            List<string[]> dataList = new();

            string[] row = Array.Empty<string>();
            if (_tableDetector.DetectFeatures.HeaderCellPatterns.Any())
            {
                // Немного улучшенный вариант. Использывание паттернов ячеек заголовка.
                var result = CalcHeaderCellsByPatterns(rawData, _tableDetector.DetectFeatures.HeaderCellPatterns);
                int actualColumnsLength = result.CheckHeaderCellByPatternResults.Length;
                row = Enumerable.Repeat("", actualColumnsLength).ToArray();
                if (result.IsSucces)
                {
                    string[] tempRow = result.CheckHeaderCellByPatternResults.Select(x => x.CellText).ToArray();
                    dataList.Add(tempRow);
                    int currentRowIndex = result.CheckHeaderCellByPatternResults.Max(x => x.EndRowIndex) + 1;
                    for (int i = currentRowIndex; i < rawData.Length; i++)
                    {
                        int rawColumnIndex = -1;
                        tempRow = Enumerable.Repeat("", actualColumnsLength).ToArray();
                        for (int k = 0; k < actualColumnsLength; k++)
                        {
                            var cellParams = result.CheckHeaderCellByPatternResults[k];
                            int offset = cellParams.EndColumnIndex - cellParams.BeginColumnIndex;
                            for (int c = 0; c < offset + 1; c++)
                            {
                                rawColumnIndex++;
                                tempRow[k] += rawData[i][rawColumnIndex].Trim() + ' ';
                            }
                        }
                        if (IsBeginNewTableRow(tempRow) && !IsEmptyRow(row))
                        {
                            dataList.Add(row);
                            row = tempRow;
                        }
                        else
                        {
                            row = ConcatRows(new string[][] { row, tempRow })
                                .Select(x => x.Trim())
                                .ToArray();
                        }
                    }
                }
            }
            else
            {
                // Обычный вариант.
                row = Enumerable.Repeat("", rawData[0].Length).ToArray();
                foreach (var rawRow in rawData)
                {
                    if (IsBeginNewTableRow(rawRow) && !IsEmptyRow(row))
                    {
                        dataList.Add(row);
                        row = rawRow;
                    }
                    else
                    {
                        row = ConcatRows(new string[][] { row, rawRow })
                            .Select(x => x.Trim())
                            .ToArray();
                    }
                }
            }
            if (!IsEmptyRow(row))
            {
                dataList.Add(row);
            }

            _data = dataList.ToArray();

            // TODO: Можно еще какие-нибудь проверочки понадобавлять для определения успешности
            return isSucces;
        }

        /// <summary>
        /// Извлечение данных следующей таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public bool ExtractNext(string text)
        {
            int index = _lastIndex + 1;
            return index < text.Length
                && Extract(text, index);
        }

        /// <summary>
        /// Извлечение данных и получение всех таблиц, которые будут обнаружены.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public IEnumerable<string[][]> GetAll(string text)
        {
            Stack<string[][]> data = new();

            while (ExtractNext(text))
            {
                data.Push(Data);
            }

            return data.AsEnumerable();
        }

        /// <summary>
        /// Очистить.
        /// </summary>
        public void Clear()
        {
            _data = Array.Empty<string[]>();
            _lastIndex = -1;
        }

        /// <summary>
        /// Получить "сырые" данные таблицы без каких либор дополнительных обработок.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="tableParameters"></param>
        /// <returns></returns>
        private string[][] GetRawTableData(string text, TableParameters tableParameters)
        {
            bool isCheckSkipOn = !string.IsNullOrWhiteSpace(_tableDetector.DetectFeatures.LineSkipPattern);
            int currentIndex = tableParameters.FirstCharIndex;

            List<string[]> dataList = new();
            while (currentIndex < tableParameters.LastCharIndex)
            {
                string[] tempRow = GetNextRowCells(text, tableParameters.BeginColumnIndexes, isCheckSkipOn, ref currentIndex);
                if (IsEmptyRow(tempRow))
                {
                    continue;
                }
                dataList.Add(tempRow);
            }
            _lastIndex = currentIndex;

            return dataList.ToArray();
        }

        /// <summary>
        /// Получение ячеек строки.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <param name="isCheckSkipOn"></param>
        /// <param name="currentIndex"></param>
        /// <returns></returns>
        private string[] GetNextRowCells(string text, int[] beginColumnIndexes, bool isCheckSkipOn, ref int currentIndex)
        {
            string textLine = "";

            int endLineIndex = text.IndexOf(Environment.NewLine, currentIndex);
            if (endLineIndex >= 0)
            {
                textLine = text[currentIndex..endLineIndex];
            }
            else
            {
                // Берем последний символ, если не найден перенос.
                endLineIndex = text.Length - 1;
                textLine = text[currentIndex..];
            }

            currentIndex = endLineIndex + Environment.NewLine.Length;

            // TODO: Тут можно, конечно же, лучше замутить,
            // но пускай пока так, не будем усложнять еще как-то.
            if (isCheckSkipOn && Regex.IsMatch(textLine, _tableDetector.DetectFeatures.LineSkipPattern))
            {
                return Enumerable.Repeat("", beginColumnIndexes.Length).ToArray();
            }
            return ExtractCellsData(textLine, beginColumnIndexes);
        }

        /// <summary>
        /// Извлечение данных ячеек.
        /// </summary>
        /// <param name="textLine"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns></returns>
        private string[] ExtractCellsData(string textLine, int[] beginColumnIndexes)
        {
            int columnsLength = beginColumnIndexes.Length;
            int textLineLength = textLine.Length;
            string[] row = new string[columnsLength];

            if (textLineLength > 0)
            {
                for (int i = 0; i < columnsLength - 1; i++)
                {
                    int beginIndex = beginColumnIndexes[i];
                    int endIndex = beginColumnIndexes[i + 1] - 1;
                    if (beginIndex < textLineLength && endIndex < textLineLength)
                    {
                        // Нормальный вариант
                        row[i] = textLine[beginColumnIndexes[i]..beginColumnIndexes[i + 1]].Trim();
                    }
                    else if (beginIndex < textLineLength)
                    {
                        // Если индекс окончания выходит за пределы
                        row[i] = textLine[beginColumnIndexes[i]..].Trim();
                    }
                    else
                    {
                        // Если совсем тоска печаль и всё за пределы
                        row[i] = "";
                    }
                }
                row[^1] = beginColumnIndexes[^1] < textLineLength
                    ? textLine[beginColumnIndexes[^1]..].Trim()
                    : "";
            }
            else
            {
                row = Enumerable.Repeat("", columnsLength).ToArray();
            }

            return row;
        }

        /// <summary>
        /// Соединение строк.
        /// </summary>
        /// <param name="rows"></param>
        /// <returns></returns>
        private string[] ConcatRows(string[][] rows)
        {
            string[] commonRow = new string[rows[0].Length];
            for (int i = 0; i < commonRow.Length; i++)
            {
                commonRow[i] = "";
            }

            foreach (var row in rows)
            {
                for (int i = 0; i < commonRow.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(commonRow[i]))
                    {
                        commonRow[i] += Environment.NewLine + row[i];
                    }
                    else
                    {
                        commonRow[i] = row[i];
                    }
                }
            }

            return commonRow;
        }

        /// <summary>
        /// Это пустая строка.
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private bool IsEmptyRow(string[] row)
        {
            bool isEmpty = true;

            if (row.Length > 0)
            {
                foreach (var item in row)
                {
                    if (!string.IsNullOrWhiteSpace(item))
                    {
                        isEmpty = false;
                        break;
                    }
                }
            }

            return isEmpty;
        }

        /// <summary>
        /// Это начало новой строки?
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private bool IsBeginNewTableRow(string[] row)
        {
            return string.IsNullOrWhiteSpace(_tableDetector.DetectFeatures.FirstBodyRowCellWordPattern)
                ? !string.IsNullOrWhiteSpace(row[0])
                : Regex.IsMatch(row[0], _tableDetector.DetectFeatures.FirstBodyRowCellWordPattern);
        }

        /// <summary>
        /// Расчитать параметры ячеек заголовка по паттернам.
        /// </summary>
        /// <param name="rawData"></param>
        /// <param name="headerCellPatterns"></param>
        /// <returns></returns>
        private CalcHeaderCellsByPatternsResult CalcHeaderCellsByPatterns(string[][] rawData, List<string> headerCellPatterns)
        {
            List<string> normalHeaderCellPatterns = headerCellPatterns.Select(x => Regex.Replace(x, @"\s{2,}", " ")).ToList();
            int columnCount = rawData[0].Length;
            int currentColumnIndex = 0;
            List<CheckHeaderCellByPatternResult> checkHeaderCellByPatternResults = new();
            foreach (var pattern in normalHeaderCellPatterns)
            {
                var checkResult = CheckHeaderCellByPattern(rawData, pattern, ref currentColumnIndex);
                if (checkResult.IsSucces)
                {
                    checkHeaderCellByPatternResults.Add(checkResult);
                }
                else
                {
                    break;
                }
            }
            bool isSucces = checkHeaderCellByPatternResults.Count == normalHeaderCellPatterns.Count;
            CalcHeaderCellsByPatternsResult result = new(isSucces,
                isSucces
                ? checkHeaderCellByPatternResults.ToArray()
                : Array.Empty<CheckHeaderCellByPatternResult>());
            return result;
        }

        /// <summary>
        /// Проверка ячеейки заголовка по паттерну.
        /// </summary>
        /// <param name="rawData"></param>
        /// <param name="headerCellPattern"></param>
        /// <param name="currentColumnIndex"></param>
        /// <returns></returns>
        private CheckHeaderCellByPatternResult CheckHeaderCellByPattern(string[][] rawData, string headerCellPattern, ref int currentColumnIndex)
        {
            CheckHeaderCellByPatternResult result = new(false, "", -1, -1, -1, -1);
            int columnCount = rawData[0].Length;
            int rowCount = rawData.Length;
            for (int x1 = currentColumnIndex; x1 < columnCount; x1++)
            {
                currentColumnIndex++;
                for (int x2 = x1; x2 < columnCount; x2++)
                {
                    for (int y1 = 0; y1 < rowCount; y1++)
                    {
                        for (int y2 = y1; y2 < rowCount; y2++)
                        {
                            int subLen1 = y2 - y1 + 1;
                            int subLen2 = x2 - x1 + 1;
                            string[][] testData = new string[subLen1][];
                            for (int i = y1; i <= y2; i++)
                            {
                                string[] row = new string[subLen2];
                                for (int k = x1; k <= x2; k++)
                                {
                                    row[k - x1] = rawData[i][k];
                                }
                                testData[i - y1] = row;
                            }
                            string testText = GetSingleTestCellText(testData);
                            if (Regex.IsMatch(testText, headerCellPattern))
                            {
                                currentColumnIndex = x2 + 1;
                                result = new(
                                    isSucces: true,
                                    cellText: testText,
                                    beginRowIndex: y1,
                                    beginColumnIndex: x1,
                                    endRowIndex: y2,
                                    endColumnIndex: x2);
                                return result;
                            }
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Получить объединенный текст с нескольких ячеек.
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private string GetSingleTestCellText(string[][] data)
        {
            StringBuilder builder = new();

            foreach (var row in data)
            {
                foreach (var cell in row)
                {
                    string cellText = Regex.Replace(cell, @"\s{2,}", " ").Trim() + ' ';
                    builder.Append(cellText);
                }
            }

            return builder.ToString().Trim();
        }
    }
}

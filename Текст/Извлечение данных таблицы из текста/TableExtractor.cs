using System;
using System.Collections.Generic;
using System.Linq;
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

        private string[,] _data = new string[0,0];

        private int _lastIndex = -1;

        public TableExtractor() {}

        public TableExtractor(TableDetectFeatures detectFeatures)
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
        public string[,] Data => _data;

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
                int len0 = _data.GetLength(0);
                int len1 = _data.GetLength(1);
                string[][] arrayOfarray = new string[len0][];
                for (int i = 0; i < len0; i++)
                {
                    string[] row = new string[len1];
                    for (int k = 0; k < len1; k++)
                    {
                        row[k] = _data[i, k];
                    }
                    arrayOfarray[i] = row;
                }
                return JsonSerializer.Serialize(arrayOfarray, options);
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
            bool isCheckSkipOn = !string.IsNullOrWhiteSpace(_tableDetector.DetectFeatures.LineSkipPattern);

            int columnsLength = tableParameters.BeginColumnIndexes.Length;
            int currentIndex = tableParameters.FirstCharIndex;

            List<string[]> dataList = new();

            string[] row = Enumerable.Repeat("", columnsLength).ToArray();
            while (currentIndex < tableParameters.LastCharIndex)
            {
                string[] tempRow = GetNextRowCells(text, tableParameters.BeginColumnIndexes, isCheckSkipOn, ref currentIndex);
                if (IsEmptyRow(tempRow))
                {
                    continue;
                }
                if (!string.IsNullOrWhiteSpace(tempRow[0]))
                {
                    if (!string.IsNullOrWhiteSpace(row[0]))
                    {
                        row = row.Select(x => x.Trim()).ToArray();
                        dataList.Add(row);
                    }
                    row = Enumerable.Repeat("", columnsLength).ToArray();
                }
                row = ConcatRows(new string[][] { row, tempRow });
            }
            if (!IsEmptyRow(row))
            {
                row = row.Select(x => x.Trim()).ToArray();
                dataList.Add(row);
            }
            _lastIndex = currentIndex;

            _data = new string[dataList.Count, columnsLength];
            int r = 0;
            foreach (var item in dataList)
            {
                for (int k = 0; k < columnsLength; k++)
                {
                    _data[r, k] = item[k];
                }
                r++;
            }

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
        public IEnumerable<string[,]> GetAll(string text)
        {
            Stack<string[,]> data = new();

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
            _data = new string[0, 0];
            _lastIndex = -1;
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
    }
}

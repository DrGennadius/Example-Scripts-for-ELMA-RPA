using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Извлекатель таблицы из текста.
    /// </summary>
    public class TableExtractor
    {
        readonly TableDetector _tableDetector = new();

        private string[,] _data = new string[0,0];

        public TableExtractor() {}

        public TableExtractor(TableDetectFeatures detectFeatures)
        {
            _tableDetector = new(detectFeatures);
        }

        public TableExtractor(TableDetector detectFeatures)
        {
            _tableDetector = detectFeatures;
        }

        public string[,] Data => _data;

        /// <summary>
        /// Извлечение данных таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public bool Extract(string text)
        {
            var tableParameters = _tableDetector.Detect(text);
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
            int currentIndex = tableParameters.FirstCharIndex;

            List<string[]> dataList = new();

            string[] row = new string[columnsLength];
            for (int i = 0; i < row.Length; i++)
            {
                row[i] = "";
            }
            while (currentIndex < tableParameters.LastCharIndex)
            {
                string[] tempRow = GetNextRowCells(text, tableParameters.BeginColumnIndexes, ref currentIndex);
                if (!string.IsNullOrWhiteSpace(tempRow[0]))
                {
                    if (!string.IsNullOrWhiteSpace(row[0]))
                    {
                        dataList.Add(row);
                    }
                    row = new string[columnsLength];
                    for (int i = 0; i < row.Length; i++)
                    {
                        row[i] = "";
                    }
                }
                row = ConcatRows(new string[][] { row, tempRow });
            }
            if (!string.IsNullOrWhiteSpace(row[0]))
            {
                dataList.Add(row);
            }

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

            return isSucces;
        }

        /// <summary>
        /// Получение ячеек строки.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <param name="currentIndex"></param>
        /// <returns></returns>
        private string[] GetNextRowCells(string text, int[] beginColumnIndexes, ref int currentIndex)
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

            for (int i = 0; i < columnsLength - 1; i++)
            {
                row[i] = beginColumnIndexes[i + 1] < textLineLength
                    ? textLine[beginColumnIndexes[i]..beginColumnIndexes[i + 1]].Trim()
                    : "";
            }
            row[^1] = beginColumnIndexes[^1] < textLineLength
                ? textLine[beginColumnIndexes[^1]..].Trim()
                : "";

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
    }
}

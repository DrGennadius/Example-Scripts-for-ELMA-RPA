﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Поисковик/определитель таблицы в тексте.
    /// </summary>
    public class TableDetector
    {
        readonly TableDetectFeatures _detectFeatures = new();

        public TableDetector() {}

        public TableDetector(TableDetectFeatures detectFeatures)
        {
            _detectFeatures = detectFeatures;
        }

        /// <summary>
        /// Определение таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public TableParameters? Detect(string text)
        {
            TableParameters? tableParameters = null;

            if (string.IsNullOrWhiteSpace(text))
            {
                return tableParameters;
            }

            int beginIndex = DetectFirstCharIndex(ref text);
            if (beginIndex == -1)
            {
                return tableParameters;
            }

            int currentIndex = beginIndex;
            var beginColumnIndexes = DetectBeginColumnIndexes(ref text, ref currentIndex);
            int endIndex = PassToEndTable(ref text, ref currentIndex, ref beginColumnIndexes);

            tableParameters = new TableParameters(beginIndex, endIndex, beginColumnIndexes);

            return tableParameters;
        }

        /// <summary>
        /// Получить индекс последней строки таблицы по последнему индексу.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="lastIndex"></param>
        /// <returns></returns>
        public int LastTableRowIndex(ref string text, int lastIndex)
        {
            int newLineLength = Environment.NewLine.Length;
            string[] rows = text.Split(Environment.NewLine);
            int index = 0;
            int i = 0;
            for (; i < rows.Length; i++)
            {
                int tempIndex = index + rows[i].Length;
                if (tempIndex >= lastIndex)
                {
                    break;
                }
                index = tempIndex + newLineLength;
            }
            return i;
        }

        /// <summary>
        /// Определение индекса первого символа.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private int DetectFirstCharIndex(ref string text)
        {
            if (!string.IsNullOrEmpty(_detectFeatures.FirstCellText))
            {
                return text.IndexOf(_detectFeatures.FirstCellText);
            }
            else
            {
                return DetectFirstCharIndexAuto(ref text);
            }
        }

        /// <summary>
        /// Автоматическое определение индекса первого символа по паттерну разделения.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private int DetectFirstCharIndexAuto(ref string text)
        {
            if (string.IsNullOrEmpty(_detectFeatures.SplitPattern))
            {
                throw new Exception("SplitPattern не назначен.");
            }

            int index = -1;
            
            var match = Regex.Match(text, _detectFeatures.SplitPattern);
            if (match.Success)
            {
                string pattern = $"({Environment.NewLine})+(?={_detectFeatures.SplitPattern})";
                var beginMatch = Regex.Match(text, pattern);
                index = beginMatch.Success && beginMatch.Index <= match.Index
                    ? beginMatch.Index
                    : 0;
            }

            return index;
        }

        /// <summary>
        /// Определение индексов начал столбцов.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="currentIndex"></param>
        /// <returns></returns>
        private int[] DetectBeginColumnIndexes(ref string text, ref int currentIndex)
        {
            int endLineIndex = text.IndexOf(Environment.NewLine, currentIndex);
            string subText = text[currentIndex..endLineIndex];
            // Ищем разделители. Тут должно быть без пропусков по идее,
            // т.к. это первая строка хидера таблицы,
            // где должны быть наименования столбцов.
            var splitMatches = Regex.Matches(subText, _detectFeatures.SplitPattern);
            // Следующий текущий индекс это следующий символ после конца текущей строки.
            currentIndex = endLineIndex + Environment.NewLine.Length;
            // Расчитываем стартовые индексы начала столбцов
            // ic = начальный индекс разделителя + длина разделителя
            int[] beginIndexes = splitMatches.Select(x => x.Index + x.Value.Length).ToArray();
            // В самом начале еще вставляем самый первый индекс.
            return (new int[] { 0 }).Concat(beginIndexes).ToArray();
        }

        /// <summary>
        /// Дойти до конца таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="currentIndex"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns></returns>
        private int PassToEndTable(ref string text, ref int currentIndex, ref int[] beginColumnIndexes)
        {
            int endIndex = currentIndex;
#if DEBUG
            char debugBeginChar = text[currentIndex];
#endif

            while (currentIndex < text.Length)
            {
                int endLineIndex = text.IndexOf(Environment.NewLine, currentIndex);
                if (endLineIndex < 0)
                {
                    break;
                }
                string textLine = text[currentIndex..endLineIndex];
                if (!IsValidRow(ref textLine, ref beginColumnIndexes))
                {
                    break;
                }
                endIndex = currentIndex;
                // Устанавливаем следующий индекс начала следующим после конца этой строки.
                currentIndex = endLineIndex + Environment.NewLine.Length;
            }

            return endIndex;
        }

        /// <summary>
        /// Это валидная строка?
        /// </summary>
        /// <param name="textLine"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns></returns>
        private bool IsValidRow(ref string textLine, ref int[] beginColumnIndexes)
        {
            if (textLine == "")
            {
                // Пропускаем.
                // Дело в том, что у нас может быть текст, в котором есть пустые строки.
                return true;
            }
            if (string.IsNullOrWhiteSpace(textLine) || beginColumnIndexes.Length <= 0)
            {
                return false;
            }
            if (beginColumnIndexes.Length == 1)
            {
                return true;
            }

            bool isValid = IsValidRowBase(ref textLine, ref beginColumnIndexes);

            if (isValid && _detectFeatures.HasStartSequentialNumberingCells)
            {
                // Дополнительная проверка нумерации.
                isValid = IsValidRowWithNumberingStart(ref textLine, ref beginColumnIndexes);
            }

            return isValid;
        }

        /// <summary>
        /// Это валидная строка (базовая проверка)?
        /// </summary>
        /// <param name="textLine"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns></returns>
        private bool IsValidRowBase(ref string textLine, ref int[] beginColumnIndexes)
        {
            bool isValid = true;

            int beginIndex;
            int endIndex;

            for (int i = 1; i < beginColumnIndexes.Length; i++)
            {
                // Будем как бы проверять предыдущий фрагмент текста (ячейку),
                // которая находится перед индексом начала текущего фрагмента.
                beginIndex = beginColumnIndexes[i - 1];
                endIndex = beginColumnIndexes[i] - 1;
                isValid = IsValidFragment(ref textLine, beginIndex, endIndex);
                if (!isValid)
                {
                    break;
                }
            }
            if (isValid)
            {
                // Т.к. мы в цикле использовали предыдущий фрагмент,
                // то сейчас остается последний, где конец - последний индекс строки.
                beginIndex = beginColumnIndexes[^1];
                endIndex = textLine.Length - 1;
                isValid = IsValidFragment(ref textLine, beginIndex, endIndex);
            }

            return isValid;
        }

        private bool IsValidFragment(ref string textLine, int beginIndex, int endIndex)
        {
            if (beginIndex >= textLine.Length || endIndex >= textLine.Length)
            {
                // Предполагаем, что если уходим за пределы конца строки,
                // то там просто пустые значения.
                return true;
            }
            // Обрезаем фрагмент предыдущий.
            string subText = textLine[beginIndex..endIndex];
            // Проверяем на разрез. Не будем что-то усложнять/оптимизировать, будем следовать простой логике.
            // По идее фрагмент содержит то, что является частью значения и разделительные символы в конце.
            // Т.е. должны получить 2 элемента, 2й должен быть пустым или 1 элемент если в конце строки.
            var elements = Regex.Split(subText, _detectFeatures.SplitPattern);
            if (!((elements.Length == 2 && elements[1] == "") || (elements.Length == 1 && endIndex == textLine.Length - 1)))
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Это валидная строка (проверка нумерации в первой ячейке)?
        /// </summary>
        /// <param name="textLine"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns></returns>
        private bool IsValidRowWithNumberingStart(ref string textLine, ref int[] beginColumnIndexes)
        {
            bool isValid = true;

            if (beginColumnIndexes.Length == 0)
            {
                return false;
            }

            int beginIndex = beginColumnIndexes[0];
            int endIndex = beginColumnIndexes.Length > 1 ? beginColumnIndexes[1] - 1 : textLine.Length - 1;

            string subText = textLine[beginIndex..endIndex].Trim();
            if (!string.IsNullOrWhiteSpace(subText))
            {
                isValid = int.TryParse(subText, out _);
            }

            return isValid;
        }
    }
}

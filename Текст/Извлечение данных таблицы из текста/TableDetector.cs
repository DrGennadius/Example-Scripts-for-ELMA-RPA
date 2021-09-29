using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Поисковик/определитель таблицы в тексте.
    /// </summary>
    public class TableDetector
    {
        public TableFeatures DetectFeatures { get; set; }

        public TableDetector() { }

        public TableDetector(TableFeatures detectFeatures)
        {
            DetectFeatures = detectFeatures;
        }

        /// <summary>
        /// Определение таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="startIndex">Индекс начального символа</param>
        /// <returns></returns>
        public TableParameters? Detect(string text, int startIndex = 0)
        {
            TableParameters? tableParameters = null;

            if (string.IsNullOrWhiteSpace(text))
            {
                return tableParameters;
            }

            int beginIndex = DetectFirstCharIndex(text, startIndex);
            if (beginIndex == -1)
            {
                return tableParameters;
            }

            int currentIndex = beginIndex;
            var beginColumnIndexes = DetectBeginColumnIndexes(text, ref currentIndex);
            var fullBeginColumnIndexes = PassToEndTable(text, beginIndex, ref currentIndex, beginColumnIndexes);

            tableParameters = new TableParameters(beginIndex, currentIndex, fullBeginColumnIndexes);

            return tableParameters;
        }

        /// <summary>
        /// Определение следующей таблицы после указанной таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="prevTable">Таблица, после которой нужно искать</param>
        /// <returns></returns>
        public TableParameters? Detect(string text, TableParameters prevTable)
        {
            return Detect(text, prevTable.LastCharIndex + 1);
        }

        /// <summary>
        /// Определить все таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public TableParameters[] DetectAll(string text)
        {
            List<TableParameters> tablesParameters = new();

            TableParameters? tableParameters = Detect(text);

            if (tableParameters.HasValue)
            {
                tablesParameters.Add(tableParameters.Value);
            }

            while (tableParameters.HasValue)
            {
                tableParameters = Detect(text, tableParameters.Value);
                if (tableParameters.HasValue)
                {
                    tablesParameters.Add(tableParameters.Value);
                }
            }

            return tablesParameters.ToArray();
        }

        /// <summary>
        /// Получить индекс последней строки таблицы по последнему индексу.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="lastIndex"></param>
        /// <returns></returns>
        public int LastTableRowIndex(string text, int lastIndex)
        {
            int newLineLength = Environment.NewLine.Length;
            string[] rows = text.Split(Environment.NewLine);
            int index = 0;
            int i = 0;
            for (; i < rows.Length; i++)
            {
#if DEBUG
                string debugRow = rows[i];
#endif
                int tempIndex = index + rows[i].Length;
                if (tempIndex >= lastIndex)
                {
                    break;
                }
                index = tempIndex + newLineLength;
            }
            return i - 1;
        }

        /// <summary>
        /// Определение индекса первого символа.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="startIndex"></param>
        /// <returns></returns>
        private int DetectFirstCharIndex(string text, int startIndex = 0)
        {
            if (!string.IsNullOrEmpty(DetectFeatures.FirstTableCellWordPattern))
            {
#if DEBUG
                // Старый вариант через string.IndexOf
                int debugIndex = text.IndexOf(DetectFeatures.FirstTableCellWordPattern, startIndex);
#endif
                string subText = text[startIndex..];
                var match = Regex.Match(subText, DetectFeatures.FirstTableCellWordPattern);
                int index = match.Success ? startIndex + match.Index : -1;
                return index;
            }
            else
            {
                return DetectFirstCharIndexAuto(text);
            }
        }

        /// <summary>
        /// Автоматическое определение индекса первого символа по паттерну разделения.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private int DetectFirstCharIndexAuto(string text)
        {
            if (string.IsNullOrEmpty(DetectFeatures.SplitPattern))
            {
                throw new Exception("SplitPattern не назначен.");
            }

            int index = -1;

            var match = Regex.Match(text, DetectFeatures.SplitPattern);
            if (match.Success)
            {
                string pattern = $"({Environment.NewLine})+(?={DetectFeatures.SplitPattern})";
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
        private int[] DetectBeginColumnIndexes(string text, ref int currentIndex)
        {
            string subText = "";
            int endLineIndex = text.IndexOf(Environment.NewLine, currentIndex);
            if (endLineIndex >= 0)
            {
                subText = text[currentIndex..endLineIndex];
            }
            else
            {
                // Берем последний символ, если не найден перенос.
                endLineIndex = text.Length - 1;
                subText = text[currentIndex..];
            }
            // Ищем разделители. Тут должно быть без пропусков по идее,
            // т.к. это первая строка хидера таблицы,
            // где должны быть наименования столбцов.
            var splitMatches = Regex.Matches(subText, DetectFeatures.SplitPattern);
            // Следующий текущий индекс это следующий символ после конца текущей строки.
            currentIndex = endLineIndex + Environment.NewLine.Length;
            if (splitMatches.Count == 0)
            {
                // Если не было разделено, то значит возвращаем только один индекс.
                return new int[] { 0 };
            }
            // Расчитываем стартовые индексы начала столбцов
            // ic = начальный индекс разделителя + длина разделителя
            int[] beginIndexes = splitMatches.Select(x => x.Index + x.Value.Length).ToArray();
            // На всякий случай удалим последний, если там ничего нет.
            if (beginIndexes[^1] == endLineIndex)
            {
                beginIndexes = beginIndexes.SkipLast(1).ToArray();
            }
            // В самом начале еще вставляем самый первый индекс.
            return (new int[] { 0 }).Concat(beginIndexes).ToArray();
        }

        /// <summary>
        /// Дойти до конца таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="startIndex"></param>
        /// <param name="currentIndex"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns>Полный массив начал строк.</returns>
        private List<BeginColumnIndexesItem> PassToEndTable(string text, int startIndex, ref int currentIndex, int[] beginColumnIndexes)
        {
#if DEBUG
            char debugBeginChar = text[currentIndex];
#endif
            bool isCheckSkipOn = !string.IsNullOrWhiteSpace(DetectFeatures.LineSkipPattern);
            BeginColumnIndexesItem lastIndexes = new(startIndex, beginColumnIndexes);
            List<BeginColumnIndexesItem> fullBeginColumnIndexes = new();
            fullBeginColumnIndexes.Add(lastIndexes);
            string textLine = "";
            while (currentIndex < text.Length)
            {
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
                // Устанавливаем следующий индекс начала следующим после конца этой строки.
                currentIndex = endLineIndex + Environment.NewLine.Length;

                bool isSkip = isCheckSkipOn && Regex.IsMatch(textLine, DetectFeatures.LineSkipPattern);

                if (isSkip)
                {
                    var newBeginColumnIndexes = CalcNewBeginColumnIndexesVariant(text, beginColumnIndexes, ref currentIndex);
                    if (newBeginColumnIndexes.TextBeginCharIndex == -1 || newBeginColumnIndexes.BeginColumnIndexes.Length == 0)
                    {
                        break;
                    }
                    else
                    {
                        lastIndexes = newBeginColumnIndexes;
                        fullBeginColumnIndexes.Add(newBeginColumnIndexes);
                    }
                }
                else if (!IsValidRow(textLine, lastIndexes.BeginColumnIndexes))
                {
                    break;
                }
            }

            return fullBeginColumnIndexes;
        }

        /// <summary>
        /// Это валидная строка?
        /// </summary>
        /// <param name="textLine"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns></returns>
        private bool IsValidRow(string textLine, int[] beginColumnIndexes)
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

            bool isValid = IsValidRowBase(textLine, beginColumnIndexes);

            if (isValid && DetectFeatures.HasStartSequentialNumberingCells)
            {
                // Дополнительная проверка нумерации.
                isValid = IsValidRowWithNumberingStart(textLine, beginColumnIndexes);
            }

            return isValid;
        }

        /// <summary>
        /// Это валидная строка (базовая проверка)?
        /// </summary>
        /// <param name="textLine"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns></returns>
        private bool IsValidRowBase(string textLine, int[] beginColumnIndexes)
        {
            bool isValid = true;

#if DEBUG
            int debugTextLineLength = textLine.Length;
#endif

            int beginIndex;
            int endIndex;

            for (int i = 1; i < beginColumnIndexes.Length; i++)
            {
                // Будем как бы проверять предыдущий фрагмент текста (ячейку),
                // которая находится перед индексом начала текущего фрагмента.
                beginIndex = beginColumnIndexes[i - 1];
                endIndex = beginColumnIndexes[i];
                isValid = IsValidFragment(textLine, beginIndex, endIndex);
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
                endIndex = textLine.Length;
                isValid = IsValidFragment(textLine, beginIndex, endIndex);
            }

            return isValid;
        }

        private bool IsValidFragment(string textLine, int beginIndex, int endIndex)
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
            var elements = Regex.Split(subText, DetectFeatures.SplitPattern);
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
        private bool IsValidRowWithNumberingStart(string textLine, int[] beginColumnIndexes)
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

        /// <summary>
        /// Расчитать новый вариант индексов начал столбцов.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="currentIndex"></param>
        /// <param name="originBeginColumnIndexes"></param>
        /// <returns></returns>
        private BeginColumnIndexesItem CalcNewBeginColumnIndexesVariant(string text, int[] originBeginColumnIndexes, ref int currentIndex)
        {
            int[] beginColumnIndexes = Array.Empty<int>();
            int storedIndex = -1;

            while (currentIndex < text.Length)
            {
                storedIndex = currentIndex;
                beginColumnIndexes = DetectBeginColumnIndexes(text, ref currentIndex);
                if (beginColumnIndexes.Length == originBeginColumnIndexes.Length)
                {
                    break;
                }
            }

            return new(storedIndex, beginColumnIndexes);
        }
    }
}

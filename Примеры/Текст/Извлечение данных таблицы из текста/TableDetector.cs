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
        private const double NullСontinuousBodyRowCellTextK = 0.5;
        private const double NullRowCellK = 1.0;
        private const double ApplyRowPatternK = 0.25;
        private const double CommonSimilarRowMinK = 0.5;

        public TableDetector() { }

        public TableDetector(TableFeatures detectFeatures)
        {
            DetectFeatures = detectFeatures;
        }

        /// <summary>
        /// Признаки определения таблицы.
        /// </summary>
        public TableFeatures DetectFeatures { get; set; }

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
            beginIndexes = (new int[] { 0 }).Concat(beginIndexes).ToArray();
            return beginIndexes;
        }

        /// <summary>
        /// Дойти до конца таблицы.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="startIndex"></param>
        /// <param name="currentIndex"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns>Полный массив начал строк.</returns>
        private List<RowInfoItem> PassToEndTable(string text, int startIndex, ref int currentIndex, int[] beginColumnIndexes)
        {
#if DEBUG
            char debugBeginChar = text[currentIndex];
#endif
            bool isCheckSkipOn = !string.IsNullOrWhiteSpace(DetectFeatures.LineSkipPattern);
            RowInfoItem prevLastRowInfoItem = null;
            RowInfoItem lastRowInfoItem = new(startIndex, beginColumnIndexes);
            RowInfoItem newRowInfoItem = null;
            List<RowInfoItem> fullRowInfoItems = new();
            fullRowInfoItems.Add(lastRowInfoItem);
            string textLine = "";
            int emptyLineCount = 0;
            bool isSkipping = false;
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
                if (string.IsNullOrWhiteSpace(textLine))
                {
                    // Устанавливаем следующий индекс начала следующим после конца этой строки.
                    currentIndex += Environment.NewLine.Length;
                    emptyLineCount++;
                    continue;
                }
                emptyLineCount = 0;

                bool isSkipNow = isCheckSkipOn && Regex.IsMatch(textLine, DetectFeatures.LineSkipPattern);

                if (!isSkipNow)
                {
                    if (isSkipping)
                    {
                        newRowInfoItem = CalcNewRowInfoItemVariant(text, lastRowInfoItem, currentIndex);
                        if (newRowInfoItem.TextBeginCharIndex == -1 || newRowInfoItem.BeginColumnIndexes.Length == 0)
                        {
                            break;
                        }
                        if (CommonTableHelper.IsEmptyRow(lastRowInfoItem.СontinuousBodyRowCellTexts))
                        {
                            CalcBodyRowFeaturesForce(text, lastRowInfoItem, currentIndex);
                        }
                        lastRowInfoItem.GenerateBodyRowPattern();
                        prevLastRowInfoItem = lastRowInfoItem;
                        lastRowInfoItem = newRowInfoItem;
                        fullRowInfoItems.Add(lastRowInfoItem);
                    }
                    var result = ValidationRow(textLine, prevLastRowInfoItem, lastRowInfoItem);
                    if (result.IsValid)
                    {
#if DEBUG
                        if (newRowInfoItem != null)
                        {
                            string debugPartOfText = "";
                            int debugEndLineIndex = text.IndexOf(Environment.NewLine, newRowInfoItem.TextBeginCharIndex);
                            if (debugEndLineIndex >= 0)
                            {
                                debugPartOfText = text[newRowInfoItem.TextBeginCharIndex..debugEndLineIndex];
                            }
                            else
                            {
                                // Берем последний символ, если не найден перенос.
                                debugEndLineIndex = text.Length - 1;
                                debugPartOfText = text[newRowInfoItem.TextBeginCharIndex..];
                            }
                            if (!string.IsNullOrWhiteSpace(debugPartOfText))
                            {
                                string[] debugCells = GetRowCellsForce(debugPartOfText, newRowInfoItem);
                            }
                        }
#endif
                        if (result.CorrectRowInfo != null && result.CorrectRowInfo.IsAutoCorrected)
                        {
                            lastRowInfoItem = result.CorrectRowInfo.NewRowInfoItem;
                        }
                        // Копим признаки для дальнейшего сравнения.
                        CalcBodyRowFeatures(textLine, lastRowInfoItem);
                    }
                    else
                    {
                        break;
                    }
                }
                // Устанавливаем следующий индекс начала следующим после конца этой строки.
                currentIndex = endLineIndex + Environment.NewLine.Length;
                emptyLineCount++;
                isSkipping = isSkipNow;
            }

            // Компенсируем переход на пустые строки в конце.
            int dropEndChars = Environment.NewLine.Length * emptyLineCount;
            if (dropEndChars > 0)
            {
                currentIndex -= dropEndChars;
            }

            return fullRowInfoItems;
        }

        /// <summary>
        /// Это валидная строка?
        /// </summary>
        /// <param name="textLine"></param>
        /// <param name="prevRowInfoItem"></param>
        /// <param name="rowInfoItem"></param>
        /// <returns></returns>
        private ValidationRowResult ValidationRow(string textLine, RowInfoItem prevRowInfoItem, RowInfoItem rowInfoItem)
        {
            ValidationRowResult result = new()
            {
                IsValid = false
            };

            if (textLine == "")
            {
                // Пропускаем.
                // Дело в том, что у нас может быть текст, в котором есть пустые строки.
                result.IsValid = true;
                return result;
            }
            if (string.IsNullOrWhiteSpace(textLine) || rowInfoItem.BeginColumnIndexes.Length <= 0)
            {
                return result;
            }
            if (rowInfoItem.BeginColumnIndexes.Length == 1)
            {
                result.IsValid = true;
                return result;
            }

            bool isValid = IsValidRowBase(textLine, rowInfoItem.BeginColumnIndexes);
            // prevRowInfoItem для случаев, если после разреза страницы сразу же
            // какой-нибудь косяк и можно было бы проверить по сохраненным данным с прошлой страницы.
            bool isUsePrevRowInfoItem = rowInfoItem == null
                || (string.IsNullOrWhiteSpace(rowInfoItem.BodyRowPattern)
                && CommonTableHelper.IsEmptyRow(rowInfoItem.СontinuousBodyRowCellTexts));
            RowInfoItem tragetRowInfoItem = isUsePrevRowInfoItem
                ? prevRowInfoItem
                : rowInfoItem;
            if (!isValid
                && tragetRowInfoItem != null
                && (!string.IsNullOrWhiteSpace(tragetRowInfoItem.BodyRowPattern)
                || !CommonTableHelper.IsEmptyRow(tragetRowInfoItem.СontinuousBodyRowCellTexts)))
            {
                // Перепроверяем с неопределенной увереностью, что строка всё таки может быть валидной.
                CorrectRowInfoItem correctRowInfoItem = CorrectRowInfo(textLine, tragetRowInfoItem, isUsePrevRowInfoItem);
                isValid = correctRowInfoItem.IsValid;
                result.CorrectRowInfo = correctRowInfoItem;
            }

            if (isValid && DetectFeatures.HasStartSequentialNumberingCells)
            {
                // Дополнительная проверка нумерации.
                isValid = IsValidRowWithNumberingStart(textLine, rowInfoItem.BeginColumnIndexes);
            }

            result.IsValid = isValid;
            return result;
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
                // TODO: Можно какую-нибудь нечеткую проверку добавить с предыдущими строками.
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
        /// <param name="lastRowInfoItem"></param>
        /// <returns></returns>
        private RowInfoItem CalcNewRowInfoItemVariant(string text, RowInfoItem lastRowInfoItem, int currentIndex)
        {
            int[] beginColumnIndexes = Array.Empty<int>();
            RowInfoItem rowInfoItem = new(currentIndex, beginColumnIndexes);

            int prevCurrentIndex = currentIndex;
            while (currentIndex < text.Length)
            {
                beginColumnIndexes = DetectBeginColumnIndexes(text, ref currentIndex);
                if (beginColumnIndexes.Length == lastRowInfoItem.BeginColumnIndexes.Length)
                {
                    // Найден предположительно.
                    // TODO: Тут потом можно еще всякие проверочки добавить.
                    // А еще может быть ситуация когда количество столбцов изначально может быть некорректным.
                    rowInfoItem.BeginColumnIndexes = beginColumnIndexes;
                    rowInfoItem.СontinuousBodyRowCellTexts = Enumerable.Repeat("", beginColumnIndexes.Length).ToArray();
                    break;
                }
                else if (beginColumnIndexes.Length == lastRowInfoItem.BeginColumnIndexes.Length + 1
                    && beginColumnIndexes[0] == 0)
                {
                    // Проверка на смещение, из-за которого в начале может получится "пропуск",
                    // который опознан как разделитель
                    int startIndex = rowInfoItem.TextBeginCharIndex;
                    int endIndex = startIndex + beginColumnIndexes[1];
                    string subText = text[startIndex..endIndex];
                    if (string.IsNullOrWhiteSpace(subText))
                    {
                        beginColumnIndexes = beginColumnIndexes.Skip(1).ToArray();
                        startIndex = prevCurrentIndex;
                        endIndex = currentIndex - Environment.NewLine.Length;
                        string rowText = text[startIndex..endIndex];
                        double k = CalcSimilarRowK(GetPrimitiveCellText(rowText), beginColumnIndexes, lastRowInfoItem.СontinuousBodyRowCellTexts);
                        if (k >= CommonSimilarRowMinK)
                        {
                            rowInfoItem.BeginColumnIndexes = beginColumnIndexes;
                            rowInfoItem.СontinuousBodyRowCellTexts = Enumerable.Repeat("", beginColumnIndexes.Length).ToArray();
                            break;
                        }
                    }
                }
                prevCurrentIndex = currentIndex;
            }

            return rowInfoItem;
        }

        /// <summary>
        /// Намерено расчитать признаки строки тела таблицы по тому, что есть.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="rowInfoItem"></param>
        /// <param name="allowOverBoundAsEmpty"></param>
        private void CalcBodyRowFeaturesForce(string text, RowInfoItem rowInfoItem, int limitIndex)
        {
            bool isCheckSkipOn = !string.IsNullOrWhiteSpace(DetectFeatures.LineSkipPattern);
            int currentIndex = rowInfoItem.TextBeginCharIndex;
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
            // Устанавливаем следующий индекс начала следующим после конца этой строки.
            currentIndex = endLineIndex + Environment.NewLine.Length;
            while (currentIndex < limitIndex)
            {
                endLineIndex = text.IndexOf(Environment.NewLine, currentIndex);
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
                if (string.IsNullOrWhiteSpace(textLine))
                {
                    continue;
                }

                bool isSkipNow = isCheckSkipOn && Regex.IsMatch(textLine, DetectFeatures.LineSkipPattern);

                if (!isSkipNow)
                {
                    CalcBodyRowFeatures(textLine, rowInfoItem, true);
                }
            }
        }

        /// <summary>
        /// Расчитать признаки строки тела таблицы.
        /// По факту пока расчитывается по строке, которая имеет вхождениие в диапазон начальных индексов.
        /// </summary>
        /// <param name="subText"></param>
        /// <param name="rowInfoItem"></param>
        /// <param name="allowOverBoundAsEmpty"></param>
        private void CalcBodyRowFeatures(string subText, RowInfoItem rowInfoItem, bool allowOverBoundAsEmpty = false)
        {
            if (string.IsNullOrWhiteSpace(subText))
            {
                return;
            }

            int startIndex = -1;
            int endIndex = -1;
            string[] cells = new string[rowInfoItem.BeginColumnIndexes.Length];
            for (int i = 0; i < rowInfoItem.BeginColumnIndexes.Length - 1; i++)
            {
                startIndex = rowInfoItem.BeginColumnIndexes[i];
                endIndex = rowInfoItem.BeginColumnIndexes[i + 1] - 1;
                if (allowOverBoundAsEmpty)
                {
                    if (startIndex >= subText.Length)
                    {
                        cells[i] = "";
                    }
                    else if (endIndex >= subText.Length)
                    {
                        cells[i] = subText[rowInfoItem.BeginColumnIndexes[i]..];
                    }
                    else
                    {
                        cells[i] = subText[rowInfoItem.BeginColumnIndexes[i]..(rowInfoItem.BeginColumnIndexes[i + 1] - 1)];
                    }
                }
                else if (startIndex >= subText.Length || endIndex >= subText.Length)
                {
                    return;
                }
                else
                {
                    cells[i] = subText[rowInfoItem.BeginColumnIndexes[i]..(rowInfoItem.BeginColumnIndexes[i + 1] - 1)];
                }
            }
            startIndex = rowInfoItem.BeginColumnIndexes[^1];
            if (allowOverBoundAsEmpty)
            {
                if (startIndex >= subText.Length)
                {
                    cells[^1] = "";
                }
                else
                {
                    cells[^1] = subText[rowInfoItem.BeginColumnIndexes[^1]..];
                }
            }
            else if (startIndex >= subText.Length || endIndex >= subText.Length)
            {
                return;
            }
            else
            {
                cells[^1] = subText[rowInfoItem.BeginColumnIndexes[^1]..];
            }
            cells = cells.Select(x => GetPrimitiveCellText(x)).ToArray();

            for (int i = 0; i < cells.Length; i++)
            {
                string continuousCellText = GetСontinuousCellText(rowInfoItem.СontinuousBodyRowCellTexts[i], cells[i]);
                rowInfoItem.СontinuousBodyRowCellTexts[i] = continuousCellText;
            }
        }

        /// <summary>
        /// Получить примитивный текст. На текущий момент только заменяет числа на 0.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string GetPrimitiveCellText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return text;
            }

            char[] result = text.ToArray();

            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsDigit(text[i]))
                {
                    result[i] = '0';
                }
            }

            return new string(result);
        }

        /// <summary>
        /// Получить продолжающий текст. Т.е. 
        /// Оставляет общие символы двух строк.
        /// Т.е. по факту работает как пересечение массивов символов, 
        /// но при этом вставляет пустые символы в местах различий.
        /// </summary>
        /// <param name="text1"></param>
        /// <param name="text2"></param>
        /// <returns></returns>
        private string GetСontinuousCellText(string text1, string text2)
        {
            if (string.IsNullOrWhiteSpace(text1) && string.IsNullOrWhiteSpace(text2))
            {
                return "";
            }
            if (string.IsNullOrWhiteSpace(text1))
            {
                return text2;
            }
            if (string.IsNullOrWhiteSpace(text2))
            {
                return text1;
            }

            char[] charArray1 = text1.ToArray();
            char[] charArray2 = text2.ToArray();

            int len1 = charArray1.Length;
            int len2 = charArray2.Length;

            int lenDiff = len2 - len1;

            if (lenDiff >= 0)
            {
                for (int i = 0; i < len1; i++)
                {
                    if (charArray1[i] != charArray2[i])
                    {
                        charArray2[i] = ' ';
                    }
                }
                for (int i = len1; i < len1 + lenDiff; i++)
                {
                    charArray2[i] = ' ';
                }
                return new string(charArray2);
            }
            else
            {
                for (int i = 0; i < len2; i++)
                {
                    if (charArray2[i] != charArray1[i])
                    {
                        charArray1[i] = ' ';
                    }
                }
                for (int i = len2; i < len2 - lenDiff; i++)
                {
                    charArray1[i] = ' ';
                }
                return new string(charArray1);
            }
        }

        /// <summary>
        /// Это может быть валидная строка, либо исправленная?
        /// </summary>
        /// <param name="rowText"></param>
        /// <param name="rowInfoItem"></param>
        /// <param name="isUsePrevRowInfoItem"></param>
        /// <returns></returns>
        private CorrectRowInfoItem CorrectRowInfo(string rowText, RowInfoItem rowInfoItem, bool isUsePrevRowInfoItem)
        {
            CorrectRowInfoItem correctRowInfoItem = new()
            {
                OriginRowInfoItem = rowInfoItem
            };
            // 1. Первым делом проверим строку по паттерну.
            double applyPatternK = 0.0;
            if (!string.IsNullOrWhiteSpace(rowInfoItem.BodyRowPattern)
                && Regex.IsMatch(rowText, rowInfoItem.BodyRowPattern))
            {
                applyPatternK = ApplyRowPatternK;
            }
            // 2. Считаем схожесть по продолжительным символам.
            double similarRowK = 0.0;
            if (!CommonTableHelper.IsEmptyRow(rowInfoItem.СontinuousBodyRowCellTexts))
            {
                if (isUsePrevRowInfoItem)
                {
                    correctRowInfoItem = AutoCorrectRowInfo(rowText, rowInfoItem, correctRowInfoItem);
                    similarRowK = correctRowInfoItem.SimilarCoef;
                }
                else
                {
                    similarRowK = CalcSimilarRowK(rowText, rowInfoItem);
                    correctRowInfoItem.SimilarCoef = similarRowK;
                }
            }

            double k = (applyPatternK + similarRowK) / 2;
            correctRowInfoItem.IsValid = k >= CommonSimilarRowMinK;

            if (correctRowInfoItem.IsAutoCorrected)
            {
                correctRowInfoItem.NewRowInfoItem.TextBeginCharIndex = rowInfoItem.TextBeginCharIndex;
            }

            return correctRowInfoItem;
        }

        /// <summary>
        /// Посчитать коэфицент схожести строки.
        /// </summary>
        /// <param name="rowText"></param>
        /// <param name="rowInfoItem"></param>
        /// <returns></returns>
        private double CalcSimilarRowK(string rowText, RowInfoItem rowInfoItem)
        {
            string[] rowCells = GetRowCellsForce(rowText, rowInfoItem);

            return CalcSimilarRowK(rowCells, rowInfoItem.СontinuousBodyRowCellTexts);
        }

        /// <summary>
        /// Посчитать коэфицент схожести строки.
        /// </summary>
        /// <param name="rowText"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <param name="continuousBodyRowCellTexts"></param>
        /// <returns></returns>
        private double CalcSimilarRowK(string rowText, int[] beginColumnIndexes, string[] continuousBodyRowCellTexts)
        {
            string[] rowCells = GetRowCellsForce(rowText, beginColumnIndexes);

            return CalcSimilarRowK(rowCells, continuousBodyRowCellTexts);
        }

        /// <summary>
        /// Посчитать коэфицент схожести строки.
        /// </summary>
        /// <param name="rowCells"></param>
        /// <param name="continuousBodyRowCellTexts"></param>
        /// <returns></returns>
        private double CalcSimilarRowK(string[] rowCells, string[] continuousBodyRowCellTexts)
        {
            double k = 0.0;

            if (continuousBodyRowCellTexts.Length > 0)
            {
                for (int i = 0; i < continuousBodyRowCellTexts.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(continuousBodyRowCellTexts[i]))
                    {
                        k += NullСontinuousBodyRowCellTextK;
                    }
                    if (string.IsNullOrWhiteSpace(rowCells[i]))
                    {
                        k += NullRowCellK;
                    }
                    else
                    {
                        int maxLen = Math.Max(rowCells[i].Length, continuousBodyRowCellTexts[i].Length);
                        int d = LevenshteinHelper.GetDistance(rowCells[i], continuousBodyRowCellTexts[i]);
                        k += ((double)(maxLen - d)) / maxLen;
                    }
                }

                k /= continuousBodyRowCellTexts.Length;
            }

            return k;
        }

        /// <summary>
        /// Получить в любом случае ячейки строки, пусть они и пустые будут.
        /// </summary>
        /// <param name="rowText"></param>
        /// <param name="beginColumnIndexes"></param>
        /// <returns></returns>
        private string[] GetRowCellsForce(string rowText, int[] beginColumnIndexes)
        {
            int startIndex = -1;
            int endIndex = -1;
            string[] cells = new string[beginColumnIndexes.Length];
            for (int i = 0; i < beginColumnIndexes.Length - 1; i++)
            {
                startIndex = beginColumnIndexes[i];
                endIndex = beginColumnIndexes[i + 1] - 1;
                if (startIndex >= rowText.Length)
                {
                    cells[i] = "";
                }
                else if (endIndex >= rowText.Length)
                {
                    cells[i] = rowText[beginColumnIndexes[i]..];
                }
                else
                {
                    cells[i] = rowText[beginColumnIndexes[i]..(beginColumnIndexes[i + 1] - 1)];
                }
            }
            startIndex = beginColumnIndexes[^1];
            if (startIndex >= rowText.Length)
            {
                cells[^1] = "";
            }
            else
            {
                cells[^1] = rowText[beginColumnIndexes[^1]..];
            }
            return cells;
        }

        /// <summary>
        /// Получить в любом случае ячейки строки, пусть они и пустые будут.
        /// </summary>
        /// <param name="rowText"></param>
        /// <param name="rowInfoItem"></param>
        /// <returns></returns>
        private string[] GetRowCellsForce(string rowText, RowInfoItem rowInfoItem)
        {
            return GetRowCellsForce(rowText, rowInfoItem.BeginColumnIndexes);
        }

        /// <summary>
        /// Автоматическое определение ячеек и создание новой информации о строке.
        /// </summary>
        /// <param name="rowText"></param>
        /// <param name="rowInfoItem"></param>
        /// <param name="correctRowInfoItem"></param>
        /// <returns></returns>
        private CorrectRowInfoItem AutoCorrectRowInfo(string rowText, RowInfoItem rowInfoItem, CorrectRowInfoItem correctRowInfoItem)
        {
            double k = 0.0;

            // Сначала пробуем харкорное обычное
            string[] cells = GetRowCellsForce(rowText, rowInfoItem);

            // Проверяем немного
            bool isOk = true;

            // Примитивная проверка на то, что первые символы не пробелы в значимых ячейках.
            foreach (var cell in cells)
            {
                if (!string.IsNullOrWhiteSpace(cell) && cell[0] == ' ')
                {
                    isOk = false;
                    break;
                }
            }

            if (isOk)
            {
                k = CalcSimilarRowK(cells.Select(x => GetPrimitiveCellText(x)).ToArray(), rowInfoItem.СontinuousBodyRowCellTexts);
            }
            CorrectRowInfoItem autoCorrectRowInfoItem = GetAutoDetectRowCellsAutoAndCreateNewRowInfo(rowText, rowInfoItem, correctRowInfoItem);
            if (autoCorrectRowInfoItem.SimilarCoef > k)
            {
                correctRowInfoItem = autoCorrectRowInfoItem;
                // * 2 для увеличения шансов при автоматическом режиме.
                correctRowInfoItem.SimilarCoef *= 2;
            }
            else
            {
                correctRowInfoItem.SimilarCoef = k;
            }

            return correctRowInfoItem;
        }

        /// <summary>
        /// Автоопределение ячеек и новых данных о строке.
        /// </summary>
        /// <param name="rowText"></param>
        /// <param name="rowInfoItem"></param>
        /// <param name="correctRowInfoItem"></param>
        /// <returns></returns>
        private CorrectRowInfoItem GetAutoDetectRowCellsAutoAndCreateNewRowInfo(string rowText, RowInfoItem rowInfoItem, CorrectRowInfoItem correctRowInfoItem)
        {
            RowInfoItem newRowInfoItem = null;

            string[] cells = Enumerable.Repeat("", rowInfoItem.BeginColumnIndexes.Length).ToArray();

            int currentIndex = 0;
            int[] defaultBeginColumnIndexes = DetectBeginColumnIndexes(rowText, ref currentIndex);

            int diffLen = defaultBeginColumnIndexes.Length - rowInfoItem.BeginColumnIndexes.Length;

            double bestK = 0.0;
            int[] bestBeginIndexesArray = Array.Empty<int>();

            if (diffLen == 1)
            {
                newRowInfoItem = new RowInfoItem(rowInfoItem);
                List<int> tempBeginIndexes = new();
                for (int m = 1; m < defaultBeginColumnIndexes.Length - 1; m++)
                {
                    tempBeginIndexes.Clear();
                    // m - индекс мержа
                    for (int i = 0; i < defaultBeginColumnIndexes.Length; i++)
                    {
                        if (i == m)
                        {
                            continue;
                        }
                        tempBeginIndexes.Add(defaultBeginColumnIndexes[i]);
                    }
                    int[] tempBeginIndexesArray = tempBeginIndexes.ToArray();
                    string[] tempCells = GetRowCellsForce(rowText, tempBeginIndexesArray);
                    // TODO: Надо бы хранить где-то схожесть, чтобы потом не пересчитывать по новой.
                    double tempK = CalcSimilarRowK(
                        tempCells.Select(x => GetPrimitiveCellText(x)).ToArray(),
                        rowInfoItem.СontinuousBodyRowCellTexts);
                    if (tempK > bestK)
                    {
                        bestK = tempK;
                        bestBeginIndexesArray = tempBeginIndexesArray;
                        cells = tempCells;
                    }
                }
            }

            if (correctRowInfoItem == null)
            {
                correctRowInfoItem = new()
                {
                    OriginRowInfoItem = rowInfoItem
                };
            }
            correctRowInfoItem.NewRowInfoItem = rowInfoItem;
            correctRowInfoItem.SimilarCoef = bestK;

            return correctRowInfoItem;
        }
    }
}

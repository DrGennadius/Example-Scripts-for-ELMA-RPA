using System.Linq;
using System.Text.RegularExpressions;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Уникальный элемент индексов начала столбцов.
    /// </summary>
    public class RowInfoItem
    {
        /// <summary>
        /// Индекс символа в тексте.
        /// </summary>
        public int TextBeginCharIndex { get; set; }

        /// <summary>
        /// Индексы начала столбцов.
        /// </summary>
        public int[] BeginColumnIndexes { get; set; }

        /// <summary>
        /// Продолжительные символы ячеек.
        /// </summary>
        public string[] СontinuousBodyRowCellTexts { get; set; }

        /// <summary>
        /// Вычисляемый паттерн строки тела таблицы.
        /// </summary>
        public string BodyRowPattern { get; set; }

        public RowInfoItem(RowInfoItem otherRowInfoItem)
        {
            TextBeginCharIndex = otherRowInfoItem.TextBeginCharIndex;
            BeginColumnIndexes = otherRowInfoItem.BeginColumnIndexes;
            СontinuousBodyRowCellTexts = otherRowInfoItem.СontinuousBodyRowCellTexts;
            BodyRowPattern = otherRowInfoItem.BodyRowPattern;
        }

        public RowInfoItem(int textBeginCharIndex, int[] beginColumnIndexes)
        {
            TextBeginCharIndex = textBeginCharIndex;
            BeginColumnIndexes = beginColumnIndexes;
            СontinuousBodyRowCellTexts = Enumerable.Repeat("", beginColumnIndexes.Length).ToArray();
            BodyRowPattern = "";
        }

        public RowInfoItem(int textBeginCharIndex, int[] beginColumnIndexes, string[] сontinuousBodyRowCellTexts) 
            : this(textBeginCharIndex, beginColumnIndexes)
        {
            СontinuousBodyRowCellTexts = сontinuousBodyRowCellTexts;
        }

        public RowInfoItem(int textBeginCharIndex, int[] beginColumnIndexes, string[] сontinuousBodyRowCellTexts, string bodyRowPattern) 
            : this(textBeginCharIndex, beginColumnIndexes, сontinuousBodyRowCellTexts)
        {
            BodyRowPattern = bodyRowPattern;
        }

        /// <summary>
        /// Генерировать паттерн строки.
        /// </summary>
        public void GenerateBodyRowPattern()
        {
            string pattern = "";

            foreach (var item in СontinuousBodyRowCellTexts)
            {
                string str = Regex.Replace(item, @"\s{2,}", @".+");
                str = Regex.Replace(str, @"\d", @"\d");
                pattern += str;
            }

            BodyRowPattern = pattern;
        }
    }
}

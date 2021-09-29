using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Параметры таблицы.
    /// </summary>
    public struct TableParameters
    {
        /// <summary>
        /// Индекс первого символа.
        /// </summary>
        public int FirstCharIndex;

        /// <summary>
        /// Индекс последнего символа.
        /// </summary>
        public int LastCharIndex;

        /// <summary>
        /// Индексы начала столбцов в каждой строке.
        /// </summary>
        public List<BeginColumnIndexesItem> BeginColumnIndexesItems;

        public TableParameters(int firstCharIndex, int lastCharIndex, List<BeginColumnIndexesItem> beginColumnIndexesItem)
        {
            FirstCharIndex = firstCharIndex;
            LastCharIndex = lastCharIndex;
            BeginColumnIndexesItems = beginColumnIndexesItem;
        }
    }

    /// <summary>
    /// Уникальный элемент индексов начала столбцов.
    /// </summary>
    public struct BeginColumnIndexesItem
    {
        /// <summary>
        /// Индекс символа в тексте.
        /// </summary>
        public int TextBeginCharIndex;

        /// <summary>
        /// Индексы начала столбцов.
        /// </summary>
        public int[] BeginColumnIndexes;

        public BeginColumnIndexesItem(int textBeginCharIndex, int[] beginColumnIndexes)
        {
            TextBeginCharIndex = textBeginCharIndex;
            BeginColumnIndexes = beginColumnIndexes;
        }
    }
}

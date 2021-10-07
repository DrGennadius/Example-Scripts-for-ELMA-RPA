using System;
using System.Collections.Generic;
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
        /// Список информации о строках.
        /// </summary>
        public List<RowInfoItem> RowInfoItems;

        public TableParameters(int firstCharIndex, int lastCharIndex, List<RowInfoItem> rowInfoItems)
        {
            FirstCharIndex = firstCharIndex;
            LastCharIndex = lastCharIndex;
            RowInfoItems = rowInfoItems;
        }
    }
}

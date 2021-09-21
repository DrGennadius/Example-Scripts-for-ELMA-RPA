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
        public int[] BeginColumnIndexes;

        public TableParameters(int firstCharIndex, int lastCharIndex, int[] beginColumnIndexes)
        {
            FirstCharIndex = firstCharIndex;
            LastCharIndex = lastCharIndex;
            BeginColumnIndexes = beginColumnIndexes;
        }
    }
}

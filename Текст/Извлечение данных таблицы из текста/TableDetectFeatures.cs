using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Признаки для определения таблицы.
    /// </summary>
    public class TableDetectFeatures
    {
        /// <summary>
        /// Текст первой ячейки (слева вверху). Стандартно '№'.
        /// </summary>
        public string FirstCellText { get; set; } = "№";

        /// <summary>
        /// Имеет последовательную нумерацию в первых ячейках. Стандартно true.
        /// </summary>
        public bool HasStartSequentialNumberingCells { get; set; } = true;

        /// <summary>
        /// Паттерн разделения столбцов. Стандартно /\s{2,}/.
        /// </summary>
        public string SplitPattern { get; set; } = @"\s{2,}";
    }
}

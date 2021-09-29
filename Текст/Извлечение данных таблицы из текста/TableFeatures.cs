using System.Collections.Generic;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Признаки для определения таблицы.
    /// </summary>
    public class TableFeatures
    {
        /// <summary>
        /// Паттерн первой строки первой ячейки таблицы (слева вверху). Стандартно '№'.
        /// </summary>
        public string FirstTableCellWordPattern { get; set; } = "№";

        /// <summary>
        /// Паттерн первой строки первой ячейки строки тела таблицы (слева). Стандартно /\d+/
        /// </summary>
        public string FirstBodyRowCellWordPattern { get; set; } = @"\d+";

        /// <summary>
        /// Имеет последовательную нумерацию в первых ячейках. Стандартно true.
        /// </summary>
        public bool HasStartSequentialNumberingCells { get; set; } = true;

        /// <summary>
        /// Паттерн разделения столбцов. Стандартно /\s{2,}/.
        /// </summary>
        public string SplitPattern { get; set; } = @"\s{2,}";

        /// <summary>
        /// Паттерн пропуска строк.
        /// Например, если есть фрагменты колонтитулов,
        /// которые сохранились в тексте, например,
        /// после извлечения текта с текстового слоя PDF.
        /// </summary>
        public string LineSkipPattern { get; set; }

        /// <summary>
        /// Список паттернов ячеек заголовка таблицы.
        /// </summary>
        public List<string> HeaderCellPatterns { get; set; } = new();
    }
}

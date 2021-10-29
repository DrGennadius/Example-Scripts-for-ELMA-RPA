using System;
using System.IO;
using System.Collections.Generic;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Контекст процесса
    /// <summary>
    public class Context
    {
        /// <summary>
        /// Путь к выходному файлу CSV для суммы
        /// </summary>
        public String OutputCSVFilePathSumm
        {
            get;
            set;
        }

        /// <summary>
        /// Путь к выходному файлу CSV для таблицы 1
        /// </summary>
        public String OutputCSVFilePathTable1
        {
            get;
            set;
        }

        /// <summary>
        /// Путь к выходному файлу CSV для таблицы 2
        /// </summary>
        public String OutputCSVFilePathTable2
        {
            get;
            set;
        }

        /// <summary>
        /// Итоговая сумма
        /// </summary>
        public Nullable<Double> Summ
        {
            get;
            set;
        }
    }
}
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
        /// Инфо о файле Excel
        /// </summary>
        public FileInfo ExcelFileInfo
        {
            get;
            set;
        }

        /// <summary>
        /// Сумма
        /// </summary>
        public Nullable<Double> Summ
        {
            get;
            set;
        }

        /// <summary>
        /// Корректный формат
        /// </summary>
        public Nullable<Boolean> IsCorrectFormat
        {
            get;
            set;
        }

        /// <summary>
        /// Есть ошибка в столбце "Итоговая стоимость"
        /// </summary>
        public Nullable<Boolean> HasSummError
        {
            get;
            set;
        }
    }
}
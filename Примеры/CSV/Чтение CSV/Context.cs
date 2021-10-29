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
        /// Путь к выходному файлу CSV для таблицы 1
        /// </summary>
        public String In_CSVFilePathTable1
        {
            get;
            set;
        }

        /// <summary>
        /// Путь к выходному файлу CSV для таблицы 2
        /// </summary>
        public String In_CSVFilePathTable2
        {
            get;
            set;
        }

        /// <summary>
        /// Строки из первого файла
        /// </summary>
        public List<String> Out_Rows1
        {
            get;
            set;
        }
    }
}
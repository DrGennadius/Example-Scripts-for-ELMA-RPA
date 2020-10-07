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
        /// Путь к файлу Excel
        /// </summary>
        public String ExcelFilePath
        {
            get;
            set;
        }

        /// <summary>
        /// Наименование листа
        /// </summary>
        public String SheetName
        {
            get;
            set;
        }
    }
}
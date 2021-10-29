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
        /// Путь файла настроек
        /// </summary>
        public String ConfigFilePath
        {
            get;
            set;
        }

        /// <summary>
        /// Путь файла настроек
        /// </summary>
        public String WorkFolder
        {
            get;
            set;
        }

        /// <summary>
        /// Имена файлов
        /// </summary>
        public List<String> FileNames
        {
            get;
            set;
        }
    }
}
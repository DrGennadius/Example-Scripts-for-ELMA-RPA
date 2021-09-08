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
        /// Путь к файлу конфигурации.
        /// </summary>
        public String ConfigFilePath
        {
            get;
            set;
        }

        /// <summary>
        /// Параметр 1 из листа 1.
        /// </summary>
        public String StringParam1s1
        {
            get;
            set;
        }

        /// <summary>
        /// Параметр 1 из листа 2.
        /// </summary>
        public String StringParam2s1
        {
            get;
            set;
        }

        /// <summary>
        /// Параметр 2 из листа 2.
        /// </summary>
        public String StringParam2s2
        {
            get;
            set;
        }
    }
}
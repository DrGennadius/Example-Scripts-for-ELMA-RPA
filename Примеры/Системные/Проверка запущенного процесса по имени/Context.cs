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
        /// Имя процесса
        /// </summary>
        public String ProcessName
        {
            get;
            set;
        }

        /// <summary>
        /// Процесс существует
        /// </summary>
        public Nullable<Boolean> ProcessExists
        {
            get;
            set;
        }
    }
}
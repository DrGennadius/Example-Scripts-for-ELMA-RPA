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
        // <value>Локальное имя хоста.</value>
        public String HostName
        {
            get;
            set;
        }

        // <value>Локальный IP хоста.</value>
        public String HostIP
        {
            get;
            set;
        }
    }
}
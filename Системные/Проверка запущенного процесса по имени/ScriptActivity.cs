using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Экземпляр данного класса будет создан при выполнении скрипта.
    /// <summary>
    public class ScriptActivity
    {
        /// <summary>
        /// Данная функция является точкой входа.
        /// <summary>
        public void Execute(Context context)
        {
            context.ProcessExists = Process.GetProcesses().Any(p => p.ProcessName == context.ProcessName);
            // Если нужно найти соответствие части названия
            //context.ProcessExists = Process.GetProcesses().Any(p => p.ProcessName.Contains(context.ProcessName));
        }
    }
}

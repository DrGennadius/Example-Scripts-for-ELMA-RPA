using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;

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
            ExcelConfigurationManager configurationManager = new ExcelConfigurationManager(context.ConfigFilePath);
            configurationManager.Read();
            context.StringParam1s1 = configurationManager.SingleParams["Лист1"]["Параметр 1"];
            context.StringParam2s1 = configurationManager.SingleParams["Лист2"]["Параметр 2.1"];
            context.StringParam2s2 = configurationManager.SingleParams["Лист2"]["Параметр 2.2"];
            context.StringListParam = configurationManager.MultipleParams["Лист3"]["Список"].ToList();
        }
    }
}

using System;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Используется только для целей debug в внешней IDE
    /// </summary>
    class Program
    {
        static void Main()
        {
            var context = new Context()
            {
                OutputCSVFilePathSumm = "Итоговая сумма.csv",
                OutputCSVFilePathTable1 = "Таблица 1.csv",
                OutputCSVFilePathTable2 = "Таблица 2.csv",
                Summ = 123456.44
            };
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);
            Console.ReadKey();
        }
    }
}

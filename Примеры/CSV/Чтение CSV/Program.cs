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
                In_CSVFilePathTable1 = "Таблица 1.csv",
                In_CSVFilePathTable2 = "Таблица 2.csv"
            };
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);
            foreach (var row in context.Out_Rows1)
            {
                Console.WriteLine(row);
                // При желании можно разбить по ячейкам.
            }
            Console.ReadKey();
        }
    }
}

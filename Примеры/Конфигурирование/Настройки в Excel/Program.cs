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
            Context context = new()
            {
                ConfigFilePath = "Пример.xlsx"
            };

            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);

            Console.WriteLine($"Лист 1 параметр 1 = '{context.StringParam1s1}'");
            Console.WriteLine($"Лист 2 параметр 1 = '{context.StringParam2s1}'");
            Console.WriteLine($"Лист 2 параметр 2 = '{context.StringParam2s2}'");
            Console.WriteLine($"Лист 3 множественный параметр = '{string.Join(", ", context.StringListParam)}'");

            Console.ReadKey();
        }
    }
}

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
                FilePath = "Пример текста.txt"
            };
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);

            Console.ReadKey();
        }
    }
}

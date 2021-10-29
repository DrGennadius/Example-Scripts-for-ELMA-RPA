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
            var context = new Context();
            context.ProcessName = "notepad";
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);
            Console.WriteLine($"{context.ProcessName} is running: {context.ProcessExists}");
            Console.ReadKey();
        }
    }
}

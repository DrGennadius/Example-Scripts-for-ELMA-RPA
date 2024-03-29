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
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);
            Console.WriteLine(context.HostName);
            Console.WriteLine(context.HostIP);
            Console.ReadKey();
        }
    }
}

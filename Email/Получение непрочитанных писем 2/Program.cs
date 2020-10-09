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
            var devConfig = new DevSettings("devconfig.json");
            var context = new Context()
            {
                WorkDirectory = "Вложения",
                Host = devConfig.Host,
                Port = devConfig.Port,
                UseSsl = devConfig.UseSsl,
                UserName = devConfig.UserName,
                Password = devConfig.Password,
                SubjectText = devConfig.SubjectText,
                SenderText = devConfig.SenderText
            };

            // Скрипт №1
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);

            // Скрипт №2
            var scriptActivity2 = new ScriptActivity2();
            scriptActivity2.Execute(context);

            Console.ReadKey();
        }
    }
}

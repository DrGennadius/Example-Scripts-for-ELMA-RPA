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
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);
            Console.ReadKey();
        }
    }
}

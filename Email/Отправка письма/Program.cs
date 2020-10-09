using System;
using System.Collections.Generic;
using System.IO;

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
                Host = devConfig.Host,
                Port = devConfig.Port,
                UseSsl = devConfig.UseSsl,
                UserName = devConfig.UserName,
                Password = devConfig.Password,
                Subject = devConfig.Subject,
                Sender = devConfig.Sender,
                Reciver = devConfig.Reciver,
                MessageText = devConfig.MessageText,
                Attachments = new List<FileInfo>()
                {
                    new FileInfo("Пример.xlsx"),
                    new FileInfo("Пример 2.txt")
                }
            };
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);
            Console.ReadKey();
        }
    }
}

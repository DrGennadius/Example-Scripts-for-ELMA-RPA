using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Экземпляр данного класса будет создан при выполнении скрипта.
    /// Это второй скрипт.
    /// <summary>
    public class ScriptActivity2
    {
        /// <summary>
        /// Данная функция является точкой входа.
        /// <summary>
        public void Execute(Context context)
        {
            StringBuilder sb = new StringBuilder();

            // Получаем список контекстов писем
            var messageContexts = context.MessageContextStrings.Select(x => new MessageContext(x));
            foreach (var item in messageContexts)
            {
                // Какие-то действия с телом сообщения и вложениями
                sb.Append(string.Format("Сообщение:\n{0}\nВложения:\n{1}", item.TextBody, string.Join(Environment.NewLine, item.AttachmentPaths)));
            }

            // Продолжаем выполнять какие-то действия
            byte[] bytes = Encoding.Default.GetBytes(sb.ToString());
            string outputStr = Encoding.UTF8.GetString(bytes);

            string outputPath = Path.Combine(context.WorkDirectory, string.Format("{0:yyyy_MM_dd_HH_mm_ss}.txt", DateTime.Now));
            File.WriteAllText(outputPath, outputStr, Encoding.UTF8);
        }
    }
}

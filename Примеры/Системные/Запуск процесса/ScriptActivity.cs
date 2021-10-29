using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Экземпляр данного класса будет создан при выполнении скрипта.
    /// <summary>
    public class ScriptActivity
    {
        /// <summary>
        /// Данная функция является точкой входа.
        /// <summary>
        public void Execute(Context context)
        {
            // Делаем первичные проверочки на существование необходимых файлов.
            if (!File.Exists(context.PdfFilePath))
            {
                context.CommandExecuteResult = "Исполняемого файла утилиты pdftotext не существует.";
            }
            if (!File.Exists(context.PdfFilePath))
            {
                context.CommandExecuteResult = "Файла PDF не существует.";
            }
            // ...
            // Там еще какие-нибудь проверочки, если нужно...
            // ...

            // Создаем строку с аргументам.
            string arguments = $"-enc UTF-8 -table \"{context.PdfFilePath}\" \"{context.ExtractedTxtFilePath}\"";
            // Напоминаем, что строку можно собрать разными способами:
            //string arguments = string.Format("-enc UTF-8 -table \"{0}\" \"{1}\"", context.PdfFilePath, context.ExtractedTxtFilePath);
            //string arguments = "-enc UTF-8 -table \"" + context.PdfFilePath + "\" \"" + context.ExtractedTxtFilePath + "\"";

            // Создаем экземпляр информации запускаемого процесса.
            ProcessStartInfo procStartInfo = new()
            {
                // Имя исполняемого файла
                FileName = context.PdftotextPath,
                // Аргументы для испольняемого файла.
                // Тут приведен пример использования кодировки UTF8 на случай,
                // если есть проблемы в системе использования, например, кириллицы.
                // В принципе можно в любом случае использовать.
                Arguments = Encoding.Default.GetString(Encoding.UTF8.GetBytes(arguments)),
                // Перенаправление стандартного вывода.
                // В общем это когда в консоле что-то пишет какая-то команда,
                // то вот с помощью такой настройки мы можем это получить в коде.
                RedirectStandardOutput = true
            };

            // Создаем процесс, запускаем и ждем.
            using (Process process = new())
            {
                process.StartInfo = procStartInfo;
                // Запускаем процесс.
                bool isStarted = process.Start();
                if (isStarted)
                {
                    // Если запущен.
                    // Ожидаем завершения процесса (команды).
                    process.WaitForExit();

                    // Сохраняем вывод в контексте робота.
                    context.CommandExecuteResult = process.StandardOutput.ReadToEnd();
                }
                else
                {
                    // Если не запущен, то записываем это в контекст.
                    context.CommandExecuteResult = "Процесс не запущен.";
                }
            }
        }
    }
}

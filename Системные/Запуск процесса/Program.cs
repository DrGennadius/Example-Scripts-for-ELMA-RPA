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
                PdftotextPath = @"pdftotext.exe",
                PdfFilePath = @"Документ на кириллице.pdf",
                ExtractedTxtFilePath = @"Извлеченный текст\Документ на кириллице.txt"
            };
            var scriptActivity = new ScriptActivity();
            scriptActivity.Execute(context);
            Console.WriteLine(context.CommandExecuteResult);
            Console.ReadKey();
        }
    }
}

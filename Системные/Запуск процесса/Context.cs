using System;
using System.IO;
using System.Collections.Generic;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Контекст процесса
    /// <summary>
    public class Context
    {
        // <value>Путь к файлу PDF</value>
        public String PdfFilePath
        {
            get;
            set;
        }

        // <value>Путь к pdftotext.exe</value>
        public String PdftotextPath
        {
            get;
            set;
        }

        // <value>Наименование файла с извлеченным текстом</value>
        public String ExtractedTxtFilePath
        {
            get;
            set;
        }

        // <value>Результат выполнения комманды</value>
        public String CommandExecuteResult
        {
            get;
            set;
        }
    }
}
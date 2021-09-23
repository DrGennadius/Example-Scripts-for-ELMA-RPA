using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Data;

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
            if (
                !string.IsNullOrWhiteSpace(context.FilePath) 
                && File.Exists(context.FilePath)
                )
            {

                // Считываем весь текст.
                string text = File.ReadAllText(context.FilePath);
                
                // Конфигурируем признаки, по которым будем искать таблицу.
                TableDetectFeatures tableDetectFeatures = new()
                {
                    HasStartSequentialNumberingCells = true
                };

                TableExtractor tableExtractor = new(tableDetectFeatures);

                // Извлекаем таблицу из текста (данные таблицы храняться в экземпляре tableExtractor).
                bool isSucces = tableExtractor.Extract(text);

                // Далее можно обращаться к данным, если успешно
                if (isSucces)
                {
                    string data1 = tableExtractor.Data[0, 0];
                    // ...
                    // Еще какой-нибудь код...
                    // ...
                }
            }
        }
    }
}

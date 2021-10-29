using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using ExcelDataReader;
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
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            // Почти также как и через FileInfo
            using (var stream = File.Open(context.ExcelFilePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                reader = ExcelReaderFactory.CreateReader(stream);
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        // Использовать в таблицах заголовки
                        UseHeaderRow = true
                    }
                };

                // Получение DataSet (набор данных)
                // Подробнее https://docs.microsoft.com/ru-ru/dotnet/api/system.data.dataset?view=netcore-3.1
                DataSet dataSet = reader.AsDataSet(conf);

                // Некоторые примеры использования:

                // Получение DataTable (определенная таблица, Sheet)
                // Подробнее https://docs.microsoft.com/ru-ru/dotnet/api/system.data.datatable?view=netcore-3.1
                DataTable dataTable = dataSet.Tables[0];

                // Получение DataRowCollection (набор строк)
                DataRowCollection rows = dataTable.Rows;

                // Получение DataColumnCollection (набор столбцов)
                DataColumnCollection columns = dataTable.Columns;

                // Проверка существует ли столбец "Итоговая стоимость"
                if (columns.Contains("Итоговая стоимость"))
                {
                    context.IsCorrectFormat = true;

                    // В качестве примера посчитаем сумму, но с учетом того, что последняя строка итоговая.
                    int rowsCount = rows.Count;
                    double summ = 0.0;
                    for (int i = 0; i < rowsCount - 1; i++)
                    {
                        summ += (double)rows[i]["Итоговая стоимость"];
                    }
                    context.Summ = summ;

                    // Мы можем также проверить еще что-нибудь, например, расчитанную и итоговую.
                    // И в случае ошибки отметить это в контексте.
                    if (Math.Round(summ, 2) != Math.Round((double)rows[rowsCount - 1]["Итоговая стоимость"], 2))
                    {
                        context.HasSummError = true;
                    }
                }
                else
                {
                    // В качестве примера обозначим, что если не нашли столбец "Итоговая стоимость",
                    // то считаем некорректный формат. Далее это можно использовать роботом,
                    // например, в шлюзе и выполнить какие-то определенные действия для этого случая.
                    context.IsCorrectFormat = false;
                }

                // Подробности по объектам и по другим случаям чтения можно найти в подразделах:
                // https://docs.microsoft.com/ru-ru/dotnet/api/system.data?view=netcore-3.1
            }
        }
    }
}

using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

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
            // Разделитель (в прогамме должен быть настроен этот разделитель,
            // а также в файле используется именно этот разделитель).
            // Частенько используется ',', ';' или '\t'(Tab).
            char separateChar = ';';

            // Пример 1. Прочитаем файл CSV и сохраним строки в контекст.
            // В дальнейшем эти строки можно использовать в роботе далее.
            // Например, создать цикл "Повторять для каждого" на графической схеме и использовать этот список.
            context.Out_Rows1 = GetRowsFromCSV(context.In_CSVFilePathTable1);

            // Пример 2. Прочитаем второй файл CSV и обработаем его, например, создадим DataTable
            DataTable dataTable = GetDataTableFromCSV(context.In_CSVFilePathTable2, separateChar);
        }

        /// <summary>
        /// Получение списка строк из файла CSV
        /// </summary>
        /// <param name="filePath">Путь к файлу CSV</param>
        /// <returns></returns>
        private List<string> GetRowsFromCSV(string filePath)
        {
            List<string> output = new List<string>();

            using (StreamReader reader = new StreamReader(File.OpenRead(filePath)))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    output.Add(line);
                }
            }

            return output;
        }

        /// <summary>
        /// Получение DataTable из файла CSV
        /// </summary>
        /// <param name="filePath">Путь к файлу CSV</param>
        /// <param name="separateChar">Символ разделителя</param>
        /// <returns></returns>
        private DataTable GetDataTableFromCSV(string filePath, char separateChar)
        {
            DataTable dataTable = new DataTable("Наименование таблицы");

            using (StreamReader reader = new StreamReader(File.OpenRead(filePath)))
            {
                // Для того, чтобы сохранить 1ую строку как заголовки и создать столбцы
                bool isFirstLine = true;
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        string[] values = line.Split(separateChar);
                        if (isFirstLine)
                        {
                            for (int i = 0; i < values.Length; i++)
                            {
                                if (i == 0)
                                {
                                    // Для первого создаем столбец с типом int
                                    dataTable.Columns.Add(
                                        new DataColumn()
                                        {
                                            ColumnName = values[i],
                                            DataType = typeof(int)
                                        }
                                    );
                                }
                                else
                                {
                                    // Для остальных создаем с типом string (стандартное)
                                    dataTable.Columns.Add(new DataColumn(values[i]));
                                }
                                // Можно было бы еще обработать даты
                            }
                            isFirstLine = false;
                        }
                        else
                        {
                            DataRow row = dataTable.NewRow();
                            // Спсиок объектов для удобства хранения значений разных типов
                            List<object> cellValues = new List<object>();
                            for (int i = 0; i < values.Length; i++)
                            {
                                if (i == 0)
                                {
                                    // Для первого создаем с типом int
                                    cellValues.Add(Convert.ToInt32(values[i]));
                                }
                                else
                                {
                                    // Для остальных создаем с типом string (как есть уже)
                                    cellValues.Add(values[i]);
                                }
                                // Можно было бы еще обработать даты
                            }
                            // Добавляем в строку
                            row.ItemArray = cellValues.ToArray();
                            // Добавляем строку в список
                            dataTable.Rows.Add(row);
                        }
                    }
                }
            }

            return dataTable;
        }
    }
}

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
            // Разделитель (в прогамме должен быть настроен этот разделитель).
            // Частенько используется ',', ';' или '\t'(Tab).
            string ss = ";";

            // Пример 1. Запись в одну ячейку
            if (!context.Summ.HasValue)
            {
                context.Summ = 0;
            }
            File.WriteAllText(context.OutputCSVFilePathSumm, context.Summ.Value.ToString("0.00"));

            // Пример 2. Создаем простую таблицу сразу в строку, а точнее будем использовать StringBuilder
            StringBuilder sb1 = new StringBuilder();
            sb1.AppendLine("Имя" + ss + "Возраст");
            sb1.AppendLine("Александра" + ss + "21");
            sb1.AppendLine("Петр" + ss + "33");

            // Далее можно закодировать текст, например UTF8
            byte[] bytes1 = Encoding.Default.GetBytes(sb1.ToString());
            string outputStr1 = Encoding.UTF8.GetString(bytes1);

            // И записать в файл уже с указанной кодировкой
            File.WriteAllText(context.OutputCSVFilePathTable1, outputStr1, Encoding.UTF8);

            // Пример 3. Запись DataTable в CSV
            DataTable dataTable = GetTable();

            StringBuilder sb2 = new StringBuilder();

            IEnumerable<string> columnNames = dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
            sb2.AppendLine(string.Join(ss, columnNames));

            foreach (DataRow row in dataTable.Rows)
            {
                // Если нужно просто создать массив из строк без доп. обработки, просто конвертирование в строку:
                //IEnumerable<string> cells = row.ItemArray.Select(x => x.ToString());

                // Чтобы форматировать определенные типы
                IEnumerable<string> cells = row.ItemArray.Select(x => x is DateTime ? ((DateTime)x).ToShortDateString() : x.ToString());
                sb2.AppendLine(string.Join(ss, cells));
            }

            byte[] bytes2 = Encoding.Default.GetBytes(sb2.ToString());
            string outputStr2 = Encoding.UTF8.GetString(bytes2);

            File.WriteAllText(context.OutputCSVFilePathTable2, outputStr2, Encoding.UTF8);

            // На выходе получили 3 файла по 3м примерам.
        }

        /// <summary>
        /// Пример простой таблицы
        /// </summary>
        /// <returns></returns>
        private DataTable GetTable()
        {
            int id = 101;
            DataTable table = new DataTable()
            {
                // Наименование таблицы:
                TableName = "Тестовая таблица",
                // Столбцы:
                Columns =
                {
                    { "Id", typeof(int) },
                    { "Фамилия", typeof(string) },
                    { "Имя", typeof(string) },
                    { "Отчество", typeof(string) },
                    { "Дата начала работы", typeof(DateTime) }
                },
                // Строки:
                Rows =
                {
                    { ++id, "Иванов", "Иван", "Иванович", new DateTime(2020, 7, 7) },
                    { ++id, "Абушев", "Олег", "Иванович", new DateTime(2020, 7, 10) },
                    { ++id, "Иванов", "Виталий", "Тестирович", new DateTime(2020, 8, 15) },
                    { ++id, "Программист", "Кирилл", "Дотнетович", new DateTime(2020, 9, 1) },
                }
            };
            return table;
        }
    }
}

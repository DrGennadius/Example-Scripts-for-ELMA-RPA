using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
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
            using (var spreadsheetDocument = SpreadsheetDocument.Create(context.ExcelFilePath, SpreadsheetDocumentType.Workbook))
            {
                // Получаем главную часть книги
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();

                // Создаем обязательно объекты Workbook и Sheets
                spreadsheetDocument.WorkbookPart.Workbook = new Workbook();
                spreadsheetDocument.WorkbookPart.Workbook.Sheets = new Sheets();

                // Получаем таблицу
                DataTable table = GetTable();

                // Добавляем в нашу часть книги новую часть - рабоичий лист, куда будем записывать данные из таблицы типа DataTable
                var sheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                // Создаем SheetData, где хранятся все значения листа
                var sheetData = new SheetData();
                // Создаем объекты типа Worksheet и Sheets
                sheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                // Сохраняем id части листа
                string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(sheetPart);

                // Далее создаем объект типа Sheet и добавлем его в список sheets,
                // но нужно задать id. Для этого ставим 1 или найдем максимальный из сщуествующих и прибавим 1
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }
                Sheet sheet = new Sheet() 
                {
                    Id = relationshipId,
                    SheetId = sheetId,
                    // В данном примере имя листа берем по имени таблицы
                    Name = table.TableName
                };
                sheets.Append(sheet);

                // Создаем заголовочную строку
                Row headerRow = new Row();

                List<string> columns = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.CellValue = new CellValue(column.ColumnName);
                    cell.DataType = CellValues.String;
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                // Далее создаем остальные строки
                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (string col in columns)
                    {
                        Cell cell = new Cell();
                        // Для значний типа DateTime будем формировать короткий формат даты, остальное - просто получаем как строку
                        cell.CellValue = new CellValue(dsrow[col] is DateTime ? ((DateTime)dsrow[col]).ToShortDateString() : dsrow[col].ToString());
                        // В примере используем везде тип Строка, при желании можно доп. определять тип
                        cell.DataType = CellValues.String;
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }
            }
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

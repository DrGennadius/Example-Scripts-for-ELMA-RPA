using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

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
            SpreadsheetDocument spreadsheetDocument = null;
            // Проверка, существует ли файл Excel по пути, указанному в контекстной переменной.
            // На самом деле не обязательно. Если нужно всегда создавать новый (с перезаписью),
            // то можно всегда использовать SpreadsheetDocument.Create()
            // Но в качестве примера рассмотрим ситуацию с изменением уже созданного файла Excel
            bool excelExists = File.Exists(context.ExcelFilePath);
            if (excelExists)
            {
                // При существующем файле открываем
                spreadsheetDocument = SpreadsheetDocument.Open(context.ExcelFilePath, true);
            }
            else
            {
                // Если файл не существует, то создаем
                spreadsheetDocument = SpreadsheetDocument.Create(context.ExcelFilePath, SpreadsheetDocumentType.Workbook);
            }

            // Workbook - книга (по другому можно сказать), является элементом верхнего уровня,
            // которая содержит несколько таблиц (или другие встречаемые термины: страницы, листы)) и другие атрибуты.
            // Первым делом нужно получить его или создать новый.
            WorkbookPart workbookPart = excelExists ? spreadsheetDocument.WorkbookPart : spreadsheetDocument.AddWorkbookPart();
            if (!excelExists || workbookPart == null)
            {
                workbookPart.Workbook = new Workbook();
            }

            Sheet sheet = null;

            // Worksheet - рабочий лист. Необходимо найти по имени, а не первый,
            // который по сути будет последним активным при сохранении Excel.
            WorksheetPart worksheetPart = null;
            if (excelExists)
            {
                sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == context.SheetName);
                if (sheet == null)
                {
                    worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                }
                else
                {
                    worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                }
            }
            if (worksheetPart == null)
            {
                worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            }

            // Проверяем таблицу, которая искалась ранее
            if (sheet == null)
            {
                // Если таблица не найдена, то создаем новый
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                uint newSheetId = 1U;
                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                if (sheets != null)
                {
                    // Нужно задать верный id для нового листа, после последнего
                    Sheet lastSheet = sheets.Elements<Sheet>().OrderBy(x => x.SheetId).LastOrDefault();
                    if (lastSheet != null)
                    {
                        newSheetId = lastSheet.SheetId + 1;
                    }
                }
                else
                {
                    // Если элемента с таким типом еще нет, то обязательно нужно создать:
                    sheets = workbookPart.Workbook.AppendChild(new Sheets());
                }

                sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = newSheetId,
                    Name = context.SheetName
                };
                sheets.Append(sheet);

                // Далее можно настроить столбцы, но необязательно
                // Если не настраивать, то будет стандартный вид.
                Columns columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                bool needToInsertColumns = false;
                if (columns == null)
                {
                    columns = new Columns();
                    needToInsertColumns = true;
                }
                // Для примера создадим столбец с настроенной длиной
                Column column1 = new Column() { Min = 1U, Max = 1U, Width = 20D, CustomWidth = true };
                columns.Append(column1);
                if (needToInsertColumns)
                {
                    worksheetPart.Worksheet.InsertAt(columns, 0);
                }
            }

            // Страницу можно также получить по id, например, первую так:
            // sheet = (Sheet)workbookPart.Workbook.Sheets.ChildElements.GetItem(0);

            // Получим данные листа
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Далее нужно по необходимости. Если предпологается, что всегда будет создаваться новый,
            // то можно просто установить rowIndex = 1U (именно 1, не 0).
            // Если всегда предполагается добавление новых строк с существующим,
            // то нужно установить значение rowIndex, относительно количества существующих строк:
            var rows = sheetData.Elements<Row>();
            uint rowIndex = excelExists ? (uint)rows.Count() + 1U : 1U;

            // Создание строки с индексом
            Row row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);

            // Пример новой ячейки с текстом:
            Cell newCell1 = InsertCellInWorksheet("A", rowIndex, worksheetPart);

            newCell1.CellValue = new CellValue("Какой-то текст");
            newCell1.DataType = new EnumValue<CellValues>(CellValues.String);

            // Инкрементируем индекс
            rowIndex++;

            // Пример новый ячейки с числом
            Cell newCell2 = InsertCellInWorksheet("A", rowIndex, worksheetPart);

            newCell2.CellValue = new CellValue("12345");
            newCell2.DataType = new EnumValue<CellValues>(CellValues.Number);

            // Не забываем закрывать, если не использовали блок using:
            spreadsheetDocument.Close();
        }

        /// <summary>
        /// Метод для получения или создания ячейки и вставки данных в эту ячейку
        /// </summary>
        /// <param name="columnName">Имя столбца в формате как в Excel [A-Z]</param>
        /// <param name="rowIndex">Индекс строки</param>
        /// <param name="worksheetPart"></param>
        /// <returns>Ячейка</returns>
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // Если рабочий лист не содержит строки с указанным индексом строки, то просто добавляем ее.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // Если ячейки с указанным именем столбца нет, то добавляем.
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Ячейки должны располагаться в последовательном порядке согласно CellReference.
                // Необходимо определить, куда вставить новую ячейку.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
    }
}

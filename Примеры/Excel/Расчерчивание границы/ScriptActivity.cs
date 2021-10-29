using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

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
            GetData(context, out SpreadsheetDocument spreadsheetDocument, out WorksheetPart worksheetPart);

            // 1. Получаем/создаем объект типа WorkbookStylesPart
            var stylesPart = InitWorkbookStylesPart(spreadsheetDocument);

            // 2. Создаем свою комбинацию стилей для рамок и стили форматов ячеек для них,
            // хотя можно сделать по другому - создавать недостающие находу (может быть чуть позже в другом примере будет).
            AddAllOwnBorderStyles(stylesPart, BorderStyleValues.Medium);

            // 3. Пример новой ячейки с границами
            Cell newCell0 = InsertCellInWorksheet("B", 2, worksheetPart);
            newCell0.CellValue = new CellValue("АБВ");
            newCell0.DataType = new EnumValue<CellValues>(CellValues.String);
            List<BorderPropertiesType> borderPropertiesTypes0 = new List<BorderPropertiesType>()
            {
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder()
            };
            newCell0.StyleIndex = GetCellFormatIndex(stylesPart, borderPropertiesTypes0);

            // 4. Пример создания 4 ячейки внутри общей границы
            List<BorderPropertiesType> borderPropertiesTypes1 = new List<BorderPropertiesType>()
            {
                new LeftBorder(),
                new TopBorder()
            };
            List<BorderPropertiesType> borderPropertiesTypes2 = new List<BorderPropertiesType>()
            {
                new TopBorder(),
                new RightBorder()
            };
            List<BorderPropertiesType> borderPropertiesTypes3 = new List<BorderPropertiesType>()
            {
                new LeftBorder(),
                new BottomBorder()
            };
            List<BorderPropertiesType> borderPropertiesTypes4 = new List<BorderPropertiesType>()
            {
                new RightBorder(),
                new BottomBorder()
            };

            Cell newCell1 = InsertCellInWorksheet("C", 3, worksheetPart, "C3");
            newCell1.StyleIndex = GetCellFormatIndex(stylesPart, borderPropertiesTypes1);

            Cell newCell2 = InsertCellInWorksheet("D", 3, worksheetPart, "D3");
            newCell2.StyleIndex = GetCellFormatIndex(stylesPart, borderPropertiesTypes2);

            Cell newCell3 = InsertCellInWorksheet("C", 4, worksheetPart, "C4");
            newCell3.StyleIndex = GetCellFormatIndex(stylesPart, borderPropertiesTypes3);

            Cell newCell4 = InsertCellInWorksheet("D", 4, worksheetPart, "D4");
            newCell4.StyleIndex = GetCellFormatIndex(stylesPart, borderPropertiesTypes4);

            // Не забываем закрывать, если не использовали блок using:
            spreadsheetDocument.Close();
        }

        private void AddAllOwnBorderStyles(WorkbookStylesPart stylesPart, BorderStyleValues borderStyle = BorderStyleValues.Thin)
        {
            var propertiesTypesCombinations = GetAllCombinations(
                new List<BorderPropertiesType> {
                    new LeftBorder() { Style = borderStyle },
                    new RightBorder() { Style = borderStyle },
                    new TopBorder() { Style = borderStyle },
                    new BottomBorder() { Style = borderStyle }
                }
            );
            foreach (var item in propertiesTypesCombinations)
            {
                var borderPropertiesTypeClones = item.Select(x => (BorderPropertiesType)x.Clone());
                AddOwnBorderStyle(stylesPart, borderPropertiesTypeClones, borderStyle);
                UpdateCellFormat(stylesPart, borderPropertiesTypeClones);
            }
        }

        private List<List<T>> GetAllCombinations<T>(List<T> list)
        {
            int comboCount = (int)Math.Pow(2, list.Count) - 1;
            List<List<T>> result = new List<List<T>>();
            for (int i = 1; i < comboCount + 1; i++)
            {
                result.Add(new List<T>());
                for (int j = 0; j < list.Count; j++)
                {
                    if ((i >> j) % 2 != 0)
                        result.Last().Add(list[j]);
                }
            }
            return result;
        }

        private void AddOwnBorderStyle(WorkbookStylesPart stylesPart, IEnumerable<BorderPropertiesType> borderPropertiesTypes, BorderStyleValues borderStyle = BorderStyleValues.Thin)
        {
            var borderElements = stylesPart.Stylesheet.Borders.ChildElements.Where(x => x.ChildElements.Count == borderPropertiesTypes.Count());
            foreach (var borderElement in borderElements)
            {
                bool existsAll = true;
                foreach (var item in borderPropertiesTypes)
                {
                    if (!borderElement.ChildElements.Where(x => x.GetType() == item.GetType()).Any())
                    {
                        existsAll = false;
                    }
                }
                if (existsAll)
                {
                    return;
                }
            }
            Border border = new Border();
            foreach (var item in borderPropertiesTypes)
            {
                border.Append(item);
            }
            stylesPart.Stylesheet.Borders.Append(border);
        }

        private UInt32Value GetBorderIndex(WorkbookStylesPart stylesPart, IEnumerable<BorderPropertiesType> borderPropertiesTypes)
        {
            UInt32Value index = 0U;
            var borderElements = stylesPart.Stylesheet.Borders.ChildElements;
            for (int i = 0; i < borderElements.Count; i++)
            {
                bool existsAll = true;
                foreach (var item in borderPropertiesTypes)
                {
                    if (!borderElements.ElementAt(i).ChildElements.Where(x => x.GetType() == item.GetType()).Any())
                    {
                        existsAll = false;
                    }
                }
                if (existsAll)
                {
                    index = Convert.ToUInt32(i);
                    break;
                }
            }
            return index;
        }

        private UInt32Value GetCellFormatIndex(WorkbookStylesPart stylesPart, UInt32Value borderId)
        {
            UInt32Value index = 0U;
            for (int i = 0; i < stylesPart.Stylesheet.CellFormats.ChildElements.Count; i++)
            {
                if ((((CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt(i))).BorderId == borderId)
                {
                    index = Convert.ToUInt32(i);
                    break;
                }
            }
            return index;
        }

        private UInt32Value GetCellFormatIndex(WorkbookStylesPart stylesPart, IEnumerable<BorderPropertiesType> borderPropertiesTypes)
        {
            return GetCellFormatIndex(stylesPart, GetBorderIndex(stylesPart, borderPropertiesTypes));
        }

        private void UpdateCellFormat(WorkbookStylesPart stylesPart, IEnumerable<BorderPropertiesType> borderPropertiesTypes)
        {
            UInt32Value borderId = GetBorderIndex(stylesPart, borderPropertiesTypes);
            bool exists = false;
            for (int i = 0; i < stylesPart.Stylesheet.CellFormats.ChildElements.Count; i++)
            {
                if ((((CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt(i))).BorderId == borderId)
                {
                    exists = true;
                    break;
                }
            }
            if (!exists)
            {
                CellFormat cellformat = new CellFormat() { FontId = 0, FillId = 0, BorderId = borderId };
                stylesPart.Stylesheet.CellFormats.Append(cellformat);
            }
        }

        /// <summary>
        /// Инициализация части стилей. При создании создаются стандартные элементы.
        /// </summary>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="workbookStylesPart"></param>
        private WorkbookStylesPart InitWorkbookStylesPart(SpreadsheetDocument spreadsheetDocument)
        {
            WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.WorkbookStylesPart;

            // 1. Создаем отсутствующие элементы.
            if (workbookStylesPart == null)
            {
                workbookStylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            }
            if (workbookStylesPart.Stylesheet == null)
            {
                workbookStylesPart.Stylesheet = new Stylesheet();
            }
            if (workbookStylesPart.Stylesheet.Fonts == null)
            {
                workbookStylesPart.Stylesheet.Fonts = new Fonts();
            }
            if (workbookStylesPart.Stylesheet.CellFormats == null)
            {
                workbookStylesPart.Stylesheet.CellFormats = new CellFormats();
            }
            if (workbookStylesPart.Stylesheet.Fills == null)
            {
                workbookStylesPart.Stylesheet.Fills = new Fills();
            }
            if (workbookStylesPart.Stylesheet.Borders == null)
            {
                workbookStylesPart.Stylesheet.Borders = new Borders();
            }

            // 2. Проверяем шрифты. Предпологаем, что если нет, то создаем.
            if (workbookStylesPart.Stylesheet.Fonts.ChildElements.Count == 0)
            {
                // Создаем некий "стандартный" шрифт, который пустой, и добавляем его в список
                Font defaultFont = new Font();
                workbookStylesPart.Stylesheet.Fonts.Append(defaultFont);
            }

            // 3. Проверяем заполнения. Предпологаем, что если нет, то создаем хотя бы 1 "стандартный".
            if (workbookStylesPart.Stylesheet.Fills.ChildElements.Count == 0)
            {
                Fill defaultFill = new Fill();
                workbookStylesPart.Stylesheet.Fills.Append(defaultFill);
            }

            // 4. Проверяем рамки. Предпологаем, что если нет, то создаем хотя бы 1 "стандартные".
            if (workbookStylesPart.Stylesheet.Borders.ChildElements.Count == 0)
            {
                Border defaultBorder = new Border();
                workbookStylesPart.Stylesheet.Borders.Append(defaultBorder);
            }

            // 5. Проверяем форматы ячейки. Предпологаем, что если нет, то создаем.
            if (workbookStylesPart.Stylesheet.CellFormats.ChildElements.Count == 0)
            {
                // Создаем "стандартный" формат ячейки и добавляем его в список
                CellFormat defaultCellformat = new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 };
                workbookStylesPart.Stylesheet.CellFormats.Append(defaultCellformat);
            }

            return workbookStylesPart;
        }

        /// <summary>
        /// Получаем данные: SpreadsheetDocument, worksheetPart и sheetData
        /// </summary>
        /// <param name="context"></param>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="worksheetPart"></param>
        /// <param name="sheetData"></param>
        private void GetData(Context context, out SpreadsheetDocument spreadsheetDocument, out WorksheetPart worksheetPart)
        {
            // В данном методе открывается/создается файл excel и получаем/создаем элементы структуры.
            // C комментариями можно разобрать тут:
            // https://github.com/DrGennadius/Example-Scripts-for-ELMA-RPA/blob/master/Excel/%D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5%20%D0%BD%D0%BE%D0%B2%D0%BE%D0%B3%D0%BE%20%D0%BB%D0%B8%D1%81%D1%82%D0%B0/ScriptActivity.cs
            spreadsheetDocument = null;
            bool excelExists = File.Exists(context.ExcelFilePath);
            if (excelExists)
            {
                spreadsheetDocument = SpreadsheetDocument.Open(context.ExcelFilePath, true);
            }
            else
            {
                spreadsheetDocument = SpreadsheetDocument.Create(context.ExcelFilePath, SpreadsheetDocumentType.Workbook);
            }
            WorkbookPart workbookPart = excelExists ? spreadsheetDocument.WorkbookPart : spreadsheetDocument.AddWorkbookPart();
            if (!excelExists || workbookPart == null)
            {
                workbookPart.Workbook = new Workbook();
            }
            Sheet sheet = null;
            worksheetPart = null;
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
            if (sheet == null)
            {
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                uint newSheetId = 1U;
                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                if (sheets != null)
                {
                    Sheet lastSheet = sheets.Elements<Sheet>().OrderBy(x => x.SheetId).LastOrDefault();
                    if (lastSheet != null)
                    {
                        newSheetId = lastSheet.SheetId + 1;
                    }
                }
                else
                {
                    sheets = workbookPart.Workbook.AppendChild(new Sheets());
                }

                sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = newSheetId,
                    Name = context.SheetName
                };
                sheets.Append(sheet);
            }
        }

        /// <summary>
        /// Вставить данные в ячейку листа
        /// </summary>
        /// <param name="columnName">Имя столбца в формате как в Excel [A-Z...]</param>
        /// <param name="rowIndex">Индекс строки</param>
        /// <param name="worksheetPart"></param>
        /// <param name="value">Значение</param>
        /// <param name="type">Тип значения</param>
        /// <returns></returns>
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart, string value, CellValues type = CellValues.String)
        {
            Cell newCell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);
            newCell.CellValue = new CellValue(value);
            newCell.DataType = new EnumValue<CellValues>(type);
            return newCell;
        }

        /// <summary>
        /// Метод для получения или создания ячейки и вставки данных в эту ячейку
        /// </summary>
        /// <param name="columnName">Имя столбца в формате как в Excel [A-Z...]</param>
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
                // Обязательно необходимо соблюсти последовательность индексов
                Row nextRow = sheetData.Elements<Row>().OrderBy(r => r.RowIndex).FirstOrDefault(r => r.RowIndex > rowIndex);
                row = new Row() { RowIndex = rowIndex };
                sheetData.InsertBefore(row, nextRow);
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

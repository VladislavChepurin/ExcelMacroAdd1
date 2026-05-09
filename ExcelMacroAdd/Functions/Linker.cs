using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;


namespace ExcelMacroAdd.Functions
{
    internal sealed class Linker : AbstractFunctions
    {
        private readonly ICorrectFontResources correctFontResources;

        public Linker(ICorrectFontResources correctFontResources)
        {
            this.correctFontResources = correctFontResources;
        }

        public override void Start()
        {
            Workbook workbook = null;
            Worksheet worksheet = null;
            Sheets sheets = null;
            Range headersRange = null;
            Range columnWidthRange = null;
            Range formattingRange = null;
            Range formattingCells = null;
            Font formattingFont = null;
            Range fontRange = null;
            Font fontRangeFont = null;
            Range bordersRange = null;
            Borders borders = null;
            Range focusRange = null;

            try
            {
                workbook = WorkBook;
                worksheet = Worksheet;

                //Создаем коллекцию
                List<int> list = new List<int>();

                // В коллекцию помещаем парсированые имена листов
                sheets = workbook.Sheets;
                int sheetCount = sheets.Count;
                for (int index = 1; index <= sheetCount; index++)
                {
                    Worksheet sheet = null;
                    try
                    {
                        sheet = (Worksheet)sheets[index];
                        int.TryParse(sheet.Name, out int result);
                        list.Add(result);
                    }
                    finally
                    {
                        ReleaseComObject(sheet);
                    }
                }

                int worksheetIndex = worksheet.Index;
                //Проверяем на налиечие этих номеров в коллекции, если есть, то имя вкладке не присваиваем
                if (!list.Contains(worksheetIndex - 1))
                {
                    //имя вкладки в зависимости от индекса
                    worksheet.Name = (worksheetIndex - 1).ToString();
                }

                // Заголовки столбцов
                headersRange = worksheet.Range["A1", "J1"];
                headersRange.Value2 = new object[,]
                {
                    {
                        "Артикул",
                        "Описание",
                        "Кол-во",
                        "Кратность",
                        "Пр-ль",
                        "Скидка",
                        "Цена",
                        "Цена со скидкой",
                        "Стоимость",
                        "Дата и время"
                    }
                };

                // Ширина столбцов
                var columnWidths = new Dictionary<string, double>
                {
                    ["A:A"] = 21,
                    ["B:B"] = 80,
                    ["C:C"] = 10,
                    ["D:I"] = 13,
                    ["J:J"] = 0
                };

                foreach (var width in columnWidths)
                {
                    columnWidthRange = worksheet.Range[width.Key];
                    columnWidthRange.ColumnWidth = width.Value;
                    ReleaseComObject(columnWidthRange);
                    columnWidthRange = null;
                }

                //размечаем границы и правим шрифты
                formattingRange = worksheet.Range["A1", "J500"];
                formattingCells = formattingRange.Cells;
                formattingFont = formattingCells.Font;
                formattingFont.Name = correctFontResources.NameFont;
                formattingFont.Size = correctFontResources.SizeFont;

                // Форматирование
                fontRange = worksheet.Range["A1", "J100"];
                fontRangeFont = fontRange.Font;
                fontRangeFont.Name = correctFontResources.NameFont;
                fontRangeFont.Size = correctFontResources.SizeFont;

                bordersRange = worksheet.Range["A1", "J11"];
                borders = bordersRange.Borders;
                borders.LineStyle = XlLineStyle.xlContinuous;  // Добавлено оформление границ

                focusRange = worksheet.Range["A2", Type.Missing];
                focusRange.Select();   //Фокус на ячейку А2
            }
            finally
            {
                ReleaseComObject(focusRange);
                ReleaseComObject(borders);
                ReleaseComObject(bordersRange);
                ReleaseComObject(fontRangeFont);
                ReleaseComObject(fontRange);
                ReleaseComObject(formattingFont);
                ReleaseComObject(formattingCells);
                ReleaseComObject(formattingRange);
                ReleaseComObject(columnWidthRange);
                ReleaseComObject(headersRange);
                ReleaseComObject(sheets);
                ReleaseComObject(worksheet);
                ReleaseComObject(workbook);
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}

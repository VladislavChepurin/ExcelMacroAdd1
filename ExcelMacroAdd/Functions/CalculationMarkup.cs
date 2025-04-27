using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelMacroAdd.Functions
{
    internal sealed class CalculationMarkup : AbstractFunctions
    {
        private const int ExcelAutoColor = 0;
        private readonly ICorrectFontResources correctFontResources;

        public CalculationMarkup(ICorrectFontResources correctFontResources)
        {
            this.correctFontResources = correctFontResources;
        }

        public override void Start()
        {
            //Проверяем наличие данных в таблице, A1:H9
            bool allCellsAreNull = true;

            // Loop through columns 1-9 and rows 1-8
            for (int col = 1; col <= 9 && allCellsAreNull; col++)
            {
                for (int row = 1; row <= 8 && allCellsAreNull; row++)
                {
                    if (Worksheet.Cells[col, row].Value2 != null)
                    {
                        allCellsAreNull = false;
                        // Exit both loops immediately when first non-null cell is found
                        break;
                    }
                }
            }

            //Проверяем результат переменной
            if (allCellsAreNull)
            {
                // Заголовки столбцов
                var headers = new List<(string Cell, string Value)>
                {
                    ("A1", "Наименование проекта"),
                    ("A2", "Производитель коммутационной аппаратуры"),
                    ("A3", "№п/п"),
                    ("B3", "Наименование щита"),
                    ("C3", "Номер схемы"),
                    ("D3", "Кол-во"),
                    ("E3", "Цена"),
                    ("F3", "Стоимость"),
                    ("G3", "Тип шкафа"),
                    ("H3", "Примечания")
                };

                foreach (var header in headers)
                {
                    Worksheet.Range[header.Cell].Value2 = header.Value;
                }

                // Заливка ячеек
                var coloredCells = new Dictionary<string, XlRgbColor>
                {
                    ["B1"] = XlRgbColor.rgbYellow,
                    ["B2"] = XlRgbColor.rgbGreen
                };

                foreach (var cell in coloredCells)
                {
                    Worksheet.Range[cell.Key].Interior.Color = cell.Value;
                }

                // Ширина столбцов
                var columnWidths = new Dictionary<string, double>
                {
                    ["A:A"] = 22,
                    ["B:B"] = 50,
                    ["C:C"] = 40,
                    ["D:G"] = 10,
                    ["H:H"] = 45
                };

                foreach (var width in columnWidths)
                {
                    Worksheet.Range[width.Key].ColumnWidth = width.Value;
                }

                // Формулы и нумерация
                const int startRow = 4;
                const int endRow = 23;

                for (int row = startRow; row <= endRow; row++)
                {
                    Worksheet.Range[$"F{row}"].Formula = $"=D{row}*E{row}";
                    Worksheet.Range[$"A{row}"].Value2 = (row - 3).ToString();
                }                              

                // Форматирование
                var fontRange = Worksheet.Range[$"A1:H{endRow}"];
                fontRange.Font.Name = correctFontResources.NameFont;
                fontRange.Font.Size = correctFontResources.SizeFont;
                fontRange.Borders.LineStyle = XlLineStyle.xlContinuous;  // Добавлено оформление границ
                fontRange.Rows.AutoFit();
                fontRange.WrapText = true;

                // Оптимизация производительности
                Marshal.ReleaseComObject(fontRange);
            }
            else
            {
                MessageWarning("Внимание! На листе есть данные",
                    "Ошибка разметки");
            }
        }
    }
}

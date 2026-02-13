using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Drawing;
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
                    ("B2", "Наименование проекта"),
                    ("B3", "Производитель коммутационной аппаратуры"),
                    ("B4", "Приннцип расчета"),
                    ("B5", "Дополнительная информация"),
                    ("A7", "№"),
                    ("B7", "Наименование"),
                    ("C7", "Шифр рабочей документации"),
                    ("E7", "Кол-во"),
                    ("F7", "Цена"),
                    ("G7", "Стоимость"),
                    ("H7", "Примечание"),
                    ("I7", "Тип шкафа"),
                    ("J7", "Коментарии")
                };

                foreach (var header in headers)
                {
                    Worksheet.Range[header.Cell].Value2 = header.Value;
                }

                // Заливка ячеек
                var coloredCells = new Dictionary<string, int>
                {
                    ["B2:B6"] = ColorTranslator.ToOle(Color.FromArgb(221, 235, 247)),
                    ["A7:E27"] = ColorTranslator.ToOle(Color.FromArgb(221, 235, 247))
                };

                foreach (var cell in coloredCells)
                {
                    Worksheet.Range[cell.Key].Interior.Color = cell.Value;
                }

                // Ширина столбцов
                var columnWidths = new Dictionary<string, double>
                {
                    ["A:A"] = 2.86,
                    ["B:C"] = 28.57,
                    ["D:D"] = 33.57,                    
                    ["E:E"] = 6.57,
                    ["F:F"] = 14.86,
                    ["G:G"] = 9.71,
                    ["H:H"] = 33.57,
                    ["I:I"] = 11.71,
                    ["J:J"] = 37.57
                };

                foreach (var width in columnWidths)
                {
                    Worksheet.Range[width.Key].ColumnWidth = width.Value;
                }

                // Формулы и нумерация
                const int startRow = 8;
                const int endRow = 27;

                for (int row = startRow; row <= endRow; row++)
                {
                    Worksheet.Range[$"G{row}"].Formula = $"=F{row}*E{row}";
                    Worksheet.Range[$"A{row}"].Value2 = (row - 7).ToString();
                }                              

                // Форматирование
                var fontRange = Worksheet.Range[$"A1:J{endRow}"];
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

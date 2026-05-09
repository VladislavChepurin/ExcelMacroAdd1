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
            Worksheet worksheet = null;
            Range validationRange = null;
            Range projectHeadersRange = null;
            Range tableHeadersRange = null;
            Range coloredRange = null;
            Interior coloredInterior = null;
            Range columnWidthRange = null;
            Range formulaRange = null;
            Range numberRange = null;
            Range formattingRange = null;
            Microsoft.Office.Interop.Excel.Font formattingFont = null;
            Borders formattingBorders = null;
            Range formattingRows = null;

            try
            {
                worksheet = Worksheet;

                //Проверяем наличие данных в таблице, A1:H9
                validationRange = worksheet.Range["A1", "H9"];
                bool allCellsAreNull = AreAllCellsNull(validationRange.Value2);

                //Проверяем результат переменной
                if (allCellsAreNull)
                {
                    projectHeadersRange = worksheet.Range["B2", "B5"];
                    projectHeadersRange.Value2 = new object[,]
                    {
                        { "Наименование проекта" },
                        { "Производитель коммутационной аппаратуры" },
                        { "Приннцип расчета" },
                        { "Дополнительная информация" }
                    };

                    tableHeadersRange = worksheet.Range["A7", "J7"];
                    tableHeadersRange.Value2 = new object[,]
                    {
                        {
                            "№",
                            "Наименование",
                            "Шифр рабочей документации",
                            null,
                            "Кол-во",
                            "Цена",
                            "Стоимость",
                            "Примечание",
                            "Тип шкафа",
                            "Коментарии"
                        }
                    };

                    // Заливка ячеек
                    var coloredCells = new Dictionary<string, int>
                    {
                        ["B2:B6"] = ColorTranslator.ToOle(Color.FromArgb(221, 235, 247)),
                        ["A7:E27"] = ColorTranslator.ToOle(Color.FromArgb(221, 235, 247))
                    };

                    foreach (var cell in coloredCells)
                    {
                        coloredRange = worksheet.Range[cell.Key];
                        coloredInterior = coloredRange.Interior;
                        coloredInterior.Color = cell.Value;

                        ReleaseComObject(coloredInterior);
                        coloredInterior = null;

                        ReleaseComObject(coloredRange);
                        coloredRange = null;
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
                        columnWidthRange = worksheet.Range[width.Key];
                        columnWidthRange.ColumnWidth = width.Value;

                        ReleaseComObject(columnWidthRange);
                        columnWidthRange = null;
                    }

                    // Формулы и нумерация
                    const int startRow = 8;
                    const int endRow = 27;

                    formulaRange = worksheet.Range[$"G{startRow}", $"G{endRow}"];
                    formulaRange.Formula = BuildFormulaValues(startRow, endRow);

                    numberRange = worksheet.Range[$"A{startRow}", $"A{endRow}"];
                    numberRange.Value2 = BuildNumberValues(startRow, endRow);

                    // Форматирование
                    formattingRange = worksheet.Range["A1", $"J{endRow}"];
                    formattingFont = formattingRange.Font;
                    formattingFont.Name = correctFontResources.NameFont;
                    formattingFont.Size = correctFontResources.SizeFont;

                    formattingBorders = formattingRange.Borders;
                    formattingBorders.LineStyle = XlLineStyle.xlContinuous;  // Добавлено оформление границ

                    formattingRows = formattingRange.Rows;
                    formattingRows.AutoFit();
                    formattingRange.WrapText = true;
                }
                else
                {
                    MessageWarning("Внимание! На листе есть данные",
                        "Ошибка разметки");
                }
            }
            finally
            {
                ReleaseComObject(formattingRows);
                ReleaseComObject(formattingBorders);
                ReleaseComObject(formattingFont);
                ReleaseComObject(formattingRange);
                ReleaseComObject(numberRange);
                ReleaseComObject(formulaRange);
                ReleaseComObject(columnWidthRange);
                ReleaseComObject(coloredInterior);
                ReleaseComObject(coloredRange);
                ReleaseComObject(tableHeadersRange);
                ReleaseComObject(projectHeadersRange);
                ReleaseComObject(validationRange);
                ReleaseComObject(worksheet);
            }
        }

        private static bool AreAllCellsNull(object rangeValues)
        {
            if (rangeValues == null)
            {
                return true;
            }

            if (rangeValues is object[,] cellValues)
            {
                for (int row = cellValues.GetLowerBound(0); row <= cellValues.GetUpperBound(0); row++)
                {
                    for (int column = cellValues.GetLowerBound(1); column <= cellValues.GetUpperBound(1); column++)
                    {
                        if (cellValues[row, column] != null)
                        {
                            return false;
                        }
                    }
                }

                return true;
            }

            return false;
        }

        private static object[,] BuildFormulaValues(int startRow, int endRow)
        {
            int rowCount = endRow - startRow + 1;
            var formulas = new object[rowCount, 1];

            for (int row = startRow; row <= endRow; row++)
            {
                formulas[row - startRow, 0] = $"=F{row}*E{row}";
            }

            return formulas;
        }

        private static object[,] BuildNumberValues(int startRow, int endRow)
        {
            int rowCount = endRow - startRow + 1;
            var values = new object[rowCount, 1];

            for (int row = startRow; row <= endRow; row++)
            {
                values[row - startRow, 0] = (row - 7).ToString();
            }

            return values;
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

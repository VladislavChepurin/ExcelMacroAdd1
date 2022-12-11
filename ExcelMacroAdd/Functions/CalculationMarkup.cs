using ExcelMacroAdd.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelMacroAdd.Functions
{
    internal sealed class CalculationMarkup : AbstractFunctions
    {
        private readonly ICorrectFontResources correctFontResources;

        public CalculationMarkup(ICorrectFontResources correctFontResources)
        {
            this.correctFontResources = correctFontResources;
        }

        public override void Start()
        {
            //Проверяем наличие данных в таблице, A1:H9
            Boolean resultCellNull = true;
            for (int column = 1; column <= 9; column++)
            {
                for (int row = 1; row <= 8; row++)
                {
                    if (Worksheet.Cells[column, row].Value2 != null)
                    {
                        resultCellNull = false;
                    }
                }
            }
            //Проверяем результат переменной
            if (resultCellNull)
            {
                //состовляем надписи колонок           
                Worksheet.Range["A1", Type.Missing].Value2 = "Наименование проекта";
                Worksheet.Range["A2", Type.Missing].Value2 = "Производитель коммутационной аппаратуры";
                Worksheet.Range["A3", Type.Missing].Value2 = "№п/п";
                Worksheet.Range["B3", Type.Missing].Value2 = "Наименование щита";
                Worksheet.Range["C3", Type.Missing].Value2 = "Номер схемы";
                Worksheet.Range["D3", Type.Missing].Value2 = "Кол-во";
                Worksheet.Range["E3", Type.Missing].Value2 = "Цена";
                Worksheet.Range["F3", Type.Missing].Value2 = "Стоимость";
                Worksheet.Range["G3", Type.Missing].Value2 = "Тип шкафа";
                Worksheet.Range["H3", Type.Missing].Value2 = "Примечания";

                Worksheet.Range["B1", Type.Missing].Interior.Color = XlRgbColor.rgbYellow;
                Worksheet.Range["B2", Type.Missing].Interior.Color = XlRgbColor.rgbGreen;

                //увеличиваем размер по ширине диапазон ячеек
                Worksheet.Range["A1", Type.Missing].EntireColumn.ColumnWidth = 22;
                Worksheet.Range["B1", Type.Missing].EntireColumn.ColumnWidth = 50;
                Worksheet.Range["C1", Type.Missing].EntireColumn.ColumnWidth = 40;
                Worksheet.Range["D1", "G1"].EntireColumn.ColumnWidth = 10;
                Worksheet.Range["H1", Type.Missing].EntireColumn.ColumnWidth = 45;

                //Вставка формул
                for (int i = 4; i < 10; i++)
                {
                    Worksheet.Range["F" + i, Type.Missing].Formula = string.Format("=D{0}*E{0}", i);
                    Worksheet.Range["A" + i, Type.Missing].Value2 = (i - 3).ToString();
                }

                //размечаем границы и правим шрифты
                Worksheet.Range["A1", "H100"].Cells.Font.Name = correctFontResources.NameFont;
                Worksheet.Range["A1", "H100"].Cells.Font.Size = correctFontResources.SizeFont;

                var excelCells = Worksheet.Range["A1", "H9"];

                excelCells.Rows.AutoFit();
                excelCells.WrapText = true;

                var borderIndex = XlBordersIndex.xlEdgeLeft; //Левая граница
                excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelCells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlEdgeTop; //Верхняя граница
                excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelCells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlEdgeBottom; //Нижняя граница
                excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelCells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlEdgeRight;  //Правая граница
                excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelCells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlInsideHorizontal;  //Внутренняя горизонтальня граница
                excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelCells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlInsideVertical;  //Внутренняя горизонтальня граница
                excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelCells.Borders[borderIndex].ColorIndex = 0;
            }
            else
            {
                MessageWarning("Внимание! На листе есть данные",
                    "Ошибка разметки");
            }
        }
    }
}

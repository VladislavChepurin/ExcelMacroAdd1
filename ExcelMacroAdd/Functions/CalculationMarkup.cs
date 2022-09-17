using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal class CalculationMarkup : AbstractFunctions
    {
        public sealed override void Start()
        {
            //Проверяем наличие данных в таблице, A1:H9
            Boolean resultCellNull = true;
            for (int column = 1; column <= 9; column++)
            {
                for (int row = 1; row <= 8; row++)
                {
                    if (!(worksheet.Cells[column, row].Value2 == null))
                    {
                        resultCellNull = false;
                    }
                }
            }
            //Проверяем результат переменной
            if (resultCellNull)
            {
                //состовляем надписи колонок           
                worksheet.get_Range("A1", Type.Missing).Value2 = "Наименование проекта";
                worksheet.get_Range("A2", Type.Missing).Value2 = "Производитель коммутационной аппаратуры";
                worksheet.get_Range("A3", Type.Missing).Value2 = "№п/п";
                worksheet.get_Range("B3", Type.Missing).Value2 = "Наименование щита";
                worksheet.get_Range("C3", Type.Missing).Value2 = "Номер схемы";
                worksheet.get_Range("D3", Type.Missing).Value2 = "Кол-во";
                worksheet.get_Range("E3", Type.Missing).Value2 = "Цена";
                worksheet.get_Range("F3", Type.Missing).Value2 = "Стоимость";
                worksheet.get_Range("G3", Type.Missing).Value2 = "Тип шкафа";
                worksheet.get_Range("H3", Type.Missing).Value2 = "Примечания";

                worksheet.get_Range("B1", Type.Missing).Interior.Color = XlRgbColor.rgbYellow;
                worksheet.get_Range("B2", Type.Missing).Interior.Color = XlRgbColor.rgbGreen;

                //увеличиваем размер по ширине диапазон ячеек
                worksheet.get_Range("A1", Type.Missing).EntireColumn.ColumnWidth = 22;
                worksheet.get_Range("B1", Type.Missing).EntireColumn.ColumnWidth = 50;
                worksheet.get_Range("C1", Type.Missing).EntireColumn.ColumnWidth = 40;
                worksheet.get_Range("D1", "G1").EntireColumn.ColumnWidth = 10;
                worksheet.get_Range("H1", Type.Missing).EntireColumn.ColumnWidth = 45;

                //Вставка формул
                for (int i = 4; i < 10; i++)
                {
                    worksheet.get_Range("F" + i, Type.Missing).Formula = String.Format("=D{0}*E{0}", i, i);
                    worksheet.get_Range("A" + i, Type.Missing).Value2 = (i - 3).ToString();
                }

                //размечаем границы и правим шрифты
                worksheet.get_Range("A1", "H100").Cells.Font.Name = "Calibri";
                worksheet.get_Range("A1", "H100").Cells.Font.Size = 11;

                var excelcells = worksheet.get_Range("A1", "H9");

                excelcells.Rows.AutoFit();
                excelcells.WrapText = true;

                XlBordersIndex borderIndex;

                borderIndex = XlBordersIndex.xlEdgeLeft; //Левая граница
                excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelcells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlEdgeTop; //Верхняя граница
                excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelcells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlEdgeBottom; //Нижняя граница
                excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelcells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlEdgeRight;  //Правая граница
                excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelcells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlInsideHorizontal;  //Внутренняя горизонтальня граница
                excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelcells.Borders[borderIndex].ColorIndex = 0;

                borderIndex = XlBordersIndex.xlInsideVertical;  //Внутренняя горизонтальня граница
                excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
                excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
                excelcells.Borders[borderIndex].ColorIndex = 0;
            }
            else
            {
                MessageWarning("Внимание! На листе есть данные",
                    "Ошибка разметки");
            }
        }
    }
}

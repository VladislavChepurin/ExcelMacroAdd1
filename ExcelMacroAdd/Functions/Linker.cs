using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;


namespace ExcelMacroAdd.Functions
{
    internal class Linker : AbstractFunctions
    {
        public sealed override void Start()
        {    
            //Создаем коллекцию
            List<int> L = new List<int>();
            // В коллекцию помещаем парсированые имена листов
            foreach (Worksheet sheet in workBook.Sheets)
            {
                int.TryParse(sheet.Name, out int result);
                L.Add(result);
            }
            //Проверяем на налиечие этих номеров в коллекции, если есть, то имя вкладке не присваиваем                 
            if (!L.Contains(worksheet.Index - 1))
            {
                //имя вкладки в зависимости от индекса
                worksheet.Name = (worksheet.Index - 1).ToString();
            }

            //состовляем надписи колонок           
            worksheet.get_Range("A1", Type.Missing).Value2 = "Артикул";
            worksheet.get_Range("B1", Type.Missing).Value2 = "Описание";
            worksheet.get_Range("C1", Type.Missing).Value2 = "Кол-во";
            worksheet.get_Range("D1", Type.Missing).Value2 = "Кратность";
            worksheet.get_Range("E1", Type.Missing).Value2 = "Пр-ль";
            worksheet.get_Range("F1", Type.Missing).Value2 = "Скидка";
            worksheet.get_Range("G1", Type.Missing).Value2 = "Цена";
            worksheet.get_Range("H1", Type.Missing).Value2 = "Цена со скидкой";
            worksheet.get_Range("I1", Type.Missing).Value2 = "Стоимость";

            //увеличиваем размер по ширине диапазон ячеек
            worksheet.get_Range("A1", Type.Missing).EntireColumn.ColumnWidth = 21;
            worksheet.get_Range("B1", Type.Missing).EntireColumn.ColumnWidth = 80;
            worksheet.get_Range("C1", Type.Missing).EntireColumn.ColumnWidth = 10;
            worksheet.get_Range("D1", "I1").EntireColumn.ColumnWidth = 13;

            //размечаем границы и правим шрифты
            worksheet.get_Range("A1", "i500").Cells.Font.Name = "Calibri";
            worksheet.get_Range("A1", "i500").Cells.Font.Size = 11;

            var excelcells = worksheet.get_Range("A1", "I11");
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
    }
}

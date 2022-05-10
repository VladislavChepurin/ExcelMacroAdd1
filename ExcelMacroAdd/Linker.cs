using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelMacroAdd
{
    internal class Linker
    {
        public Linker()
        {
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Workbook workBook = Globals.ThisAddIn.GetActiveWorkBook();

            //Создаем коллекцию
            List<int> L = new List<int>();
            // В коллекцию помещаем парсированые имена листов
            foreach (Excel.Worksheet sheet in workBook.Sheets)
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

            var Excelcells = worksheet.get_Range("A1", "I11");
            Excel.XlBordersIndex BorderIndex;

            BorderIndex = Excel.XlBordersIndex.xlEdgeLeft; //Левая граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlEdgeTop; //Верхняя граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlEdgeBottom; //Нижняя граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlEdgeRight;  //Правая граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlInsideHorizontal;  //Внутренняя горизонтальня граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlInsideVertical;  //Внутренняя горизонтальня граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;
        }
    }
}

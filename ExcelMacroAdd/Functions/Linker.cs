using ExcelMacroAdd.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;


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
            //Создаем коллекцию
            List<int> list = new List<int>();
            // В коллекцию помещаем парсированые имена листов
            foreach (Worksheet sheet in WorkBook.Sheets)
            {
                int.TryParse(sheet.Name, out int result);
                list.Add(result);
            }
            //Проверяем на налиечие этих номеров в коллекции, если есть, то имя вкладке не присваиваем                 
            if (!list.Contains(Worksheet.Index - 1))
            {
                //имя вкладки в зависимости от индекса
                Worksheet.Name = (Worksheet.Index - 1).ToString();
            }

            //состовляем надписи колонок           
            Worksheet.Range["A1", Type.Missing].Value2 = "Артикул";
            Worksheet.Range["B1", Type.Missing].Value2 = "Описание";
            Worksheet.Range["C1", Type.Missing].Value2 = "Кол-во";
            Worksheet.Range["D1", Type.Missing].Value2 = "Кратность";
            Worksheet.Range["E1", Type.Missing].Value2 = "Пр-ль";
            Worksheet.Range["F1", Type.Missing].Value2 = "Скидка";
            Worksheet.Range["G1", Type.Missing].Value2 = "Цена";
            Worksheet.Range["H1", Type.Missing].Value2 = "Цена со скидкой";
            Worksheet.Range["I1", Type.Missing].Value2 = "Стоимость";

            //увеличиваем размер по ширине диапазон ячеек
            Worksheet.Range["A1", Type.Missing].EntireColumn.ColumnWidth = 21;
            Worksheet.Range["B1", Type.Missing].EntireColumn.ColumnWidth = 80;
            Worksheet.Range["C1", Type.Missing].EntireColumn.ColumnWidth = 10;
            Worksheet.Range["D1", "I1"].EntireColumn.ColumnWidth = 13;

            //размечаем границы и правим шрифты
            Worksheet.Range["A1", "i500"].Cells.Font.Name = correctFontResources.NameFont;
            Worksheet.Range["A1", "i500"].Cells.Font.Size = correctFontResources.SizeFont;

            var excelCells = Worksheet.Range["A1", "I11"];

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
    }
}

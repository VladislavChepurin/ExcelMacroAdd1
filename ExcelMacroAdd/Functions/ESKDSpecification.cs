using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMacroAdd.Functions
{
    internal sealed class ESKDSpecification : AbstractFunctions
    {
        public override void Start()
        {          
            //CreateSheetFirst(); //Взависимости от условия вызывается либо этот метод
            CreateSheetNext();    //Либо этот (условие я допишу позже)
        }

        public void CreateSheetNext() //Создаем второй и последующие листы
        {
            MergeCellsSheetNext();
            DrawLineHorizontalSheetNext();
            DrawLineVerticalSheetNext();
            SetingsSheetNext();
            SetTextSheetNext();
        }

        public void CreateSheetFirst() //Создаем первый лист
        {
            MergeCellsSheetFirst();
            DrawLineHorizontalSheetFirst();
            DrawLineVerticalSheetFirst();
            SetingsSheetFirst();
            SetTextSheetFirst();
        }

        private IDictionary<string, double> ColumnWidthsDictionary()
        {
           // Ширина столбцов
           return new Dictionary<string, double>
            {
                ["A:A"] = 2.0,
                ["B:B"] = 2.86,
                ["C:D"] = 2.43,
                ["E:E"] = 1.0,
                ["F:F"] = 1.71,
                ["G:G"] = 3.57,
                ["H:I"] = 4.0,
                ["J:J"] = 2.57,
                ["K:L"] = 2.0,
                ["M:N"] = 2.57,
                ["O:O"] = 2.29,
                ["P:P"] = 6.71,
                ["Q:Q"] = 7.86,
                ["R:R"] = 6.0,
                ["S:X"] = 2.0,
                ["Y:Y"] = 4.71,
                ["Z:Z"] = 2.0
           };
        }

        private IDictionary<string, double> RowHeightsDictionarySheetFirst()
        {            
            return new Dictionary<string, double>
            {
                ["1:1"] = 42.75,  
                ["2:29"] = 23.25,     
                ["30:38"] = 15.0      
            };
        }

        private IDictionary<string, double> RowHeightsDictionarySheetNext()
        {
            return new Dictionary<string, double>
            {
                ["1:1"] = 42.75,    
                ["2:32"] = 23.25,
                ["33:33"] = 0.0,
                ["34:34"] = 15.0,
                ["35:36"] = 7.5,
                ["37:38"] = 15.0
            };
        }


        private void MergeCellsSheetFirst()
        {
            for (int row = 1; row <= 29; row++)
            {
                Worksheet.Range[$"E{row}:F{row}"].Merge();
                Worksheet.Range[$"G{row}:O{row}"].Merge();
                Worksheet.Range[$"P{row}:U{row}"].Merge();
                Worksheet.Range[$"V{row}:W{row}"].Merge();
                Worksheet.Range[$"X{row}:Z{row}"].Merge();
            }

            Worksheet.Range["A18:A21"].Merge();
            Worksheet.Range["B18:B21"].Merge();
            Worksheet.Range["A22:A24"].Merge();
            Worksheet.Range["B22:B24"].Merge();
            Worksheet.Range["A25:A27"].Merge();
            Worksheet.Range["B25:B27"].Merge();
            Worksheet.Range["A28:A32"].Merge();
            Worksheet.Range["B28:B32"].Merge();
            Worksheet.Range["A33:A37"].Merge();
            Worksheet.Range["B33:B37"].Merge();

            for (int row = 30; row <= 32; row++)
            {
                Worksheet.Range[$"D{row}:E{row}"].Merge();
            }

            for (int row = 33; row <= 37; row++)
            {
                Worksheet.Range[$"C{row}:E{row}"].Merge();
            }

            for (int row = 30; row <= 37; row++)
            {
                Worksheet.Range[$"F{row}:H{row}"].Merge();
                Worksheet.Range[$"I{row}:J{row}"].Merge();
                Worksheet.Range[$"K{row}:L{row}"].Merge();
            }

            Worksheet.Range["M30:Z32"].Merge(); // Документ
            Worksheet.Range["M33:R37"].Merge(); // Объект
            Worksheet.Range["S35:Z37"].Merge(); // Производитель 
            Worksheet.Range["S33:U33"].Merge(); // Лит.
            Worksheet.Range["V33:X33"].Merge(); // Лист
            Worksheet.Range["V34:X34"].Merge(); // Лист
            Worksheet.Range["Y33:Z33"].Merge(); // Листов
            Worksheet.Range["Y34:Z34"].Merge(); // Листов
            Worksheet.Range["M38:R38"].Merge(); // Копировал
            Worksheet.Range["S38:Z38"].Merge(); // Формат А4
        }


        private void DrawLineHorizontalSheetFirst()
        {                     
            // Устанавливаем горизонтальные линии
            Worksheet.Range["C1:Z2"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z2"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["A17:B37"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["A17:B37"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C2:Z29"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C2:Z29"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            Worksheet.Range["C29:Z30"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C29:Z30"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C31:L32"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C31:L32"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C32:Z33"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C32:Z33"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["M33:Z37"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["M33:Z37"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;
                       
            Worksheet.Range["C30:L31"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C30:L31"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            Worksheet.Range["C33:L37"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C33:L37"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            // Внешние горизонтальные границы
            Worksheet.Range["C1:Z1"].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z1"].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["A37:Z37"].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["A37:Z37"].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;

        }

        private void DrawLineVerticalSheetFirst()
        {            
            Worksheet.Range["C1:Z29"].Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z29"].Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["A18:B37"].Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["A18:B37"].Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C30:Z37"].Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C30:Z37"].Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;

            // Внешние вертикальные границы
            Worksheet.Range["C1:Z37"].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z37"].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C1:Z37"].Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z37"].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["A18:A37"].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["A18:A37"].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

        }

        private void SetingsSheetFirst()
        {
            // Установка ширины столбцов
            foreach (var width in ColumnWidthsDictionary())
            {
                Worksheet.Range[width.Key].ColumnWidth = width.Value;
            }

            // Установка высоты строк
            foreach (var height in RowHeightsDictionarySheetFirst())
            {
                Worksheet.Range[height.Key].RowHeight = height.Value;
            }

            Worksheet.Range["A1:Z38"].Font.Name = "ISOCPEUR";
            Worksheet.Range["C1:Z29"].Font.Size = 12;
            Worksheet.Range["A18:B37"].Font.Size = 10;
            Worksheet.Range["C30:L37"].Font.Size = 10;
            Worksheet.Range["S33:Z34"].Font.Size = 10;
            Worksheet.Range["M30:Z32"].Font.Size = 18;
            Worksheet.Range["M33:R37"].Font.Size = 14;
            Worksheet.Range["S35:Z37"].Font.Size = 11;
            Worksheet.Range["S35:Z37"].Font.Size = 11;
            Worksheet.Range["M38:Z38"].Font.Size = 10;

            Worksheet.Range["A1:Z1"].WrapText = true;
            Worksheet.Range["A1:F1"].Orientation = 90;
            Worksheet.Range["V1:W1"].Orientation = 90;
            Worksheet.Range["A18:B37"].Orientation = 90;

            Worksheet.Range["A1:Z37"].VerticalAlignment = XlVAlign.xlVAlignCenter;
            Worksheet.Range["A1:Z1"].HorizontalAlignment = XlVAlign.xlVAlignCenter;
            Worksheet.Range["C30:L32"].HorizontalAlignment = XlVAlign.xlVAlignCenter;
            Worksheet.Range["M30:Z37"].HorizontalAlignment = XlVAlign.xlVAlignCenter;
        }

        private void SetTextSheetFirst()
        {
            Worksheet.Range["C1", Type.Missing].Value2 = "Формат";
            Worksheet.Range["D1", Type.Missing].Value2 = "Зона";
            Worksheet.Range["E1", Type.Missing].Value2 = "Поз.";
            Worksheet.Range["G1", Type.Missing].Value2 = "Обозначение";
            Worksheet.Range["P1", Type.Missing].Value2 = "Наименование";
            Worksheet.Range["V1", Type.Missing].Value2 = "Кол.";
            Worksheet.Range["X1", Type.Missing].Value2 = "Примечаение";

            Worksheet.Range["A18", Type.Missing].Value2 = "Подпись и дата";
            Worksheet.Range["A22", Type.Missing].Value2 = "Инв. № дубл.";
            Worksheet.Range["A25", Type.Missing].Value2 = "Взам. инв. №";
            Worksheet.Range["A28", Type.Missing].Value2 = "Подпись и дата";
            Worksheet.Range["A33", Type.Missing].Value2 = "Инв. № подл.";

            Worksheet.Range["C32", Type.Missing].Value2 = "Изм.";
            Worksheet.Range["D32", Type.Missing].Value2 = "Лист";
            Worksheet.Range["F32", Type.Missing].Value2 = "№ докум.";
            Worksheet.Range["I32", Type.Missing].Value2 = "Подп.";
            Worksheet.Range["K32", Type.Missing].Value2 = "Дата";
            Worksheet.Range["C33", Type.Missing].Value2 = "Разраб.";
            Worksheet.Range["C34", Type.Missing].Value2 = "Пров.";
            Worksheet.Range["C36", Type.Missing].Value2 = "Н. контр.";
            Worksheet.Range["C37", Type.Missing].Value2 = "Утв.";

            Worksheet.Range["M30", Type.Missing].Value2 = "Документ";
            Worksheet.Range["M33", Type.Missing].Value2 = "Объект";
            Worksheet.Range["S35", Type.Missing].Value2 = "OOO ...";

            Worksheet.Range["S33", Type.Missing].Value2 = "Лит.";
            Worksheet.Range["V33", Type.Missing].Value2 = "Лист";
            Worksheet.Range["Y33", Type.Missing].Value2 = "Листов";

            Worksheet.Range["M38", Type.Missing].Value2 = "Копировал";
            Worksheet.Range["S38", Type.Missing].Value2 = "Формат А4";

        }

        private void SetingsSheetNext()
        {
            // Установка ширины столбцов
            foreach (var width in ColumnWidthsDictionary())
            {
                Worksheet.Range[width.Key].ColumnWidth = width.Value;
            }

            // Установка высоты строк
            foreach (var height in RowHeightsDictionarySheetNext())
            {
                Worksheet.Range[height.Key].RowHeight = height.Value;
            }

            Worksheet.Range["A1:Z38"].Font.Name = "ISOCPEUR";
            Worksheet.Range["C1:Z32"].Font.Size = 12;
            Worksheet.Range["A18:B37"].Font.Size = 10;
            Worksheet.Range["M34:X37"].Font.Size = 18;

            Worksheet.Range["A1:Z1"].WrapText = true;
            Worksheet.Range["A1:F1"].Orientation = 90;
            Worksheet.Range["V1:W1"].Orientation = 90;
            Worksheet.Range["A18:B37"].Orientation = 90;

            Worksheet.Range["A1:Z37"].VerticalAlignment = XlVAlign.xlVAlignCenter;
            Worksheet.Range["A1:Z1"].HorizontalAlignment = XlVAlign.xlVAlignCenter;
            Worksheet.Range["C34:Z37"].HorizontalAlignment = XlVAlign.xlVAlignCenter;          
        }

        private void DrawLineHorizontalSheetNext()
        {
            // Устанавливаем горизонтальные линии
            Worksheet.Range["C1:Z2"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z2"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["A17:B37"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["A17:B37"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C2:Z32"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C2:Z32"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            Worksheet.Range["C34:L35"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C34:L35"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            Worksheet.Range["C36:L37"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C36:L37"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C32:Z33"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C32:Z33"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["Y33:Z36"].Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["Y33:Z36"].Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlMedium;

            // Внешние горизонтальные границы
            Worksheet.Range["C1:Z1"].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z1"].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["A37:Z37"].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["A37:Z37"].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;

        }

        private void DrawLineVerticalSheetNext()
        {
            Worksheet.Range["C1:Z32"].Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z32"].Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["A18:B37"].Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["A18:B37"].Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C33:Z37"].Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C33:Z37"].Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;

            // Внешние вертикальные границы
            Worksheet.Range["C1:Z37"].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z37"].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["C1:Z37"].Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["C1:Z37"].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;

            Worksheet.Range["A18:A37"].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            Worksheet.Range["A18:A37"].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
        } 

        private void MergeCellsSheetNext()
        {
            for (int row = 1; row <= 32; row++)
            {
                Worksheet.Range[$"E{row}:F{row}"].Merge();
                Worksheet.Range[$"G{row}:O{row}"].Merge();
                Worksheet.Range[$"P{row}:U{row}"].Merge();
                Worksheet.Range[$"V{row}:W{row}"].Merge();
                Worksheet.Range[$"X{row}:Z{row}"].Merge();
            }

            Worksheet.Range["A18:A21"].Merge();
            Worksheet.Range["B18:B21"].Merge();
            Worksheet.Range["A22:A24"].Merge();
            Worksheet.Range["B22:B24"].Merge();
            Worksheet.Range["A25:A27"].Merge();
            Worksheet.Range["B25:B27"].Merge();
            Worksheet.Range["A28:A31"].Merge();
            Worksheet.Range["B28:B31"].Merge();
            Worksheet.Range["A32:A37"].Merge();
            Worksheet.Range["B32:B37"].Merge();

            Worksheet.Range["M34:X37"].Merge();
            Worksheet.Range["Y34:Z35"].Merge();
            Worksheet.Range["Y36:Z37"].Merge();

            Worksheet.Range["C35:C36"].Merge();

            Worksheet.Range["D34:E34"].Merge();
            Worksheet.Range["D35:E36"].Merge();
            Worksheet.Range["D37:E37"].Merge();

            Worksheet.Range["F34:H34"].Merge();
            Worksheet.Range["F35:H36"].Merge();
            Worksheet.Range["F37:H37"].Merge();

            Worksheet.Range["I34:J34"].Merge();
            Worksheet.Range["I35:J36"].Merge();
            Worksheet.Range["I37:J37"].Merge();

            Worksheet.Range["K34:L34"].Merge();
            Worksheet.Range["K35:L36"].Merge();
            Worksheet.Range["K37:L37"].Merge();

            Worksheet.Range["M38:R38"].Merge(); // Копировал
            Worksheet.Range["S38:Z38"].Merge(); // Формат А4
        }

        private void SetTextSheetNext()
        {
            Worksheet.Range["C1", Type.Missing].Value2 = "Формат";
            Worksheet.Range["D1", Type.Missing].Value2 = "Зона";
            Worksheet.Range["E1", Type.Missing].Value2 = "Поз.";
            Worksheet.Range["G1", Type.Missing].Value2 = "Обозначение";
            Worksheet.Range["P1", Type.Missing].Value2 = "Наименование";
            Worksheet.Range["V1", Type.Missing].Value2 = "Кол.";
            Worksheet.Range["X1", Type.Missing].Value2 = "Примечаение";

            Worksheet.Range["A18", Type.Missing].Value2 = "Подпись и дата";
            Worksheet.Range["A22", Type.Missing].Value2 = "Инв. № дубл.";
            Worksheet.Range["A25", Type.Missing].Value2 = "Взам. инв. №";
            Worksheet.Range["A28", Type.Missing].Value2 = "Подпись и дата";
            Worksheet.Range["A32", Type.Missing].Value2 = "Инв. № подл.";

            Worksheet.Range["C37", Type.Missing].Value2 = "Изм.";
            Worksheet.Range["D37", Type.Missing].Value2 = "Лист";
            Worksheet.Range["F37", Type.Missing].Value2 = "№ докум.";
            Worksheet.Range["I37", Type.Missing].Value2 = "Подп.";
            Worksheet.Range["K37", Type.Missing].Value2 = "Дата";
            Worksheet.Range["Y34", Type.Missing].Value2 = "Лист";

            Worksheet.Range["M34", Type.Missing].Value2 = "Документ";

            Worksheet.Range["M38", Type.Missing].Value2 = "Копировал";
            Worksheet.Range["S38", Type.Missing].Value2 = "Формат А4";
        }
    }
}

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd
{
    public partial class Ribbon1
    {
            private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void button1_Click(object sender, RibbonControlEventArgs e) //Удаление формул
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;
            cell.Value = cell.Value;                    //Удаляем формулы
            worksheet.get_Range("A1", "A1").Select();   //Фокус на ячейку А1
        }

        private void button2_Click(object sender, RibbonControlEventArgs e) //Разметка шаблона расчетов
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;

            //состовляем надписи колонок           
            worksheet.get_Range("A1").Value2 = "Артикул";
            worksheet.get_Range("B1").Value2 = "Описание";
            worksheet.get_Range("C1").Value2 = "Кол-во";
            worksheet.get_Range("D1").Value2 = "Кратность";
            worksheet.get_Range("E1").Value2 = "Пр-ль";
            worksheet.get_Range("F1").Value2 = "Скидка";
            worksheet.get_Range("G1").Value2 = "Цена";
            worksheet.get_Range("H1").Value2 = "Цена со скидкой";
            worksheet.get_Range("I1").Value2 = "Стоимость";

            //увеличиваем размер по ширине диапазон ячеек
            worksheet.get_Range("A1", "A1").EntireColumn.ColumnWidth = 21;
            worksheet.get_Range("B1", "B1").EntireColumn.ColumnWidth = 80;
            worksheet.get_Range("C1", "C1").EntireColumn.ColumnWidth = 10;
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

        private void button3_Click(object sender, RibbonControlEventArgs e) //Корпуса щитов
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;

            if  (application.ActiveWorkbook.Name == "_Журнал учета НКУ 2022.xlsx") {
                
            
            
            
            
            }
            else
            {
                MessageBox.Show(
                "Программа работает только в файле _Журнал учета НКУ 2022.xlsx \n Пожайлуста откройте целевую книгу и запустите программу.",
                "Ошибка вызова",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
        }
    }
}

using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelMacroAdd
{
    internal class WriteExcel
    {
        public WriteExcel(DataInXml dataInXml, int rowsLine = default, string getArticle = null, int quantity = default, bool link = false)
        {
            int endRow = default;
            //Стороки подключения к Excel
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Range cell = Globals.ThisAddIn.GetActiveCell();
            //Вычисляем столбец на который установлен фокус
            int firstRow = cell.Row;
            if (firstRow == 1) firstRow++;
            int countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
            if (countRow == 1 || !(getArticle is null))
            {
                firstRow += rowsLine ;
            }
            else
            {
                endRow = firstRow + countRow;
            }                   
            // Заполняем таблицу
            do
            {
                if (!(getArticle is null))
                {
                    worksheet.get_Range("A" + firstRow).Value2 = getArticle;
                }
                worksheet.get_Range("B" + firstRow).FormulaLocal = String.Format(
                    dataInXml.GetDataInXml("Formula_1"), firstRow);    //Столбец "Описание". Вызывает формулу Formula_1
                if (!(quantity == 0)) worksheet.get_Range("C" + firstRow).Value2 = quantity;
                worksheet.get_Range("D" + firstRow).FormulaLocal = String.Format(
                    dataInXml.GetDataInXml("Formula_2"), firstRow);    //Столбец "Кратность". Вызывает формулу Formula_2
                worksheet.get_Range("E" + firstRow).Value2 = Replace.RepleceVendorTable(dataInXml.Vendor);
                worksheet.get_Range("F" + firstRow).Value2 = dataInXml.GetDataInXml("Discont");         //Столбец "Скидка". Вызывает значение Discont
                worksheet.get_Range("G" + firstRow).FormulaLocal = String.Format(
                    dataInXml.GetDataInXml("Formula_3"), firstRow);     //Столбец "Цена". Вызывает формулу Formula_3
                worksheet.get_Range("H" + firstRow).Formula = String.Format("=G{0}*(100-F{0})/100", firstRow);
                worksheet.get_Range("I" + firstRow).Formula = String.Format("=H{0}*C{0}", firstRow);
                firstRow++;
            }
            while (endRow > firstRow);
            if (link) _ = new Linker();                          // Если стоит галочка, то запускается разметчик листов
        }
    }
}

using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;


namespace ExcelMacroAdd
{
    internal class WriteExcel
    {

        public WriteExcel(string getArticle, string vendor, int rows, int quantity, Boolean link)
        {
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Workbook workBook = Globals.ThisAddIn.GetActiveWorkBook();
            Range cell = Globals.ThisAddIn.GetActiveCell();
            
            int firstRow = cell.Row;
            if (firstRow == 1) firstRow++;

            int lastRow = rows + firstRow;
            // Заполняем таблицу
            worksheet.get_Range("A" + lastRow).Value2 = getArticle;
            worksheet.get_Range("B" + lastRow).Formula = GetFormulaInXml(FuncReplece(vendor), 1, lastRow);
            worksheet.get_Range("C" + lastRow).Value2 = quantity;
            worksheet.get_Range("D" + lastRow).Formula = GetFormulaInXml(FuncReplece(vendor), 2, lastRow);
            worksheet.get_Range("E" + lastRow).Value2 = FuncReplece(vendor);
            worksheet.get_Range("F" + lastRow).Formula = GetFormulaInXml(FuncReplece(vendor), 4, lastRow);
            worksheet.get_Range("G" + lastRow).Formula = GetFormulaInXml(FuncReplece(vendor), 3, lastRow);
            worksheet.get_Range("H" + lastRow).Formula = String.Format("=G{0}*(100-F{0})/100", lastRow);
            worksheet.get_Range("I" + lastRow).Formula = String.Format("=H{0}*C{0}", lastRow);

            if (link) _ = new Linker(); // Если стоит галочка, то запускается разметчик листов
                 
        }

        private string FuncReplece(string mReplase)                          // Функция замены // индус заплачит от умиления IEK ВА47 - кирилица, IEK BA47М - латиница Переписать!!!
        {
            return mReplase.Replace("IEK ВА47", "IEK").Replace("IEK BA47М", "IEK").Replace("EKF PROxima", "EKF").Replace("EKF AVERS", "EKF").Replace("Schneider", "SE");
        }

        private string GetFormulaInXml(string vendor, int formulaNum, int rows)
        {
            return String.Empty;
        }




    }
}

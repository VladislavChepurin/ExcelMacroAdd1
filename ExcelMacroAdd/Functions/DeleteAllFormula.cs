using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal sealed class DeleteAllFormula : AbstractFunctions
    {
        public override void Start()
        {
            bool isVisible = true;
            var works = WorkBook.Sheets;
            foreach (Excel.Worksheet sheet in works)
            {
                if (sheet.Visible == Excel.XlSheetVisibility.xlSheetHidden)
                {
                    sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    isVisible = false;
                }
                sheet.Activate();
                if (sheet.Index == 1) continue;
                sheet.Range["A2", "G500"].Value = sheet.Range["A2", "G500"].Value;
                sheet.Range["A1", Type.Missing].Select();   //Фокус на ячейку А1

                if (!isVisible)
                {
                    sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    isVisible = true;
                }
            }
        }
    }
}
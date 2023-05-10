using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal sealed class DeleteAllFormula : AbstractFunctions
    {
        public override void Start()
        {   
            var works = WorkBook.Sheets;
            foreach (Excel.Worksheet sheet in works)
            {
                sheet.Activate();

                if (sheet.Index == 1)
                    continue;

                if (sheet.Visible == Excel.XlSheetVisibility.xlSheetHidden)
                    continue;                  

                sheet.Range["A2", "G500"].Value = sheet.Range["A2", "G500"].Value;
                sheet.Range["A1", Type.Missing].Select();   //Фокус на ячейку А1                
            }
        }
    }
}
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal class DeleteAllFormula : AbstractFunctions
    {
        public sealed override void Start()
        {
            foreach (Excel.Worksheet sheet in WorkBook.Sheets)
            {
                sheet.Activate();
                if (sheet.Index == 1) continue;
                sheet.Range["A2", "G500"].Value = sheet.Range["A2", "G500"].Value;
                sheet.Range["A1", Type.Missing].Select();   //Фокус на ячейку А1
            }
        }
    }
}

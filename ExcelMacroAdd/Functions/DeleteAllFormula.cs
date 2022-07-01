using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal class DeleteAllFormula : AbstractFunctions
    {
        public override void Start()
        {
            foreach (Excel.Worksheet sheet in workBook.Sheets)
            {
                sheet.Activate();
                if (!(sheet.Index == 1))
                {
                    sheet.get_Range("A2", "G500").Value = sheet.get_Range("A2", "G500").Value;
                    sheet.get_Range("A1", Type.Missing).Select();   //Фокус на ячейку А1
                }
            }
        }
    }
}

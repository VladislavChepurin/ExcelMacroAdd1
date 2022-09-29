using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelMacroAdd.Functions
{
    internal class EditCalculation : AbstractFunctions
    {
        public sealed override void Start()
        {
            foreach (Worksheet sheet in WorkBook.Sheets)
            {
                sheet.Activate();
                if (sheet.Index == 1) continue;
                sheet.Range["A1", "i500"].Cells.Font.Name = "Calibri";
                sheet.Range["A1", "i500"].Cells.Font.Size = 11;
                sheet.Range["D1", Type.Missing].EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                sheet.Range["D1", Type.Missing].Value2 = "Кратность";
                sheet.Range["D1", Type.Missing].EntireColumn.ColumnWidth = 10;
            }
        }
    }
}

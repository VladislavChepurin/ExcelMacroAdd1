using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelMacroAdd.Functions
{
    internal class EditCalculation : AbstractFunctions
    {
        protected internal override void Start()
        {
            foreach (Worksheet sheet in workBook.Sheets)
            {
                sheet.Activate();
                if (sheet.Index != 1)
                {
                    sheet.get_Range("A1", "i500").Cells.Font.Name = "Calibri";
                    sheet.get_Range("A1", "i500").Cells.Font.Size = 11;
                    sheet.get_Range("D1", Type.Missing).EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                    sheet.get_Range("D1", Type.Missing).Value2 = "Кратность";
                    sheet.get_Range("D1", Type.Missing).EntireColumn.ColumnWidth = 10;
                }
            }
        }
    }
}

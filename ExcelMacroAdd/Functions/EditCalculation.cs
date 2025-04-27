using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelMacroAdd.Functions
{
    internal sealed class EditCalculation : AbstractFunctions
    {
        private readonly ICorrectFontResources correctFontResources;

        public EditCalculation(ICorrectFontResources correctFontResources)
        {
            this.correctFontResources = correctFontResources;
        }

        public override void Start()
        {
            foreach (Worksheet sheet in WorkBook.Sheets)
            {
                sheet.Activate();
                if (sheet.Index == 1) continue;
                sheet.Range["A1", "i500"].Cells.Font.Name = correctFontResources.NameFont;
                sheet.Range["A1", "i500"].Cells.Font.Size = correctFontResources.SizeFont;
                sheet.Range["D1", Type.Missing].EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                sheet.Range["D1", Type.Missing].Value2 = "Кратность";
                sheet.Range["D1", Type.Missing].EntireColumn.ColumnWidth = 10;
            }        
        }
    }
}

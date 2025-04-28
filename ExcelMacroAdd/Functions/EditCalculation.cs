using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

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
                // Пропускаем лист по индексу 1 и скрытые листы
                if (sheet.Index == 1 || sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    continue;              
                sheet.Range["A1", "i500"].Cells.Font.Name = correctFontResources.NameFont;
                sheet.Range["A1", "i500"].Cells.Font.Size = correctFontResources.SizeFont;
                sheet.Cells[1, 4].EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                sheet.Cells[1, 4].Value2 = "Кратность";
                sheet.Cells[1, 4].EntireColumn.ColumnWidth = 10;
            }        
        }
    }
}

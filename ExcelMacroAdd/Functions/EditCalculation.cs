using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
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
            Workbook workbook = null;
            Sheets sheets = null;

            try
            {
                workbook = WorkBook;
                sheets = workbook.Sheets;
                int sheetCount = sheets.Count;

                for (int index = 1; index <= sheetCount; index++)
                {
                    Worksheet sheet = null;
                    Range formattingRange = null;
                    Range formattingCells = null;
                    Font formattingFont = null;
                    Range insertCell = null;
                    Range insertedColumn = null;

                    try
                    {
                        sheet = (Worksheet)sheets[index];

                        // Пропускаем лист по индексу 1 и скрытые листы
                        if (sheet.Index == 1 || sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            continue;
                        }

                        formattingRange = sheet.Range["A1", "I500"];
                        formattingCells = formattingRange.Cells;
                        formattingFont = formattingCells.Font;
                        formattingFont.Name = correctFontResources.NameFont;
                        formattingFont.Size = correctFontResources.SizeFont;

                        insertCell = (Range)sheet.Cells[1, 4];
                        insertedColumn = insertCell.EntireColumn;
                        insertedColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        insertCell.Value2 = "Кратность";
                        insertedColumn.ColumnWidth = 10;
                    }
                    finally
                    {
                        ReleaseComObject(insertedColumn);
                        ReleaseComObject(insertCell);
                        ReleaseComObject(formattingFont);
                        ReleaseComObject(formattingCells);
                        ReleaseComObject(formattingRange);
                        ReleaseComObject(sheet);
                    }
                }
            }
            finally
            {
                ReleaseComObject(sheets);
                ReleaseComObject(workbook);
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}

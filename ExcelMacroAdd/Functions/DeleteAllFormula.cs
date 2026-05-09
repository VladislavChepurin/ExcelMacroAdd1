using ExcelMacroAdd.Services;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal sealed class DeleteAllFormula : AbstractFunctions
    {
        public override void Start()
        {
            // Отключаем обновление интерфейса для повышения производительности
            Application.ScreenUpdating = false;

            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet activeSheet = null;
            Excel.Range focusRange = null;

            try
            {
                workbook = WorkBook;
                sheets = workbook.Worksheets;
                int sheetCount = sheets.Count;

                for (int index = 1; index <= sheetCount; index++)
                {
                    Excel.Worksheet sheet = null;
                    Excel.Range targetRange = null;

                    try
                    {
                        sheet = (Excel.Worksheet)sheets[index];
                        // Пропускаем лист по индексу 1 и скрытые листы
                        if (sheet.Index == 1 || sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            continue;
                        }

                        // Получаем диапазон без активации листа
                        targetRange = sheet.Range["A2:G500"];

                        // Заменяем формулы на статические значения
                        object values = targetRange.Value2;
                        targetRange.Value2 = values;
                    }
                    finally
                    {
                        ReleaseComObject(targetRange);
                        ReleaseComObject(sheet);
                    }
                }

                // Возвращаем фокус на A1 активного листа (опционально)
                activeSheet = Application.ActiveSheet as Excel.Worksheet;
                if (activeSheet != null)
                {
                    focusRange = activeSheet.Range["A1"];
                    focusRange.Select();
                }
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка: {ex.Message}\n{ex.StackTrace}",
                               "Ошибка обработки");
                Logger.LogException(ex);
            }
            finally
            {
                // Восстанавливаем обновление экрана
                Application.ScreenUpdating = true;

                ReleaseComObject(focusRange);
                ReleaseComObject(activeSheet);
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

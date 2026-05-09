using ExcelMacroAdd.Services;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace ExcelMacroAdd.Functions
{
    internal sealed class DeleteFormula : AbstractFunctions
    {
        public override void Start()
        {
            Worksheet worksheet = null;
            Range cell = null;
            Range focusRange = null;

            try
            {
                worksheet = Worksheet;
                cell = Cell;
                cell.Value2 = cell.Value2;                      //Удаляем формулы
                focusRange = worksheet.Range["A1", Type.Missing];
                focusRange.Select();   //Фокус на ячейку А1
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка: {ex.Message}\n{ex.StackTrace}",
                               "Ошибка обработки");
                Logger.LogException(ex);
            }
            finally
            {
                ReleaseComObject(focusRange);
                ReleaseComObject(cell);
                ReleaseComObject(worksheet);
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

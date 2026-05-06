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
            Range cell = Cell;
            try
            {
                cell.Value2 = cell.Value2;                      //Удаляем формулы
                Worksheet.Range["A1", Type.Missing].Select();   //Фокус на ячейку А1   
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка: {ex.Message}\n{ex.StackTrace}",
                               "Ошибка обработки");
                Logger.LogException(ex);
            }
            finally
            {
                if (cell != null)
                {
                    Marshal.ReleaseComObject(cell);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
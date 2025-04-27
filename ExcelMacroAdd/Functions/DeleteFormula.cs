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
            try
            {
                Cell.Value2 = Cell.Value2;                      //Удаляем формулы
                Worksheet.Range["A1", Type.Missing].Select();   //Фокус на ячейку А1   
                Marshal.ReleaseComObject(Cell);
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка: {ex.Message}\n{ex.StackTrace}",
                               "Ошибка обработки");
                Logger.LogException(ex);
            }
            finally
            {
                if (Cell != null)
                {
                    Marshal.ReleaseComObject(Cell);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}


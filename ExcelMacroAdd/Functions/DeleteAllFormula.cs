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

            // Инициализируем переменную для коллекции листов
            Excel.Sheets sheets = null;

            try
            {
                sheets = WorkBook.Worksheets;

                foreach (Excel.Worksheet sheet in sheets)
                {
                    // Пропускаем лист по индексу 1 и скрытые листы
                    if (sheet.Index == 1 || sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                        continue;

                    // Получаем диапазон без активации листа
                    Excel.Range targetRange = sheet.Range["A2:G500"];

                    // Заменяем формулы на статические значения
                    object[,] values = (object[,])targetRange.Value2;
                    targetRange.Value2 = values;

                    // Явное освобождение ресурсов диапазона
                    Marshal.ReleaseComObject(targetRange);
                }

                // Возвращаем фокус на A1 активного листа (опционально)
                ((Excel.Worksheet)Application.ActiveSheet)?.Range["A1"].Select();
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

                // Освобождаем COM-объекты
                if (sheets != null)
                {
                    Marshal.ReleaseComObject(sheets);                    
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
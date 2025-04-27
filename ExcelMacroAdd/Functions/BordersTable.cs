using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;
using ExcelMacroAdd.Services;

//Rewiew OK 21.04.2025
namespace ExcelMacroAdd.Functions
{
    internal sealed class BordersTable : AbstractFunctions
    {
        private const int ExcelAutoColor = 0;

        public override void Start()
        {
            try
            {
                Application.ScreenUpdating = false;

                var excelCells = Application.Selection as Range;

                if (excelCells == null)
                {
                    MessageInformation("Выделите диапазон ячеек.", "Внимание!");                   
                    return;
                } 
                
                excelCells.Borders.LineStyle = XlLineStyle.xlContinuous;  // Добавлено оформление границ  
            }
            catch (COMException ex)
            {
                Logger.LogException(ex, "Ошибка Excel");             
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Неизвестная ошибка");                
            }
            finally
            {
                Application.ScreenUpdating = true;
            }
        }       
    }
}

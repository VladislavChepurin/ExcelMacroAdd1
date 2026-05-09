using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelMacroAdd.Functions
{
    internal sealed class CorrectFont : AbstractFunctions
    {
        private readonly ICorrectFontResources correctFontResources;

        public CorrectFont(ICorrectFontResources correctFontResources)
        {
            this.correctFontResources = correctFontResources;
        }

        public override void Start()
        {
            Range excelCells = null;
            Font selectionFont = null;

            try
            {
                excelCells = Application.Selection as Range;
                if (excelCells == null)
                {
                    MessageInformation("Выделите диапазон ячеек.", "Внимание!");
                    return;
                }

                selectionFont = excelCells.Font;
                selectionFont.Name = correctFontResources.NameFont;
                selectionFont.Size = correctFontResources.SizeFont;
            }
            catch (COMException ex)
            {
                MessageError("В файле appSettings.json установлены не верные параметры шрифта, пожайлуста установите правильные и доступные значения.", "Ошибка параметров шрифта");
                Logger.LogException(ex);
            }
            finally
            {
                ReleaseComObject(selectionFont);
                ReleaseComObject(excelCells);
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

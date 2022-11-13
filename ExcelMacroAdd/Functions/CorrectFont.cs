using System;
using System.Runtime.InteropServices;
using ExcelMacroAdd.Interfaces;

namespace ExcelMacroAdd.Functions
{
    internal class CorrectFont : AbstractFunctions
    {
        private readonly ICorrectFontResources correctFontResources;
        public CorrectFont(ICorrectFontResources correctFontResources)
        {
            this.correctFontResources = correctFontResources;
        }

        public sealed override void Start()
        {
            var excelCells = Application.Selection;

            try
            {
                excelCells.Font.Name = correctFontResources.NameFont;
                excelCells.Font.Size = correctFontResources.SizeFont;
            }
            catch (COMException)
            {
                MessageError( "В файле appSettings.json установлены не верные параметры шрифта, пожайлуста установите правильные и доступные значения.","Ошибка параметров шрифта" );
            }
        }
    }
}

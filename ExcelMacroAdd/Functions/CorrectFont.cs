using ExcelMacroAdd.Serializable.Entity.Interfaces;
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
            var excelCells = Application.Selection;

            try
            {
                excelCells.Font.Name = correctFontResources.NameFont;
                excelCells.Font.Size = correctFontResources.SizeFont;
            }
            catch (COMException)
            {
                MessageError("В файле appSettings.json установлены не верные параметры шрифта, пожайлуста установите правильные и доступные значения.", "Ошибка параметров шрифта");
            }
        }
    }
}

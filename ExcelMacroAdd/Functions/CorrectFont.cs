namespace ExcelMacroAdd.Functions
{
    internal class CorrectFont : AbstractFunctions
    {
        public sealed override void Start()
        {
            var excelCells = Application.Selection;
            excelCells.Font.Name = "Calibri";
            excelCells.Font.Size = 11;
        }
    }
}

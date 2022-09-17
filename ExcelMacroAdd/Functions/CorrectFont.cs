namespace ExcelMacroAdd.Functions
{
    internal class CorrectFont : AbstractFunctions
    {
        public sealed override void Start()
        {
            var excelcells = application.Selection;
            excelcells.Font.Name = "Calibri";
            excelcells.Font.Size = 11;
        }
    }
}

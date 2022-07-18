namespace ExcelMacroAdd.Functions
{
    internal class CorrectFont : AbstractFunctions
    {
        protected internal override void Start()
        {
            var excelcells = application.Selection;
            excelcells.Font.Name = "Calibri";
            excelcells.Font.Size = 11;
        }
    }
}

using Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    abstract class AbstractFunctions
    {
        internal readonly Application application = Globals.ThisAddIn.GetApplication();
        internal readonly Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
        internal readonly Range cell = Globals.ThisAddIn.GetActiveCell();
        internal readonly Workbook workBook = Globals.ThisAddIn.GetActiveWorkBook();
        public abstract void Start();
    }
}

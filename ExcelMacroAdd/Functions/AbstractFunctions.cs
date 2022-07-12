using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    abstract class AbstractFunctions
    {
        protected readonly Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.GetApplication();
        protected readonly Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
        protected readonly Range cell = Globals.ThisAddIn.GetActiveCell();
        protected readonly Workbook workBook = Globals.ThisAddIn.GetActiveWorkBook();
        public abstract void Start();
        public void MessageWrongNameJournal()
        {
            MessageBox.Show(
                "Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                "Имя книги не совпадает с целевой",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}

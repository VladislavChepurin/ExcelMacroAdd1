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
        protected internal abstract void Start();

        protected internal void MessageInformation(string textMessage, string textAtribute)
        {
            MessageBox.Show(textMessage,
                textAtribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        protected internal void MessageWarning(string textMessage, string textAtribute)
        {
            MessageBox.Show(textMessage,
                textAtribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        protected internal void MessageError(string textMessage, string textAtribute)
        {
            MessageBox.Show(textMessage,
                textAtribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}

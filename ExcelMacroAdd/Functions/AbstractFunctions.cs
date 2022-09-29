using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal abstract class AbstractFunctions
    {
        protected readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisAddIn.GetApplication();
        protected readonly Worksheet Worksheet = Globals.ThisAddIn.GetActiveWorksheet();
        protected readonly Range Cell = Globals.ThisAddIn.GetActiveCell();
        protected readonly Workbook WorkBook = Globals.ThisAddIn.GetActiveWorkBook();
        public abstract void Start();

        protected internal void MessageInformation(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        protected internal void MessageWarning(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        protected internal void MessageError(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}

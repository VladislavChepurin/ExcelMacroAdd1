using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd
{
    public sealed partial class ThisAddIn
    {
        private NewRibbon _ribbon;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new NewRibbon();
            return _ribbon;
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Освобождаем DbContext, MemoryCache и другие ресурсы аддина
            _ribbon?.Dispose();
        }

        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        public Excel.Workbook GetActiveWorkBook()
        {
            return (Excel.Workbook)Application.ActiveWorkbook;
        }

        public Excel.Range GetActiveCell()
        {
            return (Excel.Range)Application.Selection;
        }

        public Excel.Application GetApplication()
        {
            return Application;
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
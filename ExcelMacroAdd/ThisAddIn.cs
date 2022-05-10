using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelMacroAdd
{
    public sealed partial class ThisAddIn

    {       

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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

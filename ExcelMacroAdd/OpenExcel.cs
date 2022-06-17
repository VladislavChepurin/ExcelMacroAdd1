using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMacroAdd
{
    internal class OpenExcel
    {
        private string openVendor;

        public string OpenVendor
        {
            set { openVendor = value;}             

            get {throw new Exception("Попытка доступа к закрытому свойству");}
        }
        public OpenExcel()
        {
            //Класс для откррытия Excel файла и управления жизненым циклом.
        }

        public void TempOpenExcelFile()
        {

        }

    }
}

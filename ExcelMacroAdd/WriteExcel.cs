using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ExcelMacroAdd
{
    internal class WriteExcel
    {
        // Folders AppData content Settings.xml
        readonly string file = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\AddIns\ExcelMacroAdd\Settings.xml";

        public WriteExcel()
        {
            // Конструктор класса, прям Капитан Очевидность
        }
        public void FuncWrite()
        {
            string _getArticle= GetArticle;
            string _vendor = Vendor;
            int _quantity = Quantity;
            int _rows = Rows;
            Boolean _link = Link;
            //Стороки подключения к Excel
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Workbook workBook = Globals.ThisAddIn.GetActiveWorkBook();
            Range cell = Globals.ThisAddIn.GetActiveCell();
            //Вычисляем столбец на который установлен фокус
            int firstRow = cell.Row; 
            if (firstRow == 1) firstRow++;
            int lastRow = _rows + firstRow;

            // Заполняем таблицу
            worksheet.get_Range("A" + lastRow).Value2 = _getArticle;
            worksheet.get_Range("B" + lastRow).FormulaLocal = String.Format(
                GetDataInXml(_vendor, "Formula_1"), lastRow);    //Столбец "Описание". Вызывает формулу Formula_1
            worksheet.get_Range("C" + lastRow).Value2 = _quantity;
            worksheet.get_Range("D" + lastRow).FormulaLocal = String.Format(
                GetDataInXml(_vendor, "Formula_2"), lastRow);    //Столбец "Кратность". Вызывает формулу Formula_2
            worksheet.get_Range("E" + lastRow).Value2 = Replace.RepleceVendorTable(_vendor);
            worksheet.get_Range("F" + lastRow).Value2 = GetDataInXml(_vendor, "Discont");         //Столбец "Скидка". Вызывает значение Discont
            worksheet.get_Range("G" + lastRow).FormulaLocal = String.Format(
                GetDataInXml(_vendor, "Formula_3"), lastRow);     //Столбец "Цена". Вызывает формулу Formula_3
            worksheet.get_Range("H" + lastRow).Formula = String.Format("=G{0}*(100-F{0})/100", lastRow);
            worksheet.get_Range("I" + lastRow).Formula = String.Format("=H{0}*C{0}", lastRow);

            if (_link) _ = new Linker();                          // Если стоит галочка, то запускается разметчик листов
        }
        
        private string GetDataInXml(string vendor, string element)
        {           
            string middle = default;
            try
            {
                XDocument xdoc = XDocument.Load(file);
                var toDiscont = xdoc.Element("MetaSettings")?   // получаем корневой узел MetaSettings
                    .Elements("Vendor")                         // получаем все элементы Vendor                               
                    .Where(p => p.Attribute("vendor")?.Value == Replace.RepleceVendorTable(vendor))
                    .Select(p => new                            // для каждого объекта создаем анонимный объект
                    {
                        dataXml = p.Element(element)?.Value
                    });

                if (toDiscont != null)
                {
                    foreach (var data in toDiscont)
                    {
                      middle = data.dataXml;                       
                    }
                }
                return middle ?? String.Empty;
            }
            catch (Exception)
            {
                return String.Empty;
            }        
        }

        public string GetArticle { get; set; }

        public string Vendor { get; set; }

        public int Rows { get; set; }

        public int Quantity { get; set; }

        public Boolean Link { get; set; }

    }
}

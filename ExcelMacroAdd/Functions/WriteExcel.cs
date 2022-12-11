using ExcelMacroAdd.Interfaces;
using System.Collections.Generic;

namespace ExcelMacroAdd.Functions
{
    internal sealed class WriteExcel : AbstractFunctions 
    {
        private readonly IDataInXml dataInXml;
        private readonly string vendor;
        private readonly int rowsLine;
        private readonly string getArticle;
        private readonly int quantity;

        public WriteExcel(IDataInXml dataInXml, string vendor ,int rowsLine = default, string getArticle = null, int quantity = default)
        {
            this.dataInXml = dataInXml;
            this.vendor = vendor;
            this.rowsLine = rowsLine;
            this.getArticle = getArticle;
            this.quantity = quantity;
        }

        public override void Start()
        {         
            int endRow = default;          
            //Вычисляем столбец на который установлен фокус
            int firstRow = Cell.Row;
            if (firstRow == 1) firstRow++;
            int countRow = Cell.Rows.Count;          // Вычисляем кол-во выделенных строк
            if (countRow == 1 || !(getArticle is null))
            {
                firstRow += rowsLine;
            }
            else
            {
                endRow = firstRow + countRow;
            }
            // Заполняем таблицу
            do
            {
                if (!(getArticle is null)) Worksheet.Range["A" + firstRow].Value2 = getArticle;
                /*
                 * Функция dataInXml.Value.ReadFileXml("Formula_1") возвращает строку типа:
                 * =ВПР(A{0};'C:\Users\..
                 * Функция String.Format подставляет вместо {0} числовое значение firstRow
                */
                var objectVendor = dataInXml.ReadElementXml(vendor, dataInXml.ReadFileXml());
                Worksheet.Range["B" + firstRow].FormulaLocal = string.Format(objectVendor.Formula_1, firstRow);    //Столбец "Описание". Вызывает формулу Formula_1            
                if (quantity != 0) Worksheet.Range["C" + firstRow].Value2 = quantity;
                Worksheet.Range["D" + firstRow].FormulaLocal = string.Format(objectVendor.Formula_2, firstRow);    //Столбец "Кратность". Вызывает формулу Formula_2
                Worksheet.Range["E" + firstRow].Value2 = ReplaceVendorTable()[vendor];
                Worksheet.Range["F" + firstRow].Value2 = objectVendor.Discount;         //Столбец "Скидка". Вызывает значение Discount
                Worksheet.Range["G" + firstRow].FormulaLocal = string.Format(objectVendor.Formula_3, firstRow);     //Столбец "Цена". Вызывает формулу Formula_3
                Worksheet.Range["H" + firstRow].Formula = string.Format("=G{0}*(100-F{0})/100", firstRow);
                Worksheet.Range["I" + firstRow].Formula = string.Format("=H{0}*C{0}", firstRow);                
                firstRow++;
            }
            while (endRow > firstRow);          
        }

        private static IDictionary<string, string> ReplaceVendorTable()
        {
            Dictionary<string, string> dictionaryVendor = new Dictionary<string, string>()
            {
                {"Iek", "IEK"},
                {"Ekf", "EKF"},
                {"IekVa47", "IEK"},
                {"IekVa47m", "IEK"},
                {"IekArmat", "IEK"},
                {"EkfProxima", "EKF"},
                {"EkfAvers", "EKF"},
                {"Abb", "ABB"},
                {"Keaz", "KEAZ"},
                {"Dkc", "DKC"},
                {"Dekraft", "DEKraft"},
                {"Schneider", "Schneider"},
                {"Tdm", "TDM"}
            };
            return dictionaryVendor;
        }
    }
}

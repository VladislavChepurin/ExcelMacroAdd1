using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.Servises;
using System;

namespace ExcelMacroAdd.Functions
{
    internal class WriteExcel : AbstractFunctions 
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

        protected internal override void Start()
        {         
            int endRow = default;          
            //Вычисляем столбец на который установлен фокус
            int firstRow = cell.Row;
            if (firstRow == 1) firstRow++;
            int countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
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
                if (!(getArticle is null)) worksheet.get_Range("A" + firstRow).Value2 = getArticle;
                /*
                 * Функция dataInXml.Value.ReadFileXml("Formula_1") возвращает строку типа:
                 * =ВПР(A{0};'C:\Users\..
                 * Функция String.Format подставляет вместо {0} числовое значение firstRow
                */
                var objectVendor = dataInXml.ReadElementXml(vendor, dataInXml.ReadFileXml());
                worksheet.get_Range("B" + firstRow).FormulaLocal = String.Format(objectVendor.Formula_1, firstRow);    //Столбец "Описание". Вызывает формулу Formula_1            
                if (!(quantity == 0)) worksheet.get_Range("C" + firstRow).Value2 = quantity;
                worksheet.get_Range("D" + firstRow).FormulaLocal = String.Format(objectVendor.Formula_2, firstRow);    //Столбец "Кратность". Вызывает формулу Formula_2
                worksheet.get_Range("E" + firstRow).Value2 = Replace.RepleceVendorTable(vendor);
                worksheet.get_Range("F" + firstRow).Value2 = objectVendor.Discont;         //Столбец "Скидка". Вызывает значение Discont
                worksheet.get_Range("G" + firstRow).FormulaLocal = String.Format(objectVendor.Formula_3, firstRow);     //Столбец "Цена". Вызывает формулу Formula_3
                worksheet.get_Range("H" + firstRow).Formula = String.Format("=G{0}*(100-F{0})/100", firstRow);
                worksheet.get_Range("I" + firstRow).Formula = String.Format("=H{0}*C{0}", firstRow);                
                firstRow++;
            }
            while (endRow > firstRow);          
        }
    }
}

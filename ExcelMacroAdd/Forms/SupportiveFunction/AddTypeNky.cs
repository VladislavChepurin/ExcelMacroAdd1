using ExcelMacroAdd.Functions;
using Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Forms.SupportiveFunction
{
    internal class AddTypeNky : AbstractFunctions
    {
        private readonly string Type;
        public AddTypeNky(string type)
        {
            Type = type;
        }

        public override void Start()
        {
            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var currentCollum = Cell.Column;
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow;
            do
            {
                var currentValue = Worksheet.Cells[firstRow, currentCollum].Value;

                if (currentValue is null)
                {
                    Worksheet.Cells[firstRow, currentCollum].Value2 = Type;
                    firstRow++;
                    continue;
                }     
                var getType = currentValue.GetType();

                if (getType.FullName == "System.Double")
                {
                    var tempValue = Worksheet.Cells[firstRow, currentCollum].Value2.ToString();
                    Worksheet.Cells[firstRow, currentCollum].Value2 = tempValue + '+' + Type;
                }  
                
                firstRow++;
            }
            while (endRow > firstRow);
        }
    }
}

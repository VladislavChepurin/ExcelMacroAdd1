using ExcelMacroAdd.Functions;

namespace ExcelMacroAdd.Forms.SupportiveFunction
{
    internal class DeleteTypeNky : AbstractFunctions
    {
        public override void Start()
        {
            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var currentCollum = Cell.Column;
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow;
            do
            {
                Worksheet.Cells[firstRow, currentCollum].Value2 = null;
                firstRow++;
            }
            while (endRow > firstRow);
        }
    }
}

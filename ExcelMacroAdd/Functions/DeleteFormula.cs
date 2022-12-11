using System;

namespace ExcelMacroAdd.Functions
{
    internal sealed class DeleteFormula : AbstractFunctions
    {
        public override void Start()
        {
            Cell.Value = Cell.Value;                            //Удаляем формулы
            Worksheet.Range["A1", Type.Missing].Select();   //Фокус на ячейку А1   
        }
    }
}


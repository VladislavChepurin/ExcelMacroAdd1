﻿using System;

namespace ExcelMacroAdd.Functions
{
    internal class DeleteFormula : AbstractFunctions
    {
        public sealed override void Start()
        {
            cell.Value = cell.Value;                            //Удаляем формулы
            worksheet.get_Range("A1", Type.Missing).Select();   //Фокус на ячейку А1   
        }
    }
}


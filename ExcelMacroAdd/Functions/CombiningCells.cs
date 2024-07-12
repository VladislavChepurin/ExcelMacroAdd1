using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;

namespace ExcelMacroAdd.Functions
{
    internal class CombiningCells : AbstractFunctions
    {
        public override void Start()
        {          
            var value = Cell.Value;
            if (value != null && value is Object[,])
            {
                List<string> list = new List<string>();
                foreach (var item in value)
                {
                    if (item != null)
                        list.Add(item.ToString());
                }
                //Очистить выделенный диапазон
                var rng = Worksheet.get_Range(Cell.Address, Type.Missing);
                rng.ClearContents();
                //Вствить значения
                Application.ActiveCell.Value2 = String.Join(" ", list);
            }
        }
    }
}
